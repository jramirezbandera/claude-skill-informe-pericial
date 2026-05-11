"""Procesa una revisión del informe.

Detecta el último `informe_pericial_vN.docx` en `99_salida/`, acepta los
tracked changes, extrae los comentarios al margen, y produce
`informe_pericial_v(N+1).docx` SIN tracked changes ni comentarios.

Los comentarios extraídos se guardan en `_skill_workspace/comentarios_revision.md`
con anchor (texto al que apuntaban) y contexto, para que la skill (Claude
conversacional) los procese y aplique las ediciones que sugieran.

Uso:
    python revisar.py --encargo "<ruta>"

    # forzar versión de salida (default = siguiente)
    python revisar.py --encargo "<ruta>" --salida v3
"""
from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
from copy import deepcopy
from pathlib import Path

from lxml import etree

try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

W = "{%s}" % NS["w"]


def _qn(tag: str) -> str:
    """'w:p' -> '{namespace}p'."""
    if ":" in tag:
        prefix, local = tag.split(":")
        return "{%s}%s" % (NS[prefix], local)
    return tag


# ────────── Aceptar tracked changes ──────────

def aceptar_tracked_changes(doc_xml: bytes) -> bytes:
    """Devuelve el document.xml con tracked changes aceptados."""
    root = etree.fromstring(doc_xml)

    # 1. <w:del> y <w:moveFrom>: eliminar (texto borrado/movido fuera)
    for tag in ("w:del", "w:moveFrom"):
        for el in root.iter(_qn(tag)):
            parent = el.getparent()
            if parent is not None:
                parent.remove(el)

    # 2. <w:ins> y <w:moveTo>: reemplazar por su contenido
    for tag in ("w:ins", "w:moveTo"):
        # Importante: list() para no iterar mientras mutamos
        for el in list(root.iter(_qn(tag))):
            parent = el.getparent()
            if parent is None:
                continue
            idx = list(parent).index(el)
            # Insertar hijos en lugar del wrapper
            for child in reversed(list(el)):
                parent.insert(idx, child)
            parent.remove(el)

    # 3. Cambios de formato: <w:rPrChange>, <w:pPrChange>, <w:cellIns>, <w:cellDel>
    for tag in ("w:rPrChange", "w:pPrChange", "w:cellIns", "w:cellDel",
                "w:cellMerge", "w:trPrChange", "w:tcPrChange",
                "w:sectPrChange", "w:numberingChange"):
        for el in root.iter(_qn(tag)):
            parent = el.getparent()
            if parent is not None:
                parent.remove(el)

    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


# ────────── Extraer comentarios ──────────

def extraer_comentarios(zip_path: Path) -> list[dict]:
    """Extrae comentarios + sus anchors. Devuelve lista de dicts."""
    if not zip_path.exists():
        return []
    with zipfile.ZipFile(zip_path) as z:
        names = z.namelist()
        if "word/comments.xml" not in names or "word/document.xml" not in names:
            return []
        comments_raw = z.read("word/comments.xml")
        document_raw = z.read("word/document.xml")

    comments_root = etree.fromstring(comments_raw)
    document_root = etree.fromstring(document_raw)

    # comment id -> {author, date, text}
    by_id: dict[str, dict] = {}
    for c in comments_root.iter(_qn("w:comment")):
        cid = c.get(_qn("w:id"))
        author = c.get(_qn("w:author"), "") or ""
        date = c.get(_qn("w:date"), "") or ""
        # Concatenar todos los <w:t> del comentario
        texts = ["".join(t.text or "" for t in c.iter(_qn("w:t")))]
        by_id[cid] = {
            "id": cid,
            "author": author,
            "date": date,
            "text": "\n".join(t for t in texts if t).strip(),
            "anchor": "",
            "context": "",
        }

    # Buscar anchors: texto entre <w:commentRangeStart w:id="N"/> y <w:commentRangeEnd w:id="N"/>
    # Los range-start y range-end están dispersos en el body, posiblemente partidos
    # entre runs. Estrategia: recorrer el body preorder, llevar un set de "abiertos",
    # y acumular texto en runs cuando haya algún id abierto.
    body = document_root.find(_qn("w:body"))
    if body is None:
        return list(by_id.values())

    abiertos: set[str] = set()
    # Para contexto: acumular últimos 80 chars del párrafo donde se abrió cada comentario
    contexto_pre: dict[str, str] = {}
    parrafo_actual_text: list[str] = []
    parrafo_de_apertura: dict[str, list[str]] = {}

    def visitar(el):
        tag = etree.QName(el.tag).localname
        if tag == "p":
            parrafo_actual_text.clear()
        if tag == "commentRangeStart":
            cid = el.get(_qn("w:id"))
            if cid in by_id:
                abiertos.add(cid)
                contexto_pre[cid] = "".join(parrafo_actual_text)[-80:]
                parrafo_de_apertura[cid] = list(parrafo_actual_text)
        elif tag == "commentRangeEnd":
            cid = el.get(_qn("w:id"))
            if cid in by_id:
                abiertos.discard(cid)
        elif tag == "t":
            txt = el.text or ""
            parrafo_actual_text.append(txt)
            for cid in abiertos:
                by_id[cid]["anchor"] += txt

        for child in el:
            visitar(child)

    visitar(body)

    # Para el contexto: buscar el párrafo entero que contenía el anchor (limitado)
    # Mejor enfoque: buscar el texto del primer párrafo donde se abrió el comentario
    for cid, comment in by_id.items():
        comment["anchor"] = comment["anchor"].strip()
        # contexto = el texto del párrafo entero donde empieza, hasta 200 chars
        pre = contexto_pre.get(cid, "").strip()
        if pre:
            comment["context"] = pre

    return list(by_id.values())


# ────────── Construir docx limpio (sin comments, sin track-changes) ──────────

PARTES_COMENTARIO = {
    "word/comments.xml",
    "word/commentsExtended.xml",
    "word/commentsExtensible.xml",
    "word/commentsIds.xml",
    "word/people.xml",
    "word/threadedComments.xml",
}


def limpiar_referencias_comentarios(doc_xml: bytes) -> bytes:
    """Elimina commentRangeStart, commentRangeEnd, commentReference del body."""
    root = etree.fromstring(doc_xml)
    for tag in ("w:commentRangeStart", "w:commentRangeEnd", "w:commentReference"):
        for el in list(root.iter(_qn(tag))):
            parent = el.getparent()
            if parent is not None:
                parent.remove(el)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def limpiar_rels_y_content_types(zin: zipfile.ZipFile, name: str, data: bytes) -> bytes:
    """Quita relaciones y content types que apunten a partes de comentarios."""
    if name == "word/_rels/document.xml.rels":
        root = etree.fromstring(data)
        ns_pkg = "http://schemas.openxmlformats.org/package/2006/relationships"
        rt_comment_keywords = ("comment", "people", "threadedComment")
        for rel in list(root):
            target = rel.get("Target", "") or ""
            rtype = rel.get("Type", "") or ""
            if any(k in target.lower() or k in rtype.lower() for k in rt_comment_keywords):
                root.remove(rel)
        return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    if name == "[Content_Types].xml":
        root = etree.fromstring(data)
        for el in list(root):
            partname = el.get("PartName", "") or ""
            if "comment" in partname.lower() or "/people.xml" in partname.lower() or "threadedComment" in partname:
                root.remove(el)
        return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    return data


def construir_docx_revisado(src: Path, dst: Path) -> None:
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename in PARTES_COMENTARIO:
                continue  # eliminar parte entera
            data = zin.read(item.filename)
            if item.filename == "word/document.xml":
                data = aceptar_tracked_changes(data)
                data = limpiar_referencias_comentarios(data)
            elif item.filename in ("word/_rels/document.xml.rels", "[Content_Types].xml"):
                data = limpiar_rels_y_content_types(zin, item.filename, data)
            zout.writestr(item, data)


# ────────── Detección de versión ──────────

RE_VERSION = re.compile(r"^informe_pericial_v(\d+)\.docx$", re.IGNORECASE)


def encontrar_ultimo_v(salida: Path) -> tuple[Path, int] | None:
    candidatos = []
    for p in salida.iterdir():
        if not p.is_file():
            continue
        m = RE_VERSION.match(p.name)
        if m:
            candidatos.append((p, int(m.group(1))))
    if not candidatos:
        return None
    candidatos.sort(key=lambda x: x[1], reverse=True)
    return candidatos[0]


# ────────── Markdown de comentarios ──────────

def comentarios_a_markdown(comentarios: list[dict], origen: str, destino: str) -> str:
    if not comentarios:
        return (f"# Comentarios de revisión ({origen} → {destino})\n\n"
                "No se encontraron comentarios al margen.\n")
    lines = [f"# Comentarios de revisión ({origen} → {destino})", ""]
    lines.append(f"Total: **{len(comentarios)}**")
    lines.append("")
    lines.append("Para cada comentario: anchor (texto subrayado en Word), comentario al margen y contexto. "
                 "La skill aplica las ediciones que sugieran sobre el docx limpio.")
    lines.append("")
    for i, c in enumerate(comentarios, 1):
        lines.append(f"## Comentario #{i}")
        if c.get("author"):
            lines.append(f"- **Autor**: {c['author']}")
        if c.get("date"):
            lines.append(f"- **Fecha**: {c['date']}")
        anchor = c.get("anchor", "").strip()
        if anchor:
            lines.append(f"- **Anchor**: «{anchor}»")
        ctx = c.get("context", "").strip()
        if ctx and ctx != anchor:
            lines.append(f"- **Contexto previo**: …{ctx}")
        lines.append("- **Comentario**:")
        for ln in (c.get("text", "") or "").splitlines():
            lines.append(f"  > {ln}" if ln.strip() else "  >")
        lines.append("")
    return "\n".join(lines) + "\n"


# ────────── Main ──────────

def main(argv: list[str]) -> int:
    p = argparse.ArgumentParser()
    p.add_argument("--encargo", required=True)
    p.add_argument("--salida", default=None,
                   help="Versión de salida (vN). Default: siguiente al último encontrado.")
    args = p.parse_args(argv)

    encargo = Path(args.encargo).resolve()
    salida_dir = encargo / "99_salida"
    if not salida_dir.exists():
        print(f"ERROR: no existe {salida_dir}", file=sys.stderr)
        return 1

    detectado = encontrar_ultimo_v(salida_dir)
    if detectado is None:
        print(f"ERROR: no hay informe_pericial_vN.docx en {salida_dir}", file=sys.stderr)
        return 1
    src, n_actual = detectado

    # Determinar versión destino
    if args.salida:
        m = re.match(r"v?(\d+)$", args.salida.lower())
        if not m:
            print(f"ERROR: --salida mal formada {args.salida!r}", file=sys.stderr)
            return 1
        n_dst = int(m.group(1))
    else:
        n_dst = n_actual + 1

    dst = salida_dir / f"informe_pericial_v{n_dst}.docx"
    if dst.exists():
        print(f"ERROR: ya existe {dst}. Borra o usa --salida diferente.", file=sys.stderr)
        return 1

    print(f"Origen:  {src.name}")
    print(f"Destino: {dst.name}")

    # Extraer comentarios ANTES de procesar (sobre el src)
    comentarios = extraer_comentarios(src)
    print(f"Comentarios encontrados: {len(comentarios)}")

    # Construir docx revisado (acepta tracked changes + elimina comentarios)
    construir_docx_revisado(src, dst)
    print(f"Tracked changes aceptados y comentarios eliminados.")

    # Escribir markdown de comentarios
    ws = encargo / "_skill_workspace"
    ws.mkdir(parents=True, exist_ok=True)
    md = comentarios_a_markdown(comentarios, src.name, dst.name)
    md_path = ws / "comentarios_revision.md"
    md_path.write_text(md, encoding="utf-8")
    print(f"Comentarios extraídos a: {md_path}")

    # Resumen JSON para programático (opcional)
    json_path = ws / "comentarios_revision.json"
    json_path.write_text(json.dumps(comentarios, indent=2, ensure_ascii=False), encoding="utf-8")

    print(f"\nOK — {dst.name} ({dst.stat().st_size/1024:.0f} KB)")
    if comentarios:
        print(f"     Procesa los {len(comentarios)} comentarios desde {md_path.name}")
        print(f"     y aplica las ediciones que sugieran sobre {dst.name}.")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
