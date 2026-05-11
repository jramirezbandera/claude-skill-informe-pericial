"""Inspecciona el estado de un encargo y produce una checklist legible.

Diseño: la documentación es heterogénea y casi nunca encaja con un nombre
"esperado". El checklist:

1. Comprueba que `caso.yaml` existe y tiene los campos básicos completos.
2. Cuenta archivos por carpeta raíz (00..06, 99) recursivamente, con tabla.
3. Para `04_inspeccion/`, separa fotos pendientes vs clasificadas y notas Notability.
4. Avisa si hay informe contraria detectada pero su toggle está apagado en yaml.

Salida: human-readable por defecto. Con --json escupe estructura para que la skill la procese.

Uso:
    python checklist.py --encargo "Z:\\03-INFORME\\Encargos\\..."
    python checklist.py --encargo "..." --json
"""
from __future__ import annotations

import argparse
import json
import os
import sys
from pathlib import Path

import yaml

try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass


CARPETAS_DOC = [
    "00_encargo",
    "01_documentacion",
    "02_parte_contraria",
    "03_comunicaciones",
    "05_ensayos",
    "06_referencias",
]

EXT_FOTO = {".jpg", ".jpeg", ".png", ".heic", ".heif", ".webp"}
EXT_DOCS = {".pdf", ".docx", ".doc", ".md", ".txt", ".xlsx", ".xls", ".csv"}
EXCLUDED_DIRS = {"_skill_workspace", "_pages"}


def es_archivo_generado(nombre: str) -> bool:
    n = nombre.lower()
    return n.endswith((".pdf.txt", ".pdf.preview.png", ".docx.txt", ".doc.txt"))


def es_visible(f: Path) -> bool:
    return f.is_file() and not f.name.startswith(".") and not es_archivo_generado(f.name)


def walk_visible(carpeta: Path):
    for ruta in carpeta.rglob("*"):
        if any(part in EXCLUDED_DIRS for part in ruta.parts):
            continue
        if es_visible(ruta):
            yield ruta


def contar_carpeta(carpeta: Path) -> tuple[int, list[str]]:
    """Devuelve (n_archivos, lista de subcarpetas con archivos dentro)."""
    if not carpeta.exists():
        return 0, []
    archivos = list(walk_visible(carpeta))
    subcarpetas = sorted({str(a.parent.relative_to(carpeta)) for a in archivos if a.parent != carpeta})
    return len(archivos), subcarpetas


def contar_fotos(encargo: Path) -> tuple[int, int]:
    pend = encargo / "04_inspeccion" / "fotos"
    clas = encargo / "04_inspeccion" / "fotos_renombradas"
    n_pend = sum(1 for f in pend.iterdir() if f.is_file() and f.suffix.lower() in EXT_FOTO) if pend.exists() else 0
    n_clas = sum(1 for f in clas.iterdir() if f.is_file() and f.suffix.lower() in EXT_FOTO) if clas.exists() else 0
    return n_pend, n_clas


def listar_pdfs_notas(encargo: Path) -> list[Path]:
    notas = encargo / "04_inspeccion" / "notas_tablet"
    if not notas.exists():
        return []
    return sorted(p for p in notas.iterdir() if p.is_file() and p.suffix.lower() == ".pdf")


def construir_checklist(encargo: Path) -> dict:
    caso_path = encargo / "caso.yaml"
    if not caso_path.exists():
        raise SystemExit(f"ERROR: no existe {caso_path}")
    with caso_path.open(encoding="utf-8") as f:
        caso = yaml.safe_load(f) or {}

    items = []

    # Datos básicos
    cliente = caso.get("cliente", {}) or {}
    inmueble = caso.get("inmueble", {}) or {}
    fechas = caso.get("fechas", {}) or {}
    obj = caso.get("objeto_pericial", "") or ""
    falta = []
    if not cliente.get("nombre"): falta.append("cliente.nombre")
    if not cliente.get("rol"): falta.append("cliente.rol")
    if not inmueble.get("direccion"): falta.append("inmueble.direccion")
    if not inmueble.get("tipologia"): falta.append("inmueble.tipologia")
    if not fechas.get("encargo"): falta.append("fechas.encargo")
    if not fechas.get("visita"): falta.append("fechas.visita")
    if not obj.strip() or "Pendiente" in obj or "RELLENAR" in obj.upper(): falta.append("objeto_pericial")
    if falta:
        items.append(("WARN", f"caso.yaml incompleto: faltan {', '.join(falta)}"))
    else:
        items.append(("OK", "caso.yaml — datos básicos completos"))

    # Carpetas de documentación: enumerar todo lo que hay
    conteo_por_carpeta = {}
    for sub in CARPETAS_DOC:
        n, subdirs = contar_carpeta(encargo / sub)
        conteo_por_carpeta[sub] = (n, subdirs)
        if n == 0:
            items.append(("INFO", f"{sub}/: vacía"))
        else:
            txt = f"{sub}/: {n} archivo(s)"
            if subdirs:
                txt += f" en {len(subdirs)} subcarpeta(s): {', '.join(subdirs[:3])}{'...' if len(subdirs)>3 else ''}"
            items.append(("OK", txt))

    # Detección informe contraria por contenido de 02_parte_contraria/
    n_contra, _ = contar_carpeta(encargo / "02_parte_contraria")
    sec = caso.get("secciones_a_incluir", {}) or {}
    if n_contra > 0 and not sec.get("analisis_critico_contraria"):
        items.append(("WARN", "Hay archivos en 02_parte_contraria/ pero el toggle analisis_critico_contraria está a false"))
    elif n_contra == 0 and sec.get("analisis_critico_contraria"):
        items.append(("WARN", "Toggle analisis_critico_contraria=true pero 02_parte_contraria/ está vacío"))

    # Inspección
    n_pend, n_clas = contar_fotos(encargo)
    notas = listar_pdfs_notas(encargo)
    if n_pend == 0 and n_clas == 0:
        items.append(("WARN", "04_inspeccion/fotos*/: sin fotos volcadas todavía"))
    elif n_pend > 0 and n_clas == 0:
        items.append(("INFO", f"04_inspeccion/fotos/: {n_pend} fotos sin renombrar — ejecuta procesar_fotos --renombrar"))
    elif n_pend > 0 and n_clas > 0:
        items.append(("INFO", f"04_inspeccion/: {n_pend} pendientes en fotos/, {n_clas} clasificadas"))
    else:
        items.append(("OK", f"04_inspeccion/fotos_renombradas/: {n_clas} fotos clasificadas"))

    if not notas:
        items.append(("INFO", "04_inspeccion/notas_tablet/: sin notas Notability"))
    else:
        items.append(("OK", f"04_inspeccion/notas_tablet/: {len(notas)} PDF(s)"))

    # Deficiencias
    deficiencias = caso.get("deficiencias", []) or []
    if not deficiencias:
        items.append(("INFO", "Sin deficiencias declaradas todavía en caso.yaml"))
    else:
        items.append(("OK", f"{len(deficiencias)} deficiencia(s) declarada(s) en caso.yaml"))

    # Inventario extraído?
    inv = encargo / "_skill_workspace" / "inventario.md"
    if not inv.exists():
        items.append(("INFO", "Inventario no generado — ejecuta extraer_textos.py para indexar la documentación"))
    else:
        items.append(("OK", "Inventario disponible en _skill_workspace/inventario.md"))

    return {
        "encargo": str(encargo),
        "codigo": caso.get("codigo"),
        "titulo_corto": caso.get("titulo_corto"),
        "items": items,
        "conteo_por_carpeta": conteo_por_carpeta,
        "fotos_pendientes": n_pend,
        "fotos_clasificadas": n_clas,
        "notas_pdfs": [n.name for n in notas],
        "deficiencias_count": len(deficiencias),
    }


ICONS_UNICODE = {"OK": "✅", "WARN": "⚠️ ", "FAIL": "❌", "INFO": "ℹ️ "}
ICONS_ASCII = {"OK": "[OK]  ", "WARN": "[WARN]", "FAIL": "[FAIL]", "INFO": "[INFO]"}
ICONS = ICONS_UNICODE if os.environ.get("INFORME_PERICIAL_ASCII") != "1" else ICONS_ASCII


def imprimir(checklist: dict) -> None:
    print(f"Encargo: {checklist['codigo']} — {checklist['titulo_corto']}")
    print(f"Path:    {checklist['encargo']}")
    print()
    for nivel, msg in checklist["items"]:
        print(f"  {ICONS.get(nivel, '-')} {msg}")
    print()


def main(argv: list[str]) -> int:
    p = argparse.ArgumentParser()
    p.add_argument("--encargo", required=True)
    p.add_argument("--json", action="store_true")
    args = p.parse_args(argv)

    encargo = Path(args.encargo)
    chk = construir_checklist(encargo)

    if args.json:
        print(json.dumps(chk, ensure_ascii=False, indent=2, default=str))
    else:
        imprimir(chk)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
