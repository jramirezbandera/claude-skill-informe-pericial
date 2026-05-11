"""Extrae una plantilla base .docx LIMPIA desde un informe existente.

Toma un .docx de referencia (típicamente CAORZA), conserva sus estilos
(Heading 1/2/3, Caption, fuentes, theme, numeración, page setup) y produce
una plantilla con:
- body vacío (sólo sectPr con referencias a header/footer)
- header genérico: "DICTAMEN PERICIAL — {{titulo_corto}}"
- footer con paginación: "Página X de Y"
- ningún archivo en word/media/ (sin imágenes embebidas)
- sin comments, sin headers/footers heredados, sin contenido residual

Trabaja a nivel de zip + XML para asegurar limpieza total — python-docx
no es suficiente porque no borra binarios ni partes del zip.

Uso:
    python extraer_plantilla.py --src "C:\\...\\INFORME CAORZA BENALMADENA_Vfinal.docx" \\
                                --dst "C:\\Users\\el usuario\\.claude\\skills\\informe-pericial\\plantillas\\plantilla_base.docx"
"""
from __future__ import annotations

import argparse
import re
import sys
import zipfile
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

# Partes del zip que conservamos sin tocar (heredan estilos / config)
PASS_THROUGH = {
    "word/styles.xml",
    "word/theme/theme1.xml",
    "word/fontTable.xml",
    "word/numbering.xml",
    "word/settings.xml",
    "word/webSettings.xml",
    "docProps/app.xml",
    "docProps/core.xml",
}

# Partes que reescribimos
REWRITE = {
    "word/document.xml",
    "word/_rels/document.xml.rels",
    "word/header1.xml",
    "word/footer1.xml",
    "[Content_Types].xml",
    "_rels/.rels",
}

# Partes que eliminamos (no las copiamos)
SKIP_PREFIXES = (
    "word/media/",
    "word/embeddings/",
)
SKIP_EXACT = {
    "word/header2.xml",
    "word/header3.xml",
    "word/footer2.xml",
    "word/footer3.xml",
    "word/comments.xml",
    "word/commentsExtended.xml",
    "word/commentsExtensible.xml",
    "word/commentsIds.xml",
    "word/endnotes.xml",
    "word/footnotes.xml",
    "word/people.xml",
    "word/glossary/document.xml",
}


# Nuevo document.xml: body vacío con una sectPr que referencia header1 y footer1.
# Tomamos el page setup (pgSz, pgMar) del docx original mediante extracción.

NEW_HEADER_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:pPr>
      <w:pStyle w:val="Header"/>
      <w:jc w:val="right"/>
    </w:pPr>
    <w:r>
      <w:rPr><w:sz w:val="18"/><w:color w:val="595959"/></w:rPr>
      <w:t xml:space="preserve">DICTAMEN PERICIAL — {{titulo_corto}}</w:t>
    </w:r>
  </w:p>
</w:hdr>
"""

NEW_FOOTER_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:pPr>
      <w:pStyle w:val="Footer"/>
      <w:jc w:val="center"/>
    </w:pPr>
    <w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:t xml:space="preserve">Página </w:t></w:r>
    <w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:instrText xml:space="preserve">PAGE</w:instrText></w:r>
    <w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:t xml:space="preserve"> de </w:t></w:r>
    <w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:instrText xml:space="preserve">NUMPAGES</w:instrText></w:r>
    <w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:fldChar w:fldCharType="end"/></w:r>
  </w:p>
</w:ftr>
"""

NEW_DOCUMENT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings" Target="webSettings.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
  <Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
  <Relationship Id="rIdHdr" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
  <Relationship Id="rIdFtr" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
</Relationships>
"""

NEW_PACKAGE_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
"""

NEW_CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
  <Override PartName="/word/webSettings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml"/>
  <Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>
  <Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
</Types>
"""


def extraer_sectpr(document_xml: str) -> str:
    """Extrae el último w:sectPr (page setup) del document.xml original.
    Si no se encuentra, devuelve uno por defecto A4 con márgenes razonables."""
    m = re.search(r"<w:sectPr[^>]*>.*?</w:sectPr>", document_xml, re.DOTALL)
    if m:
        sectpr = m.group(0)
        # Quitar referencias a headers/footers viejos y dejar las nuevas
        sectpr = re.sub(r"<w:headerReference[^/]*/>", "", sectpr)
        sectpr = re.sub(r"<w:footerReference[^/]*/>", "", sectpr)
        # Insertar nuestras referencias justo después de la apertura del sectPr
        sectpr = re.sub(
            r"(<w:sectPr[^>]*>)",
            r'\1<w:headerReference w:type="default" r:id="rIdHdr"/><w:footerReference w:type="default" r:id="rIdFtr"/>',
            sectpr,
            count=1,
        )
        return sectpr
    # Fallback A4
    return (
        '<w:sectPr>'
        '<w:headerReference w:type="default" r:id="rIdHdr"/>'
        '<w:footerReference w:type="default" r:id="rIdFtr"/>'
        '<w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="1417" w:right="1701" w:bottom="1417" w:left="1701" w:header="708" w:footer="708" w:gutter="0"/>'
        '<w:cols w:space="708"/>'
        '</w:sectPr>'
    )


def construir_document_xml(sectpr_xml: str) -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:document '
        'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<w:body>'
        + sectpr_xml +
        '</w:body>'
        '</w:document>'
    )


def main(argv: list[str]) -> int:
    p = argparse.ArgumentParser()
    p.add_argument("--src", required=True, help="docx fuente (informe completo)")
    p.add_argument("--dst", required=True, help="docx destino (plantilla)")
    args = p.parse_args(argv)

    src = Path(args.src)
    dst = Path(args.dst)
    if not src.exists():
        print(f"ERROR: no existe {src}", file=sys.stderr)
        return 1
    dst.parent.mkdir(parents=True, exist_ok=True)

    # Leer el zip fuente
    with zipfile.ZipFile(src) as z_in:
        original_doc_xml = z_in.read("word/document.xml").decode("utf-8")
        # Recolectar partes que pasan tal cual
        passthrough_data = {}
        for name in PASS_THROUGH:
            try:
                passthrough_data[name] = z_in.read(name)
            except KeyError:
                pass  # algunos pueden no existir, no es crítico

    # Extraer sectPr (page setup) original con nuestras referencias inyectadas
    sectpr = extraer_sectpr(original_doc_xml)
    new_document_xml = construir_document_xml(sectpr)

    # Escribir el zip destino
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z_out:
        # Pass-through
        for name, data in passthrough_data.items():
            z_out.writestr(name, data)
        # Reescritos
        z_out.writestr("[Content_Types].xml", NEW_CONTENT_TYPES)
        z_out.writestr("_rels/.rels", NEW_PACKAGE_RELS)
        z_out.writestr("word/_rels/document.xml.rels", NEW_DOCUMENT_RELS)
        z_out.writestr("word/document.xml", new_document_xml)
        z_out.writestr("word/header1.xml", NEW_HEADER_XML)
        z_out.writestr("word/footer1.xml", NEW_FOOTER_XML)

    tamano_kb = dst.stat().st_size / 1024
    print(f"OK — plantilla limpia creada en: {dst}")
    print(f"     Tamaño: {tamano_kb:.0f} KB (vs {src.stat().st_size/1024/1024:.0f} MB del fuente)")
    print(f"     Estilos heredados de: {src.name}")
    print()
    print("Contiene: estilos Heading 1/2/3, Caption, fuentes, theme, numeración, page setup.")
    print("Header genérico: 'DICTAMEN PERICIAL — {{titulo_corto}}'")
    print("Footer: paginación 'Página X de Y'")
    print("Body: vacío. Sin imágenes. Sin comments. Sin headers heredados.")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
