"""Lee notas de Notability (PDF) y prepara su contenido para Claude.

Modo 1 — texto: si el PDF tiene OCR de manuscrito, extrae texto plano.
Modo 2 — visión: si no hay texto, renderiza cada página como PNG en
         <pdf>/_pages/pagN.png para que Claude las lea con visión.

Uso:
    python leer_notability.py --pdf "ruta\\Nota.pdf"

Output:
    - Si hay texto: lo imprime en stdout y guarda <pdf>.txt
    - Si no hay texto: crea <pdf>/_pages/ con un PNG por página y avisa.
    - Siempre imprime un resumen al final con la decisión tomada.
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

import fitz  # pymupdf

try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass


MIN_TEXT_CHARS = 50  # umbral para considerar que hay texto extraíble


def main(argv: list[str]) -> int:
    p = argparse.ArgumentParser()
    p.add_argument("--pdf", required=True, help="Ruta al PDF de Notability")
    p.add_argument("--dpi", type=int, default=180, help="DPI para renderizado en modo visión (default 180)")
    p.add_argument("--force-vision", action="store_true", help="Forzar modo visión aunque haya texto")
    args = p.parse_args(argv)

    pdf_path = Path(args.pdf)
    if not pdf_path.exists():
        print(f"ERROR: no existe {pdf_path}", file=sys.stderr)
        return 1

    doc = fitz.open(str(pdf_path))
    n = doc.page_count
    text_chars = sum(len(doc[i].get_text()) for i in range(n))
    has_text = text_chars >= MIN_TEXT_CHARS

    print(f"PDF: {pdf_path.name}")
    print(f"Páginas: {n}")
    print(f"Caracteres extraíbles: {text_chars}")

    if has_text and not args.force_vision:
        # Modo texto
        out_txt = pdf_path.with_suffix(".txt")
        with out_txt.open("w", encoding="utf-8") as f:
            for i in range(n):
                f.write(f"\n=== PÁGINA {i+1} ===\n")
                f.write(doc[i].get_text())
        print(f"MODO: TEXTO (OCR de Notability detectado)")
        print(f"Texto guardado en: {out_txt}")
        return 0

    # Modo visión
    pages_dir = pdf_path.parent / "_pages"
    pages_dir.mkdir(exist_ok=True)
    saved = []
    for i in range(n):
        pix = doc[i].get_pixmap(dpi=args.dpi)
        out = pages_dir / f"{pdf_path.stem}_p{i+1:02d}.png"
        pix.save(str(out))
        saved.append(out)

    print(f"MODO: VISIÓN (sin OCR — Notability exportado sin reconocer manuscrito)")
    print(f"Renderizadas {n} páginas en {pages_dir} a {args.dpi} dpi")
    print(f"AVISO PARA JAVIER: la próxima vez exporta desde Notability con")
    print(f'                  "Reconocer texto manuscrito" activado para evitar el modo visión.')
    print()
    print("Páginas listas para que Claude las lea:")
    for s in saved:
        print(f"  {s}")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
