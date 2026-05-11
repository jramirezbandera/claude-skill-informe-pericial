"""Crea SÓLO la estructura de carpetas vacía para un encargo nuevo.

NO genera caso.yaml — eso lo escribe la skill (Claude) directamente con
los datos que recopile en la conversación interactiva.

Por defecto crea la carpeta DENTRO del directorio de trabajo actual (cwd).
Javier hace `cd` al sitio donde quiere el encargo y ahí se materializa.

Uso:
    cd "Z:\\03-INFORME\\AÑO 2026"
    python nuevo_encargo.py --apellido "García López" --localidad "Vélez-Málaga" --codigo "2026_001"

    # o explícito:
    python nuevo_encargo.py --apellido "..." --localidad "..." --codigo "..." --raiz "C:\\algun\\sitio"

Crea:
    <cwd>/<codigo>_<ApellidoSinEspacios>_<Localidad>/
        00_encargo/
        01_documentacion/
        02_parte_contraria/
        03_comunicaciones/
        04_inspeccion/{notas_tablet, fotos, fotos_renombradas}
        05_ensayos/
        06_referencias/
        99_salida/{presupuesto}

Imprime al final la ruta absoluta de la carpeta creada (única línea, prefijo "PATH=").
"""
from __future__ import annotations

import argparse
import os
import re
import sys
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

SUBDIRS = [
    "00_encargo",
    "01_documentacion",
    "02_parte_contraria",
    "03_comunicaciones",
    "04_inspeccion/notas_tablet",
    "04_inspeccion/fotos",
    "04_inspeccion/fotos_renombradas",
    "05_ensayos",
    "06_referencias",
    "99_salida/presupuesto",
]


def slug_apellido(apellido: str) -> str:
    parts = re.split(r"\s+", apellido.strip())
    return "".join(p.capitalize() for p in parts if p)


def slug_localidad(localidad: str) -> str:
    s = localidad.strip().replace(" ", "")
    s = s.replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u")
    s = s.replace("Á", "A").replace("É", "E").replace("Í", "I").replace("Ó", "O").replace("Ú", "U")
    s = s.replace("ñ", "n").replace("Ñ", "N")
    s = s.replace("-", "")
    return s


def main(argv: list[str]) -> int:
    p = argparse.ArgumentParser()
    p.add_argument("--apellido", required=True)
    p.add_argument("--localidad", required=True)
    p.add_argument("--codigo", required=True)
    p.add_argument("--raiz", default=None,
                   help="Directorio donde crear la carpeta del encargo. Por defecto: directorio de trabajo actual (cwd).")
    args = p.parse_args(argv)

    raiz = Path(args.raiz) if args.raiz else Path(os.getcwd())

    if not re.match(r"^\d{4}_\d{3}$", args.codigo):
        print(f"ERROR: --codigo debe tener formato YYYY_NNN, recibido {args.codigo!r}", file=sys.stderr)
        return 1

    nombre = f"{args.codigo}_{slug_apellido(args.apellido)}_{slug_localidad(args.localidad)}"
    encargo = (raiz / nombre).resolve()
    if encargo.exists():
        print(f"ERROR: ya existe {encargo}", file=sys.stderr)
        return 1

    encargo.mkdir(parents=True)
    for sub in SUBDIRS:
        (encargo / sub).mkdir(parents=True, exist_ok=True)

    print(f"OK — estructura creada en: {encargo}")
    print()
    print(f"PATH={encargo}")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
