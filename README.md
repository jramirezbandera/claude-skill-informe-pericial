# claude-skill-informe-pericial

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![Made with Claude Code](https://img.shields.io/badge/Made%20with-Claude%20Code-D97757?logo=anthropic&logoColor=white)](https://claude.com/claude-code)
[![Lang: español](https://img.shields.io/badge/lang-español-red.svg)](README.md)
[![PRs Welcome](https://img.shields.io/badge/PRs-welcome-brightgreen.svg)](https://github.com/jramirezbandera/claude-skill-informe-pericial/pulls)

Skill para [Claude Code](https://claude.com/claude-code) que asiste a peritos arquitectos en la **redacción de dictámenes periciales de patologías constructivas** en `.docx`. Cubre el flujo completo: creación del expediente, ingesta de documentación, clasificación de fotografías, redacción, inyección del presupuesto exportado de Presto/Arquímedes y aplicación de revisiones.

> ⚠️ **No sustituye al criterio profesional.** Es una herramienta de ayuda. El perito firmante es responsable de revisar, validar y, en su caso, modificar todo lo que la skill produzca.

---

## ¿Qué hace exactamente?

Un dictamen pericial de patologías típico tiene una estructura repetitiva (juramento, bases, antecedentes, metodología, deficiencias, cuadro resumen, firma…) y mucho trabajo mecánico que NO añade valor profesional: clasificar fotos, ordenarlas por fecha, mapear partidas de Presto a cada defecto, calcular totales por capítulo, aceptar tracked changes…

La skill automatiza ese trabajo mecánico, deja al perito centrarse en lo que sí requiere criterio (causa, atribución, proporcionalidad de la reparación), y produce un `.docx` con los estilos profesionales habituales.

## Flujo de trabajo

```
1. /informe nuevo               → 3 preguntas → carpeta + caso.yaml esqueleto
2. (volcado: documentación, notas Notability, fotos a carpeta Drive)
3. /informe ingestar            → briefing (5 preguntas) + ingesta autónoma de PDFs
                                   + Excel para clasificar fotos + confirmación de huecos
4. /informe deficiencias        → fija la lista de deficiencias
5. /informe redactar            → genera informe_pericial_v1.docx (sin presupuestos)
6. (medición en Presto, exportar RTF a 99_salida/presupuesto/)
7. /informe presupuestar        → inyecta capítulos en cada deficiencia → v2.docx
8. (revisión del usuario en Word: tracked changes + comentarios)
9. /informe revisar             → acepta cambios, extrae comentarios → v3.docx
```

Cada subcomando es idempotente: puedes repetirlo si añades material nuevo.

## Subcomandos

| Comando | Qué hace |
|---|---|
| `/informe nuevo` | Crea la carpeta del encargo y `caso.yaml` esqueleto con 3 preguntas mínimas (apellido, localidad, código `YYYY_NNN`). |
| `/informe ingestar` | Briefing narrativo corto + ingesta autónoma de PDFs/DOCX al `caso.yaml` + procesamiento de fotos (Drive + EXIF + Excel) + notas Notability. Genera `notas_por_documento.md` y `cronologia.md` como memoria externa persistente. |
| `/informe deficiencias` | Fija la lista cerrada de deficiencias (modo dictado, edición manual de `deficiencias.md`, o aceptar propuesta autogenerada). |
| `/informe redactar` | Ensambla `informe_pericial_v1.docx` desde plantilla + boilerplate + prosa redactada por la skill + fotos clasificadas. |
| `/informe presupuestar` | Lee el RTF de Presto/Arquímedes, mapea subcapítulos a deficiencias y produce `v2.docx` con tablas de partidas + cuadro resumen + hoja resumen. |
| `/informe revisar` | Acepta tracked changes, extrae comentarios al margen y produce `v(N+1).docx`. La skill aplica luego las ediciones que sugieran los comentarios. |
| `/informe checklist` | Muestra el estado del encargo (✅/⚠️/❌). |

## Requisitos

- **Python 3.10+** (probado en 3.13)
- **Claude Code** instalado ([guía oficial](https://docs.claude.com/claude-code))
- **Microsoft Word** ≥2016 (para abrir y revisar los `.docx` finales)
- **Microsoft Excel** o LibreOffice Calc (para clasificar fotos en `.xlsx`)
- **Google Drive for Desktop** (recomendado, para que las fotos del móvil lleguen con EXIF)
- **Presto / Arquímedes** (opcional, para presupuestos)

### Paquetes Python

```
pip install python-docx pillow pillow-heif piexif pyyaml jinja2 openpyxl striprtf lxml
```

## Instalación

### 1. Clona el repo donde quieras

```bash
git clone https://github.com/jramirezbandera/claude-skill-informe-pericial.git
cd claude-skill-informe-pericial
```

### 2. Crea tu `datos_autor.yaml`

```bash
cp boilerplate/datos_autor.example.yaml boilerplate/datos_autor.yaml
# edita boilerplate/datos_autor.yaml con tus datos profesionales
```

Este fichero NO se subirá al repositorio (está en `.gitignore`).

### 3. Instálala como skill global de Claude Code

```bash
# Crea un enlace simbólico o copia el contenido a la carpeta de skills:
# Windows (PowerShell, como Admin):
New-Item -ItemType Junction -Path "$env:USERPROFILE\.claude\skills\informe-pericial" -Target "<ruta-donde-clonaste>"

# macOS / Linux:
ln -s "$(pwd)" "$HOME/.claude/skills/informe-pericial"
```

O simplemente copia toda la carpeta a `~/.claude/skills/informe-pericial/`.

### 4. (Opcional) Instala el slash command

Copia `commands/informe.md` a `~/.claude/commands/informe.md` para poder invocar `/informe nuevo`, `/informe ingestar`, etc. desde Claude Code CLI.

> En interfaces que no sean Claude Code CLI (Claude Desktop, web, etc.), los slash commands custom no funcionan. Habla a Claude en lenguaje natural: "crea un encargo nuevo", "ingesta el encargo", "redacta el v1"…

### 5. Personaliza la plantilla (opcional)

`plantillas/plantilla_base.docx` viene con estilos Heading 1/2/3, Caption, header/footer y page setup. Si quieres usar tus propios estilos:

1. Toma uno de tus informes finales `.docx`.
2. Ejecuta:
   ```bash
   python scripts/extraer_plantilla.py --src "<tu_informe>.docx" --dst plantillas/plantilla_base.docx
   ```
3. El script vacía el contenido, limpia metadatos personales, conserva estilos.

## Estructura de un encargo

Cuando ejecutas `/informe nuevo` se crea:

```
YYYY_NNN_ApellidoCliente_Localidad/
├── caso.yaml                       ← datos estructurados (apellido, juzgado, agentes, deficiencias…)
├── deficiencias.md                 ← fuente de verdad legible
├── 00_encargo/
│   ├── hoja_encargo.pdf
│   └── objeto_pericial.md
├── 01_documentacion/               ← flexible: cualquier PDF/DOCX que recibas
├── 02_parte_contraria/             ← informes de la otra parte
├── 03_comunicaciones/              ← emails, burofax
├── 04_inspeccion/
│   ├── fecha_inspeccion.txt
│   ├── notas_tablet/               ← Notability PDF (con OCR de manuscrito)
│   ├── fotos/                      ← copia local desde Drive (con EXIF)
│   ├── fotos_renombradas/          ← clasificadas: P001_DEF01_humedad_muro.jpg
│   └── clasificacion_fotos.xlsx    ← Excel con miniaturas para clasificar
├── 05_ensayos/                     ← opcional
├── 06_referencias/                 ← normativa, jurisprudencia
├── 99_salida/
│   ├── presupuesto/                ← RTFs exportados de Presto
│   └── informe_pericial_vN.docx
└── _skill_workspace/               ← memoria externa persistente
    ├── inventario.md
    ├── briefing.md
    ├── notas_por_documento.md
    ├── cronologia.md
    ├── redaccion_deficiencias.md
    └── ...
```

## Personalización del estilo

Los textos fijos del informe están en `boilerplate/` y son fáciles de adaptar:

| Fichero | Qué contiene | Sustituciones Jinja |
|---|---|---|
| `juramento_objetividad.md` | Texto literal del juramento (LEC) | — |
| `metodologia.md` | Sección 5 del informe | `{{ proyectista }}`, `{{ hay_informe_contraria }}` |
| `pie_firma.md` | Cierre del documento | `{{ lugar }}`, `{{ fecha_entrega_largo }}`, `{{ nombre_completo }}`, `{{ titulacion }}`, `{{ colegiado }}`, `{{ colegio }}`, `{{ despacho }}`, `{{ email }}` |
| `datos_autor.yaml` | Tus datos profesionales (privado) | — |

La estructura del informe está en `reference/estructura_macro.md`. La plantilla por deficiencia en `reference/plantilla_deficiencia.md`. El estilo de redacción (voz, tiempos, frases tipo) en `reference/estilo_redaccion.md`.

## Scripts disponibles

| Script | Para qué |
|---|---|
| `nuevo_encargo.py` | Crea la estructura de carpetas del encargo |
| `extraer_textos.py` | Vuelca a `.txt` el contenido de todos los PDFs/DOCX |
| `leer_notability.py` | Procesa notas manuscritas con/sin OCR |
| `procesar_fotos.py` | Copia desde Drive, ordena por EXIF, genera Excel de clasificación |
| `redactar_v1.py` | Ensambla el `.docx` v1 |
| `presupuestar.py` | Inyecta el presupuesto (RTF) en el v1 → v2 |
| `revisar.py` | Acepta tracked changes y extrae comentarios → v(N+1) |
| `checklist.py` | Muestra estado del encargo |
| `extraer_plantilla.py` | Regenera `plantilla_base.docx` desde un informe existente |

Todos llevan `--help` y pueden usarse de forma standalone.

## Filosofía de diseño

- **Separación inteligencia / maquinaria.** La skill (Claude conversacional) toma las decisiones que requieren criterio profesional (causa, atribución, proporcionalidad de la reparación) y redacta la prosa. Los scripts solo ensamblan, no piensan.
- **Memoria externa persistente.** Todo lo aprendido durante la ingesta se guarda en `_skill_workspace/` para sobrevivir a compactaciones de contexto.
- **El humano aprueba antes de fases pesadas.** No se redactan 50 páginas sin checkpoint.
- **El presupuesto lo hace Arquímedes, no la skill.** La skill solo trocea e inyecta el RTF de Presto.
- **Datos faltantes son la norma.** Muchos encargos no tienen proyecto visado ni informe contraria. La skill degrada secciones, no se rompe.

## Limitaciones conocidas

- Formato de presupuesto: sólo se ha probado con RTF exportado de **Presto/Arquímedes**. Otros programas (TCQ, Menfis, etc.) pueden requerir adaptar el parser en `scripts/presupuestar.py`.
- Tracked changes complejos (filas insertadas/borradas en tablas con `<w:cellIns>`/`<w:cellDel>`) se aceptan, pero casos edge pueden requerir abrir el `.docx` en Word y aceptar manualmente lo que no se haya procesado.
- El índice (TOC) del `.docx` se inserta como campo Word — hay que pulsar F9 en Word para regenerarlo al abrir el documento.
- HEIC del iPhone: se convierten a JPG en `_skill_workspace/jpegs/` antes de embeberlos en el `.docx`. Requiere `pillow-heif`.

## Aviso legal

Esta skill es una herramienta de **asistencia técnica**. El dictamen pericial firmado es responsabilidad exclusiva del perito firmante, que debe revisar y validar todo el contenido antes de su entrega. Los autores de esta skill no se hacen responsables del uso que terceros hagan de ella ni de los dictámenes que se elaboren con su ayuda.

La estructura, fórmulas tipo y estilo de redacción incluidos en `boilerplate/` y `reference/` reflejan una forma habitual de redactar dictámenes periciales de patologías constructivas en España, pero **no son la única válida** ni la prescrita por ninguna norma específica. Adáptalos a tu criterio profesional y a las particularidades de cada caso.

## Contribuciones

Las contribuciones son bienvenidas — especialmente:

- Parsers para otros programas de presupuesto (TCQ, Menfis…)
- Soporte para otros idiomas / colegios profesionales fuera de España
- Mejoras en la plantilla por deficiencia
- Tests automatizados

Abre un issue o un PR.

## Licencia

[MIT](LICENSE). Usa, modifica, comparte y comercializa con atribución.

## Autor original

Javier Ramírez Bandera — Arquitecto, COA Málaga.

Esta skill fue diseñada iterativamente con [Claude Code](https://claude.com/claude-code) a partir del flujo de trabajo real del autor, y publicada como herramienta abierta para la comunidad pericial.
