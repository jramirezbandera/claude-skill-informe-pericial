---
name: informe-pericial
description: Redacta dictámenes periciales de patologías constructivas en .docx, con el estilo de Javier Ramírez Bandera (arquitecto, COA Málaga). Úsalo cuando el usuario pida crear un encargo nuevo, redactar, presupuestar o revisar un informe pericial. Trigger: "informe pericial", "dictamen pericial", "informe de patologías", "encargo nuevo", "redacta el informe del [caso]", "/informe ...".
---

# Skill: informe-pericial

Genera dictámenes periciales en .docx siguiendo la estructura y estilo de Javier Ramírez Bandera.

**Ubicación de los encargos:** cada encargo se crea como subcarpeta del **directorio de trabajo actual** (cwd) cuando se invoca `/informe nuevo`. Javier hace `cd` a donde quiere que viva el encargo (ej. `Z:\03-INFORME\AÑO 2026\`) y allí se crea `<YYYY_NNN_apellido_localidad>/`. NO hay raíz fija predeterminada.

## Flujo de trabajo

```
1. /informe nuevo               → conversación guiada → carpeta + caso.yaml + objeto_pericial.md
2. (Javier vuelca documentación, fotos, notas Notability)
3. /informe ingestar            → indexa documentación, procesa fotos, lee notas
4. /informe deficiencias        → SE FIJA LA LISTA DE DEFICIENCIAS (dictada o propuesta)
5. /informe redactar            → genera v1 .docx (sin presupuestos)
6. (Javier mide en Presto, exporta los 2 docx a 99_salida/presupuesto/)
7. /informe presupuestar        → inyecta capítulos en cada deficiencia → v2
8. (Javier revisa en Word)
9. /informe revisar             → aplica cambios y genera vN+1
```

## Filosofía

- **`/informe nuevo` es MÍNIMO.** Sólo 3 preguntas para crear la carpeta. Nada más. El resto del `caso.yaml` se rellena en `/informe ingestar` leyendo los PDFs.
- **`/informe ingestar` empieza con un BRIEFING narrativo corto** (qué le ha pedido el cliente, para qué parte actúa, qué hay en juego) y termina pidiendo SOLO confirmación de los datos que no encontró o que son ambiguos.
- **No preguntar lo que está en los PDFs.** Dirección del inmueble, agentes de obra, fechas, juzgado, ref. catastral, abogados — todo eso lo extrae la ingesta.
- **Memoria externa persistente.** Para que la ventana de contexto no se sature en encargos grandes, durante la ingesta se escriben `_skill_workspace/notas_por_documento.md` (resumen estructurado por documento) y `_skill_workspace/cronologia.md` (línea de tiempo). En fases posteriores, **leer SIEMPRE primero esos dos ficheros** antes de tocar los `.txt` o PDFs originales.
- **`01_documentacion/` es flexible.** Acepta cualquier archivo y subcarpeta (presupuestos del cliente, correos, ensayos, fotos antiguas, informes de terceros, notas de reunión…). La skill ingesta todo y lo tiene en cuenta al redactar.
- **Datos faltantes son la norma.** Muchos encargos no tienen proyecto visado, ni informe contraria. La skill degrada secciones, no se rompe.
- **El humano aprueba antes de fases pesadas.** No redactar 50 páginas sin checkpoint.
- **Los presupuestos los hace Arquímedes, no la skill.** La skill sólo trocea e inyecta.
- **Cada deficiencia se redacta sin coste primero (v1).** Los costes entran en v2 desde los docx que Javier exporta.

---

## Subcomandos

### `/informe nuevo` (también: "encargo nuevo de…", "nuevo encargo", "vamos a crear el encargo")

**Objetivo: crear la carpeta lo más rápido posible.** Sólo 3 preguntas. Todo lo demás se rellena después en `/informe ingestar` leyendo los PDFs.

**Conducta:**

1. **3 preguntas mínimas, una por turno**, en lenguaje natural:

   - **"¿Apellido o razón social del cliente?"** (ej. "García López", "Comunidad de Propietarios Villa Carmen")
   - **"¿Localidad del inmueble?"** (ej. "Vélez-Málaga", "Mijas")
   - **"¿Código del encargo? (formato YYYY_NNN, ej. 2026_007)"** — Si Javier dice "el siguiente" o "no sé", escanea las carpetas hermanas en cwd con patrón `YYYY_NNN_*` del año actual y propón el siguiente número libre.

   **NO preguntar nada más.** Ni dirección, ni agentes, ni juzgado, ni para qué parte actúa, ni objeto pericial. Eso viene en la ingesta.

2. **Crea la estructura** ejecutando:
   ```
   python scripts/nuevo_encargo.py --apellido "..." --localidad "..." --codigo "..."
   ```
   El script crea la carpeta DENTRO del directorio de trabajo actual del usuario (cwd). Captura `PATH=<ruta>` para los pasos siguientes.

3. **Escribe `caso.yaml` SKELETO** en `<ruta>/caso.yaml` con sólo los 3 datos recopilados + toggles por defecto + el resto de campos vacíos. Usa el schema de `reference/caso_yaml_schema.yaml`. Toggles por defecto:
   - `metodologia: true`
   - `desarrollo_anormal_obra: false`
   - `analisis_critico_contraria: false` (se activará en ingesta si aparece informe contrario)
   - `penalizaciones_retraso: false`
   - `hoja_resumen_presupuesto: true`
   - `template_deficiencia: "estandar"`
   - `fecha_encargo: <hoy en ISO>`

4. **NO escribas `00_encargo/objeto_pericial.md` todavía.** Eso lo redactará la ingesta a partir del briefing narrativo + lectura de la hoja de encargo.

5. **Mensaje de cierre, breve**:
   > Carpeta creada en `<ruta>`.
   >
   > Ahora vuelca el material:
   > - **Documentación** (hoja de encargo, demanda, contrato, proyecto, correos, presupuestos del cliente, informe contrario…) → `00_encargo/`, `01_documentacion/`, `02_parte_contraria/` o `03_comunicaciones/` según convenga.
   > - **Notas Notability** (con OCR de manuscrito activado) → `04_inspeccion/notas_tablet/`.
   > - **Fotos** → a tu carpeta de Drive del encargo (app Drive desde galería del móvil — preserva EXIF). Cuando ingestemos, te pido la ruta de Drive y las copio aquí.
   >
   > Cuando lo tengas, dime **"ingesta el encargo"** y arrancamos: te haré 4-5 preguntas rápidas de contexto y luego proceso todo de golpe.

### `/informe ingestar` (o "ingesta el encargo", "lee toda la documentación")

Procesa todo lo que Javier ha volcado tras `/informe nuevo`. Se puede ejecutar varias veces conforme se añade material; es idempotente.

**Estructura: BRIEFING → INGESTA AUTÓNOMA → CONFIRMACIÓN DE HUECOS.**

#### Fase A — BRIEFING (preguntas de contexto, antes de leer nada)

Conversación corta, **una pregunta por turno**, en lenguaje natural. El objetivo es darle a la skill el "marco" del caso para que luego sepa qué buscar en los PDFs y cómo interpretarlos.

1. **"Cuéntame en 2-3 frases: ¿qué te ha pedido el cliente y qué problema hay con la obra?"**
   → Esto pasa a `00_encargo/objeto_pericial.md` (Javier puede afinar luego).

2. **"¿Para qué parte actúas? (actora, demandada o extrajudicial)"**

3. **"¿Hay procedimiento judicial abierto?"** (sí/no — si dice sí, NO pidas juzgado/autos: lo extraerá la ingesta de la demanda.)

4. **"¿Hay informe de la parte contraria que tenga que rebatir?"** (sí/no — si dice sí, activa `analisis_critico_contraria: true`.)

5. **"¿Algo importante que NO esté en los documentos y debas saber para leerlos en contexto?"** (texto libre, opcional — guarda en `_skill_workspace/briefing.md` para tenerlo a mano al redactar.)

Tras el briefing, **escribe `00_encargo/objeto_pericial.md`** con la respuesta de la pregunta 1, y actualiza `caso.yaml` con las respuestas 2-4. La 5 va a `_skill_workspace/briefing.md`.

#### Fase B — INGESTA AUTÓNOMA (sin más preguntas)

1. **Indexar documentación** — `python scripts/extraer_textos.py --encargo <path>`. Recorre recursivamente `00_encargo/`, `01_documentacion/`, `02_parte_contraria/`, `03_comunicaciones/`, `05_ensayos/`, `06_referencias/` y vuelca texto plano de cada PDF/DOCX a un `.txt` paralelo. Genera `_skill_workspace/inventario.md`.

2. **Leer inventario** — Read tool sobre `_skill_workspace/inventario.md`.

3. **Extracción autónoma de datos al `caso.yaml`** — leer los `.txt` relevantes (hoja de encargo, demanda, contrato, proyecto, escritura, informe contrario) y rellenar TODO lo que se pueda:
   - **Inmueble**: dirección completa, CP, municipio, tipología, ref. catastral, superficie.
   - **Procedimiento**: juzgado, nº de autos, abogado y procurador del cliente.
   - **Agentes de obra**: promotor, constructor, proyectista, DF, DEO.
   - **Fechas**: inicio/fin de obra, recepción, encargo (si aparece distinto al de hoy).
   - **Toggles**: si aparece informe contrario y aún no estaba activado, activarlo. Si la documentación habla de incidencias graves de obra → considerar `desarrollo_anormal_obra: true`.

   **Regla de oro**: si un dato aparece en un PDF, NO se le vuelve a preguntar a Javier. Sólo se confirma al final si hay ambigüedad (ej. dos direcciones distintas en dos documentos).

3 bis. **MEMORIA EXTERNA: notas por documento + cronología** *(crítico para encargos con mucha documentación)*

   Al ingestar, mi ventana de contexto puede saturarse si los PDFs son grandes. Para que ni `/informe redactar` ni `/informe revisar` tengan que releer los PDFs originales, durante esta fase escribo DOS ficheros persistentes que después serán mi primera lectura:

   **a) `_skill_workspace/notas_por_documento.md`** — un bloque por cada documento RELEVANTE leído (no para volcados periféricos; sí para hoja de encargo, demanda, contrato, proyecto, informe contraria, certificaciones, comunicaciones clave, ensayos). Formato exacto:

   ```markdown
   # Notas por documento

   ## <ruta-relativa-al-encargo>
   - **Tipo**: <hoja de encargo | demanda | contrato | proyecto | informe contraria | certificación | comunicación | ensayo | otro>
   - **Autor / emisor**: <quién>
   - **Fecha**: <YYYY-MM-DD si se conoce>
   - **Datos clave**:
     - <dato 1 con valor concreto>
     - <dato 2>
     - ...
   - **Citas literales relevantes**:
     - p. <N>: «<cita textual entre comillas tipográficas>»
     - ...
   - **Anomalías / a tener en cuenta**:
     - <observación 1>
   - **Útil para deficiencias**: <slug1>, <slug2>  (o "general" si es contexto transversal)
   ```

   Reglas:
   - 5–15 bullets por documento. Si un documento no aporta nada relevante, NO crear bloque (mejor un inventario corto que un dump masivo).
   - **Citas literales con número de página** — son lo más valioso porque permiten redactar sin volver al PDF.
   - **"Útil para deficiencias"** acelera la fase de redacción: al redactar `humedades_garaje`, busco en el .md los bloques que mencionen ese slug.
   - Idempotente: si el .md ya existe, releerlo y actualizar/añadir, no sobrescribir.

   **b) `_skill_workspace/cronologia.md`** — línea de tiempo del caso construida cruzando fechas de todos los documentos. Formato exacto:

   ```markdown
   # Cronología del caso

   ## YYYY
   - **DD/MM** — <hecho breve>. *[fuente: <ruta-relativa>, p.<N>]*
   - **DD/MM** — <otro hecho>. *[fuente: ...]*

   ## YYYY+1
   - ...
   ```

   Reglas:
   - Agrupado por año. Dentro, orden cronológico.
   - Cada línea con fuente entre corchetes en cursiva.
   - Hechos relevantes: contratos, actas de obra, certificaciones, comunicaciones (burofax, requerimientos), recepciones, daños reportados, peritajes previos, demandas, visitas de Javier, encargo. NO incluir trivialidades.
   - Útil sobre todo para la sección de "ANTECEDENTES" y para "DESARROLLO ANORMAL DE LA OBRA" si toca.

   Estos dos ficheros son lectura PRIORITARIA en fases posteriores y permiten que la conversación no dependa de tener los PDFs cargados en contexto.

4. **Procesar fotos** (origen: carpeta de Drive del encargo, vía Google Drive for Desktop). Flujo nuevo basado en Excel:

   **Paso 4.1 — Copiar y renombrar:**
   - Preguntar a Javier la ruta de la carpeta de Drive del encargo (típicamente bajo `G:\Mi unidad\...`).
   - Ejecutar `python scripts/procesar_fotos.py --encargo <path> --renombrar --source "<ruta_drive>"`.
   - El script copia las fotos a `04_inspeccion/fotos/`, las ordena por EXIF, las renombra `P001..PNNN` cronológicamente, escribe `fecha_inspeccion.txt` y **genera `04_inspeccion/clasificacion_fotos.xlsx`** con miniaturas embebidas + agrupación temporal (gap >5 min = nuevo grupo) + columnas `Deficiencia` (con dropdown desde `deficiencias.md`), `Descripción` y `Notas`.
   - **Pre-requisito clave**: para que el dropdown de Deficiencia del Excel salga útil, conviene que `deficiencias.md` esté ya creado (modo borrador es suficiente — Javier puede ajustarlo en `/informe deficiencias`). Si no existe aún, el Excel se genera sin dropdown y Javier teclea los slugs a mano.

   **Paso 4.2 — Javier rellena el Excel** (no en chat, en su Excel):
   - Para cada foto: asigna `Deficiencia` (slug, o `general`, o `descartar`) y `Descripción` (5-15 palabras: lo que se ve, será el pie de imagen).
   - Guarda el .xlsx.

   **Paso 4.3 — Aplicar la clasificación:**
   - Cuando Javier diga "aplica la clasificación", ejecutar `python scripts/procesar_fotos.py --encargo <path> --aplicar-excel`.
   - El script lee el .xlsx, mueve cada foto a `04_inspeccion/fotos_renombradas/` con nombre `P001_DEF01_slug-de-la-descripcion.jpg` (el slug se deriva automáticamente de la Descripción), descarta las marcadas como `descartar`, y escribe `_skill_workspace/fotos_descripciones.json` con todas las descripciones (formato `{P001: {deficiencia, descripcion, ...}}`).
   - Avisa si hay fotos sin Deficiencia rellena.

   **Cómo se usa luego**: `redactar_v1.py` lee `fotos_descripciones.json` y usa la `Descripción` que escribió Javier como **pie de imagen** ("Imagen NN: descripción") en el docx, en vez de derivar del nombre del archivo. Las fotos con `Deficiencia=general` no entran en ninguna sección 7.x — se reservan para que la skill las inserte conversacionalmente en Antecedentes o Portada si conviene.

   **Modo legacy** (sin Excel, dictando rangos): `--clasificar --map "1-12:DEF01:slug"`. Sigue funcionando pero no es el flujo recomendado.

   **Regenerar el Excel**: si Javier añade más fotos a Drive después o quiere reiniciar la clasificación, ejecutar `--regenerar-excel` para reconstruir el .xlsx (atención: pierde lo que hubiera rellenado).

5. **Leer notas Notability** — para cada PDF en `04_inspeccion/notas_tablet/`:
   - `python scripts/leer_notability.py --pdf <ruta>`.
   - Si modo TEXTO → Read del `.txt` paralelo. Si modo VISIÓN → Read de cada PNG en `_pages/`.
   - Avisar si una nota no tiene OCR.

6. **Generar BORRADOR de deficiencias.md** — combinando briefing + notas + fotos + documentación, escribe `deficiencias.md` en la raíz del encargo. NO toca aún `caso.yaml > deficiencias`.

#### Fase C — CONFIRMACIÓN DE HUECOS (sólo lo que falta o es ambiguo)

Tras la ingesta autónoma, mostrar a Javier un **resumen estructurado** con tres listas:

```
✅ ENCONTRADO (no preguntar):
   - Dirección: <calle nº, CP, municipio>  [fuente: escritura.pdf]
   - Promotor: <nombre>  [fuente: contrato.pdf]
   - Juzgado: <…>  [fuente: demanda.pdf]
   - …

⚠️ AMBIGUO / CONFLICTIVO (necesito tu confirmación):
   - Fecha de fin de obra: el contrato dice 2023-06-15, la recepción 2023-09-30. ¿Cuál uso?
   - Hay dos direcciones: <A> en escritura, <B> en demanda. ¿Cuál es la del inmueble peritado?

❌ NO ENCONTRADO (rellena si lo sabes, o "no sé"):
   - Referencia catastral
   - Abogado del cliente
   - Fecha de inicio de obra
```

Pedir las respuestas en bloque (Javier puede contestar todas seguidas en un solo mensaje o decir "no sé" donde no las tenga). Aplicar al `caso.yaml`.

7. **Ejecutar checklist** — `python scripts/checklist.py --encargo <path>` para mostrar el estado final.

8. **Cierre**:
   > Ingesta completada. He generado una propuesta de deficiencias en `deficiencias.md`. Cuando estés conforme con la lista, dime **"fijemos las deficiencias"** (o `/informe deficiencias`) para cerrarla y poder redactar.

### `/informe deficiencias` (o "fijemos las deficiencias", "dicto las deficiencias", "lee deficiencias.md")

Momento EXPLÍCITO del flujo donde se cierra la lista de deficiencias antes de redactar. Tres modos según cómo lo invoque Javier:

**Modo A — Dictado:** Javier dice algo tipo "son cuatro: humedades en el garaje, fisuras en fachada norte, problemas en la escollera del lindero sur, y filtraciones en la sala de calderas". Claude:
1. Parsea la frase y propone slugs cortos (snake_case): `humedades_garaje`, `fisuras_fachada_norte`, `escollera_sur`, `filtraciones_calderas`.
2. Para cada una asigna id correlativo y título legible (capitalizado).
3. Si tras `/informe ingestar` ya hay material asociado (notas/fotos), enlaza el origen.
4. Escribe `deficiencias.md` y muestra la lista para confirmación.

**Modo B — Editar el fichero:** Javier ha editado `deficiencias.md` a mano (añadido, quitado o renombrado). Claude lo lee y sincroniza `caso.yaml > deficiencias`.

**Modo C — Aceptar la propuesta de `/informe ingestar`:** si Javier dice "vale la propuesta" o similar, Claude lee el `deficiencias.md` que generó `ingestar` y lo sincroniza con `caso.yaml`.

En cualquier modo, al terminar:
- `deficiencias.md` queda como fuente de verdad legible para Javier.
- `caso.yaml > deficiencias` queda sincronizado (id, slug, titulo, capitulo_presupuesto, origen).
- Mostrar resumen: "Fijadas N deficiencias: [lista]. Listo para `/informe redactar`."

**Formato de `deficiencias.md`** (Claude lo crea/edita con Read/Write/Edit, sin script):

```markdown
# Deficiencias — encargo 2026_005

1. **humedades_garaje** · Humedades en muros del garaje
   - Origen: notas tablet pag 3, fotos P012-P018
   - Notas: filtración por muro este, agravada en lluvia
   - Capítulo presupuesto: CAPÍTULO 01 — HUMEDADES GARAJE

2. **fisuras_fachada_norte** · Fisuras en fachada norte
   - Origen: notas tablet pag 1, fotos P001-P008
   - Notas: fisuras horizontales en planta 2ª, junto a ventanal
   - Capítulo presupuesto: CAPÍTULO 02 — FISURAS FACHADA NORTE

3. **escollera_sur** · Desprendimiento de ripios en escollera sur
   - Origen: dictada por Javier, sin fotos asociadas todavía
   - Capítulo presupuesto: CAPÍTULO 03 — ESCOLLERA SUR
```

Reglas:
- El número del item es el `id`.
- La palabra en negrita es el `slug` (snake_case, sin acentos, corto).
- Tras el · va el `título` (capitalizado, frase corta).
- "Capítulo presupuesto" lo necesita la fase 2 — si Javier no lo ha asignado, Claude lo deriva del slug en mayúsculas y deja el número correlativo.

**`/informe deficiencias` se puede ejecutar tantas veces como haga falta.** Si tras redactar el v1 Javier descubre que falta una deficiencia, puede volver aquí, añadirla, y volver a `/informe redactar`.

### `/informe checklist` (o "qué tengo y qué falta", "estado del encargo")

`python scripts/checklist.py --encargo <path>`. Pinta ✅/⚠️/❌/ℹ️ con:
- caso.yaml — campos básicos completos o lista de los que faltan
- Conteo de archivos por carpeta raíz (incluye subcarpetas)
- Avisos sobre toggles vs realidad (informe contraria detectado pero toggle off, etc.)
- Fotos pendientes vs clasificadas
- Notas Notability disponibles
- Deficiencias declaradas

Variante: `--json` para procesar la salida programáticamente.

### `/informe redactar` (o "redacta el v1", "genera el informe")

**Pre-requisitos:**
- `/informe ingestar` ejecutado.
- `/informe deficiencias` ejecutado (lista fijada en `deficiencias.md` y reflejada en `caso.yaml`).
- Checklist sin ❌ críticos.

Si falta algo, NO redactar — pedir a Javier que complete primero.

**Arquitectura: separación inteligencia / maquinaria.**

La skill (Claude conversacional) genera la prosa de cada sección que requiere redacción experta y la guarda en `_skill_workspace/`. Después, el script `redactar_v1.py` ENSAMBLA el `.docx` desde esas piezas + boilerplate + `caso.yaml` + plantilla. El script NO redacta nada, solo monta.

#### Fase 0 — Cargar memoria externa (OBLIGATORIO antes de redactar)

**Leer SIEMPRE primero**, en este orden, antes de tocar PDFs o `.txt`:

1. `_skill_workspace/briefing.md` — contexto narrativo del caso.
2. `_skill_workspace/notas_por_documento.md` — síntesis estructurada por documento (citas literales con página).
3. `_skill_workspace/cronologia.md` — línea de tiempo del caso.
4. `caso.yaml` — datos estructurados.
5. `deficiencias.md` — lista cerrada.

Estos cinco ficheros contienen el 95% de lo que necesito para redactar. Solo recurrir a un `.txt` paralelo de un PDF si necesito **una cita exacta** que no está en `notas_por_documento.md` o **verificar** algo concreto. **Nunca releer un PDF entero en esta fase**.

Si alguno de los cinco falta (notas o cronología vacíos), volver a `/informe ingestar` para que se completen antes de redactar — no improvisar leyendo PDFs sueltos.

#### Fase 1 — Generar la prosa (Claude, conversacional)

En este orden, escribir cada fichero leyendo lo que sea necesario y aplicando `reference/estilo_redaccion.md` y `reference/plantilla_deficiencia.md`:

1. **`_skill_workspace/documentacion_consultada.md`** — un párrafo redactado (no listado plano) que enumere los documentos efectivamente consultados durante la ingesta. La fuente es `_skill_workspace/inventario.md` más la clasificación que tú hagas (hoja de encargo, demanda, contrato, proyecto, certificaciones, informe contraria, comunicaciones, reportaje fotográfico, notas manuscritas…). Estilo: prosa formal, p.ej. *"Para la elaboración del presente dictamen se ha consultado: hoja de encargo de fecha…, demanda de la actora con número de autos… presentada ante…, contrato de obra…, proyecto de ejecución visado el… redactado por…, certificaciones de obra nº 1 a nº 12, reportaje fotográfico realizado durante la visita y notas manuscritas tomadas in situ"*.

2. **`_skill_workspace/antecedentes.md`** *(opcional)* — si quieres redactar la sección 4 con tu propia prosa en vez de dejar que el script monte un fallback estructurado por agentes/fechas. Útil cuando hay desarrollo cronológico interesante o detalles del entorno que merezcan prosa.

3. **`_skill_workspace/redaccion_deficiencias.md`** — la pieza grande. Para cada deficiencia de `deficiencias.md`, redactar los 5 sub-apartados de `reference/plantilla_deficiencia.md` aplicando `reference/estilo_redaccion.md`. **Formato exacto** (el script parsea por estos encabezados):

   ```markdown
   # Redacción de deficiencias

   ## DEF01 — slug · Título legible

   ### Existencia
   Prosa…

   ### Definición en proyecto
   Prosa…

   ### Causa
   Prosa…

   ### Certificación
   (opcional — omitir el bloque entero si no aplica)

   ### Propuesta de reparación
   Prosa…
   ```

   **Reglas críticas**:
   - El código `DEFNN` se infiere del id en `deficiencias.md` (id 1 → DEF01, id 12 → DEF12).
   - Usar tercera persona impersonal ("se ha podido constatar"), nunca "yo".
   - Citas literales del proyecto entre comillas tipográficas («…»).
   - Si falta proyecto visado, sustituir la fórmula por la de `plantilla_deficiencia.md`.
   - **NO incluir las fotos en el .md**: las fotos se insertan automáticamente desde `04_inspeccion/fotos_renombradas/` por el script (busca por patrón `*_DEF{NN}_*` y las pone en "Existencia" con caption `Imagen NN: <descripción del slug>`).
   - **NO incluir el bloque de presupuesto en "Propuesta de reparación"**: el script añade automáticamente `[[PRESUPUESTO_DEFICIENCIA: <slug>]]` para que `/informe presupuestar` lo sustituya en fase 2.

4. **`_skill_workspace/desarrollo_anormal.md`** *(opcional)* — sólo si `caso.yaml > secciones_a_incluir.desarrollo_anormal_obra: true`.

5. **`_skill_workspace/analisis_critico_contraria.md`** *(opcional)* — sólo si `analisis_critico_contraria: true`. Aplicar el approach validado: análisis global con incisiones puntuales en contradicciones (no defecto a defecto).

6. **`_skill_workspace/penalizaciones.md`** *(opcional)* — sólo si `penalizaciones_retraso: true`.

**Iteración con Javier**: redactar deficiencia a deficiencia (no las 5 de golpe), enseñarle cada bloque, recoger correcciones, fijar y pasar a la siguiente. Esto da control y evita re-trabajo masivo.

#### Fase 2 — Ensamblar el docx (script)

Cuando todas las piezas estén generadas y aprobadas:

```
python scripts/redactar_v1.py --encargo "<ruta-encargo>" \
    [--lugar "Málaga"] [--fecha-entrega 2026-05-20]
```

El script:
- Carga la plantilla `plantillas/plantilla_base.docx` (limpia, 24 KB) e inyecta `titulo_corto` en el header.
- Monta portada, índice, las 12 secciones según `reference/estructura_macro.md` (omitiendo opcionales no activadas en `caso.yaml`).
- Sustituye boilerplate (`juramento`, `metodologia`, `pie_firma`) con renderizado Jinja2.
- Para cada deficiencia: sub-apartados desde `redaccion_deficiencias.md`, fotos automáticas desde `fotos_renombradas/` con caption `Imagen NN: descripción`, marcador `[[PRESUPUESTO_DEFICIENCIA: <slug>]]`.
- Si una deficiencia no tiene prosa redactada, deja `[[FALTA: redactar X]]` en el docx — el script no falla, sólo avisa al final.
- HEIC/HEIF se convierten automáticamente a JPG en `_skill_workspace/jpegs/`.
- Cuadro resumen de conclusiones se monta como tabla con columnas `Nº | Deficiencia | Causa breve | Imputable a | Coste reparación` (las dos últimas con marcador para `/informe presupuestar`).

Salida: `99_salida/informe_pericial_v1.docx`.

**Tras ejecutar el script**: avisa a Javier que abra el docx y pulse F9 sobre el índice para regenerar la TOC, y mostrar el resumen de avisos del script.

### `/informe presupuestar` (o "inyecta el presupuesto")

**Pre-requisito:** Javier ha exportado de Presto/Arquímedes **un fichero `.RTF`** a `99_salida/presupuesto/`. El nombre puede ser cualquiera — si hay varios, se usa el más reciente. Si hay versiones (p.ej. `..._02_02_...`, `..._03_03_...`), el script toma el último mtime.

**Estructura del RTF de Presto** (lo que el parser espera):
- Capítulos: línea `CAPÍTULO NN <título>`
- Subcapítulos: línea `SUBCAPÍTULO NN.MM <título>` — **cada subcapítulo mapea a UNA deficiencia**
- Partidas: línea `NN.MM.PP <ud> <descripción>` + descripción larga + líneas de medición + total partida
- Totales: `TOTAL SUBCAPÍTULO NN.MM ... <importe>`, `TOTAL CAPÍTULO NN ... <importe>`, `TOTAL <importe general>`

**Mapeo deficiencia ↔ subcapítulo** (en este orden de prioridad):
1. `caso.yaml > deficiencias[i].subcapitulo_presupuesto: "04.05"` (código directo). Es el que la skill rellena al fijar deficiencias o como parte de la conversación con Javier antes de presupuestar.
2. Match fuzzy entre el slug normalizado y el título del subcapítulo (sin acentos, sin guiones bajos).
3. Si no hay match único → marcador `[[SUBCAPÍTULO NO ENCONTRADO]]` y aviso al final.

**Conducta antes de invocar el script:**

1. Pedir a Javier que confirme que ha exportado el RTF a `99_salida/presupuesto/`.
2. Parsear el RTF con `presupuestar.py` (en modo dry-run conceptual: leerlo y mostrar la lista de subcapítulos detectados).
3. Para cada deficiencia que NO tenga `subcapitulo_presupuesto` en `caso.yaml`, presentar a Javier los subcapítulos del RTF y pedirle que asigne el código (`04.05`, etc.). Aplicar al `caso.yaml` y a `deficiencias.md`.
4. Ejecutar `presupuestar.py`:
   ```
   python scripts/presupuestar.py --encargo "<ruta>"
   ```

**Lo que hace el script:**
- Parsea el RTF, agrupa partidas por subcapítulo, calcula totales.
- Carga `informe_pericial_v1.docx`.
- Para cada `[[PRESUPUESTO_DEFICIENCIA: <slug>]]` del v1, lo sustituye por una **tabla** con las partidas del subcapítulo (Código, Ud, Descripción, Cantidad, Precio €, Importe €) más una fila final con `TOTAL SUBCAPÍTULO NN.MM`.
- Rellena la columna `Coste reparación` del cuadro resumen final con el total de cada deficiencia.
- Sustituye `[[HOJA_RESUMEN_PRESUPUESTO ...]]` por una tabla `Capítulo | Importe (€)` con el total de cada capítulo + fila `TOTAL GENERAL`.
- Guarda como `99_salida/informe_pericial_v2.docx`.

**Tras ejecutar**:
- Mostrar a Javier el resumen del script (subcapítulos inyectados, costes en cuadro, hoja resumen).
- Si quedan marcadores `[[CAUSA BREVE]]` o `[[IMPUTABLE A]]` en el cuadro resumen del v2, **redactarlos ahora conversacionalmente** (1 frase breve por deficiencia) y aplicar como Edit sobre el docx. Estos no los rellena el script porque dependen de la causa y la atribución de responsabilidad, que son criterio del perito.

### `/informe revisar` (o "aplica los cambios", "dame el v3")

Procesa una revisión hecha por Javier en Word. Tres modos no exclusivos:
1. **Tracked changes** (Control de cambios activado en Word): inserciones y borrados.
2. **Comentarios al margen** del docx: cada uno con anchor (texto al que apunta) y mensaje.
3. **Texto plano por chat**: "en deficiencia 3 cambia X por Y", "añade un párrafo a la conclusión sobre Z…".

**Arquitectura: separación inteligencia / maquinaria** (igual que `/informe redactar`).

#### Fase 1 — Procesamiento mecánico (script)

```
python scripts/revisar.py --encargo "<ruta>"
```

El script:
- Detecta el último `informe_pericial_vN.docx` en `99_salida/` (busca por sufijo numérico, no por mtime).
- **Acepta todos los tracked changes**: inserciones (`<w:ins>`) se mantienen como texto, borrados (`<w:del>`) se descartan, cambios de formato (`<w:rPrChange>`, `<w:pPrChange>`, `<w:cellIns>`, `<w:cellDel>`, `<w:moveFrom>`, `<w:moveTo>`, `<w:sectPrChange>`) se eliminan dejando el resultado.
- **Extrae todos los comentarios** con su anchor (texto subrayado en Word), autor, fecha y mensaje, y los guarda en:
  - `_skill_workspace/comentarios_revision.md` — formato legible para la skill.
  - `_skill_workspace/comentarios_revision.json` — formato programático.
- **Elimina los comentarios** del docx (purga `commentRangeStart/End`, `commentReference`, `word/comments.xml`, `word/people.xml`, content-types y rels asociadas).
- Guarda el resultado como `informe_pericial_v(N+1).docx`.

Por defecto la versión sale como N+1. Forzar otra con `--salida v5`.

#### Fase 2 — Aplicación de comentarios e instrucciones por chat (skill)

Tras el script, la skill (Claude conversacional):

0. **Carga memoria externa primero** (`notas_por_documento.md`, `cronologia.md`, `briefing.md`, `caso.yaml`, `deficiencias.md`). Si una corrección de Javier exige verificar un dato del expediente, consultar esos ficheros antes que los PDFs.

1. **Lee `comentarios_revision.md`** y lo presenta a Javier en pantalla, comentario por comentario.
2. **Para cada comentario**, propone la edición concreta a aplicar sobre el v(N+1).docx:
   - "Comentario #1 dice 'cambiar X por Y' sobre el anchor «...». Aplico `find/replace` de X→Y limitado a ese párrafo. ¿OK?"
   - Si el comentario es ambiguo o requiere criterio (p.ej. "ampliar este argumento"), redactar la propuesta antes de aplicarla.
3. **Aplica** las ediciones aprobadas con `python-docx` (Read del docx, modificar, Write). Para reemplazos simples, find/replace en runs; para inserciones de párrafos, `add_paragraph` en posición; para tablas, `cell.text = …`.
4. **Instrucciones por chat** (modo 3) se aplican igual: la skill identifica el target y edita el docx.
5. **Re-genera el cuadro resumen** si algún cambio afectó costes o títulos de deficiencias.
6. Si tras todo Javier quiere otra versión, vuelve a `/informe revisar` partiendo del v(N+1).

**Nota sobre tracked changes complejos**: tablas con `<w:cellIns>`/`<w:cellDel>` (inserción/borrado de filas) se aceptan como están. Si Word genera un cambio que el script no contempla y queda en el docx, abrir el v(N+1) en Word, "Aceptar todos" y guardar — el script no debería romper la apertura aunque le falte algún caso edge.

---

## Estructura de carpeta de encargo

```
YYYY_NNN_ApellidoCliente_Localidad/
├── caso.yaml
├── deficiencias.md                 ← fuente de verdad legible (la mantiene /informe deficiencias)
├── 00_encargo/
│   ├── hoja_encargo.pdf            (subido por Javier)
│   ├── objeto_pericial.md          (escrito por la skill en /informe nuevo)
│   └── ... (cualquier otra cosa relacionada con el encargo)
├── 01_documentacion/               ← FLEXIBLE: cualquier archivo y subcarpeta
│   ├── proyecto_visado.pdf
│   ├── contrato.pdf
│   ├── presupuestos_cliente/
│   ├── ACTUACIONES VILLA CARMEN.pdf
│   ├── Certificaciones y fras/
│   └── ... (todo lo que el cliente envíe + lo que recopiles)
├── 02_parte_contraria/             ← informes y alegaciones de la otra parte
├── 03_comunicaciones/              ← emails, burofax, whatsapp_export.txt
├── 04_inspeccion/
│   ├── fecha_inspeccion.txt
│   ├── notas_tablet/               ← Notability PDF (idealmente con OCR de manuscrito)
│   ├── fotos/                      ← copia local desde la carpeta de Drive del encargo (con EXIF)
│   └── fotos_renombradas/          ← P001_DEFNN_slug.jpg tras /informe ingestar
├── 05_ensayos/                     ← opcional
├── 06_referencias/                 ← normativa, jurisprudencia, fichas técnicas
├── 99_salida/
│   ├── presupuesto/                ← los 2 docx de Arquímedes
│   └── informe_pericial_vN.docx
└── _skill_workspace/               ← memoria externa persistente:
                                    inventario.md, briefing.md,
                                    notas_por_documento.md, cronologia.md,
                                    redaccion_deficiencias.md,
                                    documentacion_consultada.md,
                                    comentarios_revision.md, jpegs/, etc.
```

## Archivos de referencia

| Archivo | Cuándo |
|--------|--------|
| `reference/estructura_macro.md` | Antes de redactar v1 |
| `reference/plantilla_deficiencia.md` | Al redactar cada deficiencia |
| `reference/estilo_redaccion.md` | Siempre que generes prosa |
| `reference/caso_yaml_schema.yaml` | Al crear/leer caso.yaml |
| `boilerplate/juramento_objetividad.md` | Texto fijo, copia literal |
| `boilerplate/metodologia.md` | Texto con sustituciones de variables |
| `boilerplate/pie_firma.md` | Cierre del documento |
| `boilerplate/datos_autor.yaml` | Datos del autor para portada y pie |
| `plantillas/plantilla_base.docx` | Plantilla con estilos heredados de CAORZA |

## Reglas estrictas

1. **Nunca inventar fechas, nombres ni hechos.** Si falta un dato, dejar marcador `[[FALTA: <dato>]]` y avisar.
2. **Citas literales del proyecto entre comillas tipográficas «»** — nunca parafrasear como si fuera del proyecto.
3. **Toda foto del informe debe existir en `04_inspeccion/fotos_renombradas/`.**
4. **Si un toggle de sección está a false en `caso.yaml`, NO incluir esa sección.**
5. **El presupuesto NO se redacta**: viene de Arquímedes.
6. **`/informe nuevo` es interactivo y produce yaml COMPLETO.** No dejar campos "RELLENAR" salvo que Javier expresamente diga "lo añado luego".
7. **`01_documentacion/` acepta cualquier estructura.** No imponer nombres de archivo concretos.
