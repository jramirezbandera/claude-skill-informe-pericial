# Plantilla por deficiencia (sub-estructura obligatoria)

Cada deficiencia se redacta como una subsección de "DEFICIENCIAS OBSERVADAS" con un título descriptivo en Heading 2 (en mayúsculas), y dentro lleva 5 sub-apartados en Heading 3 (en este orden):

## 1. Existencia

Descripción de lo constatado in situ. Empezar siempre con una fórmula del tipo:

> "Se ha podido constatar mediante visita al inmueble la existencia de…"
> "Durante la visita realizada el día {{fecha_visita}}, se observó…"

Mencionar el alcance material (extensión, ubicación, número de elementos afectados). Insertar foto(s) del móvil con caption "Imagen NN: <breve descripción>".

Si la deficiencia tiene riesgo añadido (caída de cascotes, accesibilidad, fugas con daños a terceros…), añadir un párrafo final advirtiendo del riesgo.

## 2. Definición en proyecto y mediciones

Citar literalmente del proyecto o memoria de calidades cuando se disponga, entrecomillado con comillas tipográficas («…» o "…"). Indicar:

- Plano o página de la memoria donde aparece (si se conoce).
- Partida de mediciones que afecte al elemento (código, descripción larga textual entre comillas).

Si NO hay proyecto visado, sustituir la fórmula por:

> "No se dispone de proyecto de ejecución visado. La definición contractual del elemento se extrae de [memoria de calidades anexa al contrato / infografía comercial / declaración del propietario]…"

Si sí hay contrato pero no proyecto técnico, citar el contrato.

## 3. Causa

Análisis técnico de por qué se ha producido la deficiencia. Estructura habitual:

- Descripción del mecanismo correcto de ejecución/funcionamiento.
- Qué se hizo mal o no se hizo (con concreción: defecto de ejecución, mantenimiento, diseño, uso, alteración posterior…).
- Si hay evidencia documental (libro de órdenes, mensajes WhatsApp con DF/constructora, actas), citarla. Insertar capturas como ilustración.

Diferenciar siempre entre defecto de ejecución, defecto de proyecto, defecto de mantenimiento y deterioro normal por uso.

## 4. Certificación

OPCIONAL — incluir sólo si se dispone de las certificaciones de obra y la deficiencia se corresponde con una unidad certificada y abonada.

> "En la certificación nº {{n}}, de fecha {{fecha}}, se certificó el {{capítulo/subcapítulo}} {{descripción}}, por lo que fueron abonados los trabajos realizados de…"

Si no aplica, OMITIR este sub-apartado completo (no dejar título vacío).

## 5. Propuesta de reparación

Dos partes:

**5.1 — Descripción técnica de la reparación.** Cómo debe abordarse: secuencia de trabajos, criterios constructivos, materiales, advertencias. Si la solución implica cambios estéticos o funcionales, justificar la proporcionalidad respecto al daño.

**5.2 — Valoración económica (FASE 2).** En el v1 redactado por la skill, este bloque queda como marcador:

```
[[PRESUPUESTO_DEFICIENCIA: <slug-deficiencia>]]
```

En la fase 2 (subcomando `/informe presupuestar`), la skill localiza el capítulo correspondiente en el docx de desglose exportado de Arquímedes y lo sustituye en este punto.

Cuando una deficiencia es sólo descripción técnica sin valoración, sustituir el marcador por:

> "Esta deficiencia se valora como reparación técnica sin partidas presupuestarias específicas, por integrarse dentro de los trabajos generales de mantenimiento de la propiedad."

## Mapeo deficiencia ↔ capítulo de presupuesto

Para que la fase 2 funcione, la convención es:

- Cada deficiencia tiene un `slug` corto (ej. `humedades_garaje`).
- Cuando el usuario monta el presupuesto en Presto/Arquímedes, cada CAPÍTULO debe nombrarse con el mismo slug en mayúsculas (ej. `CAPÍTULO 03 — HUMEDADES GARAJE`).
- La skill detecta el match por nombre de capítulo y lo inyecta en el marcador correspondiente.

Si una deficiencia no tiene capítulo en el presupuesto, la skill lo notifica como warning y deja el marcador con texto "PENDIENTE DE VALORACIÓN".
