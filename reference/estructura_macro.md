# Estructura macro del informe pericial

Orden y obligatoriedad de cada sección:

| # | Sección | Obligatoria | Activación |
|---|---------|-------------|-----------|
| 1 | CONSIDERACIONES PREVIAS Y OBJETO DE TRABAJO | Sí | Siempre |
| 2 | JURAMENTO DE ACTUAR CON OBJETIVIDAD | Sí | Siempre |
| 3 | BASES DEL DICTAMEN. VISITA E INSPECCIÓN | Sí | Siempre |
|   |  · Visita de inspección | Sí | Siempre |
|   |  · Documentación examinada | Sí | Siempre |
|   |  · Alcance de esta base documental | Sí | Siempre |
| 4 | ANTECEDENTES (Y DESCRIPCIÓN DEL CONJUNTO) | Sí | Siempre |
|   |  · Situación | Sí | Siempre |
|   |  · Fechas de obra y agentes que intervinieron | Sí | Siempre |
|   |  · Ámbito de la prueba pericial solicitada | Sí | Siempre |
|   |  · Tipología y entorno | Sí | Siempre |
|   |  · Delimitación del ámbito pericial | Opcional | Cuando hay zonas dudosas |
| 5 | METODOLOGÍA DE ANÁLISIS | Opcional | `caso.yaml` → secciones_a_incluir.metodologia |
| 6 | DESARROLLO ANORMAL DE LA OBRA | Opcional | Sólo si hubo incidencias documentadas en obra (libro de órdenes, actas) |
| 7 | DEFICIENCIAS OBSERVADAS / DEFECTOS RECLAMADOS | Sí | Siempre. Subsección por deficiencia (ver `plantilla_deficiencia.md`) |
| 8 | ANÁLISIS CRÍTICO DEL INFORME DE LA ACTORA | Opcional | Sólo si existe `02_parte_contraria/informe_*.pdf` |
| 9 | PENALIZACIONES Y RETRASO | Opcional | `caso.yaml` → secciones_a_incluir.penalizaciones_retraso |
| 10 | CUADRO RESUMEN Y CONCLUSIONES | Sí | Siempre |
| 11 | HOJA RESUMEN DE PRESUPUESTO | Opcional | `caso.yaml` → secciones_a_incluir.hoja_resumen_presupuesto |
| 12 | PIE DE FIRMA | Sí | Siempre |

## Numeración

- Encabezados de sección 1 a 12 → estilo Word "Heading 1"
- Subsecciones (ej. 4.1, 4.2…) → "Heading 2"
- Sub-apartados de cada deficiencia (ej. 7.3.1 Existencia) → "Heading 3"

El estilo de la plantilla base ya tiene definidos Heading 1/2/3. No usar otros niveles.

## Comportamiento ante secciones omitidas

Si una sección opcional no se incluye, NO renumerar visualmente con saltos: Word genera la numeración desde la plantilla. La sección simplemente no se inserta. La tabla de contenido debe regenerarse al abrir el .docx (Word avisa).

## Tono general por sección

- Consideraciones previas, Juramento, Bases, Antecedentes, Metodología → tono formal, expositivo, tercera persona impersonal ("se ha podido constatar", "el suscrito manifiesta", "la documentación examinada permite…").
- Cada deficiencia → tono técnico-descriptivo, neutral. Evitar adjetivos valorativos salvo en el sub-apartado de Causa cuando se atribuya responsabilidad.
- Cuadro resumen y conclusiones → frases breves, directas. Lista numerada de hallazgos.
