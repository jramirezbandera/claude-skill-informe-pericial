---
description: Subcomandos de la skill informe-pericial (nuevo, ingestar, deficiencias, redactar, presupuestar, revisar, checklist).
argument-hint: "<subcomando> [extras]"
---

Activa la skill `informe-pericial` y ejecuta el subcomando solicitado.

**Subcomando solicitado:** `$ARGUMENTS`

## Qué hacer

1. **Lee primero** `C:\Users\el usuario\.claude\skills\informe-pericial\SKILL.md` para el contexto completo del workflow (si no lo has leído ya en esta conversación).

2. **Identifica el subcomando** dentro de `$ARGUMENTS`. Subcomandos válidos:
   - `nuevo` — **3 preguntas mínimas** (apellido, localidad, código YYYY_NNN). Crea carpeta + `caso.yaml` skeleto. NADA MÁS. El resto del yaml se rellena en `ingestar`.
   - `ingestar` — **BRIEFING corto (5 preguntas) → ingesta autónoma de PDFs + fotos desde Drive + notas Notability → resumen ✅/⚠️/❌ y confirmación final de huecos**. Genera `objeto_pericial.md` y borrador de `deficiencias.md`.
   - `deficiencias` — fija la lista de deficiencias (modo dictado, edición manual, o aceptar propuesta)
   - `redactar` — genera v1 del informe (sin presupuestos)
   - `presupuestar` — inyecta los capítulos de Arquímedes en cada deficiencia → v2
   - `revisar` — aplica cambios sobre vN → vN+1
   - `checklist` — muestra estado del encargo (✅/⚠️/❌)
   - `fotos` — atajo a la fase de fotos de `ingestar` (sólo procesar el volcado)
   - `notas` — atajo a la fase de notas Notability de `ingestar`

3. **Si `$ARGUMENTS` está vacío o no coincide con ningún subcomando**, muestra esta lista al usuario y pregunta cuál quiere ejecutar. NO asumas.

4. **Ejecuta el subcomando** siguiendo al pie de la letra la sección correspondiente de `SKILL.md`. No improvises — la skill define cada paso (qué scripts ejecutar, qué preguntar, qué escribir).

5. **Argumentos extra**: si después del nombre del subcomando hay más texto (p.ej. `/informe nuevo García López Mijas`), trátalo como contexto inicial que adelanta respuestas del bloque interactivo. No es obligatorio.

## Recordatorio de paths

- Skill: `C:\Users\el usuario\.claude\skills\informe-pericial\`
- Encargos: subcarpeta del cwd actual (NO hay raíz fija)
- Drive (fotos): `G:\Mi unidad\...` (la ruta concreta del encargo la pregunta el subcomando `ingestar`)
