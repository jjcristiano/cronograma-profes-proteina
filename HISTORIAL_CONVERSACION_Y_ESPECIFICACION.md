# HISTORIAL DE CONVERSACIÓN, DECISIONES TÉCNICAS Y ESPECIFICACIÓN FUNCIONAL (v2.1)
## Proyecto: `cronograma-profes-proteina` (Proteína Gym · Proteína Box · Boxeo)

> **Nota para continuar en otra sesión de Antigravity:**
> Puedes referenciar este archivo con `@HISTORIAL_CONVERSACION_Y_ESPECIFICACION.md` o mencionar la conversación anterior usando `@` en el chat de Antigravity (`82232c32-b986-4759-bcbd-a67e53a0acbc`).

---

## 1. Decisiones de Diseño Acordadas con el Usuario

1. **Modalidad de implementación**:
   - Entrega combinada de:
     1. Archivo `.xlsx` completo listo para Google Sheets (`SISTEMA_GESTION_STAFF_PROTEINA_v2.xlsx`).
     2. Automatizaciones en Google Apps Script (`GoogleAppsScript_Proteina.gs`).
     3. Aplicación Web Administrativa interactiva (`webapp/index.html`).
2. **Estructura de las 3 Unidades en cada Pestaña Mensual (`NOVIEMBRE 2026` a `DICIEMBRE 2027`)**:
   - Cada hoja mensual contiene **3 bloques de calendario día por día (del 1 al 31)**:
     - **Sección 1: Proteína Gym** *(Cronograma General)*
     - **Sección 2: Proteína Box** *(Horarios Box)*
     - **Sección 3: Boxeo**
   - Incluye además la pestaña **`PLANTILLA SEMANAL`** con los turnos base de Lunes a Domingo para pre-cargar cualquier mes con un clic.
3. **Unificación del Staff en `PROFES` y Criterio de Pago de Ventas**:
   - Todos los integrantes figuran una sola vez en la pestaña **`PROFES`**, separando `"Rob/sil"` en **Robert** y **Silvina**.
   - Lista inicial de 13 profesores activos con colores asignados: `Agustín`, `Robert`, `Silvina`, `Valentina`, `Jessi`, `Luna`, `Irina`, `Cecilia`, `Oscar`, `Yami`, `Lucas`, `Fede` y `Martín`.
   - Las horas de `VENTAS` también se liquidan al mismo valor hora único dentro de su unidad correspondiente.
4. **Reemplazos y Selección Múltiple (`0,5 hs` completas para cada profesor)**:
   - Cada bloque horario es de **30 minutos (`0,5 hs`)**.
   - Si en una misma celda hay dos o más profesores (por compartir franja o por reemplazo `Titular + Reemplazante`), **cada uno recibe `0,5 hs` completas** (Reglas `R01`, `R02`, `R03`).
   - Se incluye además un botón/menú específico para registrar el detalle de **Profesor Titular vs. Profesor Reemplazante** y el motivo del cambio.
5. **Funciones Rápidas Solicitadas Posteriormente**:
   - **Cambiar el nombre de un profesor (`✏️ Cambiar Nombre`)**: Actualiza el nombre en `PROFES` y reemplaza automáticamente todas sus apariciones en los 14 cronogramas mensuales, la plantilla semanal y el historial de reemplazos.
   - **Carga rápida de día completo o rango horario (`⚡ Asignar Rango Horario Rápido`, ej. `08:00 a 22:00`)**: Permite asignar a un profesor en todo un rango de horas de una sola vez (sumándolo a quienes ya estén o reemplazando el rango) sin tener que ir celda por celda.

---

## 2. Reglas Obligatorias Implementadas (`R01` a `R16`)

- **R01**: Cada celda representa 30 minutos (`0,5 horas`).
- **R02**: Un profesor en una celda $\rightarrow$ `0,5 horas` para ese profesor.
- **R03**: Dos o más profesores en una celda $\rightarrow$ `0,5 horas` para cada uno.
- **R04**: Cada asignación identifica obligatoriamente su unidad (`Proteína Gym`, `Proteína Box` o `Boxeo`).
- **R05**: Lunes a sábado normal $\rightarrow$ Tarifa $\times 1$.
- **R06**: Domingo $\rightarrow$ Tarifa $\times 2$.
- **R07**: Feriado (en pestaña `FERIADOS`) $\rightarrow$ Tarifa $\times 2$.
- **R08**: Domingo y feriado simultáneos (ej. `20/06/2027`) $\rightarrow$ Tarifa $\times 2$ (nunca cuádruple).
- **R09**: Valor hora único para todos los profesores (configurable en `CONFIGURACIÓN!B6`, inicializado en `\$5.000`).
- **R10**: Nuevo profesor en `PROFES` se incorpora automáticamente a las listas desplegables y reportes.
- **R11**: Profesor en estado `Inactivo` deja de aparecer en las listas desplegables para nuevas asignaciones.
- **R12**: Al desactivar un profesor se conserva todo su historial y liquidaciones anteriores.
- **R13**: Liquidación cerrada (`LIQUIDADO`) conserva su tarifa congelada y advierte ante intentos de modificación.
- **R14**: Si un profesor trabaja en varias unidades (ej. `Valentina`, `Robert`, `Fede`), se consolidan sus horas sin duplicar pagos.
- **R15**: Reporte por unidad muestra horas normales, dobles y costos de cada sector por separado.
- **R16**: Reporte general suma exactamente las tres unidades sin duplicaciones.

---

## 3. Validación de Importes — Mes Base: Noviembre 2026 (Tarifa `\$5.000/h`)

- **Proteína Gym**: `822,0 hs` normales + `51,0 hs` dobles = **`873,0 hs`** $\rightarrow$ **`\$4.620.000`**
- **Proteína Box**: `109,0 hs` normales + `3,0 hs` dobles = **`112,0 hs`** $\rightarrow$ **`\$575.000`**
- **Boxeo**: `78,0 hs` normales + `3,5 hs` dobles = **`81,5 hs`** $\rightarrow$ **`\$425.000`**
- **TOTAL CONSOLIDADO (3 UNIDADES)**: `1.009,0 hs` normales + `57,5 hs` dobles = **`1.066,5 hs`** $\rightarrow$ **`\$5.620.000`**
