# INSTRUCTIVO DE USO PARA ADMINISTRADORES Y RESPONSABLES
## Sistema Integral de Gestión de Horarios, Staff y Liquidaciones (v2.0)
### Proteína Gym · Proteína Box · Boxeo

---

## 1. Archivos Entregados en el Proyecto

1. **`SISTEMA_GESTION_STAFF_PROTEINA_v2.xlsx`**: Archivo maestro con las **23 pestañas** listas para abrir en Microsoft Excel o subir directamente a **Google Sheets**.
2. **`GoogleAppsScript_Proteina.gs`**: Script de automatización para Google Sheets que habilita el menú personalizado **`🏋️ Sistema Proteína`** (chips multicolor, registro guiado de reemplazos con trazabilidad Titular/Suplente, pre-carga mensual desde la plantilla semanal, sincronización de `REGISTRO DE HORAS` y cierre de liquidaciones).
3. **`webapp/index.html`**: **Aplicación Web Administrativa Interactiva** lista para abrir en cualquier navegador (computadora o celular) con todos los módulos integrados.

---

## 2. Cómo Subir y Activar el Sistema en Google Sheets (En 2 Minutos)

1. Ingrese a **Google Drive** (`drive.google.com`) -> botón **Nuevo** -> **Subir archivo** y seleccione `SISTEMA_GESTION_STAFF_PROTEINA_v2.xlsx`.
2. Ábralo con **Hojas de cálculo de Google (Google Sheets)** y en el menú superior elija **Archivo -> Guardar como Hojas de cálculo de Google**.
3. Para activar el menú de automatizaciones **`🏋️ Sistema Proteína`**:
   - Vaya a **Extensiones -> Apps Script**.
   - Borre el contenido vacío que aparece por defecto y pegue el contenido íntegro del archivo `GoogleAppsScript_Proteina.gs`.
   - Presione el ícono **Guardar (💾)** y recargue la pestaña de Google Sheets.
   - En la barra superior aparecerá el menú **`🏋️ Sistema Proteína`**.

---

## 3. Operaciones Cotidianas Paso a Paso

### A. Asignar horarios en las 3 Unidades (`NOVIEMBRE 2026` a `DICIEMBRE 2027`)
- Ingrese a la pestaña del mes correspondiente.
- Cada hoja mensual está dividida en **3 bloques claramente identificados**:
  1. **Sección 1: Proteína Gym (Cronograma General)** (Filas 9 a 40)
  2. **Sección 2: Proteína Box (Horarios Box)** (Filas 44 a 75)
  3. **Sección 3: Boxeo** (Filas 79 a 110)
- Cada celda verde representa **30 minutos (`0,5 horas`)** para una fecha y unidad específicas.
- Si en una misma media hora trabajan dos o más personas (por ejemplo, profesor de sala + ventas, o dos profesores), selecciónelos o escríbalos separados por coma (ej.: `Robert, Silvina` o `Cecilia, Luna, Valentina`). **Cada integrante recibe automáticamente `0,5 horas` completas** (Reglas `R01`, `R02` y `R03`).

### A.2. ⚡ Cómo Cargar un Día Completo o Rango Horario de Una Sola Vez (ej. `08:00 a 22:00` sin ir celda por celda)
Tienes **3 maneras rápidas** de hacerlo sin marcar 28 celdas una por una:
1. **Desde el menú `🏋️ Sistema Proteína` en Google Sheets (`⚡ Asignar Rango Horario Rápido`)**:
   - Párate en la columna del día que quieras cargar, abre el menú **`🏋️ Sistema Proteína -> ⚡ Asignar Rango Horario Rápido (ej. 08:00 a 22:00)`**, escribe el nombre del profesor y el rango `08:00 a 22:00`. El sistema completa automáticamente las 28 celdas de media hora (pudiendo elegir si **suma** al profesor junto a los que ya estaban o si **reemplaza** esas horas).
2. **Arrastrando o Copiando/Pegando en Google Sheets / Excel**:
   - Selecciona al profesor en la celda de las `08:00 a 08:30`, toma el **puntito azul de la esquina inferior derecha de la celda** y arrástralo hacia abajo hasta las `21:30 a 22:00` (o presiona `Ctrl+C`, selecciona hasta las `22:00` y presiona `Ctrl+V`).
3. **En la Web App Administrativa (`webapp/index.html`)**:
   - Arriba del cronograma tienes la barra **"⚡ Carga Rápida de Turno Completo"**: eliges el profesor, el día (o *"Todos los Lunes del mes"*), presionas el botón **`🕗 Día Completo (08:00 a 22:00)`** y haces clic en **`✅ Aplicar Rango de Una Vez`**.
   - O bien, al hacer clic en la celda de las `08:00`, adentro de la misma ventana eliges en **"⚡ ¿Repetir de corrido en este día hasta las?:" -> `Hasta las 22:00 (14.0 hs de corrido)`** y presionas Guardar.

### B. Registrar un Reemplazo (Titular + Reemplazante)
- **Opción 1 (Directa en la celda)**: Agregue el nombre del profesor reemplazante junto al titular separado por coma (ej. `Robert, Agustín`). Ambos sumarán `0,5 horas` para la liquidación.
- **Opción 2 (Con trazabilidad de motivo)**: Ubíquese en la celda del horario y elija en el menú superior **`🏋️ Sistema Proteína -> 🔄 Registrar Reemplazo en Celda Activa`**. El sistema registrará a ambos en la celda y guardará una nota de auditoría detallando quién fue el **Profesor Titular**, quién fue el **Profesor Reemplazante** y el motivo del reemplazo.

### C. Incorporar un Profesor Nuevo, Cambiarle el Nombre o Desactivar una Baja (`PROFES`)
- **Alta nueva**: Vaya a la pestaña **`PROFES`**, escriba el nombre del nuevo profesor en la siguiente fila verde vacía y seleccione el estado **`Activo`**. Automáticamente aparecerá en las listas desplegables de todos los meses y en todos los reportes mensuales, anuales y por unidad sin tocar ninguna fórmula (Regla `R10`).
- **✏️ Cambiar el nombre de un profesor existente**:
  - **En Google Sheets**: Use el menú **`🏋️ Sistema Proteína -> ✏️ Cambiar Nombre de un Profesor`**. Ingrese el nombre actual (ej. `Robert`) y el nuevo nombre (ej. `Roberto`): el sistema lo cambia en la pestaña `PROFES` **y actualiza automáticamente todas las celdas de los 14 meses y la plantilla semanal** para que no pierda ni media hora ya cargada.
  - **En la Web App**: En la pestaña **`👥 PROFES`**, haga clic en el botón **`✏️ Cambiar Nombre`** al lado del profesor; se actualizará en toda la base de datos al instante.
- **Baja / Desactivación**: Nunca borre la fila de un profesor que ya trabajó. Simplemente cambie su columna **Estado** de `Activo` a **`Inactivo`**. Dejará de aparecer en las listas desplegables para nuevas asignaciones (`R11`), pero conservará intacto su historial y sus liquidaciones de meses anteriores (`R12`).

### D. Cambiar el Valor de la Hora (`CONFIGURACIÓN`)
- Vaya a la pestaña **`CONFIGURACIÓN`** y modifique la celda azul **`B6`** (*Valor de la Hora Normal*).
- Automáticamente se recalculará el valor de la hora doble (`×2`) en `B7` y se actualizarán todos los meses que estén en estado **`ABIERTO`**, sin alterar los meses anteriores que ya hayan sido marcados como **`LIQUIDADO`** (`R09`, `R13`).

### E. Agregar Feriados (`FERIADOS`)
- Vaya a la pestaña **`FERIADOS`** y agregue la fecha en la siguiente fila verde vacía.
- Todos los cronogramas detectan automáticamente tanto los **Domingos** como las fechas listadas en **`FERIADOS`** y les aplican el multiplicador **`×2` (Doble)**. Si un feriado cae en día domingo (ej. `20/06/2027`), el sistema aplica `×2` y nunca `×4` (Regla `R08`).

### F. Consultar Pagos y Cerrar el Mes (`RESUMEN MENSUAL` y `REPORTES POR UNIDAD`)
- En **`RESUMEN MENSUAL`**, elija el mes en la celda verde **`C4`**.
- Verá al instante:
  1. **Liquidación General Consolidada por Profesor** (sumando sus horas de Proteína Gym, Proteína Box y Boxeo sin duplicar pagos).
  2. **Desglose de Horas y Costos por Unidad** (cuánto representa económicamente cada unidad y su porcentaje sobre el total).
  3. **Ficha Individual por Profesor** (seleccionando al profesor en la tabla 3 para imprimir o exportar su recibo en PDF).
- Una vez revisado y pagado el mes, cambie su estado a **`LIQUIDADO`** en `CONFIGURACIÓN` (o desde el menú **`🏋️ Sistema Proteína -> 🔒 Cerrar / Congelar Liquidación del Mes`**) para blindar los importes históricos.
