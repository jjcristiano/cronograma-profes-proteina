# 🏋️ Sistema Integral de Gestión de Horarios, Staff y Liquidaciones (v2.1)
## Proteína Gym · Proteína Box · Boxeo

Repositorio oficial del sistema administrativo centralizado para la gestión de cronogramas mensuales (`Noviembre 2026` a `Diciembre 2027`), administración de profesores, registro de reemplazos y liquidación automática de remuneraciones de las tres unidades de trabajo:
- **Proteína Gym** *(Cronograma General)*
- **Proteína Box** *(Horarios Box)*
- **Boxeo** *(Escuela y Entrenamiento de Boxeo)*

---

## 📂 Estructura de Archivos del Repositorio

| Archivo / Carpeta | Descripción |
| :--- | :--- |
| `SISTEMA_GESTION_STAFF_PROTEINA_v2.xlsx` | **Planilla Maestra (23 pestañas)** compatible con Google Sheets y Microsoft Excel, con todas las fórmulas, validaciones, feriados y meses de Nov 2026 a Dic 2027. |
| `GoogleAppsScript_Proteina.gs` | **Código de Google Apps Script** para pegar en `Extensiones -> Apps Script` dentro de Google Sheets. Activa el menú `🏋️ Sistema Proteína` (carga rápida de rangos horarios como `08:00 a 22:00`, registro de reemplazos Titular + Suplente, cambio de nombre de profesor en todo el historial, sincronización de `REGISTRO DE HORAS` y cierre mensual). |
| `webapp/index.html` | **Aplicación Web Administrativa Interactiva** lista para abrir en cualquier navegador (PC o celular), con chips multicolor, multiselección, carga rápida de jornada completa, alta/baja/renombrado de profesores, liquidaciones y exportación a PDF. |
| `INSTRUCTIVO_Y_MANUAL_DE_USO.md` | Manual paso a paso para los administradores y responsables del cronograma y de los pagos. |
| `HISTORIAL_CONVERSACION_Y_ESPECIFICACION.md` | Documento completo con la Especificación Funcional y Técnica v2.0, decisiones de diseño y resumen de la conversación de desarrollo. |
| `generar_sistema_proteina.py` | Script en Python (`openpyxl`) que genera y estructura el libro `SISTEMA_GESTION_STAFF_PROTEINA_v2.xlsx`. |
| `verificar_22_pruebas.py` | Script de auditoría y validación de las 22 pruebas obligatorias y de los importes mensuales por unidad y por profesor. |

---

## 🚀 Inicio Rápido

### Opción 1: Usar en Google Sheets
1. Sube `SISTEMA_GESTION_STAFF_PROTEINA_v2.xlsx` a tu Google Drive y ábrelo como **Hoja de cálculo de Google** (`Archivo -> Guardar como Hojas de cálculo de Google`).
2. Ve a **`Extensiones -> Apps Script`**, pega el contenido de `GoogleAppsScript_Proteina.gs`, guarda (`💾`) y recarga la página para habilitar el menú **`🏋️ Sistema Proteína`**.

### Opción 2: Usar la Web App Administrativa
1. Abre el archivo `webapp/index.html` directamente en tu navegador web (Chrome, Edge, Safari, móvil o escritorio).
