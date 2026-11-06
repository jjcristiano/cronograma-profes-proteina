# HISTORIAL DE CONVERSACIÓN, DECISIONES TÉCNICAS Y ESPECIFICACIÓN FUNCIONAL (v3.0 - Versión Web)
## Proyecto: `cronograma-profes-proteina` (Proteína Gym · Proteína Box · Boxeo)

---

## 1. Decisiones de Diseño Acordadas con el Usuario

1. **Modalidad 100% Web App (Descarte de Spreadsheets)**:
   - Se descartaron los archivos de Excel/Google Sheets (`.xlsx`, `.gs` y scripts `.py`) para mantener únicamente la **Aplicación Web Administrativa (`index.html` / `webapp/index.html`)** con persistencia automática en `localStorage`.
2. **Interfaz Simplificada para Usuarios +50 e Identidad Visual Proteína Gym**:
   - Integración del logo oficial (`LOGO_GYM_PROTEINA.PNG`) y paleta de colores institucional **Negro y Amarillo (`#0a0a0a` y `#f5e600`)** con alto contraste y tipografía legible.
   - Textos cortos y directos sin descripciones técnicas:
     - `Gestión de Cronogramas Mensuales`
     - Unidades: `1. Proteína Gym`, `2. Proteína Box`, `3. Boxeo`
     - `Carga Rápida de Turnos`
   - Elementos eliminados para evitar confusión:
     - Eliminado el botón de pre-cargar plantilla semanal.
     - Eliminados los botones de atajos rápidos de horario.
     - Eliminado el campo "Motivo de reemplazo" (se conserva la selección de Titular y Reemplazante).
     - Eliminada la columna y campo de "Unidades y Rol" en la sección `PROFES`.
3. **Listado Actualizado de Profesores (14 integrantes)**:
   - `Sil`
   - `Robert`
   - `Ceci`
   - `Jesi`
   - `Valentin`
   - `Agustin`
   - `Oscar`
   - `Irina`
   - `Luna`
   - `Cande`
   - `Agustin Castro`
   - `Alejandra Stefani`
   - `Fernanda Firpo`
   - `Nahuel`

---

## 2. Reglas de Liquidación Activas

- Cada bloque horario equivale a **30 minutos (`0,5 horas`)**.
- Si en una misma celda hay dos o más profesores (por compartir turno o por reemplazo `Titular + Reemplazante`), **cada uno recibe `0,5 horas` completas**.
- Lunes a sábado normal $\rightarrow$ Tarifa $\times 1$.
- Domingos y feriados $\rightarrow$ Tarifa $\times 2$ (si coincide domingo y feriado se paga $\times 2$, nunca cuádruple).
- Valor hora único para todos los profesores, configurable en `⚙️ VALOR HORA Y FERIADOS`.
