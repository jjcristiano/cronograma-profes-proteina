/**
 * ============================================================================
 * SISTEMA INTEGRAL DE GESTIÓN DE HORARIOS, STAFF Y LIQUIDACIONES (v2.1)
 * Proteína Gym · Proteína Box · Boxeo
 * ============================================================================
 * Instrucciones de instalación en Google Sheets:
 * 1. Abra el archivo SISTEMA_GESTION_STAFF_PROTEINA_v2.xlsx en Google Sheets
 *    (Archivo -> Guardar como Hojas de cálculo de Google).
 * 2. Vaya al menú Extensiones -> Apps Script.
 * 3. Pegue todo este código en el archivo Código.gs y presione Guardar (💾).
 * 4. Recargue la hoja de cálculo: aparecerá el menú "🏋️ Sistema Proteína".
 * ============================================================================
 */

const MESES_SISTEMA = [
  "NOVIEMBRE 2026", "DICIEMBRE 2026", "ENERO 2027", "FEBRERO 2027",
  "MARZO 2027", "ABRIL 2027", "MAYO 2027", "JUNIO 2027",
  "JULIO 2027", "AGOSTO 2027", "SEPTIEMBRE 2027", "OCTUBRE 2027",
  "NOVIEMBRE 2027", "DICIEMBRE 2027"
];

const BLOQUES_UNIDADES = [
  { nombre: "Proteína Gym", seccion: "General", filaInicio: 9, filaFin: 40, filaPlantilla: 6 },
  { nombre: "Proteína Box", seccion: "Box", filaInicio: 44, filaFin: 75, filaPlantilla: 41 },
  { nombre: "Boxeo", seccion: "Boxeo", filaInicio: 79, filaFin: 110, filaPlantilla: 76 }
];

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("🏋️ Sistema Proteína")
    .addItem("⚡ Asignar Rango Horario Rápido (ej. 08:00 a 22:00)", "asignarRangoHorarioRapido")
    .addItem("🔄 Registrar Reemplazo en Celda Activa (Titular + Suplente)", "registrarReemplazoCeldaActiva")
    .addItem("✏️ Cambiar Nombre de un Profesor (Actualiza todo el historial)", "renombrarProfesorEnTodoElSistema")
    .addSeparator()
    .addItem("📅 Pre-cargar Mes Activo desde Plantilla Semanal", "precargarMesDesdePlantilla")
    .addItem("🎨 Actualizar Listas Desplegables y Colores de Profes", "actualizarValidacionesYColores")
    .addItem("🗄️ Sincronizar Base de Datos (REGISTRO DE HORAS)", "sincronizarRegistroDeHoras")
    .addSeparator()
    .addItem("🔒 Cerrar / Congelar Liquidación del Mes en Consulta", "cerrarLiquidacionMensual")
    .addItem("🛡️ Proteger Fórmulas y Celdas Automáticas", "protegerFormulasSistema")
    .addToUi();
}

/**
 * Advierte si el usuario edita un mes ya marcado como LIQUIDADO (Regla R13)
 */
function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  const mesIdx = MESES_SISTEMA.indexOf(sheetName);
  if (mesIdx === -1) return;

  const row = e.range.getRow();
  const col = e.range.getColumn();
  if (col < 2 || row < 9 || row > 110) return;

  const ss = e.source;
  const cfgSheet = ss.getSheetByName("CONFIGURACIÓN");
  if (!cfgSheet) return;

  const estadoMes = cfgSheet.getRange(14 + mesIdx, 2).getValue();
  if (estadoMes === "LIQUIDADO") {
    SpreadsheetApp.getUi().alert(
      "⚠️ Período Cerrado (LIQUIDADO)",
      "El mes de " + sheetName + " ya fue cerrado y liquidado (Regla R13).\n" +
      "Si necesita realizar un ajuste retroactivo autorizado, cambie primero el estado a ABIERTO en la pestaña CONFIGURACIÓN.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * NUEVO 1: Asignar un profesor en todo un rango horario de una sola vez (ej. 08:00 a 22:00)
 * sin tener que ir celda por celda.
 */
function asignarRangoHorarioRapido() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const sheetName = sheet.getName();

  if (MESES_SISTEMA.indexOf(sheetName) === -1) {
    ui.alert("Abra primero la pestaña del mes donde desea cargar el horario (ej. NOVIEMBRE 2026) y ubíquese en la columna del día deseado.");
    return;
  }

  const activeCell = sheet.getActiveCell();
  const col = activeCell.getColumn() >= 2 ? activeCell.getColumn() : 2;
  const row = activeCell.getRow();

  // Detectar unidad según dónde esté parado el usuario (por defecto Proteína Gym)
  let bloqueActivo = BLOQUES_UNIDADES[0];
  for (let i = 0; i < BLOQUES_UNIDADES.length; i++) {
    if (row >= BLOQUES_UNIDADES[i].filaInicio && row <= BLOQUES_UNIDADES[i].filaFin) {
      bloqueActivo = BLOQUES_UNIDADES[i];
      break;
    }
  }

  const fechaCol = sheet.getRange(5, col).getDisplayValue();
  const diaCol = sheet.getRange(4, col).getDisplayValue();

  const pProf = ui.prompt(
    "⚡ Asignación Rápida por Rango Horario (" + bloqueActivo.nombre + ")",
    "Día seleccionado: " + diaCol + " " + fechaCol + "\n\n" +
    "1. Ingrese el nombre del Profesor (o varios separados por coma, ej: Agustín):",
    ui.ButtonSet.OK_CANCEL
  );
  if (pProf.getSelectedButton() !== ui.Button.OK) return;
  const profInput = pProf.getResponseText().trim();
  if (!profInput) return;

  const pRango = ui.prompt(
    "⏰ Rango de Horario (ej. 08:00 a 22:00)",
    "Ingrese la hora de inicio y fin separadas por 'a' o guion (ejemplos: '08:00 a 22:00' o '8 a 22'):",
    ui.ButtonSet.OK_CANCEL
  );
  if (pRango.getSelectedButton() !== ui.Button.OK) return;
  const textoRango = pRango.getResponseText().trim().toLowerCase().replace("-", " a ");
  const partes = textoRango.split("a").map(function(s) { return s.trim(); });
  if (partes.length < 2) {
    ui.alert("Formato inválido. Escriba por ejemplo: 08:00 a 22:00");
    return;
  }

  function parseHoraMin(str) {
    const p = str.split(":");
    const h = parseInt(p[0], 10);
    const m = p.length > 1 ? parseInt(p[1], 10) : 0;
    return h * 60 + m;
  }

  const minInicio = parseHoraMin(partes[0]);
  const minFin = parseHoraMin(partes[1]);
  if (isNaN(minInicio) || isNaN(minFin) || minFin <= minInicio) {
    ui.alert("Verifique el horario ingresado (ejemplo válido: 08:00 a 22:00).");
    return;
  }

  const modoResp = ui.alert(
    "¿Cómo desea aplicar el rango?",
    "• Presione SÍ para SUMAR a '" + profInput + "' junto a los profesores que ya estén en esas celdas.\n" +
    "• Presione NO para REEMPLAZAR el contenido de esas celdas dejando únicamente a '" + profInput + "'.",
    ui.ButtonSet.YES_NO_CANCEL
  );
  if (modoResp === ui.Button.CANCEL) return;
  const sumarAExistentes = (modoResp === ui.Button.YES);

  const nuevosProfes = profInput.split(",").map(function(s) { return s.trim(); }).filter(Boolean);
  const valoresCol = sheet.getRange(bloqueActivo.filaInicio, col, 32, 1).getValues();
  let bloquesModificados = 0;

  for (let s = 0; s < 32; s++) {
    const slotStartMin = (7 * 60) + (s * 30);
    const slotEndMin = slotStartMin + 30;
    if (slotStartMin >= minInicio && slotEndMin <= minFin) {
      if (sumarAExistentes) {
        const actuales = String(valoresCol[s][0] || "")
          .split(",")
          .map(function(x) { return x.trim(); })
          .filter(Boolean);
        nuevosProfes.forEach(function(np) {
          if (actuales.indexOf(np) === -1) actuales.push(np);
        });
        valoresCol[s][0] = actuales.join(", ");
      } else {
        valoresCol[s][0] = nuevosProfes.join(", ");
      }
      bloquesModificados++;
    }
  }

  sheet.getRange(bloqueActivo.filaInicio, col, 32, 1).setValues(valoresCol);
  ui.alert(
    "✅ Rango Horario Aplicado",
    "Se cargó a '" + nuevosProfes.join(", ") + "' en " + bloquesModificados + " bloques de 30 min (" +
    (bloquesModificados * 0.5).toFixed(1) + " horas) el " + diaCol + " " + fechaCol + " en " + bloqueActivo.nombre + ".",
    ui.ButtonSet.OK
  );
}

/**
 * NUEVO 2: Cambiar el nombre de un profesor en PROFES y actualizar automáticamente
 * todas sus apariciones en PLANTILLA SEMANAL y en los 14 cronogramas mensuales.
 */
function renombrarProfesorEnTodoElSistema() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const respViejo = ui.prompt(
    "✏️ Cambiar Nombre de un Profesor (Paso 1 de 2)",
    "Escriba el nombre ACTUAL del profesor que desea cambiar (tal como figura hoy, ej: Robert):",
    ui.ButtonSet.OK_CANCEL
  );
  if (respViejo.getSelectedButton() !== ui.Button.OK) return;
  const nombreActual = respViejo.getResponseText().trim();
  if (!nombreActual) return;

  const respNuevo = ui.prompt(
    "✏️ Cambiar Nombre de un Profesor (Paso 2 de 2)",
    "Escriba el NUEVO nombre para '" + nombreActual + "' (ej: Roberto):",
    ui.ButtonSet.OK_CANCEL
  );
  if (respNuevo.getSelectedButton() !== ui.Button.OK) return;
  const nombreNuevo = respNuevo.getResponseText().trim();
  if (!nombreNuevo || nombreNuevo === nombreActual) return;

  const profSheet = ss.getSheetByName("PROFES");
  const datosProfes = profSheet.getRange("A5:C39").getValues();
  let encontradoEnProfes = false;

  for (let i = 0; i < datosProfes.length; i++) {
    if (String(datosProfes[i][0]).trim().toLowerCase() === nombreActual.toLowerCase()) {
      profSheet.getRange(5 + i, 1).setValue(nombreNuevo);
      profSheet.getRange(5 + i, 3).setValue(nombreNuevo);
      encontradoEnProfes = true;
      break;
    }
  }

  if (!encontradoEnProfes) {
    ui.alert("⚠️ No se encontró a '" + nombreActual + "' en la pestaña PROFES. Verifique cómo está escrito.");
    return;
  }

  // Función auxiliar para reemplazar el nombre exacto dentro de celdas simples o con múltiples nombres ("Robert, Agustín")
  function reemplazarEnRango(rng) {
    const vals = rng.getValues();
    let cambios = 0;
    for (let r = 0; r < vals.length; r++) {
      for (let c = 0; c < vals[r].length; c++) {
        const txt = String(vals[r][c] || "").trim();
        if (!txt) continue;
        const lista = txt.split(",").map(function(item) {
          const limpio = item.trim();
          if (limpio.toLowerCase() === nombreActual.toLowerCase()) {
            cambios++;
            return nombreNuevo;
          }
          return limpio;
        });
        vals[r][c] = lista.join(", ");
      }
    }
    if (cambios > 0) rng.setValues(vals);
    return cambios;
  }

  let totalCeldasActualizadas = 0;
  const shPlantilla = ss.getSheetByName("PLANTILLA SEMANAL");
  if (shPlantilla) {
    totalCeldasActualizadas += reemplazarEnRango(shPlantilla.getRange("B6:H108"));
  }

  MESES_SISTEMA.forEach(function(nombreMes) {
    const sh = ss.getSheetByName(nombreMes);
    if (sh) {
      totalCeldasActualizadas += reemplazarEnRango(sh.getRange("B9:AF40"));
      totalCeldasActualizadas += reemplazarEnRango(sh.getRange("B44:AF75"));
      totalCeldasActualizadas += reemplazarEnRango(sh.getRange("B79:AF110"));
    }
  });

  actualizarValidacionesYColores();

  ui.alert(
    "✅ Nombre Actualizado en Todo el Sistema",
    "Se cambió '" + nombreActual + "' por '" + nombreNuevo + "' en la pestaña PROFES y se actualizaron automáticamente " +
    totalCeldasActualizadas + " celdas de horarios en los cronogramas y plantillas.",
    ui.ButtonSet.OK
  );
}

/**
 * 1. Registrar un reemplazo conservando al Profesor Titular y al Reemplazante
 *    en la misma franja horaria (0,5 hs para cada uno según Regla R03 y Sección 7).
 */
function registrarReemplazoCeldaActiva() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const sheetName = sheet.getName();

  if (MESES_SISTEMA.indexOf(sheetName) === -1) {
    ui.alert("Seleccione primero una celda de horario dentro de una pestaña mensual (ej. NOVIEMBRE 2026).");
    return;
  }

  const cell = sheet.getActiveCell();
  const row = cell.getRow();
  const col = cell.getColumn();

  let unidadDetectada = null;
  for (let i = 0; i < BLOQUES_UNIDADES.length; i++) {
    if (row >= BLOQUES_UNIDADES[i].filaInicio && row <= BLOQUES_UNIDADES[i].filaFin) {
      unidadDetectada = BLOQUES_UNIDADES[i];
      break;
    }
  }

  if (!unidadDetectada || col < 2) {
    ui.alert("Por favor ubíquese sobre una celda de asignación horaria (columnas de fecha, dentro de Proteína Gym, Proteína Box o Boxeo).");
    return;
  }

  const valorActual = String(cell.getValue() || "").trim();
  const horario = sheet.getRange(row, 1).getDisplayValue();
  const fecha = sheet.getRange(5, col).getDisplayValue();

  const respTitular = ui.prompt(
    "🔄 Registro de Reemplazo (" + unidadDetectada.nombre + " - " + fecha + " " + horario + ")",
    "Profesor Titular (actual en celda: '" + (valorActual || "ninguno") + "'):",
    ui.ButtonSet.OK_CANCEL
  );
  if (respTitular.getSelectedButton() !== ui.Button.OK) return;
  const titular = respTitular.getResponseText().trim() || valorActual;

  const respSuplente = ui.prompt(
    "🔄 Profesor Reemplazante",
    "Ingrese el nombre del Profesor Reemplazante (ambos sumarán 0,5 hs según Regla R03):",
    ui.ButtonSet.OK_CANCEL
  );
  if (respSuplente.getSelectedButton() !== ui.Button.OK) return;
  const suplente = respSuplente.getResponseText().trim();
  if (!suplente) return;

  const respMotivo = ui.prompt(
    "📝 Motivo del Reemplazo (Opcional)",
    "Indique brevemente el motivo del reemplazo para el historial:",
    ui.ButtonSet.OK_CANCEL
  );
  const motivo = respMotivo.getSelectedButton() === ui.Button.OK ? respMotivo.getResponseText().trim() : "";

  let listaActual = valorActual ? valorActual.split(",").map(function(s) { return s.trim(); }).filter(Boolean) : [];
  if (titular && listaActual.indexOf(titular) === -1) listaActual.push(titular);
  if (listaActual.indexOf(suplente) === -1) listaActual.push(suplente);

  cell.setValue(listaActual.join(", "));
  cell.setNote("REEMPLAZO REGISTRADO\nTitular: " + titular + "\nReemplazante: " + suplente + (motivo ? "\nMotivo: " + motivo : ""));

  ui.alert(
    "✅ Reemplazo Registrado",
    "Se asignó '" + listaActual.join(", ") + "' en " + unidadDetectada.nombre + " (" + fecha + " " + horario + ").\n" +
    "Tanto el titular (" + titular + ") como el reemplazante (" + suplente + ") contabilizan 0,5 hs automáticamente.",
    ui.ButtonSet.OK
  );
}

/**
 * 2. Pre-cargar el mes activo tomando los horarios de PLANTILLA SEMANAL
 */
function precargarMesDesdePlantilla() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const sheetName = sheet.getName();

  if (MESES_SISTEMA.indexOf(sheetName) === -1) {
    ui.alert("Abra primero la pestaña del mes que desea pre-cargar (ej. ENERO 2027) y vuelva a ejecutar esta opción.");
    return;
  }

  const confirm = ui.alert(
    "📅 Pre-cargar " + sheetName,
    "¿Desea completar los horarios de las 3 unidades (Proteína Gym, Proteína Box y Boxeo) en " + sheetName + " usando la PLANTILLA SEMANAL?",
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  const plantilla = ss.getSheetByName("PLANTILLA SEMANAL");
  const diasSemanaOrden = ["LUNES", "MARTES", "MIÉRCOLES", "JUEVES", "VIERNES", "SÁBADO", "DOMINGO"];
  const encabezadosDia = sheet.getRange(4, 2, 1, 31).getDisplayValues()[0];

  BLOQUES_UNIDADES.forEach(function(bloque) {
    const datosPlantilla = plantilla.getRange(bloque.filaPlantilla, 2, 32, 7).getValues();
    for (let c = 0; c < 31; c++) {
      const nombreDia = String(encabezadosDia[c] || "").trim().toUpperCase();
      const idxDia = diasSemanaOrden.indexOf(nombreDia);
      if (idxDia !== -1) {
        const colValores = [];
        for (let r = 0; r < 32; r++) {
          colValores.push([datosPlantilla[r][idxDia]]);
        }
        sheet.getRange(bloque.filaInicio, 2 + c, 32, 1).setValues(colValores);
      }
    }
  });

  ui.alert("✅ " + sheetName + " fue pre-cargado exitosamente con las 3 unidades desde la Plantilla Semanal.");
}

/**
 * 3. Actualizar listas desplegables dinámicas y colores condicionales desde PROFES
 */
function actualizarValidacionesYColores() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const profSheet = ss.getSheetByName("PROFES");
  const rangoActivos = profSheet.getRange("J5:J39");
  const datosProfes = profSheet.getRange("A5:D39").getValues();

  const reglaValidacion = SpreadsheetApp.newDataValidation()
    .requireValueInRange(rangoActivos, true)
    .setAllowInvalid(true)
    .build();

  MESES_SISTEMA.forEach(function(nombreMes) {
    const sh = ss.getSheetByName(nombreMes);
    if (!sh) return;

    const rangosGrilla = [
      sh.getRange("B9:AF40"),
      sh.getRange("B44:AF75"),
      sh.getRange("B79:AF110")
    ];

    rangosGrilla.forEach(function(rng) {
      rng.setDataValidation(reglaValidacion);
    });

    const reglasFormato = [];
    datosProfes.forEach(function(fila) {
      const nombre = String(fila[0] || "").trim();
      const hexBg = String(fila[3] || "").trim();
      if (nombre && hexBg && hexBg.indexOf("#") === 0) {
        const rule = SpreadsheetApp.newConditionalFormatRule()
          .whenTextContains(nombre)
          .setBackground(hexBg)
          .setRanges(rangosGrilla)
          .build();
        reglasFormato.push(rule);
      }
    });
    sh.setConditionalFormatRules(reglasFormato);
  });

  SpreadsheetApp.getUi().alert("✅ Listas desplegables y colores de profesores actualizados en los 14 meses.");
}

/**
 * 4. Sincronizar la base de datos interna REGISTRO DE HORAS normalizando selecciones múltiples
 */
function sincronizarRegistroDeHoras() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const bdSheet = ss.getSheetByName("REGISTRO DE HORAS");
  const cfgSheet = ss.getSheetByName("CONFIGURACIÓN");
  const cfgData = cfgSheet.getRange("A14:E27").getValues();

  const filasBD = [];
  let contadorId = 1;

  MESES_SISTEMA.forEach(function(nombreMes, mIdx) {
    const sh = ss.getSheetByName(nombreMes);
    if (!sh) return;

    const estadoMes = cfgData[mIdx][1];
    const valorHoraMes = Number(cfgData[mIdx][3]) || 5000;
    const diasRow = sh.getRange(4, 2, 1, 31).getDisplayValues()[0];
    const fechasRow = sh.getRange(5, 2, 1, 31).getDisplayValues()[0];
    const tarifaRow = sh.getRange(6, 2, 1, 31).getDisplayValues()[0];
    const horariosCol = sh.getRange(9, 1, 32, 1).getDisplayValues();

    BLOQUES_UNIDADES.forEach(function(bloque) {
      const valores = sh.getRange(bloque.filaInicio, 2, 32, 31).getDisplayValues();
      const notas = sh.getRange(bloque.filaInicio, 2, 32, 31).getNotes();

      for (let r = 0; r < 32; r++) {
        const partesHora = String(horariosCol[r][0] || "07:00 a 07:30").split(" a ");
        const hIni = partesHora[0] || "";
        const hFin = partesHora[1] || "";

        for (let c = 0; c < 31; c++) {
          const celdaStr = String(valores[r][c] || "").trim();
          if (!celdaStr || !fechasRow[c]) continue;

          const esDoble = String(tarifaRow[c]).indexOf("DOBLE") !== -1;
          const mult = esDoble ? 2 : 1;
          const tieneNotaReemplazo = String(notas[r][c] || "").indexOf("REEMPLAZO") !== -1;
          const nombres = celdaStr.split(",").map(function(n) { return n.trim(); }).filter(Boolean);

          nombres.forEach(function(profesor, nIdx) {
            const idStr = "ASIG-" + ("00000" + contadorId).slice(-5);
            const tipoAsig = (tieneNotaReemplazo && nIdx === nombres.length - 1 && nombres.length > 1) ? "Reemplazo" : "Titular";
            filasBD.push([
              idStr,
              nombreMes,
              fechasRow[c],
              diasRow[c],
              hIni,
              hFin,
              0.5,
              profesor,
              bloque.nombre,
              bloque.seccion,
              tipoAsig,
              esDoble ? "Doble (×2)" : "Normal (×1)",
              valorHoraMes,
              0.5 * valorHoraMes * mult,
              estadoMes
            ]);
            contadorId++;
          });
        }
      }
    });
  });

  if (bdSheet.getLastRow() > 4) {
    bdSheet.getRange(5, 1, bdSheet.getLastRow() - 4, 15).clearContent();
  }
  if (filasBD.length > 0) {
    bdSheet.getRange(5, 1, filasBD.length, 15).setValues(filasBD);
  }

  SpreadsheetApp.getUi().alert("✅ Base de datos REGISTRO DE HORAS sincronizada: " + filasBD.length + " registros individuales procesados.");
}

/**
 * 5. Cerrar / Congelar la liquidación mensual seleccionada en RESUMEN MENSUAL (Sección 11.5 - Regla R13)
 */
function cerrarLiquidacionMensual() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resSheet = ss.getSheetByName("RESUMEN MENSUAL");
  const cfgSheet = ss.getSheetByName("CONFIGURACIÓN");

  const mesSeleccionado = String(resSheet.getRange("C4").getValue() || "").trim();
  const mIdx = MESES_SISTEMA.indexOf(mesSeleccionado);
  if (mIdx === -1) {
    ui.alert("Seleccione un período válido en la celda C4 de RESUMEN MENSUAL.");
    return;
  }

  const tarifaVigente = Number(cfgSheet.getRange("B6").getValue()) || 5000;
  const resp = ui.alert(
    "🔒 Confirmar Cierre de Liquidación: " + mesSeleccionado,
    "Se marcará el período " + mesSeleccionado + " como LIQUIDADO congelando su Valor Hora en $" + tarifaVigente.toLocaleString("es-AR") + ".\n" +
    "De esta forma, futuros aumentos de tarifa no modificarán retroactivamente esta liquidación (Regla R13).\n\n¿Desea continuar?",
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  const filaCfg = 14 + mIdx;
  cfgSheet.getRange(filaCfg, 2).setValue("LIQUIDADO");
  cfgSheet.getRange(filaCfg, 3).setValue(tarifaVigente);
  cfgSheet.getRange(filaCfg, 8).setValue(new Date());
  cfgSheet.getRange(filaCfg, 9).setValue("Cerrado y congelado por Administrador");

  ui.alert("✅ El período " + mesSeleccionado + " ha quedado cerrado como LIQUIDADO.");
}

/**
 * 6. Proteger las hojas técnicas y bloques de fórmulas para evitar borrados accidentales (Sección 17)
 */
function protegerFormulasSistema() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojasProtegidas = ["RESUMEN ANUAL", "REPORTES POR UNIDAD", "REGISTRO DE HORAS"];

  hojasProtegidas.forEach(function(nombreHoja) {
    const sh = ss.getSheetByName(nombreHoja);
    if (sh) {
      const prot = sh.protect().setDescription("Protección automática de reportes y base de datos");
      prot.setWarningOnly(true);
    }
  });

  MESES_SISTEMA.forEach(function(nombreMes) {
    const sh = ss.getSheetByName(nombreMes);
    if (sh) {
      const protCalc = sh.getRange("A114:Q151").protect().setDescription("Cálculos mensuales automáticos");
      protCalc.setWarningOnly(true);
    }
  });

  SpreadsheetApp.getUi().alert("🛡️ Se activó la protección contra modificaciones accidentales en fórmulas y reportes.");
}
