// =====================================================================
// HORARIO AULA MAGNA — Apps Script (arquitectura dinámica v2)
// =====================================================================

const COLORES = {
  'Ensayo':             { bg: '#B3D9FF', text: '#003366' },
  'Evento/Conferencia': { bg: '#FFD9B3', text: '#663300' },
  'Ceremonia':          { bg: '#E8D5FF', text: '#3D0066' },
  'Clase/Taller':       { bg: '#B3FFD9', text: '#003322' },
  'Otro':               { bg: '#F1F3F4', text: '#3C4043' }
};

// Horas en formato 24h (identificador interno)
const HORAS = [
  '07:00','08:00','09:00','10:00','11:00','12:00',
  '13:00','14:00','15:00','16:00','17:00','18:00','19:00','20:00'
];

// Versión visual con emojis para la columna HORA en HOJA1
const HORAS_DISPLAY = [
  '7:00 🌅','8:00','9:00','10:00','11:00','12:00',
  '13:00','14:00 🌇','15:00','16:00','17:00','18:00','19:00','20:00 🌙'
];

// Gradiente: amarillo claro (mañana) → amarillo fuerte (mediodía) → azul fuerte → azul claro (tarde)
const HORA_BG = [
  '#FFFDE7','#FFF9C4','#FFF176','#FFEE58','#FDD835','#F9CE1F',
  '#F5C000','#1565C0','#1976D2','#2196F3','#42A5F5','#64B5F6',
  '#90CAF9','#BBDEFB'
];
const HORA_TEXT = [
  '#827717','#827717','#827717','#827717','#827717','#827717',
  '#5D4037','#FFFFFF','#FFFFFF','#FFFFFF','#FFFFFF','#0D47A1',
  '#0D47A1','#0D47A1'
];

// Lunes a Sábado (sin Domingo)
const DIAS = ['Lunes','Martes','Miércoles','Jueves','Viernes','Sabado'];

const HOJA_DATOS  = 'DATOS';
const HOJA_VISTA  = 'HOJA1';
const FILA_INICIO_CONTENIDO = 5; // La semana empieza en la fila 5

const MESES_IDX = {
  'ENERO':0,'FEBRERO':1,'MARZO':2,'ABRIL':3,'MAYO':4,'JUNIO':5,
  'JULIO':6,'AGOSTO':7,'SEPTIEMBRE':8,'OCTUBRE':9,'NOVIEMBRE':10,'DICIEMBRE':11
};

// =====================================================================
// MENÚ — solo para editores
// =====================================================================
function onOpen(e) {
  if (e && e.authMode === ScriptApp.AuthMode.NONE) return;
  try {
    if (esEditor()) {
      SpreadsheetApp.getUi()
        .createMenu('Aula Magna')
        .addItem('Agregar Evento',        'mostrarAgregarEvento')
        .addItem('Eliminar Evento',        'mostrarEliminarEvento')
        .addSeparator()
        .addItem('Ver Todos los Eventos', 'mostrarTodosEventos')
        .addSeparator()
        .addItem('Configurar HOJA1',      'configurarHoja1')
        .addToUi();
    }
  } catch (err) {}
}

function esEditor() {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const email = Session.getActiveUser().getEmail();
    if (!email) return false;
    const editores = ss.getEditors().map(u => u.getEmail());
    return editores.includes(email) || email === ss.getOwner().getEmail();
  } catch (e) { return false; }
}

// =====================================================================
// onEdit — maneja los dropdowns de HOJA1
// =====================================================================
function onEdit(e) {
  try {
    const sheet = e.range.getSheet();
    if (sheet.getName() !== HOJA_VISTA) return;

    const row = e.range.getRow();
    const col = e.range.getColumn();

    // C2 (col 3) = selección de mes
    if (row === 2 && col === 3) {
      actualizarDropdownSemanas(sheet, e.value || '');
      limpiarVista(sheet);
    }

    // G2 (col 7) = selección de semana
    if (row === 2 && col === 7) {
      const mes      = String(sheet.getRange(2, 3).getValue()).trim();
      const semLabel = String(e.value || '').trim();
      if (!mes || !semLabel) return;

      const mesIdx = MESES_IDX[mes.toUpperCase()];
      if (mesIdx === undefined) return;

      const semanas = calcularSemanasDelMes(mesIdx, new Date().getFullYear());
      const idx     = semanas.findIndex(s => s.label === semLabel);
      if (idx >= 0) renderizarSemana(sheet, mes, idx);
    }
  } catch (err) {}
}

// =====================================================================
// CONFIGURAR / INICIALIZAR HOJA1
// =====================================================================
function configurarHoja1() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Crear o limpiar HOJA1
  let sheet = ss.getSheetByName(HOJA_VISTA);
  if (!sheet) {
    sheet = ss.insertSheet(HOJA_VISTA, 0);
  } else {
    sheet.clearContents();
    sheet.clearFormats();
    try { sheet.getRange(1, 1, sheet.getMaxRows(), 8).breakApart(); } catch(e) {}
  }

  // ── Fila 1: Título AULA MAGNA
  sheet.getRange(1, 1, 1, 7).merge()
       .setValue('AULA MAGNA')
       .setFontFamily('Poppins').setFontSize(22).setFontWeight('bold')
       .setHorizontalAlignment('center').setVerticalAlignment('middle')
       .setBackground('#FFFFFF').setFontColor('#202124');
  sheet.setRowHeight(1, 50);

  // ── Fila 2: controles de selección
  sheet.getRange(2, 1).setValue('SELECCIONE MES:')
       .setFontFamily('Poppins').setFontSize(11).setFontWeight('bold')
       .setVerticalAlignment('middle');

  // Dropdown de mes en C2
  const reglaMes = SpreadsheetApp.newDataValidation()
    .requireValueInList(Object.keys(MESES_IDX), true)
    .setAllowInvalid(false).build();
  sheet.getRange(2, 3).setDataValidation(reglaMes).clearContent()
       .setFontFamily('Poppins').setFontSize(11);

  sheet.getRange(2, 5).setValue('SEMANA:')
       .setFontFamily('Poppins').setFontSize(11).setFontWeight('bold')
       .setVerticalAlignment('middle');

  // Dropdown de semana en G2 (se puebla cuando se elige mes)
  const reglaSem = SpreadsheetApp.newDataValidation()
    .requireValueInList(['— Selecciona un mes primero —'], true)
    .setAllowInvalid(false).build();
  sheet.getRange(2, 7).setDataValidation(reglaSem).clearContent()
       .setFontFamily('Poppins').setFontSize(11);

  sheet.setRowHeight(2, 35);

  // ── Filas 3-4: separadores finos
  sheet.setRowHeight(3, 6);
  sheet.setRowHeight(4, 6);

  // ── Anchos de columna
  sheet.setColumnWidth(1, 97);              // HORA
  for (let c = 2; c <= 7; c++) sheet.setColumnWidth(c, 155); // 6 días

  // ── Congelar las 4 filas de control
  sheet.setFrozenRows(4);

  // ── Ocultar hojas de meses individuales (si existen)
  ss.getSheets().forEach(s => {
    const nombre = s.getName().toUpperCase();
    if (MESES_IDX[nombre] !== undefined) {
      try { s.hideSheet(); } catch (e) {}
    }
  });

  // ── Asegurarse de que HOJA1 sea visible y activa
  sheet.showSheet();
  ss.setActiveSheet(sheet);

  SpreadsheetApp.getUi().alert('✅ HOJA1 configurada correctamente.\n\nAhora selecciona un mes y una semana para ver el horario.');
}

// =====================================================================
// ACTUALIZAR DROPDOWN DE SEMANAS (cuando cambia el mes)
// =====================================================================
function actualizarDropdownSemanas(sheet, mes) {
  const mesIdx = mes ? MESES_IDX[mes.toUpperCase()] : undefined;

  let regla;
  if (mesIdx !== undefined) {
    const labels = calcularSemanasDelMes(mesIdx, new Date().getFullYear()).map(s => s.label);
    regla = SpreadsheetApp.newDataValidation()
      .requireValueInList(labels, true).setAllowInvalid(false).build();
  } else {
    regla = SpreadsheetApp.newDataValidation()
      .requireValueInList(['— Selecciona un mes primero —'], true)
      .setAllowInvalid(false).build();
  }
  sheet.getRange(2, 7).setDataValidation(regla).clearContent();
}

// =====================================================================
// RENDERIZAR LA SEMANA SELECCIONADA EN HOJA1
// =====================================================================
function renderizarSemana(sheet, mes, indiceSemana) {
  const mesIdx = MESES_IDX[mes.toUpperCase()];
  if (mesIdx === undefined) return;

  const year    = new Date().getFullYear();
  const semanas = calcularSemanasDelMes(mesIdx, year);
  const semana  = semanas[indiceSemana];
  if (!semana) return;

  limpiarVista(sheet);

  const diasValidos = calcularDiasValidos(mesIdx, year, semana.inicio, semana.fin);
  const F = FILA_INICIO_CONTENIDO;

  // ── Fila F: título de la semana
  sheet.getRange(F, 1, 1, 7).merge()
       .setValue(semana.label)
       .setFontFamily('Poppins').setFontSize(15).setFontWeight('normal')
       .setHorizontalAlignment('center').setVerticalAlignment('middle')
       .setBackground('#FFFFFF').setFontColor('#202124')
       .setBorder(false, false, true, false, false, false,
                  '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);
  sheet.setRowHeight(F, 32);

  // ── Fila F+1: encabezados HORA + días (todo en mayúscula)
  const encabezados = ['HORA', ...DIAS.map(d => d.toUpperCase())];
  sheet.getRange(F + 1, 1, 1, 7).setValues([encabezados])
       .setFontFamily('Poppins').setFontSize(15).setFontWeight('normal')
       .setHorizontalAlignment('center').setVerticalAlignment('middle')
       .setBackground('#FFFFFF').setFontColor('#202124')
       .setBorder(false, false, true, false, false, false,
                  '#202124', SpreadsheetApp.BorderStyle.SOLID);
  sheet.setRowHeight(F + 1, 34);

  // ── Filas de horas con gradiente
  for (let h = 0; h < HORAS.length; h++) {
    const row = F + 2 + h;
    sheet.setRowHeight(row, 37);

    // Celda HORA (con gradiente de color)
    sheet.getRange(row, 1)
         .setValue(HORAS_DISPLAY[h])
         .setFontFamily('Poppins').setFontSize(11)
         .setFontWeight('normal').setFontStyle('normal')
         .setBackground(HORA_BG[h]).setFontColor(HORA_TEXT[h])
         .setHorizontalAlignment('center').setVerticalAlignment('middle');

    // Celdas de días (Lunes=col2 … Sabado=col7)
    for (let d = 0; d < DIAS.length; d++) {
      const col   = d + 2;
      const valid = diasValidos.map(x => x.toLowerCase()).includes(DIAS[d].toLowerCase());
      sheet.getRange(row, col)
           .setBackground(valid ? '#FFFFFF' : '#F0F0F0')
           .setFontColor('#202124')
           .setHorizontalAlignment('center').setVerticalAlignment('middle')
           .setBorder(true, true, true, true, false, false,
                      '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);
    }
  }

  // ── Cargar eventos de esta semana desde DATOS
  const eventos = obtenerTodosEventos().filter(e =>
    e.mes.toUpperCase() === mes.toUpperCase() &&
    parseInt(e.indiceSemana) === indiceSemana
  );

  eventos.forEach(evt => {
    const idxIni = HORAS.indexOf(normalizarHora(evt.horaInicio));
    const idxFin = HORAS.indexOf(normalizarHora(evt.horaFin));
    const idxDia = DIAS.map(d => d.toLowerCase()).indexOf(String(evt.dia).toLowerCase());
    if (idxIni < 0 || idxFin < 0 || idxDia < 0) return;

    const filaIni  = F + 2 + idxIni;
    const filaFin  = F + 2 + idxFin;
    const col      = idxDia + 2;
    const numFilas = filaFin - filaIni + 1;

    const rango = sheet.getRange(filaIni, col, numFilas, 1);
    if (numFilas > 1) try { rango.merge(); } catch (ex) {}

    let texto = evt.nombre;
    if (evt.responsable) texto += '\n' + evt.responsable;
    if (evt.notas)       texto += '\n' + evt.notas;

    const color = COLORES[evt.categoria] || COLORES['Otro'];
    rango.setValue(texto)
         .setBackground(color.bg).setFontColor(color.text)
         .setFontWeight('bold').setFontFamily('Poppins').setFontSize(9)
         .setWrap(true).setVerticalAlignment('middle').setHorizontalAlignment('center');
  });
}

// =====================================================================
// LIMPIAR ÁREA DE VISTA (filas 5 en adelante)
// =====================================================================
function limpiarVista(sheet) {
  const numRows = HORAS.length + 3;
  const rango   = sheet.getRange(FILA_INICIO_CONTENIDO, 1, numRows, 7);
  try { rango.breakApart(); } catch (e) {}
  rango.clearContent().clearFormat().setBackground('#FFFFFF');
}

// =====================================================================
// SIDEBARS
// =====================================================================
function mostrarAgregarEvento() {
  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createHtmlOutputFromFile('AddEvent').setTitle('Agregar Evento')
  );
}
function mostrarEliminarEvento() {
  const tmpl = HtmlService.createTemplateFromFile('RemoveEvent');
  tmpl.eventosJson = JSON.stringify(obtenerTodosEventos());
  SpreadsheetApp.getUi().showSidebar(tmpl.evaluate().setTitle('Eliminar Evento'));
}
function mostrarTodosEventos() {
  const tmpl = HtmlService.createTemplateFromFile('AllEvents');
  tmpl.eventosJson = JSON.stringify(obtenerTodosEventos());
  SpreadsheetApp.getUi().showSidebar(tmpl.evaluate().setTitle('Todos los Eventos'));
}

// =====================================================================
// DATOS PARA FORMULARIOS (AddEvent.html)
// =====================================================================
function obtenerDatosIniciales() {
  return {
    meses:      Object.keys(MESES_IDX),
    dias:       DIAS,
    horas:      HORAS,
    categorias: Object.keys(COLORES)
  };
}

function obtenerSemanasDeMes(nombreMes) {
  const mesIdx = MESES_IDX[nombreMes.toUpperCase()];
  if (mesIdx === undefined) return [];
  const year = new Date().getFullYear();
  return calcularSemanasDelMes(mesIdx, year).map((s, i) => ({
    label:       s.label,
    indice:      i,
    filaInicio:  -1,
    diasValidos: calcularDiasValidos(mesIdx, year, s.inicio, s.fin)
  }));
}

// =====================================================================
// AGREGAR EVENTO — solo escribe en DATOS y refresca la vista
// =====================================================================
function agregarEvento(datos) {
  try {
    const semanas = obtenerSemanasDeMes(datos.mes);
    const semana  = semanas[datos.indiceSemana];
    if (!semana) return { ok: false, msg: 'Semana no encontrada' };

    const idxIni = HORAS.indexOf(datos.horaInicio);
    const idxFin = HORAS.indexOf(datos.horaFin);
    if (idxIni < 0 || idxFin < 0) return { ok: false, msg: 'Hora inválida' };
    if (idxFin < idxIni)           return { ok: false, msg: 'La hora de fin debe ser igual o posterior al inicio' };

    const idxDia = DIAS.map(d => d.toLowerCase()).indexOf(datos.dia.toLowerCase());
    if (idxDia < 0) return { ok: false, msg: 'Día inválido' };

    if (semana.diasValidos && !semana.diasValidos.map(d => d.toLowerCase()).includes(datos.dia.toLowerCase())) {
      return { ok: false, msg: `"${datos.dia}" no pertenece a esta semana.` };
    }

    // Verificar solapamiento en DATOS
    const existentes = obtenerTodosEventos().filter(e =>
      e.mes.toUpperCase() === datos.mes.toUpperCase() &&
      parseInt(e.indiceSemana) === datos.indiceSemana &&
      e.dia.toLowerCase() === datos.dia.toLowerCase()
    );
    for (const evt of existentes) {
      const eIni = HORAS.indexOf(normalizarHora(evt.horaInicio));
      const eFin = HORAS.indexOf(normalizarHora(evt.horaFin));
      if (eIni <= idxFin && eFin >= idxIni) {
        return { ok: false, msg: 'Ya existe un evento en ese horario. Elimínalo primero.' };
      }
    }

    // Guardar en DATOS (las horas como texto para evitar auto-conversión de Sheets)
    const id    = Utilities.getUuid();
    const hojaD = obtenerOCrearHojaDatos();
    const nuevaFila = hojaD.getLastRow() + 1;
    hojaD.appendRow([
      id, datos.mes, datos.indiceSemana, datos.dia,
      datos.horaInicio, datos.horaFin,
      datos.nombre, datos.responsable || '', datos.notas || '',
      datos.categoria
    ]);
    // Forzar formato texto en las celdas de hora para que no se conviertan a Date
    hojaD.getRange(nuevaFila, 5, 1, 2).setNumberFormat('@');

    refrescarVistaActual(datos.mes, datos.indiceSemana);
    return { ok: true, id };
  } catch (err) {
    return { ok: false, msg: err.message };
  }
}

// =====================================================================
// ELIMINAR EVENTO — borra de DATOS y refresca la vista
// =====================================================================
function eliminarEvento(id) {
  try {
    const hojaD = obtenerOCrearHojaDatos();
    const filas  = hojaD.getDataRange().getValues();

    let filaObj = -1, mesDel = '', semDel = -1;
    for (let i = 1; i < filas.length; i++) {
      if (filas[i][0] === id) {
        filaObj = i + 1;
        mesDel  = filas[i][1];
        semDel  = parseInt(filas[i][2]);
        break;
      }
    }
    if (filaObj < 0) return { ok: false, msg: 'Evento no encontrado' };

    hojaD.deleteRow(filaObj);
    refrescarVistaActual(mesDel, semDel);
    return { ok: true };
  } catch (err) {
    return { ok: false, msg: err.message };
  }
}

// Refresca HOJA1 si está mostrando exactamente el mes/semana modificado
function refrescarVistaActual(mes, indiceSemana) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(HOJA_VISTA);
    if (!sheet) return;

    const mesActual  = String(sheet.getRange(2, 3).getValue()).trim();
    const semActual  = String(sheet.getRange(2, 7).getValue()).trim();
    if (mesActual.toUpperCase() !== String(mes).toUpperCase()) return;

    const mesIdx  = MESES_IDX[mesActual.toUpperCase()];
    if (mesIdx === undefined) return;

    const semanas = calcularSemanasDelMes(mesIdx, new Date().getFullYear());
    const idx     = semanas.findIndex(s => s.label === semActual);
    if (idx === indiceSemana) renderizarSemana(sheet, mes, indiceSemana);
  } catch (e) {}
}

// =====================================================================
// OBTENER TODOS LOS EVENTOS
// =====================================================================
function obtenerTodosEventos() {
  try {
    const hojaD = obtenerOCrearHojaDatos();
    const filas  = hojaD.getDataRange().getValues();
    const evts   = [];
    for (let i = 1; i < filas.length; i++) {
      if (filas[i][0]) {
        evts.push({
          id:           filas[i][0],
          mes:          filas[i][1],
          indiceSemana: filas[i][2],
          dia:          filas[i][3],
          horaInicio:   filas[i][4],
          horaFin:      filas[i][5],
          nombre:       filas[i][6],
          responsable:  filas[i][7],
          notas:        filas[i][8],
          categoria:    filas[i][9]
        });
      }
    }
    return evts;
  } catch (e) { return []; }
}

// =====================================================================
// HOJA DE DATOS INTERNA (oculta)
// =====================================================================
function obtenerOCrearHojaDatos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s    = ss.getSheetByName(HOJA_DATOS);
  if (!s) {
    s = ss.insertSheet(HOJA_DATOS);
    s.hideSheet();
    const headers = ['ID','Mes','Semana','Dia','HoraInicio','HoraFin',
                     'Nombre','Responsable','Notas','Categoria'];
    s.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    s.setFrozenRows(1);
  }
  return s;
}

// =====================================================================
// HELPERS — cálculo de semanas y días válidos
// =====================================================================
function calcularSemanasDelMes(mesIdx, year) {
  const primerDia      = new Date(year, mesIdx, 1);
  const ultimoDia      = new Date(year, mesIdx + 1, 0).getDate();
  const diaSemIni      = primerDia.getDay(); // 0=Dom,1=Lun,...,6=Sab
  const primerDomingo  = (diaSemIni === 0) ? 1 : (8 - diaSemIni);

  const semanas = [];
  let inicio = 1;
  let fin    = Math.min(primerDomingo, ultimoDia);
  semanas.push({ label: `Semana del ${inicio} al ${fin}`, inicio, fin });
  inicio = fin + 1;

  while (inicio <= ultimoDia) {
    fin = Math.min(inicio + 6, ultimoDia);
    semanas.push({ label: `Semana del ${inicio} al ${fin}`, inicio, fin });
    inicio = fin + 1;
  }
  return semanas;
}

function calcularDiasValidos(mesIdx, year, diaInicio, diaFin) {
  const numDias = diaFin - diaInicio + 1;
  if (numDias >= 6 || numDias < 1) return [...DIAS];

  const jsDay = new Date(year, mesIdx, diaInicio).getDay(); // 0=Dom,1=Lun,...,6=Sab
  if (jsDay === 0) {
    // Empieza en domingo — sin domingo en nuestro sistema, mostramos desde Lunes
    return DIAS.slice(0, Math.min(numDias, 6));
  }
  const posIni = jsDay - 1; // 0=Lun,...,5=Sab
  const posFin = Math.min(posIni + numDias - 1, 5);
  return DIAS.slice(posIni, posFin + 1);
}

// Convierte hora al formato 24h "HH:MM" usado internamente.
// Google Sheets auto-convierte "07:00" a un objeto Date al leerlo desde la hoja,
// por eso se detecta instanceof Date primero.
function normalizarHora(h) {
  if (h instanceof Date) {
    return String(h.getHours()).padStart(2, '0') + ':' + String(h.getMinutes()).padStart(2, '0');
  }
  h = String(h).trim().toLowerCase();
  if (/^\d{2}:\d{2}$/.test(h)) return h;
  const m = h.match(/^(\d{1,2}):(\d{2})\s*(am|pm)$/);
  if (!m) return h;
  let hr = parseInt(m[1]);
  if (m[3] === 'pm' && hr < 12) hr += 12;
  if (m[3] === 'am' && hr === 12) hr = 0;
  return String(hr).padStart(2, '0') + ':' + m[2];
}
