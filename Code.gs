// =====================================================================
// HORARIO AULA MAGNA — Apps Script principal
// =====================================================================

const COLORES = {
  'Ensayo':             { bg: '#B3D9FF', text: '#003366' },
  'Evento/Conferencia': { bg: '#FFD9B3', text: '#663300' },
  'Ceremonia':          { bg: '#E8D5FF', text: '#3D0066' },
  'Clase/Taller':       { bg: '#B3FFD9', text: '#003322' },
  'Otro':               { bg: '#F1F3F4', text: '#3C4043' }
};

const HORAS = [
  '07:00am','08:00am','09:00am','10:00am','11:00am',
  '12:00pm','01:00pm','02:00pm','03:00pm','04:00pm',
  '05:00pm','06:00pm','07:00pm','08:00pm'
];

const DIAS = ['Lunes','Martes','Miércoles','Jueves','Viernes','Sabado','Domingo'];
const HOJA_DATOS = 'DATOS';
const BG_VACIO = '#E8F4F8';

const MESES_IDX = {
  'ENERO':0,'FEBRERO':1,'MARZO':2,'ABRIL':3,'MAYO':4,'JUNIO':5,
  'JULIO':6,'AGOSTO':7,'SEPTIEMBRE':8,'OCTUBRE':9,'NOVIEMBRE':10,'DICIEMBRE':11
};

// =====================================================================
// MENÚ — solo visible para editores
// =====================================================================
function onOpen(e) {
  if (e && e.authMode === ScriptApp.AuthMode.NONE) return;
  try {
    if (esEditor()) {
      SpreadsheetApp.getUi()
        .createMenu('Aula Magna')
        .addItem('Agregar Evento', 'mostrarAgregarEvento')
        .addItem('Eliminar Evento', 'mostrarEliminarEvento')
        .addSeparator()
        .addItem('Ver Todos los Eventos', 'mostrarTodosEventos')
        .addToUi();
    }
  } catch (err) {
    // El usuario no tiene permisos — no se muestra el menú
  }
}

function esEditor() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const email = Session.getActiveUser().getEmail();
    if (!email) return false;
    const editores = ss.getEditors().map(u => u.getEmail());
    const dueno = ss.getOwner().getEmail();
    return editores.includes(email) || email === dueno;
  } catch (e) {
    return false;
  }
}

// =====================================================================
// SIDEBARS
// =====================================================================
function mostrarAgregarEvento() {
  const html = HtmlService.createHtmlOutputFromFile('AddEvent')
    .setTitle('Agregar Evento');
  SpreadsheetApp.getUi().showSidebar(html);
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
// DATOS PARA FORMULARIOS
// =====================================================================
function obtenerDatosIniciales() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const meses = ss.getSheets()
    .map(s => s.getName())
    .filter(n => n !== HOJA_DATOS);
  return {
    meses,
    dias: DIAS,
    horas: HORAS,
    categorias: Object.keys(COLORES)
  };
}

function obtenerSemanasDeMes(nombreMes) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(nombreMes);
  if (!sheet) return [];

  const ultimaFila = sheet.getLastRow();
  if (ultimaFila < 1) return [];

  const col1 = sheet.getRange(1, 1, ultimaFila, 1).getValues();
  const semanas = [];
  const mesIdx = MESES_IDX[nombreMes.toUpperCase()];
  const year = new Date().getFullYear();

  for (let i = 0; i < col1.length; i++) {
    const val = String(col1[i][0]).trim();
    if (val.toLowerCase().startsWith('semana')) {
      const match = val.match(/del\s+(\d+)\s+al\s+(\d+)/i);
      let diasValidos = [...DIAS]; // por defecto todos

      if (match && mesIdx !== undefined) {
        const diaInicio = parseInt(match[1]);
        const diaFin    = parseInt(match[2]);
        // JS getDay(): 0=Dom, 1=Lun, ..., 6=Sab → convertir a 0=Lun, ..., 6=Dom
        const jsDay     = new Date(year, mesIdx, diaInicio).getDay();
        const posInicio = jsDay === 0 ? 6 : jsDay - 1;
        const posFin    = Math.min(posInicio + (diaFin - diaInicio), 6);
        diasValidos     = DIAS.slice(posInicio, posFin + 1);
      }

      semanas.push({
        label:       val,
        indice:      semanas.length,
        filaInicio:  i + 1,
        diasValidos
      });
    }
  }
  return semanas;
}

// =====================================================================
// AGREGAR EVENTO
// =====================================================================
function agregarEvento(datos) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(datos.mes);
    if (!sheet) return { ok: false, msg: `Hoja "${datos.mes}" no encontrada` };

    const semanas = obtenerSemanasDeMes(datos.mes);
    const semana = semanas[datos.indiceSemana];
    if (!semana) return { ok: false, msg: 'Semana no encontrada' };

    const idxInicio = HORAS.indexOf(datos.horaInicio);
    const idxFin    = HORAS.indexOf(datos.horaFin);
    if (idxInicio < 0 || idxFin < 0) return { ok: false, msg: 'Hora inválida' };
    if (idxFin < idxInicio) return { ok: false, msg: 'La hora de fin debe ser igual o posterior al inicio' };

    const idxDia = DIAS.indexOf(datos.dia);
    if (idxDia < 0) return { ok: false, msg: 'Día inválido' };
    if (semana.diasValidos && !semana.diasValidos.includes(datos.dia)) {
      return { ok: false, msg: `"${datos.dia}" no pertenece a la semana "${semana.label}". Días válidos: ${semana.diasValidos.join(', ')}` };
    }

    // Fila: filaInicio de semana + 2 (encabezado semana + encabezado días) + índice hora
    const filaBase   = semana.filaInicio + 2;
    const filaInicio = filaBase + idxInicio;
    const filaFin    = filaBase + idxFin;
    const columna    = idxDia + 2; // columna A = HORA, columna B = Lunes, etc.
    const numFilas   = filaFin - filaInicio + 1;

    const rango = sheet.getRange(filaInicio, columna, numFilas, 1);

    // Verificar si ya hay contenido
    const contenido = rango.getValues().flat().filter(v => v !== '');
    if (contenido.length > 0) {
      return { ok: false, msg: 'Ya existe un evento en ese horario. Elimínalo primero.' };
    }

    // Construir texto del evento
    let texto = datos.nombre;
    if (datos.responsable) texto += '\n' + datos.responsable;
    if (datos.notas) texto += '\n' + datos.notas;

    // Aplicar contenido y formato
    if (numFilas > 1) rango.merge();
    const color = COLORES[datos.categoria] || { bg: '#FFE4B5', text: '#333333' };
    rango.setValue(texto)
         .setBackground(color.bg)
         .setFontColor(color.text)
         .setFontWeight('bold')
         .setWrap(true)
         .setVerticalAlignment('middle')
         .setHorizontalAlignment('center')
         .setFontSize(9);

    // Guardar en hoja DATOS
    const id = Utilities.getUuid();
    const hojaD = obtenerOCrearHojaDatos();
    hojaD.appendRow([
      id, datos.mes, datos.indiceSemana, datos.dia,
      datos.horaInicio, datos.horaFin,
      datos.nombre, datos.responsable || '', datos.notas || '',
      datos.categoria, filaInicio, filaFin, columna
    ]);

    return { ok: true, id };
  } catch (err) {
    return { ok: false, msg: err.message };
  }
}

// =====================================================================
// ELIMINAR EVENTO
// =====================================================================
function eliminarEvento(id) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaD = obtenerOCrearHojaDatos();
    const filas = hojaD.getDataRange().getValues();

    let filaObjetivo = -1;
    let evtDatos = null;

    for (let i = 1; i < filas.length; i++) {
      if (filas[i][0] === id) {
        filaObjetivo = i + 1;
        evtDatos = {
          mes:       filas[i][1],
          filaIni:   parseInt(filas[i][10]),
          filaFin:   parseInt(filas[i][11]),
          columna:   parseInt(filas[i][12])
        };
        break;
      }
    }

    if (!evtDatos) return { ok: false, msg: 'Evento no encontrado' };

    const sheet = ss.getSheetByName(evtDatos.mes);
    if (sheet) {
      const numFilas = evtDatos.filaFin - evtDatos.filaIni + 1;
      const rango = sheet.getRange(evtDatos.filaIni, evtDatos.columna, numFilas, 1);
      try { rango.breakApart(); } catch (e) {}
      // clearFormat() borra bordes — solo reseteamos lo que agregarEvento modificó
      rango.clearContent()
           .setBackground(BG_VACIO)
           .setFontColor('#000000')
           .setFontWeight('normal')
           .setFontSize(9)
           .setWrap(false)
           .setVerticalAlignment('middle')
           .setHorizontalAlignment('center');
    }

    hojaD.deleteRow(filaObjetivo);
    return { ok: true };
  } catch (err) {
    return { ok: false, msg: err.message };
  }
}

// =====================================================================
// OBTENER TODOS LOS EVENTOS
// =====================================================================
function obtenerTodosEventos() {
  try {
    const hojaD = obtenerOCrearHojaDatos();
    const filas = hojaD.getDataRange().getValues();
    const eventos = [];
    for (let i = 1; i < filas.length; i++) {
      if (filas[i][0]) {
        eventos.push({
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
    return eventos;
  } catch (e) {
    return [];
  }
}

// =====================================================================
// HOJA DE DATOS INTERNA
// =====================================================================
function obtenerOCrearHojaDatos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(HOJA_DATOS);
  if (!s) {
    s = ss.insertSheet(HOJA_DATOS);
    s.hideSheet();
    const headers = [
      'ID','Mes','Semana','Dia','HoraInicio','HoraFin',
      'Nombre','Responsable','Notas','Categoria',
      'FilaInicio','FilaFin','Columna'
    ];
    s.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    s.setFrozenRows(1);
  }
  return s;
}
