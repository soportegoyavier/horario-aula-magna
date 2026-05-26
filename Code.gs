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
const BG_VACIO = '#d9d9d9';

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
        .addItem('Agregar Evento',      'mostrarAgregarEvento')
        .addItem('Eliminar Evento',     'mostrarEliminarEvento')
        .addSeparator()
        .addItem('Ver Todos los Eventos', 'mostrarTodosEventos')
        .addSeparator()
        .addItem('Crear nuevo mes',     'mostrarCrearMes')
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

function mostrarCrearMes() {
  const html = HtmlService.createHtmlOutputFromFile('CreateMonth')
    .setTitle('Crear Nuevo Mes');
  SpreadsheetApp.getUi().showSidebar(html);
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
        const numDias   = diaFin - diaInicio + 1;

        if (numDias >= 7 || numDias < 1) {
          // Semana completa o etiqueta con cruce de mes → todos los días
          diasValidos = [...DIAS];
        } else {
          // Semana parcial: calcular desde qué día empieza
          const jsDay     = new Date(year, mesIdx, diaInicio).getDay(); // 0=Dom,1=Lun,...
          const posInicio = jsDay === 0 ? 6 : jsDay - 1; // convertir a 0=Lun,...,6=Dom
          const posFin    = Math.min(posInicio + numDias - 1, 6);
          diasValidos     = DIAS.slice(posInicio, posFin + 1);
        }
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

// =====================================================================
// CREAR MES
// =====================================================================
function obtenerDatosCrearMes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const existentes = ss.getSheets()
    .map(s => s.getName().toUpperCase())
    .filter(n => n !== HOJA_DATOS);
  return {
    year: new Date().getFullYear(),
    meses: Object.keys(MESES_IDX),
    existentes
  };
}

function crearMes(datos) {
  try {
    const ss      = SpreadsheetApp.getActiveSpreadsheet();
    const nombre  = datos.mes.toUpperCase();
    const year    = parseInt(datos.year);
    const mesIdx  = MESES_IDX[nombre];

    if (mesIdx === undefined) return { ok: false, msg: 'Mes no reconocido.' };
    if (ss.getSheetByName(nombre)) return { ok: false, msg: `La hoja "${nombre}" ya existe.` };

    const semanas  = calcularSemanasDelMes(mesIdx, year);
    const hojaD    = ss.getSheetByName(HOJA_DATOS);
    const posicion = hojaD ? hojaD.getIndex() - 1 : ss.getSheets().length;
    const sheet    = ss.insertSheet(nombre, posicion);

    const COLOR_FONDO  = '#5a5b5d'; // fondo encabezados y HORA
    const COLOR_CELDAS = '#d9d9d9'; // fondo celdas de horario
    const COLOR_TEXTO  = '#FFFFFF';
    const FUENTE       = 'Poppins';

    let fila = 1;

    for (const semana of semanas) {
      // Fila 1: encabezado de semana (celdas combinadas A-H)
      const rngSemana = sheet.getRange(fila, 1, 1, 8);
      rngSemana.merge()
               .setValue(semana.label)
               .setBackground(COLOR_FONDO)
               .setFontColor(COLOR_TEXTO)
               .setFontWeight('bold')
               .setFontFamily(FUENTE)
               .setFontSize(15)
               .setHorizontalAlignment('center')
               .setVerticalAlignment('middle');
      sheet.setRowHeight(fila, 32);
      fila++;

      // Fila 2: encabezados de columna (HORA + días)
      const encabezados = ['HORA','Lunes','Martes','Miércoles','Jueves','Viernes','Sabado','Domingo'];
      sheet.getRange(fila, 1, 1, 8)
           .setValues([encabezados])
           .setBackground(COLOR_FONDO)
           .setFontColor(COLOR_TEXTO)
           .setFontWeight('bold')
           .setFontFamily(FUENTE)
           .setFontSize(17)
           .setHorizontalAlignment('center')
           .setVerticalAlignment('middle');
      sheet.setRowHeight(fila, 28);
      fila++;

      // Filas de horas
      for (let h = 0; h < HORAS.length; h++) {
        sheet.setRowHeight(fila + h, 25);

        // Celda HORA
        sheet.getRange(fila + h, 1)
             .setValue(HORAS[h])
             .setBackground(COLOR_FONDO)
             .setFontColor(COLOR_TEXTO)
             .setFontWeight('bold')
             .setFontStyle('italic')
             .setFontFamily(FUENTE)
             .setFontSize(11)
             .setHorizontalAlignment('center')
             .setVerticalAlignment('middle');

        // Celdas de días (B-H)
        sheet.getRange(fila + h, 2, 1, 7)
             .setBackground(COLOR_CELDAS)
             .setHorizontalAlignment('center')
             .setVerticalAlignment('middle')
             .setBorder(true, true, true, true, true, true,
                        '#b0b0b0', SpreadsheetApp.BorderStyle.SOLID);
      }
      fila += HORAS.length;

      // Fila separadora vacía
      sheet.setRowHeight(fila, 10);
      fila++;
    }

    // Anchos de columna
    sheet.setColumnWidth(1, 97);
    for (let c = 2; c <= 8; c++) sheet.setColumnWidth(c, 155);

    return { ok: true };
  } catch (err) {
    return { ok: false, msg: err.message };
  }
}

function calcularSemanasDelMes(mesIdx, year) {
  const primerDia      = new Date(year, mesIdx, 1);
  const ultimoDia      = new Date(year, mesIdx + 1, 0).getDate();
  const diaSemanaInicio = primerDia.getDay(); // 0=Dom,1=Lun,...,6=Sab

  // Primer domingo del mes (fin de la primera semana Mon-Dom)
  const primerDomingo = (diaSemanaInicio === 0) ? 1 : (8 - diaSemanaInicio);

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
