// ─── CONFIG ───────────────────────────────────────────────────────────────────
var STAFFING_SHEET_ID   = '1fPL_GZfRAH_Jz8hHfvZtPEpxf1dAlJ-8W8JOBgsTw18';
var STAFFING_SHEET_NAME = 'Estado Actual Operación';
var STAFFING_ROW_DATA   = 11; // primera fila de datos (colaboradores/proyectos)

var SHEET_PROYECTOS  = 'Proyectos';
var SHEET_MEDICIONES = 'Mediciones';

// ─── ENTRY POINT ──────────────────────────────────────────────────────────────
function doGet() {
  return HtmlService.createHtmlOutputFromFile('isam')
    .setTitle('ISAM — Háptica')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ─── CARGA INICIAL (una sola llamada desde el cliente) ────────────────────────
function getAll() {
  return {
    proyectos:       getProyectos(),
    mediciones:      getMediciones(),
    horasEjecutadas: getHorasEjecutadas()
  };
}

// ─── PROYECTOS ────────────────────────────────────────────────────────────────
function getProyectos() {
  var sheet = _sheet(SHEET_PROYECTOS);
  if (!sheet) return [];

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var data = sheet.getRange(2, 1, lastRow - 1, 16).getValues();
  return data
    .filter(function(r) { return r[0]; })
    .map(function(r) {
      return {
        nombre:               r[0],
        host:                 r[1],
        owner:                r[2],
        doer1:                r[3],
        doer2:                r[4],
        otros1:               r[5],
        otros2:               r[6],
        horasPresupuestadas:  r[7],
        fechaInicio:          _fecha(r[8]),
        fechaCierre:          _fecha(r[9]),
        fechaProyectada:      _fecha(r[10]),
        semanasPresupuestadas: r[11],
        semanasProyectadas:   r[12],
        conocimientoCliente:  r[13],
        innovacionEquipo:     r[14],
        estado:               r[15] || 'activo'
      };
    });
}

function saveProyecto(data) {
  var sheet = _sheetOrCreate(SHEET_PROYECTOS,
    ['Nombre','Host','Owner','Doer1','Doer2','Otros1','Otros2',
     'HorasPresupuestadas','FechaInicio','FechaCierre','FechaProyectada',
     'SemanasPresupuestadas','SemanasProyectadas','ConocimientoCliente',
     'InnovacionEquipo','Estado']);

  sheet.appendRow([
    data.nombre, data.host, data.owner, data.doer1, data.doer2,
    data.otros1, data.otros2, data.horasPresupuestadas,
    data.fechaInicio    ? new Date(data.fechaInicio)    : '',
    data.fechaCierre    ? new Date(data.fechaCierre)    : '',
    data.fechaProyectada ? new Date(data.fechaProyectada) : '',
    data.semanasPresupuestadas, data.semanasProyectadas,
    data.conocimientoCliente, data.innovacionEquipo, 'activo'
  ]);
  return { ok: true };
}

function saveFechaProyectada(nombre, fecha) {
  var sheet = _sheet(SHEET_PROYECTOS);
  if (!sheet) return { ok: false };

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { ok: false };

  var nombres = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < nombres.length; i++) {
    if ((nombres[i][0] || '').toString().trim() === nombre) {
      // Column 11 = FechaProyectada (index base-1)
      sheet.getRange(i + 2, 11).setValue(fecha ? new Date(fecha) : '');
      return { ok: true };
    }
  }
  return { ok: false };
}

// ─── MEDICIONES ISAM ──────────────────────────────────────────────────────────
function getMediciones() {
  var sheet = _sheet(SHEET_MEDICIONES);
  if (!sheet) return [];

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
  return data
    .filter(function(r) { return r[0]; })
    .map(function(r) {
      return {
        proyecto:    r[0],
        fecha:       _fecha(r[1]),
        friccion:    r[2],
        dependencia: r[3],
        scopeCreep:  r[4],
        burnout:     r[5],
        promedio:    r[6],
        notas:       r[7]
      };
    });
}

function saveMedicion(data) {
  var sheet = _sheetOrCreate(SHEET_MEDICIONES,
    ['Proyecto','Fecha','Friccion','Dependencia','ScopeCreep',
     'Burnout','PromedioISAM','Notas']);

  var promedio = parseFloat(
    ((data.friccion + data.dependencia + data.scopeCreep + data.burnout) / 4).toFixed(2)
  );

  sheet.appendRow([
    data.proyecto,
    new Date(data.fecha),
    data.friccion, data.dependencia, data.scopeCreep, data.burnout,
    promedio,
    data.notas || ''
  ]);
  return { ok: true, promedio: promedio };
}

// ─── HORAS EJECUTADAS (desde staffing) ───────────────────────────────────────
/**
 * Suma todas las horas registradas por proyecto en el sheet de staffing.
 * Devuelve: { "Nombre Proyecto": totalHoras, ... }
 */
function getHorasEjecutadas() {
  try {
    var ss    = SpreadsheetApp.openById(STAFFING_SHEET_ID);
    var sheet = ss.getSheetByName(STAFFING_SHEET_NAME);
    if (!sheet) return {};

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < STAFFING_ROW_DATA) return {};

    var numRows    = lastRow - STAFFING_ROW_DATA + 1;
    var dataRange  = sheet.getRange(STAFFING_ROW_DATA, 1, numRows, lastCol);
    var values     = dataRange.getValues();
    var fontWeights = dataRange.getFontWeights();

    var resultado = {};

    for (var r = 0; r < values.length; r++) {
      var nombre  = (values[r][0] || '').toString().trim();
      var esBold  = (fontWeights[r][0] === 'bold');

      if (!nombre || esBold) continue; // fila vacía o fila de colaborador

      // Fila de proyecto: sumar todas las columnas numéricas
      var total = 0;
      for (var c = 1; c < values[r].length; c++) {
        var val = values[r][c];
        if (typeof val === 'number' && val > 0) total += val;
      }
      if (total > 0) {
        resultado[nombre] = (resultado[nombre] || 0) + total;
      }
    }
    return resultado;

  } catch (e) {
    Logger.log('Error getHorasEjecutadas: ' + e.message);
    return {};
  }
}

// ─── HELPERS INTERNOS ─────────────────────────────────────────────────────────
function _sheet(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name);
}

function _sheetOrCreate(name, headers) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function _fecha(val) {
  if (!val) return '';
  try {
    return Utilities.formatDate(new Date(val), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  } catch (e) {
    return '';
  }
}

/**
 * Ejecutar una vez manualmente para crear las hojas con sus cabeceras.
 */
function initSheets() {
  _sheetOrCreate(SHEET_PROYECTOS,
    ['Nombre','Host','Owner','Doer1','Doer2','Otros1','Otros2',
     'HorasPresupuestadas','FechaInicio','FechaCierre','FechaProyectada',
     'SemanasPresupuestadas','SemanasProyectadas','ConocimientoCliente',
     'InnovacionEquipo','Estado']);

  _sheetOrCreate(SHEET_MEDICIONES,
    ['Proyecto','Fecha','Friccion','Dependencia','ScopeCreep',
     'Burnout','PromedioISAM','Notas']);

  return { ok: true };
}
