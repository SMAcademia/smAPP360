// ===========================================================================
//  SM ACADEMIA - Google Apps Script v2.5
//  Sheets ID: 1oHJIUoyR3V5N0iweWnMagZic_8YkWtVV6CsshXloX4k
// ===========================================================================

var SPREADSHEET_ID = '1oHJIUoyR3V5N0iweWnMagZic_8YkWtVV6CsshXloX4k';

function ss_() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

// ---------------------------------------------------------------------------
// GET - JSONP desde la app (solo lectura y PING)
// ---------------------------------------------------------------------------
function doGet(e) {
  var callback = (e && e.parameter && e.parameter.callback) ? e.parameter.callback : 'callback';
  var action   = (e && e.parameter && e.parameter.action)   ? e.parameter.action   : '';

  try {
    var result;
    switch (action) {

      case 'PING':
        result = { ok: true, ts: new Date().toISOString() };
        break;

      case 'GET_ALL':
        result = getAllData();
        break;

      case 'APPEND':
        result = appendRow(
          e.parameter.sheet,
          JSON.parse(e.parameter.row || '{}')
        );
        break;

      case 'UPDATE':
        result = updateRow(
          e.parameter.sheet,
          e.parameter.id,
          JSON.parse(e.parameter.row || '{}')
        );
        break;

      case 'DELETE':
        result = deleteRow(e.parameter.sheet, e.parameter.id);
        break;

      default:
        result = { ok: false, error: 'Accion desconocida: ' + action };
    }

    return jsonp(callback, result);

  } catch (err) {
    return jsonp(callback, { ok: false, error: err.toString() });
  }
}

// ---------------------------------------------------------------------------
// POST - app interna (APPEND/UPDATE/DELETE) + formulario publico
// ---------------------------------------------------------------------------
function doPost(e) {
  try {
    var body   = JSON.parse(e.postData.contents);
    var action = body.action || 'NUEVA_PREINSCRIPCION';
    var result;

    if (action === 'APPEND') {
      result = appendRow(body.sheet, body.row || {});
    } else if (action === 'UPDATE') {
      result = updateRow(body.sheet, body.id, body.row || {});
    } else if (action === 'DELETE') {
      result = deleteRow(body.sheet, body.id);
    } else if (action === 'NUEVA_PREINSCRIPCION') {
      result = nuevaPreinscripcion(body);
    } else {
      result = { ok: false, error: 'Accion desconocida: ' + action };
    }

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ---------------------------------------------------------------------------
// Helper JSONP
// ---------------------------------------------------------------------------
function jsonp(callback, data) {
  return ContentService
    .createTextOutput(callback + '(' + JSON.stringify(data) + ')')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

// ===========================================================================
//  LEER TODAS LAS HOJAS
// ===========================================================================
function getAllData() {
  var book = ss_();
  return {
    ok:                      true,
    clientes:                sheetToJSON(book, 'CLIENTES'),
    alumnos:                 sheetToJSON(book, 'ALUMNOS'),
    actividades:             sheetToJSON(book, 'ACTIVIDADES'),
    preinscripciones:        sheetToJSON(book, 'PREINSCRIPCIONES'),
    inscripciones:           sheetToJSON(book, 'INSCRIPCIONES'),
    pagos:                   sheetToJSON(book, 'PAGOS'),
    facturas:                sheetToJSON(book, 'FACTURAS'),
    personal:                sheetToJSON(book, 'PERSONAL'),
    ausencias:               sheetToJSON(book, 'AUSENCIAS'),
    registroHorario:         sheetToJSON(book, 'REGISTRO_HORARIO'),
    sesiones:                sheetToJSON(book, 'SESIONES'),
    asistenciasFutbol:       sheetToJSON(book, 'ASISTENCIAS_FUTBOL'),
    asistenciasExtra:        sheetToJSON(book, 'ASISTENCIAS_EXTRA'),
    nominas:                 sheetToJSON(book, 'NOMINAS'),
    gastos:                  sheetToJSON(book, 'GASTOS'),
    movimientosBancarios:    sheetToJSON(book, 'MOVIMIENTOS_BANCARIOS'),
    ampaSocios:              sheetToJSON(book, 'AMPA_SOCIOS'),
    ampaAlumnos:             sheetToJSON(book, 'AMPA_ALUMNOS'),
    encuestasSatisfaccion:   sheetToJSON(book, 'ENCUESTAS_SATISFACCION'),
  };
}

// Convierte una hoja en array de objetos ([] si no existe)
function sheetToJSON(book, sheetName) {
  var sheet = book.getSheetByName(sheetName);
  if (!sheet) {
    Logger.log('AVISO: hoja no encontrada -> ' + sheetName);
    return [];
  }

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];

  var data    = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var headers = data[0].map(function(h) { return String(h).trim(); });
  var rows    = [];

  for (var i = 1; i < data.length; i++) {
    var row     = data[i];
    var isEmpty = row.every(function(c) { return c === '' || c === null || c === undefined; });
    if (isEmpty) continue;

    var obj = {};
    headers.forEach(function(h, j) {
      var v = row[j];
      if (v instanceof Date) {
        if (h.indexOf('HORA') !== -1) {
          obj[h] = Utilities.formatDate(v, Session.getScriptTimeZone(), 'HH:mm');
        } else {
          obj[h] = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        }
      } else {
        obj[h] = (v !== null && v !== undefined) ? v : '';
      }
    });
    rows.push(obj);
  }
  return rows;
}

// ===========================================================================
//  APPEND - protegido con LockService para evitar IDs duplicados
// ===========================================================================
function appendRow(sheetName, rowObj) {
  var lock   = LockService.getScriptLock();
  var locked = false;
  try {
    lock.waitLock(12000);
    locked = true;

    var book  = ss_();
    var sheet = book.getSheetByName(sheetName);
    if (!sheet) return { ok: false, error: 'Hoja no encontrada: ' + sheetName };

    var lastCol = sheet.getLastColumn();
    if (lastCol < 1) return { ok: false, error: 'Sin cabeceras: ' + sheetName };

    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
      .map(function(h) { return String(h).trim(); });

    var idCol = headers[0];

    if (idCol && !rowObj[idCol]) {
      rowObj[idCol] = nextId_(sheet, sheetName, headers);
    }

    if (idCol && rowObj[idCol]) {
      var existingId = String(rowObj[idCol]);
      var lastRow2   = sheet.getLastRow();
      if (lastRow2 > 1) {
        var existingIds = sheet.getRange(2, 1, lastRow2 - 1, 1).getValues();
        for (var k = 0; k < existingIds.length; k++) {
          if (String(existingIds[k][0]) === existingId) {
            return { ok: false, error: 'ID duplicado: ' + existingId, duplicate: true };
          }
        }
      }
    }

    var newRow = headers.map(function(h) {
      var v = (rowObj[h] !== undefined && rowObj[h] !== null) ? rowObj[h] : '';
      if (h.indexOf('HORA') !== -1 && v !== '') return String(v);
      return v;
    });

    sheet.appendRow(newRow);
    return { ok: true, id: rowObj[idCol] || '' };

  } catch (err) {
    if (!locked) return { ok: false, error: 'Escritura simultanea. Reintenta.' };
    return { ok: false, error: err.toString() };
  } finally {
    if (locked) { try { lock.releaseLock(); } catch(e2) {} }
  }
}

// ===========================================================================
//  UPDATE - protegido con LockService
// ===========================================================================
function updateRow(sheetName, id, rowObj) {
  var lock   = LockService.getScriptLock();
  var locked = false;
  try {
    lock.waitLock(12000);
    locked = true;

    var book  = ss_();
    var sheet = book.getSheetByName(sheetName);
    if (!sheet) return { ok: false, error: 'Hoja no encontrada: ' + sheetName };

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 2) return { ok: false, error: 'Hoja vacia' };

    var data    = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    var headers = data[0].map(function(h) { return String(h).trim(); });

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        var updatedRow = headers.map(function(h, j) {
          if (rowObj[h] !== undefined) {
            if (h.indexOf('HORA') !== -1 && rowObj[h] !== '') return String(rowObj[h]);
            return rowObj[h];
          }
          return data[i][j];
        });
        sheet.getRange(i + 1, 1, 1, headers.length).setValues([updatedRow]);
        return { ok: true, id: id };
      }
    }
    return { ok: false, error: 'Fila no encontrada: ' + id };

  } catch (err) {
    if (!locked) return { ok: false, error: 'Escritura simultanea. Reintenta.' };
    return { ok: false, error: err.toString() };
  } finally {
    if (locked) { try { lock.releaseLock(); } catch(e2) {} }
  }
}

// ===========================================================================
//  DELETE
// ===========================================================================
function deleteRow(sheetName, id) {
  var book  = ss_();
  var sheet = book.getSheetByName(sheetName);
  if (!sheet) return { ok: false, error: 'Hoja no encontrada: ' + sheetName };

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { ok: false, error: 'Hoja vacia' };

  var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) {
      sheet.deleteRow(i + 2);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Fila no encontrada: ' + id };
}

// ===========================================================================
//  GENERACION DE IDs AUTOMATICOS
// ===========================================================================
function nextId_(sheet, sheetName, headers) {
  var PREFIX_MAP = {
    'PREINSCRIPCIONES':       'PREINSC-',
    'INSCRIPCIONES':          'INSC-',
    'FACTURAS':               'FAC-',
    'AMPA_SOCIOS':            'SOC-',
    'AMPA_ALUMNOS':           'ALU-',
    'AUSENCIAS':              'AUS-',
    'REGISTRO_HORARIO':       'RH-',
    'ASISTENCIAS_FUTBOL':     'AF-',
    'ASISTENCIAS_EXTRA':      'AE-',
    'SESIONES':               'SES-',
    'NOMINAS':                'NOM-',
    'ENCUESTAS_SATISFACCION': 'ENC-',
  };

  var prefix  = PREFIX_MAP[sheetName];
  var lastRow = sheet.getLastRow();

  if (prefix) {
    var num = 1;
    if (lastRow > 1) {
      var vals = sheet.getRange(2, 1, lastRow - 1, 1).getValues()
        .map(function(r) { return String(r[0]); })
        .filter(function(v) { return v.indexOf(prefix) === 0; })
        .map(function(v) { return parseInt(v.replace(prefix, ''), 10) || 0; });
      if (vals.length) num = Math.max.apply(null, vals) + 1;
    }
    return prefix + (num < 10 ? '00' + num : num < 100 ? '0' + num : String(num));
  } else {
    if (lastRow < 2) return 1;
    var nums = sheet.getRange(2, 1, lastRow - 1, 1).getValues()
      .map(function(r) { return parseInt(r[0], 10) || 0; });
    return nums.length ? Math.max.apply(null, nums) + 1 : 1;
  }
}

// ===========================================================================
//  NORMALIZACION DE TEXTO (elimina tildes para comparar)
// ===========================================================================
function stripAccents_(str) {
  return String(str || '')
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .toUpperCase();
}

// ===========================================================================
//  PREINSCRIPCION PUBLICA (desde inscripcion.html)
// ===========================================================================
function nuevaPreinscripcion(data) {
  var book  = ss_();
  var sheet = book.getSheetByName('PREINSCRIPCIONES');
  if (!sheet) return { ok: false, error: 'Hoja PREINSCRIPCIONES no encontrada' };

  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(function(h) { return String(h).trim(); });

  // Acepta claves en mayusculas (manual/API) y minusculas (formulario web)
  var d = data;
  var nombreNormRaw = d.NOMBRE_ALUMNO   || d.nombre_alumno   || '';
  var emailRaw      = d.EMAIL_TUTOR     || d.email_tutor1    || d.email_tutor    || '';
  var nombreNorm    = stripAccents_(nombreNormRaw);
  var emailNorm     = String(emailRaw).toLowerCase().trim();

  var existing = sheetToJSON(book, 'PREINSCRIPCIONES');
  var dup = existing.some(function(r) {
    if (String(r.ESTADO || '').toUpperCase() === 'RECHAZADA') return false;
    var mismoNombre = nombreNorm && stripAccents_(r.NOMBRE_ALUMNO) === nombreNorm;
    var mismoEmail  = emailNorm  && String(r.EMAIL_TUTOR || '').toLowerCase() === emailNorm;
    return mismoNombre && mismoEmail;
  });

  if (dup) return { ok: false, error: 'Ya existe una preinscripcion para este alumno.' };

  var id = nextId_(sheet, 'PREINSCRIPCIONES', headers);

  var actividadInteres = d.ACTIVIDAD_INTERES || d.actividad_interes ||
    (Array.isArray(d.actividades) ? d.actividades.join(', ') : (d.actividades || ''));

  var rowObj = {
    ID:               id,
    FECHA_SOLICITUD:  Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    NOMBRE_ALUMNO:    String(nombreNormRaw).toUpperCase().trim(),
    APELLIDOS_ALUMNO: String(d.APELLIDOS_ALUMNO || d.apellidos_alumno || '').toUpperCase().trim(),
    FECHA_NAC:        d.FECHA_NAC   || d.fecha_nac   || '',
    CURSO_ALUMNO:     d.CURSO_ALUMNO|| d.curso_alumno|| '',
    CENTRO:           d.CENTRO      || d.centro      || 'ESCALERITAS',
    NOMBRE_TUTOR:     String(d.NOMBRE_TUTOR    || d.nombre_tutor1    || d.nombre_tutor    || '').toUpperCase().trim(),
    APELLIDOS_TUTOR:  String(d.APELLIDOS_TUTOR || d.apellidos_tutor1 || d.apellidos_tutor || '').toUpperCase().trim(),
    RELACION_TUTOR:   d.RELACION_TUTOR || d.relacion_tutor1 || '',
    NIF_TUTOR:        d.NIF_TUTOR      || d.nif_tutor1      || '',
    EMAIL_TUTOR:      emailNorm,
    'TELÉFONO_TUTOR': d.TELEFONO_TUTOR || d['TELÉFONO_TUTOR'] || d.telefono_tutor1 || '',
    ACTIVIDAD_INTERES:actividadInteres,
    DIAS_DISPONIBLES: d.DIAS_DISPONIBLES || d.dias_disponibles || '',
    LINEA:            d.LINEA || (stripAccents_(actividadInteres).indexOf('FUTBOL') !== -1 ? 'FUTBOL' : 'EXTRAESCOLARES'),
    AUTORIZA_IMAGEN:  d.AUTORIZA_IMAGEN || d.autoriza_imagen || 'NO',
    SALUD:            d.SALUD || d.salud || '',
    ESTADO:           'RECIBIDA',
    NOTAS:            d.NOTAS || d.observaciones || d.notas || '',
    CANAL:            d.CANAL || (d.nombre_alumno ? 'WEB' : 'MANUAL'),
  };

  var newRow = headers.map(function(h) {
    return rowObj[h] !== undefined ? rowObj[h] : '';
  });
  sheet.appendRow(newRow);

  return { ok: true, id: id };
}

// ===========================================================================
//  ALERTAS DE IMPAGOS (ejecutar manualmente o con trigger semanal)
// ===========================================================================
function sendAlertsImpagos() {
  var book     = ss_();
  var pagos    = sheetToJSON(book, 'PAGOS');
  var alumnos  = sheetToJSON(book, 'ALUMNOS');
  var clientes = sheetToJSON(book, 'CLIENTES');
  var enviados = 0;

  pagos
    .filter(function(p) { return String(p.ESTADO || '').toUpperCase() === 'PENDIENTE'; })
    .forEach(function(p) {
      var alumno  = alumnos.find(function(a) { return String(a.ID_ALUMNO) === String(p.ID_ALUMNO); });
      var cliente = alumno ? clientes.find(function(c) { return String(c.ID_CLIENTE) === String(alumno.ID_CLIENTE); }) : null;
      if (cliente && cliente.EMAIL) {
        MailApp.sendEmail({
          to:      cliente.EMAIL,
          subject: 'SM Academia - Recordatorio de pago pendiente',
          body:    'Estimado/a ' + cliente.NOMBRE + ',\n\n' +
                   'Tiene un pago pendiente:\n' +
                   '  Concepto: ' + (p.CONCEPTO || '-') + '\n' +
                   '  Mes:      ' + (p.MES || '-') + '\n' +
                   '  Importe:  ' + (p.TOTAL_FACTURADO || '-') + ' EUR\n\n' +
                   'Contacte con nosotros para regularizarlo.\n\nSM Academia',
        });
        enviados++;
      }
    });

  return { ok: true, alertas: enviados };
}

// ===========================================================================
//  DIAGNOSTICO - ejecutar desde el editor para ver errores
// ===========================================================================
function diagnostico() {
  try {
    var book   = ss_();
    var sheets = book.getSheets().map(function(s) { return s.getName(); });
    Logger.log('Hojas encontradas: ' + JSON.stringify(sheets));

    var esperadas = [
      'CLIENTES','ALUMNOS','ACTIVIDADES','PREINSCRIPCIONES','INSCRIPCIONES',
      'PAGOS','FACTURAS','PERSONAL','AUSENCIAS','REGISTRO_HORARIO',
      'SESIONES','ASISTENCIAS_FUTBOL','ASISTENCIAS_EXTRA','NOMINAS',
      'GASTOS','MOVIMIENTOS_BANCARIOS','AMPA_SOCIOS','AMPA_ALUMNOS',
      'ENCUESTAS_SATISFACCION',
    ];

    var faltantes = esperadas.filter(function(n) { return sheets.indexOf(n) === -1; });
    if (faltantes.length) {
      Logger.log('HOJAS FALTANTES: ' + JSON.stringify(faltantes));
    } else {
      Logger.log('OK: Todas las hojas encontradas');
    }

    var result = getAllData();
    Logger.log('GET_ALL ok=' + result.ok);
    Logger.log('Clientes: '  + result.clientes.length);
    Logger.log('Personal: '  + result.personal.length);
    Logger.log('Encuestas: ' + result.encuestasSatisfaccion.length);

  } catch(err) {
    Logger.log('ERROR: ' + err.toString());
  }
}
