// ===========================================================================
//  SM ACADEMIA - Google Apps Script v2.5
//  Sheets ID: 1oHJIUoyR3V5N0iweWnMagZic_8YkWtVV6CsshXloX4k
// ===========================================================================

var SPREADSHEET_ID = '1oHJIUoyR3V5N0iweWnMagZic_8YkWtVV6CsshXloX4k';

// ---------------------------------------------------------------------------
// Cuentas de correo de la empresa
// ---------------------------------------------------------------------------
var EMAIL_EMPRESA = {
  direccion:    'direccion@sm-academia.com',          // canal oficial + denuncia
  director:     'samy.martin@academiasmfutbol.com',   // dirección personal
  futbol:       'academiasmfutbol@outlook.es',        // academia fútbol
  extraescolar: 'academiasmextraescolares@outlook.es', // desuso
  asesoria:     'info@vgasesoria.es',
};

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

      case 'GET_NOMINA_PDF':
        result = getNominaPdf_(e.parameter.fileId);
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
    } else if (action === 'APPEND_BATCH') {
      result = appendBatch(body.sheet, body.rows || []);
    } else if (action === 'UPDATE') {
      result = updateRow(body.sheet, body.id, body.row || {});
    } else if (action === 'DELETE') {
      result = deleteRow(body.sheet, body.id);
    } else if (action === 'DELETE_BATCH') {
      result = deleteBatch(body.sheet, body.ids || []);
    } else if (action === 'NUEVA_PREINSCRIPCION') {
      result = nuevaPreinscripcion(body);
    } else if (action === 'SEND_FACTURA_EMAIL') {
      result = sendFacturaEmail(body);
    } else if (action === 'LISTAR_DRIVE') {
      result = listarDrive_(body.folderId || '');
    } else if (action === 'GENERAR_PDF_FIRMADO') {
      result = generarPDFFirmado_(body.firmaId);
    } else if (action === 'ENVIAR_DOCS_ASESORIA') {
      result = enviarDocsAsesoria_();
    } else if (action === 'GET_FACTURAS_DRIVE') {
      result = getFacturasDrive_();
    } else if (action === 'GET_DRIVE_FACTURAS_METADATA') {
      result = getDriveFacturasMetadata_();
    } else if (action === 'GET_OLD_DB_FACTURAS') {
      result = getOldDBFacturas_();
    } else if (action === 'IMPORTAR_FACTURAS_HISTORICAS') {
      result = importarFacturasHistoricas_(body.filas || []);
    } else if (action === 'PREPARAR_ASESORIA_DRIVE') {
      result = prepararAsesoriaEnDrive_(body.anyo || 2026);
    } else if (action === 'OBTENER_PDF_B64') {
      result = obtenerPdfB64_(body.fileId);
    } else if (action === 'GUARDAR_PDF_FIRMADO_PERSONAL') {
      result = guardarPdfFirmadoPersonal_(body);
    } else if (action === 'SUBIR_CNDS') {
      result = subirCNDS_(body);
    } else if (action === 'SUBIR_NOMINA_PDF') {
      result = subirNominaPdf_(body);
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
    solicitudes:             sheetToJSON(book, 'SOLICITUDES'),
    bancoHistorico:          sheetToJSON(book, 'BANCO_HISTORICO'),
    pagadoresExternos:       sheetToJSON(book, 'PAGADORES_EXTERNOS'),
    firmasProtocolo:         sheetToJSON(book, 'FIRMAS_PROTOCOLO'),
  };
}

// Ejecutar UNA VEZ desde el editor GAS para convertir campos de dinero de texto a número.
// SEGURO: solo procesa celdas cuyo valor es un string puro (typeof === 'string').
// Números ya almacenados como número y fechas (Date objects) se dejan intactos.
function migrarImportesANumero() {
  var book = ss_();
  var CONFIG = [
    { sheet: 'PAGOS',         cols: ['IMPORTE','DESC_EUR','TOTAL_ESPERADO','TOTAL_RECIBIDO'] },
    { sheet: 'INSCRIPCIONES', cols: ['PRECIO_ACORDADO','DESCUENTO'] },
    { sheet: 'ACTIVIDADES',   cols: ['PRECIO'] },
    { sheet: 'FACTURAS',      cols: ['BASE_IMPONIBLE','DESCUENTO','TOTAL'] },
  ];
  var totalFixed = 0;
  CONFIG.forEach(function(cfg) {
    var sheet = book.getSheetByName(cfg.sheet);
    if (!sheet) return;
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 2) return;
    var headers = sheet.getRange(1,1,1,lastCol).getValues()[0].map(function(h){ return String(h).trim(); });
    var colIdxs = [];
    cfg.cols.forEach(function(col){ var idx=headers.indexOf(col); if(idx>=0) colIdxs.push(idx); });
    if (!colIdxs.length) return;
    var data = sheet.getRange(2,1,lastRow-1,lastCol).getValues();
    data.forEach(function(row, ri) {
      colIdxs.forEach(function(ci) {
        var v = row[ci];
        if (typeof v !== 'string' || v.trim() === '') return; // solo strings, nunca fechas ni números
        var n = parseFloat(v.replace(',','.'));
        if (isNaN(n) || !isFinite(n)) return;
        sheet.getRange(ri+2, ci+1).setValue(n);
        totalFixed++;
      });
    });
    Logger.log('Procesada: ' + cfg.sheet);
  });
  Logger.log('Total celdas convertidas: ' + totalFixed);
}

// Ejecutar UNA VEZ desde el editor GAS para migrar columnas HORA a texto plano
function migrarHorasATexto() {
  var book = ss_();
  var sheets = ['SESIONES', 'ACTIVIDADES', 'REGISTRO_HORARIO'];
  sheets.forEach(function(sName) {
    var sheet = book.getSheetByName(sName);
    if (!sheet) return;
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 2) return;
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var horaIdxs = [];
    headers.forEach(function(h, i) { if (String(h).indexOf('HORA') !== -1) horaIdxs.push(i); });
    if (!horaIdxs.length) return;
    var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    data.forEach(function(row, ri) {
      horaIdxs.forEach(function(ci) {
        var v = row[ci];
        if (v instanceof Date) {
          var formatted = Utilities.formatDate(v, 'Atlantic/Canary', 'HH:mm');
          var cell = sheet.getRange(ri + 2, ci + 1);
          cell.setNumberFormat('@STRING@');
          cell.setValue(formatted);
        }
      });
    });
    Logger.log('Migrado: ' + sName);
  });
  Logger.log('Migración completada');
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

  var range        = sheet.getRange(1, 1, lastRow, lastCol);
  var data         = range.getValues();
  var displayData  = range.getDisplayValues();
  var headers      = data[0].map(function(h) { return String(h).trim(); });
  // Columnas que deben leerse siempre como texto (listas de IDs separadas por coma)
  var TEXT_COLS    = { 'ALUMNOS_PRESENTES': true };
  var rows         = [];

  for (var i = 1; i < data.length; i++) {
    var row        = data[i];
    var displayRow = displayData[i];
    var isEmpty    = row.every(function(c) { return c === '' || c === null || c === undefined; });
    if (isEmpty) continue;

    var obj = {};
    headers.forEach(function(h, j) {
      var v = row[j];
      if (v instanceof Date) {
        if (h.indexOf('HORA') !== -1) {
          // Usar 'Atlantic/Canary' para evitar offset con 'Europe/Madrid' (Islas Canarias UTC+0/+1)
          obj[h] = Utilities.formatDate(v, 'Atlantic/Canary', 'HH:mm');
        } else {
          obj[h] = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        }
      } else if (TEXT_COLS[h]) {
        // Leer valor mostrado: locale español convierte "47,12" en número 47.12;
        // getDisplayValues() devuelve "47,12" respetando el formato regional
        obj[h] = displayRow[j] || '';
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
      if (/^(IMPORTE|PRECIO|TOTAL_ESPERADO|TOTAL_RECIBIDO|DESC_EUR|DESCUENTO|CUOTA|PRECIO_ACORDADO)$/.test(h) && v !== '') {
        var n = parseFloat(v); if (!isNaN(n)) return n;
      }
      return v;
    });

    sheet.appendRow(newRow);

    // Forzar columnas HORA y ALUMNOS_PRESENTES a texto plano
    // (evita que Sheets reinterprete "47,12" como número 47.12 en locale español)
    var newRowNum = sheet.getLastRow();
    headers.forEach(function(h, i) {
      if (h.indexOf('HORA') !== -1 || h === 'ALUMNOS_PRESENTES') {
        sheet.getRange(newRowNum, i + 1).setNumberFormat('@STRING@');
      }
    });

    return { ok: true, id: rowObj[idCol] || '' };

  } catch (err) {
    if (!locked) return { ok: false, error: 'Escritura simultanea. Reintenta.' };
    return { ok: false, error: err.toString() };
  } finally {
    if (locked) { try { lock.releaseLock(); } catch(e2) {} }
  }
}

// ===========================================================================
//  APPEND_BATCH - escribe múltiples filas en una sola llamada (un único lock)
// ===========================================================================
function appendBatch(sheetName, rows) {
  if (!rows || !rows.length) return { ok: true, count: 0 };
  var lock   = LockService.getScriptLock();
  var locked = false;
  try {
    lock.waitLock(30000);
    locked = true;

    var book  = ss_();
    var sheet = book.getSheetByName(sheetName);

    // Auto-crear la hoja con cabeceras si no existe
    if (!sheet) {
      sheet = book.insertSheet(sheetName);
      var firstRow = rows[0];
      var cols = Object.keys(firstRow);
      // Añadir columna ID como primera si no está
      if (cols[0] !== 'ID') cols.unshift('ID');
      sheet.appendRow(cols);
    }

    var lastCol  = sheet.getLastColumn();
    if (lastCol < 1) return { ok: false, error: 'Sin cabeceras: ' + sheetName };

    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
      .map(function(h) { return String(h).trim(); });

    var idCol    = headers[0];
    var lastRow  = sheet.getLastRow();
    var nextNum  = lastRow;  // fila actual para generar IDs secuenciales

    var matrix = [];
    var generatedIds = [];

    rows.forEach(function(rowObj) {
      if (idCol && !rowObj[idCol]) {
        nextNum++;
        rowObj[idCol] = sheetName + '-' + String(nextNum).padStart(4, '0');
      }
      var newRow = headers.map(function(h) {
        var v = (rowObj[h] !== undefined && rowObj[h] !== null) ? rowObj[h] : '';
        if (h.indexOf('HORA') !== -1 && v !== '') return String(v);
        if (/^(IMPORTE|PRECIO|TOTAL_ESPERADO|TOTAL_RECIBIDO|DESC_EUR|DESCUENTO|CUOTA|PRECIO_ACORDADO)$/.test(h) && v !== '') {
          var n = parseFloat(v); if (!isNaN(n)) return n;
        }
        return v;
      });
      matrix.push(newRow);
      generatedIds.push(rowObj[idCol] || '');
    });

    if (matrix.length) {
      var startRow = sheet.getLastRow() + 1;
      sheet.getRange(startRow, 1, matrix.length, matrix[0].length).setValues(matrix);
      // Forzar texto en columnas que no deben ser numéricas (e.g. ALUMNOS_PRESENTES)
      headers.forEach(function(h, i) {
        if (h === 'ALUMNOS_PRESENTES' || h.indexOf('HORA') !== -1) {
          sheet.getRange(startRow, i + 1, matrix.length, 1).setNumberFormat('@STRING@');
        }
      });
    }

    return { ok: true, count: matrix.length, ids: generatedIds };

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
            if (/^(IMPORTE|PRECIO|TOTAL_ESPERADO|TOTAL_RECIBIDO|DESC_EUR|DESCUENTO|CUOTA|PRECIO_ACORDADO)$/.test(h) && rowObj[h] !== '') {
              var n = parseFloat(rowObj[h]); if (!isNaN(n)) return n;
            }
            return rowObj[h];
          }
          return data[i][j];
        });
        sheet.getRange(i + 1, 1, 1, headers.length).setValues([updatedRow]);

        // Forzar columnas HORA a texto plano tras actualizar
        headers.forEach(function(h, j) {
          if (h.indexOf('HORA') !== -1) {
            sheet.getRange(i + 1, j + 1).setNumberFormat('@STRING@');
          }
        });

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
//  DELETE_BATCH - elimina múltiples filas por ID en una pasada
// ===========================================================================
function deleteBatch(sheetName, ids) {
  if (!ids || !ids.length) return { ok: true, count: 0 };
  var book  = ss_();
  var sheet = book.getSheetByName(sheetName);
  if (!sheet) return { ok: false, error: 'Hoja no encontrada: ' + sheetName };

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { ok: true, count: 0 };

  var idSet = {};
  ids.forEach(function(id) { idSet[String(id)] = true; });

  var colIds = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  // Recorrer de abajo a arriba para que deleteRow no cambie índices
  var deleted = 0;
  for (var i = colIds.length - 1; i >= 0; i--) {
    if (idSet[String(colIds[i][0])]) {
      sheet.deleteRow(i + 2);
      deleted++;
    }
  }
  return { ok: true, count: deleted };
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
    'PAGADORES_EXTERNOS':     'PE-',
    'FIRMAS_PROTOCOLO':       'FIRMA-',
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
    .replace(/[\u0300-\u036f]/g, '')
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

// ===========================================================================
//  ENVÍO DE AVISO DE FACTURA POR EMAIL
// ===========================================================================
function sendFacturaEmail(body) {
  var to      = body.email;
  var nombre  = body.nombre  || 'cliente';
  var numFac  = body.numFac  || '';
  var total   = body.total   || '';
  var portal  = 'https://smacademia.github.io/smAPP360/';

  if (!to) return { ok: false, error: 'Sin email destinatario' };

  var asunto = 'Factura disponible · ' + numFac + ' — SM Academia';

  var html = '<div style="font-family:Arial,sans-serif;max-width:520px;margin:0 auto;color:#222">'
    + '<div style="background:#2e7d32;padding:24px 32px;border-radius:8px 8px 0 0">'
    +   '<h1 style="color:#fff;margin:0;font-size:22px">SM Academia</h1>'
    + '</div>'
    + '<div style="background:#f9f9f9;padding:28px 32px;border-radius:0 0 8px 8px;border:1px solid #e0e0e0;border-top:none">'
    +   '<p>Hola <strong>' + nombre + '</strong>,</p>'
    +   '<p>Ya tienes disponible tu última factura en la app <strong>SM APP 360°</strong>.</p>'
    +   '<table style="margin:20px 0;border-collapse:collapse;width:100%">'
    +     '<tr><td style="padding:6px 0;color:#666;width:140px">Número de factura</td>'
    +         '<td style="padding:6px 0;font-weight:bold">' + numFac + '</td></tr>'
    +     (total ? '<tr><td style="padding:6px 0;color:#666">Importe</td>'
    +              '<td style="padding:6px 0;font-weight:bold">' + total + ' €</td></tr>' : '')
    +   '</table>'
    +   '<p style="margin:28px 0 8px">'
    +     '<a href="' + portal + '" style="background:#2e7d32;color:#fff;padding:12px 28px;'
    +     'border-radius:6px;text-decoration:none;font-size:15px;background:#2e7d32">Acceder a la app</a>'
    +   '</p>'
    +   '<p style="margin-top:32px;font-size:12px;color:#999">SM Academia de Fútbol · '
    +   'Este mensaje es informativo. Puedes consultar y descargar tu factura desde el portal de familias.</p>'
    + '</div>'
    + '</div>';

  try {
    MailApp.sendEmail({
      to:       to,
      replyTo:  EMAIL_EMPRESA.direccion,
      name:     'SM Academia',
      subject:  asunto,
      htmlBody: html,
    });
    return { ok: true };
  } catch(err) {
    return { ok: false, error: err.toString() };
  }
}

// ===========================================================================
//  EXPLORADOR DE GOOGLE DRIVE
// ===========================================================================
var MIME_VISIBLES_ = [
  'application/vnd.google-apps.document',
  'application/vnd.google-apps.presentation',
  'application/vnd.google-apps.spreadsheet',
  'application/pdf',
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
];

function listarDrive_(folderId) {
  var folder;
  try {
    folder = folderId ? DriveApp.getFolderById(folderId) : DriveApp.getRootFolder();
  } catch(e) {
    return { ok: false, error: 'Carpeta no accesible: ' + e.message };
  }

  var items = [];

  var subs = folder.getFolders();
  while (subs.hasNext()) {
    var f = subs.next();
    items.push({ tipo: 'carpeta', id: f.getId(), nombre: f.getName() });
  }

  var files = folder.getFiles();
  while (files.hasNext()) {
    var f = files.next();
    if (MIME_VISIBLES_.indexOf(f.getMimeType()) < 0) continue;
    var id   = f.getId();
    var mime = f.getMimeType();
    var url  = f.getUrl();
    var prev = (mime === 'application/vnd.google-apps.document')
      ? 'https://docs.google.com/document/d/' + id + '/preview'
      : 'https://drive.google.com/file/d/' + id + '/preview';
    items.push({ tipo: 'archivo', id: id, nombre: f.getName(), url: url, previewUrl: prev, mime: mime });
  }

  items.sort(function(a, b) {
    if (a.tipo !== b.tipo) return a.tipo === 'carpeta' ? -1 : 1;
    return a.nombre.localeCompare(b.nombre, 'es');
  });

  var padre = folderId ? folder.getParents() : null;
  var padreId = (padre && padre.hasNext()) ? padre.next().getId() : '';

  return {
    ok:      true,
    id:      folder.getId(),
    nombre:  folder.getName(),
    padreId: padreId,
    items:   items,
  };
}

// ===========================================================================
//  GENERACIÓN DE PDF FIRMADO
// ===========================================================================
var PROTO_GDOC_IDS_ = {
  'lgtbiq':                '1nev6zQDFju1dDCKaUZc6wtcesD1_sG6UnxqHmWA8XCI',
  'manual-empleado':       '1fE_6Lk2QJhQ8-SG_ytILdqn39g7pBRf0ag_eQmcQMTg',
  'operativo-escaleritas': '13FOWFnuadOOpniQ14TB9wG_Y1LXOgkap',
};

function generarPDFFirmado_(firmaId) {
  var book   = ss_();
  var firmas = sheetToJSON(book, 'FIRMAS_PROTOCOLO');
  var firma  = null;
  for (var i = 0; i < firmas.length; i++) {
    if (String(firmas[i].ID) === String(firmaId)) { firma = firmas[i]; break; }
  }
  if (!firma) return { ok: false, error: 'Firma no encontrada: ' + firmaId };

  var nombre    = firma.NOMBRE    || '';
  var fechaRaw  = firma.FECHA     || '';
  var hora      = firma.HORA      || '';
  var tituloDoc = firma.TITULO_DOC || '';
  var docId     = firma.DOC_ID    || '';
  var rawB64    = String(firma.FIRMA_IMG || '').replace(/^data:image\/[a-z]+;base64,/, '');

  // Normalizar fecha a DD/MM/YYYY para mostrar en el PDF
  // Acepta ISO (2026-10-09), DD/MM/YYYY (09/10/2026) y DD/MM/YY (09/10/26)
  var fecha = fechaRaw;
  var isoMatch = fechaRaw.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (isoMatch) {
    fecha = isoMatch[3] + '/' + isoMatch[2] + '/' + isoMatch[1];
  }

  if (!rawB64) return { ok: false, error: 'Imagen de firma no encontrada' };

  var imgBlob;
  try {
    imgBlob = Utilities.newBlob(Utilities.base64Decode(rawB64), 'image/jpeg', 'firma.jpg');
  } catch(e) {
    try {
      imgBlob = Utilities.newBlob(Utilities.base64Decode(rawB64), 'image/png', 'firma.png');
    } catch(e2) {
      return { ok: false, error: 'Error decodificando firma: ' + e2.message };
    }
  }

  var driveFileId = PROTO_GDOC_IDS_[docId];
  var newDoc;
  try {
    if (driveFileId) {
      // Copiar el Google Doc original e inyectar la firma
      var copia = DriveApp.getFileById(driveFileId).makeCopy(
        'Firmado_' + nombre.replace(/\s+/g,'_') + '_' + fechaRaw.replace(/\//g,'-')
      );
      newDoc = DocumentApp.openById(copia.getId());
      var body = newDoc.getBody();

      // Reemplazar placeholder con bloque de firma
      var found = body.findText('\\{\\{FIRMA_TRABAJADOR\\}\\}');
      if (found) {
        var textEl = found.getElement();
        var parent = textEl.getParent();
        var parentType = parent.getType();

        if (parentType === DocumentApp.ElementType.PARAGRAPH &&
            parent.getParent().getType() === DocumentApp.ElementType.TABLE_CELL) {
          // Placeholder dentro de celda de tabla
          var cell     = parent.getParent();
          var row      = cell.getParent();
          var table    = row.getParent();
          var rowIndex = table.getChildIndex(row);

          // Rellenar las demás celdas del row (Nombre | DNI | Fecha | Firma)
          if (row.getNumCells() >= 4) {
            var c0 = row.getCell(0); if (!c0.getText().trim()) c0.setText(nombre);
            var c1 = row.getCell(1); if (!c1.getText().trim()) c1.setText('—');
            var c2 = row.getCell(2); if (!c2.getText().trim()) c2.setText(fecha + ' ' + hora);
          } else if (row.getNumCells() >= 1) {
            row.getCell(0).setText(nombre + '  ·  ' + fecha + ' ' + hora);
          }

          // Sustituir celda de firma con imagen
          cell.clear();
          cell.appendParagraph('').appendInlineImage(imgBlob).setWidth(110).setHeight(42);

          // Eliminar filas vacías que quedan debajo
          var nRows = table.getNumRows();
          for (var r = nRows - 1; r > rowIndex; r--) {
            var testRow = table.getRow(r);
            var isEmpty = true;
            for (var c = 0; c < testRow.getNumCells(); c++) {
              if (testRow.getCell(c).getText().trim()) { isEmpty = false; break; }
            }
            if (isEmpty) table.removeRow(r);
          }

        } else if (parentType === DocumentApp.ElementType.PARAGRAPH) {
          // Placeholder en párrafo libre del body
          // No se puede usar removeChild si es el último párrafo — reemplazar texto en su lugar
          var idx = body.getChildIndex(parent);
          parent.replaceText('\\{\\{FIRMA_TRABAJADOR\\}\\}', 'Firmado digitalmente por: ' + nombre);
          parent.setSpacingBefore(6);
          body.insertParagraph(idx + 1, 'Fecha: ' + fecha + '  ·  Hora: ' + hora);
          body.insertParagraph(idx + 2, '').appendInlineImage(imgBlob).setWidth(190).setHeight(65);
        }
      } else {
        // Sin placeholder: añadir al final
        body.appendHorizontalRule();
        body.appendParagraph('FIRMA DEL/LA TRABAJADOR/A').setBold(true);
        body.appendParagraph(nombre + '  ·  ' + fecha + '  ·  ' + hora);
        body.appendParagraph('').appendInlineImage(imgBlob).setWidth(190).setHeight(65);
      }
      newDoc.saveAndClose();
    } else {
      // Sin Google Doc: generar certificado
      newDoc = DocumentApp.create(
        'Cert_Firma_' + nombre.replace(/\s+/g,'_') + '_' + fecha.replace(/\//g,'-')
      );
      var body = newDoc.getBody();
      body.appendParagraph('SM ACADEMIA SPORTS & FOOTBALL')
        .setHeading(DocumentApp.ParagraphHeading.HEADING1)
        .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      body.appendParagraph('Certificado de firma de protocolo')
        .setHeading(DocumentApp.ParagraphHeading.HEADING2)
        .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      body.appendHorizontalRule();
      body.appendParagraph('DOCUMENTO: ' + tituloDoc);
      body.appendParagraph('');
      body.appendParagraph('Nombre y apellidos: ' + nombre);
      body.appendParagraph('Fecha: ' + fecha + '  ·  Hora: ' + hora);
      body.appendParagraph('');
      body.appendParagraph('Firma:');
      body.appendParagraph('').appendInlineImage(imgBlob).setWidth(190).setHeight(65);
      body.appendHorizontalRule();
      body.appendParagraph('Documento generado por SM Academia · ' + fecha)
        .setFontSize(8).setItalic(true);
      newDoc.saveAndClose();
    }

    // Exportar como PDF
    var pdfBlob = DriveApp.getFileById(newDoc.getId()).getAs(MimeType.PDF);
    pdfBlob.setName('Protocolo_' + nombre.replace(/\s+/g,'_') + '_' + fecha.replace(/\//g,'-') + '.pdf');

    var folder  = _carpetaFirmados_();
    var pdfFile = folder.createFile(pdfBlob);
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // Borrar copia intermedia
    DriveApp.getFileById(newDoc.getId()).setTrashed(true);

    var pdfId  = pdfFile.getId();
    var pdfUrl = pdfFile.getUrl();

    // Actualizar fila en FIRMAS_PROTOCOLO con PDF_ID y PDF_URL
    var sheet = book.getSheetByName('FIRMAS_PROTOCOLO');
    if (sheet && sheet.getLastRow() > 1) {
      var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      var colPdfId  = headers.indexOf('PDF_ID')  + 1;
      var colPdfUrl = headers.indexOf('PDF_URL') + 1;
      var colId     = headers.indexOf('ID')      + 1;
      if (colId > 0) {
        var vals = sheet.getRange(2, colId, sheet.getLastRow() - 1, 1).getValues();
        for (var r = 0; r < vals.length; r++) {
          if (String(vals[r][0]) === String(firmaId)) {
            if (colPdfId  > 0) sheet.getRange(r + 2, colPdfId ).setValue(pdfId);
            if (colPdfUrl > 0) sheet.getRange(r + 2, colPdfUrl).setValue(pdfUrl);
            break;
          }
        }
      }
    }

    return { ok: true, pdfId: pdfId, pdfUrl: pdfUrl };

  } catch(e) {
    return { ok: false, error: e.message };
  }
}

function _carpetaFirmados_() {
  var name = 'SM Academia — Documentos Firmados';
  var it   = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}

// ===========================================================================
//  ASESORÍA — LISTADO DE PDFs Y PREPARACIÓN DE CARPETA DRIVE
// ===========================================================================
var DRIVE_FAC_EMITIDAS_  = '1PqKU6DfAElQigDPDkBLnuNEklSZplRBn'; // carpeta FACTURAS
var DRIVE_FAC_RECIBIDAS_ = '1i09bvwehgbDlaywAu1ED_M0J-V_j3Oqv'; // carpeta PROOVEDORES FACTURAS
var OLD_SS_ID_           = '1nx1wbPyUKDvjsvGUSKaDjClF526S46WDGfyZXkLc4Hc'; // app anterior
var DRIVE_SM_ROOT_       = '1cnL0es8vW8wpb-Dcv0BrYyc5FkkgQCH7'; // carpeta SM ACADEMIA

function getFacturasDrive_() {
  try {
    var folder = DriveApp.getFolderById(DRIVE_FAC_EMITIDAS_);
    var files  = folder.getFiles();
    var list   = [];
    while (files.hasNext()) {
      var f = files.next();
      list.push({
        id:     f.getId(),
        nombre: f.getName(),
        fecha:  Utilities.formatDate(f.getDateCreated(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        url:    f.getUrl(),
        bytes:  f.getSize()
      });
    }
    list.sort(function(a,b){ return a.nombre.localeCompare(b.nombre,'es'); });
    return { ok: true, files: list };
  } catch(e) {
    return { ok: false, error: e.toString() };
  }
}

function prepararAsesoriaEnDrive_(anyo) {
  anyo = parseInt(anyo) || 2026;
  try {
    var raiz   = DriveApp.getFolderById(DRIVE_SM_ROOT_);
    var nombre = 'Asesoría ' + anyo + ' · SM Academia';
    var dest   = _goc_(raiz, nombre);

    // ── Ingresos ──
    var ingF  = _goc_(dest, '📥 INGRESOS – Facturas emitidas');
    var tiFs  = [
      _goc_(ingF, 'T1 · Enero–Marzo'),
      _goc_(ingF, 'T2 · Abril–Junio'),
      _goc_(ingF, 'T3 · Julio–Septiembre'),
      _goc_(ingF, 'T4 · Octubre–Diciembre'),
    ];
    var facF     = DriveApp.getFolderById(DRIVE_FAC_EMITIDAS_);
    var fFiles   = facF.getFiles();
    var nIng = 0, nDupes = 0;
    var vistos = {};
    while (fFiles.hasNext()) {
      var ff      = fFiles.next();
      var created = ff.getDateCreated();
      if (created.getFullYear() !== anyo) continue;
      var n = ff.getName();
      if (vistos[n]) { nDupes++; continue; }
      vistos[n] = true;
      var mes = created.getMonth() + 1;
      var ti  = Math.min(3, Math.floor((mes - 1) / 3));
      ff.makeCopy(n, tiFs[ti]);
      nIng++;
    }

    // ── Gastos ──
    var gasF   = _goc_(dest, '📤 GASTOS – Facturas recibidas');
    var provF  = DriveApp.getFolderById(DRIVE_FAC_RECIBIDAS_);
    var provIt = provF.getFolders();
    var nGas   = 0;
    while (provIt.hasNext()) {
      var pf    = provIt.next();
      var dest2 = _goc_(gasF, pf.getName());
      var pFils = pf.getFiles();
      while (pFils.hasNext()) {
        var pf2 = pFils.next();
        if (pf2.getDateCreated().getFullYear() !== anyo) continue;
        pf2.makeCopy(pf2.getName(), dest2);
        nGas++;
      }
    }

    return {
      ok:             true,
      url:            dest.getUrl(),
      nombre:         nombre,
      ingresosCopiadas: nIng,
      gastosCopiadas:   nGas,
      duplicadosSaltados: nDupes
    };
  } catch(e) {
    return { ok: false, error: e.toString() };
  }
}

function _goc_(parent, name) {
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

// ---------------------------------------------------------------------------
//  Metadatos de PDFs de Drive + detección de ya importados
// ---------------------------------------------------------------------------
function getDriveFacturasMetadata_() {
  try {
    var folder = DriveApp.getFolderById(DRIVE_FAC_EMITIDAS_);
    var files  = folder.getFiles();

    // Números ya en la hoja FACTURAS (serie ANT = importadas antes)
    var sheet    = ss_().getSheetByName('FACTURAS');
    var yaEnDB   = {};
    if (sheet && sheet.getLastRow() > 1) {
      var vals = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
      var hdrs = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      var iSerie = hdrs.indexOf('SERIE');
      var iNum   = hdrs.indexOf('NUMERO');
      if (iSerie >= 0 && iNum >= 0) {
        vals.forEach(function(r) {
          if (String(r[iSerie]).toUpperCase() === 'ANT') {
            yaEnDB[parseInt(r[iNum]) || 0] = true;
          }
        });
      }
    }

    var list = [];
    while (files.hasNext()) {
      var f = files.next();
      if (f.getMimeType() !== 'application/pdf') continue;
      var nombre = f.getName();

      // Extraer número: "Factura_NNN_NOMBRE.pdf" ó "Factura NNN.pdf" ó "Factura NNN NOMBRE.pdf"
      var mNum = nombre.match(/Factura[_ ](\d+)/i);
      if (!mNum) continue;
      var num = parseInt(mNum[1]);
      if (!num) continue;

      // Extraer cliente: texto después del número, sin extensión
      var mCli = nombre.match(/Factura[_ ]\d+[_ ](.+?)\.pdf$/i);
      var cliente = mCli ? mCli[1].replace(/_/g, ' ').trim() : '';

      var fechaCreada = f.getDateCreated();
      list.push({
        id:         f.getId(),
        nombre:     nombre,
        numero:     num,
        cliente:    cliente,
        fecha:      Utilities.formatDate(fechaCreada, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        url:        f.getUrl(),
        yaImportada: yaEnDB[num] === true
      });
    }

    list.sort(function(a, b) { return a.numero - b.numero; });
    return { ok: true, files: list };
  } catch(e) {
    return { ok: false, error: e.toString() };
  }
}

// ---------------------------------------------------------------------------
//  Lectura del spreadsheet antiguo para cruzar importes
// ---------------------------------------------------------------------------
function getOldDBFacturas_() {
  try {
    var ss     = SpreadsheetApp.openById(OLD_SS_ID_);
    var sheets = ss.getSheets();
    var hojas  = sheets.map(function(s) { return s.getName(); });
    var result = [];

    sheets.forEach(function(sheet) {
      if (sheet.getLastRow() < 2) return;
      var vals = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
      var hdrs = vals[0].map(function(h) { return String(h || '').trim().toUpperCase(); });

      // Detectar columnas relevantes
      var colNum = -1, colTotal = -1, colCli = -1, colFecha = -1;
      hdrs.forEach(function(h, i) {
        if (colNum   < 0 && /N[ÚU]MERO|^NUM$|^N[º°]$|FACTURA/.test(h)) colNum   = i;
        if (colTotal < 0 && /TOTAL|IMPORTE|PRECIO|CUOTA|COBRADO/.test(h) && !/DESC/.test(h)) colTotal = i;
        if (colCli   < 0 && /CLIENTE|NOMBRE|ALUMNO|FAMILIA|TITULAR/.test(h)) colCli = i;
        if (colFecha < 0 && /FECHA/.test(h)) colFecha = i;
      });

      if (colTotal < 0) return; // sin columna de importe → ignorar hoja

      for (var i = 1; i < vals.length; i++) {
        var row   = vals[i];
        var total = parseFloat(String(row[colTotal]).replace(',', '.')) || 0;
        if (!total) continue;

        var num = null;
        if (colNum >= 0) {
          var raw = String(row[colNum] || '').replace(/[^\d]/g, '');
          num = raw ? (parseInt(raw) || null) : null;
        }
        var fecha = '';
        if (colFecha >= 0 && row[colFecha]) {
          fecha = row[colFecha] instanceof Date
            ? Utilities.formatDate(row[colFecha], Session.getScriptTimeZone(), 'yyyy-MM-dd')
            : String(row[colFecha]).substring(0, 10);
        }
        result.push({
          hoja:    sheet.getName(),
          numero:  num,
          cliente: colCli >= 0 ? String(row[colCli] || '').trim() : '',
          fecha:   fecha,
          total:   total
        });
      }
    });

    return { ok: true, registros: result, hojas: hojas };
  } catch(e) {
    return { ok: false, error: e.toString() };
  }
}

// ---------------------------------------------------------------------------
//  Importación masiva de facturas históricas (serie ANT)
// ---------------------------------------------------------------------------
function importarFacturasHistoricas_(filas) {
  try {
    var sheet = ss_().getSheetByName('FACTURAS');
    if (!sheet) return { ok: false, error: 'Hoja FACTURAS no encontrada' };

    var hdrs = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var count = 0;

    filas.forEach(function(fila) {
      var num     = parseInt(fila.numero) || 0;
      var cliente = String(fila.cliente || '').trim();
      var fecha   = String(fila.fecha   || '').trim();
      var total   = parseFloat(fila.total) || 0;
      var driveId = String(fila.driveId  || '').trim();
      var urlDrv  = String(fila.url      || '').trim();
      if (!num || !fecha) return;

      var idFac = 'ANT-' + String(num).padStart(3, '0');

      var row = {};
      row['ID_FACTURA']     = idFac;
      row['FECHA_EXP']      = fecha;
      row['SERIE']          = 'ANT';
      row['NUMERO']         = num;
      row['TOTAL']          = total;
      row['BASE_IMPONIBLE'] = total;
      row['CLIENTE_NOMBRE'] = cliente;
      row['CONCEPTO']       = 'Servicio academia (histórico)';
      row['ESTADO_AEAT']    = 'COBRADA';
      row['TIPO_FACTURA']   = 'ORIGINAL';
      row['TIPO_IMPUESTO']  = 'EXENTO';
      row['CUOTA_IMPUESTO'] = 0;
      row['NOMBRE_EMISOR']  = 'SM ACADEMIA';
      row['NIF_EMISOR']     = '';
      row['HASH_ANTERIOR']  = '';
      row['HASH_ACTUAL']    = '';
      row['URL_PDF_DRIVE']  = urlDrv;

      // Construir fila en el orden de cabeceras existentes
      var newRow = hdrs.map(function(h) {
        return row[h] !== undefined ? row[h] : '';
      });
      // Si URL_PDF_DRIVE no existe en cabeceras, añadir columna
      if (hdrs.indexOf('URL_PDF_DRIVE') === -1 && urlDrv) {
        newRow.push(urlDrv);
      }
      sheet.appendRow(newRow);
      count++;
    });

    return { ok: true, count: count };
  } catch(e) {
    return { ok: false, error: e.toString() };
  }
}

// ===========================================================================
//  ENVÍO DE DOCUMENTACIÓN A ASESORÍA
// ===========================================================================
function enviarDocsAsesoria_() {
  var book   = ss_();
  var firmas = sheetToJSON(book, 'FIRMAS_PROTOCOLO');
  var conPdf = firmas.filter(function(f) { return f.PDF_ID; });

  if (!conPdf.length) return { ok: false, error: 'No hay documentos firmados con PDF generado' };

  var attachments = [];
  var resumen     = [];
  conPdf.forEach(function(f) {
    try {
      var blob = DriveApp.getFileById(f.PDF_ID).getBlob();
      blob.setName((f.TITULO_DOC || f.DOC_ID) + ' — ' + (f.NOMBRE || '') + '.pdf');
      attachments.push(blob);
      resumen.push('• ' + (f.NOMBRE || '—') + ' · ' + (f.TITULO_DOC || f.DOC_ID) + ' · ' + (f.FECHA || ''));
    } catch(e) {
      Logger.log('Error adjuntando ' + f.PDF_ID + ': ' + e.message);
    }
  });

  if (!attachments.length) return { ok: false, error: 'No se pudieron obtener los archivos PDF' };

  var asunto  = 'SM Academia — Protocolos firmados por el equipo';
  var cuerpo  = 'Se adjuntan los protocolos firmados digitalmente por el equipo de SM Academia Sports & Football.\n\n'
    + resumen.join('\n') + '\n\n'
    + 'Total adjuntos: ' + attachments.length + '\n'
    + 'Enviado: ' + new Date().toLocaleDateString('es-ES');

  MailApp.sendEmail({
    to:          EMAIL_EMPRESA.asesoria,
    replyTo:     EMAIL_EMPRESA.direccion,
    name:        'SM Academia',
    subject:     asunto,
    body:        cuerpo,
    attachments: attachments,
  });

  return { ok: true, enviados: attachments.length };
}

// ===========================================================================
//  PDF personal: obtener bytes y guardar firmado
// ===========================================================================
function obtenerPdfB64_(fileId) {
  try {
    var file  = DriveApp.getFileById(fileId);
    var bytes = file.getBlob().getBytes();
    return { ok: true, b64: Utilities.base64Encode(bytes), nombre: file.getName() };
  } catch(e) {
    return { ok: false, error: e.toString() };
  }
}

function guardarPdfFirmadoPersonal_(body) {
  try {
    var bytes  = Utilities.base64Decode(body.pdfB64);
    var nombre = String(body.nombre || 'documento_firmado').replace(/\.pdf$/i,'') + '_FIRMADO.pdf';
    var blob   = Utilities.newBlob(bytes, 'application/pdf', nombre);
    var folder = _carpetaFirmados_();
    var file   = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return { ok: true, url: file.getUrl(), fileId: file.getId(), nombre: nombre };
  } catch(e) {
    return { ok: false, error: e.toString() };
  }
}

function subirCNDS_(body) {
  try {
    var pdfB64    = body.pdfB64;
    var nombre    = String(body.nombre || 'CNDS.pdf');
    var monitorId = String(body.monitorId || '');

    if (!pdfB64)    return { ok: false, error: 'Falta el contenido del PDF' };
    if (!monitorId) return { ok: false, error: 'Falta el ID del monitor' };

    var book     = ss_();
    var personal = sheetToJSON(book, 'PERSONAL');
    var monitor  = personal.find(function(p) { return String(p.ID_MONITOR) === monitorId; });
    if (!monitor) return { ok: false, error: 'Monitor no encontrado (ID: ' + monitorId + ')' };

    var nombreCompleto = [monitor.NOMBRE, monitor.APELLIDOS].filter(Boolean).join(' ').trim();

    function getOrCreate(parent, folderName) {
      var it = parent.getFoldersByName(folderName);
      return it.hasNext() ? it.next() : parent.createFolder(folderName);
    }
    var fSMAcademia = getOrCreate(DriveApp.getRootFolder(), 'SM Academia');
    var fExtra      = getOrCreate(fSMAcademia,  'SM Extraescolares');
    var fContratos  = getOrCreate(fExtra,        'Contratos');
    var fMonitor    = getOrCreate(fContratos,    nombreCompleto);

    var bytes  = Utilities.base64Decode(pdfB64);
    var blob   = Utilities.newBlob(bytes, 'application/pdf', nombre);
    var file   = fMonitor.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    var url    = file.getUrl();
    var titulo = nombre.replace(/\.pdf$/i, '').replace(/_/g, ' ');

    var oldDocs = String(monitor.DOCS_DRIVE || '').trim();
    var newDocs = oldDocs ? oldDocs + ';;' + titulo + '|' + url : titulo + '|' + url;

    updateRow('PERSONAL', monitorId, { DOCS_DRIVE: newDocs });

    return { ok: true, url: url, fileId: file.getId(), nombre: nombre, titulo: titulo };
  } catch(e) {
    return { ok: false, error: e.toString() };
  }
}

// ===========================================================================
//  NÓMINAS — PROXY DE DESCARGA Y SUBIDA DE PDF
// ===========================================================================
function getNominaPdf_(fileId) {
  if (!fileId) return { ok: false, error: 'fileId requerido' };
  try {
    var file = DriveApp.getFileById(fileId);
    var blob = file.getBlob();
    return {
      ok:       true,
      nombre:   file.getName(),
      mimeType: blob.getContentType() || 'application/pdf',
      data:     Utilities.base64Encode(blob.getBytes())
    };
  } catch(e) {
    return { ok: false, error: 'Archivo no accesible: ' + e.message };
  }
}

function subirNominaPdf_(body) {
  var base64     = body.base64;
  var fileName   = body.fileName   || 'nomina.pdf';
  if (!base64) return { ok: false, error: 'base64 requerido' };
  try {
    var bytes  = Utilities.base64Decode(base64);
    var blob   = Utilities.newBlob(bytes, MimeType.PDF, fileName);
    var folder = _carpetaNominas_();
    var file   = folder.createFile(blob);
    return { ok: true, fileId: file.getId(), url: file.getUrl() };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

function _carpetaNominas_() {
  var name = 'SM Academia — Nóminas';
  var it   = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}
