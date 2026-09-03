var CONFIG = {
  SPREADSHEET_ID: '1epQQsmHk6jscDXDu8JRAFPRMyT_RCKmp6aguI5oMXJI', // CONFIGURAÇÃO NECESSÁRIA: confirmar
  SHEET_REGISTROS: 'REGISTROS',
  TIMEZONE: 'America/Recife'
};

// Ordem oficial das colunas na aba REGISTROS. Esta lista é a fonte única da
// verdade do layout da planilha — não reordene sem atualizar setupKronosSheets().
var COLUMNS = [
  'recordId', 'date', 'shift', 'psName', 'area', 'site',
  'metaOutbound', 'realOutbound',
  'metaFastStart', 'realFastStart',
  'metaStrongFinish', 'realStrongFinish',
  'volPlan', 'volReal',
  'weightsJson',
  'realIdle', 'realCot', 'idleCompliance',
  's5Json',
  'notes',
  'score', 'classification',
  'hourlyDataJson',
  'validationStatus', 'validatedBy', 'validationDate', 'validationComment',
  'createdAt', 'updatedAt'
];

var VALID_SHIFTS = ['T1', 'T2', 'T3'];
var REQUIRED_FIELDS = ['date', 'shift', 'psName', 'area'];

// ----------------------------------------------------------------------------
// SETUP (rodar manualmente uma vez)
// ----------------------------------------------------------------------------
function setupKronosSheets() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.SHEET_REGISTROS);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_REGISTROS);
    sheet.getRange(1, 1, 1, COLUMNS.length).setValues([COLUMNS]);
    sheet.setFrozenRows(1);
    Logger.log('Aba "' + CONFIG.SHEET_REGISTROS + '" criada com sucesso.');
    return 'Aba criada com sucesso.';
  }

  var lastCol = sheet.getLastColumn();
  var currentHeaders = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  var headersMatch = COLUMNS.length === currentHeaders.length &&
    COLUMNS.every(function (h, i) { return h === currentHeaders[i]; });

  if (headersMatch) {
    Logger.log('Aba "' + CONFIG.SHEET_REGISTROS + '" já existe e está correta.');
    return 'Aba já configurada corretamente.';
  }

  // Não sobrescreve dados existentes automaticamente — apenas avisa.
  var msg = 'ATENÇÃO: a aba "' + CONFIG.SHEET_REGISTROS + '" já existe com um layout ' +
    'diferente do esperado.\nEsperado: ' + JSON.stringify(COLUMNS) +
    '\nEncontrado: ' + JSON.stringify(currentHeaders) +
    '\nAjuste manualmente ou renomeie a aba atual e rode setupKronosSheets() novamente.';
  Logger.log(msg);
  return msg;
}

// ----------------------------------------------------------------------------
// ENTRADAS HTTP
// ----------------------------------------------------------------------------
function doGet(e) {
  try {
    var params = (e && e.parameter) || {};
    if (!params.action) {
      return jsonOutput_(successResponse_({ status: 'online' }, 'KRONOS API online.'));
    }
    var result = routeAction_(params.action, params, false);
    return jsonOutput_(result);
  } catch (err) {
    logAction_('doGet', false, err);
    return jsonOutput_(errorResponse_('Erro ao processar solicitação GET.', err.message));
  }
}

function doPost(e) {
  try {
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    var action = body.action;
    var data = body.data || {};
    if (!action) {
      return jsonOutput_(errorResponse_('Parâmetro "action" não informado.'));
    }
    var result = routeAction_(action, data, true);
    return jsonOutput_(result);
  } catch (err) {
    logAction_('doPost', false, err);
    return jsonOutput_(errorResponse_('Erro ao processar solicitação POST.', err.message));
  }
}

function routeAction_(action, params, isPost) {
  var writeActions = ['createRecord', 'updateRecord', 'deleteRecord', 'validateScore', 'rejectScore'];
  if (writeActions.indexOf(action) !== -1 && !isPost) {
    throw new Error('A ação "' + action + '" requer método POST.');
  }

  switch (action) {
    case 'getData':
      return successResponse_(getData_(), 'Dados carregados com sucesso.');
    case 'createRecord':
      return successResponse_(createRecord_(params), 'Registro criado com sucesso.');
    case 'updateRecord':
      return successResponse_(updateRecord_(params), 'Registro atualizado com sucesso.');
    case 'deleteRecord':
      return successResponse_(deleteRecord_(params), 'Registro removido com sucesso.');
    case 'validateScore':
      return successResponse_(setValidation_(params, 'VALIDATED'), 'Turno validado com sucesso.');
    case 'rejectScore':
      return successResponse_(setValidation_(params, 'REJECTED'), 'Turno reprovado com sucesso.');
    default:
      throw new Error('Ação desconhecida: ' + action);
  }
}

// ----------------------------------------------------------------------------
// LEITURA
// ----------------------------------------------------------------------------
function getData_() {
  var sheet = getSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    logAction_('getData', true, { count: 0 });
    return { records: [] };
  }

  var range = sheet.getRange(2, 1, lastRow - 1, COLUMNS.length);
  var values = range.getValues();

  var records = values
    .filter(function (row) { return row[0] !== '' && row[0] !== null; }) // ignora linhas vazias
    .map(rowToRecord_)
    .sort(function (a, b) {
      if (a.date === b.date) return (a.createdAt || '').localeCompare(b.createdAt || '');
      return a.date < b.date ? -1 : 1;
    });

  logAction_('getData', true, { count: records.length });
  return { records: records };
}

// ----------------------------------------------------------------------------
// ESCRITA
// ----------------------------------------------------------------------------
function createRecord_(data) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    validateRecordInput_(data);

    var sheet = getSheet_();
    var id = data.id || Utilities.getUuid();

    if (findRowIndexById_(sheet, id) !== -1) {
      throw new Error('Já existe um registro com este ID. Use updateRecord para alterá-lo.');
    }

    var now = getCurrentDateTime_();
    var toSave = shallowClone_(data);
    toSave.id = id;
    toSave.createdAt = now;
    toSave.updatedAt = now;
    if (!toSave.validationStatus) toSave.validationStatus = 'PENDING';

    var row = recordToRow_(toSave);
    sheet.appendRow(row);

    logAction_('createRecord', true, { id: id });
    return rowToRecord_(row);
  } catch (err) {
    logAction_('createRecord', false, err);
    throw err;
  } finally {
    lock.releaseLock();
  }
}

function updateRecord_(data) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    if (!data.id) throw new Error('ID do registro é obrigatório para atualização.');
    validateRecordInput_(data);

    var sheet = getSheet_();
    var rowIndex = findRowIndexById_(sheet, data.id);
    if (rowIndex === -1) throw new Error('Registro não encontrado: ' + data.id);

    var existingRow = sheet.getRange(rowIndex, 1, 1, COLUMNS.length).getValues()[0];
    var existingRecord = rowToRecord_(existingRow);

    // Mescla: campos não enviados (ex.: validação, quando o front-end omite
    // undefined) preservam o valor já existente na planilha.
    var merged = shallowClone_(existingRecord);
    for (var key in data) {
      if (data.hasOwnProperty(key) && data[key] !== undefined) {
        merged[key] = data[key];
      }
    }
    merged.id = data.id;
    merged.createdAt = existingRecord.createdAt;
    merged.updatedAt = getCurrentDateTime_();

    var row = recordToRow_(merged);
    sheet.getRange(rowIndex, 1, 1, COLUMNS.length).setValues([row]);

    logAction_('updateRecord', true, { id: data.id });
    return rowToRecord_(row);
  } catch (err) {
    logAction_('updateRecord', false, err);
    throw err;
  } finally {
    lock.releaseLock();
  }
}

function deleteRecord_(data) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    if (!data.id) throw new Error('ID do registro é obrigatório para exclusão.');
    var sheet = getSheet_();
    var rowIndex = findRowIndexById_(sheet, data.id);
    if (rowIndex === -1) throw new Error('Registro não encontrado: ' + data.id);

    sheet.deleteRow(rowIndex);
    logAction_('deleteRecord', true, { id: data.id });
    return { id: data.id };
  } catch (err) {
    logAction_('deleteRecord', false, err);
    throw err;
  } finally {
    lock.releaseLock();
  }
}

function setValidation_(data, status) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    if (!data.id) throw new Error('ID do registro é obrigatório.');
    if (!data.validatedBy) throw new Error('Nome do supervisor é obrigatório.');
    if (status === 'REJECTED' && !data.comment) {
      throw new Error('Comentário é obrigatório em caso de reprovação.');
    }

    var sheet = getSheet_();
    var rowIndex = findRowIndexById_(sheet, data.id);
    if (rowIndex === -1) throw new Error('Registro não encontrado: ' + data.id);

    var existingRow = sheet.getRange(rowIndex, 1, 1, COLUMNS.length).getValues()[0];
    var record = rowToRecord_(existingRow);

    record.validationStatus = status;
    record.validatedBy = data.validatedBy;
    record.validationDate = getCurrentDateTime_();
    record.validationComment = data.comment || '';
    record.updatedAt = getCurrentDateTime_();

    var row = recordToRow_(record);
    sheet.getRange(rowIndex, 1, 1, COLUMNS.length).setValues([row]);

    logAction_('setValidation', true, { id: data.id, status: status });
    return rowToRecord_(row);
  } catch (err) {
    logAction_('setValidation', false, err);
    throw err;
  } finally {
    lock.releaseLock();
  }
}

// ----------------------------------------------------------------------------
// VALIDAÇÃO (backend nunca confia só no front-end)
// ----------------------------------------------------------------------------
function validateRecordInput_(data) {
  REQUIRED_FIELDS.forEach(function (field) {
    if (data[field] === undefined || data[field] === null || data[field] === '') {
      throw new Error('Campo obrigatório ausente: ' + field);
    }
  });

  if (VALID_SHIFTS.indexOf(data.shift) === -1) {
    throw new Error('Turno inválido: ' + data.shift + '. Valores aceitos: ' + VALID_SHIFTS.join(', '));
  }

  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(data.date))) {
    throw new Error('Formato de data inválido (esperado AAAA-MM-DD): ' + data.date);
  }

  var numericFields = ['metaOutbound', 'realOutbound', 'metaFastStart', 'realFastStart',
    'metaStrongFinish', 'realStrongFinish', 'volPlan', 'volReal', 'score'];
  numericFields.forEach(function (field) {
    if (data[field] !== undefined && data[field] !== null && data[field] !== '' && isNaN(Number(data[field]))) {
      throw new Error('Campo numérico inválido: ' + field + ' = ' + data[field]);
    }
  });
}

// ----------------------------------------------------------------------------
// CONVERSÃO REGISTRO <-> LINHA DA PLANILHA
// ----------------------------------------------------------------------------
function recordToRow_(r) {
  var weights = r.weights || {};
  var s5 = r.s5 || {};
  var hourly = {
    thpMeta: r.hourlyTHPMeta || [],
    thpReal: r.hourlyTHPReal || [],
    idleMeta: r.hourlyIdleMeta || [],
    idleReal: r.hourlyIdleReal || [],
    cotMeta: r.hourlyCOTMeta || [],
    cotReal: r.hourlyCOTReal || []
  };

  var map = {
    recordId: r.id || '',
    date: r.date || '',
    shift: r.shift || '',
    psName: r.psName || '',
    area: r.area || '',
    site: r.site || 'BRFPE1',
    metaOutbound: numOrZero_(r.metaOutbound),
    realOutbound: numOrZero_(r.realOutbound),
    metaFastStart: numOrZero_(r.metaFastStart),
    realFastStart: numOrZero_(r.realFastStart),
    metaStrongFinish: numOrZero_(r.metaStrongFinish),
    realStrongFinish: numOrZero_(r.realStrongFinish),
    volPlan: numOrZero_(r.volPlan),
    volReal: numOrZero_(r.volReal),
    weightsJson: JSON.stringify(weights),
    realIdle: numOrZero_(r.realIdle),
    realCot: numOrZero_(r.realCot),
    idleCompliance: numOrZero_(r.idleCompliance),
    s5Json: JSON.stringify(s5),
    notes: r.notes || '',
    score: numOrZero_(r.score),
    classification: r.classification || '',
    hourlyDataJson: JSON.stringify(hourly),
    validationStatus: r.validationStatus || 'PENDING',
    validatedBy: r.validatedBy || '',
    validationDate: r.validationDate || '',
    validationComment: r.validationComment || '',
    createdAt: r.createdAt || getCurrentDateTime_(),
    updatedAt: r.updatedAt || getCurrentDateTime_()
  };

  return COLUMNS.map(function (key) { return map[key]; });
}

function rowToRecord_(row) {
  var raw = {};
  COLUMNS.forEach(function (key, i) { raw[key] = row[i]; });

  var weights = safeParseJson_(raw.weightsJson, {});
  var s5 = safeParseJson_(raw.s5Json, {});
  var hourly = safeParseJson_(raw.hourlyDataJson, {});

  return {
    id: raw.recordId,
    date: raw.date,
    shift: raw.shift,
    psName: raw.psName,
    area: raw.area,
    site: raw.site,
    metaOutbound: Number(raw.metaOutbound) || 0,
    realOutbound: Number(raw.realOutbound) || 0,
    metaFastStart: Number(raw.metaFastStart) || 0,
    realFastStart: Number(raw.realFastStart) || 0,
    metaStrongFinish: Number(raw.metaStrongFinish) || 0,
    realStrongFinish: Number(raw.realStrongFinish) || 0,
    volPlan: Number(raw.volPlan) || 0,
    volReal: Number(raw.volReal) || 0,
    weights: weights,
    realIdle: Number(raw.realIdle) || 0,
    realCot: Number(raw.realCot) || 0,
    idleCompliance: Number(raw.idleCompliance) || 0,
    s5: s5,
    notes: raw.notes || '',
    score: Number(raw.score) || 0,
    classification: raw.classification || '',
    hourlyTHPMeta: hourly.thpMeta || [],
    hourlyTHPReal: hourly.thpReal || [],
    hourlyIdleMeta: hourly.idleMeta || [],
    hourlyIdleReal: hourly.idleReal || [],
    hourlyCOTMeta: hourly.cotMeta || [],
    hourlyCOTReal: hourly.cotReal || [],
    validationStatus: raw.validationStatus || 'PENDING',
    validatedBy: raw.validatedBy || '',
    validationDate: raw.validationDate || '',
    validationComment: raw.validationComment || '',
    createdAt: raw.createdAt,
    updatedAt: raw.updatedAt
  };
}

// ----------------------------------------------------------------------------
// AUXILIARES
// ----------------------------------------------------------------------------
function getSheet_() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.SHEET_REGISTROS);
  if (!sheet) {
    throw new Error('CONFIGURAÇÃO NECESSÁRIA: aba "' + CONFIG.SHEET_REGISTROS +
      '" não encontrada. Rode a função setupKronosSheets() uma vez antes de usar a API.');
  }
  return sheet;
}

function findRowIndexById_(sheet, id) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (ids[i][0] === id) return i + 2; // +2: cabeçalho + índice base 1
  }
  return -1;
}

function getCurrentDateTime_() {
  return Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");
}

function numOrZero_(v) {
  var n = Number(v);
  return isNaN(n) ? 0 : n;
}

function safeParseJson_(str, fallback) {
  try {
    return str ? JSON.parse(str) : fallback;
  } catch (e) {
    return fallback;
  }
}

function shallowClone_(obj) {
  var copy = {};
  for (var k in obj) { if (obj.hasOwnProperty(k)) copy[k] = obj[k]; }
  return copy;
}

function jsonOutput_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function successResponse_(data, message) {
  return { success: true, data: data, message: message || '', timestamp: getCurrentDateTime_() };
}

function errorResponse_(message, err) {
  return { success: false, data: null, message: message || 'Erro desconhecido.', error: err || message };
}

function logAction_(action, success, extra) {
  try {
    Logger.log(JSON.stringify({
      action: action,
      success: success,
      timestamp: getCurrentDateTime_(),
      detail: (extra && extra.message) ? extra.message : extra
    }));
  } catch (e) { /* nunca deixa o log quebrar a execução */ }
}

// ----------------------------------------------------------------------------
// Mantido do arquivo original — útil para conferir a estrutura real da
// planilha antes/depois de rodar setupKronosSheets().
// ----------------------------------------------------------------------------
function inspecionarPlanilha() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheets = ss.getSheets();
  var resultado = '';
  sheets.forEach(function (sheet) {
    var nome = sheet.getName();
    var lastCol = sheet.getLastColumn();
    var lastRow = sheet.getLastRow();
    var headers = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
    resultado += '\n=== ABA: "' + nome + '" ===\n';
    resultado += 'Linhas com dados: ' + lastRow + ' | Colunas: ' + lastCol + '\n';
    resultado += 'Cabeçalhos: ' + JSON.stringify(headers) + '\n';
  });
  Logger.log(resultado);
  return resultado;
}
