// =============================================================
// Mia 咪啊 Health Tracker v3.1 — Google Apps Script Backend
// Backward compatible: reads Chinese AND English sheet names
// =============================================================

function doGet(e) {
  var action = (e.parameter.action || '').toLowerCase();
  if (action === 'ping') return jsonOut({status:'ok', time: new Date().toISOString()});
  if (action === 'load') return jsonOut(loadAllData());
  return jsonOut({status:'ready'});
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var action = data.action || 'append';
    if (action === 'append') return jsonOut(appendData(data));
    if (action === 'delete') return jsonOut(deleteRecord(data));
    if (action === 'settings') return jsonOut(saveSettings(data));
    if (action === 'fullSync') return jsonOut(fullSync(data));
    return jsonOut({status:'error', message:'Unknown action'});
  } catch(err) {
    return jsonOut({status:'error', message: err.toString()});
  }
}

// ---- Sheet name mapping (English = primary, Chinese = legacy fallback) ----
var SHEET_NAMES = {
  diet:     {en: 'Diet',     zh: '飲食紀錄'},
  health:   {en: 'Health',   zh: '健康紀錄'},
  weight:   {en: 'Weight',   zh: '體重紀錄'},
  shedding: {en: 'Shedding', zh: '掉毛紀錄'}
};

var HEADERS = {
  diet:     ['id','date','time','type','brand','flavor','amount','note'],
  health:   ['id','date','time','type','value','note'],
  weight:   ['date','weight'],
  shedding: ['id','date','level','note']
};

// ---- Find sheet: try English name first, then Chinese ----
function findSheet(ss, type) {
  var names = SHEET_NAMES[type];
  if (!names) return null;
  var sheet = ss.getSheetByName(names.en);
  if (sheet) return {sheet: sheet, source: 'en'};
  sheet = ss.getSheetByName(names.zh);
  if (sheet) return {sheet: sheet, source: 'zh'};
  return null;
}

// ---- Get or create English-named sheet ----
function getOrCreateEn(ss, type) {
  var names = SHEET_NAMES[type];
  var headers = HEADERS[type];
  var sheet = ss.getSheetByName(names.en);
  if (!sheet) {
    sheet = ss.insertSheet(names.en);
    sheet.appendRow(headers);
    sheet.getRange(1,1,1,headers.length).setFontWeight('bold');
  }
  return sheet;
}

// ---- APPEND a single record ----
function appendData(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var type = data.type;
  var r = data.record;
  var sheet = getOrCreateEn(ss, type);
  
  if (type === 'diet')     sheet.appendRow([r.id, r.date, r.time||'', r.type, r.brand||'', r.flavor||'', r.amount||'', r.note||'']);
  else if (type === 'health')  sheet.appendRow([r.id, r.date, r.time||'', r.type, r.value||'', r.note||'']);
  else if (type === 'weight')  sheet.appendRow([r.date, r.weight]);
  else if (type === 'shedding') sheet.appendRow([r.id, r.date, r.level, r.note||'']);
  
  return {status:'ok', type:type};
}

// ---- DELETE a record by id ----
function deleteRecord(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var found = findSheet(ss, data.type);
  if (!found) return {status:'ok', message:'Sheet not found'};
  var sheet = found.sheet;
  var rows = sheet.getDataRange().getValues();
  for (var i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][0]) === String(data.id || data.date)) {
      sheet.deleteRow(i + 1);
      return {status:'ok', deleted: data.id || data.date};
    }
  }
  return {status:'ok', message:'Not found'};
}

// ---- SAVE settings (protected: won't overwrite non-empty with empty) ----
function saveSettings(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settings = data.settings || {};
  
  // Read existing settings first
  var existing = {};
  var sheet = ss.getSheetByName('Settings');
  if (sheet && sheet.getLastRow() > 1) {
    var sData = sheet.getRange(2, 1, sheet.getLastRow()-1, 2).getValues();
    for (var i = 0; i < sData.length; i++) {
      var k = String(sData[i][0]), v = String(sData[i][1]);
      try { existing[k] = JSON.parse(v); } catch(e) { existing[k] = v; }
    }
  }
  
  // For each key, if new value is empty but existing has content, keep existing
  var finalSettings = {};
  var keys = Object.keys(settings);
  for (var i = 0; i < keys.length; i++) {
    var k = keys[i], newVal = settings[k], oldVal = existing[k];
    var isEmpty = isEmptyVal(newVal);
    var oldHasContent = !isEmptyVal(oldVal);
    if (isEmpty && oldHasContent) {
      finalSettings[k] = oldVal; // Keep existing
    } else {
      finalSettings[k] = newVal;
    }
  }
  // Also preserve any keys that existed but weren't sent
  Object.keys(existing).forEach(function(k){ if (!(k in finalSettings)) finalSettings[k] = existing[k]; });
  
  var s = getOrCreateSettings(ss);
  if (s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow()-1, 3).clearContent();
  
  var fKeys = Object.keys(finalSettings);
  for (var i = 0; i < fKeys.length; i++) {
    var val = finalSettings[fKeys[i]];
    s.appendRow([fKeys[i], typeof val === 'object' ? JSON.stringify(val) : String(val), new Date().toISOString()]);
  }
  return {status:'ok'};
}

function isEmptyVal(v) {
  if (v === null || v === undefined || v === '') return true;
  if (Array.isArray(v) && v.length === 0) return true;
  if (typeof v === 'object') {
    if (v.name === '' && v.cycle === 0 && v.last === '') return true; // empty med
    if (Object.keys(v).length === 0) return true;
  }
  return false;
}

function getOrCreateSettings(ss) {
  var sheet = ss.getSheetByName('Settings');
  if (!sheet) {
    sheet = ss.insertSheet('Settings');
    sheet.appendRow(['key','value','updated']);
    sheet.getRange(1,1,1,3).setFontWeight('bold');
  }
  return sheet;
}

// ---- LOAD all data (reads from Chinese OR English sheets) ----
function loadAllData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var result = {status:'ok'};
  
  // Diet - try English first, then Chinese
  result.dietLogs = loadType(ss, 'diet', function(row, src) {
    if (src === 'zh') {
      // Chinese sheet format: ID, 日期, 時間, 類型, 品牌, 口味, 份量, 備註
      return {id:s(row[0]), date:s(row[1]), time:s(row[2]), type:s(row[3]), brand:s(row[4]), flavor:s(row[5]), amount:s(row[6]), note:s(row[7])};
    }
    return {id:s(row[0]), date:s(row[1]), time:s(row[2]), type:s(row[3]), brand:s(row[4]), flavor:s(row[5]), amount:s(row[6]), note:s(row[7])};
  });
  
  // Health
  result.healthLogs = loadType(ss, 'health', function(row, src) {
    if (src === 'zh') {
      return {id:s(row[0]), date:s(row[1]), time:s(row[2]), type:s(row[3]), value:s(row[4]), note:s(row[5])};
    }
    return {id:s(row[0]), date:s(row[1]), time:s(row[2]), type:s(row[3]), value:s(row[4]), note:s(row[5])};
  });
  
  // Weight
  result.weights = loadType(ss, 'weight', function(row, src) {
    // Both formats: first col = date, second col = weight
    var dateStr = s(row[0]);
    // Clean date: if it's a Date object converted to string, extract just the date part
    if (dateStr.length > 10) dateStr = dateStr.substring(0, 10);
    return {date: dateStr, weight: Number(row[1])};
  });
  
  // Shedding
  result.sheddingLogs = loadType(ss, 'shedding', function(row, src) {
    if (src === 'zh') {
      return {id:s(row[0]), date:s(row[1]), level:Number(row[2]), note:s(row[3])};
    }
    return {id:s(row[0]), date:s(row[1]), level:Number(row[2]), note:s(row[3])};
  });
  
  // Settings
  result.settings = {};
  var settingsSheet = ss.getSheetByName('Settings');
  if (settingsSheet && settingsSheet.getLastRow() > 1) {
    var sData = settingsSheet.getRange(2, 1, settingsSheet.getLastRow()-1, 2).getValues();
    for (var i = 0; i < sData.length; i++) {
      var key = String(sData[i][0]), val = String(sData[i][1]);
      try { result.settings[key] = JSON.parse(val); } catch(e) { result.settings[key] = val; }
    }
  }
  
  return result;
}

// Load from best available sheet (English first, then Chinese)
function loadType(ss, type, mapper) {
  var found = findSheet(ss, type);
  if (!found || found.sheet.getLastRow() <= 1) return [];
  var data = found.sheet.getRange(2, 1, found.sheet.getLastRow()-1, found.sheet.getLastColumn()).getValues();
  return data.map(function(row) { return mapper(row, found.source); }).filter(function(item) {
    // Filter out rows with empty first column
    var firstVal = item.id || item.date;
    return firstVal && String(firstVal).trim() !== '' && String(firstVal) !== 'undefined';
  });
}

function s(v) { return v === null || v === undefined ? '' : String(v); }

// ---- FULL SYNC (initial migration) ----
function fullSync(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  if (data.dietLogs && data.dietLogs.length) {
    var sheet = getOrCreateEn(ss, 'diet');
    var existing = getExistingIds(sheet);
    data.dietLogs.forEach(function(d) {
      if (d.id && !existing[d.id]) sheet.appendRow([d.id, d.date, d.time||'', d.type, d.brand||'', d.flavor||'', d.amount||'', d.note||'']);
    });
  }
  
  if (data.healthLogs && data.healthLogs.length) {
    var sheet = getOrCreateEn(ss, 'health');
    var existing = getExistingIds(sheet);
    data.healthLogs.forEach(function(h) {
      if (h.id && !existing[h.id]) sheet.appendRow([h.id, h.date, h.time||'', h.type, h.value||'', h.note||'']);
    });
  }
  
  if (data.weights && data.weights.length) {
    var sheet = getOrCreateEn(ss, 'weight');
    var existingDates = {};
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2,1,sheet.getLastRow()-1,1).getValues().forEach(function(r){existingDates[String(r[0]).substring(0,10)]=true});
    }
    data.weights.forEach(function(w) {
      if (w.date && !existingDates[w.date]) sheet.appendRow([w.date, w.weight]);
    });
  }
  
  if (data.sheddingLogs && data.sheddingLogs.length) {
    var sheet = getOrCreateEn(ss, 'shedding');
    var existing = getExistingIds(sheet);
    data.sheddingLogs.forEach(function(sh) {
      if (sh.id && !existing[sh.id]) sheet.appendRow([sh.id, sh.date, sh.level, sh.note||'']);
    });
  }
  
  if (data.settings) saveSettings({settings: data.settings});
  
  return {status:'ok', message:'Full sync complete'};
}

function getExistingIds(sheet) {
  var ids = {};
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2,1,sheet.getLastRow()-1,1).getValues().forEach(function(r){ids[String(r[0])]=true});
  }
  return ids;
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}