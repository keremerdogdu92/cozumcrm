/* v3.8.0
 * CHANGELOG:
 * - Documented compatibility with SGK-aware trial quick entry offers.
 * - Keeps header-cache, slugification, and row IO utilities aligned with trial updates.
 * - Moved Quick Entry shared helpers (errors, date parsing, prefixed ids) into core utils.
 */

const HEADER_CACHE = Object.create(null);

function sh(name){ return SpreadsheetApp.getActive().getSheetByName(name); }
function ensure(name){ return sh(name) || SpreadsheetApp.getActive().insertSheet(name); }
function now(){ return new Date(); }
function fmt(d){ return Utilities.formatDate(new Date(d), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'); }
function ymdhm(d){ const z=n=>('0'+n).slice(-2); return d.getFullYear()+z(d.getMonth()+1)+z(d.getDate())+z(d.getHours())+z(d.getMinutes()); }
function rand2(){ const c='ABCDEFGHJKLMNPQRSTUVWXYZ23456789'; return c[Math.floor(Math.random()*c.length)]+c[Math.floor(Math.random()*c.length)]; }
function colLetter_(col){ let s=''; while(col>0){ let m=(col-1)%26; s=String.fromCharCode(65+m)+s; col=(col-m-1)/26|0; } return s; }
function toast(msg){ SpreadsheetApp.getActive().toast(String(msg)); }

function invalidateHeaderCache_(sheetName){ if(sheetName){ delete HEADER_CACHE[String(sheetName)]; } else { Object.keys(HEADER_CACHE).forEach(k=>delete HEADER_CACHE[k]); } }

function getHeaderMap(sheetRef){
  const sheet = typeof sheetRef === 'string' ? sh(sheetRef) : sheetRef;
  if(!sheet) throw new Error('Sheet not found: '+sheetRef);
  const name = sheet.getName();
  const lastCol = Math.max(1, sheet.getLastColumn());
  const cacheKey = name+'#'+lastCol;
  const cached = HEADER_CACHE[name];
  if(cached && cached.signature === cacheKey){ return cached.map; }

  const headers = sheet.getRange(1,1,1,lastCol).getValues()[0];
  const order = new Array(headers.length);
  const map = Object.create(null);
  for(let i=0;i<headers.length;i++){
    const raw = headers[i];
    const key = raw && String(raw).trim();
    order[i] = key || '';
    if(key){ map[key] = i+1; }
  }
  map.__order = order;
  map.__sheetName = name;
  HEADER_CACHE[name] = {signature: cacheKey, map};
  return map;
}

function headerIndex_(sheet){ return getHeaderMap(sheet); }

function getHeaderOrder_(sheetName){ return getHeaderMap(sheetName).__order; }

function toSlugTr(name){
  if(!name) return '';
  const map = {
    'Ç':'c','ç':'c','Ğ':'g','ğ':'g','İ':'i','I':'i','ı':'i','Ö':'o','ö':'o','Ş':'s','ş':'s','Ü':'u','ü':'u','Â':'a','â':'a','Ê':'e','ê':'e','Î':'i','î':'i','Ô':'o','ô':'o','Û':'u','û':'u'
  };
  const len = name.length;
  let out = '';
  let prevUnderscore = true;
  for(let i=0;i<len;i++){
    const ch = name.charAt(i);
    const lower = map.hasOwnProperty(ch) ? map[ch] : ch.toLowerCase();
    if(/[a-z0-9]/.test(lower)){
      out += lower;
      prevUnderscore = false;
      continue;
    }
    if(/[\s\-\.+]/.test(ch)){
      if(!prevUnderscore) out += '_';
      prevUnderscore = true;
    }
  }
  out = out.replace(/_+/g,'_').replace(/^_+|_+$/g,'');
  return out;
}

function qe_error_(code, message){
  const err = new Error(message);
  err.code = code;
  return err;
}

function qe_parseDate_(value, fieldName){
  if(value instanceof Date){ return value; }
  const str = String(value || '').trim();
  if(!str){ throw qe_error_('missing_'+String(fieldName).replace(/\./g,'_'), fieldName+' required'); }
  const parsed = new Date(str);
  if(isNaN(parsed.getTime())){ throw qe_error_('invalid_'+String(fieldName).replace(/\./g,'_'), fieldName+' invalid date'); }
  return parsed;
}

function qe_parseDateOptional_(value){
  if(!value){ return null; }
  if(value instanceof Date){ return value; }
  const str = String(value || '').trim();
  if(!str) return null;
  const parsed = new Date(str);
  if(isNaN(parsed.getTime())){ throw qe_error_('invalid_date','Invalid date value'); }
  return parsed;
}

function ensureSheetHeaders(sheetName, expectedKeys, labels){
  const sheet = ensure(sheetName);
  if(sheet.getMaxRows() < 2){ sheet.insertRows(1, 2 - sheet.getMaxRows()); }
  const lastCol = sheet.getLastColumn();
  const haveCols = Math.max(lastCol, 0);
  const haveKeys = haveCols ? sheet.getRange(1,1,1,haveCols).getValues()[0] : [];
  const haveLabels = haveCols ? sheet.getRange(2,1,1,haveCols).getValues()[0] : [];
  const seen = Object.create(null);
  for(let i=0;i<haveKeys.length;i++){ const key=haveKeys[i] && String(haveKeys[i]).trim(); if(key){ seen[key]=true; } }
  const addKeys = [];
  const addLabels = [];
  if(expectedKeys){
    for(let i=0;i<expectedKeys.length;i++){
      const key = expectedKeys[i];
      if(!key || seen[key]) continue;
      addKeys.push(key);
      addLabels.push(labels && labels[i] ? labels[i] : key);
      seen[key] = true;
    }
  }
  if(addKeys.length){
    const startCol = haveKeys.length ? haveKeys.length+1 : 1;
    sheet.getRange(1,startCol,1,addKeys.length).setValues([addKeys]);
    sheet.getRange(2,startCol,1,addKeys.length).setValues([addLabels]);
  } else {
    // ensure labels exist for expected keys when blank or matching key
    if(labels && labels.length){
      for(let i=0;i<expectedKeys.length;i++){
        const key = expectedKeys[i];
        const col = haveKeys.indexOf(key);
        if(col === -1) continue;
        const current = haveLabels[col];
        const desired = labels[i];
        if(!current || current === key){
          sheet.getRange(2,col+1).setValue(desired);
        }
      }
    }
  }
  invalidateHeaderCache_(sheetName);
  return getHeaderMap(sheetName);
}

function verifyHeaders(){
  const issues = [];
  if(typeof SHEET_HEADER_DEFS === 'object'){
    Object.keys(SHEET_HEADER_DEFS).forEach(name=>{
      const def = SHEET_HEADER_DEFS[name];
      const sheet = sh(name);
      if(!sheet){ issues.push('Missing sheet: '+name); return; }
      const map = getHeaderMap(sheet);
      def.keys.forEach(key=>{ if(!map[key]) issues.push(name+' missing column '+key); });
    });
  }
  if(issues.length){ issues.forEach(msg=>Logger.log(msg)); toast('Header issues logged'); }
  else { Logger.log('Headers OK'); toast('Headers verified'); }
}

function writeRows(sheetName, rowsArray){
  if(!rowsArray || !rowsArray.length) return {startRow:0, endRow:0};
  const sheet = ensure(sheetName);
  const map = getHeaderMap(sheet);
  const order = map.__order;
  const width = order.length;
  if(!width) throw new Error('No headers defined for '+sheetName);
  const rowCount = rowsArray.length;
  const buffer = new Array(rowCount);
  for(let r=0;r<rowCount;r++){
    const obj = rowsArray[r] || {};
    const row = new Array(width);
    for(let c=0;c<width;c++){
      const key = order[c];
      row[c] = key ? (obj[key] !== undefined ? obj[key] : '') : '';
    }
    buffer[r] = row;
  }
  const lastRow = sheet.getLastRow();
  const startRow = Math.max(3, lastRow+1);
  const neededRows = startRow + rowCount - sheet.getMaxRows();
  if(neededRows > 0){ sheet.insertRowsAfter(sheet.getMaxRows(), neededRows); }
  sheet.getRange(startRow,1,rowCount,width).setValues(buffer);
  return {startRow, endRow:startRow+rowCount-1};
}

function updateRow(sheetName, rowIndex, rowObject){
  if(!rowIndex || rowIndex<3) throw new Error('Row index out of bounds for '+sheetName);
  const sheet = sh(sheetName);
  if(!sheet) throw new Error('Sheet not found: '+sheetName);
  const map = getHeaderMap(sheet);
  const order = map.__order;
  const width = order.length;
  const row = new Array(width);
  for(let c=0;c<width;c++){
    const key = order[c];
    row[c] = key ? (rowObject[key] !== undefined ? rowObject[key] : '') : '';
  }
  sheet.getRange(rowIndex,1,1,width).setValues([row]);
  return {row:rowIndex};
}

function readByKey(sheetName, keyField, keyValue){
  const sheet = sh(sheetName);
  if(!sheet) return null;
  const map = getHeaderMap(sheet);
  const col = map[keyField];
  if(!col) return null;
  const lastRow = sheet.getLastRow();
  if(lastRow < 3) return null;
  const valueRange = sheet.getRange(3,col,lastRow-2,1).getValues();
  const target = keyValue;
  for(let i=0;i<valueRange.length;i++){
    const candidate = valueRange[i][0];
    if(candidate === '' && target === ''){ return readRow_(sheet, map, i+3); }
    if(candidate === target){ return readRow_(sheet, map, i+3); }
    if(candidate !== null && candidate !== '' && target !== null && target !== '' && String(candidate) === String(target)){
      return readRow_(sheet, map, i+3);
    }
  }
  return null;
}

function readRow_(sheet, map, rowIndex){
  const width = map.__order.length;
  const values = sheet.getRange(rowIndex,1,1,width).getValues()[0];
  const obj = Object.create(null);
  const order = map.__order;
  for(let i=0;i<order.length;i++){
    const key = order[i];
    if(key) obj[key] = values[i];
  }
  obj._row = rowIndex;
  return {row:rowIndex, values, object:obj};
}

function esc(value){
  if(value === null || value === undefined){ return ''; }
  return String(value)
    .replace(/&/g,'&amp;')
    .replace(/</g,'&lt;')
    .replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;')
    .replace(/'/g,'&#39;');
}

function joinDeviceSummary(devices){
  if(!Array.isArray(devices) || !devices.length){ return ''; }
  const entries = [];
  for(let i=0;i<devices.length;i++){
    const item = devices[i];
    if(!item) continue;
    const model = String(item.model || '').trim();
    if(!model) continue;
    const price = Number(item.price_net != null ? item.price_net : item.price);
    const hasPrice = isFinite(price) && price > 0;
    entries.push(hasPrice ? (model+'='+price) : model);
  }
  if(!entries.length){ return ''; }
  return 'Devices: '+entries.join('; ');
}

function cfg(key, defVal){ const c=sh('Config'); if(!c) return defVal; const f=c.createTextFinder('^'+key+'$').useRegularExpression(true).findNext(); if(!f) return defVal; const v=c.getRange(f.getRow(),2).getValue(); return v===undefined||v===null?defVal:v; }
function setCfg(key,val){ const c=ensure('Config'); const f=c.createTextFinder('^'+key+'$').useRegularExpression(true).findNext(); if(f){ c.getRange(f.getRow(),2).setValue(val); } else { c.appendRow([key,val]); } }
