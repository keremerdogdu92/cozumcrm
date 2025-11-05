/** /AppsScript/CRM/QuickEntry.js
 * v3.7.0
 * CHANGELOG:
 * - Reworked qe_upsertTrialCase for single-row trials, per-visit logs, and secure device offers.
 * - Added strict validation, note trimming, and device pricing support without touching sales sheets.
 * - Updated smoke tests to reflect trial-contact logging and header resilience.
 * HOWTO (Script Editor):
 * 1. Run installCRM() once to seed headers/validations, then open the sidebar via CRM → Trial Quick Entry.
 * 2. Execute qe_testA()/qe_testB()/qe_testC() from the Apps Script editor to simulate different visit payloads.
 * 3. Use qe_testD() to scramble/restore headers and ensure qe_upsertTrialCase keeps working after column moves.
 */

function openTrialQuickEntry(){
  SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutputFromFile('CRM_ui_QuickEntry').setTitle('CRM • Trial Quick Entry'));
}

function qe_upsertTrialCase(payload){
  const lock = LockService.getDocumentLock();
  const locked = lock.tryLock(10000);
  if(!locked){ throw qe_error_('qe_lock_timeout','Another user is updating records. Please retry in a few seconds.'); }
  try{
    const normalized = qe_normalizePayload_(payload);
    qe_prepareSheets_();
    const stamp = now();

    const patient = qe_upsertPatient_(normalized.person, stamp);
    const trial = qe_upsertTrial_(patient, normalized.trial, normalized.person, stamp);
    qe_appendTrialContactLog_(patient, trial, normalized, stamp);
    qe_writeDevices_(patient, trial, normalized.devices, stamp);
    qe_queueDashboardRefresh_();
    return {ok:true, patient_slug:patient.slug, trial_id:trial.trial_id};
  }catch(err){
    Logger.log(err);
    if(err && err.code){ return {ok:false, code:err.code, message:err.message}; }
    return {ok:false, code:'internal_error', message:(err && err.message) ? err.message : 'Unexpected error'};
  }finally{
    try{ lock.releaseLock(); }catch(e){}
  }
}

function qe_normalizePayload_(payload){
  if(!payload || typeof payload !== 'object'){ throw qe_error_('bad_payload','Payload must be an object'); }
  const personRaw = payload.person || {};
  const trialRaw = payload.trial || {};
  const saleRaw = payload.sale || {};
  const devicesRaw = Array.isArray(payload.devices) ? payload.devices : [];

  const fullName = String(personRaw.full_name || '').trim();
  if(!fullName){ throw qe_error_('missing_full_name','Full name is required'); }
  const slug = toSlugTr(fullName);
  if(!slug){ throw qe_error_('invalid_slug','Full name must contain letters to create a slug'); }

  const ref = String(personRaw.ref || '').trim();
  if(!ref){ throw qe_error_('missing_ref','Reference is required'); }

  const phone = String(personRaw.phone || '').trim();
  if(!phone){ throw qe_error_('missing_phone','Phone is required'); }

  const initialAt = qe_parseDate_(trialRaw.initial_at, 'trial.initial_at');
  const nextAt = qe_parseDate_(trialRaw.next_at, 'trial.next_at');
  const note = String(trialRaw.note || '').trim();
  if(!note){ throw qe_error_('missing_note','Note is required'); }

  const devices = [];
  for(let i=0;i<devicesRaw.length;i++){
    const raw = devicesRaw[i] || {};
    const model = String(raw.model || '').trim();
    const priceValue = raw.price_net != null ? raw.price_net : raw.price;
    const hasPriceValue = !(priceValue === undefined || priceValue === null || String(priceValue).trim() === '');
    if(!model && !hasPriceValue){ continue; }
    if(!model){ throw qe_error_('invalid_device_model','Device model is required'); }
    const price = Number(priceValue);
    if(!isFinite(price) || price <= 0){ throw qe_error_('invalid_device_price','Device price must be greater than zero'); }
    const qtyRaw = raw.qty != null ? Number(raw.qty) : 1;
    const qty = isFinite(qtyRaw) && qtyRaw > 0 ? qtyRaw : 1;
    devices.push({model, price_net: price, qty});
  }
  if(!devices.length){ throw qe_error_('missing_devices','At least one device is required'); }

  const payer = String((saleRaw && saleRaw.payer) || '').trim();
  if(!payer){ throw qe_error_('missing_sale_payer','Payer is required'); }

  const person = {
    full_name: fullName,
    slug,
    phone,
    address: personRaw.address ? String(personRaw.address).trim() : '',
    tc: personRaw.tc ? String(personRaw.tc).trim() : '',
    ref
  };

  const trial = {
    initial_at: initialAt,
    next_at: nextAt,
    note
  };

  return {
    person,
    trial,
    devices,
    sale:{ payer }
  };
}

function qe_prepareSheets_(){
  ['Hastalar','Denemeler','Cihazlar','Görüşmeler'].forEach(ensureHeaderSet);
}

function qe_upsertPatient_(person, stamp){
  const existing = readByKey('Hastalar','slug', person.slug);
  const rowObj = existing ? Object.assign({}, existing.object) : {};
  const patientId = existing ? (existing.object.patient_id || existing.object.tc || Utilities.getUuid()) : Utilities.getUuid();
  rowObj.patient_id = patientId;
  rowObj.full_name = person.full_name;
  rowObj.slug = person.slug;
  rowObj.phone = person.phone;
  rowObj.address = person.address || '';
  rowObj.ref_key = person.ref;
  if(person.tc){ rowObj.tc = person.tc; }
  if(!existing){
    rowObj.created_at = stamp;
    rowObj.updated_at = stamp;
    writeRows('Hastalar',[rowObj]);
    return {patient_id:patientId, slug:person.slug};
  }
  rowObj.created_at = existing.object.created_at || rowObj.created_at || stamp;
  const shouldUpdate = qe_hasRowChanges_(rowObj, existing.object, ['updated_at']);
  if(shouldUpdate){
    rowObj.updated_at = stamp;
    updateRow('Hastalar', existing.row, rowObj);
  }
  return {patient_id:patientId, slug:person.slug, row:existing.row};
}

function qe_upsertTrial_(patient, trial, person, stamp){
  const existing = readByKey('Denemeler','slug', patient.slug);
  const rowObj = existing ? Object.assign({}, existing.object) : {};
  const trialId = existing ? (existing.object.trial_id || Utilities.getUuid()) : Utilities.getUuid();
  rowObj.trial_id = trialId;
  rowObj.full_name = person.full_name;
  rowObj.phone = person.phone;
  if(person.tc){ rowObj.tc = person.tc; }
  rowObj.address = person.address || '';
  const existingStart = existing && existing.object && existing.object.started_at ? existing.object.started_at : null;
  rowObj.started_at = existingStart || trial.initial_at;
  rowObj.status = 'active';
  rowObj.patient_id = patient.patient_id;
  rowObj.slug = patient.slug;
  rowObj.last_contact_at = trial.initial_at;
  rowObj.next_contact_at = trial.next_at;
  if(rowObj.next_action_at !== undefined){ rowObj.next_action_at = trial.next_at; }
  rowObj.ref_key = person.ref;
  const previousCount = Number(existing && existing.object && existing.object.contact_count ? existing.object.contact_count : 0);
  rowObj.contact_count = previousCount + 1;
  const existingNotes = existing ? String(existing.object.notes || '') : '';
  const entry = fmt(trial.initial_at)+' - '+trial.note;
  const combined = existingNotes ? (existingNotes+'\n'+entry) : entry;
  rowObj.notes = qe_trimNotes_(combined);
  if(!existing){
    rowObj.created_at = stamp;
    rowObj.updated_at = stamp;
    writeRows('Denemeler',[rowObj]);
    return {trial_id:trialId};
  }
  rowObj.created_at = existing.object.created_at || rowObj.created_at || stamp;
  const shouldUpdate = qe_hasRowChanges_(rowObj, existing.object, ['updated_at']);
  if(shouldUpdate){
    rowObj.updated_at = stamp;
    updateRow('Denemeler', existing.row, rowObj);
  }
  return {trial_id:trialId, row:existing.row};
}

function qe_appendTrialContactLog_(patient, trial, normalized, stamp){
  const partsNote = qe_buildTrialNote_(normalized.trial.note, normalized.devices, normalized.sale);
  const row = {
    log_id: makePrefixedUuid_('LOG'),
    who_type: 'patient',
    who_id: patient.patient_id,
    method: 'visit',
    when: normalized.trial.initial_at,
    note: partsNote,
    type: 'trial_contact',
    next_action_at: normalized.trial.next_at,
    created_at: stamp
  };
  writeRows('Görüşmeler', [row]);
}

function qe_writeDevices_(patient, trial, devices, stamp){
  if(!devices || !devices.length) return;
  const deviceGroupId = makePrefixedUuid_('DG');
  const rows = new Array(devices.length);
  for(let i=0;i<devices.length;i++){
    const device = devices[i];
    rows[i] = {
      device_group_id: deviceGroupId,
      patient_id: patient.patient_id,
      trial_id: trial.trial_id,
      model: device.model,
      side: '',
      qty: device.qty || 1,
      price_offer: device.price_net,
      device_id: makePrefixedUuid_('DEV'),
      device_group_ref: deviceGroupId,
      created_at: stamp,
      purpose: 'trial'
    };
  }
  writeRows('Cihazlar', rows);
}

function qe_buildTrialNote_(note, devices, sale){
  const parts = [];
  const trimmed = note ? String(note).trim() : '';
  if(trimmed){ parts.push(trimmed); }
  const summary = joinDeviceSummary(devices);
  if(summary){ parts.push(summary); }
  if(sale && sale.payer){ parts.push('Payer: '+sale.payer); }
  return parts.join(' | ');
}

function qe_trimNotes_(text){
  if(!text) return '';
  return text.length > 2000 ? text.slice(-2000) : text;
}

function qe_queueDashboardRefresh_(){
  const cache = CacheService.getScriptCache();
  if(cache.get('qe_dash_refresh')) return;
  cache.put('qe_dash_refresh','1',180);
  const triggers = ScriptApp.getProjectTriggers();
  for(let i=0;i<triggers.length;i++){
    if(triggers[i].getHandlerFunction && triggers[i].getHandlerFunction()==='refreshDashboard'){ return; }
  }
  ScriptApp.newTrigger('refreshDashboard').timeBased().after(60*1000).create();
}

function qe_hasRowChanges_(nextRow, previousRow, ignoreKeys){
  const ignore = Object.create(null);
  if(Array.isArray(ignoreKeys)){
    for(let i=0;i<ignoreKeys.length;i++){ ignore[ignoreKeys[i]] = true; }
  }
  const keys = Object.keys(nextRow);
  for(let i=0;i<keys.length;i++){
    const key = keys[i];
    if(ignore[key]) continue;
    if(!qe_cellEquals_(nextRow[key], previousRow[key])){ return true; }
  }
  return false;
}

function qe_cellEquals_(a, b){
  if(a === b) return true;
  const normalize = value => {
    if(value === undefined || value === null || value === '') return '';
    if(value instanceof Date) return value.getTime();
    if(typeof value === 'number'){ return Number(value); }
    return String(value);
  };
  return normalize(a) === normalize(b);
}

function qe_testA(){
  const name = 'Test A '+Utilities.getUuid().slice(0,8);
  return qe_upsertTrialCase({
    person:{ full_name:name, phone:'+905551112233', ref:'test', address:'Line 1' },
    trial:{ initial_at: now(), next_at: new Date(Date.now()+86400000), note:'Initial consultation' },
    devices:[{model:'Alpha', price_net:5000}],
    sale:{ payer:'SGK' }
  });
}

function qe_testB(){
  const name = 'Test B '+Utilities.getUuid().slice(0,8);
  return qe_upsertTrialCase({
    person:{ full_name:name, phone:'+905550000001', ref:'expo' },
    trial:{ initial_at: now(), next_at: new Date(Date.now()+2*86400000), note:'Two device offer' },
    devices:[{model:'Alpha-L', price_net:4800},{model:'Alpha-R', price_net:4900}],
    sale:{ payer:'Private' }
  });
}

function qe_testC(){
  const name = 'Test C '+Utilities.getUuid().slice(0,6);
  const first = qe_upsertTrialCase({
    person:{ full_name:name, phone:'+905554444444', ref:'campaign' },
    trial:{ initial_at: now(), next_at: new Date(Date.now()+3600000), note:'first visit' },
    devices:[{model:'Beta', price_net:6200}],
    sale:{ payer:'Private' }
  });
  const second = qe_upsertTrialCase({
    person:{ full_name:name, phone:'+905554444445', ref:'campaign' },
    trial:{ initial_at: now(), next_at: new Date(Date.now()+2*3600000), note:'follow-up visit' },
    devices:[{model:'Beta', price_net:6400}],
    sale:{ payer:'Private' }
  });
  return {first, second};
}

function qe_testD(){
  const sheets=['Hastalar','Denemeler','Cihazlar','Görüşmeler'];
  const state = sheets.map(name=>({name, order:(sh(name)?getHeaderMap(name).__order.slice():[])}));
  try{
    sheets.forEach(name=>qe_scrambleSheetColumns_(name));
    const a = qe_testA();
    const b = qe_testB();
    return {a,b};
  }finally{
    state.forEach(info=>qe_restoreSheetColumns_(info));
  }
}

function qe_scrambleSheetColumns_(name){
  const sheet = sh(name); if(!sheet) return;
  const map = getHeaderMap(sheet);
  const order = map.__order;
  if(order.length < 3) return;
  const lastHeader = order[order.length-1];
  const lastCol = map[lastHeader];
  if(!lastCol) return;
  sheet.moveColumns(sheet.getRange(1,lastCol,sheet.getMaxRows(),1),2);
  invalidateHeaderCache_(name);
}

function qe_restoreSheetColumns_(info){
  const sheet = sh(info.name); if(!sheet) return;
  const desired = info.order;
  if(!desired || !desired.length) return;
  for(let idx=0; idx<desired.length; idx++){
    const key = desired[idx];
    if(!key) continue;
    const map = getHeaderMap(sheet);
    const current = map[key];
    if(!current) continue;
    const dest = idx+1;
    if(current !== dest){
      sheet.moveColumns(sheet.getRange(1,current,sheet.getMaxRows(),1), dest);
      invalidateHeaderCache_(sheet.getName());
    }
  }
}

/*
HOW TO TEST (Trial Quick Entry)
1. Run installCRM(), open the Trial Quick Entry sidebar, and submit the form with required fields and at least one device.
2. After each submission inspect Hastalar (updated contact info), Denemeler (single trial row with incremented contact_count), Cihazlar (new purpose="trial" offers), and Görüşmeler (appended trial_contact log).
3. Re-run qe_upsertTrialCase with the same full name to confirm Denemeler stays single-row while notes and contact_count grow.
*/
