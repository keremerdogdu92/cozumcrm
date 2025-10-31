/* v3.1.0 */
// ── Menu hook ─────────────────────────────────────────────────────────────────
function openQuickEntry(){ SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutputFromFile('CRM_ui_QuickEntry').setTitle('CRM • Hızlı Giriş')); }
function onOpen(){ // keep existing items if you already have a Menu.gs; add this line there instead
  const ui=SpreadsheetApp.getUi();
  const m=ui.createMenu('CRM');
  m.addItem('Hızlı Giriş','openQuickEntry')
   .addItem('Refresh Dashboard','refreshDashboard')
   .addItem('Sync Calendar','syncCalendarFromAppointments')
   .addItem('Update Inventory','updateInventoryStatus')
   .addItem('Send Daily Reminders','sendDailyReminders')
   .addToUi();
}

// ── Search API ────────────────────────────────────────────────────────────────
function qe_searchPeople(q){
  q=String(q||'').toLowerCase().trim();
  const out=[];
  // Patients
  const p=sh('Hastalar'); if(p){
    const idx=headerIndex_(p); const vals=p.getDataRange().getValues();
    for(let r=3;r<vals.length;r++){
      const name=String(vals[r][idx['full_name']-1]||''); const tc=String(vals[r][idx['tc']-1]||'');
      if(!name) continue;
      if(!q || name.toLowerCase().includes(q) || tc.includes(q)){
        out.push({type:'patient', id:tc, label:`[Hasta] ${name} • TC:${tc}`});
      }
    }
  }
  // Trials
  const t=sh('Denemeler')||sh('Trials'); if(t){
    const idx=headerIndex_(t); const vals=t.getDataRange().getValues();
    for(let r=3;r<vals.length;r++){
      const name=String(vals[r][idx['full_name']-1]||''); const tid=String(vals[r][idx['trial_id']-1]||'');
      if(!name || !tid) continue;
      if(!q || name.toLowerCase().includes(q) || tid.includes(q)){
        out.push({type:'trial', id:tid, label:`[Deneme] ${name} • ID:${tid}`});
      }
    }
  }
  return out.slice(0,50);
}

// ── Trial create ──────────────────────────────────────────────────────────────
function qe_createTrial(payload){
  const t=ensure('Denemeler'); const idx=headerIndex_(t); if(t.getLastRow()<3) t.insertRowsAfter(2,1);
  const id=Utilities.getUuid();
  const r=Math.max(3,t.getLastRow()+1);
  t.getRange(r,idx['trial_id']).setValue(id);
  t.getRange(r,idx['full_name']).setValue(payload.full_name||'');
  t.getRange(r,idx['phone']).setValue(payload.phone||'');
  t.getRange(r,idx['tc']).setValue(payload.tc||'');
  t.getRange(r,idx['address']).setValue(payload.address||'');
  t.getRange(r,idx['started_at']).setValue(now());
  t.getRange(r,idx['status']).setValue('active');
  t.getRange(r,idx['notes']).setValue(payload.notes||'');
  return {trial_id:id};
}

// ── Patient + Sale create (can be used for “bought during trial”) ────────────
function qe_createPatientWithSale(payload){
  const P=ensure('Hastalar'); const pIdx=headerIndex_(P); if(P.getLastRow()<3) P.insertRowsAfter(2,1);
  const S=ensure('Satışlar'); const sIdx=headerIndex_(S);

  // upsert patient by TC
  const tc=String(payload.tc||'').trim();
  if(!/^\d{11}$/.test(tc)) throw new Error('TC must be 11 digits');
  let prow=null;
  for(let r=3;r<=P.getLastRow();r++){
    if(String(P.getRange(r,pIdx['tc']).getValue())===tc){ prow=r; break; }
  }
  if(!prow){ prow=P.getLastRow()+1; }
  P.getRange(prow,pIdx['full_name']).setValue(payload.full_name||'');
  P.getRange(prow,pIdx['tc']).setValue(tc);
  if(payload.phone)   P.getRange(prow,pIdx['phone']).setValue(payload.phone);
  if(payload.address) P.getRange(prow,pIdx['address']).setValue(payload.address);
  if(payload.purchase_date) P.getRange(prow,pIdx['purchase_date']).setValue(new Date(payload.purchase_date));
  if(payload.price_net!=null) P.getRange(prow,pIdx['paid_amount_last']).setValue(Number(payload.price_net)||0);

  // create sale
  const saleId=makeSaleId_();
  const row=[ saleId, tc, '', '', '', payload.purchase_date?new Date(payload.purchase_date):now(),
              '', '', Number(payload.price_net)||0, payload.installments||'', '',
              false, (payload.payment_method||''), (payload.payment_place||''),
              Boolean(payload.uses_sgk)||false, false, false, payload.notes||'' ];
  S.appendRow(row);

  // close trial if provided
  if(payload.trial_id){
    const T=sh('Denemeler'); if(T){
      const tIdx=headerIndex_(T);
      const f=T.createTextFinder('^'+payload.trial_id+'$').useRegularExpression(true).matchCase(true).findNext();
      if(f){ const r=f.getRow(); T.getRange(r,tIdx['status']).setValue('converted'); }
    }
  }
  return {sale_id:saleId, tc};
}

// ── Interaction add for patient or trial ─────────────────────────────────────
function qe_addInteraction(payload){
  const I=ensure('Görüşmeler'); const iIdx=headerIndex_(I); if(I.getLastRow()<3) I.insertRowsAfter(2,1);
  const newR=I.getLastRow()+1;
  I.getRange(newR,iIdx['log_id']).setValue(Utilities.getUuid());
  I.getRange(newR,iIdx['who_type']).setValue(payload.type); // 'patient' | 'trial'
  I.getRange(newR,iIdx['who_id']).setValue(payload.id);
  I.getRange(newR,iIdx['method']).setValue(payload.method||'call');
  I.getRange(newR,iIdx['when']).setValue(payload.when?new Date(payload.when):now());
  I.getRange(newR,iIdx['patient_note']).setValue(payload.note||'');
  if(payload.satisfaction!=null) I.getRange(newR,iIdx['satisfaction']).setValue(Number(payload.satisfaction));
  if(payload.next_action_at) I.getRange(newR,iIdx['next_action_at']).setValue(new Date(payload.next_action_at));
  return {ok:true};
}
