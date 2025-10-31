/* v3.4.0 | Patient Card: add total_spend in patient summary; keep DTO stable */
function openPatientCard(){
  SpreadsheetApp.getUi()
    .showSidebar(HtmlService.createHtmlOutputFromFile('UiPatientCard').setTitle('CRM • Hasta Kartı'));
}
function onOpen(){
  SpreadsheetApp.getUi().createMenu('CRM')
    .addItem('Hızlı Giriş','openQuickEntry')
    .addItem('Hasta Kartı','openPatientCard')
    .addItem('Refresh Dashboard','refreshDashboard')
    .addItem('Sync Calendar','syncCalendarFromAppointments')
    .addItem('Update Inventory','updateInventoryStatus')
    .addItem('Send Daily Reminders','sendDailyReminders')
    .addToUi();
}

function pc_searchPeople(q){
  q=String(q||'').toLowerCase().trim();
  const out=[]; const P=sh('Hastalar'); if(!P) return out;
  const idx=headerIndex_(P); const vals=P.getDataRange().getValues();
  for(let r=3;r<vals.length;r++){
    const name=String(vals[r][idx['full_name']-1]||'');
    const tc  =String(vals[r][idx['tc']-1]||'');
    if(!name||!tc) continue;
    if(!q || name.toLowerCase().includes(q) || tc.includes(q)) out.push({tc, name});
  }
  return out.slice(0,50);
}

function pc_getPatientCard(tc){
  tc=String(tc||'').trim();
  const nowD = new Date();

  const P=sh('Hastalar'); const pIdx=headerIndex_(P);
  let prow=null, pRowVals=null;
  for(let r=3;r<=P.getLastRow();r++){
    if(String(P.getRange(r,pIdx['tc']).getValue())===tc){
      prow=r; pRowVals=P.getRange(r,1,1,P.getLastColumn()).getValues()[0]; break;
    }
  }
  if(!prow) throw new Error('Hasta bulunamadı');

  const S=sh('Satışlar'); const sIdx=headerIndex_(S);
  const sVals=S.getDataRange().getValues(); const sales=[];
  for(let i=1;i<sVals.length;i++){
    const r=sVals[i]; if(String(r[sIdx['patient_id']-1]||'')!==tc) continue;
    sales.push({
      sale_id: r[sIdx['sale_id']-1]||'',
      sale_date: r[sIdx['sale_date']-1]||'',
      price_net: Number(r[sIdx['price_net']-1]||0),
      payment_method: r[sIdx['payment_method']-1]||'',
      payment_place: r[sIdx['payment_place']-1]||'',
      invoice_issued: r[sIdx['invoice_issued']-1]===true,
      uses_sgk: r[sIdx['uses_sgk']-1]===true,
      sgk_report_received: r[sIdx['sgk_report_received']-1]===true,
      sgk_system_entered: r[sIdx['sgk_system_entered']-1]===true,
      serial_no: r[sIdx['serial_no']-1]||'',
      barcode:   r[sIdx['barcode']-1]||'',
      notes:     r[sIdx['notes']-1]||''
    });
  }

  const I=sh('Görüşmeler'); const iIdx=headerIndex_(I);
  const iVals=I.getDataRange().getValues(); const interactions=[];
  for(let i=1;i<iVals.length;i++){
    const r=iVals[i];
    if(String(r[iIdx['who_type']-1]||'')!=='patient') continue;
    if(String(r[iIdx['who_id']-1]||'')!==tc) continue;
    interactions.push({
      when: r[iIdx['when']-1]||'',
      method: r[iIdx['method']-1]||'',
      note: r[iIdx['patient_note']-1]||'',
      satisfaction: r[iIdx['satisfaction']-1]
    });
  }

  const A=sh('Randevular'); const aIdx=headerIndex_(A);
  const aVals=A.getDataRange().getValues(); const appts=[];
  for(let i=1;i<aVals.length;i++){
    const r=aVals[i];
    if(String(r[aIdx['who_type']-1]||'')!=='patient') continue;
    if(String(r[aIdx['who_id']-1]||'')!==tc) continue;
    appts.push({
      status: r[aIdx['status']-1]||'',
      title:  r[aIdx['title']-1]||'',
      start:  r[aIdx['start_datetime']-1]||'',
      end:    r[aIdx['end_datetime']-1]||'',
      location: r[aIdx['location']-1]||''
    });
  }

  const D=sh('Cihazlar'); const dIdx=D?headerIndex_(D):{};
  const devices=[];
  if(D){
    const dVals=D.getDataRange().getValues();
    for(let i=1;i<dVals.length;i++){
      const r=dVals[i]; if(String(r[dIdx['patient_id']-1]||'')!==tc) continue;
      devices.push({
        device_group_id: r[dIdx['device_group_id']-1]||'',
        side: r[dIdx['side']-1]||'',
        serial_no: r[dIdx['serial_no']-1]||'',
        barcode: r[dIdx['barcode']-1]||'',
        given_date: r[dIdx['given_date']-1]||'',
        power_type: r[dIdx['power_type']-1]||'',
        notes: r[dIdx['notes']-1]||''
      });
    }
  }

  const Aks=sh('Aksesuarlar'); const axIdx=Aks?headerIndex_(Aks):{};
  const accessories=[];
  if(Aks){
    const axVals=Aks.getDataRange().getValues();
    for(let i=1;i<axVals.length;i++){
      const r=axVals[i]; if(String(r[axIdx['patient_id']-1]||'')!==tc) continue;
      accessories.push({
        device_group_id: r[axIdx['device_group_id']-1]||'',
        type: r[axIdx['type']-1]||'',
        serial_no: r[axIdx['serial_no']-1]||'',
        barcode: r[axIdx['barcode']-1]||'',
        qty: r[axIdx['qty']-1]||'',
        notes: r[axIdx['notes']-1]||''
      });
    }
  }

  sales.sort((a,b)=>new Date(b.sale_date)-new Date(a.sale_date));
  interactions.sort((a,b)=>new Date(b.when)-new Date(a.when));
  appts.sort((a,b)=>new Date(a.start)-new Date(b.start));

  const totalSpend = sales.reduce((s,x)=>s+Number(x.price_net||0),0);

  const patient = {
    tc,
    full_name: pRowVals[pIdx['full_name']-1]||'',
    phone:     pRowVals[pIdx['phone']-1]||'',
    address:   pRowVals[pIdx['address']-1]||'',
    purchase_date: pRowVals[pIdx['purchase_date']-1]||'',
    paid_amount_last: pRowVals[pIdx['paid_amount_last']-1]||'',
    last_contact_at: pRowVals[pIdx['last_contact_at']-1]||'',
    next_contact_at: pRowVals[pIdx['next_contact_at']-1]||'',
    contact_count: pRowVals[pIdx['contact_count']-1]||0,
    days_since_last: pRowVals[pIdx['days_since_last']-1]||'',
    limit_days_effective: pRowVals[pIdx['limit_days_effective']-1]||'',
    stale_flag: pRowVals[pIdx['stale_flag']-1]||'',
    last_payment_place: pRowVals[pIdx['last_payment_place']-1]||'',
    last_payment_method: pRowVals[pIdx['last_payment_method']-1]||'',
    uses_sgk: pRowVals[pIdx['uses_sgk']-1]||false,
    last_sgk_report_received: pRowVals[pIdx['last_sgk_report_received']-1]||false,
    last_sgk_system_entered: pRowVals[pIdx['last_sgk_system_entered']-1]||false,
    satisfaction_last: pRowVals[pIdx['satisfaction_last']-1]||'',
    satisfaction_avg: pRowVals[pIdx['satisfaction_avg']-1]||'',
    notes: pRowVals[pIdx['notes']-1]||'',
    total_spend: totalSpend
  };

  return { now: nowD, patient, sales, interactions: interactions.slice(0,20), appts: appts.slice(0,10), devices, accessories };
}

function pc_renderToSheet(tc){
  const data=pc_getPatientCard(tc);
  const s=ensure('HastaKartı'); s.clear();
  let r=1;

  s.getRange(r,1).setValue('Hasta Kartı');
  s.getRange(r,2).setValue(Utilities.formatDate(data.now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')); r+=2;

  s.getRange(r,1).setValue('Ad Soyad'); s.getRange(r,2).setValue(data.patient.full_name); r++;
  s.getRange(r,1).setValue('TC');       s.getRange(r,2).setValue(data.patient.tc); r++;
  s.getRange(r,1).setValue('Telefon');  s.getRange(r,2).setValue(data.patient.phone); r++;
  s.getRange(r,1).setValue('Adres');    s.getRange(r,2).setValue(data.patient.address); r+=2;

  s.getRange(r,1).setValue('Son Görüşme'); s.getRange(r,2).setValue(data.patient.last_contact_at); r++;
  s.getRange(r,1).setValue('Sonraki Temas'); s.getRange(r,2).setValue(data.patient.next_contact_at); r++;
  s.getRange(r,1).setValue('Görüşme Sayısı'); s.getRange(r,2).setValue(data.patient.contact_count); r++;
  s.getRange(r,1).setValue('Uyarı'); s.getRange(r,2).setValue(data.patient.stale_flag); r++;
  s.getRange(r,1).setValue('Toplam Harcama'); s.getRange(r,2).setValue(data.patient.total_spend); r+=2;

  s.getRange(r,1).setValue('Satışlar (sondan başa)'); r++;
  const sales = [['Satış ID','Tarih','Net','Yöntem','Yer','Fatura','SGK','Seri','Barkod']];
  data.sales.forEach(x=>sales.push([x.sale_id,x.sale_date,x.price_net,x.payment_method,x.payment_place,x.invoice_issued?'Evet':'Hayır',x.uses_sgk?'Evet':'Hayır',x.serial_no,x.barcode]));
  s.getRange(r,1,sales.length,sales[0].length).setValues(sales); r+=sales.length+1;

  s.getRange(r,1).setValue('Cihazlar'); r++;
  const dev = [['GrupID','Taraf','Seri','Barkod','Teslim','Güç','Not']];
  data.devices.forEach(x=>dev.push([x.device_group_id,x.side,x.serial_no,x.barcode,x.given_date,x.power_type,x.notes]));
  s.getRange(r,1,dev.length,dev[0].length).setValues(dev); r+=dev.length+1;

  if(data.accessories.length){
    s.getRange(r,1).setValue('Aksesuarlar'); r++;
    const acc=[['GrupID','Tür','Seri','Barkod','Adet','Not']];
    data.accessories.forEach(x=>acc.push([x.device_group_id,x.type,x.serial_no,x.barcode,x.qty,x.notes]));
    s.getRange(r,1,acc.length,acc[0].length).setValues(acc); r+=acc.length+1;
  }

  s.getRange(r,1).setValue('Görüşmeler (son 20)'); r++;
  const ints=[['Zaman','Yöntem','Not','Memnuniyet']];
  data.interactions.forEach(x=>ints.push([x.when,x.method,x.note,x.satisfaction]));
  s.getRange(r,1,ints.length,ints[0].length).setValues(ints); r+=ints.length+1;

  s.getRange(r,1).setValue('Randevular (yaklaşan + geçmiş)'); r++;
  const ap=[['Durum','Başlık','Başlangıç','Bitiş','Konum']];
  data.appts.forEach(x=>ap.push([x.status,x.title,x.start,x.end,x.location]));
  s.getRange(r,1,ap.length,ap[0].length).setValues(ap);

  s.autoResizeColumns(1, 10);
  SpreadsheetApp.getActiveSpreadsheet().toast('HastaKartı sayfası güncellendi.');
  return {ok:true};
}
