/* v3.5.1 | Patient Card: slug-aware search and header-safe DTO building */
function openPatientCard(){
  SpreadsheetApp.getUi()
    .showSidebar(HtmlService.createHtmlOutputFromFile('UiPatientCard').setTitle('CRM • Hasta Kartı'));
}
function pc_searchPeople(q){
  q=String(q||'').toLowerCase().trim();
  const out=[]; const P=sh('Hastalar'); if(!P) return out;
  const idx=headerIndex_(P); const vals=P.getDataRange().getValues();
  for(let r=3;r<vals.length;r++){
    const name=String(vals[r][idx['full_name']-1]||'');
    const tc  =String(vals[r][idx['tc']-1]||'');
    const slug=idx['slug']?String(vals[r][idx['slug']-1]||''):'';
    if(!name) continue;
    const key = tc || slug;
    if(!key) continue;
    const hay=[name.toLowerCase(), tc, slug].join(' ');
    if(!q || hay.includes(q)) out.push({tc:key, name:slug?`${name} • ${slug}`:name});
  }
  return out.slice(0,50);
}

function pc_getPatientCard(identifier){
  const id=String(identifier||'').trim();
  const nowD = new Date();

  const P=sh('Hastalar'); const pIdx=headerIndex_(P);
  const read = (map, row, key) => {
    if(!map) return '';
    const col = map[key];
    return col ? row[col-1] : '';
  };
  const readBool = (map, row, key) => {
    if(!map) return false;
    const col = map[key];
    return col ? row[col-1] === true : false;
  };
  let prow=null, pRowVals=null;
  for(let r=3;r<=P.getLastRow();r++){
    const row=P.getRange(r,1,1,P.getLastColumn()).getValues()[0];
    const tcVal = String(read(pIdx, row, 'tc')||'');
    const slugVal = String(read(pIdx, row, 'slug')||'');
    if(id && (id===tcVal || id===slugVal)){
      prow=r; pRowVals=row; break;
    }
  }
  if(!prow) throw new Error('Hasta bulunamadı');

  const patientId = read(pIdx, pRowVals, 'patient_id') || id;
  const tc = String(read(pIdx, pRowVals, 'tc')||'');

  const S=sh('Satışlar'); const sIdx=headerIndex_(S);
  const sVals=S.getDataRange().getValues(); const sales=[];
  for(let i=1;i<sVals.length;i++){
    const r=sVals[i]; if(String(read(sIdx, r, 'patient_id')||'')!==String(patientId)) continue;
    sales.push({
      sale_id: read(sIdx, r, 'sale_id')||'',
      sale_date: read(sIdx, r, 'sale_date')||'',
      price_net: Number(read(sIdx, r, 'price_net')||0),
      payment_method: read(sIdx, r, 'payment_method')||'',
      payment_place: read(sIdx, r, 'payment_place')||'',
      invoice_issued: readBool(sIdx, r, 'invoice_issued'),
      uses_sgk: readBool(sIdx, r, 'uses_sgk'),
      sgk_report_received: readBool(sIdx, r, 'sgk_report_received'),
      sgk_system_entered: readBool(sIdx, r, 'sgk_system_entered'),
      serial_no: read(sIdx, r, 'serial_no')||'',
      barcode:   read(sIdx, r, 'barcode')||'',
      notes:     read(sIdx, r, 'notes')||''
    });
  }

  const I=sh('Görüşmeler'); const iIdx=headerIndex_(I);
  const iVals=I.getDataRange().getValues(); const interactions=[];
  for(let i=1;i<iVals.length;i++){
    const r=iVals[i];
    if(String(read(iIdx, r, 'who_type')||'')!=='patient') continue;
    if(String(read(iIdx, r, 'who_id')||'')!==String(patientId)) continue;
    const noteVal = read(iIdx, r, 'note') || read(iIdx, r, 'patient_note') || '';
    interactions.push({
      when: read(iIdx, r, 'when')||'',
      method: read(iIdx, r, 'method')||'',
      note: noteVal,
      satisfaction: read(iIdx, r, 'satisfaction')||''
    });
  }

  const A=sh('Randevular'); const aIdx=headerIndex_(A);
  const aVals=A.getDataRange().getValues(); const appts=[];
  for(let i=1;i<aVals.length;i++){
    const r=aVals[i];
    if(String(read(aIdx, r, 'who_type')||'')!=='patient') continue;
    if(String(read(aIdx, r, 'who_id')||'')!==String(patientId)) continue;
    appts.push({
      status: read(aIdx, r, 'status')||'',
      title:  read(aIdx, r, 'title')||'',
      start:  read(aIdx, r, 'start_datetime')||'',
      end:    read(aIdx, r, 'end_datetime')||'',
      location: read(aIdx, r, 'location')||''
    });
  }

  const D=sh('Cihazlar'); const dIdx=D?headerIndex_(D):{};
  const devices=[];
  if(D){
    const dVals=D.getDataRange().getValues();
    for(let i=1;i<dVals.length;i++){
      const r=dVals[i]; if(String(read(dIdx, r, 'patient_id')||'')!==String(patientId)) continue;
      devices.push({
        device_group_id: read(dIdx, r, 'device_group_id')||'',
        side: read(dIdx, r, 'side')||'',
        serial_no: read(dIdx, r, 'serial_no')||'',
        barcode: read(dIdx, r, 'barcode')||'',
        given_date: read(dIdx, r, 'given_date')||'',
        power_type: read(dIdx, r, 'power_type')||'',
        notes: read(dIdx, r, 'notes')||''
      });
    }
  }

  const Aks=sh('Aksesuarlar'); const axIdx=Aks?headerIndex_(Aks):{};
  const accessories=[];
  if(Aks){
    const axVals=Aks.getDataRange().getValues();
    for(let i=1;i<axVals.length;i++){
      const r=axVals[i]; if(String(read(axIdx, r, 'patient_id')||'')!==String(patientId)) continue;
      accessories.push({
        device_group_id: read(axIdx, r, 'device_group_id')||'',
        type: read(axIdx, r, 'type')||'',
        serial_no: read(axIdx, r, 'serial_no')||'',
        barcode: read(axIdx, r, 'barcode')||'',
        qty: read(axIdx, r, 'qty')||'',
        notes: read(axIdx, r, 'notes')||''
      });
    }
  }

  sales.sort((a,b)=>new Date(b.sale_date)-new Date(a.sale_date));
  interactions.sort((a,b)=>new Date(b.when)-new Date(a.when));
  appts.sort((a,b)=>new Date(a.start)-new Date(b.start));

  const totalSpend = sales.reduce((s,x)=>s+Number(x.price_net||0),0);

  const patient = {
    tc,
    full_name: read(pIdx, pRowVals, 'full_name')||'',
    slug: read(pIdx, pRowVals, 'slug')||'',
    patient_id: patientId,
    phone:     read(pIdx, pRowVals, 'phone')||'',
    address:   read(pIdx, pRowVals, 'address')||'',
    purchase_date: read(pIdx, pRowVals, 'purchase_date')||'',
    paid_amount_last: read(pIdx, pRowVals, 'paid_amount_last')||'',
    last_contact_at: read(pIdx, pRowVals, 'last_contact_at')||'',
    next_contact_at: read(pIdx, pRowVals, 'next_contact_at')||'',
    contact_count: read(pIdx, pRowVals, 'contact_count')||0,
    days_since_last: read(pIdx, pRowVals, 'days_since_last')||'',
    limit_days_effective: read(pIdx, pRowVals, 'limit_days_effective')||'',
    stale_flag: read(pIdx, pRowVals, 'stale_flag')||'',
    last_payment_place: read(pIdx, pRowVals, 'last_payment_place')||'',
    last_payment_method: read(pIdx, pRowVals, 'last_payment_method')||'',
    uses_sgk: readBool(pIdx, pRowVals, 'uses_sgk'),
    last_sgk_report_received: readBool(pIdx, pRowVals, 'last_sgk_report_received'),
    last_sgk_system_entered: readBool(pIdx, pRowVals, 'last_sgk_system_entered'),
    satisfaction_last: read(pIdx, pRowVals, 'satisfaction_last')||'',
    satisfaction_avg: read(pIdx, pRowVals, 'satisfaction_avg')||'',
    notes: read(pIdx, pRowVals, 'notes')||'',
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
  s.getRange(r,1).setValue('Slug');     s.getRange(r,2).setValue(data.patient.slug||''); r++;
  s.getRange(r,1).setValue('Patient ID'); s.getRange(r,2).setValue(data.patient.patient_id||''); r++;
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
