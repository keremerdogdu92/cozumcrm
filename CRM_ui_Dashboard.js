/* v3.5.1 | Dashboard helpers hardened for header-driven sheets */
function refreshDashboard(){
  const d=ensure('Dashboard'); d.clear();
  d.getRange(1,1).setValue('Now: '+fmt(now()));
  let r=3;

  const rev=listRevenueSummary_(); d.getRange(r,1).setValue('Revenue (net)'); r++; d.getRange(r,1,rev.length,rev[0].length).setValues(rev); r+=rev.length+1;
  const inv=listPendingInvoice_(); d.getRange(r,1).setValue('Pending Invoices'); r++; d.getRange(r,1,inv.length,inv[0].length).setValues(inv); r+=inv.length+1;
  const sgk1=listSGKReportMissing_(); d.getRange(r,1).setValue('SGK Report Missing'); r++; d.getRange(r,1,sgk1.length,sgk1[0].length).setValues(sgk1); r+=sgk1.length+1;
  const sgk2=listSGKSystemMissing_(); d.getRange(r,1).setValue('SGK System Not Entered'); r++; d.getRange(r,1,sgk2.length,sgk2[0].length).setValues(sgk2); r+=sgk2.length+1;
  const stale=listStalePatients_(); d.getRange(r,1).setValue('Stale Patients'); r++; d.getRange(r,1,stale.length,stale[0].length).setValues(stale); r+=stale.length+1;
  const miss=listSalesMissingFields_(); d.getRange(r,1).setValue('Sales Missing Fields (serial/barcode/price_net)'); r++; d.getRange(r,1,miss.length,miss[0].length).setValues(miss); r+=miss.length+1;
  const up=listUpcoming_(); d.getRange(r,1).setValue('Upcoming Appointments'); r++; d.getRange(r,1,up.length,up[0].length).setValues(up);

  d.autoResizeColumns(1,10);
}

function listRevenueSummary_(){
  const sheet=sh('Satışlar'); if(!sheet) return [['metric','value']];
  const map=getHeaderMap(sheet); const values=sheet.getDataRange().getValues();
  const idx = key => (map[key] ? map[key]-1 : null);
  const saleDateIdx = idx('sale_date');
  const priceNetIdx = idx('price_net');
  const paymentPlaceIdx = idx('payment_place');
  const paymentMethodIdx = idx('payment_method');
  const out=[['metric','value']]; let week=0, month=0;
  const d0=now(); const weekStart=new Date(d0); weekStart.setDate(d0.getDate()-(d0.getDay()+6)%7);
  const monthStart=new Date(d0.getFullYear(), d0.getMonth(), 1);
  const byPlace={}, byMethod={};
  for(let i=2;i<values.length;i++){
    const row=values[i]; const dateVal=saleDateIdx!=null?row[saleDateIdx]:null; const netVal=priceNetIdx!=null?Number(row[priceNetIdx]||0):0;
    if(!dateVal || !isFinite(netVal)) continue; const dd=new Date(dateVal);
    if(dd>=weekStart) week+=netVal; if(dd>=monthStart) month+=netVal;
    const place=(paymentPlaceIdx!=null?row[paymentPlaceIdx]:'')||'-';
    const method=(paymentMethodIdx!=null?row[paymentMethodIdx]:'')||'-';
    byPlace[place]=(byPlace[place]||0)+netVal;
    byMethod[method]=(byMethod[method]||0)+netVal;
  }
  out.push(['revenue_week',week],['revenue_month',month]);
  out.push(['breakdown_by_place','value']); Object.keys(byPlace).forEach(k=>out.push([k,byPlace[k]]));
  out.push(['breakdown_by_method','value']); Object.keys(byMethod).forEach(k=>out.push([k,byMethod[k]]));
  return out;
}

function listPendingInvoice_(){
  const sheet=sh('Satışlar'); if(!sheet) return [['sale_id','patient','net','date']];
  const map=getHeaderMap(sheet); const vals=sheet.getDataRange().getValues();
  const idx = key => (map[key] ? map[key]-1 : null);
  const saleIdIdx=idx('sale_id');
  const patientIdx=idx('patient_id');
  const netIdx=idx('price_net');
  const dateIdx=idx('sale_date');
  const invoiceIdx=idx('invoice_issued');
  if(saleIdIdx==null || patientIdx==null || netIdx==null || dateIdx==null || invoiceIdx==null){
    return [['sale_id','patient_id','net','date']];
  }
  const out=[['sale_id','patient_id','net','date']];
  for(let i=2;i<vals.length;i++){
    const r=vals[i]; if(r[invoiceIdx]===true) continue;
    out.push([r[saleIdIdx]||'', r[patientIdx]||'', r[netIdx]||'', r[dateIdx]||'']);
  }
  if(out.length===1) out.push(['-','-','-','-']); return out;
}

function listSGKReportMissing_(){
  const sheet=sh('Satışlar'); if(!sheet) return [['sale_id','patient_id','uses_sgk','report_received']];
  const map=getHeaderMap(sheet); const vals=sheet.getDataRange().getValues();
  const idx = key => (map[key] ? map[key]-1 : null);
  const saleIdIdx=idx('sale_id');
  const patientIdx=idx('patient_id');
  const usesIdx=idx('uses_sgk');
  const repIdx=idx('sgk_report_received');
  if(saleIdIdx==null || patientIdx==null || usesIdx==null || repIdx==null){
    return [['sale_id','patient_id','uses_sgk','report_received']];
  }
  const out=[['sale_id','patient_id','uses_sgk','report_received']];
  for(let i=2;i<vals.length;i++){
    const r=vals[i]; if(r[usesIdx]===true && r[repIdx]!==true){ out.push([r[saleIdIdx]||'', r[patientIdx]||'', true, Boolean(r[repIdx])]); }
  }
  if(out.length===1) out.push(['-','-','-','-']); return out;
}

function listSGKSystemMissing_(){
  const sheet=sh('Satışlar'); if(!sheet) return [['sale_id','patient_id','report_received','system_entered']];
  const map=getHeaderMap(sheet); const vals=sheet.getDataRange().getValues();
  const idx = key => (map[key] ? map[key]-1 : null);
  const saleIdIdx=idx('sale_id');
  const patientIdx=idx('patient_id');
  const usesIdx=idx('uses_sgk');
  const repIdx=idx('sgk_report_received');
  const sysIdx=idx('sgk_system_entered');
  if(saleIdIdx==null || patientIdx==null || usesIdx==null || repIdx==null || sysIdx==null){
    return [['sale_id','patient_id','report_received','system_entered']];
  }
  const out=[['sale_id','patient_id','report_received','system_entered']];
  for(let i=2;i<vals.length;i++){
    const r=vals[i]; if(r[usesIdx]===true && r[repIdx]===true && r[sysIdx]!==true){ out.push([r[saleIdIdx]||'', r[patientIdx]||'', true, Boolean(r[sysIdx])]); }
  }
  if(out.length===1) out.push(['-','-','-','-']); return out;
}

function listStalePatients_(){
  const sheet=sh('Hastalar'); if(!sheet) return [['patient_id','name','days_since_last','limit_effective','flag']];
  const map=getHeaderMap(sheet); const vals=sheet.getDataRange().getValues();
  const idx = key => (map[key] ? map[key]-1 : null);
  const idIdx=idx('patient_id');
  const nameIdx=idx('full_name');
  const daysIdx=idx('days_since_last');
  const limitIdx=idx('limit_days_effective');
  const flagIdx=idx('stale_flag');
  if(idIdx==null || nameIdx==null || daysIdx==null || limitIdx==null || flagIdx==null){
    return [['patient_id','name','days_since_last','limit_effective','flag']];
  }
  const out=[['patient_id','name','days_since_last','limit_effective','flag']];
  for(let i=2;i<vals.length;i++){
    const r=vals[i]; if(r[flagIdx]==='STALE'){ out.push([r[idIdx]||'', r[nameIdx]||'', r[daysIdx]||'', r[limitIdx]||'', r[flagIdx]||'']); }
  }
  if(out.length===1) out.push(['-','-','-','-','-']); return out;
}

function listSalesMissingFields_(){
  const sheet=sh('Satışlar'); if(!sheet) return [['sale_id','patient_id','missing']];
  const map=getHeaderMap(sheet); const vals=sheet.getDataRange().getValues();
  const idx = key => (map[key] ? map[key]-1 : null);
  const saleIdIdx=idx('sale_id');
  const patientIdx=idx('patient_id');
  const serialIdx=idx('serial_no');
  const barcodeIdx=idx('barcode');
  const netIdx=idx('price_net');
  if(saleIdIdx==null || patientIdx==null || serialIdx==null || barcodeIdx==null || netIdx==null){
    return [['sale_id','patient_id','missing']];
  }
  const out=[['sale_id','patient_id','missing']];
  for(let i=2;i<vals.length;i++){
    const r=vals[i]; const miss=[];
    if(!r[serialIdx]) miss.push('serial_no');
    if(!r[barcodeIdx]) miss.push('barcode');
    if(!r[netIdx]) miss.push('price_net');
    if(miss.length) out.push([r[saleIdIdx]||'', r[patientIdx]||'', miss.join(',')]);
  }
  if(out.length===1) out.push(['-','-','-']); return out;
}

function listUpcoming_(){
  const sheet=sh('Randevular'); if(!sheet) return [['when','title','who_type','who_id','status','location']];
  const map=getHeaderMap(sheet); const vals=sheet.getDataRange().getValues();
  const idx = key => (map[key] ? map[key]-1 : null);
  const statusIdx=idx('status');
  const startIdx=idx('start_datetime');
  const titleIdx=idx('title');
  const whoTypeIdx=idx('who_type');
  const whoIdIdx=idx('who_id');
  const locationIdx=idx('location');
  if(statusIdx==null || startIdx==null || titleIdx==null || whoTypeIdx==null || whoIdIdx==null || locationIdx==null){
    return [['when','title','who_type','who_id','status','location']];
  }
  const days=Number(cfg('UPCOMING_DAYS',7))||7; const t0=now(); const t1=new Date(t0.getTime()+days*24*3600*1000);
  const out=[['when','title','who_type','who_id','status','location']];
  for(let i=2;i<vals.length;i++){
    const r=vals[i]; const st=r[statusIdx], sd=r[startIdx];
    if(st==='scheduled' && sd && sd>=t0 && sd<=t1){ out.push([sd,r[titleIdx]||'',r[whoTypeIdx]||'',r[whoIdIdx]||'',st,r[locationIdx]||'']); }
  }
  if(out.length===1) out.push(['-','-','-','-','-','-']); return out;
}
