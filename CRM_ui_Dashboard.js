/* v3.0.0 */
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
  const s=sh('Satışlar').getDataRange().getValues();
  const out=[['metric','value']]; let week=0, month=0;
  const d0=now(), weekStart=new Date(d0); weekStart.setDate(d0.getDate()-(d0.getDay()+6)%7);
  const monthStart=new Date(d0.getFullYear(), d0.getMonth(), 1);
  const byPlace={}, byMethod={};
  for(let i=2;i<s.length;i++){
    const row=s[i]; const d=row[5], net=Number(row[8]||0);
    if(!d || isNaN(net)) continue; const dd=new Date(d);
    if(dd>=weekStart) week+=net; if(dd>=monthStart) month+=net;
    const place=row[13]||'-', method=row[12]||'-';
    byPlace[place]=(byPlace[place]||0)+net;
    byMethod[method]=(byMethod[method]||0)+net;
  }
  out.push(['revenue_week',week],['revenue_month',month]);
  out.push(['breakdown_by_place','value']); Object.keys(byPlace).forEach(k=>out.push([k,byPlace[k]]));
  out.push(['breakdown_by_method','value']); Object.keys(byMethod).forEach(k=>out.push([k,byMethod[k]]));
  return out;
}
function listPendingInvoice_(){
  const s=sh('Satışlar').getDataRange().getValues(); const out=[['sale_id','tc','net','date']];
  for(let i=2;i<s.length;i++){ const r=s[i]; if(r[11]!==true){ out.push([r[0],r[1],r[8],r[5]]); } }
  if(out.length===1) out.push(['-','-','-','-']); return out;
}
function listSGKReportMissing_(){
  const s=sh('Satışlar').getDataRange().getValues(); const out=[['sale_id','tc','uses_sgk','report_received']];
  for(let i=2;i<s.length;i++){ const r=s[i]; if(r[14]===true && r[15]!==true){ out.push([r[0],r[1],true,Boolean(r[15])]); } }
  if(out.length===1) out.push(['-','-','-','-']); return out;
}
function listSGKSystemMissing_(){
  const s=sh('Satışlar').getDataRange().getValues(); const out=[['sale_id','tc','report_received','system_entered']];
  for(let i=2;i<s.length;i++){ const r=s[i]; if(r[14]===true && r[15]===true && r[16]!==true){ out.push([r[0],r[1],true,Boolean(r[16])]); } }
  if(out.length===1) out.push(['-','-','-','-']); return out;
}
function listStalePatients_(){
  const p=sh('Hastalar').getDataRange().getValues(); const out=[['tc','name','days_since_last','limit_effective','flag']];
  for(let i=2;i<p.length;i++){ const r=p[i]; if(r[14]==='STALE'){ out.push([r[1],r[0],r[9],r[13],r[14]]); } }
  if(out.length===1) out.push(['-','-','-','-','-']); return out;
}
function listSalesMissingFields_(){
  const s=sh('Satışlar').getDataRange().getValues(); const out=[['sale_id','tc','missing']];
  for(let i=2;i<s.length;i++){
    const r=s[i]; const miss=[];
    if(!r[3]) miss.push('serial_no');
    if(!r[4]) miss.push('barcode');
    if(!r[8]) miss.push('price_net');
    if(miss.length) out.push([r[0],r[1],miss.join(',')]);
  }
  if(out.length===1) out.push(['-','-','-']); return out;
}
function listUpcoming_(){
  const ap=sh('Randevular').getDataRange().getValues(); const head=ap[0]; const idx={}; head.forEach((h,i)=>idx[h]=i);
  const days=Number(cfg('UPCOMING_DAYS',7))||7; const t0=now(); const t1=new Date(t0.getTime()+days*24*3600*1000);
  const out=[['when','title','who_type','who_id','status','location']];
  for(let i=2;i<ap.length;i++){
    const r=ap[i]; const st=r[idx['status']], sd=r[idx['start_datetime']];
    if(st==='scheduled' && sd && sd>=t0 && sd<=t1){ out.push([sd,r[idx['title']],r[idx['who_type']],r[idx['who_id']],st,r[idx['location']]||'']); }
  }
  if(out.length===1) out.push(['-','-','-','-','-','-']); return out;
}
