/* v3.0.0 */
function onEdit(e){
  try{
    const s=e.range.getSheet(); const n=s.getName();
    if(n==='Hastalar') return onEditPatients_(e);
    if(n==='Satışlar') return onEditSales_(e);
    if(n==='Randevular') return onEditAppointments_(e);
  }catch(err){ Logger.log(err); }
}
