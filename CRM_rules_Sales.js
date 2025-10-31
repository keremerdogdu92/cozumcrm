/* v3.0.0 */
function onEditSales_(e){
  const s=e.range.getSheet(); const idx=headerIndex_(s); const row=e.range.getRow(); if(row<=2) return;
  const col=e.range.getColumn();
  const uses=s.getRange(row,idx['uses_sgk']).getValue();
  const rep=s.getRange(row,idx['sgk_report_received']).getValue();
  const sys=s.getRange(row,idx['sgk_system_entered']).getValue();
  if(!uses){
    if(rep===true) s.getRange(row,idx['sgk_report_received']).setValue(false);
    if(sys===true) s.getRange(row,idx['sgk_system_entered']).setValue(false);
  }else{
    if(col===idx['sgk_system_entered'] && sys===true && rep!==true){
      s.getRange(row,idx['sgk_system_entered']).setValue(false);
      toast('Önce "SGK Raporu Geldi" işaretlenmeli.');
    }
  }
}
