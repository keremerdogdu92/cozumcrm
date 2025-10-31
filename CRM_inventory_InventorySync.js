/* v3.0.0 */
function updateInventoryStatus(){
  const inv=sh('Inventory'), sales=sh('Satışlar');
  const invData=inv.getDataRange().getValues(), hI=invData[0];
  const siI=hI.indexOf('serial_no'), stI=hI.indexOf('status');
  const sold=new Set();
  const sData=sales.getDataRange().getValues(), hS=sData[0], siS=hS.indexOf('serial_no');
  for(let i=2;i<sData.length;i++){ const sn=sData[i][siS]; if(sn) sold.add(String(sn)); }
  for(let j=1;j<invData.length;j++){
    const sn=String(invData[j][siI]||'');
    const newSt= sold.has(sn)?'sold':(invData[j][stI]||'in_stock');
    inv.getRange(j+1,stI+1).setValue(newSt);
  }
}
