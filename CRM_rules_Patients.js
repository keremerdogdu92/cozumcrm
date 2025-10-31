/* v3.0.0 */
function onEditPatients_(e){
  const s=e.range.getSheet(); const idx=headerIndex_(s); const row=e.range.getRow(); if(row<=2) return;
  const tc=s.getRange(row,idx['tc']).getValue();
  if(String(tc).length && !/^\d{11}$/.test(String(tc))){ toast('TC must be 11 digits'); s.getRange(row,idx['tc']).setValue(''); return; }
  const editedCol=e.range.getColumn();
  const editedHead=Object.keys(idx).find(k=>idx[k]===editedCol) || '';
  if(['purchase_date','paid_amount_last'].includes(editedHead)){
    const priceNet=s.getRange(row,idx['paid_amount_last']).getValue();
    const purchaseDate=s.getRange(row,idx['purchase_date']).getValue() || now();
    if(tc && (priceNet || purchaseDate)){ createOrTouchSaleForPatient_(tc, purchaseDate, Number(priceNet)||0); }
  }
}
function createOrTouchSaleForPatient_(tc, saleDate, priceNet){
  const sales=sh('Satışlar'); const idx=headerIndex_(sales);
  for(let r=3;r<=sales.getLastRow();r++){
    const pid=sales.getRange(r,idx['patient_id']).getValue();
    const sd=sales.getRange(r,idx['sale_date']).getValue();
    if(String(pid)===String(tc) && sd && fmt(sd).slice(0,10)===fmt(saleDate).slice(0,10)){
      if(priceNet) sales.getRange(r,idx['price_net']).setValue(priceNet);
      return;
    }
  }
  const saleId=makeSaleId_();
  const row=[ saleId, tc, '', '', '', saleDate, '', '', priceNet, '', '', false, '', '', false, false, false, '' ];
  sales.appendRow(row);
  toast('Sale created: '+saleId+' for TC '+tc);
}
