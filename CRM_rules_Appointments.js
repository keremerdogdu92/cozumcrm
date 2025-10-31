/* v3.0.0 */
function onEditAppointments_(e){
  const s=e.range.getSheet(); const idx=headerIndex_(s); const row=e.range.getRow(); if(row<=2) return;
  const status=s.getRange(row,idx['status']).getValue(); if(status!=='done') return;
  const whoId=s.getRange(row,idx['who_id']).getValue();
  const when=s.getRange(row,idx['end_datetime']).getValue() || s.getRange(row,idx['start_datetime']).getValue() || now();
  const title=s.getRange(row,idx['title']).getValue();
  const ish=sh('Görüşmeler'); const iIdx=headerIndex_(ish); const newR=Math.max(3, ish.getLastRow()+1);
  if(ish.getLastRow()<3) ish.insertRowsAfter(2,1);
  ish.getRange(newR,iIdx['log_id']).setValue(Utilities.getUuid());
  ish.getRange(newR,iIdx['who_type']).setValue('patient');
  ish.getRange(newR,iIdx['who_id']).setValue(whoId);
  ish.getRange(newR,iIdx['method']).setValue('visit');
  ish.getRange(newR,iIdx['when']).setValue(when);
  ish.getRange(newR,iIdx['patient_note']).setValue('Appt done: '+title);
}
