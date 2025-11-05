/* v3.5.0 */
function debugHeaders(){
  const ss=SpreadsheetApp.getActive(); const names=Object.keys(SHEET_HEADER_DEFS); let out=[];
  names.forEach(n=>{
    const s=ss.getSheetByName(n); if(!s){ out.push(n+' MISSING'); return; }
    const row1=s.getRange(1,1,1,Math.max(1,s.getLastColumn())).getValues()[0].join(' | ');
    const row2=s.getRange(2,1,1,Math.max(1,s.getLastColumn())).getValues()[0].join(' | ');
    out.push(n+': r1='+row1); out.push('   r2='+row2);
  });
  Logger.log(out.join('\n')); toast('Debug → View Logs');
}
