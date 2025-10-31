/* v3.0.0 */
function sh(n){ return SpreadsheetApp.getActive().getSheetByName(n); }
function ensure(n){ return sh(n) || SpreadsheetApp.getActive().insertSheet(n); }
function now(){ return new Date(); }
function fmt(d){ return Utilities.formatDate(new Date(d), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'); }
function ymdhm(d){ const z=n=>('0'+n).slice(-2); return d.getFullYear()+z(d.getMonth()+1)+z(d.getDate())+z(d.getHours())+z(d.getMinutes()); }
function rand2(){ const c='ABCDEFGHJKLMNPQRSTUVWXYZ23456789'; return c[Math.floor(Math.random()*c.length)]+c[Math.floor(Math.random()*c.length)]; }
function colLetter_(col){ let s=''; while(col>0){ let m=(col-1)%26; s=String.fromCharCode(65+m)+s; col=(col-m-1)/26|0; } return s; }
function toast(msg){ SpreadsheetApp.getActive().toast(String(msg)); }
function headerIndex_(s){ const head=s.getRange(1,1,1,s.getLastColumn()).getValues()[0]; const m={}; head.forEach((h,i)=>m[String(h).trim()]=i+1); return m; }
function cfg(key, defVal){ const c=sh('Config'); if(!c) return defVal; const f=c.createTextFinder('^'+key+'$').useRegularExpression(true).findNext(); if(!f) return defVal; return c.getRange(f.getRow(),2).getValue() ?? defVal; }
function setCfg(key,val){ const c=ensure('Config'); const f=c.createTextFinder('^'+key+'$').useRegularExpression(true).findNext(); if(f){ c.getRange(f.getRow(),2).setValue(val); } else { c.appendRow([key,val]); } }
