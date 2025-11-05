/* v3.6.0 */
function makeSaleId_(){ let n=Number(cfg('SALES_SEQ',0))|0; n+=1; setCfg('SALES_SEQ',n); return 'S-'+('000000'+n).slice(-6); }
function makeDeviceGroupId_(tc){ return 'DG-'+String(tc)+'-'+ymdhm(now())+'-'+rand2(); }
function makePrefixedUuid_(prefix){ return String(prefix||'ID')+'-'+Utilities.getUuid(); }
