/** /AppsScript/scripts/verify_headers.gs
 * v3.6.1
 * CHANGELOG:
 * - Updated for minimal Quick Entry columns while keeping dry-run/default behaviour intact.
 * - Still honours SHEET_HEADER_DEFS ordering and optional auto-move execution.
 */

function verifyHeadersScript(dryRun){
  const isDryRun = (dryRun === undefined) ? true : !(!dryRun);
  const summary = [];
  if(typeof SHEET_HEADER_DEFS !== 'object'){ Logger.log('SHEET_HEADER_DEFS missing'); return summary; }
  const names = Object.keys(SHEET_HEADER_DEFS);
  for(let i=0;i<names.length;i++){
    const sheetName = names[i];
    const def = SHEET_HEADER_DEFS[sheetName];
    if(!def || !def.keys || !def.keys.length) continue;
    const sheet = sh(sheetName);
    if(!sheet){
      summary.push({sheet:sheetName, status:'missing_sheet'});
      Logger.log('Missing sheet: '+sheetName);
      continue;
    }
    const expected = def.keys.filter(Boolean);
    const operations = [];
    const map = getHeaderMap(sheet);
    let cursor = 1;
    for(let k=0;k<expected.length;k++){
      const key = expected[k];
      const currentCol = map[key];
      if(!currentCol){
        operations.push({type:'missing', key});
        Logger.log(sheetName+' missing column '+key);
        cursor++;
        continue;
      }
      if(currentCol !== cursor){
        operations.push({type:'move', key, from:currentCol, to:cursor});
      }
      cursor++;
    }
    if(!operations.length){
      summary.push({sheet:sheetName, status:'ok'});
      continue;
    }
    summary.push({sheet:sheetName, status:isDryRun ? 'dry_run' : 'reordered', operations:operations.slice()});
    if(isDryRun) continue;
    for(let opIndex=0; opIndex<operations.length; opIndex++){
      const op = operations[opIndex];
      if(op.type !== 'move') continue;
      const liveMap = getHeaderMap(sheet);
      const fromCol = liveMap[op.key];
      if(!fromCol || fromCol === op.to) continue;
      sheet.moveColumns(sheet.getRange(1,fromCol,sheet.getMaxRows(),1), op.to);
      invalidateHeaderCache_(sheetName);
    }
  }
  if(isDryRun){
    Logger.log('Dry run only. Call verifyHeadersScript(false) to apply moves.');
  }else{
    toast('Header order normalised');
  }
  return summary;
}
