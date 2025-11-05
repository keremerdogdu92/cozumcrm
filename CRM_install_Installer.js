/** /AppsScript/CRM/install/Installer.gs
 * v3.8.0
 * CHANGELOG:
 * - Ensured Denemeler captures offer summaries with SGK dropdown validation.
 * - Retained prior trigger setup and device validations while avoiding Cihazlar writes.
 * - Header seeding remains additive without altering existing formulas.
 */
function installCRM(){
  Object.keys(SHEET_HEADER_DEFS).forEach(ensureHeaderSet);

  const c=ensure('Config'); if(c.getLastRow()<2) c.insertRowsAfter(1,1);
  if(c.getLastRow()===2){
    c.getRange(2,1,5,2).setValues([
      ['PATIENT.LIMIT_DEFAULT_DAYS',90],
      ['SATISFACTION_THRESHOLD',7],
      ['UPCOMING_DAYS',7],
      ['REMINDER_HOUR',9],
      ['SALES_SEQ',0]
    ]);
  }

  const listValidation = values => SpreadsheetApp.newDataValidation().setAllowInvalid(false).requireValueInList(values);

  setListValidation_('Görüşmeler','method',['visit','call','msg','payment']);
  setListValidation_('Görüşmeler','type',['trial_contact','payment','visit','call','msg','note']);
  sh('Randevular').getRange('B3:B').setDataValidation(listValidation(['patient','reference','trial']));
  sh('Randevular').getRange('I3:I').setDataValidation(listValidation(['scheduled','done','cancelled','rescheduled']));

  sh('Satışlar').getRange('L3:L').insertCheckboxes();
  sh('Satışlar').getRange('M3:M').setDataValidation(listValidation(['kart','nakit','taksit']));
  sh('Satışlar').getRange('N3:N').setDataValidation(listValidation(['firma','çözüm']));
  sh('Satışlar').getRange('O3:O').insertCheckboxes();
  sh('Satışlar').getRange('P3:P').insertCheckboxes();
  sh('Satışlar').getRange('Q3:Q').insertCheckboxes();

  setListValidation_('Cihazlar','side',['Sağ','Sol','Tek']);
  setListValidation_('Cihazlar','power_type',['tek','pilli','şarjlı']);
  setListValidation_('Cihazlar','purpose',['trial','sale']);

  sh('Stok').getRange('D3:D').setDataValidation(listValidation(['in_stock','reserved','sold','returned','service']));

  setListValidation_('Denemeler','status',['active','converted','lost']);
  setListValidation_('Denemeler','uses_sgk',['Var','Yok']);

  const salesSheet = sh('Satışlar');
  const salesHeaders = headerIndex_(salesSheet);
  const pricingCol = salesHeaders['pricing_mode'];
  if(pricingCol){
    const pricingRangeRows = Math.max(0, salesSheet.getMaxRows()-2);
    if(pricingRangeRows > 0){
      salesSheet.getRange(3, pricingCol, pricingRangeRows, 1)
        .setDataValidation(listValidation(['per_device','total']));
    }
  }

  seedPatientFormulas_();
  onOpen();

  ScriptApp.getProjectTriggers().forEach(t=>{
    const h=t.getHandlerFunction();
    if(['sendDailyReminders','syncCalendarFromAppointments','updateInventoryStatus'].includes(h)) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendDailyReminders').timeBased().everyDays(1).atHour(Number(cfg('REMINDER_HOUR',9))||9).create();
  ScriptApp.newTrigger('syncCalendarFromAppointments').timeBased().everyHours(1).create();
  ScriptApp.newTrigger('updateInventoryStatus').timeBased().everyDays(1).create();

  toast('Kurulum tamam: Satır1=key, Satır2=TR, Veri=Satır3');
}

function seedPatientFormulas_(){
  const P=sh('Hastalar'); if(P.getLastRow()<3) P.insertRowsAfter(2,1);
  const Ph=headerIndex_(P);
  const I=sh('Görüşmeler'), Ih=headerIndex_(I);
  const S=sh('Satışlar'),   Sh=headerIndex_(S);
  const r=3;

  const p_id=Ph['patient_id']||Ph['tc'];
  const p_last=Ph['last_contact_at'], p_next=Ph['next_contact_at'], p_cnt=Ph['contact_count'], p_days=Ph['days_since_last'];
  const p_mode=Ph['limit_mode'], p_preset=Ph['limit_preset'], p_ld=Ph['limit_days'], p_le=Ph['limit_days_effective'], p_stale=Ph['stale_flag'];
  const p_lpp=Ph['last_payment_place'], p_lpm=Ph['last_payment_method'];
  const p_uses=Ph['uses_sgk'], p_lrep=Ph['last_sgk_report_received'], p_lsys=Ph['last_sgk_system_entered'];
  const p_slast=Ph['satisfaction_last'], p_savg=Ph['satisfaction_avg'], p_paid=Ph['paid_amount_last'];

  const I_when=colLetter_(Ih['when']), I_whoType=colLetter_(Ih['who_type']), I_whoId=colLetter_(Ih['who_id']), I_satisf=colLetter_(Ih['satisfaction']), I_next=colLetter_(Ih['next_action_at']);
  const S_pid=colLetter_(Sh['patient_id']), S_net=colLetter_(Sh['price_net']), S_place=colLetter_(Sh['payment_place']), S_method=colLetter_(Sh['payment_method']), S_uses=colLetter_(Sh['uses_sgk']), S_rep=colLetter_(Sh['sgk_report_received']), S_sys=colLetter_(Sh['sgk_system_entered']);
  const patientKeyLetter = colLetter_(p_id);

  P.getRange(r,p_last).setFormula(`=IFERROR(MAX(FILTER(Görüşmeler!${I_when}:${I_when},Görüşmeler!${I_whoType}:${I_whoType}="patient",Görüşmeler!${I_whoId}:${I_whoId}=${patientKeyLetter}${r})),"")`);
  P.getRange(r,p_next).setFormula(`=IFERROR(MIN(FILTER(Görüşmeler!${I_next}:${I_next},Görüşmeler!${I_whoType}:${I_whoType}="patient",Görüşmeler!${I_whoId}:${I_whoId}=${patientKeyLetter}${r},Görüşmeler!${I_next}:${I_next}>=TODAY())),"")`);
  P.getRange(r,p_cnt).setFormula(`=COUNTIFS(Görüşmeler!${I_whoType}:${I_whoType},"patient",Görüşmeler!${I_whoId}:${I_whoId},${patientKeyLetter}${r})`);
  P.getRange(r,p_days).setFormula(`=IF(${colLetter_(p_last)}${r}="","",TODAY()-${colLetter_(p_last)}${r})`);
  P.getRange(r,p_mode).setValue('preset');
  P.getRange(r,p_preset).setValue('3a');
  P.getRange(r,p_le).setFormula(`=IF(${colLetter_(p_mode)}${r}="custom",${colLetter_(p_ld)}${r},IF(${colLetter_(p_mode)}${r}="preset",IF(${colLetter_(p_preset)}${r}="7g",7,IF(${colLetter_(p_preset)}${r}="10g",10,IF(${colLetter_(p_preset)}${r}="1a",30,IF(${colLetter_(p_preset)}${r}="3a",90,VALUE(Config!B2)))))),IF(${colLetter_(p_mode)}${r}="auto",IF(${colLetter_(p_cnt)}${r}<3,30,IF(${colLetter_(p_cnt)}${r}<6,60,90)),VALUE(Config!B2))))`);
  P.getRange(r,p_stale).setFormula(`=IF(AND(${colLetter_(p_days)}${r}<>"",${colLetter_(p_le)}${r}<>"",${colLetter_(p_days)}${r}>=${colLetter_(p_le)}${r}),"STALE","OK")`);
  P.getRange(r,p_lpp).setFormula(`=IFERROR(INDIRECT("Satışlar!${S_place}&"&MAX(FILTER(ROW(Satışlar!${S_pid}:${S_pid}),Satışlar!${S_pid}:${S_pid}=${patientKeyLetter}${r}))),"")`.replace(/&/g,''));
  P.getRange(r,p_lpm).setFormula(`=IFERROR(INDIRECT("Satışlar!${S_method}&"&MAX(FILTER(ROW(Satışlar!${S_pid}:${S_pid}),Satışlar!${S_pid}:${S_pid}=${patientKeyLetter}${r}))),"")`.replace(/&/g,''));
  P.getRange(r,p_uses).setFormula(`=IFERROR(INDIRECT("Satışlar!${S_uses}&"&MAX(FILTER(ROW(Satışlar!${S_pid}:${S_pid}),Satışlar!${S_pid}:${S_pid}=${patientKeyLetter}${r}))),FALSE)`.replace(/&/g,''));
  P.getRange(r,p_lrep).setFormula(`=IFERROR(INDIRECT("Satışlar!${S_rep}&"&MAX(FILTER(ROW(Satışlar!${S_pid}:${S_pid}),Satışlar!${S_pid}:${S_pid}=${patientKeyLetter}${r}))),FALSE)`.replace(/&/g,''));
  P.getRange(r,p_lsys).setFormula(`=IFERROR(INDIRECT("Satışlar!${S_sys}&"&MAX(FILTER(ROW(Satışlar!${S_pid}:${S_pid}),Satışlar!${S_pid}:${S_pid}=${patientKeyLetter}${r}))),FALSE)`.replace(/&/g,''));
  P.getRange(r,p_slast).setFormula(`=IFERROR(LOOKUP(2,1/(Görüşmeler!${I_whoId}:${I_whoId}=${patientKeyLetter}${r}),Görüşmeler!${I_satisf}:${I_satisf}),"")`);
  P.getRange(r,p_savg).setFormula(`=IFERROR(AVERAGE(FILTER(Görüşmeler!${I_satisf}:${I_satisf},(Görüşmeler!${I_whoId}:${I_whoId}=${patientKeyLetter}${r})*(Görüşmeler!${I_when}:${I_when}>=TODAY()-90))),"")`);
  P.getRange(r,p_paid).setFormula(`=IFERROR(INDIRECT("Satışlar!${S_net}&"&MAX(FILTER(ROW(Satışlar!${S_pid}:${S_pid}),Satışlar!${S_pid}:${S_pid}=${patientKeyLetter}${r}))),"")`.replace(/&/g,''));
}

function setListValidation_(sheetName, key, values){
  if(!values || !values.length) return;
  const sheet = sh(sheetName);
  if(!sheet) return;
  const headers = headerIndex_(sheet);
  const col = headers[key];
  if(!col) return;
  const rows = Math.max(sheet.getMaxRows()-2,0);
  if(rows <= 0) return;
  const validation = SpreadsheetApp.newDataValidation().setAllowInvalid(false).requireValueInList(values);
  sheet.getRange(3,col,rows,1).setDataValidation(validation);
}
