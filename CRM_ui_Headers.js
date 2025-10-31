/** /AppsScript/CRM/ui/Headers.gs
 * TR sheet adlarıyla sadeleştirilmiş başlıklar. v3.2.0
 */
const HEADERS = {
  Hastalar: {
    keys:['full_name','tc','address','phone','purchase_date','paid_amount_last','last_contact_at','next_contact_at','contact_count','days_since_last','limit_mode','limit_preset','limit_days','limit_days_effective','stale_flag','last_payment_place','last_payment_method','uses_sgk','last_sgk_report_received','last_sgk_system_entered','satisfaction_last','satisfaction_avg','notes'],
    tr:['Ad Soyad','TC Kimlik No','Adres','Telefon','Satın Alma Tarihi','Son Ödenen Tutar (Net)','Son Görüşme Tarihi','Sonraki Temas Tarihi','Görüşme Sayısı','Son Görüşmeden Beri Gün','Görüşme Limit Modu','Görüşme Limit Şablonu','Limit (Gün)','Etkin Limit (Gün)','Uyarı (Temas Gecikmiş)','Son Ödeme Yeri','Son Ödeme Yöntemi','SGK Kullanımı (Özet)','Son SGK Raporu Geldi','Son SGK Sisteme Girildi','Memnuniyet (Son)','Memnuniyet Ort. (90g)','Notlar']
  },
  "Görüşmeler": {
    keys:['log_id','who_type','who_id','method','when','patient_note','satisfaction','next_action_at'],
    tr:['Kayıt ID','Kayıt Türü','Kime Ait','Yöntem','Zaman','Açıklama','Memnuniyet (0–10)','Sonraki Aksiyon Tarihi']
  },
  "Satışlar": {
    keys:['sale_id','patient_id','device_group_id','serial_no','barcode','sale_date','price_gross','discount','price_net','installments','paid_amount','invoice_issued','payment_method','payment_place','uses_sgk','sgk_report_received','sgk_system_entered','notes'],
    tr:['Satış ID','Hasta ID (TC)','Cihaz Grup ID','Seri No','Barkod','Satış Tarihi','Brüt Tutar','İndirim','Net Tutar','Taksit Sayısı','Ödenen Tutar','Fatura Kesildi','Ödeme Yöntemi','Ödeme Yeri','SGK Kullanımı','SGK Raporu Geldi','SGK Sisteme Girildi','Notlar']
  },
  "Denemeler": {
    keys:['trial_id','full_name','phone','tc','address','started_at','last_contact_at','next_contact_at','status','notes'],
    tr:['Deneme ID','Ad Soyad','Telefon','TC','Adres','Başlangıç','Son Temas','Sonraki Temas','Durum','Notlar']
  },
  "Cihazlar": { // eski PatientDevices
    keys:['device_group_id','patient_id','side','serial_no','barcode','given_date','power_type','notes'],
    tr:['Cihaz Grup ID','Hasta ID (TC)','Taraf (Sağ/Sol/Tek)','Seri No','Barkod','Teslim Tarihi','Güç Türü (tek/pilli/şarjlı)','Notlar']
  },
  "Aksesuarlar": { // eski DeviceAccessories
    keys:['device_group_id','patient_id','type','serial_no','barcode','qty','notes'],
    tr:['Cihaz Grup ID','Hasta ID (TC)','Tür','Seri No','Barkod','Adet','Notlar']
  },
  "Randevular": {
    keys:['appt_id','who_type','who_id','title','start_datetime','end_datetime','location','calendar_event_id','status','is_first_visit','note'],
    tr:['Randevu ID','Kayıt Türü','Kime Ait','Başlık','Başlangıç','Bitiş','Konum','Takvim Etkinlik ID','Durum','İlk Ziyaret','Not']
  },
  "Referanslar": {
    keys:['ref_id','type','org_name','contact_name','phone','last_contact_at','next_contact_at','days_since_last','limit_days','stale_flag','notes'],
    tr:['Referans ID','Tür','Kurum Adı','Kişi Adı','Telefon','Son Temas','Sonraki Temas','Gün Farkı','Limit (Gün)','Uyarı','Notlar']
  },
  "Ürünler":    { keys:['name','brand','model','list_price','vat_rate'], tr:['Ürün Adı','Marka','Model','Liste Fiyatı','KDV Oranı'] },
  "Stok":       { keys:['serial_no','barcode','received_date','status'], tr:['Seri No','Barkod','Geliş Tarihi','Durum'] },
  "Config":     { keys:['key','value'], tr:['Anahtar','Değer'] },
  "Dashboard":  { keys:[], tr:[] }
};

function setHeaders_(sheetName){
  const s = ensure(sheetName);
  s.clear();
  s.getRange(1,1,s.getMaxRows(),s.getMaxColumns()).clearDataValidations();
  const keys = HEADERS[sheetName].keys, tr = HEADERS[sheetName].tr;
  if(keys.length){
    s.getRange(1,1,1,keys.length).setValues([keys]).setFontColor('#9AA0A6').setFontSize(9);
    s.getRange(2,1,1,keys.length).setValues([tr]).setFontColor('#202124').setFontWeight('bold').setFontSize(11);
    s.setFrozenRows(2);
  }
}
function forceHeadersTR(){ Object.keys(HEADERS).forEach(setHeaders_); toast('Headers reapplied'); }
