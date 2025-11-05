/** /AppsScript/CRM/ui/Headers.gs
 * v3.7.0
 * CHANGELOG:
 * - Reordered quick-entry sheets for single-row trials with contact counts and device offers.
 * - Added price_offer/purpose columns plus Görüşmeler note ordering for secure trial logging.
 * - Preserved legacy analytics columns by appending them after the required quick-entry keys.
 */
const SHEET_HEADER_DEFS = Object.freeze({
  'Hastalar': {
    keys: [
      'patient_id','full_name','slug','phone','address','ref_key','created_at','updated_at',
      'tc','purchase_date','paid_amount_last','last_contact_at','next_contact_at','contact_count',
      'days_since_last','limit_mode','limit_preset','limit_days','limit_days_effective','stale_flag',
      'last_payment_place','last_payment_method','uses_sgk','last_sgk_report_received','last_sgk_system_entered',
      'satisfaction_last','satisfaction_avg','notes'
    ],
    labels: [
      'Hasta ID','Ad Soyad','Slug','Telefon','Adres','Referans Anahtarı','Oluşturulma','Güncelleme',
      'TC Kimlik No','Satın Alma Tarihi','Son Ödenen Tutar (Net)','Son Görüşme Tarihi','Sonraki Temas Tarihi','Görüşme Sayısı',
      'Son Görüşmeden Beri Gün','Görüşme Limit Modu','Görüşme Limit Şablonu','Limit (Gün)','Etkin Limit (Gün)','Uyarı (Temas Gecikmiş)',
      'Son Ödeme Yeri','Son Ödeme Yöntemi','SGK Kullanımı (Özet)','Son SGK Raporu Geldi','Son SGK Sisteme Girildi',
      'Memnuniyet (Son)','Memnuniyet Ort. (90g)','Notlar'
    ]
  },
  'Denemeler': {
    keys: [
      'trial_id','full_name','phone','tc','address','started_at','status','notes',
      'patient_id','slug','last_contact_at','next_contact_at','contact_count',
      'ref_key','created_at','updated_at','next_action_at'
    ],
    labels: [
      'Deneme ID','Ad Soyad','Telefon','TC','Adres','Başlangıç','Durum','Notlar',
      'Hasta ID','Slug','Son Temas','Sonraki Temas','Görüşme Sayısı',
      'Referans Anahtarı','Oluşturulma','Güncelleme','Sonraki Aksiyon'
    ]
  },
  'Cihazlar': {
    keys: [
      'device_group_id','patient_id','trial_id','model','side','qty','price_offer',
      'device_id','device_group_ref','created_at','purpose','serial_no','barcode','given_date',
      'power_type','notes'
    ],
    labels: [
      'Cihaz Grup ID','Hasta ID (TC)','Deneme ID','Model','Taraf (Sağ/Sol/Tek)','Adet','Teklif Fiyatı',
      'Cihaz ID','Grup Referansı','Oluşturulma','Kullanım Amacı','Seri No','Barkod','Teslim Tarihi',
      'Güç Türü (tek/pilli/şarjlı)','Notlar'
    ]
  },
  'Satışlar': {
    keys: [
      'sale_id','patient_id','trial_id','device_group_id','serial_no','barcode','sale_date','price_gross','discount','price_net','installments','paid_amount','invoice_issued','payment_method','payment_place','uses_sgk','sgk_report_received','sgk_system_entered','notes',
      'pricing_mode','item_model','quantity','total_net','payer','created_at'
    ],
    labels: [
      'Satış ID','Hasta ID (TC)','Deneme ID','Cihaz Grup ID','Seri No','Barkod','Satış Tarihi','Brüt Tutar','İndirim','Net Tutar','Taksit Sayısı','Ödenen Tutar','Fatura Kesildi','Ödeme Yöntemi','Ödeme Yeri','SGK Kullanımı','SGK Raporu Geldi','SGK Sisteme Girildi','Notlar',
      'Fiyatlama Modu','Ürün Modeli','Adet','Toplam Net','Ödeyen','Oluşturulma'
    ]
  },
  'Görüşmeler': {
    keys: [
      'log_id','who_type','who_id','method','when','note','type','next_action_at','created_at',
      'patient_note','satisfaction','payment_amount','payment_method'
    ],
    labels: [
      'Kayıt ID','Kayıt Türü','Kime Ait','Yöntem','Zaman','Not','Tür','Sonraki Aksiyon Tarihi','Oluşturulma',
      'Açıklama','Memnuniyet (0–10)','Ödeme Tutarı','Ödeme Yöntemi'
    ]
  },
  'Aksesuarlar': {
    keys: ['device_group_id','patient_id','type','serial_no','barcode','qty','notes'],
    labels: ['Cihaz Grup ID','Hasta ID','Tür','Seri No','Barkod','Adet','Notlar']
  },
  'Randevular': {
    keys: ['appt_id','who_type','who_id','title','start_datetime','end_datetime','location','calendar_event_id','status','is_first_visit','note'],
    labels: ['Randevu ID','Kayıt Türü','Kime Ait','Başlık','Başlangıç','Bitiş','Konum','Takvim Etkinlik ID','Durum','İlk Ziyaret','Not']
  },
  'Referanslar': {
    keys: ['ref_id','type','org_name','contact_name','phone','last_contact_at','next_contact_at','days_since_last','limit_days','stale_flag','notes'],
    labels: ['Referans ID','Tür','Kurum Adı','Kişi Adı','Telefon','Son Temas','Sonraki Temas','Gün Farkı','Limit (Gün)','Uyarı','Notlar']
  },
  'Ürünler': {
    keys: ['name','brand','model','list_price','vat_rate'],
    labels: ['Ürün Adı','Marka','Model','Liste Fiyatı','KDV Oranı']
  },
  'Stok': {
    keys: ['serial_no','barcode','received_date','status'],
    labels: ['Seri No','Barkod','Geliş Tarihi','Durum']
  },
  'Dashboard': { keys: [], labels: [] },
  'Config': { keys:['key','value'], labels:['Anahtar','Değer'] }
});

function ensureHeaderSet(sheetName){
  const def = SHEET_HEADER_DEFS[sheetName];
  if(!def) return;
  ensureSheetHeaders(sheetName, def.keys, def.labels);
}

function forceHeadersTR(){
  Object.keys(SHEET_HEADER_DEFS).forEach(ensureHeaderSet);
  toast('Headers ensured');
}
