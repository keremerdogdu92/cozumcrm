/* v3.0.0 */
function ensureCalendar_(){ const name='Hearing Center CRM'; const c=CalendarApp.getCalendarsByName(name); return c.length?c[0]:CalendarApp.createCalendar(name); }
function syncCalendarFromAppointments(){
  const cal=ensureCalendar_(); const s=sh('Randevular'); const idx=headerIndex_(s);
  for(let r=3;r<=s.getLastRow();r++){
    const st=s.getRange(r,idx['status']).getValue(); const sd=s.getRange(r,idx['start_datetime']).getValue(); const eid=s.getRange(r,idx['calendar_event_id']).getValue();
    if(st==='cancelled' || !sd || eid) continue;
    const title=s.getRange(r,idx['title']).getValue()||'Appointment';
    const ed=s.getRange(r,idx['end_datetime']).getValue()||sd; const loc=s.getRange(r,idx['location']).getValue()||'';
    const ev=cal.createEvent(title,new Date(sd),new Date(ed),{location:loc}); s.getRange(r,idx['calendar_event_id']).setValue(ev.getId());
  }
}
function sendDailyReminders(){
  const up=listUpcoming_(); const st=listStalePatients_();
  const text=[
    'Upcoming:\n'+(up.length>1?up.slice(1).map(r=>`• ${fmt(r[0])} - ${r[1]} (${r[2]}:${r[3]}) @ ${r[5]}`).join('\n'):'None'),
    '\nStale:\n'+(st.length>1?st.slice(1).map(r=>`• tc:${r[0]} ${r[1]} → ${r[2]}/${r[3]} days`).join('\n'):'None')
  ].join('\n');
  MailApp.sendEmail(Session.getActiveUser().getEmail(), 'CRM Daily Summary', text);
}
