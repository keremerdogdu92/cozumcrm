/* v3.5.0 */
function onOpen(){
  const menu=SpreadsheetApp.getUi().createMenu('CRM');
  menu.addItem('Trial Quick Entry','openTrialQuickEntry')
      .addItem('Hasta Kartı','openPatientCard')
      .addItem('Refresh Dashboard','refreshDashboard')
      .addItem('Sync Calendar','syncCalendarFromAppointments')
      .addItem('Update Inventory','updateInventoryStatus')
      .addItem('Send Daily Reminders','sendDailyReminders')
      .addToUi();
}
