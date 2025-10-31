/* v3.0.0 */
function onOpen(){
  SpreadsheetApp.getUi().createMenu('CRM')
    .addItem('Refresh Dashboard','refreshDashboard')
    .addItem('Sync Calendar','syncCalendarFromAppointments')
    .addItem('Update Inventory','updateInventoryStatus')
    .addItem('Send Daily Reminders','sendDailyReminders')
    .addToUi();
}
