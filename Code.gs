const SHEET_NAME = 'Registrations';
const ADMIN_PASSWORD = 'CHANGE_THIS_ADMIN_PASSWORD';

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) sheet.appendRow(['Timestamp','Team Name','College','Team Leader','Email','Phone','Track','Team Members']);
}

function doPost(e) {
  try {
    setup();
    const data = JSON.parse(e.postData.contents);
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME).appendRow([new Date(), data.teamName || '', data.college || '', data.leader || '', data.email || '', data.phone || '', data.track || '', data.members || '']);
    return json({ok:true, message:'Registration saved'});
  } catch (err) { return json({ok:false, error:String(err)}); }
}

function doGet(e) {
  try {
    if ((e.parameter.action || '') !== 'list') return json({ok:true, message:'Hackathon registration API is running'});
    if (e.parameter.password !== ADMIN_PASSWORD) return json({ok:false, error:'Invalid admin password'});
    setup();
    const values = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME).getDataRange().getValues();
    const rows = values.slice(1).map(r => ({timestamp:formatDate(r[0]), teamName:r[1], college:r[2], leader:r[3], email:r[4], phone:r[5], track:r[6], members:r[7]}));
    const payload = JSON.stringify({ok:true, rows:rows});
    if (e.parameter.callback) return ContentService.createTextOutput(e.parameter.callback + '(' + payload + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
    return json({ok:true, rows:rows});
  } catch (err) { return json({ok:false, error:String(err)}); }
}

function formatDate(value) {
  if (value instanceof Date) return Utilities.formatDate(value, Session.getScriptTimeZone(), 'dd MMM yyyy, hh:mm a');
  return String(value || '');
}
function json(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
