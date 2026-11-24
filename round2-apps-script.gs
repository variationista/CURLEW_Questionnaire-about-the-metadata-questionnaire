function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('Round 2') || ss.insertSheet('Round 2');
    if (sh.getLastRow() === 0) sh.appendRow(['Submission ID','Submitted at (UTC)','Team member','ID','Original question / topic','Decision']);
    const safe = v => /^[=+\-@]/.test(String(v)) ? "'" + v : String(v);
    const rows = data.decisions.map(d => [safe(data.submissionId),new Date().toISOString(),safe(data.name),safe(d.id),safe(d.question),safe(d.decision)]);
    sh.getRange(sh.getLastRow()+1,1,rows.length,6).setValues(rows);
    return ContentService.createTextOutput('OK');
  } finally { lock.releaseLock(); }
}