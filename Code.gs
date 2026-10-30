const SHEET_NAME = 'Registrations';
const HEADERS = ['เวลา','เลขทะเบียน','ชื่อ-นามสกุล','เบอร์โทร','หน่วยงาน','ผู้ติดตาม','หมายเหตุ'];

function doGet() {
  return ContentService.createTextOutput('MESOOK registration API is running');
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sh.setFrozenRows(1);
  } else if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function doPost(e) {
  try {
    const data = JSON.parse((e.postData && e.postData.contents) || '{}');
    const fullName = String(data.fullName || '').trim();
    if (!fullName) return json_({ok:false, message:'กรุณากรอกชื่อ-นามสกุล'});

    const sh = getSheet_();
    const lock = LockService.getScriptLock();
    lock.waitLock(5000);
    try {
      const row = sh.getLastRow() + 1;
      const regNo = 'MS-' + Utilities.formatString('%05d', Math.max(1, row - 1));
      sh.getRange(row, 1, 1, 7).setValues([[
        new Date(), regNo, fullName,
        String(data.phone || '').trim(),
        String(data.organization || '').trim(),
        Number(data.guests || 0),
        String(data.note || '').trim()
      ]]);
      SpreadsheetApp.flush();
      return json_({ok:true, regNo:regNo, name:fullName});
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return json_({ok:false, message:err && err.message ? err.message : 'บันทึกไม่สำเร็จ'});
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
