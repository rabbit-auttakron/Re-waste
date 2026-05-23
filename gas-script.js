/**
 * ================================================================
 *  Re-waste · Google Apps Script Backend
 *  วิธีใช้:
 *  1. ไปที่ script.google.com → สร้างโปรเจกต์ใหม่
 *  2. วางโค้ดนี้ทับ Code.gs
 *  3. กด Deploy → New Deployment
 *  4. Type: Web app | Execute as: Me | Access: Anyone
 *  5. Copy URL แล้วใส่ใน GAS_URL ของ pet-donate.html / pet-admin.html
 * ================================================================
 */

const SHEET_NAME = 'Donations';

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['id','ts','name','empId','dept','destination','bottles','kgPET','kgCO2e']);
    sheet.setFrozenRows(1);
    sheet.getRange(1,1,1,9).setFontWeight('bold').setBackground('#14532d').setFontColor('#ffffff');
  }
  return sheet;
}

function doGet(e) {
  const action = (e.parameter && e.parameter.action) || 'get';

  // ── บันทึกรายการใหม่ ──────────────────────────────────────
  if (action === 'save') {
    try {
      const rec = JSON.parse(decodeURIComponent(e.parameter.data));
      const sheet = getSheet();
      sheet.appendRow([
        rec.id, rec.ts, rec.name, rec.empId, rec.dept,
        rec.destination, rec.bottles, rec.kgPET, rec.kgCO2
      ]);
      return ok({ saved: true });
    } catch(err) {
      return err_({ error: err.message });
    }
  }

  // ── ลบรายการ ──────────────────────────────────────────────
  if (action === 'delete') {
    try {
      const id = e.parameter.id;
      const sheet = getSheet();
      const vals = sheet.getDataRange().getValues();
      for (let i = 1; i < vals.length; i++) {
        if (String(vals[i][0]) === String(id)) { sheet.deleteRow(i + 1); break; }
      }
      return ok({ deleted: true });
    } catch(err) {
      return err_({ error: err.message });
    }
  }

  // ── ดึงข้อมูลทั้งหมด (default) ───────────────────────────
  try {
    const sheet = getSheet();
    const vals  = sheet.getDataRange().getValues();
    if (vals.length <= 1) return ok({ records: [] });
    const keys  = ['id','ts','name','empId','dept','destination','bottles','kgPET','kgCO2'];
    const records = vals.slice(1).map(row => {
      const r = {};
      keys.forEach((k, i) => r[k] = row[i]);
      r.bottles = Number(r.bottles); r.kgPET = Number(r.kgPET); r.kgCO2 = Number(r.kgCO2);
      return r;
    }).filter(r => r.id);          // ข้ามแถวว่าง
    return ok({ records });
  } catch(err) {
    return err_({ error: err.message });
  }
}

function ok(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
function err_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
