/**
 * Code.gs
 * --------
 * Backend ของแบบสำรวจสภาพอาคาร — เขียนด้วย Google Apps Script แทน Python
 * รันบนเซิร์ฟเวอร์ของ Google เอง ฟรี ไม่ต้องมี Hosting แยก ไม่ต้องมี Service Account Key
 *
 * วิธีติดตั้ง (ทำครั้งเดียว):
 *   1. ไปที่ script.google.com → New project
 *   2. ลบโค้ดเปล่าเริ่มต้นออก แล้ววางโค้ดทั้งไฟล์นี้แทน
 *   3. เมนูซ้าย "Project Settings" (รูปเฟือง) → Script Properties → Add script property
 *      เพิ่ม 3 ค่านี้:
 *        DRIVE_FOLDER_ID = <Folder ID ของโฟลเดอร์แม่บน Drive ที่เตรียมไว้>
 *        SHEET_ID        = <Sheet ID ของ Google Sheet ที่เตรียมไว้สำหรับเก็บคำตอบ>
 *        SHEET_TAB       = Responses   (หรือชื่อแท็บอื่นที่ต้องการ)
 *   4. ปุ่ม Deploy (มุมขวาบน) → New deployment → เลือกประเภท "Web app"
 *        - Execute as: Me
 *        - Who has access: Anyone
 *      กด Deploy → คัดลอก "Web app URL" ที่ได้มา (ลงท้ายด้วย /exec)
 *   5. เอา URL นั้นไปวางใน templates/index.html ตรงตัวแปร APPS_SCRIPT_URL ด้านบนของ <script>
 *
 * หมายเหตุสำคัญ:
 *   - ไม่มีการย่อ/บีบอัดรูปก่อนอัปโหลด (Apps Script ไม่มีเครื่องมือแบบ Pillow ในตัว)
 *     รูปจะขึ้น Drive ตามขนาดจริงที่ถ่ายมา — เผื่อพื้นที่ Drive ให้พอ
 *   - ตอนนี้ตัดการสร้าง PDF ออกก่อน (เก็บข้อมูลลง Sheet + Drive อย่างเดียว) เพื่อลดความซับซ้อน
 *     ตอนดีบัก — เพิ่มกลับมาได้ทีหลังถ้าต้องการสรุปเป็นรายงาน PDF อีกครั้ง
 *   - ใน Google Sheet เดียวกันจะมี 2 แท็บ: "BranchIndex" (1 แถวต่อสาขา ใช้เช็คซ้ำเร็ว)
 *     กับ "Responses" (คำตอบทุกข้อ) — ทั้งสองแท็บถูกสร้างให้อัตโนมัติตอนรันครั้งแรก ไม่ต้องสร้างเอง
 *   - ระบบล็อกคิว (LockService) ครอบเฉพาะขั้นตอน "เช็คซ้ำ+จองสิทธิ์สาขา" ที่ใช้เวลาสั้นมาก
 *     ไม่ได้ล็อกทั้งขั้นตอนอัปโหลด Drive/เขียน Sheet เพื่อลดคอขวดตอนมีคนส่งพร้อมกันหลายคน
 */

function getProp_(key) {
  return PropertiesService.getScriptProperties().getProperty(key);
}

const HEADER_ROW = [
  'ลำดับคิว', 'เวลาที่ส่ง', 'Submission ID', 'AFS ID', 'ID สาขา', 'ชื่อสาขา', 'พื้นที่',
  'โซน/หัวข้อ', 'ข้อที่', 'คำถาม', 'ประเภทคำตอบ', 'สถานะ', 'จำนวน', 'รายละเอียด', 'มีรูปแนบ',
];

// ============================================================
// Entry point — รับ POST จากหน้าเว็บ
// ============================================================
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;

    if (action === 'checkBranch') {
      return jsonResponse_(checkBranch_(payload));
    }
    if (action === 'submit') {
      return jsonResponse_(submitSurvey_(payload));
    }
    return jsonResponse_({ status: 'error', message: 'ไม่รู้จัก action: ' + action });
  } catch (err) {
    return jsonResponse_({ status: 'error', message: String(err) });
  }
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function norm_(v) {
  return (v || '').toString().trim().toLowerCase();
}

function sanitizeFilename_(v) {
  return (v || '').toString().replace(/[\\\/:*?"<>|]/g, '').trim() || 'branch';
}

// ============================================================
// Google Sheet — 2 แท็บ:
//   "BranchIndex" = 1 แถวต่อ 1 สาขา ใช้เช็คซ้ำเร็ว (ไม่โตตามจำนวนคำตอบ)
//   "Responses"   = คำตอบ Text ทุกข้อ (1 แถวต่อ 1 คำตอบ) ไว้เปิดดู/วิเคราะห์
// ============================================================
function getResponsesSheet_() {
  const ss = SpreadsheetApp.openById(getProp_('SHEET_ID'));
  const tabName = getProp_('SHEET_TAB') || 'Responses';
  let sheet = ss.getSheetByName(tabName);
  if (!sheet) sheet = ss.insertSheet(tabName);
  if (sheet.getLastRow() === 0) sheet.appendRow(HEADER_ROW);
  return sheet;
}

const INDEX_TAB_ = 'BranchIndex';
const INDEX_HEADER_ = ['AFS ID', 'ID สาขา', 'ชื่อสาขา', 'Submission ID', 'เวลาที่ส่ง'];

function getIndexSheet_() {
  const ss = SpreadsheetApp.openById(getProp_('SHEET_ID'));
  let sheet = ss.getSheetByName(INDEX_TAB_);
  if (!sheet) sheet = ss.insertSheet(INDEX_TAB_);
  if (sheet.getLastRow() === 0) sheet.appendRow(INDEX_HEADER_);
  return sheet;
}

/** เช็คว่าสาขานี้ (ID + ชื่อ + AFS ตรงกันทั้ง 3 ช่อง) เคยส่งแบบสำรวจแล้วหรือยัง — เช็คจาก BranchIndex
 *  (แท็บเล็ก 1 แถวต่อสาขา ไม่ใช่แท็บ Responses ที่โตเรื่อยๆ ตามจำนวนคำตอบ — เร็วกว่ามากเมื่อข้อมูลเยอะ) */
function findExistingBranch_(id, name, afs) {
  const sheet = getIndexSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;
  const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues(); // AFS ID, ID สาขา, ชื่อสาขา
  const targetAfs = norm_(afs), targetId = norm_(id), targetName = norm_(name);
  for (let i = 0; i < data.length; i++) {
    if (norm_(data[i][0]) === targetAfs && norm_(data[i][1]) === targetId && norm_(data[i][2]) === targetName) {
      return true;
    }
  }
  return false;
}

/** "จอง" สิทธิ์สาขานี้ทันทีด้วยการเพิ่มแถวใน BranchIndex — เรียกใช้ภายใต้ lock เท่านั้น
 *  เพื่อกันสองคนเช็คไม่เจอซ้ำพร้อมกันแล้วผ่านทั้งคู่ (race condition) */
function claimBranch_(id, name, afs, submissionId) {
  const sheet = getIndexSheet_();
  const now = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm:ss');
  sheet.appendRow([afs || '', id || '', name || '', submissionId, now]);
}

function checkBranch_(payload) {
  return { duplicate: findExistingBranch_(payload.id, payload.name, payload.afs) };
}

// ============================================================
// รูปภาพ — ดึงรายการรูปทั้งหมดจากคำตอบ พร้อมชื่อไฟล์
// ============================================================
function decodeDataUrl_(dataUrl) {
  const parts = (dataUrl || '').split(',');
  if (parts.length < 2) return null;
  const mimeMatch = parts[0].match(/data:(.*?);base64/);
  const mime = mimeMatch ? mimeMatch[1] : 'image/jpeg';
  try {
    return Utilities.newBlob(Utilities.base64Decode(parts[1]), mime);
  } catch (err) {
    return null;
  }
}

function photoEntries_(answers) {
  const entries = [];
  answers.forEach(function (ans) {
    const photos = ans.photos || [];
    if (photos.length && typeof photos[0] === 'object') {
      photos.forEach(function (p) {
        if (p && p.data) entries.push({ name: p.name || p.label || 'รูปภาพ', data: p.data });
      });
    } else {
      photos.forEach(function (data, i) {
        if (data) entries.push({ name: (ans.label || 'รูปภาพ') + ' (' + (i + 1) + ')', data: data });
      });
    }
  });
  return entries;
}

// ============================================================
// Google Drive — โฟลเดอร์ต่อสาขา (ตั้งชื่อตาม AFS ID) + อัปโหลดรูป/PDF เข้าไป
// ============================================================
function getOrCreateBranchFolder_(afsId) {
  const root = DriveApp.getFolderById(getProp_('DRIVE_FOLDER_ID'));
  const existing = root.getFoldersByName(afsId);
  if (existing.hasNext()) return existing.next();
  return root.createFolder(afsId);
}

// ============================================================
// mapping สถานะสำหรับเขียนลง Google Sheet
// ============================================================
const STATUS_TEXT_ = { ok: 'ปกติ', bad: 'ไม่ปกติ', none: 'ไม่มีอุปกรณ์' };

// ============================================================
// Submit — เช็คซ้ำ+จองสิทธิ์ (ใช้ lock สั้นๆ) แล้วค่อยอัปโหลด Drive + เขียน Sheet
// (ไม่ล็อกยาวทั้งฟังก์ชัน กันคอขวดตอนมีคนส่งพร้อมกันหลายคน)
// ============================================================
function submitSurvey_(payload) {
  const branch = payload.branch || {};
  const answers = payload.answers || [];
  const submissionId = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyyMMdd_HHmmss')
    + '_' + Math.random().toString(36).substring(2, 8);

  // ---------- ส่วนที่ต้องล็อกคิว: เช็คซ้ำ + จองสิทธิ์สาขาทันที (เร็ว ไม่กี่ร้อย ms) ----------
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  let isDuplicate = false;
  try {
    isDuplicate = findExistingBranch_(branch.id, branch.name, branch.afs);
    if (!isDuplicate) {
      claimBranch_(branch.id, branch.name, branch.afs, submissionId);
    }
  } finally {
    lock.releaseLock();
  }

  if (isDuplicate) {
    return {
      status: 'duplicate',
      message: 'สาขาของท่านได้ทำแบบสำรวจเรียบร้อย รบกวนตรวจสอบข้อมูลของท่านหากท่านใส่ข้อมูลผิด',
    };
  }

  // ---------- ส่วนที่เหลือ: ทำได้โดยไม่ต้องรอคิวคนอื่น (ไม่ชนกันเพราะจองสิทธิ์ไปแล้วข้างบน) ----------
  const result = {
    status: 'ok', submission_id: submissionId,
    drive_uploaded: false, drive_photos_uploaded: 0, drive_photos_total: 0,
    sheets_uploaded: false,
  };

  // 1. หา/สร้างโฟลเดอร์ Drive ของสาขา + อัปโหลดรูปทุกใบ
  try {
    const branchFolder = getOrCreateBranchFolder_(sanitizeFilename_(branch.afs || '0000'));
    const entries = photoEntries_(answers);
    result.drive_photos_total = entries.length;
    let uploaded = 0;
    entries.forEach(function (entry) {
      const blob = decodeDataUrl_(entry.data);
      if (!blob) return;
      const name = /\.(jpg|jpeg|png)$/i.test(entry.name) ? entry.name : entry.name + '.jpg';
      blob.setName(name);
      branchFolder.createFile(blob);
      uploaded++;
    });
    result.drive_photos_uploaded = uploaded;
    result.drive_uploaded = true;
  } catch (err) {
    result.drive_error = 'อัปโหลด Google Drive ไม่สำเร็จ: ' + String(err);
  }

  // 2. เขียนคำตอบ Text ลง Google Sheet แท็บ Responses (ทีละแถวต่อ 1 คำตอบ)
  try {
    const sheet = getResponsesSheet_();
    const now = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm:ss');
    const startRow = sheet.getLastRow() + 1;
    const rows = answers.map(function (ans, idx) {
      const hasPhoto = photoEntries_([ans]).length > 0 ? 'มี' : '';
      return [
        startRow - 1 + idx, now, submissionId,
        branch.afs || '', branch.id || '', branch.name || '', branch.area || '',
        ans.zone || '', ans.question || '', ans.label || '', ans.kind || '',
        STATUS_TEXT_[ans.status] || '',
        (ans.quantity !== undefined && ans.quantity !== null) ? ans.quantity : '',
        (ans.detail || '').toString().trim(), hasPhoto,
      ];
    });
    if (rows.length) {
      sheet.getRange(startRow, 1, rows.length, HEADER_ROW.length).setValues(rows);
    }
    result.sheets_uploaded = true;
    result.sheets_rows = rows.length;
  } catch (err) {
    result.sheets_error = 'เขียน Google Sheet ไม่สำเร็จ: ' + String(err);
  }

  return result;
}
