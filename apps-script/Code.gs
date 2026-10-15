/**
 * Re-design 길잡이 · 연수실 결과표 저장용 Apps Script
 * 구글 시트에 붙여 넣고 [배포 > 새 배포 > 웹 앱]으로 배포하세요. (SETUP.md 참고)
 * 이름·이메일은 받지 않고, 선택값만 한 사람당 한 줄로 저장합니다.
 */
const SHEET_NAME = 'submissions';
const RESET_PASSWORD = 'reset'; // 연수실 초기화 비밀번호 (바꾸려면 여기만 고치고 새 버전으로 배포)
const HEADERS = ['id','시간','영역','상황','원인','먼저 손볼 요소','담은 처방','area','sit','cause','scenes','picked','top'];

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(SHEET_NAME);
  if (!s) { s = ss.insertSheet(SHEET_NAME); s.appendRow(HEADERS); s.setFrozenRows(1); }
  return s;
}
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function str_(v, n) { return String(v == null ? '' : v).slice(0, n); }
function arr_(v, n, m) { return (Array.isArray(v) ? v : []).slice(0, n).map(function (x) { return str_(x, m); }); }

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.action === 'reset') {
      if (String(d.password || '') !== RESET_PASSWORD) return json_({ ok: false, error: 'wrong password' });
      const sh = sheet_(), n = sh.getLastRow();
      if (n > 1) sh.deleteRows(2, n - 1);
      return json_({ ok: true, cleared: Math.max(n - 1, 0) });
    }
    const id = str_(d.id, 80);
    if (!id || !Number.isInteger(d.area) || !Number.isInteger(d.sit)) throw new Error('bad request');
    const L = d.labels || {};
    const picked = arr_(d.picked, 60, 40), top = arr_(d.top, 6, 4);
    const scenes = (Array.isArray(d.scenes) ? d.scenes : []).slice(0, 40).filter(Number.isInteger);
    const row = [
      id, new Date(), str_(L.area, 100), str_(L.sit, 200), str_(L.cause, 200),
      arr_(L.top, 6, 60).join(' / '), arr_(L.picked, 60, 200).join(' / '),
      d.area, d.sit, Number.isInteger(d.cause) ? d.cause : -1,
      JSON.stringify(scenes), JSON.stringify(picked), JSON.stringify(top)
    ];
    const s = sheet_();
    const last = s.getLastRow();
    let at = -1;
    if (last > 1) at = s.getRange(2, 1, last - 1, 1).getValues().map(function (r) { return r[0]; }).indexOf(id);
    if (at >= 0) s.getRange(at + 2, 1, 1, row.length).setValues([row]);
    else s.appendRow(row);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  try {
    const s = sheet_();
    const last = s.getLastRow();
    const rows = [];
    if (last > 1) {
      s.getRange(2, 1, last - 1, HEADERS.length).getValues().forEach(function (r) {
        try {
          rows.push({
            area: Number(r[7]), sit: Number(r[8]), cause: Number(r[9]),
            scenes: JSON.parse(r[10] || '[]'), picked: JSON.parse(r[11] || '[]'), top: JSON.parse(r[12] || '[]')
          });
        } catch (_) {}
      });
    }
    return json_({ ok: true, rows: rows });
  } catch (err) {
    return json_({ ok: false, error: String(err), rows: [] });
  }
}
