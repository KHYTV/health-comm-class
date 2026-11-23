/**
 * 13주차 아세트아미노펜 체크 — 제출 저장용 Google Apps Script
 *
 * 설치 (한 번만, 약 5분)
 * 1. Google 드라이브에서 새 스프레드시트를 만든다 (이름 예: 아세트아미노펜_개인활동_제출).
 * 2. 메뉴 [확장 프로그램 → Apps Script]를 열고, 이 파일 내용을 통째로 붙여 넣고 저장한다.
 * 3. 왼쪽 [프로젝트 설정(톱니바퀴)] → 스크립트 속성 → 속성 추가:
 *      ADMIN_KEY = 관리자 비밀번호 (학생에게 알려 주지 않는 값)
 * 4. [배포 → 새 배포] → 유형: 웹 앱
 *      다음 사용자 인증 정보로 실행: 나
 *      액세스 권한이 있는 사용자: 모든 사용자
 *    → 배포 → 권한 승인 → 나오는 웹 앱 URL(…/exec)을 복사한다.
 * 5. 저장소의 w13/config.js 에서 APAP_ENDPOINT = "복사한 URL" 로 바꾼다.
 *
 * 저장 위치: 이 스프레드시트의 '제출' 시트, 사진은 내 드라이브의 '아세트아미노펜_개인활동_사진' 폴더(비공개).
 * 스크립트를 고친 뒤에는 [배포 → 배포 관리 → 수정 → 새 버전]으로 다시 배포해야 반영된다.
 */

const SHEET = '제출';
const FOLDER = '아세트아미노펜_개인활동_사진';
const HEAD = ['제출시각', '학번', '학과', '이름', '제품 수', '제품 요약', '함께 먹으면 하루 최대(mg)', '한 줄 평', '사진 링크', '원자료(JSON)'];

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) { sh = ss.insertSheet(SHEET); sh.appendRow(HEAD); sh.setFrozenRows(1); }
  return sh;
}
function folder_() {
  const it = DriveApp.getFoldersByName(FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER);
}
function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function clip_(s, n) { return String(s == null ? '' : s).slice(0, n); }

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const d = JSON.parse(e.postData.contents);
    if (d.action !== 'submit') return out_({ ok: false, error: 'bad_action' });
    if (!d.studentNo || !d.name || !Array.isArray(d.products) || d.products.length < 2) return out_({ ok: false, error: 'missing' });
    const when = new Date();
    const stamp = Utilities.formatDate(when, 'Asia/Seoul', 'yyyyMMdd_HHmmss');
    const links = [];
    const fol = folder_();
    (d.photos || []).slice(0, 6).forEach(function (p, i) {
      if (!p || !p.data) return;
      const bytes = Utilities.base64Decode(String(p.data).split(',').pop());
      if (bytes.length > 3 * 1024 * 1024) return;
      const f = fol.createFile(Utilities.newBlob(bytes, 'image/jpeg', clip_(d.studentNo, 20) + '_' + stamp + '_' + (i + 1) + '.jpg'));
      links.push(f.getUrl());
    });
    const rows = (d.result && d.result.rows) || [];
    const summary = rows.map(function (r) {
      return r.letter + ' ' + clip_(r.name, 40) + ': 1회 ' + Math.round(r.dose) + 'mg / 하루 ' + Math.round(r.day) + 'mg (' + r.pct + '%, ' + r.levelLabel + ')';
    }).join('\n');
    const raw = { products: d.products, result: d.result };
    sheet_().appendRow([
      Utilities.formatDate(when, 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss'),
      "'" + clip_(d.studentNo, 20), clip_(d.dept, 40), clip_(d.name, 30),
      d.products.length, summary, Math.round((d.result && d.result.sumDay) || 0),
      clip_(d.comment, 120), links.join('\n'), clip_(JSON.stringify(raw), 45000)
    ]);
    return out_({ ok: true, photos: links.length });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action === 'ping') return out_({ ok: true });
  if (p.action !== 'list') return out_({ ok: false, error: 'bad_action' });
  const key = PropertiesService.getScriptProperties().getProperty('ADMIN_KEY');
  if (!key || p.key !== key) return out_({ ok: false, error: 'unauthorized' });
  const values = sheet_().getDataRange().getDisplayValues();
  const head = values.shift() || HEAD;
  const rows = values.map(function (v) { const o = {}; head.forEach(function (h, i) { o[h] = v[i]; }); return o; });
  return out_({ ok: true, rows: rows });
}
