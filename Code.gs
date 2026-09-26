/**
 * 수학교육론 임용 대비 웹앱 — 답안·채점 기록용 Google Apps Script
 *
 * 사용 방법 (README.md 참고)
 * 1) 기록을 남길 구글 시트를 만들고 [확장 프로그램] → [Apps Script]를 엽니다.
 * 2) 이 파일 내용을 Code.gs에 붙여 넣고 저장합니다.
 * 3) 함수 선택에서 setup 을 한 번 실행합니다(권한 승인).
 * 4) [배포] → [새 배포] → 유형: 웹 앱
 *      - 다음 사용자 인증 정보로 실행: 나
 *      - 액세스 권한이 있는 사용자: 모든 사용자
 *    → 웹 앱 URL(.../exec)을 복사해 웹앱의 config.js 의 GAS_URL 에 붙여 넣습니다.
 */

const SHEET_RECORD = '기록';
const SHEET_SUMMARY = '학생별 요약';
const HEADERS = ['제출 시각', '학번', '이름', '모드', '문항 ID', '구분', '단원/이론', '문항',
  '제출 답안', '점수', '만점', '득점률(%)', '포함 키워드', '누락 키워드', '소요 시간(초)', '제한 시간(초)', '시간 초과'];

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_RECORD) || ss.insertSheet(SHEET_RECORD);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#e2eeee');
    sh.setColumnWidth(8, 260);
    sh.setColumnWidth(9, 360);
  }
  let sm = ss.getSheetByName(SHEET_SUMMARY) || ss.insertSheet(SHEET_SUMMARY);
  sm.clear();
  sm.getRange('A1').setFormula(
    "=QUERY('" + SHEET_RECORD + "'!A:Q," +
    "\"select B, C, D, count(E), avg(L), sum(O) where B is not null group by B, C, D " +
    "label count(E) '제출 수', avg(L) '평균 득점률(%)', sum(O) '총 소요(초)'\", 1)");
  sm.getRange('H1').setValue('※ 이 시트는 자동 집계입니다. 직접 수정하지 마세요.');
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(SHEET_RECORD);
    if (!sh) { setup(); sh = ss.getSheetByName(SHEET_RECORD); }
    const ts = d.ts ? new Date(d.ts) : new Date();
    const row = [
      ts, String(d.sid || ''), String(d.name || ''), d.mode || '', d.qid || '', d.group || '', d.unit || '',
      d.question || '', d.answer || '', Number(d.score) || 0, Number(d.max) || 0, Number(d.rate) || 0,
      d.hits || '', d.miss || '', Number(d.secs) || 0, d.limit === '' ? '' : Number(d.limit) || '', d.over || ''
    ];
    // 답안이 '=' 등으로 시작하면 수식으로 해석되지 않도록 처리
    for (let i = 1; i < row.length; i++) {
      if (typeof row[i] === 'string' && /^[=+\-@]/.test(row[i])) row[i] = "'" + row[i];
    }
    const r = sh.getLastRow() + 1;
    // 학번의 앞자리 0이 사라지지 않도록 학번 칸을 먼저 텍스트 서식으로 지정
    sh.getRange(r, 2).setNumberFormat('@');
    sh.getRange(r, 1, 1, row.length).setValues([row]);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  return json_({ ok: true, message: '수학교육론 기록 서버가 동작 중입니다.' });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
