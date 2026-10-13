/**
 * 벨류에셋 상담 신청 → 구글 스프레드시트 저장용 Apps Script
 *
 * [설정 방법]
 * 1. 구글 스프레드시트를 새로 만들고, 메뉴에서 [확장 프로그램] → [Apps Script] 를 엽니다.
 * 2. 기본 코드를 모두 지우고 이 파일 내용을 붙여넣은 뒤 저장합니다.
 * 3. 오른쪽 위 [배포] → [새 배포] → 유형 선택(톱니바퀴)에서 [웹 앱] 을 고릅니다.
 *      - 다음 사용자로 실행: 나
 *      - 액세스 권한이 있는 사용자: 모든 사용자
 *    [배포] 를 누르고 권한 승인을 마치면 '웹 앱 URL'(…/exec)이 나옵니다.
 * 4. index.html 의 <script> 맨 위 SHEET_ENDPOINT = '' 따옴표 안에 그 URL을 붙여넣습니다.
 * 5. 코드를 수정했다면 [배포] → [배포 관리] → 연필 아이콘 → 버전 '새 버전' 으로 다시 배포해야 반영됩니다.
 *    (URL은 그대로 유지됩니다)
 *
 * 신청이 들어올 때 메일 알림을 받으려면 NOTIFY_EMAIL 에 주소를 넣으세요.
 */

var SHEET_NAME = '상담신청';
var NOTIFY_EMAIL = '';   // 예: 'manager@example.com'  (비워두면 메일 보내지 않음)

var COLUMNS = [
  ['접수일시',       function (p) { return new Date(); }],
  ['성함',           function (p) { return p.name; }],
  ['전화번호',       function (p) { return p.phone; }],
  ['관심 상품',      function (p) { return p.product; }],
  ['희망가전 1',     function (p) { return p.appliance1; }],
  ['희망가전 2',     function (p) { return p.appliance2; }],
  ['개인정보 동의',  function (p) { return p.agree === 'Y' ? '동의' : ''; }],
  ['신청 페이지',    function (p) { return p.page; }],
  ['유입 경로',      function (p) { return p.referrer; }],
  ['기기 정보',      function (p) { return p.userAgent; }]
];

function doPost(e) {
  var p = (e && e.parameter) || {};

  // 스팸 봇(숨김 칸 입력) · 필수값 누락은 저장하지 않음
  if (p.company) return json_({ result: 'ignored' });
  if (!p.name || !p.phone || !p.product || p.agree !== 'Y') return json_({ result: 'error', message: 'missing fields' });

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = getSheet_();
    var row = COLUMNS.map(function (c) { return clean_(c[1](p)); });
    sheet.appendRow(row);
  } finally {
    lock.releaseLock();
  }

  if (NOTIFY_EMAIL) {
    MailApp.sendEmail(NOTIFY_EMAIL, '[벨류에셋] 새 상담 신청 - ' + p.name,
      COLUMNS.slice(1, 7).map(function (c) { return c[0] + ': ' + (c[1](p) || '-'); }).join('\n'));
  }
  return json_({ result: 'success' });
}

// 브라우저로 웹 앱 URL을 열었을 때 동작 확인용
function doGet() {
  return ContentService.createTextOutput('OK - 벨류에셋 상담 신청 수신 중');
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(COLUMNS.map(function (c) { return c[0]; }));
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, COLUMNS.length).setFontWeight('bold').setBackground('#feebd2');
    sheet.getRange('C:C').setNumberFormat('@');   // 전화번호 앞자리 0 유지
  }
  return sheet;
}

// 값 정리: 길이 제한 + 수식 주입(=, +, -, @ 로 시작) 방지
function clean_(v) {
  if (v instanceof Date) return v;
  v = String(v == null ? '' : v).slice(0, 500);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
