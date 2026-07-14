/**
 * ============================================================
 * © 2026 GEG 화성(깊이 e끌림). All rights reserved.
 *
 * 본 코드는 「저작권법」의 보호를 받는 저작물입니다.
 * - 복제권(제16조)·공중송신권(제18조)·배포권(제20조)은
 *   저작권자에게 있습니다.
 * - 정상 경로로 받은 이용자라도 코드의 무단 복제·재배포·
 *   재판매·리브랜딩은 허용되지 않습니다.
 * - 무단 이용 시 「저작권법」 제136조(5년 이하 징역 또는
 *   5천만 원 이하 벌금) 및 제125조(손해배상) 적용 대상이
 *   될 수 있습니다.
 * - 이용 문의: bacusiki777@gmail.com, for2102@jimj.kr
 * ============================================================
 */

// 빌드 서명
const _BUILD_SIG = 'GEGHS-DEEPE-2026';

// 출처 확인용 함수
function getBuildInfo() {
  return {
    sig: _BUILD_SIG,
    owner: 'GEG 화성(깊이 e끌림)',
    year: 2026
  };
}

// ---- 탭 이름 / 미션 / 서식 상수 ---------------------------------
const SHEET_ROSTER = '학생명단';   // A=번호 B=이름 + 각 학생의 최근 활동
const SHEET_LOG    = '활동 기록';   // 시각순으로 쌓이는 자세한 활동 로그
const SHEET_GUIDE  = '사용 설명';   // 선생님 안내

const MISSION_IDS = ['classify', 'prism', 'pyramid', 'pattern', 'net'];
const MISSION_LABELS = { classify: '분류', prism: '각기둥', pyramid: '각뿔', pattern: '규칙', net: '전개도' };

// '학생명단' 탭 열 구성: A번호 B이름 C최근활동 D최근미션 E분류 F각기둥 G각뿔 H규칙 I전개도 J완료수
const ROSTER_HEADER = ['번호', '이름', '최근 활동', '최근 미션', '분류', '각기둥', '각뿔', '규칙', '전개도', '완료 수'];
const ROSTER_FIRST_MISSION_COL = 5; // E열
const ROSTER_DONE_COL = 10;         // J열

// '활동 기록' 탭 열 구성
const LOG_HEADER = ['시각', '이름', '미션', '상태'];

const THEME_BG = '#fde68a'; // 데이터 탭 헤더 색(앱 도형 테마)

/**
 * 시트를 열 때 상단에 '도형 왕국' 메뉴를 만들어 줍니다.
 * 눈에 잘 띄도록 메뉴 이름과 항목에 그림(아이콘)을 넣었습니다.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📦 도형 왕국')
    .addItem('📊 학습 참여 통계 보기', 'openStatsSidebar')
    .addItem('📝 활동 기록 열기', 'openRecordTab')
    .addSeparator()
    .addItem('📖 사용 설명 탭 만들기', 'setupGuideSheet')
    .addItem('👥 학생 명단 탭 만들기', 'setupRosterSheet')
    .addToUi();
}

/**
 * 1. 웹 화면(index.html)을 브라우저에 표시해주는 함수
 * 웹 앱 링크로 학생들이 접속할 때 index.html의 내용으로 웹 브라우저 화면을 그려줍니다.
 * 스프레드시트에 묶인(bound) 스크립트라서, 사본을 만들면 그 사본 시트에 자동으로 연결됩니다.
 *
 * 학생 화면이 다른 주소에 배포된 경우에는, 요청에 담긴 항목(action)에 따라
 * 명단이나 기록 자료를 돌려주는 응답도 함께 처리합니다.
 */
function doGet(e) {
  var params = (e && e.parameter) || {};
  var action = params.action;

  // 학생 화면에서 이 시트의 자료(명단 읽기·기록 남기기)를 요청한 경우
  if (action) {
    var data;
    if (action === 'list') {
      data = { students: getStudentList() };
    } else if (action === 'record') {
      var result = recordMission(params.name, params.mission);
      data = { result: result };
    } else {
      data = { error: '알 수 없는 요청입니다.' };
    }

    var body = JSON.stringify(data);
    var cb = params.callback;
    // 학생 화면이 다른 주소에서도 자료를 안전하게 받을 수 있도록,
    // 요청에 담긴 callback 이름으로 감싸서 돌려줍니다.
    if (cb) {
      return ContentService.createTextOutput(cb + '(' + body + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(body)
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 기본: 앱 화면(index.html)을 그려 줍니다.
  var htmlOutput = HtmlService.createHtmlOutputFromFile('index');
  htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
  htmlOutput.setTitle('도형 왕국의 잃어버린 일지');
  htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return htmlOutput;
}

/**
 * 2. '학생명단' 탭에서 명단 읽어오기 함수
 * 웹 페이지가 시작될 때 학생 이름 목록을 가져와 전달합니다.
 * (google.script.run.getStudentList() 로 화면에서 호출되거나, doGet의 명단 요청에 사용됩니다.)
 * '학생명단' 탭은 A열=번호, B열=이름 으로 되어 있습니다.
 */
function getStudentList() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ROSTER);
    if (!sheet) return [];

    var values = sheet.getDataRange().getValues();
    var students = [];

    // 0번째 행(1행)은 헤더이므로 제외하고 1번째 행(2행)부터 이름(B열)을 모읍니다.
    for (var i = 1; i < values.length; i++) {
      var name = (values[i][1] !== undefined && values[i][1] !== '') ? values[i][1] : values[i][0];
      if (name !== undefined && name !== null && String(name).trim() !== '') {
        students.push(String(name).trim());
      }
    }
    return students;
  } catch (err) {
    return [];
  }
}

/**
 * 3. 학생이 미션을 마쳤을 때 기록하는 함수
 * (google.script.run.recordMission(name, missionId) 로 화면에서 호출되거나, doGet의 기록 요청에 사용됩니다.)
 * - '활동 기록' 탭에 [시각·이름·미션·상태] 를 한 줄씩 시간 순으로 쌓습니다(자세한 기록).
 * - '학생명단' 탭의 그 학생 줄에 최근 활동/미션과 미션별 완료를 갱신합니다(최신 상태).
 * @param {string} name - 학생 이름 (예: "학생1")
 * @param {string} missionId - 미션 고유 ID ('classify', 'prism', 'pyramid', 'pattern', 'net')
 */
function recordMission(name, missionId) {
  try {
    var label = MISSION_LABELS[missionId] || missionId;
    var now = new Date();

    // (1) 자세한 활동 로그 — 활동이 있을 때마다 한 줄씩 쌓습니다.
    var log = ensureLogSheet_();
    log.appendRow([now, name, label, '완료']);

    // (2) 학생 명단 — 그 학생 줄에 최근 활동을 갱신합니다.
    updateRosterLatest_(name, missionId, label, now);

    return 'success';
  } catch (err) {
    return 'error: ' + err.toString();
  }
}

/**
 * '학생명단'에서 해당 학생 줄을 찾아 최근 활동/미션과 미션별 완료, 완료 수를 갱신합니다.
 * 명단에 없는 이름이면 명단은 건드리지 않습니다(활동 로그에는 그대로 남습니다).
 */
function updateRosterLatest_(name, missionId, label, now) {
  ensureRosterHeader_();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var roster = ss.getSheetByName(SHEET_ROSTER);
  if (!roster) return;

  var values = roster.getDataRange().getValues();
  var rowIndex = -1;
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][1]).trim() === String(name).trim()) { rowIndex = i + 1; break; }
  }
  if (rowIndex === -1) return;

  roster.getRange(rowIndex, 3).setValue(now);   // C 최근 활동(시각)
  roster.getRange(rowIndex, 4).setValue(label); // D 최근 미션

  var mIdx = MISSION_IDS.indexOf(missionId);
  if (mIdx >= 0) roster.getRange(rowIndex, ROSTER_FIRST_MISSION_COL + mIdx).setValue('완료');

  // J 완료 수 재계산
  var missionCells = roster.getRange(rowIndex, ROSTER_FIRST_MISSION_COL, 1, MISSION_IDS.length).getValues()[0];
  var done = 0;
  for (var k = 0; k < missionCells.length; k++) {
    if (String(missionCells[k]).trim() === '완료') done++;
  }
  roster.getRange(rowIndex, ROSTER_DONE_COL).setValue(done);
}

/**
 * '학생 명단 탭 만들기' 메뉴에서 실행합니다.
 * 없으면 A=번호/B=이름 + 최근 활동 칸으로 만들고 학생1~학생30을 채웁니다.
 * 이미 있으면 이름은 그대로 두고 헤더(제목 줄)만 정리합니다.
 */
function setupRosterSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  cleanupLegacy_();

  var sheet = ss.getSheetByName(SHEET_ROSTER);
  if (sheet) {
    ensureRosterHeader_();
    ui.alert("'학생명단' 탭 헤더(제목 줄)를 정리했어요. 이름 칸(B열)을 우리 반 학생 이름으로 바꿔 사용하세요.");
    return;
  }

  sheet = ss.insertSheet(SHEET_ROSTER);
  var rows = [ROSTER_HEADER.slice()];
  for (var i = 1; i <= 30; i++) {
    var r = [i, '학생' + i];
    while (r.length < ROSTER_HEADER.length) r.push('');
    rows.push(r);
  }
  sheet.getRange(1, 1, rows.length, ROSTER_HEADER.length).setValues(rows);

  sheet.getRange(1, 1, 1, ROSTER_HEADER.length)
    .setFontWeight('bold').setBackground(THEME_BG).setHorizontalAlignment('center');
  sheet.setColumnWidth(1, 60);
  sheet.setColumnWidth(2, 140);
  sheet.setColumnWidth(3, 160);
  sheet.setColumnWidth(4, 90);
  sheet.getRange(1, 1, rows.length, ROSTER_HEADER.length).setBorder(true, true, true, true, true, true);
  sheet.setFrozenRows(1);

  ui.alert("'학생명단' 탭을 만들었어요. 이름 칸(B열)을 우리 반 학생 이름으로 바꿔 사용하세요.");
}

/**
 * '학생명단' 헤더(제목 줄)가 없거나 옛 형식(2칸)이면 새 헤더로 정리합니다. (이름 데이터는 보존)
 */
function ensureRosterHeader_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var roster = ss.getSheetByName(SHEET_ROSTER);
  if (!roster) return;
  var a1 = String(roster.getRange(1, 1).getValue()).trim();
  var c1 = roster.getRange(1, 3).getValue();
  if (a1 !== '번호' || c1 === '' || c1 === null) {
    roster.getRange(1, 1, 1, ROSTER_HEADER.length).setValues([ROSTER_HEADER]);
    roster.getRange(1, 1, 1, ROSTER_HEADER.length)
      .setFontWeight('bold').setBackground(THEME_BG).setHorizontalAlignment('center');
    roster.setFrozenRows(1);
  }
}

/**
 * '활동 기록' 탭이 없으면 만들어 돌려줍니다. 헤더(제목 줄)가 없으면 채웁니다.
 * (처음 만들 때 학생 명단 바로 옆에 자리잡음)
 */
function ensureLogSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_LOG);
  var created = false;
  if (!sheet) { sheet = ss.insertSheet(SHEET_LOG); created = true; }

  var firstCell = sheet.getRange(1, 1).getValue();
  if (firstCell === '' || firstCell === null) {
    sheet.getRange(1, 1, 1, LOG_HEADER.length).setValues([LOG_HEADER]);
    sheet.getRange(1, 1, 1, LOG_HEADER.length)
      .setFontWeight('bold').setBackground(THEME_BG).setHorizontalAlignment('center');
    sheet.setColumnWidth(1, 170);
    sheet.setColumnWidth(2, 140);
    sheet.setColumnWidth(3, 110);
    sheet.setColumnWidth(4, 80);
    sheet.setFrozenRows(1);
  }
  if (created) positionRecordNextToRoster_();
  return sheet;
}

/**
 * '활동 기록' 탭을 '학생명단' 탭 바로 오른쪽으로 옮깁니다. (명단 옆에서 바로 확인)
 */
function positionRecordNextToRoster_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var roster = ss.getSheetByName(SHEET_ROSTER);
  var rec = ss.getSheetByName(SHEET_LOG);
  if (!roster || !rec) return;
  ss.setActiveSheet(rec);
  ss.moveActiveSheet(roster.getIndex() + 1);
}

/**
 * '활동 기록 열기' 메뉴 — 학생 명단 옆의 '활동 기록' 탭을 열어 바로 확인합니다.
 */
function openRecordTab() {
  cleanupLegacy_();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureLogSheet_();
  positionRecordNextToRoster_();
  ss.setActiveSheet(ss.getSheetByName(SHEET_LOG));
}

/**
 * 옛 '학습기록' 탭(최신 상태만 있던 탭)을 정리(삭제)합니다.
 */
function cleanupLegacy_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var old = ss.getSheetByName('학습기록');
  if (old && ss.getSheets().length > 1) { try { ss.deleteSheet(old); } catch (e) {} }
}

/**
 * 학습 참여 정도를 계산해 사이드바로 돌려줍니다. (google.script.run 으로 호출)
 * '학생명단' 탭의 미션별 완료 칸을 읽어 집계합니다.
 */
function getParticipationStats() {
  ensureRosterHeader_();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var roster = ss.getSheetByName(SHEET_ROSTER);

  var students = [];
  var missionDone = [0, 0, 0, 0, 0];
  var allDoneCount = 0;
  var totalStudents = 0;

  if (roster) {
    var values = roster.getDataRange().getValues();
    for (var i = 1; i < values.length; i++) {
      var nm = values[i][1];
      if (nm === undefined || nm === null || String(nm).trim() === '') continue;
      totalStudents++;
      var cnt = 0;
      for (var m = 0; m < MISSION_IDS.length; m++) {
        var cell = values[i][ROSTER_FIRST_MISSION_COL - 1 + m];
        if (String(cell).trim() === '완료') { missionDone[m]++; cnt++; }
      }
      var isAll = (cnt === MISSION_IDS.length);
      if (isAll) allDoneCount++;
      students.push({ name: String(nm), count: cnt, allDone: isAll });
    }
  }

  students.sort(function (a, b) { return b.count - a.count; });

  return {
    totalStudents: totalStudents,
    recorded: students.filter(function (s) { return s.count > 0; }).length,
    allDone: allDoneCount,
    missions: MISSION_IDS.map(function (id, idx) { return { label: MISSION_LABELS[id], done: missionDone[idx] }; }),
    students: students
  };
}

/**
 * '학습 참여 통계 보기' 메뉴 — 오른쪽에 통계 사이드바를 엽니다.
 * 사이드바 화면은 아래 문자열에 그대로 담겨 있어 외부 연결 없이 동작합니다.
 */
function openStatsSidebar() {
  var html = HtmlService.createHtmlOutput(STATS_SIDEBAR_HTML_)
    .setTitle('학습 참여 통계');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * '사용 설명 탭 만들기' 메뉴에서 실행합니다.
 * 선생님을 위한 '사용 설명' 탭을 첫 번째 위치에 새로 만듭니다.
 */
function setupGuideSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  cleanupLegacy_();

  // 옛 안내 탭(있다면 모두)을 지우고 새로 만듭니다.
  var legacyNames = [SHEET_GUIDE, '사용법', '📘 사용법', 'guide', '가이드'];
  for (var g = 0; g < legacyNames.length; g++) {
    var s0 = ss.getSheetByName(legacyNames[g]);
    if (s0) { try { ss.deleteSheet(s0); } catch (e) {} }
  }
  var sheet = ss.insertSheet(SHEET_GUIDE, 0); // 첫 번째 탭으로

  var title = '도형 왕국 - 사용 설명';
  var intro = '이 탭은 선생님을 위한 안내입니다. 학생 화면에는 나타나지 않습니다.';

  var setupSteps = [
    '이 시트를 선생님 계정으로 사본을 만드세요. (파일 메뉴 > 사본 만들기)',
    '상단의 확장 프로그램 > Apps Script 를 열어 코드가 함께 복사되었는지 확인하세요.',
    'Apps Script 오른쪽 위 배포 > 새 배포 를 누르고 유형을 웹 앱으로 선택하세요.',
    '액세스 권한을 모든 사용자로 설정한 뒤 배포하고, 만들어진 웹 앱 주소(끝이 exec 인 주소)를 복사하세요.',
    "학생용 앱 시작 화면 아래의 '선생님 설정'을 열어 그 주소를 붙여넣고 연결하기를 누르세요.",
    '화면에 생긴 학생용 링크를 복사해 학생들에게 나눠 주세요. 학생은 링크만 누르면 바로 시작합니다.'
  ];
  var useSteps = [
    "'학생명단' 탭의 이름 칸(B열)을 우리 반 학생 이름으로 바꾸면, 앱의 이름 선택 목록에 그대로 나타납니다.",
    "학생이 미션을 마치면 '학생명단'의 그 학생 줄에 최근 활동과 완료가 표시됩니다.",
    "미션마다의 자세한 기록은 '활동 기록' 탭에 시간 순서대로 한 줄씩 쌓입니다.",
    "상단 '도형 왕국' 메뉴의 '학습 참여 통계 보기'로 반 전체 참여 정도를 한눈에 볼 수 있습니다.",
    '아직 시트를 연결하지 않았다면 앱은 체험 모드로 열리며, 결과는 학생 기기에만 임시로 저장됩니다.'
  ];

  var rows = [];
  var titleRow, introRow, sectionRows = [];
  rows.push([title]); titleRow = rows.length;
  rows.push([intro]); introRow = rows.length;
  rows.push(['']);
  rows.push(['처음 준비 (한 번만)']); sectionRows.push(rows.length);
  for (var i = 0; i < setupSteps.length; i++) rows.push([(i + 1) + '. ' + setupSteps[i]]);
  rows.push(['']);
  rows.push(['수업에서 사용하기']); sectionRows.push(rows.length);
  for (var j = 0; j < useSteps.length; j++) rows.push([(j + 1) + '. ' + useSteps[j]]);
  rows.push(['']);
  rows.push(['데이터나 설정을 바꿀 때는 앱 화면이 아니라 해당 시트 탭에서 직접 수정하세요. 탭 이름은 코드와 연결되어 있으므로 삭제하거나 변경하지 마세요.']);
  rows.push(['']);
  rows.push(['저작권 안내']); sectionRows.push(rows.length);
  rows.push(['이 자료의 저작권은 GEG 화성(깊이 e끌림)에 있으며 「저작권법」에 따라 보호받습니다.']);
  rows.push(['정상 경로로 받은 이용자라도 무단 복제·재배포·재판매·리브랜딩은 허용되지 않습니다.']);
  rows.push(['이용 문의: bacusiki777@gmail.com, for2102@jimj.kr']);

  var total = rows.length;
  sheet.getRange(1, 1, total, 1).setValues(rows);

  sheet.setColumnWidth(1, 720);
  var all = sheet.getRange(1, 1, total, 1);
  all.setWrap(true).setVerticalAlignment('top');
  all.setBorder(true, true, true, true, false, false);

  sheet.getRange(titleRow, 1).setFontSize(14).setFontWeight('bold').setBackground(THEME_BG);
  sheet.getRange(introRow, 1).setFontStyle('italic').setFontColor('#6b7280');
  for (var s = 0; s < sectionRows.length; s++) {
    sheet.getRange(sectionRows[s], 1).setFontWeight('bold').setBackground('#fef3c7');
  }
}

// 참여 통계 사이드바 화면 (외부 연결 없이 자체 완결)
const STATS_SIDEBAR_HTML_ = `<!DOCTYPE html>
<html>
<head><base target="_top">
<style>
  body{font-family:'Malgun Gothic','Apple SD Gothic Neo',sans-serif;margin:0;padding:12px;color:#3a2a12;background:#fffdf5}
  h2{font-size:16px;margin:0 0 8px}
  .card{background:#fff;border:2px solid #ecdfc2;border-radius:12px;padding:10px;margin-bottom:10px}
  .big{display:flex;gap:8px;border:0;padding:0;background:none}
  .stat{flex:1;text-align:center;background:#fef3c7;border-radius:10px;padding:8px}
  .stat .num{font-size:22px;font-weight:800;color:#b45309}
  .stat .lbl{font-size:11px;color:#8a6d3b;margin-top:2px}
  .h{font-weight:700;margin-bottom:6px;font-size:13px}
  .mrow{margin:7px 0}
  .mrow .top{display:flex;justify-content:space-between;font-size:12px;margin-bottom:3px}
  .bar{height:10px;background:#f1e9d6;border-radius:6px;overflow:hidden}
  .bar>span{display:block;height:100%;background:#f59e0b}
  table{width:100%;border-collapse:collapse;font-size:12px}
  th,td{border-bottom:1px solid #eee;padding:4px;text-align:left}
  th{color:#8a6d3b}
  .dots{letter-spacing:1px;white-space:nowrap}
  .star{color:#059669;font-weight:700}
  button{width:100%;padding:8px;border:0;border-radius:8px;background:#b45309;color:#fff;font-size:13px;cursor:pointer;margin-bottom:10px}
  .muted{color:#9a8c72;font-size:12px;text-align:center;padding:16px}
</style>
</head>
<body>
  <h2>학습 참여 통계</h2>
  <button onclick="load()">새로고침</button>
  <div id="content"><div class="muted">불러오는 중...</div></div>
<script>
  function load(){
    document.getElementById('content').innerHTML='<div class="muted">불러오는 중...</div>';
    google.script.run.withSuccessHandler(render).withFailureHandler(fail).getParticipationStats();
  }
  function fail(){ document.getElementById('content').innerHTML='<div class="muted">불러오지 못했어요. 다시 시도해 주세요.</div>'; }
  function pct(a,b){ return b? Math.round(a*100/b):0; }
  function esc(t){ return String(t).replace(/[&<>]/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;'}[c];}); }
  function render(d){
    var t=d.totalStudents||0, html='';
    html+='<div class="card big">';
    html+='<div class="stat"><div class="num">'+t+'</div><div class="lbl">전체 학생</div></div>';
    html+='<div class="stat"><div class="num">'+(d.recorded||0)+'</div><div class="lbl">참여</div></div>';
    html+='<div class="stat"><div class="num">'+(d.allDone||0)+'</div><div class="lbl">모두 완료</div></div>';
    html+='</div>';
    html+='<div class="card"><div class="h">미션별 완료 인원</div>';
    for(var i=0;i<d.missions.length;i++){
      var m=d.missions[i], p=pct(m.done,t);
      html+='<div class="mrow"><div class="top"><span>'+esc(m.label)+'</span><span>'+m.done+' / '+t+' ('+p+'%)</span></div>';
      html+='<div class="bar"><span style="width:'+p+'%"></span></div></div>';
    }
    html+='</div>';
    html+='<div class="card"><div class="h">학생별 진행</div>';
    if(!d.students||!d.students.length){ html+='<div class="muted">아직 참여 기록이 없어요.</div>'; }
    else{
      html+='<table><tr><th>이름</th><th>완료</th></tr>';
      for(var j=0;j<d.students.length;j++){
        var s=d.students[j], dots='';
        for(var k=0;k<5;k++){ dots+=(k<s.count?'●':'○'); }
        html+='<tr><td>'+esc(s.name)+(s.allDone?' <span class="star">★</span>':'')+'</td><td class="dots">'+dots+' '+s.count+'/5</td></tr>';
      }
      html+='</table>';
    }
    html+='</div>';
    document.getElementById('content').innerHTML=html;
  }
  load();
</script>
</body>
</html>`;
