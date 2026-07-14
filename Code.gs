/**
 * ============================================================
 * © 2026 GEG화성(깊이 e끌림). All rights reserved.
 *
 * 본 코드는 「저작권법」의 보호받는 저작물입니다.
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
  return { sig: _BUILD_SIG, owner: 'GEG화성(깊이 e끌림)', year: 2026 };
}

/* ============================================================
 * 식물의 구조와 기능 — 진행 저장/조회 API (구글 시트 연동)
 * 앱 화면(index.html)은 GitHub Pages에 올라가 있고,
 * 이 스크립트는 선생님 각자의 시트에 붙여 웹 앱으로 배포합니다.
 * 배포 후 나오는 .../exec 주소를 앱 ⚙️ 설정에 넣으면 연결됩니다.
 * ============================================================ */

const 앱_이름 = '식물의 비밀 공장';
const 탭_명단 = '학생 명단';
const 탭_현황 = '학습 현황';
const 탭_설명 = '사용 설명';
const 학생수 = 30;

/* ---- 시트가 열릴 때 앱 이름으로 관리 메뉴를 단다 ---- */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🌱 ' + 앱_이름)
    .addItem('학생 명단·학습 현황 만들기', 'setupSheets')
    .addItem('사용 설명 탭 만들기', 'setupGuideSheet')
    .addItem('학습 대시보드 열기', 'openDashboard')
    .addToUi();
}

/* ============================================================
 * 웹 앱 진입점 — JSONP(GET)로만 통신한다.
 * ?callback=..&action=roster|load|save|ping&student=..&data=..
 * ============================================================ */
function doGet(e) {
  const p = (e && e.parameter) || {};
  const cb = p.callback || '';
  // 앱이 아니라 사람이 주소를 직접 열었을 때 안내
  if (!cb) {
    return ContentService
      .createTextOutput('이 주소는 앱의 데이터 연결용입니다. 앱 ⚙️ 설정에 붙여 넣어 사용하세요.')
      .setMimeType(ContentService.MimeType.TEXT);
  }
  let out;
  try {
    const action = p.action || '';
    if (action === 'roster') {
      out = { ok: true, roster: 명단_읽기_() };
    } else if (action === 'load') {
      out = { ok: true, data: 진행_읽기_(p.student) };
    } else if (action === 'save') {
      진행_저장_(p.student, p.data);
      out = { ok: true };
    } else if (action === 'ping') {
      out = { ok: true, info: getBuildInfo(), app: 앱_이름 };
    } else {
      out = { ok: false, error: '알 수 없는 요청' };
    }
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err) };
  }
  return ContentService
    .createTextOutput(cb + '(' + JSON.stringify(out) + ')')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

/* ---- 학생 명단 읽기 → [{no, name}] ---- */
function 명단_읽기_() {
  const sh = 명단시트_();
  const last = Math.max(2, sh.getLastRow());
  const vals = sh.getRange(2, 1, last - 1, 2).getValues();
  const list = [];
  for (let i = 0; i < vals.length; i++) {
    const no = vals[i][0], name = vals[i][1];
    if (no === '' || no === null) continue;
    list.push({ no: Number(no), name: String(name || ('학생' + no)) });
  }
  return list;
}

/* ---- 한 학생의 진행 JSON 읽기 ---- */
function 진행_읽기_(student) {
  const no = Number(student);
  if (!no) return null;
  const sh = 명단시트_();
  const row = no + 1; // 1행은 머리글
  const cell = sh.getRange(row, 3).getValue(); // C열 = 진행(JSON)
  if (!cell) return null;
  try { return JSON.parse(cell); } catch (e) { return null; }
}

/* ---- 한 학생의 진행 저장 + 학습 현황 갱신 ---- */
function 진행_저장_(student, dataStr) {
  const no = Number(student);
  if (!no) throw new Error('학생 번호가 없습니다.');
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const payload = (typeof dataStr === 'string') ? JSON.parse(dataStr) : (dataStr || {});
    const sh = 명단시트_();
    const row = no + 1;
    const now = new Date();
    // 원본 JSON은 명단 탭 C열, 저장 시각은 D열
    sh.getRange(row, 3).setValue(JSON.stringify(payload));
    sh.getRange(row, 4).setValue(now);
    // 사람이 읽는 요약은 학습 현황 탭
    const name = String(sh.getRange(row, 2).getValue() || ('학생' + no));
    현황_갱신_(no, name, payload && payload.summary, now);
  } finally {
    lock.releaseLock();
  }
}

/* ---- 학습 현황 탭의 해당 학생 줄을 갱신 ---- */
function 현황_갱신_(no, name, summary, when) {
  const sh = 현황시트_();
  const s = summary || {};
  const row = no + 1;
  sh.getRange(row, 1, 1, 6).setValues([[
    no,
    name,
    (s.sprouts != null ? s.sprouts : '') + (s.sproutsTotal != null ? ' / ' + s.sproutsTotal : ''),
    (s.stars != null ? s.stars : '') + (s.starsTotal != null ? ' / ' + s.starsTotal : ''),
    s.allComplete ? '완료' : '진행 중',
    when
  ]]);
}

/* ============================================================
 * 시트 준비 함수들
 * ============================================================ */

/* ---- 학생 명단 + 학습 현황 탭 만들기(없으면 생성, 있으면 유지) ---- */
function setupSheets() {
  명단시트_();
  현황시트_();
  SpreadsheetApp.getActive().toast('학생 명단·학습 현황 탭이 준비되었습니다.', 앱_이름, 5);
}

function 명단시트_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(탭_명단);
  if (!sh) {
    sh = ss.insertSheet(탭_명단, 0);
    const head = [['번호', '이름', '진행(자동)', '수정 시각(자동)']];
    sh.getRange(1, 1, 1, 4).setValues(head).setFontWeight('bold').setBackground('#e8f0e2');
    const rows = [];
    for (let i = 1; i <= 학생수; i++) rows.push([i, '학생' + i, '', '']);
    sh.getRange(2, 1, 학생수, 4).setValues(rows);
    sh.setColumnWidth(1, 60); sh.setColumnWidth(2, 120);
    sh.setColumnWidth(3, 320); sh.setColumnWidth(4, 160);
    sh.setFrozenRows(1);
  }
  return sh;
}

function 현황시트_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(탭_현황);
  if (!sh) {
    sh = ss.insertSheet(탭_현황);
    const head = [['번호', '이름', '모은 새싹', '획득 별', '전체 완료', '수정 시각']];
    sh.getRange(1, 1, 1, 6).setValues(head).setFontWeight('bold').setBackground('#e8f0e2');
    const rows = [];
    for (let i = 1; i <= 학생수; i++) rows.push([i, '학생' + i, '', '', '', '']);
    sh.getRange(2, 1, 학생수, 6).setValues(rows);
    sh.setColumnWidths(1, 6, 110);
    sh.setColumnWidth(2, 120);
    sh.setFrozenRows(1);
  }
  return sh;
}

/* ============================================================
 * 사용 설명 탭 — 항상 새로(첫 번째 위치) 만든다
 * 안내문 항목 번호는 자동으로 매긴다. 이모티콘은 쓰지 않는다.
 * ============================================================ */
function setupGuideSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // 기존 사용 설명/사용법 탭은 모두 지우고 새로 만든다
  ['사용 설명', '사용법', '사용 방법'].forEach(function (nm) {
    const old = ss.getSheetByName(nm);
    if (old) ss.deleteSheet(old);
  });
  const sh = ss.insertSheet(탭_설명, 0);

  const sections = [
    ['준비하기', [
      '이 시트를 웹 앱으로 배포하면 .../exec 로 끝나는 주소가 나옵니다.',
      '앱 로그인 화면에서 시작하기 버튼 아래의 선생님 설정을 눌러 그 주소를 붙여 넣고 연결하기를 누르면 이 시트와 연결됩니다.',
      '연결 뒤 설정에 나오는 학생용 주소를 복사해 학생들에게 나눠 주세요.'
    ]],
    ['학생 명단 탭', [
      '학생 명단 탭의 이름 칸을 학급 실제 이름으로 바꾸면 앱의 학생 선택 목록에 그대로 나옵니다.',
      '진행 칸과 수정 시각 칸은 앱이 자동으로 채웁니다. 사람이 손대지 않습니다.',
      '학생은 앱에서 자기 이름을 고른 뒤 활동을 이어서 합니다.'
    ]],
    ['학습 현황 탭', [
      '학생이 활동을 마칠 때마다 모은 새싹, 획득 별, 완료 여부가 자동으로 채워집니다.',
      '반 전체의 진행을 한눈에 보려면 상단 ' + 앱_이름 + ' 메뉴의 학습 대시보드 열기를 누르세요.'
    ]],
    ['주의', [
      '데이터나 설정을 바꿀 때는 앱 화면이 아니라 해당 시트 탭에서 직접 수정하세요.',
      '탭 이름은 코드와 연결되어 있으므로 삭제하거나 바꾸지 마세요.'
    ]],
    ['저작권 안내', [
      '이 앱과 시트, 코드의 저작권은 GEG화성(깊이 e끌림)에 있습니다.',
      '저작권법의 보호를 받는 저작물로, 무단 복제·재배포·재판매·리브랜딩을 금지합니다.',
      '정상 경로로 받은 이용자라도 위 행위는 허용되지 않으며, 위반 시 저작권법에 따른 책임이 따를 수 있습니다.',
      '이용 문의: bacusiki777@gmail.com, for2102@jimj.kr'
    ]]
  ];

  // 2차원 배열로 한 번에 입력
  const rows = [];
  rows.push([앱_이름 + ' 사용 설명', '']);
  rows.push(['', '']);
  let n = 1;
  sections.forEach(function (sec) {
    rows.push([sec[0], '']);
    sec[1].forEach(function (line) {
      rows.push([n + '.', line]);
      n++;
    });
    rows.push(['', '']);
  });

  const width = 2;
  sh.getRange(1, 1, rows.length, width).setValues(rows);

  // 서식
  sh.getRange(1, 1, 1, width).merge().setFontSize(14).setFontWeight('bold')
    .setBackground('#dfe6d8').setHorizontalAlignment('center');
  // 섹션 제목 줄 강조 (B칸이 빈 줄 중 A칸에 글자가 있는 줄)
  for (let r = 3; r <= rows.length; r++) {
    const a = rows[r - 1][0], b = rows[r - 1][1];
    if (a && !b && !/^\d+\.$/.test(a)) {
      sh.getRange(r, 1, 1, width).merge().setFontWeight('bold').setBackground('#eef3ea');
    }
  }
  sh.getRange(1, 1, rows.length, width).setVerticalAlignment('middle').setWrap(true);
  sh.getRange(1, 1, rows.length, width).setBorder(true, true, true, true, true, true);
  sh.setColumnWidth(1, 90);
  sh.setColumnWidth(2, 560);
  sh.setFrozenRows(1);

  // 안내는 한 번만
  ss.toast('사용 설명 탭을 만들었습니다.', 앱_이름, 5);
}

/* ============================================================
 * 학습 대시보드 — 외부 CDN 없이 내장 HTML 팝업으로 띄운다
 * ============================================================ */
function openDashboard() {
  const html = HtmlService.createHtmlOutput(DASHBOARD_HTML_())
    .setWidth(920).setHeight(640).setTitle(앱_이름 + ' 학습 대시보드');
  SpreadsheetApp.getUi().showModalDialog(html, 앱_이름 + ' 학습 대시보드');
}

/* 대시보드가 부르는 데이터 제공 함수 (google.script.run 으로 호출)
   주의: 이름 끝에 _ 를 붙이면 비공개가 되어 google.script.run 이 못 부른다. */
function getDashboardData() {
  const sh = 현황시트_();
  const last = Math.max(2, sh.getLastRow());
  const vals = sh.getRange(2, 1, last - 1, 6).getValues();
  const rows = [];
  for (let i = 0; i < vals.length; i++) {
    const no = vals[i][0];
    if (no === '' || no === null) continue;
    rows.push({
      no: Number(no),
      name: String(vals[i][1] || ('학생' + no)),
      sprouts: String(vals[i][2] || ''),
      stars: String(vals[i][3] || ''),
      done: String(vals[i][4] || ''),
      when: vals[i][5] ? new Date(vals[i][5]).toLocaleString('ko-KR') : ''
    });
  }
  return { app: 앱_이름, rows: rows };
}

function DASHBOARD_HTML_() {
  return [
'<!DOCTYPE html><html lang="ko"><head><meta charset="UTF-8"><style>',
'  body{font-family:"Noto Sans KR",sans-serif;color:#3b3020;margin:0;padding:16px;background:#fbf7ee;}',
'  h2{margin:0 0 12px;font-size:20px;}',
'  .sum{display:flex;gap:10px;margin-bottom:14px;flex-wrap:wrap;}',
'  .card{background:#fff;border:1px solid #e7d9bd;border-radius:10px;padding:10px 16px;min-width:120px;}',
'  .card .n{font-size:24px;font-weight:700;color:#3f8c58;}',
'  .card .l{font-size:13px;color:#8a7a5b;}',
'  table{border-collapse:collapse;width:100%;background:#fff;font-size:14px;}',
'  th,td{border:1px solid #e7d9bd;padding:7px 9px;text-align:center;}',
'  th{background:#eef3ea;}',
'  .bar{height:12px;background:#eadfc7;border-radius:6px;overflow:hidden;}',
'  .bar > i{display:block;height:100%;background:#5bb377;}',
'  .muted{color:#a99b7c;}',
'</style></head><body>',
'  <h2>반 전체 학습 현황</h2>',
'  <div class="sum" id="sum"><span class="muted">불러오는 중…</span></div>',
'  <table id="tbl"><thead><tr><th>번호</th><th>이름</th><th>모은 새싹</th><th>진행</th><th>획득 별</th><th>완료</th><th>수정 시각</th></tr></thead><tbody></tbody></table>',
'<script>',
'  function pct(s){var m=String(s).match(/(\\d+)\\s*\\/\\s*(\\d+)/);if(!m||+m[2]===0)return 0;return Math.round(+m[1]/+m[2]*100);}',
'  function render(d){',
'    var rows=d.rows||[];',
'    var started=rows.filter(function(r){return r.sprouts||r.done;});',
'    var done=rows.filter(function(r){return r.done==="완료";});',
'    document.getElementById("sum").innerHTML=',
'      card(rows.length,"명단 학생")+card(started.length,"시작한 학생")+card(done.length,"전체 완료");',
'    var tb=document.querySelector("#tbl tbody");tb.innerHTML="";',
'    rows.forEach(function(r){',
'      var p=pct(r.sprouts);',
'      var tr=document.createElement("tr");',
'      tr.innerHTML="<td>"+r.no+"</td><td>"+esc(r.name)+"</td><td>"+esc(r.sprouts)+"</td>"+',
'        "<td><div class=\\"bar\\"><i style=\\"width:"+p+"%\\"></i></div></td>"+',
'        "<td>"+esc(r.stars)+"</td><td>"+(r.done==="완료"?"✔":"")+"</td><td class=\\"muted\\">"+esc(r.when)+"</td>";',
'      tb.appendChild(tr);',
'    });',
'  }',
'  function card(n,l){return "<div class=\\"card\\"><div class=\\"n\\">"+n+"</div><div class=\\"l\\">"+l+"</div></div>";}',
'  function esc(s){return String(s==null?"":s).replace(/[&<>]/g,function(c){return{"&":"&amp;","<":"&lt;",">":"&gt;"}[c];});}',
'  google.script.run.withSuccessHandler(render).withFailureHandler(function(e){',
'    document.getElementById("sum").innerHTML="<span class=\\"muted\\">불러오기 실패: "+e.message+"</span>";',
'  }).getDashboardData();',
'</'+'script></body></html>'
  ].join('\n');
}
