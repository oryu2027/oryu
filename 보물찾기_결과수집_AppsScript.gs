/**
 * 오류고 VR탐방 게임 — 참여 결과 수집용 Google Apps Script
 * (VR 보물찾기와 틀린 그림 찾기가 웹 앱 주소 하나를 함께 씁니다)
 *
 * 설치 방법
 * 1. 구글 시트 'VR 게임 결과'를 엽니다.
 * 2. 메뉴 [확장 프로그램] → [Apps Script]를 엽니다. (script.google.com에서 만든 프로젝트여도 됩니다)
 * 3. 편집기에 있던 내용을 모두 지우고 이 파일 내용을 그대로 붙여 넣은 뒤 저장합니다.
 * 4. 오른쪽 위 [배포] → [새 배포] → 유형 선택(톱니바퀴) → [웹 앱]
 *      - 다음 사용자 인증 정보로 실행: 나
 *      - 액세스 권한이 있는 사용자: 모든 사용자
 *    → [배포]를 누르고, 권한 승인 창이 뜨면 본인 계정으로 허용합니다.
 * 5. 마지막에 나오는 "웹 앱 URL"(https://script.google.com/macros/s/.../exec)을 복사합니다.
 * 6. treasure-hunt.html과 spot-the-difference.html 두 파일 모두에서
 *      var RESULT_URL = '';
 *    의 따옴표 안에 그 주소를 붙여 넣고 GitHub에 다시 올립니다.
 *
 * 결과는 게임별로 다른 탭에 날짜와 함께 한 줄씩 계속 쌓입니다(지워지지 않음).
 *   - VR 보물찾기   → '결과' 탭
 *   - 틀린 그림 찾기 → '틀린그림 결과' 탭
 * 게임의 결과판에서 날짜를 골라 볼 수 있고, 엑셀로 내려받으면 날짜별 시트로 나뉩니다.
 * 시트에서 날짜별로 보려면 1행의 '날짜' 칸에 필터를 걸면 됩니다.
 * 기록을 지우고 싶으면 시트에서 2행부터 아래 줄을 지우면 됩니다(1행 제목은 남겨 두세요).
 *
 * 이미 예전 판을 배포했다면: 내용을 바꿔 저장한 뒤 [배포] → [배포 관리] → 연필 아이콘 →
 * 버전 '새 버전' → [배포]를 누르세요. 웹 앱 주소는 그대로 유지됩니다.
 *
 * 편집기에서 확인: 위쪽 함수 목록에서 setup 을 골라 [실행] → 시트에 '결과'·'틀린그림 결과' 탭이 생기면 정상.
 *   (doPost를 직접 실행하면 'postData' 오류가 나는 것은 정상입니다. 게임이 보낼 때만 동작하는 함수입니다.)
 * 관리자 비밀번호: 아래 ADMIN_KEY 의 따옴표 안에 정하고 저장 → 새 버전으로 배포하세요.
 *   게임의 결과판은 누구에게나 순위·닉네임만 보여 주고, [관리자]에서 이 비밀번호를 넣어야
 *   상세 결과와 엑셀 내려받기가 열립니다. 구글 시트 자체는 '공유'하지 말고 선생님만 보세요.
 *
 * 연결 확인: 웹 앱 주소를 시크릿 창(로그인하지 않은 창)에서 열었을 때 {"admin":false,"rows":[...]} 처럼 짧은 글자가
 * 보이면 정상입니다. 로그인 화면이 나오면 배포의 '액세스 권한이 있는 사용자'를 '모든 사용자'로 바꿔야 합니다.
 */

var TZ = 'Asia/Seoul';
// 결과를 적을 구글 시트('VR 게임 결과')의 ID.
// 스크립트를 시트의 [확장 프로그램]에서 만들었든, script.google.com에서 따로 만들었든 이 시트에 기록합니다.
var SHEET_ID = '1nBMavCAOtQz7pywBI6cWoxcyAaCU1X2071ITHHp6NZs';

// ★ 관리자 비밀번호: 게임 결과판의 [관리자] 화면(상세 결과·엑셀 내려받기)을 열 때 입력하는 비밀번호입니다.
//   따옴표 안을 원하는 비밀번호로 바꾼 뒤 저장하고, [배포] → [배포 관리] → 새 버전으로 배포하세요.
//   학생들에게는 순위와 닉네임만 전달되고, 이 비밀번호가 맞을 때만 상세 정보가 전달됩니다.
//   (비밀번호는 이 스크립트 안에만 있고 GitHub의 게임 파일에는 들어가지 않습니다.)
//   비워 두거나 아래 예시 그대로 두면 관리자 화면이 열리지 않습니다.
var ADMIN_KEY = '여기에-관리자-비밀번호';

function book_() {
  try { var a = SpreadsheetApp.getActiveSpreadsheet(); if (a && a.getId() === SHEET_ID) return a; } catch (err) {}
  return SpreadsheetApp.openById(SHEET_ID);
}

// ---------------- VR 보물찾기 ----------------
var T_SHEET = '결과';
var T_PLACES = ['오류고 전경', '작은 정원', '청운실', '체육관', '클라이밍실', '도서관', '과학실'];
var T_HEAD = ['기록ID', '날짜', '닉네임', '찾은 보물', '전체 보물', '둘러본 장소', '전체 장소', '걸린 시간(초)', '참여 일시'].concat(T_PLACES);

// ---------------- 틀린 그림 찾기 ----------------
var S_SHEET = '틀린그림 결과';
var S_PLACES = ['도서관', '과학실', '급식실', '매점', '클라이밍실'];
var S_HEAD = ['기록ID', '날짜', '닉네임', '레벨', '칭호', '완료 장소', '전체 장소', '찾은 곳', '전체 틀린 곳', '걸린 시간(초)', '참여 일시'].concat(S_PLACES);

function sheet_(name, head) {
  var ss = book_();
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.appendRow(head);
    sh.getRange(1, 1, 1, head.length).setFontWeight('bold').setBackground('#E3E7DF');
    sh.setFrozenRows(1);
  }
  return sh;
}

function clean_(v, max) {
  return String(v == null ? '' : v).replace(/^[=+\-@]/, "'").slice(0, max || 40);
}
function num_(v, max) {
  var n = Number(v);
  return isFinite(n) && n >= 0 && n <= max ? Math.round(n) : 0;
}
function day_(d, now) {
  return /^\d{4}-\d{2}-\d{2}$/.test(d.day) ? d.day : Utilities.formatDate(now, TZ, 'yyyy-MM-dd');
}
function dayOut_(v) {
  return v instanceof Date ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd') : String(v);
}
function atOut_(v) {
  return v instanceof Date ? v.toISOString() : String(v);
}

// 편집기에서 실행해 보는 연결 확인용 함수입니다.
// 위쪽 함수 목록에서 'setup'을 고르고 [실행]을 누르면, 두 결과 탭을 만들고 실행 로그에 확인 결과를 보여 줍니다.
// (doPost·doGet은 게임이 부를 때만 동작하므로 편집기에서 직접 실행하지 마세요.)
function setup() {
  var book = book_();
  var t = sheet_(T_SHEET, T_HEAD), s = sheet_(S_SHEET, S_HEAD);
  Logger.log('연결된 시트: ' + book.getName());
  Logger.log("'" + T_SHEET + "' 탭 기록 " + (t.getLastRow() - 1) + "건, '" + S_SHEET + "' 탭 기록 " + (s.getLastRow() - 1) + '건');
  Logger.log('준비 완료! 이제 [배포] → [배포 관리] → 새 버전으로 배포하세요.');
}

// 게임이 끝날 때 결과 한 건을 받아 시트에 추가합니다.
function doPost(e) {
  if (!e || !e.postData) {
    Logger.log('doPost는 게임이 결과를 보낼 때만 동작합니다. 연결을 확인하려면 setup 함수를 실행하세요.');
    return out_({ ok: false, error: 'no postData' });
  }
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var d = JSON.parse(e.postData.contents);
    if (!d || !d.nick) return out_({ ok: false });
    var per = d.per || {};
    var now = new Date();
    var row;
    if (d.game === 'spot') {
      row = [
        clean_(d.id, 30), "'" + day_(d, now), clean_(d.nick, 20),
        num_(d.level, 10), clean_(d.levelName, 20), num_(d.cleared, 20), num_(d.placeTotal, 20),
        num_(d.found, 100), num_(d.total, 100), num_(d.sec, 86400), now
      ].concat(S_PLACES.map(function (p) { return num_(per[p], 10); }));
      sheet_(S_SHEET, S_HEAD).appendRow(row);
    } else {
      row = [
        clean_(d.id, 30), "'" + day_(d, now), clean_(d.nick, 20),
        num_(d.found, 100), num_(d.total, 100), num_(d.places, 20), num_(d.placeTotal, 20),
        num_(d.sec, 86400), now
      ].concat(T_PLACES.map(function (p) { return num_(per[p], 10); }));
      sheet_(T_SHEET, T_HEAD).appendRow(row);
    }
    return out_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

function isAdmin_(key) {
  if (!ADMIN_KEY || ADMIN_KEY === '여기에-관리자-비밀번호') return false;
  return String(key || '') === ADMIN_KEY;
}

// 게임의 "참여자 결과판"이 이 주소에서 기록을 읽어 갑니다. (?game=spot 이면 틀린 그림 찾기)
//  - 비밀번호 없이: 순위 순서대로 닉네임과 날짜만 보냅니다(학생용).
//  - ?key=관리자 비밀번호: 상세 기록 전체를 보냅니다(관리자 화면·엑셀용).
function doGet(e) {
  var p = (e && e.parameter) || {};
  var game = p.game;
  var rows = [];
  if (game === 'spot') {
    var ss = sheet_(S_SHEET, S_HEAD), n = ss.getLastRow() - 1;
    if (n > 0) ss.getRange(2, 1, n, S_HEAD.length).getValues().forEach(function (r) {
      if (!r[2]) return;
      var per = {};
      S_PLACES.forEach(function (p, i) { per[p] = Number(r[11 + i]) || 0; });
      rows.push({
        id: String(r[0]), game: 'spot', day: dayOut_(r[1]), nick: String(r[2]),
        level: Number(r[3]) || 1, levelName: String(r[4]),
        cleared: Number(r[5]) || 0, placeTotal: Number(r[6]) || 0,
        found: Number(r[7]) || 0, total: Number(r[8]) || 0,
        sec: Number(r[9]) || 0, at: atOut_(r[10]), per: per
      });
    });
  } else {
    var sh = sheet_(T_SHEET, T_HEAD), m = sh.getLastRow() - 1;
    if (m > 0) sh.getRange(2, 1, m, T_HEAD.length).getValues().forEach(function (r) {
      if (!r[2]) return;
      var per = {};
      T_PLACES.forEach(function (p, i) { per[p] = Number(r[9 + i]) || 0; });
      rows.push({
        id: String(r[0]), day: dayOut_(r[1]), nick: String(r[2]),
        found: Number(r[3]) || 0, total: Number(r[4]) || 0,
        places: Number(r[5]) || 0, placeTotal: Number(r[6]) || 0,
        sec: Number(r[7]) || 0, at: atOut_(r[8]), per: per
      });
    });
  }
  if (p.key && isAdmin_(p.key)) return out_({ admin: true, rows: rows });
  if (p.key) Utilities.sleep(1500); // 비밀번호를 마구 넣어 보는 것을 늦춥니다.
  // 학생용: 순위를 서버에서 매긴 뒤 순서·닉네임·날짜만 보냅니다(점수·시간 등 상세 정보는 보내지 않음).
  rows.sort(function (a, b) {
    if (game === 'spot' && b.cleared !== a.cleared) return b.cleared - a.cleared;
    if (b.found !== a.found) return b.found - a.found;
    if (a.sec !== b.sec) return a.sec - b.sec;
    return a.at < b.at ? -1 : a.at > b.at ? 1 : 0;
  });
  return out_({ admin: false, rows: rows.map(function (r) { return { id: r.id, nick: r.nick, day: r.day }; }) });
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
