/**
 * 우리 아이 앱 — 가족 공유 백엔드 (Google Apps Script)
 *
 * 설치(엄마가 처음 한 번만):
 * 1) 구글드라이브에서 새 [Google 스프레드시트] 만들기 (이 시트가 우리 가족 데이터 저장소)
 * 2) 시트 상단 메뉴 [확장 프로그램] → [Apps Script] 열기
 * 3) 기본 코드 지우고 이 파일 내용 전체 붙여넣기 → 저장
 * 4) [프로젝트 설정(톱니)] → [스크립트 속성] → 속성 추가:
 *      이름: FAMILY_KEY   값: (가족끼리 정한 비밀 코드, 예: uri-seojun-2026)
 * 5) [배포] → [새 배포] → 유형 [웹 앱]
 *      - 실행: 나
 *      - 액세스 권한: "모든 사용자"
 *    → 배포 → 나오는 [웹 앱 URL](.../exec)을 복사
 * 6) 앱에서 👨‍👩‍👧 → 웹앱 주소(그 URL) + 가족 키(FAMILY_KEY 값) 입력 → 연결
 *
 * ※ 코드 수정 후 재배포는 반드시 [배포 관리] → 연필(수정) → [새 버전] 으로.
 *   "새 배포"를 또 만들면 URL이 바뀌어 기존 연결이 끊깁니다.
 */

var SHEET_NAME = 'data';   // 항목 저장 시트
var DEL_NAME   = 'deleted'; // 삭제표시 저장 시트

function KEY_() {
  return PropertiesService.getScriptProperties().getProperty('FAMILY_KEY') || '';
}
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
function sheet_(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(headers); }
  return sh;
}

// 저장소 읽기 → {entries:[...], deleted:{id:ts}}
function read_() {
  var ds = sheet_(SHEET_NAME, ['id', 'updatedAt', 'json']);
  var del = sheet_(DEL_NAME, ['id', 'deletedAt']);
  var entries = [];
  var dv = ds.getDataRange().getValues();
  for (var i = 1; i < dv.length; i++) {
    if (!dv[i][0]) continue;
    try { entries.push(JSON.parse(dv[i][2])); } catch (e) {}
  }
  var deleted = {};
  var dd = del.getDataRange().getValues();
  for (var j = 1; j < dd.length; j++) {
    if (dd[j][0]) deleted[dd[j][0]] = Number(dd[j][1]) || 0;
  }
  return { entries: entries, deleted: deleted };
}

// 저장소 통째로 다시 쓰기
function write_(state) {
  var ds = sheet_(SHEET_NAME, ['id', 'updatedAt', 'json']);
  var del = sheet_(DEL_NAME, ['id', 'deletedAt']);
  ds.clearContents(); ds.appendRow(['id', 'updatedAt', 'json']);
  var rows = state.entries.map(function (e) {
    return [e.id, e.updatedAt || 0, JSON.stringify(e)];
  });
  if (rows.length) ds.getRange(2, 1, rows.length, 3).setValues(rows);
  del.clearContents(); del.appendRow(['id', 'deletedAt']);
  var drows = Object.keys(state.deleted).map(function (id) { return [id, state.deleted[id]]; });
  if (drows.length) del.getRange(2, 1, drows.length, 2).setValues(drows);
}

// 클라이언트와 동일한 병합 규칙 (id별 최신 우선 + 삭제표시). 서버에서도 합쳐 충돌 최소화.
function merge_(a, b) {
  var deleted = {}, id;
  for (id in a.deleted) deleted[id] = a.deleted[id];
  for (id in b.deleted) deleted[id] = Math.max(deleted[id] || 0, b.deleted[id]);
  var byId = {};
  [].concat(a.entries, b.entries).forEach(function (e) {
    var cur = byId[e.id];
    if (!cur || (e.updatedAt || 0) > (cur.updatedAt || 0)) byId[e.id] = e;
  });
  var entries = [];
  Object.keys(byId).forEach(function (k) {
    var e = byId[k];
    if (!(deleted[e.id] && deleted[e.id] >= (e.updatedAt || 0))) entries.push(e);
  });
  return { entries: entries, deleted: deleted };
}

function doGet(e) {
  if ((e.parameter.key || '') !== KEY_()) return json_({ ok: false, error: 'key' });
  return json_({ ok: true, state: read_() });
}

function doPost(e) {
  var body = {};
  try { body = JSON.parse(e.postData.contents || '{}'); } catch (err) { return json_({ ok: false, error: 'badjson' }); }
  if ((body.key || '') !== KEY_()) return json_({ ok: false, error: 'key' });

  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (err) { return json_({ ok: false, error: 'busy' }); }
  try {
    var incoming = body.state || { entries: [], deleted: {} };
    var merged = merge_(read_(), incoming);
    write_(merged);
    return json_({ ok: true, state: merged });
  } finally {
    lock.releaseLock();
  }
}
