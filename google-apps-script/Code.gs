/**
 * WBS 구글 시트 동기화 백엔드 (부서별 시트 분리)
 *
 * 사용법:
 * 1. 구글 시트를 새로 만든다.
 * 2. 확장 프로그램 > Apps Script 메뉴를 열고, 기본 코드를 지운 뒤 이 파일 내용을 전부 붙여넣는다.
 * 3. 저장 후 배포 > 새 배포 > (톱니바퀴) 유형: 웹 앱 선택.
 *    - 실행 계정: 나
 *    - 액세스 권한이 있는 사용자: 전체
 * 4. 배포 후 나오는 웹 앱 URL(.../exec 로 끝남)을 복사한다.
 * 5. wbs.html의 "동기화 설정"에 그 URL을 붙여넣는다.
 *
 * 이미 배포해서 쓰고 있다면: 이 내용으로 코드를 교체한 뒤
 * 배포 > 배포 관리 > (연필 아이콘) > 버전: 새 버전 > 배포 를 누르면
 * URL은 그대로 유지한 채 갱신된다.
 *
 * 시트 구성
 *   _Depts                  부서 목록 (숨김 시트). A열 id, B열 이름.
 *   _Works                  업무 목록 (숨김 시트). A열 부서id, B열 업무id, C열 이름.
 *   Data                    "기본" 부서 / "기본" 업무 (기존 시트를 그대로 쓴다)
 *   Data_<dept>             그 부서의 "기본" 업무
 *   Data_<dept>__<work>     그 부서의 그 외 업무
 *
 * 업무는 부서 안에서만 의미가 있다. 부서마다 업무 목록이 따로다.
 * "기본" 업무는 등록 없이 늘 존재하며, 기존 시트가 곧 그 부서의 기본 업무다.
 * 그래서 이 코드로 갱신해도 기존 데이터는 그대로 보인다.
 *
 * 각 데이터 시트의 첫 줄
 *   A1  WBS 전체 JSON 문자열 (그룹/업무 구조를 그대로 유지하기 위해 통째로 보관)
 *   B1  마지막 저장 시각
 *   C1  리비전 번호 (저장할 때마다 1씩 증가)
 *
 * C1의 리비전이 동시 편집 충돌 감지에 쓰인다. 클라이언트는 마지막으로 읽은 리비전을
 * _baseRev로 함께 보내고, 그 사이 다른 사람이 저장해 리비전이 올라갔으면
 * 덮어쓰지 않고 conflict 응답으로 현재 서버 내용을 돌려준다. 리비전은 부서마다 따로다.
 *
 * 관리자 모드 (쓰기 보호)
 *   스크립트 속성에 ADMIN_KEY 를 넣어 두면 쓰기(POST)에 그 열쇠를 요구한다.
 *   읽기(GET)는 그대로 열려 있어 담당자는 URL만으로 조회할 수 있다.
 *
 *   설정 방법: Apps Script 편집기 좌측 ⚙ 프로젝트 설정
 *             > 스크립트 속성 > 속성 추가
 *             속성 = ADMIN_KEY , 값 = 원하는 열쇠 문자열
 *
 *   ADMIN_KEY 를 설정하지 않으면 지금까지처럼 누구나 저장할 수 있다.
 *   (재배포 직후 열쇠를 넣기 전에 아무도 저장 못 하는 일을 막기 위한 기본값)
 *   열쇠가 새면 이 속성 값만 바꾸면 된다. 재배포는 필요 없다.
 *
 * API
 *   GET  ?dept=<id>&work=<id>  그 부서/업무의 데이터 + rev + updatedAt + 목록 + authRequired
 *   GET  ?dept=<id>&action=rev 리비전만 (주기 확인용, 가볍다)
 *   GET  ?action=depts         부서 목록만
 *   POST {_action:'checkKey', _key}                      열쇠 확인만 (저장 안 함)
 *   POST {..state, _dept, _work, _baseRev, _force}       데이터 저장
 *   POST {_action:'createDept', _name}                   부서 추가
 *   POST {_action:'renameDept', _dept, _name}            부서 이름 변경
 *   POST {_action:'removeDept', _dept}                   목록에서 제거 (시트는 남긴다)
 *   POST {_action:'createWork', _dept, _name}            업무 추가
 *   POST {_action:'renameWork', _dept, _work, _name}     업무 이름 변경
 *   POST {_action:'removeWork', _dept, _work}            목록에서 제거 (시트는 남긴다)
 */

var DATA_SHEET = 'Data';
var DEPT_SHEET = '_Depts';
var DEFAULT_DEPT = 'default';
var WORK_SHEET = '_Works';
var DEFAULT_WORK = 'default';

/* ---------------- 공통 ---------------- */

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ---------------- 관리자 열쇠 ---------------- */

function adminKey_() {
  try {
    return PropertiesService.getScriptProperties().getProperty('ADMIN_KEY') || '';
  } catch (err) {
    return '';
  }
}

// 열쇠를 설정하지 않았으면 잠그지 않는다.
function keyOk_(parsed) {
  var key = adminKey_();
  if (!key) return true;
  return !!parsed && typeof parsed._key === 'string' && parsed._key === key;
}

function authFail_() {
  return json_({
    ok: false,
    authRequired: true,
    error: '조회 전용입니다. 관리자 열쇠가 필요합니다.'
  });
}

/* ---------------- 부서 목록 ---------------- */

// 기존 Data 시트만 있던 상태에서 처음 열리면 그 시트를 "기본" 부서로 등록한다.
function getDeptRegistry_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(DEPT_SHEET);
  if (!sh) {
    sh = ss.insertSheet(DEPT_SHEET);
    sh.getRange('A1:B1').setValues([['id', 'name']]);
    sh.getRange('A2:B2').setValues([[DEFAULT_DEPT, '기본']]);
    sh.setFrozenRows(1);
    try { sh.hideSheet(); } catch (err) {}
  }
  return sh;
}

function listDepts_() {
  var sh = getDeptRegistry_();
  var last = sh.getLastRow();
  var out = [];
  if (last >= 2) {
    var vals = sh.getRange(2, 1, last - 1, 2).getValues();
    for (var i = 0; i < vals.length; i++) {
      var id = String(vals[i][0] || '').trim();
      if (!id) continue;
      out.push({ id: id, name: String(vals[i][1] || id).trim() || id });
    }
  }
  if (!out.length) out.push({ id: DEFAULT_DEPT, name: '기본' });
  return out;
}

function deptExists_(id) {
  var list = listDepts_();
  for (var i = 0; i < list.length; i++) if (list[i].id === id) return true;
  return false;
}

function deptRowIndex_(id) {
  var sh = getDeptRegistry_();
  var last = sh.getLastRow();
  if (last < 2) return -1;
  var vals = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0] || '').trim() === id) return i + 2;
  }
  return -1;
}

function newDeptId_() {
  return 'd' + String(Date.now()) + String(Math.floor(Math.random() * 1000));
}

/* ---------------- 업무 목록 ---------------- */

// 부서 안에서만 의미가 있으므로 (부서, 업무) 쌍으로 기록한다.
function getWorkRegistry_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(WORK_SHEET);
  if (!sh) {
    sh = ss.insertSheet(WORK_SHEET);
    sh.getRange('A1:C1').setValues([['deptId', 'workId', 'name']]);
    sh.setFrozenRows(1);
    try { sh.hideSheet(); } catch (err) {}
  }
  return sh;
}

// "기본" 업무는 등록하지 않아도 늘 존재한다. 기존 시트가 곧 그 부서의 기본 업무라,
// 이렇게 두면 이 코드로 갱신해도 기존 데이터를 옮길 필요가 없다.
function listWorks_(deptId) {
  var sh = getWorkRegistry_();
  var last = sh.getLastRow();
  var out = [];
  var hasDefault = false;
  if (last >= 2) {
    var vals = sh.getRange(2, 1, last - 1, 3).getValues();
    for (var i = 0; i < vals.length; i++) {
      if (String(vals[i][0] || '').trim() !== deptId) continue;
      var id = String(vals[i][1] || '').trim();
      if (!id) continue;
      if (id === DEFAULT_WORK) hasDefault = true;
      out.push({ id: id, name: String(vals[i][2] || id).trim() || id });
    }
  }
  if (!hasDefault) {
    out.unshift({ id: DEFAULT_WORK, name: '기본' });
  } else {
    // 이름을 바꾸면서 등록 행이 뒤에 붙었을 수 있다. 기본 업무는 늘 맨 앞에 둔다.
    for (var k = 0; k < out.length; k++) {
      if (out[k].id === DEFAULT_WORK) { out.unshift(out.splice(k, 1)[0]); break; }
    }
  }
  return out;
}

function workExists_(deptId, workId) {
  if (workId === DEFAULT_WORK) return true;
  var list = listWorks_(deptId);
  for (var i = 0; i < list.length; i++) if (list[i].id === workId) return true;
  return false;
}

function workRowIndex_(deptId, workId) {
  var sh = getWorkRegistry_();
  var last = sh.getLastRow();
  if (last < 2) return -1;
  var vals = sh.getRange(2, 1, last - 1, 3).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0] || '').trim() === deptId &&
        String(vals[i][1] || '').trim() === workId) return i + 2;
  }
  return -1;
}

function newWorkId_() {
  return 'w' + String(Date.now()) + String(Math.floor(Math.random() * 1000));
}

/* ---------------- 데이터 시트 ---------------- */

// 기본 업무는 예전 이름 그대로 쓴다. 이름이 바뀌면 기존 데이터를 못 찾는다.
function scopeSheetName_(deptId, workId) {
  var base = (!deptId || deptId === DEFAULT_DEPT) ? DATA_SHEET : DATA_SHEET + '_' + deptId;
  return (!workId || workId === DEFAULT_WORK) ? base : base + '__' + workId;
}

function getSheet_(deptIdArg, workIdArg) {
  var deptId = deptIdArg || DEFAULT_DEPT;
  var workId = workIdArg || DEFAULT_WORK;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var name = scopeSheetName_(deptId, workId);
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange('A1').setValue('{"groups":[]}');
    sheet.getRange('C1').setValue(0);
    // 메모는 여기서 한 번만 붙인다. 저장할 때마다 다시 쓰면 그만큼 느려진다.
    sheet.getRange('A1').setNote('WBS 데이터 (JSON) - 직접 수정하지 마세요');
    sheet.getRange('B1').setNote('마지막 저장 시각');
    sheet.getRange('C1').setNote('리비전 번호 - 직접 수정하지 마세요');
  }
  return sheet;
}

function readRev_(sheet) {
  var v = Number(sheet.getRange('C1').getValue());
  return isNaN(v) ? 0 : v;
}

function readUpdatedAt_(sheet) {
  var v = sheet.getRange('B1').getValue();
  if (!v) return '';
  if (v instanceof Date) return v.toISOString();
  return String(v);
}

function readData_(sheet) {
  var raw = sheet.getRange('A1').getValue() || '{"groups":[]}';
  var parsed;
  try {
    parsed = JSON.parse(raw);
  } catch (err) {
    parsed = null;
  }
  if (!parsed || typeof parsed !== 'object' || !Array.isArray(parsed.groups)) {
    parsed = { groups: [] };
  }
  return parsed;
}

/* ---------------- GET ---------------- */

function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.action === 'depts') {
    return json_({ depts: listDepts_(), authRequired: !!adminKey_() });
  }
  var deptId = p.dept || DEFAULT_DEPT;
  if (!deptExists_(deptId)) deptId = DEFAULT_DEPT;
  var workId = p.work || DEFAULT_WORK;
  if (!workExists_(deptId, workId)) workId = DEFAULT_WORK;

  // 화면이 주기적으로 "그새 누가 저장했나"만 물어본다. 전체 데이터를 직렬화하지
  // 않으므로 훨씬 가볍다. Apps Script 는 하루 실행 시간 총량에 제한이 있어,
  // 이 구분이 없으면 주기 조회만으로 그 시간을 다 써 버린다.
  if (p.action === 'rev') {
    var s = getSheet_(deptId, workId);
    return json_({ ok: true, rev: readRev_(s), updatedAt: readUpdatedAt_(s),
                   dept: deptId, work: workId });
  }

  var sheet = getSheet_(deptId, workId);
  var payload = readData_(sheet);
  payload.rev = readRev_(sheet);
  payload.updatedAt = readUpdatedAt_(sheet);
  payload.dept = deptId;
  payload.work = workId;
  // 왕복을 줄이려고 부서 목록을 매번 같이 내려준다.
  payload.depts = listDepts_();
  payload.works = listWorks_(deptId);
  // 화면이 조회 전용으로 시작할지 판단하는 데 쓴다. 열쇠 자체는 내려주지 않는다.
  payload.authRequired = !!adminKey_();
  return json_(payload);
}

/* ---------------- POST ---------------- */

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
  } catch (err) {
    return json_({ ok: false, error: '다른 저장이 진행 중입니다. 잠시 후 다시 시도해주세요.' });
  }
  try {
    var parsed = JSON.parse(e.postData.contents);
    var action = parsed && parsed._action ? String(parsed._action) : '';

    // 열쇠 확인만 하고 아무것도 바꾸지 않는다. 관리자 모드 진입에 쓴다.
    if (action === 'checkKey') {
      return json_({
        ok: keyOk_(parsed),
        authRequired: !!adminKey_(),
        error: keyOk_(parsed) ? '' : '관리자 열쇠가 올바르지 않습니다.'
      });
    }

    // 여기부터는 전부 쓰기 작업이라 열쇠가 필요하다.
    if (!keyOk_(parsed)) return authFail_();

    if (action === 'createDept') return createDept_(parsed);
    if (action === 'renameDept') return renameDept_(parsed);
    if (action === 'removeDept') return removeDept_(parsed);
    if (action === 'createWork') return createWork_(parsed);
    if (action === 'renameWork') return renameWork_(parsed);
    if (action === 'removeWork') return removeWork_(parsed);

    return saveData_(parsed);
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function createDept_(parsed) {
  var name = String(parsed._name || '').trim();
  if (!name) throw new Error('부서 이름이 비어 있습니다');
  var list = listDepts_();
  for (var i = 0; i < list.length; i++) {
    if (list[i].name === name) throw new Error('같은 이름의 부서가 이미 있습니다: ' + name);
  }
  var id = newDeptId_();
  var sh = getDeptRegistry_();
  sh.appendRow([id, name]);
  getSheet_(id); // 데이터 시트를 미리 만들어 둔다
  SpreadsheetApp.flush();
  return json_({ ok: true, dept: id, depts: listDepts_() });
}

function renameDept_(parsed) {
  var id = String(parsed._dept || '').trim();
  var name = String(parsed._name || '').trim();
  if (!id || !name) throw new Error('부서와 이름이 필요합니다');
  var row = deptRowIndex_(id);
  if (row < 0) throw new Error('없는 부서입니다');
  var list = listDepts_();
  for (var i = 0; i < list.length; i++) {
    if (list[i].name === name && list[i].id !== id) {
      throw new Error('같은 이름의 부서가 이미 있습니다: ' + name);
    }
  }
  getDeptRegistry_().getRange(row, 2).setValue(name);
  SpreadsheetApp.flush();
  return json_({ ok: true, depts: listDepts_() });
}

// 목록에서만 뺀다. 데이터 시트는 그대로 남겨서 실수로 지워도 되살릴 수 있게 한다.
function removeDept_(parsed) {
  var id = String(parsed._dept || '').trim();
  if (!id) throw new Error('부서가 필요합니다');
  if (listDepts_().length <= 1) throw new Error('부서가 하나뿐이라 제거할 수 없습니다');
  var row = deptRowIndex_(id);
  if (row < 0) throw new Error('없는 부서입니다');
  getDeptRegistry_().deleteRow(row);
  SpreadsheetApp.flush();
  return json_({ ok: true, depts: listDepts_() });
}

function createWork_(parsed) {
  var deptId = String(parsed._dept || DEFAULT_DEPT);
  if (!deptExists_(deptId)) throw new Error('없는 부서입니다: ' + deptId);
  var name = String(parsed._name || '').trim();
  if (!name) throw new Error('업무 이름이 비어 있습니다');
  var list = listWorks_(deptId);
  for (var i = 0; i < list.length; i++) {
    if (list[i].name === name) throw new Error('같은 이름의 업무가 이미 있습니다: ' + name);
  }
  var id = newWorkId_();
  getWorkRegistry_().appendRow([deptId, id, name]);
  getSheet_(deptId, id); // 데이터 시트를 미리 만들어 둔다
  SpreadsheetApp.flush();
  return json_({ ok: true, dept: deptId, work: id, works: listWorks_(deptId) });
}

function renameWork_(parsed) {
  var deptId = String(parsed._dept || DEFAULT_DEPT);
  var workId = String(parsed._work || '').trim();
  var name = String(parsed._name || '').trim();
  if (!workId || !name) throw new Error('업무와 이름이 필요합니다');
  if (!workExists_(deptId, workId)) throw new Error('없는 업무입니다');
  var list = listWorks_(deptId);
  for (var i = 0; i < list.length; i++) {
    if (list[i].name === name && list[i].id !== workId) {
      throw new Error('같은 이름의 업무가 이미 있습니다: ' + name);
    }
  }
  var row = workRowIndex_(deptId, workId);
  if (row < 0) {
    // 등록 없이 존재하던 기본 업무다. 이름을 바꾸려면 이때 처음 기록한다.
    getWorkRegistry_().appendRow([deptId, workId, name]);
  } else {
    getWorkRegistry_().getRange(row, 3).setValue(name);
  }
  SpreadsheetApp.flush();
  return json_({ ok: true, works: listWorks_(deptId) });
}

// 목록에서만 뺀다. 데이터 시트는 그대로 남겨서 실수로 지워도 되살릴 수 있게 한다.
function removeWork_(parsed) {
  var deptId = String(parsed._dept || DEFAULT_DEPT);
  var workId = String(parsed._work || '').trim();
  if (!workId) throw new Error('업무가 필요합니다');
  if (workId === DEFAULT_WORK) throw new Error('기본 업무는 제거할 수 없습니다');
  var row = workRowIndex_(deptId, workId);
  if (row < 0) throw new Error('없는 업무입니다');
  getWorkRegistry_().deleteRow(row);
  SpreadsheetApp.flush();
  return json_({ ok: true, works: listWorks_(deptId) });
}

function saveData_(parsed) {
  if (!parsed || !Array.isArray(parsed.groups)) {
    throw new Error('groups 배열이 없는 데이터입니다');
  }
  var deptId = String(parsed._dept || DEFAULT_DEPT);
  if (!deptExists_(deptId)) throw new Error('없는 부서입니다: ' + deptId);
  var workId = String(parsed._work || DEFAULT_WORK);
  if (!workExists_(deptId, workId)) throw new Error('없는 업무입니다: ' + workId);

  var sheet = getSheet_(deptId, workId);
  var baseRev = (typeof parsed._baseRev === 'number') ? parsed._baseRev : null;
  var force = parsed._force === true;
  var currentRev = readRev_(sheet);

  // 클라이언트가 마지막으로 읽은 리비전과 서버의 현재 리비전이 다르면
  // 그 사이 다른 사람이 저장한 것이므로 덮어쓰지 않고 현재 내용을 돌려준다.
  if (baseRev !== null && !force && baseRev !== currentRev) {
    return json_({
      ok: false,
      conflict: true,
      dept: deptId,
      work: workId,
      rev: currentRev,
      updatedAt: readUpdatedAt_(sheet),
      data: readData_(sheet)
    });
  }

  delete parsed._baseRev;
  delete parsed._force;
  delete parsed._dept;
  delete parsed._work;
  delete parsed._action;
  delete parsed._name;
  delete parsed._key;

  var nextRev = currentRev + 1;
  var now = new Date();
  // A1·B1·C1 을 한 번에 쓴다. 따로 쓰면 그때마다 시트 서버까지 다녀오느라
  // 저장이 그만큼 길어진다. 메모는 시트를 만들 때 붙여 뒀으므로 건드리지 않는다.
  sheet.getRange('A1:C1').setValues([[JSON.stringify(parsed), now, nextRev]]);
  SpreadsheetApp.flush();

  // 방금 쓴 값을 다시 읽지 않는다. 읽어 봐야 같은 값이고 왕복만 한 번 더 든다.
  return json_({ ok: true, dept: deptId, work: workId, rev: nextRev,
                 updatedAt: now.toISOString() });
}
