/**
 * WBS 구글 시트 동기화 백엔드.
 *
 * 사용법:
 * 1. 구글 시트를 새로 만든다.
 * 2. 확장 프로그램 > Apps Script 메뉴를 열고, 기본 코드를 지운 뒤 이 파일 내용을 전부 붙여넣는다.
 * 3. 저장 후 배포 > 새 배포 > 유형: 웹 앱 선택.
 *    - 실행 계정: 나
 *    - 액세스 권한이 있는 사용자: 전체
 * 4. 배포 후 나오는 웹 앱 URL(.../exec 로 끝남)을 복사한다.
 * 5. wbs.html의 "동기화 설정"에 그 URL을 붙여넣는다.
 *
 * 이미 배포해서 쓰고 있다면: 이 내용으로 코드를 교체한 뒤
 * 배포 > 배포 관리 > (연필 아이콘) > 버전: 새 버전 > 배포 를 누르면
 * URL은 그대로 유지한 채 갱신된다.
 *
 * 시트 "Data" 탭에 다음을 저장한다.
 *   A1 : WBS 전체 JSON 문자열 (그룹/업무 구조를 그대로 유지하기 위해 통째로 보관)
 *   B1 : 마지막 저장 시각
 *   C1 : 리비전 번호 (저장할 때마다 1씩 증가)
 *
 * C1의 리비전이 동시 편집 충돌 감지에 쓰인다. 클라이언트는 마지막으로 읽은 리비전을
 * _baseRev로 함께 보내고, 그 사이 다른 사람이 저장해 리비전이 올라갔으면
 * 덮어쓰지 않고 conflict 응답으로 현재 서버 내용을 돌려준다.
 */

var SHEET_NAME = 'Data';

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange('A1').setValue('{"groups":[]}');
    sheet.getRange('A1').setNote('WBS 데이터 (JSON) - 직접 수정하지 마세요');
    sheet.getRange('C1').setValue(0);
    sheet.getRange('C1').setNote('리비전 번호 - 직접 수정하지 마세요');
  }
  return sheet;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
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

function doGet(e) {
  var sheet = getSheet_();
  var payload = readData_(sheet);
  payload.rev = readRev_(sheet);
  payload.updatedAt = readUpdatedAt_(sheet);
  return json_(payload);
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
  } catch (err) {
    return json_({ ok: false, error: '다른 저장이 진행 중입니다. 잠시 후 다시 시도해주세요.' });
  }
  try {
    var sheet = getSheet_();
    var parsed = JSON.parse(e.postData.contents);
    if (!parsed || !Array.isArray(parsed.groups)) {
      throw new Error('groups 배열이 없는 데이터입니다');
    }

    var baseRev = (typeof parsed._baseRev === 'number') ? parsed._baseRev : null;
    var force = parsed._force === true;
    var currentRev = readRev_(sheet);

    // 클라이언트가 마지막으로 읽은 리비전과 서버의 현재 리비전이 다르면
    // 그 사이 다른 사람이 저장한 것이므로 덮어쓰지 않고 현재 내용을 돌려준다.
    if (baseRev !== null && !force && baseRev !== currentRev) {
      return json_({
        ok: false,
        conflict: true,
        rev: currentRev,
        updatedAt: readUpdatedAt_(sheet),
        data: readData_(sheet)
      });
    }

    delete parsed._baseRev;
    delete parsed._force;

    var nextRev = currentRev + 1;
    sheet.getRange('A1').setValue(JSON.stringify(parsed));
    sheet.getRange('B1').setValue(new Date());
    sheet.getRange('B1').setNote('마지막 저장 시각');
    sheet.getRange('C1').setValue(nextRev);
    sheet.getRange('C1').setNote('리비전 번호 - 직접 수정하지 마세요');
    SpreadsheetApp.flush();

    return json_({ ok: true, rev: nextRev, updatedAt: readUpdatedAt_(sheet) });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
