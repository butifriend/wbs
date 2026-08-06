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
 * 데이터는 시트의 "Data" 탭 A1 셀에 WBS 전체를 JSON 문자열 그대로 저장한다.
 * (그룹/업무 구조를 그대로 유지하기 위해 스프레드시트 행/열로 쪼개지 않고 JSON 통째로 보관)
 */

var SHEET_NAME = 'Data';

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange('A1').setValue('{"groups":[]}');
    sheet.getRange('A1').setNote('WBS 데이터 (JSON) - 직접 수정하지 마세요');
  }
  return sheet;
}

function doGet(e) {
  var sheet = getSheet_();
  var json = sheet.getRange('A1').getValue() || '{"groups":[]}';
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    var sheet = getSheet_();
    var body = e.postData.contents;
    var parsed = JSON.parse(body);
    if (!parsed || !Array.isArray(parsed.groups)) {
      throw new Error('groups 배열이 없는 데이터입니다');
    }
    sheet.getRange('A1').setValue(body);
    sheet.getRange('B1').setValue(new Date());
    sheet.getRange('B1').setNote('마지막 저장 시각');
    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
