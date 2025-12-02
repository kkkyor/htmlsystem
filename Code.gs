// -----------------------------------------------------------------
// [메인 함수] 웹 앱 접속 시 HTML 페이지를 반환
// -----------------------------------------------------------------
function doGet(e) {
  const t = HtmlService.createTemplateFromFile('index');
  t.view = (e && e.parameter && e.parameter.view) || 'list';
  t.isDualMode = (e && e.parameter && e.parameter.dual) === 'true';
  t.scriptUrl = ScriptApp.getService().getUrl();

  return HtmlService.createHtmlOutput(t.evaluate().getContent())
      .setTitle('고객 상담일지 기록시스템')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createTemplateFromFile(filename).getRawContent();
}

// -----------------------------------------------------------------
// [신규] 계약 데이터를 구글 시트에 추가하는 함수
// -----------------------------------------------------------------
function appendContractToSheet(data) {
  // 1. 대상 스프레드시트 ID (공유해주신 시트 URL에서 추출)
  const SHEET_ID = '10CxVE3BYlEwO2MOBIhaQ3WCe4JEGVyJ-CYPZ0vz8mRE';
  const SHEET_NAME = '9월'; // 시트 이름 확인 필요 (기본값: 시트1)

  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0]; // 이름 없으면 첫 번째 시트 사용
    
    // 2. 입력할 데이터 배열 생성 (열 순서에 맞춤)
    // A열: (비워둠 or 체크박스), B: 고객명, C: 차종, D: 담당자, E: 계약접수, F: (비워둠), G: 유입경로
    // 날짜는 포맷팅해서 입력
    const rowData = [
      false,          // A열: 체크박스 (False로 시작)
      data.name,      // B열: 고객명
      data.car,       // C열: 차종
      data.manager,   // D열: 담당자
      data.receiptDate, // E열: 계약접수
      "",             // F열
      data.leadSource // G열: 유입경로
    ];

    // 3. 마지막 행 다음 줄에 추가
    sheet.appendRow(rowData);
    
    // 4. (선택사항) A열에 체크박스 데이터 유효성 검사 추가 (이미 설정되어 있다면 생략 가능)
    const lastRow = sheet.getLastRow();
    const checkboxCell = sheet.getRange(lastRow, 1);
    const rule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
    checkboxCell.setDataValidation(rule);

    return { success: true };
  } catch (e) {
    Logger.log("Error appending to sheet: " + e.toString());
    throw new Error("시트 저장 실패: " + e.message);
  }
}
