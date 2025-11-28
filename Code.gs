
// -----------------------------------------------------------------
// [메인 함수] 웹 앱 접속 시 HTML 페이지를 반환
// -----------------------------------------------------------------
function doGet(e) {
  // 1. 템플릿 파일 생성 ('index.html'을 바라봄)
  const t = HtmlService.createTemplateFromFile('index');

  // 2. URL 파라미터 전달 (필요한 경우 유지)
  // 예: ?view=detail&dual=true
  t.view = (e && e.parameter && e.parameter.view) || 'list';
  t.isDualMode = (e && e.parameter && e.parameter.dual) === 'true';

  // 3. 스크립트 실행 URL 전달 (폼 제출 등이 필요할 때 사용)
  t.scriptUrl = ScriptApp.getService().getUrl();

  // 4. HTML 변환 및 반환
  // X-Frame-Options: ALLOWALL (iframe 임베딩 허용, 필요 시 조절)
  return HtmlService.createHtmlOutput(t.evaluate().getContent())
      .setTitle('고객 상담일지 기록시스템')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


function include(filename) {
  // TemplateFromFile은 내용을 '있는 그대로의 텍스트'로 가져옴 (성공)
  return HtmlService.createTemplateFromFile(filename).getRawContent();
}
