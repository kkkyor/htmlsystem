// [GeminiService.gs]
// 대용량 파일 처리 & 빈 응답 자동 복구(Fallback) 기능 탑재

const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_KEY');
const MODEL_NAME = 'gemini-2.0-flash'; 

/**
 * [1단계] 업로드 세션 시작
 */
function initGeminiUploadSession(mimeType, fileSize, fileName) {
  try {
    const initUrl = `https://generativelanguage.googleapis.com/upload/v1beta/files?key=${GEMINI_API_KEY}`;
    const headers = {
      'X-Goog-Upload-Protocol': 'resumable',
      'X-Goog-Upload-Command': 'start',
      'X-Goog-Upload-Header-Content-Length': fileSize.toString(),
      'X-Goog-Upload-Header-Content-Type': mimeType,
      'Content-Type': 'application/json'
    };
    const payload = JSON.stringify({ file: { display_name: fileName } });

    const response = UrlFetchApp.fetch(initUrl, {
      method: 'post',
      headers: headers,
      payload: payload,
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) {
      throw new Error(`[초기화 실패] ${response.getContentText()}`);
    }

    const uploadUrl = response.getHeaders()['x-goog-upload-url'];
    return { success: true, uploadUrl: uploadUrl };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * [2단계] 조각(Chunk) 업로드 중계 (★자동 복구 기능 추가됨★)
 */
function uploadChunkToGemini(uploadUrl, base64Data, offset, totalSize, fileName) {
  try {
    const blob = Utilities.newBlob(Utilities.base64Decode(base64Data));
    const chunkSize = blob.getBytes().length;
    const end = offset + chunkSize - 1;
    const command = (end + 1 === totalSize) ? 'upload, finalize' : 'upload';
    
    const headers = {
      'Content-Range': `bytes ${offset}-${end}/${totalSize}`,
      'X-Goog-Upload-Command': command,
      'X-Goog-Upload-Offset': offset.toString()
    };

    const response = UrlFetchApp.fetch(uploadUrl, {
      method: 'post', 
      headers: headers,
      payload: blob,
      muteHttpExceptions: true
    });

    const code = response.getResponseCode();
    const text = response.getContentText();

    // 1. 진행 중 (308): 정상
    if (code === 308) {
      return { success: true, isDone: false };
    }

    // 2. 완료 (200/201): 정상
    if (code === 200 || code === 201) {
      let fileUri = null;
      
      // (1) 정상적으로 JSON이 온 경우
      if (text) {
        try {
          const json = JSON.parse(text);
          fileUri = json.file.uri;
        } catch (e) {
          // JSON 파싱 실패 시 아래 복구 로직으로 이동
          Logger.log("JSON 파싱 실패, 복구 시도: " + e);
        }
      }

      // (2) [에러 해결 핵심] 내용은 비어있지만 200 OK인 경우 -> 직접 찾는다
      if (!fileUri && fileName) {
         Logger.log("업로드 성공했으나 URI 누락됨. 파일명으로 검색 시도...");
         fileUri = recoverFileUriByName(fileName);
      }

      if (fileUri) {
         return { success: true, isDone: true, fileUri: fileUri };
      } else {
         throw new Error("[업로드 확인 불가] 200 OK를 받았으나 파일 주소를 찾을 수 없습니다.");
      }
    }

    // 3. 그 외 에러
    throw new Error(`[업로드 실패] 상태코드(${code}): ${text}`);

  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// [신규] 파일 주소 찾기 (Fallback 함수)
function recoverFileUriByName(fileName) {
  try {
    const listUrl = `https://generativelanguage.googleapis.com/v1beta/files?key=${GEMINI_API_KEY}`;
    const res = UrlFetchApp.fetch(listUrl);
    const data = JSON.parse(res.getContentText());
    
    if (data.files && data.files.length > 0) {
      // 최신순으로 되어있으므로, 이름이 같은 가장 첫 번째 파일을 찾음
      const match = data.files.find(f => f.displayName === fileName);
      if (match) return match.uri;
    }
  } catch (e) {
    Logger.log("파일 검색 실패: " + e);
  }
  return null;
}

/**
 * [3단계] 분석 수행
 */
function analyzeGeminiFile(fileUri, mimeType) {
  try {
    waitForFileProcessing(fileUri);

    const url = `https://generativelanguage.googleapis.com/v1beta/models/${MODEL_NAME}:generateContent?key=${GEMINI_API_KEY}`;
    
    const systemPrompt = `
      너는 베테랑 자동차 장기렌트 상담원이야. 통화 녹음 파일을 분석해서 JSON으로 반환해.
      
      [지침]
      1. 반드시 **JSON 포맷**만 반환할 것. (Markdown 없이)
      2. 불확실한 값은 null 또는 빈문자열.
      
      [JSON Key]
      customer_name, customer_phone, customer_type, gender, age_group,
      address_city, income_type, credit_info, occupation,
      interested_car_model, comparison_car_model, owned_car_model,
      car_usage, driving_distance, driver_scope,
      expected_contract_timing, desired_contract_term,
      desired_initial_cost_type, maintenance_service_level,
      summary, details
    `;

    const payload = {
      contents: [{
        parts: [
          { text: systemPrompt },
          { file_data: { mime_type: mimeType, file_uri: fileUri } }
        ]
      }],
      generationConfig: { response_mime_type: "application/json" }
    };

    const response = UrlFetchApp.fetch(url, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const code = response.getResponseCode();
    const text = response.getContentText();

    if (code !== 200) throw new Error(`[분석 요청 실패] 코드(${code}): ${text}`);
    
    let json;
    try { json = JSON.parse(text); } 
    catch (e) { throw new Error(`[분석 파싱 오류] JSON 아님: ${text.substring(0, 100)}`); }

    if (!json.candidates || !json.candidates[0] || !json.candidates[0].content) {
        throw new Error(`[분석 실패] 응답 없음: ${JSON.stringify(json.promptFeedback)}`);
    }

    let rawText = json.candidates[0].content.parts[0].text;
    rawText = rawText.replace(/```json/g, "").replace(/```/g, "").trim();
    
    return { success: true, data: JSON.parse(rawText) };

  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function waitForFileProcessing(fileUri) {
  const fileName = fileUri.split("/v1beta/")[1];
  const checkUrl = `https://generativelanguage.googleapis.com/v1beta/${fileName}?key=${GEMINI_API_KEY}`;
  
  for (let i = 0; i < 45; i++) {
    Utilities.sleep(2000);
    const response = UrlFetchApp.fetch(checkUrl, { muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) throw new Error("파일 확인 에러");
    
    const state = JSON.parse(response.getContentText()).state;
    if (state === "ACTIVE") return;
    if (state === "FAILED") throw new Error("Gemini 파일 처리 실패 (FAILED)");
  }
  throw new Error("파일 처리 시간 초과");
}
