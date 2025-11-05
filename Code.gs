// Code.gs

const ss = SpreadsheetApp.getActiveSpreadsheet();
const assignmentSheet = ss.getSheetByName("배정고객");
const logSheet = ss.getSheetByName("상담기록");
const contractsSheet = ss.getSheetByName("계약기록"); // [신규] Task 1
const configSheet = ss.getSheetByName("Config");

// --- Select Option Definitions ---
const SELECT_OPTIONS = {
  // A. 일반 선택 옵션들
  leadSource: ['홈페이지', '네이버검색', '다음검색', '구글검색', 'SNS 광고', '유튜브 광고', '소개', '기존고객 재렌트', '파트너사 전달', '오프라인 광고', '기타'],
  ageGroup: ['20대(만21~25)', '20대(만26~)', '30대', '40대', '50대', '60대 이상', '미확인'],
  incomeType: ['4대보험 직장인', '4대보험 제외 직장인', '전문직', '프리랜서', '일용직', '무직/학생', '미확인'],
  creditInfo: ['1~3 등급(상)', '4~6 등급(중)', '7~9 등급(하)', '심사 불가', '심사 전', '미확인'],
  carUsage: ['출퇴근용', '업무용(영업/출장)', '레저/패밀리용', '법인 임원용', '개인 사업장 운영', '기타', '미확인'],
  expectedContractTiming: ['즉시(1개월 내)', '3개월 내', '6개월 내', '미정'],
  desiredContractTerm: ['12개월', '24개월', '36개월', '48개월', '60개월','미정'],
  desiredInitialCostType: ['무보증', '보증금', '선납금', '보증증권', '미정'],
 
  maintenanceServiceLevel: ['미포함(Self)', '포함(기본)', '포함(고급)', '미정'],
  customerStatus: ['신규 문의', '견적 발송', '가망 고객', '심사 진행중', '심사 완료', '계약 진행중', '계약 완료', '출고 완료', '상담 보류', '상담 이탈', '기존 고객'],

  // B. 주소 데이터 (시/도는 전체 목록, 시/군/구는 맵 형태)
  addressCities: ['서울특별시', '부산광역시', '대구광역시', '인천광역시', '광주광역시', '대전광역시', '울산광역시', '세종특별자치시', '경기도', '강원특별자치도', '충청북도', '충청남도', '전북특별자치도', '전라남도', '경상북도', '경상남도', '제주특별자치도'],
  addressDistricts: {
    // 예시: 필요에 따라 실제 행정구역 목록으로 채워야 합니다.(양이 많을 수 있음)
    '서울특별시': ['강남구', '강동구', '강북구', '강서구', '관악구', '광진구', '구로구', '금천구', '노원구', '도봉구', '동대문구', '동작구', '마포구', '서대문구', '서초구', '성동구', '성북구', '송파구', '양천구', '영등포구', '용산구', '은평구', '종로구', '중구', '중랑구'],
    '경기도': ['수원시 장안구', '수원시 권선구', '수원시 팔달구', '수원시 영통구', '성남시 수정구', '성남시 중원구', '성남시 분당구', '의정부시', '안양시 만안구', '안양시 동안구', /* ... 다른 시/군 ... */ '가평군', '연천군'],
    '인천광역시': ['계양구', '미추홀구', '남동구', '동구', '부평구', '서구', '연수구', '중구', '강화군', '옹진군']
    // ... 다른 시/도에 대한 시/군/구 목록 추가 ...
  }
};

/**
 * [Issue #4 적용]
 * 시트의 헤더 이름을 기반으로 {headerName: 'A'} 형태의 맵을 생성하고 캐시합니다.
 * 헤더 행 자체를 해시하여 캐시 키를 생성하므로, 헤더 변경 시 캐시가 자동 갱신됩니다.
 * * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 대상 시트
 * @returns {Object} 예: {assignmentId: 'A', customerName: 'C', ...}
 */
function getHeaderColumnLetterMap_(sheet) {
  if (!sheet) { // [신규] 시트가 null일 경우 방어 코드
    // sheet.getName()을 호출할 수 없으므로,
    // 전역 변수와 비교하여 시트 이름을 추론합니다. (정확하지 않을 수 있음)
    let sheetName = '알 수 없는 시트';
    if (sheet === assignmentSheet) sheetName = "배정고객";
    if (sheet === logSheet) sheetName = "상담기록";
    if (sheet === contractsSheet) sheetName = "계약기록";
    if (sheet === configSheet) sheetName = "Config";
    
    logError_('getHeaderColumnLetterMap_FATAL', new Error('시트 객체가 null입니다.'), { 
      sheetName: sheetName 
    });
    throw new Error(`시트 객체를 찾을 수 없습니다. (시트 이름: ${sheetName}). 시트가 삭제되었거나 Code.gs 상단 변수 이름이 잘못되었을 수 있습니다.`);
  }

  const sheetId = sheet.getSheetId();
  const cache = CacheService.getScriptCache();
  
  // 1. 헤더 행을 직접 읽어옵니다. (Gviz보다 빠름)
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  // 2. 헤더를 기반으로 해시(Hash) 키를 생성합니다.
  // MD5로도 충분하며, 8바이트로 잘라서 짧게 만듭니다.
  const headerHash = Utilities.computeDigest(
    Utilities.DigestAlgorithm.MD5,
    headers.join(',') // 헤더 배열을 쉼표로 구분된 문자열로 변환
  ).slice(0, 8).map(b => (b & 0xFF).toString(16).padStart(2, '0')).join('');

  // 3. 해시를 포함한 캐시 키를 사용합니다.
  const cacheKey = `header_col_map_v3_${sheetId}_${headerHash}`;

  const cachedMap = cache.get(cacheKey);
  if (cachedMap) {
      return JSON.parse(cachedMap);
  }

  // 4. 캐시가 없으면 맵을 생성합니다.(기존 로직 동일)
  const colMap = {};
  headers.forEach((header, i) => {
    if (header) {
      // 0-based index (i)를 1-based index (i + 1)로 변환
      let colLetter = '';
      let n = i + 1;
      while (n > 0) {
        const remainder = (n - 1) % 26;
        colLetter = String.fromCharCode(65 + remainder) + colLetter;
        n 
= Math.floor((n - 1) / 26);
      }
      colMap[header] = colLetter;
    }
  });
  cache.put(cacheKey, JSON.stringify(colMap), 3600); // 1시간 캐시
  return colMap;
}

// --- 헬퍼 함수 ---

function escapeQueryString_(str) {
  if (typeof str !== 'string') return str;
  return str.replace(/'/g, "\\'").replace(/\\/g, "\\\\");
}

// [수정] 4번 제안: Lock 타임아웃 연장 및 지수 백오프 적용
function acquireLockWithRetry_(operationName = "작업") {
  const lock = LockService.getScriptLock();
  const maxRetries = 5; // 3 → 5
  const baseSleep = 2000;
  // 1000 → 2000

  for (let i = 0; i < maxRetries; i++) {
    if (lock.tryLock(10000)) { // 5000 → 10000 (10초)
      return lock;
    }

    if (i < maxRetries - 1) {
      const jitter = Math.random() * 1000;
      const sleepTime = baseSleep * (i + 1) + jitter;
      // 지수 백오프
      Logger.log(`Lock 대기 중... (${i + 1}/${maxRetries}) - ${operationName}`);
      Utilities.sleep(sleepTime);
    }
  }

  // [수정] 에러 메시지에 operationName 포함
  throw new Error(`시스템이 혼잡합니다 (${operationName} 작업). 1분 후 다시 시도해주세요.`);
}

function measurePerformance_(fnName, fn) {
  const start = new Date().getTime();
  try {
    return fn();
  } finally {
    const elapsed = new Date().getTime() - start;
    if (elapsed > 3000) { 
      Logger.log(`[PERF_WARNING] ${fnName} took ${elapsed}ms`);
    }
  }
}

function formatPhoneNumber(phone) {
  if (!phone) return "";
  const digits = phone.replace(/\D/g, "");
  if (!digits) return "";
  return "'" + digits; 
}

// [Code.gs] - getConfigurations 함수를 아래 코드로 교체합니다.
function getConfigurations() {
  Logger.log("⚙️ [Debug 1/6] getConfigurations started."); // 1단계
  
  try {
    const cache = CacheService.getScriptCache();
    Logger.log("⚙️ [Debug 2/6] CacheService obtained."); // 2단계

    // configSheet 변수가 유효한지 다시 확인
    if (!configSheet) {
      Logger.log("❌ [Debug FATAL] 'configSheet' is null. Code.gs 상단의 시트 이름을 다시 확인하세요.");
      throw new Error("'Config' 시트를 찾을 수 없습니다. (Code.gs 상단 변수 확인)");
    }
    Logger.log("⚙️ [Debug 3/6] 'configSheet' variable is valid.");
    // 3단계

    // A1 셀의 메모(Note) 읽기 시도
    const configVersion = configSheet.getRange("A1").getNote() || "v1.0.0";
    Logger.log("⚙️ [Debug 4/6] Got configVersion: " + configVersion); // 4단계
    
    const cacheKey = `config_data_v4_${configVersion}`;
    const cachedConfig = cache.get(cacheKey); 

    if (cachedConfig) {
      Logger.log("⚙️ [Debug 5/6] Cache HIT. (캐시에서 데이터를 반환합니다)");
      // 캐시가 있으면 여기서 실행이 (성공적으로) 종료됨
      return JSON.parse(cachedConfig);
    }
    Logger.log("⚙️ [Debug 5/6] Cache MISS. (새 데이터를 빌드합니다)");
    // 5단계

    // measurePerformance_ 래퍼 내부 실행
    return measurePerformance_('getConfigurations_CacheMiss', () => {
      Logger.log("⚙️ [Debug 6/6] measurePerformance_ block started."); // 6단계
      
      const data = configSheet.getDataRange().getValues();
      Logger.log("⚙️ [Debug 7/6] configSheet.getDataRange() successful."); // 7단계
      
      const headers = data.shift();
      if (!headers || headers.length === 0) {
        Logger.log("❌ [Debug FATAL] 'Config' sheet is empty or has no header(1행) row.");
        throw new Error("'Config' 시트가 비어있거나 헤더(1행)가 없습니다.");
      }
      Logger.log("⚙️ [Debug 8/6] Headers processed."); // 8단계

      const salespersonEmails = new Set();
      const dbTypes = new Set();
      const emailToNameMap = {}; 
      const emailIndex = headers.indexOf('SalespersonEmail');
      const nameIndex = headers.indexOf('SalespersonName'); 
      const dbTypeIndex = headers.indexOf('DbType');

      if 
 (emailIndex === -1) {
          Logger.log("❌ [Debug FATAL] 'Config' 시트 1행에 'SalespersonEmail' 헤더가 없습니다.");
          throw new Error("'Config' 시트 1행에 'SalespersonEmail' 헤더가 없습니다.");
      }
      Logger.log("⚙️ [Debug 9/6] Header indexes found.");
      // 9단계

      data.forEach(row => {
        const email = row[emailIndex];
        const dbType = row[dbTypeIndex];
        
        if (email) {
          salespersonEmails.add(email);
          const name = (nameIndex > -1 && row[nameIndex]) ? row[nameIndex] : email; 
          emailToNameMap[email] = name;
       
         }
        if (dbType) dbTypes.add(dbType);
      });
      Logger.log("⚙️ [Debug 10/6] Data iteration complete."); // 10단계

      const optionsForClient = { ...SELECT_OPTIONS };
      delete optionsForClient.addressDistricts; 
      Logger.log("⚙️ [Debug 11/6] SELECT_OPTIONS processed."); // 11단계

      const config = {
        salespersonEmails: [...salespersonEmails],
        dbTypes: [...dbTypes].sort(),
        emailToNameMap: emailToNameMap,
        selectOptions: optionsForClient 
      };
      cache.put(cacheKey, JSON.stringify(config), 3600); // 1시간
      Logger.log("✅ [Debug FINAL] getConfigurations finished successfully.");
      // 최종
      return config;
    });
  } catch (e) {
    // try 블록 전체에서 오류 발생 시 이 로그가 찍힘
    Logger.log(`❌ [Debug CATCH] getConfigurations FAILED: ${e.message}`);
    Logger.log(e.stack); // 오류 상세 스택
    throw e;
    // 오류를 클라이언트로 다시 던져서 UI에 실패가 표시되도록 함
  }
}

function findRowById_(sheetName, idColumnName, idToFind) {
  return measurePerformance_(`findRowById_(${idToFind})`, () => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error(`시트 '${sheetName}'을 찾을 수 없습니다.`);

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return null;

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const idColIndex = headers.indexOf(idColumnName);
    if (idColIndex === -1) {
      throw new Error(`'${idColumnName}' 열을 '${sheetName}' 시트에서 찾을 수 없습니다.`);
    }

  
    // --- 개선된 부분 ---
    // 1. ID 컬럼 전체 값을 한 번에 읽어옵니다.
    const idColumnValues = sheet.getRange(2, idColIndex + 1, lastRow - 1, 1)
                              .getValues()
                              .flat(); // 2D 배열을 1D 배열로 변환

  
    // 2. 메모리(JS)에서 인덱스를 찾습니다. (createTextFinder보다 훨씬 빠름)
    const rowIndexInArray = idColumnValues.indexOf(idToFind);

    if (rowIndexInArray > -1) {
      // 3. 실제 시트의 행 번호를 계산합니다.(배열은 0부터 시작, 시트는 1부터, 헤더 1줄 제외)
      const rowNum = rowIndexInArray + 2;
      const rowValues = sheet.getRange(rowNum, 1, 1, headers.length).getValues()[0];
      // --- 개선 끝 ---

      const rowData = {};
      headers.forEach((header, j) => {
        let value = rowValues[j];
        if (value instanceof Date) {
          value = value.toISOString();
        }
        rowData[header] = value;
      });
      return {
        rowData: rowData,
        rowNum: rowNum,
        rowValues: rowValues,
        headers: headers
      };
    }
    return null;
  });
}

// [Code.gs] - findLogsByAssignmentId_ 대체

function findLogsByAssignmentId_(assignmentId) {
  return measurePerformance_(`findLogsByAssignmentId_(${assignmentId})`, () => {
    if (!logSheet) throw new Error("시트 '상담기록'을 찾을 수 없습니다.");

    // '상담기록' 시트의 동적 열 맵 가져오기
    const COLS_LOG = getHeaderColumnLetterMap_(logSheet);

    if (!COLS_LOG.assignmentId || !COLS_LOG.logTimestamp || !COLS_LOG.logId) {
        throw new Error("'상담기록' 시트 1행에 'assignmentId', 'logTimestamp' 또는 'logId' 헤더가 없습니다.");
    }

    let confirmedLogs = []; // Gviz로 가져온 "확정된" 로그 (시트에 저장된 로그)
    let 
pendingLogs = [];   // PropertiesService 큐에서 가져온 "대기 중인" 로그

    // --- 1. Gviz로 "확정된" 로그 가져오기 (기존 로직) ---
    try {
      const logSheetGid = logSheet.getSheetId();
      const spreadsheetId = ss.getId();
      const query = `SELECT * WHERE ${COLS_LOG.assignmentId} = '${escapeQueryString_(assignmentId)}' ORDER BY ${COLS_LOG.logTimestamp} DESC`;

      const tqUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/gviz/tq?gid=${logSheetGid}&tq=${encodeURIComponent(query)}&headers=1`;
      const token = ScriptApp.getOAuthToken();
      const response = UrlFetchApp.fetch(tqUrl, {
        headers: { 'Authorization': 'Bearer ' + token }
      });
      const jsonResponse = JSON.parse(response.getContentText().match(/google\.visualization\.Query\.setResponse\(([\s\S\w]+)\);/)[1]);

      if (jsonResponse.status === 'error') {
        throw new Error(`Gviz API 오류 (상담기록): ${jsonResponse.errors[0].detailed_message}`);
      }

      const headers = jsonResponse.table.cols.map(col => col.label || col.id);
      const headerMap = {};
      headers.forEach((h, i) => headerMap[h] = i);
      
      const timestampHeaderName = 'logTimestamp';
      confirmedLogs = jsonResponse.table.rows.map(row => {
        const logObj = {};
        headers.forEach(header => {
          const index = headerMap[header];
          let cell = row.c[index];
          let value = null;
          if (cell) {
            if (header === timestampHeaderName && cell.v) {
      
         const parsedDate = parseGvizDateObject_(cell.v);
              value = parsedDate ? parsedDate.toISOString() : (cell.f || cell.v);
            }
            else if (cell.f) { value = cell.f; } 
            else if (cell.v !== null && cell.v !== undefined) { value = cell.v; }
          
 }
          logObj[header] = value;
        });
        return logObj;
      });
    } catch (e) {
      // Gviz가 실패해도 큐는 읽어야 하므로 throw하지 않고 에러만 로깅합니다.
      logError_('findLogsByAssignmentId_Gviz', e, { assignmentId: assignmentId });
      // Gviz 쿼리 자체에 문제가 생겼음을 클라이언트에 알릴 필요가 있다면 throw
      // throw new Error("상담 기록(Gviz) 조회 중 오류가 발생했습니다.");
    }

    // --- 2. [신규] PropertiesService 큐에서 "대기 중인" 로그 가져오기 ---
    try {
        const scriptProperties = PropertiesService.getScriptProperties();
        // ★ BatchWorker.gs의 우회로(getProperties() + filter)를 동일하게 사용
        const allProperties = scriptProperties.getProperties();
        const logKeys = Object.keys(allProperties).filter(k => k.startsWith('log_queue_'));

        if (logKeys.length > 0) {
            const pendingLogsMap = {};
// (중복 방지를 위해 Map 사용)

            logKeys.forEach(key => {
                try {
                    const logDataString = allProperties[key];
                    if (!logDataString) return;
                   
   
                    const logData = JSON.parse(logDataString);
                    
                    // [중요] 현재 조회하려는 assignmentId와 일치하는 로그만 필터링
                    if (logData.assignmentId === assignmentId) {
     
                        // Gviz가 반환하는 형식과 동일하게 맞춥니다.
                        const logForClient = {
                            logId: logData.logId,
                    
         assignmentId: logData.assignmentId,
                            logTimestamp: logData.logTimestamp, // ISO 문자열
                            logContent: logData.logContent,
                            userName: logData.userName
 
                        };
                        pendingLogsMap[logData.logId] = logForClient;
                    }
                } catch (e) {
            
                    // 개별 로그 파싱 오류는 무시 (ErrorLog에 남겨도 좋음)
                    logError_('findLogs_QueueParse', e, { propertyKey: key });
                }
            });
            pendingLogs = Object.values(pendingLogsMap);
        }
    } catch (e) {
        logError_('findLogsByAssignmentId_QueueRead', e, { assignmentId: assignmentId });
        // 이 작업이 실패해도 Gviz 로그는 반환해야 하므로 throw하지 않습니다.
    }

    // --- 3. [신규] 두 로그 병합 및 정렬 ---
    
    // Gviz 로그 ID 맵을 만들어 큐에 있는 로그가 이미 Gviz에 있는지(시트에 저장됐는지) 확인
    const confirmedLogIds = new Set(confirmedLogs.map(log => log.logId));
    // Gviz에 없는 "대기 중인" 로그만 필터링
    const uniquePendingLogs = pendingLogs.filter(pLog => !confirmedLogIds.has(pLog.logId));
    const combinedLogs = [...uniquePendingLogs, ...confirmedLogs];

    // 최종적으로 시간 역순 정렬
    combinedLogs.sort((a, b) => new Date(b.logTimestamp) - new Date(a.logTimestamp));
    return combinedLogs;
  });
}

/**
 * [신규] Task 4 헬퍼
 * 특정 assignmentId에 연결된 모든 계약을 '계약기록' 시트에서 조회합니다.
 */
function findContractsByAssignmentId_(assignmentId) {
  return measurePerformance_(`findContractsByAssignmentId_(${assignmentId})`, () => {
    if (!contractsSheet) {
      logError_('findContractsByAssignmentId_', new Error("시트 '계약기록'을 찾을 수 없습니다."), { assignmentId: assignmentId });
      return []; // [수정] 오류 발생 시 빈 배열 반환 (고객 정보는 로드되도록)
    }

    const COLS_CONTRACT = getHeaderColumnLetterMap_(contractsSheet);
    if (!COLS_CONTRACT.assignmentId || !COLS_CONTRACT.contractDate) {
      logError_('findContractsByAssignmentId_', new Error("'계약기록' 시트 1행에 'assignmentId' 또는 'contractDate' 헤더가 없습니다."), { assignmentId: assignmentId });
      return []; // [수정] 오류 발생 시 빈 배열 반환
    }

    let contracts = [];
    try {
      const contractSheetGid = contractsSheet.getSheetId();
      const spreadsheetId = ss.getId();
      // contractDate가 최신인 순으로 정렬
      const query = `SELECT * WHERE ${COLS_CONTRACT.assignmentId} = '${escapeQueryString_(assignmentId)}' ORDER BY ${COLS_CONTRACT.contractDate} DESC`;

      const tqUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/gviz/tq?gid=${contractSheetGid}&tq=${encodeURIComponent(query)}&headers=1`;
      const token = ScriptApp.getOAuthToken();
      const response = UrlFetchApp.fetch(tqUrl, {
        headers: { 'Authorization': 'Bearer ' + token }
      });
      const jsonResponse = JSON.parse(response.getContentText().match(/google\.visualization\.Query\.setResponse\(([\s\S\w]+)\);/)[1]);

      if (jsonResponse.status === 'error') {
        throw new Error(`Gviz API 오류 (계약기록): ${jsonResponse.errors[0].detailed_message}`);
      }

      const headers = jsonResponse.table.cols.map(col => col.label || col.id);
      const headerMap = {};
      headers.forEach((h, i) => headerMap[h] = i);
      
      const dateHeaders = ['contractDate']; // 계약기록 시트의 날짜 열

      contracts = jsonResponse.table.rows.map(row => {
        const contractObj = {};
        headers.forEach(header => {
          const index = headerMap[header];
          let cell = row.c[index];
          let value = null;
          if (cell) {
            if (dateHeaders.includes(header) && cell.v) {
              const parsedDate = parseGvizDateObject_(cell.v);
              value = parsedDate ? parsedDate.toISOString() : (cell.f || cell.v);
            }
            else if (cell.f) { value = cell.f; } 
            else if (cell.v !== null && cell.v !== undefined) { value = cell.v; }
          }
          contractObj[header] = value;
        });
        return contractObj;
      });
    } catch (e) {
      logError_('findContractsByAssignmentId_Gviz', e, { assignmentId: assignmentId });
      // [수정] 오류가 발생해도 빈 배열을 반환 (고객 정보는 로드되도록)
      // throw new Error("계약 기록(Gviz) 조회 중 오류가 발생했습니다.");
    }
    return contracts;
  });
}


function bustUserListCache_() {
  try {
    const userProperties = PropertiesService.getUserProperties();
    const currentVersion = parseInt((userProperties.getProperty('DATA_VERSION') || '1'), 10);
    userProperties.setProperty('DATA_VERSION', (currentVersion + 1).toString());
    Logger.log(`Busted user list cache. New version: ${currentVersion + 1}`);
  } catch (e) {
    Logger.log(`Failed to bust user cache: ${e.message}`);
  }
}

// [신규] 7번 제안: 중앙 집중식 에러 로깅 함수
function logError_(context, error, additionalInfo = {}) {
  try {
    const errorLog = {
      timestamp: new Date().toISOString(),
      context: context,
      error: error.message,
      stack: error.stack ?
 error.stack : 'No stack trace available',
      user: Session.getActiveUser().getEmail(),
      additionalInfo: additionalInfo
    };
    const errorString = `[ERROR] ${JSON.stringify(errorLog)}`;
    Logger.log(errorString); // Apps Script 기본 로거에도 기록

    // 'ErrorLog' 시트에 기록
    const errorSheet = ss.getSheetByName("ErrorLog");
    if (errorSheet) {
      // 헤더 순서: Timestamp, User, Context, Error, Stack, Info(JSON)
      errorSheet.appendRow([
        errorLog.timestamp,
        errorLog.user,
        errorLog.context,
        errorLog.error,
        errorLog.stack,
        JSON.stringify(errorLog.additionalInfo)
      ]);
    }
  } catch (e) {
    // 에러 로깅 함수 자체에서 오류가 날 경우
    Logger.log(`[FATAL_LOGGING_ERROR] 에러 로깅 실패: ${e.message}`);
    Logger.log(`[ORIGINAL_ERROR] ${context}: ${error.message}`);
  }
}

// --- Main Functions ---

// -----------------------------------------------------------------
// ✨ [수정] doGet - 'isDualMode' 파라미터 추가
// -----------------------------------------------------------------
function doGet(e) { 
  const userEmail = Session.getActiveUser().getEmail();
  const configData = getConfigurations();
  if (!configData.salespersonEmails.includes(userEmail)) {
    return HtmlService.createHtmlOutput('<h1>접근 권한이 없습니다. 관리자에게 문의하세요.</h1>');
  }

  // ✨ [수정] createTemplateFromFile 대신 createHtmlOutput 사용
    const t = HtmlService.createTemplateFromFile('index');
    t.view = e.parameter.view ||
 'list'; 
    t.isDualMode = e.parameter.dual === 'true';

    // ▼▼▼ [신규] 스크립트의 /exec URL을 템플릿에 전달 ▼▼▼
    t.scriptUrl = ScriptApp.getService().getUrl();
    // ▲▲▲ [신규] ▲▲▲

    // ✨ [수정] 템플릿을 먼저 evaluate() 하여 HTML 문자열을 가져옵니다.
    const htmlContent = t.evaluate().getContent();

    // ✨ [수정] 최종 HTML을 생성하고 X-Frame-Options을 설정합니다.
    return HtmlService.createHtmlOutput(htmlContent)
        .setTitle('고객 상담일지 기록시스템')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
        // (선택 사항: 듀얼 모드 등에서 iframe 임베딩 시 필요할 수 있음)
}

/**
 * [수정된 메인 함수]
 * 기존 'COLS' 전역 상수 대신, 동적으로 생성된 COLS 맵을 사용합니다.
 */
function getAssignedCustomers(filters) {
  return measurePerformance_('getAssignedCustomers', () => {
    // [수정] 함수가 호출될 때마다 동적 맵을 (캐시에서) 불러옵니다.
    const COLS = getHeaderColumnLetterMap_(assignmentSheet);

    // [추가] 필수 열이 맵에 존재하는지 확인 (안정성)
    // [수정] Task 2: 'contractStatus'가 '배정고객' 시트에서 제거되었으므로, COLS 검사에서도 제거합니다.
    const requiredCols = ['assignedTo', 'assignmentDate', 'consultationStatus', 'dbType', 'customerName', 'customerPhoneNumber'];
    for (const col of requiredCols) {
        if (!COLS[col]) {
            throw new Error(`'배정고객' 시트 1행에서 필수 헤더 '${col}'를 찾을 수 없습니다. (열 이름이 변경되었는지 확인하세요)`);
    
         }
    }

    const userEmail = Session.getActiveUser().getEmail();
    const configData = getConfigurations();
    const userName = configData.emailToNameMap[userEmail] || userEmail; 
    
    const cache = CacheService.getUserCache();
    const offset = filters.offset || 0;
    const limit = filters.limit || 30;
    let cacheKey;
    
    if (offset === 0) {
      cacheKey = getDefaultCacheKey_(userEmail, filters); 
      const cachedResult = cache.get(cacheKey);
      if (cachedResult) {
        Logger.log('캐시 히트: ' + userName + " | key: " + cacheKey);
        try {
          const parsedResult = JSON.parse(cachedResult);
          if (parsedResult && parsedResult.customers) { 
            return parsedResult;
          } else {
            cache.remove(cacheKey);
          }
        // [수정]
        } catch (e) {
          // Logger.log("QUERY 실패: " + e.message + " | 쿼리: " + queryString);
// [삭제]
          logError_('getAssignedCustomers_Gviz', e, { query: queryString });
// [추가]
          throw new Error("데이터를 조회하는 중 오류가 발생했습니다. (QUERY 실패)");
        }
      }
      Logger.log('캐시 미스: ' + userName + " | key: " + cacheKey);
    } else {
      Logger.log('캐시 스킵 (offset > 0): ' + userName);
    }

    Logger.log('QUERY 실행: ' + userName + ", " + JSON.stringify(filters));
    // escapeQueryString_() 헬퍼 함수를 사용하여 userName을 이스케이프합니다.
    let queryString = `SELECT * WHERE ${COLS.assignedTo} = '${escapeQueryString_(userName)}'`;
    if (filters.dateFrom) {
      const dateFromStr = Utilities.formatDate(new Date(filters.dateFrom), "GMT+9", "yyyy-MM-dd");
      queryString += ` AND ${COLS.assignmentDate} >= DATE '${dateFromStr}'`;
    }
    
    if (filters.dateTo) {
      const dateToObj = new Date(filters.dateTo);
      dateToObj.setDate(dateToObj.getDate() + 1);
      const dateToStr = Utilities.formatDate(dateToObj, "GMT+9", "yyyy-MM-dd");
      queryString += ` AND ${COLS.assignmentDate} < DATE '${dateToStr}'`;
    }
    
    if (filters.consultStatus) {
      queryString += ` AND ${COLS.consultationStatus} = '${escapeQueryString_(filters.consultStatus)}'`;
    }
    
    // [수정] Task 2: 'contractStatus' 필터는 '배정고객' 시트에 더 이상 존재하지 않습니다.
    // ※참고: 이 필터가 UI에서 제거되지 않으면 Gviz 쿼리 오류가 발생합니다.
    //      (일단 서버에서는 이 필터링을 제거합니다.)
    if (filters.contractStatus) {
      // queryString += ` AND ${COLS.contractStatus} = '${escapeQueryString_(filters.contractStatus)}'`; // [삭제]
      Logger.log(`[WARN] 'contractStatus' 필터는 '배정고객' 시트에서 제거되어 무시됩니다: ${filters.contractStatus}`);
    }
    
    if (filters.dbType) {
      queryString += ` AND ${COLS.dbType} = '${escapeQueryString_(filters.dbType)}'`;
    }

    if (filters.searchTerm) {
      // ✨ [Issue #2 적용] LOWER() 및 matches '.*...*' 대신 CONTAINS 사용
      
      // 1. searchTerm도 소문자, 하이픈 제거 (미리 저장된 형식과 맞춤)
      const term = escapeQueryString_(filters.searchTerm.toLowerCase().trim().replace(/-/g, ''));
      // 2. CONTAINS 쿼리로 변경 (훨씬 빠름)
      if (term) { // 빈 문자열이 아닐 때만 쿼리 추가
        queryString += ` AND ${COLS.SearchHelper} CONTAINS '${term}'`;
      }
    }

    queryString += ` ORDER BY ${COLS.assignmentDate} DESC`;
    queryString += ` LIMIT ${limit + 1}`;
    if (offset > 0) {
      queryString += ` OFFSET ${offset}`;
    }

    let data, headers;
    try {
      const spreadsheetId = ss.getId();
      const gid = assignmentSheet.getSheetId();
      const tqUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/gviz/tq?gid=${gid}&tq=${encodeURIComponent(queryString)}&headers=1`;
      
      const token = ScriptApp.getOAuthToken();
      const response = UrlFetchApp.fetch(tqUrl, {
        headers: { 'Authorization': 'Bearer ' + token }
      });
      const jsonResponse = JSON.parse(response.getContentText().match(/google\.visualization\.Query\.setResponse\(([\s\S\w]+)\);/)[1]);

      if (jsonResponse.status === 'error') {
        throw new Error(`Visualization API 오류: ${jsonResponse.errors[0].detailed_message}`);
      }
      
      headers = jsonResponse.table.cols.map(col => col.label || col.id);
      const dateHeaders = ['assignmentDate', 'lastLogDate']; // 'contractDate'는 '배정고객' 시트에서 제거됨
      
      data = jsonResponse.table.rows.map(row => {
        return row.c.map((cell, index) => { 
          
          if (!cell) return null;
          
          const headerName = headers[index];
          
          if (cell.v && dateHeaders.includes(headerName)) {
       
           const parsedDate = parseGvizDateObject_(cell.v); 
            if (parsedDate) {
              return parsedDate; // Date 객체 자체를 반환
            }
          }

          if (cell.f && !dateHeaders.includes(headerName)) {
            return cell.f;
        
         }

          if (cell.v !== null && cell.v !== undefined) {
            return cell.v;
          }

          return null;
        });
      });
    } catch (e) {
      Logger.log("QUERY 실패: " + e.message + " | 쿼리: " + queryString);
      throw new Error("데이터를 조회하는 중 오류가 발생했습니다. (QUERY 실패)");
    }

    let hasMore = false;
    if (data.length > limit) {
      hasMore = true;
      data.pop();
    }
    
    const headerMap = {};
    headers.forEach((header, i) => headerMap[header] = i);
    const customers = data.map(row => {
      const obj = {};
      for (const header in headerMap) {
        const index = headerMap[header];
        let value = row[index];
        if (value && value instanceof Date) {
          obj[header] = value.toISOString();
        } else {
          obj[header] = value;
     
         }
      }
      return obj;
    });
    const totalCount = offset + customers.length + (hasMore ? 1 : 0);
    const result = {
      customers: customers,
      totalCount: totalCount
    };
    if (offset === 0 && cacheKey) {
      try {
        cache.put(cacheKey, JSON.stringify(result), 180);
        // 3분
        Logger.log('캐시 저장: ' + cacheKey);
      } catch (e) {
        Logger.log('캐시 저장 실패 (데이터 크기 초과): ' + e.message);
      }
    }

    return result;
  });
  // End of measurePerformance_
}

// [수정] Task 4: 1:N 계약 목록을 함께 조회합니다.
function getAssignmentDetails(assignmentId) {
  return measurePerformance_(`getAssignmentDetails(${assignmentId})`, () => {
    const found = findRowById_("배정고객", "assignmentId", assignmentId);
    if (!found) {
      throw new Error("해당 배정 ID의 고객을 찾을 수 없습니다.");
    }

    const cache = CacheService.getUserCache();
    cache.put(`rowNum_${assignmentId}`, found.rowNum, 3600); 

    const logs = findLogsByAssignmentId_(assignmentId); 
    const contracts = findContractsByAssignmentId_(assignmentId); // [신규]

    return {
      assignment: found.rowData,
      logs,
      contracts // [신규]
    };
  });
  // End of measurePerformance_
}

// [개선된 saveDetails 함수 제안]
function saveDetails(assignmentId, detailsFromClient) {
  return measurePerformance_(`saveDetails(${assignmentId})`, () => {
    
    // [수정] 작업 이름 전달
    const lock = acquireLockWithRetry_(`saveDetails: ${assignmentId}`); 
    let updatedRowDataForClient;
    let finalLogs;
    let finalContracts; // [신규] Task 4
    
    try {
      // --- 1. 읽기 (Lock 내부) ---
      // Lock을 획득한 직후의 최신 데이터를 읽어옵니다.
      const found = findRowById_("배정고객", "assignmentId", assignmentId);
      if (!found) throw new Error("데이터를 찾는 중 오류 발생 (Lock 내부)");
      
      const { rowNum, headers, rowValues } = found;
      const valuesToUpdate = [...rowValues]; // ★ 최신 원본 데이터를 복사

      // --- 2. 수정 (Lock 내부) ---
      // 클라이언트가 보낸 변경 사항(detailsFromClient)만 최신 원본에 덮어씁니다.
      Object.keys(detailsFromClient).forEach(key => {
        const colIndex = headers.indexOf(key);
        if (colIndex > -1) {
   
           let value = detailsFromClient[key];
          if (key === 'customerPhoneNumber') {
            value = formatPhoneNumber(value);
          }
          valuesToUpdate[colIndex] = value;
        }
      });
      // --- 3. 쓰기 (Lock 내부) ---
      assignmentSheet.getRange(rowNum, 1, 1, headers.length).setValues([valuesToUpdate]);
      SpreadsheetApp.flush();
      // --- 4. 반환할 데이터 준비 (Lock 내부) ---
      // 시트를 다시 읽을 필요 없이, 방금 수정한 배열을 객체로 변환합니다.
      const updatedRowData = {};
      headers.forEach((header, j) => {
        let value = valuesToUpdate[j];
        if (value instanceof Date) {
          value = value.toISOString(); // 클라이언트 반환을 위해 ISO 문자열로
        }
        updatedRowData[header] = value;
      });
      updatedRowDataForClient = updatedRowData;

    } finally {
      lock.releaseLock();
    }

    // --- 5. 후속 작업 (Lock 외부) ---
    bustUserListCache_();
    // 로그는 Lock 외부에서 조회 (Gviz 쿼리는 Lock이 필요 없음)
    finalLogs = findLogsByAssignmentId_(assignmentId);
    finalContracts = findContractsByAssignmentId_(assignmentId); // [신규] Task 4

    return {
      assignment: updatedRowDataForClient,
      logs: finalLogs,
      contracts: finalContracts // [신규] Task 4
    };
  });
}

// --- Data Writing Functions (with LockService) ---

function saveCustomerDetails(assignmentId, details) {
  const userEmail = Session.getActiveUser().getEmail();
  const configData = getConfigurations();
  const userName = configData.emailToNameMap[userEmail] || userEmail;
  const cache = CacheService.getUserCache();
  // --- 1. findRowById_를 항상 호출하여 rowNum과 headers를 확보합니다. (권한 검사 포함) ---
  // (위 2번 항목에서 findRowById_가 최적화되었기 때문에 부담이 적습니다)
  const found = findRowById_("배정고객", "assignmentId", assignmentId);
  if (!found) throw new Error("해당 배정 ID를 찾을 수 없습니다.");
  if (found.rowData.assignedTo !== userName) { 
    throw new Error("본인에게 배정된 고객만 수정할 수 있습니다.");
  }
  
  // rowNum 캐시는 findRowById_ 내부에서 처리하거나, 여기서 put 처리
  cache.put(`rowNum_${assignmentId}`, found.rowNum, 3600);
  
  // --- 2. EDITABLE_FIELDS 필터링 ---
  // [수정] Task 5: 'contractStatus' 제거
  const EDITABLE_FIELDS = [
    // 기본 정보
    'leadSource', 'customerType', 'gender', 'ageGroup', 'addressCity', 'addressDistrict',
    'incomeType', 'creditInfo',
    // 고객 니즈
    'interestedCarModel', 'comparisonCarModel', 'ownedCarModel', 'carUsage', 'driverScope',
    // 희망 계약 조건
    'expectedContractTiming', 'desiredContractTerm', 'desiredInitialCostType',
    'maintenanceServiceLevel', 'salesCondition', 'isRepurchase', 'paymentMethod',
    // 상담/관리 상태
    'customerStatus', 'consultationStatus',
    // 'contractStatus', // [삭제] Task 5
    // 연락처 (별도 저장 버튼이 있지만, 여기서도 허용 가능)
    'customerPhoneNumber',
    // ▼▼▼ [신규] 고객 메모 필드 추가 ▼▼▼
    'customerMemo'
  ];
  const sanitizedDetails = {};
  EDITABLE_FIELDS.forEach(key => {
    if (details[key] !== undefined) {
      sanitizedDetails[key] = details[key];
    }
  });
  
  // --- 3. saveDetails 호출 ---
  // saveDetails는 이제 'found' 객체(rowNum, headers, rowValues 포함)를
  // 인자로 받아 Lock 범위 최소화 로직을 수행하도록 수정하는 것이 이상적입니다.
  // (여기서는 기존 saveDetails 구조를 따른다고 가정)
  return saveDetails(assignmentId, sanitizedDetails);
}

// [삭제] Task 5: saveContractDetails 함수 전체 삭제
/*
function saveContractDetails(assignmentId, details) {
  ... (기존 함수 내용) ...
}
*/


function addConsultationLog(assignmentId, logContent) {
  return measurePerformance_(`addConsultationLog_Batch(${assignmentId})`, () => {
    const logTimestamp = new Date();
    const userEmail = Session.getActiveUser().getEmail();
    const configData = getConfigurations();
    const userName = configData.emailToNameMap[userEmail] || userEmail;
    const cache = CacheService.getUserCache();

    // --- 1. 헤더와 인덱스를 함수 시작 시 한 번만 읽기 ---
    const headers = assignmentSheet.getRange(1, 1, 1, assignmentSheet.getLastColumn()).getValues()[0];
    const assignedToColIndex = headers.indexOf('assignedTo');
    const lastLogDateColIndex = headers.indexOf('lastLogDate');

    if (assignedToColIndex === -1 || lastLogDateColIndex === -1) {
  
       throw new Error("'assignedTo' 또는 'lastLogDate' 열을 '배정고객' 시트에서 찾을 수 없습니다.");
    }

    // --- 2. 읽기 및 권한 확인 (Lock 외부) ---
    let rowNumToUpdate;
    let foundAssignmentData = null;
    const cachedRowNum = cache.get(`rowNum_${assignmentId}`);

    if (cachedRowNum) {
      rowNumToUpdate = parseInt(cachedRowNum, 10);
      const assignedToValue = assignmentSheet.getRange(rowNumToUpdate, assignedToColIndex + 1).getValue(); 
      if (assignedToValue !== userName) {
        throw new Error("본인에게 배정된 고객의 기록만 추가할 수 있습니다.");
      }
    } else {
      const found = findRowById_("배정고객", "assignmentId", assignmentId);
      if (!found) throw new Error("해당 배정 ID를 찾을 수 없습니다.");
      if (found.rowData.assignedTo !== userName) { 
        throw new Error("본인에게 배정된 고객의 기록만 추가할 수 있습니다.");
      }
      rowNumToUpdate = found.rowNum;
      foundAssignmentData = found.rowData; 
      cache.put(`rowNum_${assignmentId}`, rowNumToUpdate, 3600);
    }

    // --- 3. [변경] 쓰기 작업 (Lock 내부) ---
    // appendRow()가 제거되었으므로 Lock 범위가 매우 짧아집니다.
    // [수정] 작업 이름 전달
    const lock = acquireLockWithRetry_(`addLog: ${assignmentId}`);
    try {
      // 'lastLogDate' 업데이트 (빠른 작업)
      assignmentSheet.getRange(rowNumToUpdate, lastLogDateColIndex + 1).setValue(logTimestamp);
      SpreadsheetApp.flush(); 
    } finally {
      lock.releaseLock();
    }

    // --- 4. [신규] 로그 큐(PropertiesService)에 로그 데이터 저장 ---
    const newLogId = "LOG_" + logTimestamp.getTime() + "_" + Math.random().toString(36).substr(2, 9);
    // ✨ [추가] 9KB 제한(UTF-8 약 9000자)보다 훨씬 전에 차단 (예: 8000자)
    if (logContent.length > 8000) {
        throw new Error("상담 기록이 너무 깁니다. 8000자 이내로 나누어 저장해주세요.");
    }
    // 큐에 저장할 데이터 객체
    const logDataToQueue = {
      logId: newLogId,
      assignmentId: assignmentId,
      logTimestamp: logTimestamp.toISOString(), // ISO 문자열로 저장
      logContent: logContent,
      userName: userName
    };
    try {
      // 고유한 키로 스크립트 속성에 저장
      const logQueueKey = 'log_queue_' + newLogId;
      PropertiesService.getScriptProperties().setProperty(logQueueKey, JSON.stringify(logDataToQueue));
    // [수정]
      } catch (e) {
        // Logger.log(`로그 큐 저장 실패: ${e.message}`);
// [삭제]
        logError_('addConsultationLog_Queue', e, { assignmentId: assignmentId });
// [추가]
        throw new Error("로그를 임시 저장하는 데 실패했습니다. 잠시 후 다시 시도해주세요.");
    }

    // --- 5. [변경] 후속 작업 (Lock 외부) ---
    bustUserListCache_();
    // 캐시 무효화

    // Gviz 쿼리로 "시트에 이미 저장된" 로그 + "큐에 있는" 로그를 모두 가져옵니다.
    // [수정] existingLogs -> finalLogs로 변수명 변경 (의미 명확화)
    const finalLogs = findLogsByAssignmentId_(assignmentId);
    // [!!!]
    // [제거] findLogsByAssignmentId_가 큐에 있는 새 로그를 이미 가져오므로
    //       수동으로 newLogForClient를 만들 필요가 없습니다.
    /*
    const newLogForClient = {
      logId: newLogId,
      assignmentId: assignmentId,
      logTimestamp: logTimestamp.toISOString(), // Gviz 결과와 맞춤
      logContent: logContent,
      userName: userName
    };
    */
    
    // [제거] 위 객체를 목록 맨 앞에 추가하는 로직 제거
    // const combinedLogs = [newLogForClient, ...existingLogs];
// [삭제]

    // 반환할 고객 데이터 조합
    if (!foundAssignmentData) {
      const found = findRowById_("배정고객", "assignmentId", assignmentId);
      if (found) {
        foundAssignmentData = found.rowData;
      } else {
        throw new Error("최종 고객 데이터를 조회하는 데 실패했습니다.");
      }
    } else {
      // (이 부분은 'lastLogDate'가 즉시 UI에 반영되도록 기존 로직 유지)
      foundAssignmentData.lastLogDate = logTimestamp.toISOString();
    }
    
    // [신규] Task 4: 계약 목록도 함께 반환
    const finalContracts = findContractsByAssignmentId_(assignmentId);

    return {
      assignment: foundAssignmentData,
      // [수정] combinedLogs 대신 finalLogs (findLogsByAssignmentId_의 결과)를 반환
      logs: finalLogs,
      contracts: finalContracts // [신규]
    };
  });
}

/**
 * [중요] 아키텍처 결정 사항: 이 함수는 의도적으로 '큐(PropertiesService)'를 사용하지 않습니다.
 * * 1. 문제 상황:
 * 이 함수가 addConsultationLog처럼 '큐'를 사용(비동기 처리)할 경우, 
 * 클라이언트(UI)는 즉시 응답을 받아 고객이 목록에 추가된 것처럼 보이지만,
 * 실제 '배정고객' 시트에는 batchWriteLogs_v2 트리거가 실행될 때까지 (최대 1~5분) 데이터가 없습니다.
 * * 2. 오류 시나리오:
 * 사용자가 신규 고객을 등록한 직후(1~5분 이내), 해당 고객을 클릭해 상담 기록을 추가하려 하면
 * addConsultationLog 함수는 권한/정보 확인을 위해 findRowById_를 호출합니다.
 * 이때 시트에 고객 데이터가 아직 없으므로 "해당 배정 ID를 찾을 수 없습니다" 오류가 발생합니다.
 *
 * 3. 해결책 (현재 방식):
 * '신규 고객 등록' 작업은 '상담 기록 추가'보다 빈도가 훨씬 낮습니다.
 * 따라서 이 작업은 약간의 UI 지연(Lock 획득 및 시트 쓰기 시간 0.5~1초)을 감수하더라도,
 * 데이터 일관성(즉시 시트에 반영됨)을 보장하는 것이 더 중요합니다.
 * * 이에 따라 이 함수는 LockService를 사용하여 '배정고객' 시트에 직접 appendRow(동기 쓰기)를 수행합니다.
 * * ※ '상담 기록 추가(addConsultationLog)'는 빈도가 높고 즉각적인 UI 피드백이 중요하므로 '큐'를 사용하는 것이 맞으며,
 * 두 함수의 아키텍처는 의도적으로 다르게 설계되었습니다.
 */

function addNewCustomer(customerData) {
  // [수정] Lock을 다시 사용하여 시트 일관성을 보장합니다.
  return measurePerformance_('addNewCustomer_SheetWrite', () => {
    // [수정] 작업 이름 전달
    const lock = acquireLockWithRetry_('addNewCustomer'); 
    
    try {
      const timestamp = new Date(); // Date 객체
      const userEmail = Session.getActiveUser().getEmail();
      const configData = getConfigurations();
      const userName = configData.emailToNameMap[userEmail] || userEmail;
      
      const assignmentId = "A_" + timestamp.getTime();
      const customerId = "C_" + timestamp.getTime();

 
       // [수정] 시트에 직접 쓸 데이터 객체 (클라이언트 반환용)
      // (클라이언트 피드백을 위해 ISO 문자열 사용)
      const newCustomerData = {
        assignmentId: assignmentId,
        customerId: customerId,
        customerName: customerData.customerName,
        customerPhoneNumber: formatPhoneNumber(customerData.customerPhoneNumber), // 서버 포맷
        assignedTo: userName,
        dbType: customerData.dbType,
        assignmentDate: 
 timestamp.toISOString(), // 클라이언트 반환용 ISO 문자열
        consultationStatus: "배정됨",
        // [수정] Task 2: 'contractStatus'는 '배정고객' 시트에서 제거됨
        // contractStatus: "미해당" // [삭제]
      };
      // [수정] 헤더 맵을 동적으로 가져옵니다.
      const headers = assignmentSheet.getRange(1, 1, 1, assignmentSheet.getLastColumn()).getValues()[0];
      const headerMap = {};
      headers.forEach((h, i) => { if(h) headerMap[h] = i; });

      const newRow = Array(headers.length).fill(null);
      // 객체 데이터를 배열 순서에 맞게 매핑
      for (const headerKey in newCustomerData) {
        if (headerMap[headerKey] !== undefined) {
          let value = newCustomerData[headerKey];
          // [중요] 시트에 쓸 때는 Date 객체로 변환
          if (headerKey === 'assignmentDate') { 
            value = new Date(value);
 // ISO 문자열을 다시 Date 객체로
          }
          newRow[headerMap[headerKey]] = value;
        }
      }
      
      // [중요] SearchHelper 열은 ARRAYFORMULA가 채우도록 비워둡니다 (newRow[headerMap['SearchHelper']] = null).
      // [수정] 시트에 직접 appendRow 실행 (Lock 내부)
      assignmentSheet.appendRow(newRow);
      SpreadsheetApp.flush();
      // (선택 사항이지만 Lock 내부에선 권장)

      // [제거] 큐(PropertiesService) 관련 로직 모두 제거
      
      bustUserListCache_();
      // 캐시 무효화는 동일하게 실행

      // [수정] 시트에 방금 쓴 객체를 클라이언트에 반환
      return newCustomerData;
    } catch (e) {
        logError_('addNewCustomer', e, { customerName: customerData.customerName });
        throw new Error("고객 추가 중 오류가 발생했습니다.");
      } finally {
        // [수정] Lock 해제
        if (lock) lock.releaseLock();
      }
  }); // End of measurePerformance_
}

/**
 * [신규] Task 3: 신규 계약 추가
 * 클라이언트에서 받은 계약 정보를 '계약기록' 시트에 추가합니다.
 */
function addNewContract(assignmentId, contractData) {
  // [신규] Task 3: 신규 계약 추가
  return measurePerformance_('addNewContract', () => {
    if (!contractsSheet) throw new Error("'계약기록' 시트를 찾을 수 없습니다.");
    if (!assignmentId) throw new Error("고객 ID(assignmentId)가 없습니다.");

    // 1. 권한 확인 (이 고객이 본인 고객인지)
    const userEmail = Session.getActiveUser().getEmail();
    const configData = getConfigurations();
    const userName = configData.emailToNameMap[userEmail] || userEmail;
    
    const found = findRowById_("배정고객", "assignmentId", assignmentId);
    if (!found) throw new Error("계약할 고객을 찾을 수 없습니다.");
    if (found.rowData.assignedTo !== userName) { 
      throw new Error("본인에게 배정된 고객의 계약만 추가할 수 있습니다.");
    }

    // 2. Lock 획득 (appendRow는 동기 쓰기이므로)
    const lock = acquireLockWithRetry_(`addNewContract: ${assignmentId}`);
    
    try {
      const timestamp = new Date();
      const contractId = "C_" + timestamp.getTime() + "_" + Math.random().toString(36).substr(2, 5);

      // 3. 헤더 맵 가져오기
      const headers = contractsSheet.getRange(1, 1, 1, contractsSheet.getLastColumn()).getValues()[0];
      const headerMap = {};
      headers.forEach((h, i) => { if(h) headerMap[h] = i; });
      
      // 4. 데이터 매핑
      // (주의: contractData의 키는 '계약기록' 시트의 헤더와 일치해야 함)
      const newRow = Array(headers.length).fill(null);
      
      // 필수/기본값 설정
      newRow[headerMap['contractId']] = contractId;
      newRow[headerMap['assignmentId']] = assignmentId;
      
      // 클라이언트에서 받은 데이터 매핑
      for (const headerKey in contractData) {
        if (headerMap[headerKey] !== undefined) {
          let value = contractData[headerKey];
          // 날짜/숫자 변환 (필요시)
          if (headerKey === 'contractDate' && value) {
            value = new Date(value);
          }
          if (headerKey === 'contractAmount' && value) {
            value = parseFloat(value);
          }
          newRow[headerMap[headerKey]] = value;
        }
      }

      // 5. 시트에 쓰기
      contractsSheet.appendRow(newRow);
      SpreadsheetApp.flush();

      // 6. 클라이언트에 새로고침할 계약 목록 반환
      // (appendRow는 즉시 반영되므로, Gviz 쿼리를 다시 실행하면 새 데이터가 포함됨)
      bustUserListCache_(); // 고객 목록 캐시도 혹시 모르니 (LTV 등)
      return findContractsByAssignmentId_(assignmentId);

    } catch (e) {
      logError_('addNewContract_SheetWrite', e, { assignmentId: assignmentId, data: contractData });
      throw new Error("신규 계약을 저장하는 중 오류가 발생했습니다.");
    } finally {
      if (lock) lock.releaseLock();
    }
  });
}


/**
 * Gviz API의 날짜 응답 문자열(예: "Date(2025,9,27,13,53,40)")을
 * 'Date' 객체로 변환합니다.
 * @param {string | Date | number} gvizDateValue - Gviz 응답의 cell.v 값
 * @returns {Date |
 null} Date 객체 또는 null
 */
function parseGvizDateObject_(gvizDateValue) {
  if (!gvizDateValue) return null;
  // 1. "Date(..." 문자열 형식 처리
  if (typeof gvizDateValue === 'string' && gvizDateValue.startsWith('Date(')) {
    try {
      // "Date(2025,9,27,13,53,40)" -> [2025, 9, 27, 13, 53, 40]
      // JavaScript의 Month는 0-based이므로, gvizDateValue[1] (월)은 그대로 사용합니다.
      const params = gvizDateValue.substring(5, gvizDateValue.length - 1).split(',').map(Number);
      
      const date = new Date(
        params[0], // Year
        params[1], // Month (0-based)
     
       params[2], // Day
        params[3] || 0, // Hours
        params[4] || 0, // Minutes
        params[5] || 0, // Seconds
        params[6] || 0  // Milliseconds
      );
      
      if (!isNaN(date.getTime())) {
        return date; // Date 객체 반환
      }
    } catch (e) 
 {
      Logger.log(`Gviz Date 파싱 오류 (String): ${gvizDateValue} | ${e.message}`);
      return null; 
    }
  }

  // 2. 이미 Date 객체이거나 유효한 날짜 문자열인 경우
  try {
    const parsedDate = new Date(gvizDateValue);
    if (!isNaN(parsedDate.getTime())) {
      return parsedDate;
// Date 객체 반환
    }
  } catch (e) {
    Logger.log(`Gviz Date 파싱 오류 (Other): ${gvizDateValue} | ${e.message}`);
  }

  return null; // 파싱 불가 시 null
}

// Code.gs

// -----------------------------------------------------------------
// ✨ [개선된 배치 쓰기 함수]
// -----------------------------------------------------------------

/**
 * [배치 쓰기 함수 - v3 / 우회로 적용 / 알림 수정 완료 / ✨'대규모 백로그' 처리 개선]
 * PropertiesService에 쌓인 큐를 읽어와 시트에 일괄 기록합니다.
 * 한 번에 MAX_ITEMS_PER_RUN 개수만큼만 처리하여 큐가 과도하게 쌓여도 안전하게 분할 처리합니다.
 */
function batchWriteLogs_v2() {
  // [1.업무 시간 필터]
  try {
    const now = new Date();
    const hour = now.getHours();
    const isOffHours = (hour < 6); 
    
    if (isOffHours) {
      Logger.log(`batchWriteLogs_v2 SKIPPED - 업무 시간 아님 (현재 ${hour}시)`);
      return; 
    }
  } catch (e) {
    Logger.log(`업무 시간 필터 중 오류 발생: ${e.message}`);
  }
  
  const scriptProperties = PropertiesService.getScriptProperties();

  // ✨ [개선] 한 번의 실행으로 처리할 최대 항목 수 (로그, 신규 고객 각각)
  const MAX_ITEMS_PER_RUN = 200;
  // ✨ [개선] 큐 모니터링 및 getProperties()를 Lock 획득 *전에* 한 번만 호출
  let allProperties;
  let allKeys;
  try {
    allProperties = scriptProperties.getProperties();
    allKeys = Object.keys(allProperties);
  } catch (e) {
    // getProperties() 자체가 실패하는 심각한 상황 (500KB 초과 등)
    logError_('batchWriteLogs_v2_GetProperties_FATAL', e, {});
    // 치명적 알림 전송 (1시간 1회)
    sendCriticalAlert_("getProperties() 호출 실패", e);
    return;
    // Lock 획득 시도조차 하지 않고 종료
  }

  // [신규] 큐 깊이 모니터링 (Lock 획득 전)
  try {
    const logQueueDepth = allKeys.filter(k => k.startsWith('log_queue_')).length;
    const custQueueDepth = allKeys.filter(k => k.startsWith('new_cust_queue_')).length;
    const totalQueueDepth = logQueueDepth + custQueueDepth;

    const CRITICAL_DEPTH = 500;
// 🚨 임계값 (예: 500개)

    if (totalQueueDepth > CRITICAL_DEPTH) {
        Logger.log(`[CRITICAL_QUEUE_DEPTH] 큐가 ${totalQueueDepth}개로 위험 수위입니다.`);
        const cache = CacheService.getScriptCache();
        const alertCacheKey = 'batch_worker_QUEUE_DEPTH_alert_sent';
        
        if (!cache.get(alertCacheKey)) {
            const adminEmail = "gyumin4660@gmail.com";
// 🚨 관리자 이메일
            const subject = "[경고] 고객상담 시스템 큐(Queue) 적체 심각";
            const body = `
                배치 작업 큐가 ${totalQueueDepth}개 (로그: ${logQueueDepth} / 신규: ${custQueueDepth})로
                위험 수위(${CRITICAL_DEPTH}개)를 초과했습니다.
                getProperties() 호출이 실패하기 전에 즉시 점검이 필요합니다.
                'ErrorLog' 시트나 Apps Script 대시보드에서 batchWriteLogs_v2 트리거가
                정상 동작하는지 확인해주세요.
            `;
            MailApp.sendEmail(adminEmail, subject, body);
            cache.put(alertCacheKey, 'true', 3600); // 1시간 1회
        }
    }
  } catch (e) {
      logError_('batchWriteLogs_v2_QueueMonitor', e, {});
  }
  // [신규] 모니터링 끝

  // [2. Lock 획득]
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) { 
    Logger.log("batchWriteLogs_v2 SKIPPED - Lock 획득 실패 (이미 실행 중)");
    return;
  }

  // --- 시트 객체 정의 ---
  // [수정] const ss = SpreadsheetApp.getActiveSpreadsheet(); // 전역 변수 ss 사용
  // const assignmentSheet = ss.getSheetByName("배정고객");
// const logSheet = ss.getSheetByName("상담기록");
  // [수정] '배정고객' 및 '상담기록' 시트 객체는 전역 변수(assignmentSheet, logSheet)를 사용합니다.

  try { // -------------------------------------
        // --- 메인 Try 블록 시작 ---
        // -------------------------------------
    
    // -----------------------------------------
    // [3-1.상담 로그 큐 처리]
    // -----------------------------------------
    
    // ✨ [개선] 미리 읽어둔 allKeys 사용
    let logKeys = allKeys.filter(k => k.startsWith('log_queue_'));
    // ✨ [개선] 처리할 키의 개수를 제한합니다.
    const logKeysToProcess = logKeys.slice(0, MAX_ITEMS_PER_RUN);
    if (logKeysToProcess.length > 0) {
      // ✨ [개선] 전체 큐 개수와 함께 로깅
      Logger.log(`batchWriteLogs_v2 - ${logKeysToProcess.length}개/${logKeys.length}개의 로그를 처리합니다.`);
      if (!logSheet) throw new Error("'상담기록' 시트를 찾을 수 없습니다.");
      
      const logSheetHeaders = logSheet.getRange(1, 1, 1, logSheet.getLastColumn()).getValues()[0];
      const logHeaderMap = {};
      logSheetHeaders.forEach((header, index) => { if (header) logHeaderMap[header] = index; });

      const requiredLogCols = ['logId', 'assignmentId', 'logTimestamp', 'logContent', 'userName'];
      for (const col of requiredLogCols) {
        if (logHeaderMap[col] === undefined) {
          throw new Error(`'상담기록' 시트 1행에서 필수 헤더 '${col}'를 찾을 수 없습니다.`);
        }
      }

      // ✨ [개선] 미리 읽어둔 allProperties 사용
      const logsData = {};
      logKeysToProcess.forEach(key => { // [개선]
        if (allProperties[key]) {
          logsData[key] = allProperties[key];
        }
      });
      let logsToWrite = [];
      for (const key of logKeysToProcess) { // [개선]
        const logDataString = logsData[key];
        if (!logDataString) continue;

        try {
          // --- 개별 로그 파싱 Try ---
          const logData = JSON.parse(logDataString);
          const newRow = Array(logSheetHeaders.length).fill(null);
          
          newRow[logHeaderMap.logId] = logData.logId;
          newRow[logHeaderMap.assignmentId] = logData.assignmentId;
          newRow[logHeaderMap.logTimestamp] = new Date(logData.logTimestamp);
          newRow[logHeaderMap.logContent] = logData.logContent;
          newRow[logHeaderMap.userName] = logData.userName;
          logsToWrite.push(newRow);
        } catch (e) {
        // ✨ [Issue #6 적용] Dead Letter Queue 로직
        Logger.log(`로그 데이터 파싱 오류 (키: ${key}): ${e.message}. 실패한 큐로 이동합니다.`);
        // 1. 실패한 항목을 별도 큐로 이동
        scriptProperties.setProperty(
          `failed_log_queue_${key}`, // "failed_" 접두사 추가
          JSON.stringify({
            originalData: logDataString,
            error: e.message,
            timestamp: new Date().toISOString()
          })
        );
// 2. 원본 큐에서는 삭제 (루프가 끝난 후 logKeysToProcess.forEach에서 어차피 삭제됨)
        
        // 3. 에러 로그 남기기
        logError_('batchWriteLogs_v2_ParseFail_Log', e, { 
          propertyKey: key, 
          movedToFailedQueue: true 
        });
        }
      }

      if (logsToWrite.length > 0) {
        const lastRow = logSheet.getLastRow();
        logSheet.getRange(lastRow + 1, 1, logsToWrite.length, logSheetHeaders.length)
                .setValues(logsToWrite);
        Logger.log(`batchWriteLogs_v2 - ${logsToWrite.length}개 로그 시트 쓰기 완료.`);
      }

      logKeysToProcess.forEach(key => { // [개선]
        scriptProperties.deleteProperty(key);
      });
      Logger.log(`batchWriteLogs_v2 - ${logKeysToProcess.length}개 로그 큐 삭제 완료.`);
    
    } else {
      Logger.log("batchWriteLogs_v2 - 쓸 로그가 없습니다.");
    }

    // -----------------------------------------
    // [3-2.신규 고객 큐 처리]
    // -----------------------------------------
    
    // ✨ [개선] 미리 읽어둔 allKeys 사용
    const custKeys = allKeys.filter(k => k.startsWith('new_cust_queue_'));
    // ✨ [개선] 처리할 키의 개수를 제한합니다.
    const custKeysToProcess = custKeys.slice(0, MAX_ITEMS_PER_RUN);
    if (custKeysToProcess.length > 0) {
      // ✨ [개선] 전체 큐 개수와 함께 로깅
      Logger.log(`batchWriteLogs_v2 - ${custKeysToProcess.length}명/${custKeys.length}명의 신규 고객을 처리합니다.`);
      if (!assignmentSheet) throw new Error("'배정고객' 시트를 찾을 수 없습니다.");
      
      const assignHeaders = assignmentSheet.getRange(1, 1, 1, assignmentSheet.getLastColumn()).getValues()[0];
      const assignHeaderMap = {};
      assignHeaders.forEach((h, i) => { if(h) assignHeaderMap[h] = i; });

      // [수정] Task 2: 'contractStatus' 제거
      const requiredCustCols = ['assignmentId', 'customerId', 'customerName', 'customerPhoneNumber', 'assignedTo', 'dbType', 'assignmentDate', 'consultationStatus'];
      for (const col of requiredCustCols) {
        if (assignHeaderMap[col] === undefined) {
          throw new Error(`'배정고객' 시트 1행에서 필수 헤더 '${col}'를 찾을 수 없습니다.`);
        }
      }

      // ✨ [개선] 미리 읽어둔 allProperties 사용
      const custDataMap = {};
      custKeysToProcess.forEach(key => { // [개선]
        if (allProperties[key]) {
          custDataMap[key] = allProperties[key];
        }
      });
      const customersToWrite = [];
      for (const key of custKeysToProcess) { // [개선]
        const custDataString = custDataMap[key];
        if (!custDataString) continue;
        
        try {
        // --- 개별 고객 파싱 Try ---
        const custData = JSON.parse(custDataString);
        const newRow = Array(assignHeaders.length).fill(null);
        
        for (const headerKey in custData) {
          if (assignHeaderMap[headerKey] !== undefined) {
            let value = custData[headerKey];
            if (headerKey === 'assignmentDate') { 
              value = new Date(value);
            }
            newRow[assignHeaderMap[headerKey]] = value;
          }
        }

        // ✨ [Issue #2 적용] SearchHelper 열을 미리 소문자로 채웁니다.
        /*
        const searchHelperColIndex = assignHeaderMap['SearchHelper'];
        if (searchHelperColIndex !== undefined) {
          const name = custData.customerName || '';
          const phone = (custData.customerPhoneNumber || '').replace(/\D/g, ''); // 숫자만
          // (이름 + 전화번호) 소문자 조합을 SearchHelper 열에 저장
          newRow[searchHelperColIndex] = (name + phone).toLowerCase();
        }
        */
        
        customersToWrite.push(newRow);
        } catch (e) {
        // ✨ [Issue #6 적용] Dead Letter Queue 로직
        Logger.log(`신규 고객 큐 파싱 오류 (키: ${key}): ${e.message}. 실패한 큐로 이동합니다.`);
        // 1. 실패한 항목을 별도 큐로 이동
        scriptProperties.setProperty(
          `failed_cust_queue_${key}`, // "failed_" 접두사 추가
          JSON.stringify({
            originalData: custDataString,
            error: e.message,
            timestamp: new Date().toISOString()
          })
        );
// 2. 원본 큐에서는 삭제 (마찬가지로 루프 후 삭제됨)
        
        // 3. 에러 로그 남기기
        logError_('batchWriteLogs_v2_ParseFail_Cust', e, { 
          propertyKey: key, 
          movedToFailedQueue: true 
          });
        }
      }

      if (customersToWrite.length > 0) {
        const lastRow = assignmentSheet.getLastRow();
        assignmentSheet.getRange(lastRow + 1, 1, customersToWrite.length, assignHeaders.length)
                             .setValues(customersToWrite);
        Logger.log(`batchWriteLogs_v2 - ${customersToWrite.length}명 신규 고객 시트 쓰기 완료.`);
      }

      custKeysToProcess.forEach(key => { // [개선]
        scriptProperties.deleteProperty(key);
      });
      Logger.log(`batchWriteLogs_v2 - ${custKeysToProcess.length}개 신규 고객 큐 삭제 완료.`);
    } else {
      Logger.log("batchWriteLogs_v2 - 추가할 신규 고객이 없습니다.");
    }

  } catch (e) { // -------------------------------------
            // --- 메인 Catch 블록 (치명적 오류 발생 시) ---
            // -------------------------------------
    Logger.log(`batchWriteLogs_v2 실행 중 심각한 오류 발생: ${e.message} ${e.stack}`);
    // [수정] logError_를 사용하여 시트에도 기록
    try {
        logError_('batchWriteLogs_v2_FATAL', e, {});
    } catch (logErr) {
        // logError_ 자체도 실패할 경우 대비
        Logger.log(`[FATAL_LOGGING_ERROR] logError_ 실패: ${logErr.message} | 원본 오류: ${e.message}`);
    }
    
    // [개선] 관리자 이메일 알림 (1시간 1회) - 헬퍼 함수 호출
    sendCriticalAlert_("batchWriteLogs_v2 실행 실패", e);
  } finally {
    lock.releaseLock(); 
  }
}

// -----------------------------------------------------------------
// ✨ [헬퍼 함수] - batchWriteLogs_v2가 의존하는 함수
// -----------------------------------------------------------------

/**
 * [신규 헬퍼] 치명적 오류 발생 시 관리자에게 1시간에 1회 알림 이메일을 전송합니다.
 * @param {string} subjectPrefix - 이메일 제목에 포함될 오류 컨텍스트 (예: "getProperties() 호출 실패")
 * @param {Error} error - 발생한 오류 객체
 */
function sendCriticalAlert_(subjectPrefix, error) {
  const cache = CacheService.getScriptCache();
  // 오류 컨텍스트별로 고유한 캐시 키를 생성하여 알림이 중복되지 않도록 함
  const alertCacheKey = `batch_worker_alert_sent_${subjectPrefix.replace(/[\s\(\)]/g, '_')}`;
  if (!cache.get(alertCacheKey)) {
      const adminEmail = "gyumin4660@gmail.com";
// 🚨 관리자 이메일
      const subject = `[긴급] ${subjectPrefix} - 고객상담 시스템 배치 작업`;
      const body = `
          ${subjectPrefix} 중 심각한 오류가 발생했습니다.
          PropertiesService의 큐가 시트로 기록되지 않을 수 있습니다.
          
          Error: ${error.message}
          Stack: ${error.stack ||
 'No stack trace'}
          
          'ErrorLog' 시트를 확인해주세요.
      `;
      try {
          MailApp.sendEmail(adminEmail, subject, body);
          // 1시간 동안 알림 중복 방지
          cache.put(alertCacheKey, 'true', 3600);
      } catch (mailErr) {
          Logger.log(`[FATAL] 관리자 알림 이메일 발송 실패: ${mailErr.message}`);
      }
  }
}

/**
 * [기존 헬퍼] 중앙 집중식 에러 로깅 함수
 * (참고: 이 함수는 Code.gs 상단()에 이미 정의되어 있습니다.)
 */
/*
function logError_(context, error, additionalInfo = {}) {
  // ... (중복 정의) ...
}
*/

/**
 * [임시 청소 함수]
 * PropertiesService에 꼬여있는 큐를 강제로 삭제합니다.
 */
function manualClearQueue_v2() {
  const scriptProperties = PropertiesService.getScriptProperties();
  
  // 1. 로그 큐 삭제
  const logKeys = scriptProperties.getKeys().filter(k => k.startsWith('log_queue_'));
  if (logKeys.length > 0) {
    Logger.log(`[수동삭제] ${logKeys.length}개의 로그 큐를 삭제합니다...`);
    logKeys.forEach(key => {
      scriptProperties.deleteProperty(key);
    });
    Logger.log("[수동삭제] 로그 큐 삭제 완료.");
  } else {
    Logger.log("삭제할 로그 큐가 없습니다.");
  }

  // 2. 신규 고객 큐 삭제
  const custKeys = scriptProperties.getKeys().filter(k => k.startsWith('new_cust_queue_'));
  if (custKeys.length > 0) {
    Logger.log(`[수동삭제] ${custKeys.length}개의 신규 고객 큐를 삭제합니다...`);
    custKeys.forEach(key => {
      scriptProperties.deleteProperty(key);
    });
    Logger.log("[수동삭제] 신규 고객 큐 삭제 완료.");
  } else {
    Logger.log("삭제할 신규 고객 큐가 없습니다.");
  }
}

/**
 * [신규] 8번 제안: 야간 정리 함수
 * 매일 새벽 2~3시경 실행되도록 트리거 설정
 */
function nightlyCleanup() {
  // [수정] const ss = SpreadsheetApp.getActiveSpreadsheet(); // 전역 변수 ss 사용
  Logger.log("야간 정리 작업을 시작합니다.");

  // 1. 오래된 스크립트 캐시 정리 (Config, 헤더 맵)
  try {
    const cache = CacheService.getScriptCache();
    // 참고:removeAll은 패턴 매칭을 지원하지 않습니다. 
    // 키를 직접 명시하거나, getKeys로 가져와서 필터링 후 remove해야 하나,
    // Config/Header 맵은 put할 때 만료시간(3600s)이 있어 자동 만료되므로 
    // 굳이 수동 삭제할 필요는 없습니다.(Logger.log로 기록만 남깁니다)
    Logger.log("스크립트 캐시(Config, Header)는 자동 만료(1시간)됩니다.");
  } catch (e) {
    Logger.log(`[ERROR] 캐시 정리 실패: ${e.message}`);
  }

  // 2. 오래된 에러 로그 삭제 (ErrorLog 시트가 있는 경우)
  try {
    const errorSheet = ss.getSheetByName("ErrorLog");
    if (errorSheet && errorSheet.getLastRow() > 500) { // 500줄 이상 쌓이면
    // 최근 300줄만 남기고 삭제
    const rowsToDelete = errorSheet.getLastRow() - 300 - 1;
// (1행 헤더 제외)
    
    if (rowsToDelete > 0) {
      // 1행(헤더) 다음인 2행부터 `rowsToDelete`개 만큼 삭제
      errorSheet.deleteRows(2, rowsToDelete);
// ★ [수정] '-1' 제거
      Logger.log(`오래된 에러 로그 ${rowsToDelete}줄 삭제 완료.`);
    }
  }
  } catch (e) {
    Logger.log(`[ERROR] 에러 로그 시트 정리 실패: ${e.message}`);
  }

  Logger.log("야간 정리 작업을 완료했습니다.");
}

// [Code.gs] 파일 하단에 새 헬퍼 함수 추가

/**
 * [Issue #6] 실패한 큐(Dead Letter Queue)에 쌓인 항목을
 * 관리자가 수동으로 검토할 수 있도록 로깅합니다.
 * (실행 -> Apps Script -> 실행 -> reviewFailedQueue 선택 후 실행)
 */
function reviewFailedQueue() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const failedLogKeys = scriptProperties.getKeys().filter(k => k.startsWith('failed_log_queue_'));
  const failedCustKeys = scriptProperties.getKeys().filter(k => k.startsWith('failed_cust_queue_'));

  Logger.log(`--- 실패한 로그 큐 (${failedLogKeys.length}개) ---`);
  failedLogKeys.forEach(key => {
    const data = scriptProperties.getProperty(key);
    Logger.log(`[${key}]: ${data}`);
    // 예: 여기서 데이터를 수동 파싱 후 시트에 강제 삽입
    // scriptProperties.deleteProperty(key); // 관리자가 수동 처리 후 삭제
  });
  Logger.log(`--- 실패한 신규 고객 큐 (${failedCustKeys.length}개) ---`);
  failedCustKeys.forEach(key => {
    const data = scriptProperties.getProperty(key);
    Logger.log(`[${key}]: ${data}`);
    // scriptProperties.deleteProperty(key); // 관리자가 수동 처리 후 삭제
  });
  if (failedLogKeys.length === 0 && failedCustKeys.length === 0) {
    Logger.log("실패한 큐(DLQ)가 없습니다.");
    SpreadsheetApp.getUi().alert("실패한 큐(DLQ)가 없습니다.");
  } else {
    Logger.log("--- Apps Script 로그(Ctrl+Enter)에서 상세 내용을 확인하세요. ---");
    SpreadsheetApp.getUi().alert(`실패한 큐 ${failedLogKeys.length + failedCustKeys.length}건이 발견되었습니다. Apps Script 로그(Ctrl+Enter)를 확인하세요.`);
  }
}

// [Code.gs] 파일 하단에 새 함수를 추가합니다.
/**
 * [Issue #12] 매일 새벽 시트 전체를 백업 파일로 생성합니다.
 * (참고: '1ShOx3fcJx44ZrO5KXPd_t8C0BH2osB08'는 실제 구글 드라이브 폴더 ID로 변경해야 합니다.)
 */
function dailyBackup() {
  try {
    // 🚨 여기를 실제 구글 드라이브 폴더 ID로 변경하세요.
// (폴더 URL이 "https://drive.google.com/drive/folders/ABCDEFG" 라면 "ABCDEFG"가 ID입니다.)
    const BACKUP_FOLDER_ID = "1ShOx3fcJx44ZrO5KXPd_t8C0BH2osB08";
    if (BACKUP_FOLDER_ID === "YOUR_BACKUP_FOLDER_ID_HERE") {
      Logger.log("백업 폴더 ID가 설정되지 않아 dailyBackup()을 건너뜁니다.");
      return;
    }

    const backupFolder = DriveApp.getFolderById(BACKUP_FOLDER_ID);
    
    const timestamp = Utilities.formatDate(new Date(), 'GMT+9', 'yyyyMMdd');
    const fileName = `[백업] 고객상담_${timestamp}`;
    // 1. 시트 파일 백업
    ss.copy(fileName).moveTo(backupFolder);
    Logger.log(`백업 생성 완료: ${fileName}`);
    // 2. 30일 이상 된 오래된 백업 삭제
    const thirtyDaysAgo = new Date(Date.now() - 30 * 24 * 60 * 60 * 1000);
    const oldBackups = backupFolder.getFiles();
    
    let deletedCount = 0;
    while (oldBackups.hasNext()) {
      const file = oldBackups.next();
      // 백업 파일이고, 생성일이 30일이 지났는지 확인
      if (file.getName().startsWith('[백업]') && file.getDateCreated() < thirtyDaysAgo) {
        file.setTrashed(true);
// 휴지통으로 이동
        deletedCount++;
      }
    }
    if (deletedCount > 0) {
      Logger.log(`오래된 백업 ${deletedCount}개 삭제 완료.`);
    }

  } catch (e) {
    logError_('dailyBackup_FATAL', e, { folderId: BACKUP_FOLDER_ID });
    // 백업 실패 시 관리자에게 알림
    sendCriticalAlert_('자동 백업 실패', e);
  }
}

// [Code.gs] - 파일 하단에 아래 함수 추가

/**
 * 특정 시/도 이름에 해당하는 시/군/구 목록을 반환합니다.
 * @param {string} cityName - 조회할 시/도 이름 (예: "서울특별시")
 * @returns {Array<string>} 시/군/구 이름 배열 (예: ["강남구", "강동구", ...])
 */
function getDistrictsForCity(cityName) {
  // SELECT_OPTIONS 상수에서 해당 시/도의 시/군/구 배열을 찾아 반환
  // 만약 해당 시/도 정보가 없으면 빈 배열 반환
  return SELECT_OPTIONS.addressDistricts[cityName] ||
 [];
}

/**
 * 사용자 이메일과 필터를 기반으로 기본 캐시 키를 생성합니다.
 * @param {string} userEmail 사용자 이메일
 * @param {object} filters 적용된 필터 객체
 * @returns {string} 캐시 키 문자열
 */
function getDefaultCacheKey_(userEmail, filters) {
  // 필터 객체의 키를 정렬하여 순서에 상관없이 동일한 키 생성
  const filterKeys = Object.keys(filters).sort();
  const filterString = filterKeys.map(key => `${key}:${filters[key]}`).join('|');
  
  // 데이터 버전 포함 (bustUserListCache_에서 증가시키는 버전)
  const dataVersion = PropertiesService.getUserProperties().getProperty('DATA_VERSION') || '1';
  // MD5 해시를 사용하여 키 길이를 줄임 (선택 사항이지만 권장)
  const hashInput = `${userEmail}|${filterString}|v${dataVersion}`;
  const hash = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, hashInput)
                 .map(byte => (byte & 0xFF).toString(16).padStart(2, '0'))
                 .join('');
  return `assigned_cust_${hash}`; 
}
