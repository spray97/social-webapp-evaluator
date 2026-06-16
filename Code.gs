// =====================================================
// 사회문제 해결 웹앱 제작하기 - 수행평가 채점 도구
// K(11): 코드상태(AI) / L(12): 코드분석(AI) / M(13): 연관성분석(AI)
// N(14): 교사의견 / O(15): 점수 / P(16): 교과세특(AI)
// =====================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📋 수행평가 채점')
    .addItem('▶ 채점 도구 열기', 'openEvaluator')
    .addToUi();
}

function openEvaluator() {
  var html = HtmlService.createHtmlOutputFromFile('Evaluator')
    .setWidth(1500)
    .setHeight(950)
    .setTitle('사회문제 해결 웹앱 채점 시스템');
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

// ── 데이터 읽기 ──────────────────────────────────────

function getStudentData(rowIndex) {
  rowIndex = rowIndex || 1;
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  var targetRow = Math.max(1, Math.min(rowIndex, lastRow - 1));
  var row = sheet.getRange(targetRow + 1, 1, 1, 16).getValues()[0];

  return {
    rowIndex:          targetRow,
    totalStudents:     lastRow - 1,
    studentId:         row[2]  || '-',
    name:              row[3]  || '이름없음',
    appName:           row[4]  || '(미제출)',
    appLink:           row[5]  || '',
    codeLink:          row[6]  || '',
    problem:           row[7]  || '(미제출)',
    features:          row[8]  || '(미제출)',
    tools:             row[9]  || '(미제출)',
    codeStatus:        row[10] || '',  // K(11)
    codeAnalysis:      row[11] || '',  // L(12)
    relevanceAnalysis: row[12] || '',  // M(13)
    comment:           row[13] || '',  // N(14)
    score:             row[14] || '',  // O(15)
    setech:            row[15] || ''   // P(16): 교과세특
  };
}

// ── 데이터 저장 ──────────────────────────────────────

function saveResult(rowIndex, codeStatus, codeAnalysis, relevanceAnalysis, comment, score) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var actualRow = parseInt(rowIndex) + 1;
  sheet.getRange(actualRow, 11).setValue(codeStatus        || '');  // K(11): 코드상태
  sheet.getRange(actualRow, 12).setValue(codeAnalysis      || '');  // L(12): 코드분석
  sheet.getRange(actualRow, 13).setValue(relevanceAnalysis || '');  // M(13): 연관성분석
  sheet.getRange(actualRow, 14).setValue(comment           || '');  // N(14): 교사의견
  sheet.getRange(actualRow, 15).setValue(score             || '');  // O(15): 점수
  return true;
}

// ── 시트2: 3차 수행평가 데이터 읽기 ─────────────────

function getSheet2Data(studentId) {
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var sheet2 = ss.getSheetByName('시트2');
  if (!sheet2) return null;

  var lastRow = sheet2.getLastRow();
  if (lastRow < 2) return null;

  var data = sheet2.getRange(2, 1, lastRow - 1, 3).getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(studentId).trim()) {
      return { studentId: data[i][0], name: data[i][1], content: data[i][2] || '' };
    }
  }
  return null;
}

// ── Drive 파일 헬퍼 ──────────────────────────────────

// G열 셀 예시:
//   "https://drive.google.com/open?id=ABC123"
//   "https://drive.google.com/open?id=ABC123, https://drive.google.com/open?id=DEF456"
function parseDriveLinks(linkString) {
  if (!linkString) return [];
  var matches = linkString.match(/id=([a-zA-Z0-9_-]+)/g);
  if (!matches) return [];
  return matches.map(function(m) { return m.replace('id=', ''); });
}

function readDriveFile(fileId) {
  try {
    return DriveApp.getFileById(fileId).getBlob().getDataAsString('UTF-8');
  } catch (e) {
    return null;
  }
}

// ── OpenRouter AI 분석 ────────────────────────────────

function callOpenRouter(prompt) {
  var apiKey = PropertiesService.getScriptProperties()
                 .getProperty('OPENROUTER_API_KEY');
  if (!apiKey) throw new Error('OPENROUTER_API_KEY가 Script Properties에 설정되지 않았습니다.');

  var response = UrlFetchApp.fetch('https://openrouter.ai/api/v1/chat/completions', {
    method: 'post',
    headers: {
      'Authorization':  'Bearer ' + apiKey,
      'Content-Type':   'application/json',
      'HTTP-Referer':   'https://script.google.com',
      'X-Title':        'Social Webapp Evaluator'
    },
    payload: JSON.stringify({
      model:      'google/gemini-2.5-flash-lite',
      messages:   [{ role: 'user', content: prompt }],
      max_tokens: 800
    }),
    muteHttpExceptions: true
  });

  var parsed = JSON.parse(response.getContentText());
  if (parsed.error) throw new Error(parsed.error.message);
  return parsed.choices[0].message.content.trim();
}

function analyzeCode(rowIndex) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    var row   = sheet.getRange(parseInt(rowIndex) + 1, 1, 1, 14).getValues()[0];

    var codeLink = row[6] || '';
    var problem  = row[7] || '';
    var features = row[8] || '';

    var fileIds = parseDriveLinks(codeLink);
    if (fileIds.length === 0) {
      return { codeStatus: '오류코드', codeAnalysis: '코드 제출 링크가 없거나 파싱할 수 없습니다.' };
    }

    var codeParts = [];
    fileIds.forEach(function(id, i) {
      var content = readDriveFile(id);
      if (content) codeParts.push('=== 파일 ' + (i + 1) + ' ===\n' + content);
    });

    if (codeParts.length === 0) {
      return { codeStatus: '오류코드', codeAnalysis: 'Drive 파일을 읽을 수 없습니다. 파일 공유 권한을 확인하세요.' };
    }

    var codeText = codeParts.join('\n\n');

    var MAX_CODE_CHARS = 15000;
    var truncated = codeText.length > MAX_CODE_CHARS;
    if (truncated) {
      codeText = codeText.substring(0, MAX_CODE_CHARS) +
        '\n\n[주의: 코드가 너무 길어 앞부분 ' + MAX_CODE_CHARS + '자까지만 분석합니다]';
    }

    var combinedPrompt =
      '다음 학생이 제출한 코드를 분석하세요.\n\n' +
      '응답 형식을 반드시 지켜주세요:\n' +
      '첫 번째 줄: "정상코드" 또는 "오류코드" 중 하나만 작성\n' +
      '두 번째 줄부터: 코드가 학생의 문제 분석·핵심 기능과 얼마나 일치하는지 200자 이내 한국어로 작성\n\n' +
      '[학생이 주목한 문제점 및 대책]\n' + problem + '\n\n' +
      '[학생이 제시한 핵심 기능]\n' + features + '\n\n' +
      '[제출된 코드]\n' + codeText;

    var combinedRaw  = callOpenRouter(combinedPrompt);
    var lines        = combinedRaw.split('\n');
    var codeStatus   = lines[0].includes('정상') ? '정상코드' : '오류코드';
    var codeAnalysis = lines.slice(1).join('\n').trim();

    // 3차 수행평가 연관성 분석
    var studentId     = row[2];
    var appName       = row[4];
    var sheet2Data    = getSheet2Data(studentId);
    var sheet2Content = '';
    var relevanceAnalysis = '';

    if (sheet2Data && sheet2Data.content) {
      sheet2Content = sheet2Data.content;
      var relevancePrompt =
        '다음은 한 학생의 3차 수행평가(사회문제 분석)와 4차 수행평가(웹앱 개발) 결과입니다.\n' +
        '4차 웹앱이 3차에서 제시한 사회문제/사회현상을 해결(완화)하려는 방향과 연관성이 있는지 분석하세요.\n' +
        '반드시 첫 번째 줄에 [높음], [보통], [낮음] 중 정확히 하나만 작성하고, ' +
        '두 번째 줄부터 이유를 200자 이내 한국어로 작성하세요.\n\n' +
        '[3차 수행평가: 사회문제/사회현상 및 기대효과]\n' + sheet2Content + '\n\n' +
        '[4차 수행평가: 개발한 웹앱]\n' +
        '앱 이름: ' + appName + '\n' +
        '문제점 원인 및 대책: ' + problem + '\n' +
        '핵심 기능: ' + features;
      relevanceAnalysis = callOpenRouter(relevancePrompt);
    } else {
      relevanceAnalysis = '시트2에서 학번(' + studentId + ')에 해당하는 3차 수행평가 자료를 찾을 수 없습니다.';
    }

    return {
      codeStatus:        codeStatus,
      codeAnalysis:      codeAnalysis,
      sheet2Content:     sheet2Content,
      relevanceAnalysis: relevanceAnalysis
    };

  } catch (e) {
    return { codeStatus: '오류코드', codeAnalysis: '분석 중 오류: ' + e.message, sheet2Content: '', relevanceAnalysis: '' };
  }
}

// ── 검증용 (GAS 에디터에서 직접 실행) ─────────────────

function getNextUnscored(fromRowIndex) {
  var sheet   = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  for (var i = fromRowIndex + 1; i <= lastRow - 1; i++) {
    var score = sheet.getRange(i + 1, 15).getValue(); // O열(15번째): 점수
    if (!score && score !== 0) return i;
  }
  return null;
}

function generateSetech(rowIndex) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    var row   = sheet.getRange(parseInt(rowIndex) + 1, 1, 1, 15).getValues()[0];

    var studentId = row[2] || '';
    var name      = row[3] || '학생';
    var appName   = row[4] || '';
    var problem   = row[7] || '';
    var features  = row[8] || '';
    var tools     = row[9] || '';

    var sheet2Data    = getSheet2Data(studentId);
    var sheet3Content = sheet2Data ? sheet2Data.content : '(3차 수행평가 자료 없음)';

    var setechPrompt =
      '당신은 고등학교 교사로서 학생의 교과세특(세부능력특기사항)을 작성합니다.\n\n' +
      '아래 학생의 3차 수행평가(선행연구자료 수집)와 4차 수행평가(사회문제 해결 웹앱 개발) 결과를 바탕으로, ' +
      '두 활동을 유기적으로 연계한 교과세특을 작성하세요.\n\n' +
      '## 작성 규칙\n' +
      '- 반드시 350~400자(한글 기준, 공백 포함)로 작성\n' +
      '- 블룸 분류학 4~6단계 동사를 자연스럽게 활용: 분석·추론·평가·검토·설계·구성·해결·통합·제안 등\n' +
      '- "~를 원용하여", "~개념을 원용하고" 형태 표현을 적절히 포함\n' +
      '- 영문 약어 사용 절대 금지 (API→외부 응용 프로그래밍 인터페이스, GAS→구글 앱스 스크립트, AI→인공지능)\n' +
      '- 교과명을 본문에 직접 나열하지 말 것 → "범교과적으로", "융합적 관점에서" 등으로 대체\n' +
      '- 5단계 전개: [문제 인식] → [탐구 방법] → [구현 결과] → [분석·통찰] → [성장·마무리]\n' +
      '- 학생 이름을 주어로 하는 3인칭 서술 (예: "' + name + '은(는) ~을 분석함", "~을 설계함")\n' +
      '- 수치 단순 나열 금지 → 서술형으로 맥락화\n' +
      '- 앱 이름은 따옴표 없이 자연스럽게 문장에 포함\n\n' +
      '## 학생 이름\n' +
      name + '\n\n' +
      '## 3차 수행평가: 사회문제 선행연구자료 수집\n' +
      sheet3Content + '\n\n' +
      '## 4차 수행평가: 사회문제 해결 웹앱 개발\n' +
      '앱 이름: ' + appName + '\n' +
      '문제점 원인 및 대책: ' + problem + '\n' +
      '핵심 기능: ' + features + '\n' +
      '활용 도구: ' + tools + '\n\n' +
      '교과세특을 작성하세요. 완성 후 글자 수(공백 포함)를 확인하여 반드시 350~400자 범위에 맞춰 주세요.';

    var setech = callOpenRouter(setechPrompt);

    // P열(16번째)에 저장
    sheet.getRange(parseInt(rowIndex) + 1, 16).setValue(setech);

    return { setech: setech };
  } catch (e) {
    return { setech: '생성 오류: ' + e.message };
  }
}

function testApiKey() {
  var key = PropertiesService.getScriptProperties()
              .getProperty('OPENROUTER_API_KEY');
  Logger.log(key ? '✅ 키 등록됨: ' + key.substring(0, 10) + '...' : '❌ 키 없음');
}

function testAnalyze() {
  var result = analyzeCode(1);
  Logger.log(JSON.stringify(result));
}
