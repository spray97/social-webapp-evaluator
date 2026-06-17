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

function callOpenRouter(prompt, model) {
  var apiKey = PropertiesService.getScriptProperties()
                 .getProperty('OPENROUTER_API_KEY');
  if (!apiKey) throw new Error('OPENROUTER_API_KEY가 Script Properties에 설정되지 않았습니다.');

  var usedModel = model || 'google/gemini-3.5-flash';

  var response = UrlFetchApp.fetch('https://openrouter.ai/api/v1/chat/completions', {
    method: 'post',
    headers: {
      'Authorization':  'Bearer ' + apiKey,
      'Content-Type':   'application/json',
      'HTTP-Referer':   'https://script.google.com',
      'X-Title':        'Social Webapp Evaluator'
    },
    payload: JSON.stringify({
      model:      usedModel,
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
      '다음은 고등학생이 "사회문제 해결 웹앱"으로 제출한 코드입니다.\n' +
      '코드 텍스트만으로 분석하며, 실제 실행 결과는 확인할 수 없음을 전제로 합니다.\n\n' +
      '아래 형식을 반드시 지켜 출력하세요. 각 태그를 그대로 포함하세요.\n\n' +
      '[코드상태]\n' +
      '다음 세 가지 중 하나만 작성:\n' +
      '- 정상코드: 문법 오류 없고 실행 가능한 수준\n' +
      '- 오류코드: 명백한 문법 오류·미완성 코드 존재\n' +
      '- 기능미흡: 실행은 되나 단순 텍스트 표시·메모장 수준으로 사회문제 해결과 무관한 기능만 존재\n\n' +
      '[난이도]\n' +
      '다음 세 가지 중 하나만 작성:\n' +
      '- 기초: 단순 UI 표시, 로직·계산 없음, 사용자 입력만 저장하는 수준\n' +
      '- 중급: 입력·저장·조회 로직, 조건 분기, 간단한 계산 포함\n' +
      '- 고급: 외부 API 연동, 데이터 연산·알고리즘, 다중 기능 연동\n\n' +
      '[구현기능]\n' +
      '코드에서 실제로 확인된 기능만 나열 (200자 이내, 추정 금지)\n\n' +
      '[미구현의심]\n' +
      '학생이 제출 설명에 기재했으나 코드에서 확인되지 않거나 placeholder 수준인 기능.\n' +
      '없으면 "없음"이라고만 작성.\n\n' +
      '[채점주의]\n' +
      '교사가 앱을 직접 실행하여 확인해야 할 핵심 포인트 (100자 이내)\n\n' +
      '[학생이 주목한 문제점 및 대책]\n' + problem + '\n\n' +
      '[학생이 제시한 핵심 기능]\n' + features + '\n\n' +
      '[제출된 코드]\n' + codeText;

    var combinedRaw = callOpenRouter(combinedPrompt);

    function extractSection(raw, tag) {
      var re = new RegExp('\\[' + tag + '\\]\\s*([\\s\\S]*?)(?=\\[|$)');
      var m  = raw.match(re);
      return m ? m[1].trim() : '';
    }

    var codeStatusRaw = extractSection(combinedRaw, '코드상태');
    var difficulty    = extractSection(combinedRaw, '난이도');
    var implemented   = extractSection(combinedRaw, '구현기능');
    var notImpl       = extractSection(combinedRaw, '미구현의심');
    var caution       = extractSection(combinedRaw, '채점주의');

    var codeStatus = codeStatusRaw.includes('정상') ? '정상코드'
                   : codeStatusRaw.includes('미흡') ? '기능미흡'
                   : '오류코드';

    var codeAnalysis =
      '【난이도】' + difficulty + '\n' +
      '【구현기능】' + implemented + '\n' +
      '【미구현의심】' + notImpl + '\n' +
      '【채점주의】' + caution;

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
      difficulty:        difficulty,
      implemented:       implemented,
      notImpl:           notImpl,
      caution:           caution,
      codeAnalysis:      codeAnalysis,
      sheet2Content:     sheet2Content,
      relevanceAnalysis: relevanceAnalysis
    };

  } catch (e) {
    return {
      codeStatus: '오류코드', difficulty: '', implemented: '', notImpl: '', caution: '',
      codeAnalysis: '분석 중 오류: ' + e.message, sheet2Content: '', relevanceAnalysis: ''
    };
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
  var SETECH_MODEL = 'anthropic/claude-sonnet-4-6';
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
      '아래 3차·4차 수행평가 내용을 하나의 탐구 서사로 연결하여 교과세특을 작성하세요.\n\n' +
      '## 작성 규칙\n\n' +
      '### 1. 글자 수·구조 (가장 중요한 규칙)\n' +
      '- 반드시 **3개의 문장**으로만 구성 (문장 추가·분리 금지)\n' +
      '- 각 문장 목표 길이: 120~130자\n' +
      '- 전체 합계: 반드시 350~400자(한글 기준, 공백 포함)\n' +
      '- 문단 구분 없이 하나의 흐름으로 작성\n\n' +
      '### 2. 글머리 고정 패턴 (반드시 이 형식으로 시작)\n' +
      '  사회문제 해결 프로젝트 활동에서 \'[탐구주제]\'를 탐구 주제로 설정하고,\n' +
      '- [탐구주제]는 3차·4차 수행평가 내용에서 핵심 사회 문제를 추출하여 채움\n\n' +
      '### 3. 주어 처리\n' +
      '- 학생 이름·"학생은" 등 주어를 생략하고 "~함", "~함으로써" 형태로 통일\n' +
      '- 능동적 주체성이 드러나도록 서술 (수동태 금지)\n\n' +
      '### 4. 탐구 흐름 3단 구조 (인과 흐름 필수)\n' +
      '  [이론 탐구] → [분석·추론] → [설계·구현·제안]\n' +
      '- 단순 활동 나열이 아닌, 문제 인식에서 결과물까지 인과 흐름으로 구성\n\n' +
      '### 5. 블룸 분류학 동사 심화 순서\n' +
      '- 분석·추론(4단계) → 평가·검토(5단계) → 설계·통합·제안(6단계) 순으로 인지가 심화되도록 배치\n' +
      '- 같은 단계 동사를 반복 사용하지 않음\n\n' +
      '### 6. "원용하다" 활용\n' +
      '- 선행연구의 이론이나 개념을 분석 틀로 적용할 때 "~을 원용하여" 형태로 1회 사용\n\n' +
      '### 7. 마무리 평가 문장 (반드시 포함)\n' +
      '- 마지막 문장은 교사가 학생의 역량을 종합 평가하는 어투로 마무리\n' +
      '- 패턴: "~하는 등 남다른 [역량명]이 돋보임", "~[역량]을 보여줌"\n' +
      '- 예: "남다른 융합적 관점과 문제 해결 역량이 돋보임"\n' +
      '- 예: "남다른 인공지능 활용 및 웹 개발 능력을 보여줌"\n\n' +
      '### 8. 금지 사항\n' +
      '- 영문 약어 절대 금지: API→외부 응용 프로그래밍 인터페이스, AI→인공지능, HTML→웹 문서 구조\n' +
      '- 특정 상호명·제품명 기재 금지: 제미나이·챗지피티·클로드·앱스크립트 등\n' +
      '  → 대체어: 생성형 인공지능, 대형 언어 모델, 웹 스크립트 환경, 인공지능 도구\n' +
      '- 교과명 직접 나열 금지 → "범교과적", "융합적 관점에서" 등으로 대체\n' +
      '- 앱 이름은 작은따옴표로 표기 (예: \'앱이름\')\n\n' +
      '### 9. 대학 입시 역량 (역량 이름을 본문에 쓰지 말 것)\n' +
      '- 선행연구 분석·이론 원용 → 학업역량이 활동으로 드러남\n' +
      '- 전공 연관 개념·도구 활용, 결과물 → 진로역량이 활동으로 드러남\n' +
      '- 사회 문제 설정 배경·사회에 기여하는 결과물 → 공동체역량이 활동으로 드러남\n\n' +
      '## 참고 샘플 (이 구조와 어투를 따를 것)\n' +
      '"사회문제 해결 프로젝트 활동에서 \'공간 불평등 문제\'를 탐구 주제로 설정하고, ' +
      '선행연구자료의 선형 계획 모델을 원용하여 인구 구조적 수요, 교통망 인프라, 개인별 한계 이동 거리를 ' +
      '다각적으로 분석함으로써 수도권 자원 초집중이 지역 간 정주 여건 양극화를 심화시키는 구조적 원인임을 추론함. ' +
      '이를 바탕으로 공공데이터 좌표를 활용해 전국 226개 시·군·구의 상급종합병원·교육 거점까지의 물리적 이격 거리를 ' +
      '실시간 연산하고, 공간불평등 지수를 산출하여 등급별 행정 정책 건의서를 자동 생성하는 웹앱 \'이퀄로케이터\'를 ' +
      '설계·구현하여 소개함으로써 남다른 인공지능 활용 및 웹 개발 능력을 보여줌. ' +
      '지방소멸 위험 지역에 거점 연계형 복지 대안을 제시하고 인프라 포화 지역에 운영 최적화 방안을 제언하는 기능을 통해 ' +
      '데이터 기반 국토 균형 발전 정책의 실효성을 검증하는 등 남다른 융합적 관점과 문제 해결 역량이 돋보임."\n\n' +
      '## 3차 수행평가: 사회문제 선행연구자료 수집\n' +
      sheet3Content + '\n\n' +
      '## 4차 수행평가: 사회문제 해결 웹앱 개발\n' +
      '앱 이름: ' + appName + '\n' +
      '문제점 원인 및 대책: ' + problem + '\n' +
      '핵심 기능: ' + features + '\n' +
      '활용 도구: ' + tools + '\n\n' +
      '교과세특 본문만 출력하세요. 글자 수, 부가 설명, 기호 등 부가 정보는 절대 포함하지 마세요.\n' +
      '출력 전 반드시 글자 수(공백 포함)를 세어 350~400자인지 확인하고, 초과하면 각 문장을 줄여 조정하세요. ' +
      '참고 샘플의 글자 수는 무시하고, 오직 350~400자 기준만 따르세요.';

    var setech = callOpenRouter(setechPrompt, SETECH_MODEL);

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
