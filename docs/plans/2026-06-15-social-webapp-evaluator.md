# Social Webapp Evaluator Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Google Apps Script 기반 채점 도구를 제작한다. 학생의 폼 응답을 시각화하고, Google Drive에 업로드된 코드 파일(txt)을 OpenRouter AI로 분석해 코드 유효성(K열)과 구현 일치 분석(L열)을 자동 생성하며, 교사가 의견(M열)과 점수(N열)를 직접 입력·저장할 수 있다.

**Architecture:** `Code.gs`(서버)와 `Evaluator.html`(클라이언트) 두 파일로 구성된 GAS 프로젝트. gas-evaluator 패턴을 계승하되 OpenRouter API 호출, Google Drive 파일 읽기, 열 구조 확장(K~N)을 추가한다. AI 분석은 교사가 버튼을 눌렀을 때만 실행(on-demand)된다.

**Tech Stack:** Google Apps Script, SpreadsheetApp, DriveApp, UrlFetchApp (OpenRouter), HTML/CSS/JS (GAS modal dialog)

**추천 OpenRouter 모델:** `google/gemini-2.0-flash-001`
- 이유: 한국어 처리 우수, 코드 분석 능력 충분, 토큰당 비용 최저 수준, 응답 속도 빠름
- 대안: `anthropic/claude-haiku-4-5` (더 정확하나 비용 약 3배)

---

## 열 구조 (최종 확정)

| 열 | 인덱스 | 내용 | 담당 |
|----|--------|------|------|
| A | 0 | 타임스탬프 | 폼 |
| B | 1 | 이메일 | 폼 |
| C | 2 | 학번 | 폼 |
| D | 3 | 이름 | 폼 |
| E | 4 | 앱 서비스 이름 | 폼 |
| F | 5 | 앱 서비스 링크 | 폼 |
| G | 6 | 코드 제출 링크 (Drive txt) | 폼 |
| H | 7 | 문제점 원인 및 대책 | 폼 |
| I | 8 | 핵심 기능 및 이유 | 폼 |
| J | 9 | 활용 프로그램 | 폼 |
| **K** | **10** | **코드 상태** (`정상코드`/`오류코드`) | **AI 자동** |
| **L** | **11** | **코드 분석 200자** | **AI 자동** |
| **M** | **12** | **교사 의견** | **교사 입력** |
| **N** | **13** | **점수** (30~12) | **교사 선택** |

---

## 파일 구조

```
social-webapp-evaluator/
├── Code.gs          # 서버: 데이터 읽기/쓰기, Drive 파일 읽기, OpenRouter 호출
├── Evaluator.html   # 클라이언트: 채점 UI (modal dialog)
├── CLAUDE.md        # 프로젝트 문서
└── docs/plans/
    └── 2026-06-15-social-webapp-evaluator.md
```

---

## Task 1: API 키 설정 및 CLAUDE.md 작성

**Files:**
- Create: `social-webapp-evaluator/CLAUDE.md`

### 1-1. GAS Script Properties에 API 키 등록

GAS 에디터에서 다음 순서로 진행:

```
1. Google Sheets 열기 → 확장 프로그램 → Apps Script
2. 상단 메뉴: 프로젝트 설정 (⚙️ 톱니바퀴)
3. "스크립트 속성" 섹션 → "스크립트 속성 추가"
4. 속성: OPENROUTER_API_KEY
   값: [본인의 OpenRouter API 키 붙여넣기]
5. 저장
```

- [ ] **Step 1: GAS 에디터를 열고 Script Properties 설정 화면으로 이동**

  Google Sheets → 확장 프로그램 → Apps Script → 프로젝트 설정(⚙️) → 스크립트 속성

- [ ] **Step 2: OPENROUTER_API_KEY 속성 추가 및 저장**

  ```
  속성명: OPENROUTER_API_KEY
  값: sk-or-v1-xxxxxxxxxxxx  (본인 키)
  ```

- [ ] **Step 3: CLAUDE.md 생성**

  아래 내용으로 `social-webapp-evaluator/CLAUDE.md` 작성 (Task 완료 후 파일 생성)

---

## Task 2: Code.gs — 데이터 레이어

**Files:**
- Create: `social-webapp-evaluator/Code.gs`

- [ ] **Step 1: 기본 구조 및 메뉴 함수 작성**

```javascript
// =====================================================
// 사회문제 해결 웹앱 제작하기 - 수행평가 채점 도구
// K열: 코드상태 / L열: 코드분석 / M열: 교사의견 / N열: 점수
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
```

- [ ] **Step 2: getStudentData 함수 작성 (A~N열, 14개 열)**

```javascript
function getStudentData(rowIndex) {
  rowIndex = rowIndex || 1;
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  var targetRow = Math.max(1, Math.min(rowIndex, lastRow - 1));
  var row = sheet.getRange(targetRow + 1, 1, 1, 14).getValues()[0];

  return {
    rowIndex:     targetRow,
    totalStudents: lastRow - 1,
    studentId:    row[2]  || '-',
    name:         row[3]  || '이름없음',
    appName:      row[4]  || '(미제출)',
    appLink:      row[5]  || '',
    codeLink:     row[6]  || '',
    problem:      row[7]  || '(미제출)',
    features:     row[8]  || '(미제출)',
    tools:        row[9]  || '(미제출)',
    codeStatus:   row[10] || '',
    codeAnalysis: row[11] || '',
    comment:      row[12] || '',
    score:        row[13] || ''
  };
}
```

- [ ] **Step 3: saveResult 함수 작성 (K·L·M·N열 동시 저장)**

```javascript
function saveResult(rowIndex, codeStatus, codeAnalysis, comment, score) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var actualRow = parseInt(rowIndex) + 1;
  sheet.getRange(actualRow, 11).setValue(codeStatus   || '');
  sheet.getRange(actualRow, 12).setValue(codeAnalysis || '');
  sheet.getRange(actualRow, 13).setValue(comment      || '');
  sheet.getRange(actualRow, 14).setValue(score        || '');
  return true;
}
```

- [ ] **Step 4: Drive 링크 파싱 헬퍼 함수 작성**

```javascript
// G열 셀 값 예시:
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
```

- [ ] **Step 5: 수동 검증 — getStudentData 단독 실행**

  GAS 에디터에서 `getStudentData` 함수 선택 → ▶ 실행  
  로그(Ctrl+Enter)에서 `{rowIndex:1, name:'강혜원', ...}` 형태로 출력되면 정상

---

## Task 3: Code.gs — OpenRouter AI 통합

**Files:**
- Modify: `social-webapp-evaluator/Code.gs` (함수 추가)

- [ ] **Step 1: callOpenRouter 헬퍼 함수 작성**

```javascript
function callOpenRouter(prompt) {
  var apiKey = PropertiesService.getScriptProperties()
                 .getProperty('OPENROUTER_API_KEY');
  if (!apiKey) throw new Error('OPENROUTER_API_KEY가 Script Properties에 설정되지 않았습니다.');

  var response = UrlFetchApp.fetch('https://openrouter.ai/api/v1/chat/completions', {
    method: 'post',
    headers: {
      'Authorization': 'Bearer ' + apiKey,
      'Content-Type': 'application/json',
      'HTTP-Referer': 'https://script.google.com',
      'X-Title': 'Social Webapp Evaluator'
    },
    payload: JSON.stringify({
      model: 'google/gemini-2.0-flash-001',
      messages: [{ role: 'user', content: prompt }],
      max_tokens: 600
    }),
    muteHttpExceptions: true
  });

  var parsed = JSON.parse(response.getContentText());
  if (parsed.error) throw new Error(parsed.error.message);
  return parsed.choices[0].message.content.trim();
}
```

- [ ] **Step 2: analyzeCode 함수 작성 (AI 분석 메인)**

```javascript
function analyzeCode(rowIndex) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    var row = sheet.getRange(parseInt(rowIndex) + 1, 1, 1, 14).getValues()[0];

    var codeLink   = row[6] || '';
    var problem    = row[7] || '';
    var features   = row[8] || '';

    // Drive에서 코드 파일 읽기
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

    // 프롬프트 1: 코드 유효성 판단
    var validityPrompt =
      '다음은 학생이 제출한 Google Apps Script(code.gs)와 HTML(index.html) 코드입니다.\n' +
      '이 코드가 문법적으로 유효하고 실행 가능한지 판단하세요.\n' +
      '반드시 "정상코드" 또는 "오류코드" 단어 하나만 답하세요. 다른 설명 없이.\n\n' +
      '코드:\n' + codeText;

    // 프롬프트 2: 구현 일치 분석
    var analysisPrompt =
      '학생이 제출한 코드가 학생의 문제 분석 및 핵심 기능과 얼마나 일치하는지 분석하세요.\n' +
      '반드시 200자 이내 한국어로 작성하세요. 예시 형식: "학생이 제시한 ~기능이 코드에서 ~로 구현되어 ~"\n\n' +
      '[학생이 주목한 문제점 및 대책]\n' + problem + '\n\n' +
      '[학생이 제시한 핵심 기능]\n' + features + '\n\n' +
      '[제출된 코드]\n' + codeText;

    var statusRaw  = callOpenRouter(validityPrompt);
    var codeStatus = statusRaw.includes('정상') ? '정상코드' : '오류코드';

    var codeAnalysis = callOpenRouter(analysisPrompt);
    if (codeAnalysis.length > 200) codeAnalysis = codeAnalysis.substring(0, 200);

    return { codeStatus: codeStatus, codeAnalysis: codeAnalysis };

  } catch (e) {
    return { codeStatus: '오류코드', codeAnalysis: '분석 중 오류: ' + e.message };
  }
}
```

- [ ] **Step 3: 수동 검증 — analyzeCode(1) 직접 실행**

  GAS 에디터에서 임시 테스트 함수 실행:
  ```javascript
  function testAnalyze() {
    var result = analyzeCode(1);
    Logger.log(result);
  }
  ```
  실행 로그에서 `{codeStatus: '정상코드', codeAnalysis: '...'}` 형태 확인

---

## Task 4: Evaluator.html — 레이아웃 및 스타일

**Files:**
- Create: `social-webapp-evaluator/Evaluator.html`

- [ ] **Step 1: HTML 기본 구조 및 CSS 작성**

```html
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <link href="https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@400;700&display=swap" rel="stylesheet">
  <style>
    :root { --body-font-size: 17px; }
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      font-family: 'Noto Sans KR', sans-serif;
      background: #f0f2f5;
      height: 100vh;
      display: flex;
      flex-direction: column;
      overflow: hidden;
    }

    /* ── 헤더 ── */
    .header {
      background: #1a3a5c;
      color: white;
      padding: 10px 24px;
      display: flex;
      justify-content: space-between;
      align-items: center;
      flex-shrink: 0;
    }
    .student-info h2 { font-size: 22px; }
    .student-info span { font-size: 13px; opacity: 0.75; margin-left: 12px; }

    .controls { display: flex; align-items: center; gap: 10px; }
    .btn-sm { padding: 6px 12px; font-size: 13px; font-weight: bold; border: none; border-radius: 4px; cursor: pointer; color: white; background: #3a5a7c; }
    .btn-sm:hover { background: #4a6a8c; }
    .btn-fullscreen { background: #c0392b; }
    .btn-fullscreen:hover { background: #a93226; }
    .nav-btns { display: flex; align-items: center; gap: 7px; border-left: 1px solid #3a5a7c; padding-left: 12px; }
    .btn-nav { padding: 7px 14px; font-size: 13px; font-weight: bold; cursor: pointer; border-radius: 4px; border: 1px solid #ccc; background: white; color: #333; }
    .btn-nav:hover { background: #e0e4ea; }
    .btn-nav-next { background: #0056b3; color: white; border: none; }
    .btn-nav-next:hover { background: #004494; }
    .jump-input { width: 50px; padding: 6px; text-align: center; font-size: 13px; border: 1px solid #ccc; border-radius: 4px; }
    #prog_text { font-size: 14px; font-weight: bold; min-width: 55px; text-align: center; }

    /* ── 메인 ── */
    .main-container {
      flex: 1;
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 16px;
      padding: 16px;
      overflow: hidden;
    }
    .panel {
      background: white;
      border-radius: 8px;
      padding: 20px 24px;
      overflow-y: auto;
      box-shadow: 0 2px 8px rgba(0,0,0,0.07);
      display: flex;
      flex-direction: column;
      gap: 20px;
    }

    /* 앱 정보 카드 */
    .app-card {
      background: #eef4ff;
      border: 1px solid #c5d8f5;
      border-radius: 8px;
      padding: 14px 18px;
    }
    .app-card .app-name { font-size: 18px; font-weight: bold; color: #1a3a5c; margin-bottom: 10px; }
    .app-card .link-btns { display: flex; gap: 10px; flex-wrap: wrap; }
    .btn-link {
      padding: 7px 16px; font-size: 13px; font-weight: bold; border-radius: 5px;
      cursor: pointer; border: none; color: white; text-decoration: none; display: inline-block;
    }
    .btn-app-link  { background: #0056b3; }
    .btn-code-link { background: #5a6a7c; }
    .btn-link:hover { opacity: 0.85; }

    /* AI 분석 띠 */
    .ai-strip {
      border-top: 2px solid #e0e4ea;
      padding: 10px 16px;
      background: #fafbfc;
      display: flex;
      align-items: center;
      gap: 16px;
      flex-shrink: 0;
    }
    .btn-analyze {
      padding: 8px 20px; font-size: 14px; font-weight: bold; border: none;
      border-radius: 6px; cursor: pointer; background: #6f42c1; color: white; white-space: nowrap;
    }
    .btn-analyze:hover { background: #5a32a3; }
    .btn-analyze:disabled { background: #aaa; cursor: default; }
    .ai-badge {
      padding: 5px 14px; border-radius: 20px; font-size: 13px; font-weight: bold;
      white-space: nowrap; display: none;
    }
    .ai-badge.normal { background: #d4edda; color: #155724; display: inline-block; }
    .ai-badge.error  { background: #f8d7da; color: #721c24; display: inline-block; }
    .ai-analysis-text { font-size: 13px; color: #444; line-height: 1.6; flex: 1; }

    /* 섹션 */
    .q-label { font-size: 14px; font-weight: bold; color: #0056b3; margin-bottom: 8px; padding-bottom: 6px; border-bottom: 2px solid #eef; }
    .q-text  { font-size: var(--body-font-size); line-height: 1.8; color: #1a1a1a; white-space: pre-wrap; word-break: keep-all; }

    /* ── 푸터 ── */
    .footer-bar {
      background: white;
      border-top: 2px solid #dee2e6;
      padding: 12px 24px;
      display: flex;
      align-items: center;
      gap: 20px;
      flex-shrink: 0;
    }
    .comment-wrap { flex: 1; }
    .comment-wrap textarea {
      width: 100%; height: 56px; padding: 10px 14px;
      font-size: 15px; border: 1px solid #ced4da; border-radius: 5px; resize: none;
      font-family: 'Noto Sans KR', sans-serif;
    }
    .score-area { display: flex; flex-direction: column; gap: 8px; align-items: flex-end; }
    .score-btns { display: flex; gap: 6px; flex-wrap: wrap; justify-content: flex-end; max-width: 520px; }
    .btn-score {
      width: 52px; height: 42px; font-size: 16px; font-weight: bold;
      border: 2px solid #1a3a5c; background: white; color: #1a3a5c;
      border-radius: 6px; cursor: pointer; transition: 0.1s;
    }
    .btn-score.active { background: #1a3a5c; color: white; transform: scale(1.08); }
    .btn-score.unsubmit { border-color: #c0392b; color: #c0392b; }
    .btn-score.unsubmit.active { background: #c0392b; color: white; }
    .btn-save {
      padding: 10px 22px; background: #28a745; color: white; border: none;
      border-radius: 6px; font-size: 16px; font-weight: bold; cursor: pointer; white-space: nowrap;
    }
    .btn-save:hover { background: #218838; }
  </style>
</head>
```

- [ ] **Step 2: HTML body 마크업 작성**

```html
<body>
  <!-- 헤더 -->
  <header class="header">
    <div class="student-info">
      <h2 id="disp_name">학생 성명</h2>
      <span id="disp_id">학번: ----- | 로딩 중...</span>
    </div>
    <div class="controls">
      <button class="btn-sm btn-fullscreen" onclick="toggleFullScreen()">🖥️ 전체화면</button>
      <span style="font-size:12px; opacity:.7">글자:</span>
      <button class="btn-sm" onclick="adjustFont(2)">➕</button>
      <button class="btn-sm" onclick="adjustFont(-2)">➖</button>
      <div class="nav-btns">
        <button class="btn-nav" onclick="move(-1)">◀ 이전</button>
        <input type="number" id="jump_idx" class="jump-input" placeholder="번호" onkeydown="if(event.key==='Enter') jump()">
        <span id="prog_text">1 / 0</span>
        <button class="btn-nav btn-nav-next" onclick="move(1)">다음 ▶</button>
      </div>
    </div>
  </header>

  <!-- 메인 2열 레이아웃 -->
  <main class="main-container">
    <!-- 왼쪽: 앱 정보 + 문제점 -->
    <div class="panel">
      <div class="app-card">
        <div class="app-name" id="disp_appName">앱 이름</div>
        <div class="link-btns">
          <a id="link_app" class="btn-link btn-app-link" href="#" target="_blank">🔗 앱 실행하기</a>
          <a id="link_code" class="btn-link btn-code-link" href="#" target="_blank">📄 코드 보기</a>
        </div>
      </div>
      <div>
        <div class="q-label">6. 문제점 원인 및 대책</div>
        <div class="q-text" id="disp_problem">...</div>
      </div>
    </div>

    <!-- 오른쪽: 핵심 기능 + 활용 프로그램 -->
    <div class="panel">
      <div>
        <div class="q-label">7. 핵심 기능 및 이유</div>
        <div class="q-text" id="disp_features">...</div>
      </div>
      <div>
        <div class="q-label">8. 활용 프로그램</div>
        <div class="q-text" id="disp_tools">...</div>
      </div>
    </div>
  </main>

  <!-- AI 분석 띠 -->
  <div class="ai-strip">
    <button class="btn-analyze" id="btn_analyze" onclick="runAnalysis()">🤖 AI 코드 분석</button>
    <span class="ai-badge" id="ai_badge"></span>
    <span class="ai-analysis-text" id="ai_text">분석 버튼을 클릭하면 코드 유효성 및 구현 일치 여부를 자동 분석합니다.</span>
  </div>

  <!-- 푸터 -->
  <footer class="footer-bar">
    <div class="comment-wrap">
      <textarea id="comment_input" placeholder="교사 의견을 입력하세요 (M열 저장)"></textarea>
    </div>
    <div class="score-area">
      <div class="score-btns">
        <button class="btn-score" onclick="setScore(30)">30</button>
        <button class="btn-score" onclick="setScore(29)">29</button>
        <button class="btn-score" onclick="setScore(28)">28</button>
        <button class="btn-score" onclick="setScore(27)">27</button>
        <button class="btn-score" onclick="setScore(26)">26</button>
        <button class="btn-score" onclick="setScore(25)">25</button>
        <button class="btn-score" onclick="setScore(24)">24</button>
        <button class="btn-score" onclick="setScore(23)">23</button>
        <button class="btn-score" onclick="setScore(22)">22</button>
        <button class="btn-score unsubmit" onclick="setScore(12)">미제출</button>
      </div>
      <button class="btn-save" id="btn_save" onclick="save()">💾 저장 후 다음 학생</button>
    </div>
  </footer>
</body>
```

---

## Task 5: Evaluator.html — JavaScript 로직

**Files:**
- Modify: `social-webapp-evaluator/Evaluator.html` (script 태그 추가)

- [ ] **Step 1: 상태 변수 및 유틸 함수**

```html
<script>
  let currentIdx   = 1;
  let selectedScore = null;
  let aiResult     = { codeStatus: '', codeAnalysis: '' };

  function toggleFullScreen() {
    if (!document.fullscreenElement) {
      document.documentElement.requestFullscreen().catch(e => alert('전체화면 실행 불가: ' + e.message));
    } else {
      document.exitFullscreen();
    }
  }

  function adjustFont(delta) {
    const root = document.documentElement;
    const cur  = parseFloat(getComputedStyle(root).getPropertyValue('--body-font-size')) || 17;
    root.style.setProperty('--body-font-size', (cur + delta) + 'px');
  }
```

- [ ] **Step 2: load 함수 — 학생 데이터 렌더링**

```javascript
  function load(idx) {
    currentIdx    = idx;
    selectedScore = null;
    aiResult      = { codeStatus: '', codeAnalysis: '' };

    document.querySelectorAll('.btn-score').forEach(b => b.classList.remove('active'));
    document.getElementById('btn_save').innerText    = '데이터 불러오는 중...';
    document.getElementById('btn_analyze').disabled  = true;
    document.getElementById('btn_analyze').innerText = '⏳ 로딩 중...';
    document.getElementById('ai_badge').className    = 'ai-badge';
    document.getElementById('ai_text').innerText     = '분석 버튼을 클릭하면 코드 유효성 및 구현 일치 여부를 자동 분석합니다.';

    google.script.run
      .withSuccessHandler(function(data) {
        if (!data) {
          alert('마지막 학생입니다.');
          document.getElementById('btn_save').innerText    = '💾 저장 후 다음 학생';
          document.getElementById('btn_analyze').disabled  = false;
          document.getElementById('btn_analyze').innerText = '🤖 AI 코드 분석';
          return;
        }

        document.getElementById('disp_name').innerText    = data.name;
        document.getElementById('disp_id').innerText      = `학번: ${data.studentId} | 순번: ${data.rowIndex} `;
        document.getElementById('prog_text').innerText    = `${data.rowIndex} / ${data.totalStudents}`;
        document.getElementById('disp_appName').innerText = data.appName;

        const linkApp  = document.getElementById('link_app');
        const linkCode = document.getElementById('link_code');
        linkApp.href  = data.appLink  || '#';
        linkCode.href = data.codeLink || '#';
        linkApp.style.opacity  = data.appLink  ? '1' : '0.4';
        linkCode.style.opacity = data.codeLink ? '1' : '0.4';

        document.getElementById('disp_problem').innerText  = data.problem;
        document.getElementById('disp_features').innerText = data.features;
        document.getElementById('disp_tools').innerText    = data.tools;
        document.getElementById('comment_input').value     = data.comment || '';

        // 기존 AI 분석 결과 복원
        if (data.codeStatus) {
          showAiResult(data.codeStatus, data.codeAnalysis);
          aiResult = { codeStatus: data.codeStatus, codeAnalysis: data.codeAnalysis };
        }

        // 기존 점수 복원
        if (data.score) setScore(parseInt(data.score));

        document.getElementById('btn_save').innerText    = '💾 저장 후 다음 학생';
        document.getElementById('btn_analyze').disabled  = false;
        document.getElementById('btn_analyze').innerText = '🤖 AI 코드 분석';
      })
      .getStudentData(idx);
  }
```

- [ ] **Step 3: runAnalysis — AI 분석 실행**

```javascript
  function runAnalysis() {
    const btn = document.getElementById('btn_analyze');
    btn.disabled  = true;
    btn.innerText = '⏳ AI 분석 중...';
    document.getElementById('ai_text').innerText = 'Drive 파일을 읽고 AI가 분석 중입니다. 잠시 기다려 주세요...';
    document.getElementById('ai_badge').className = 'ai-badge';

    google.script.run
      .withSuccessHandler(function(result) {
        aiResult = result;
        showAiResult(result.codeStatus, result.codeAnalysis);
        btn.disabled  = false;
        btn.innerText = '🔄 재분석';
      })
      .withFailureHandler(function(err) {
        document.getElementById('ai_text').innerText = '오류: ' + err.message;
        btn.disabled  = false;
        btn.innerText = '🤖 AI 코드 분석';
      })
      .analyzeCode(currentIdx);
  }

  function showAiResult(status, analysis) {
    const badge = document.getElementById('ai_badge');
    badge.innerText   = status;
    badge.className   = 'ai-badge ' + (status === '정상코드' ? 'normal' : 'error');
    document.getElementById('ai_text').innerText = analysis || '';
  }
```

- [ ] **Step 4: setScore, save, move, jump 함수**

```javascript
  function setScore(s) {
    selectedScore = s;
    document.querySelectorAll('.btn-score').forEach(b => {
      const val = parseInt(b.innerText) || 12;
      b.classList.toggle('active', val === s || (b.innerText === '미제출' && s === 12));
    });
  }

  function save() {
    if (!selectedScore) { alert('점수를 먼저 선택해 주세요.'); return; }
    document.getElementById('btn_save').innerText = '저장 중...';
    google.script.run
      .withSuccessHandler(() => load(currentIdx + 1))
      .saveResult(
        currentIdx,
        aiResult.codeStatus   || '',
        aiResult.codeAnalysis || '',
        document.getElementById('comment_input').value,
        selectedScore
      );
  }

  function move(dir) { load(currentIdx + dir); }
  function jump() {
    const val = parseInt(document.getElementById('jump_idx').value);
    if (val > 0) load(val);
  }

  window.onload = () => load(1);
</script>
```

- [ ] **Step 5: 전체 파일 통합 확인**

  위 Task 4·5의 모든 HTML 블록을 올바른 순서로 하나의 `Evaluator.html` 파일에 조합:
  ```
  <head> ... CSS ... </head>
  <body>
    header → main → ai-strip → footer
    <script> 상태변수 → toggleFullScreen/adjustFont → load → runAnalysis/showAiResult → setScore/save/move/jump → window.onload </script>
  </body>
  ```

- [ ] **Step 6: GAS 에디터에 붙여넣기 및 수동 검증**

  1. Google Sheets → 확장 프로그램 → Apps Script
  2. `Code.gs` 파일: `Code.gs` 전체 내용 붙여넣기
  3. 새 HTML 파일 생성 이름 `Evaluator` → `Evaluator.html` 전체 내용 붙여넣기
  4. 저장 (Ctrl+S) → 스프레드시트 새로고침
  5. 메뉴 `📋 수행평가 채점` → `▶ 채점 도구 열기` 클릭
  6. 확인 항목:
     - [ ] 학생 이름·학번·앱 이름 표시됨
     - [ ] 앱 링크 / 코드 링크 버튼 클릭 시 Drive 열림
     - [ ] 이전/다음 버튼으로 학생 이동
     - [ ] `🤖 AI 코드 분석` 클릭 → 배지(정상코드/오류코드) + 분석 텍스트 표시
     - [ ] 점수 버튼 클릭 시 하이라이트
     - [ ] `💾 저장 후 다음 학생` 클릭 → K~N열에 데이터 저장 확인

---

## CLAUDE.md 최종 내용

```markdown
# social-webapp-evaluator

## 프로젝트 개요
지식재산 일반 / 사회문제 해결 웹앱 제작하기 수행평가 채점 도구 (Google Apps Script)

## 배포 방법
Google Sheets → 확장 프로그램 → Apps Script
- Code.gs, Evaluator(HTML) 두 파일로 구성
- Script Properties에 OPENROUTER_API_KEY 필수 등록

## 열 구조
A: 타임스탬프 / B: 이메일 / C: 학번 / D: 이름 / E: 앱 이름
F: 앱 링크 / G: 코드링크(Drive txt) / H: 문제점/대책 / I: 핵심기능 / J: 활용프로그램
K: 코드상태(AI) / L: 코드분석(AI) / M: 교사의견 / N: 점수

## AI 모델
OpenRouter - google/gemini-2.0-flash-001
대안: anthropic/claude-haiku-4-5 (더 정확, 비용 3배)

## 점수 체계
30 / 29 / 28 / 27 / 26 / 25 / 24 / 23 / 22 / 12(미제출)
```
