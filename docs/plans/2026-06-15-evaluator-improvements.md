# Evaluator Improvements Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 채점 도구의 6가지 항목을 개선한다: Drive 파일 크기 대응, API 호출 최적화(3→2회), 연관성 배지 판단 정확도 개선, 연관성 분석 결과 O열 저장, 채점 완료 표시, 미채점 바로가기 버튼.

**Architecture:** `Code.gs`(서버)와 `Evaluator.html`(클라이언트) 두 파일만 수정. 각 Task는 독립적으로 적용 가능하며 GAS 에디터에서 붙여넣기로 반영한다. 실행 순서: Task 7 → 1 → 2 → 3 → 4 → 5 (의존성 최소화 순).

**Tech Stack:** Google Apps Script, SpreadsheetApp, DriveApp, UrlFetchApp (OpenRouter `google/gemini-2.5-flash-lite`), HTML/CSS/JS

---

## 수정 파일 목록

| 파일 | 수정 내용 |
|------|-----------|
| `Code.gs` | Task 7: 파일 크기 절단 / Task 1: 프롬프트 통합 / Task 2: 연관성 프롬프트 / Task 3: saveResult·getStudentData 시그니처 변경 / Task 5: getNextUnscored 추가 |
| `Evaluator.html` | Task 2: showRelevance 파싱 변경 / Task 3: save()·load() 업데이트 / Task 4: 채점완료 배지 / Task 5: 미채점 버튼·moveToUnscored() |

---

## Task 7: Drive 파일 크기 초과 대응

**Files:**
- Modify: `Code.gs` — `analyzeCode()` 내 codeText 빌드 직후

코드가 너무 길면 OpenRouter API 토큰 한도를 초과해 오류가 발생한다. 최대 15,000자로 절단하고 프롬프트에 절단 사실을 명시한다.

- [ ] **Step 1: analyzeCode()에서 codeText 생성 직후 절단 로직 추가**

  기존 코드에서 `var codeText = codeParts.join('\n\n');` 바로 아래에 다음을 추가:

  ```javascript
  var MAX_CODE_CHARS = 15000;
  var truncated = codeText.length > MAX_CODE_CHARS;
  if (truncated) {
    codeText = codeText.substring(0, MAX_CODE_CHARS) +
      '\n\n[주의: 코드가 너무 길어 앞부분 ' + MAX_CODE_CHARS + '자까지만 분석합니다]';
  }
  ```

- [ ] **Step 2: 수동 검증**

  GAS 에디터에서 임시 함수로 긴 코드 제출 학생의 `analyzeCode(n)` 호출 후 오류 없이 결과가 반환되면 정상.

---

## Task 1: API 호출 3번 → 2번 최적화

**Files:**
- Modify: `Code.gs` — `analyzeCode()` 내 프롬프트 및 파싱 로직

현재 "코드 유효성"과 "구현 분석"을 별도 API 호출로 처리한다. 하나의 프롬프트로 합쳐 응답 첫 줄에 `정상코드`/`오류코드`, 이후 줄에 분석 텍스트를 받도록 변경한다.

- [ ] **Step 1: validityPrompt + analysisPrompt를 combinedPrompt 하나로 교체**

  기존:
  ```javascript
  var validityPrompt = '...';
  var analysisPrompt = '...';
  var statusRaw    = callOpenRouter(validityPrompt);
  var codeStatus   = statusRaw.includes('정상') ? '정상코드' : '오류코드';
  var codeAnalysis = callOpenRouter(analysisPrompt);
  ```

  교체 후:
  ```javascript
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
  ```

- [ ] **Step 2: callOpenRouter의 max_tokens를 800으로 상향**

  합쳐진 응답이 잘리지 않도록:
  ```javascript
  max_tokens: 800
  ```

- [ ] **Step 3: 수동 검증**

  `testAnalyze()` 실행 후 로그에서 `codeStatus`가 `정상코드` 또는 `오류코드`이고 `codeAnalysis`에 분석 텍스트가 있으면 정상.

---

## Task 2: 연관성 배지 키워드 판단 개선

**Files:**
- Modify: `Code.gs` — `analyzeCode()` 내 `relevancePrompt`
- Modify: `Evaluator.html` — `showRelevance()` 파싱 로직

현재 AI 응답에 "높음"/"낮음" 단어가 없으면 배지가 항상 `보통`으로 표시된다. 프롬프트로 응답 첫 줄을 `[높음]`/`[보통]`/`[낮음]` 중 하나로 강제하고, JS에서 대괄호 형식으로 정확히 파싱한다.

- [ ] **Step 1: Code.gs — relevancePrompt 마지막 지시문 교체**

  기존:
  ```javascript
  '연관성 수준(높음/보통/낮음)을 먼저 밝히고, 이유를 200자 이내 한국어로 작성하세요.\n\n' +
  ```

  교체 후:
  ```javascript
  '반드시 첫 번째 줄에 [높음], [보통], [낮음] 중 정확히 하나만 작성하고, ' +
  '두 번째 줄부터 이유를 200자 이내 한국어로 작성하세요.\n\n' +
  ```

- [ ] **Step 2: Evaluator.html — showRelevance() 내 배지 판단 로직 교체**

  기존:
  ```javascript
  const text  = relevanceText;
  let level = 'medium';
  if (text.includes('높음') || text.includes('매우 높') || text.includes('잘 부합') || text.includes('직접적')) level = 'high';
  else if (text.includes('낮음') || text.includes('관련 없') || text.includes('무관')) level = 'low';
  ```

  교체 후:
  ```javascript
  const firstLine = relevanceText.split('\n')[0];
  let level = 'medium';
  if (firstLine.includes('[높음]')) level = 'high';
  else if (firstLine.includes('[낮음]')) level = 'low';
  const bodyText = relevanceText.split('\n').slice(1).join('\n').trim();
  document.getElementById('relevance_text').innerText = bodyText || relevanceText;
  ```

  그리고 아래에 있던 `document.getElementById('relevance_text').innerText = relevanceText;` 줄은 삭제 (위에서 처리).

- [ ] **Step 3: 수동 검증**

  AI 분석 실행 후 연관성 배지가 `[높음]`/`[보통]`/`[낮음]` 중 하나로 정확히 표시되고, 배지 아래 텍스트에 대괄호 줄이 제거된 본문만 표시되면 정상.

---

## Task 3: 연관성 분석 결과 O열 저장

**Files:**
- Modify: `Code.gs` — `getStudentData()`, `saveResult()` 시그니처 변경
- Modify: `Evaluator.html` — `save()`, `load()` 업데이트

현재 연관성 분석 텍스트는 화면에만 표시되고 다음 학생으로 넘어가면 사라진다. O열(인덱스 14)에 저장해 재방문 시 복원한다.

- [ ] **Step 1: Code.gs — getStudentData() 에서 15개 열(A~O)로 확장**

  기존:
  ```javascript
  var row = sheet.getRange(targetRow + 1, 1, 1, 14).getValues()[0];
  ```
  교체 후:
  ```javascript
  var row = sheet.getRange(targetRow + 1, 1, 1, 15).getValues()[0];
  ```

  반환 객체에 `relevanceAnalysis` 추가:
  ```javascript
  relevanceAnalysis: row[14] || '',   // O열 (인덱스 14)
  ```

- [ ] **Step 2: Code.gs — saveResult() 시그니처에 relevanceAnalysis 추가**

  기존:
  ```javascript
  function saveResult(rowIndex, codeStatus, codeAnalysis, comment, score) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    var actualRow = parseInt(rowIndex) + 1;
    sheet.getRange(actualRow, 11).setValue(codeStatus   || '');  // K열
    sheet.getRange(actualRow, 12).setValue(codeAnalysis || '');  // L열
    sheet.getRange(actualRow, 13).setValue(comment      || '');  // M열
    sheet.getRange(actualRow, 14).setValue(score        || '');  // N열
    return true;
  }
  ```

  교체 후:
  ```javascript
  function saveResult(rowIndex, codeStatus, codeAnalysis, relevanceAnalysis, comment, score) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    var actualRow = parseInt(rowIndex) + 1;
    sheet.getRange(actualRow, 11).setValue(codeStatus        || '');  // K열
    sheet.getRange(actualRow, 12).setValue(codeAnalysis      || '');  // L열
    sheet.getRange(actualRow, 13).setValue(relevanceAnalysis || '');  // O열 → 실제론 M열 다음이므로 주의: 아래 참고
    sheet.getRange(actualRow, 14).setValue(comment           || '');  // N열(교사의견 → P열로 밀림 방지 위해 열 번호 재확인)
    sheet.getRange(actualRow, 15).setValue(score             || '');  // O열
    return true;
  }
  ```

  > ⚠️ **열 번호 주의:** O열을 기존 M(교사의견), N(점수) 뒤에 추가하면 점수가 O열로 밀린다. 아래와 같이 열 순서를 정리한다:
  > - K(11): 코드상태
  > - L(12): 코드분석
  > - M(13): **연관성분석** ← 새로 삽입
  > - N(14): 교사의견
  > - O(15): 점수

  실제 최종 saveResult():
  ```javascript
  function saveResult(rowIndex, codeStatus, codeAnalysis, relevanceAnalysis, comment, score) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    var actualRow = parseInt(rowIndex) + 1;
    sheet.getRange(actualRow, 11).setValue(codeStatus        || '');  // K: 코드상태
    sheet.getRange(actualRow, 12).setValue(codeAnalysis      || '');  // L: 코드분석
    sheet.getRange(actualRow, 13).setValue(relevanceAnalysis || '');  // M: 연관성분석
    sheet.getRange(actualRow, 14).setValue(comment           || '');  // N: 교사의견
    sheet.getRange(actualRow, 15).setValue(score             || '');  // O: 점수
    return true;
  }
  ```

- [ ] **Step 3: Code.gs — getStudentData() 반환 객체 열 인덱스 업데이트**

  기존 열 구조가 바뀌었으므로:
  ```javascript
  codeStatus:        row[10] || '',   // K(11)
  codeAnalysis:      row[11] || '',   // L(12)
  relevanceAnalysis: row[12] || '',   // M(13) ← 새로 추가
  comment:           row[13] || '',   // N(14)
  score:             row[14] || ''    // O(15)
  ```

  > ⚠️ 기존 스프레드시트에 이미 M열(교사의견), N열(점수) 데이터가 있다면 열을 직접 시트에서 수동으로 조정해야 한다. GAS 에디터 적용 전에 시트에서 K열 다음에 빈 열 하나를 삽입(우클릭 → 열 삽입)하면 기존 데이터가 밀리지 않는다.

- [ ] **Step 4: Evaluator.html — save() 에 relevanceAnalysis 인자 추가**

  기존:
  ```javascript
  .saveResult(
    currentIdx,
    aiResult.codeStatus   || '',
    aiResult.codeAnalysis || '',
    document.getElementById('comment_input').value,
    selectedScore
  );
  ```

  교체 후:
  ```javascript
  .saveResult(
    currentIdx,
    aiResult.codeStatus        || '',
    aiResult.codeAnalysis      || '',
    aiResult.relevanceAnalysis || '',
    document.getElementById('comment_input').value,
    selectedScore
  );
  ```

- [ ] **Step 5: Evaluator.html — load() 에서 기존 연관성 분석 복원**

  기존 AI 분석 결과 복원 블록:
  ```javascript
  if (data.codeStatus) {
    showAiResult(data.codeStatus, data.codeAnalysis);
    aiResult = { codeStatus: data.codeStatus, codeAnalysis: data.codeAnalysis, sheet2Content: '', relevanceAnalysis: '' };
  }
  ```

  교체 후:
  ```javascript
  if (data.codeStatus) {
    showAiResult(data.codeStatus, data.codeAnalysis);
    aiResult = {
      codeStatus:        data.codeStatus,
      codeAnalysis:      data.codeAnalysis,
      sheet2Content:     '',
      relevanceAnalysis: data.relevanceAnalysis || ''
    };
    if (data.relevanceAnalysis) {
      showRelevance('', data.relevanceAnalysis);
    }
  }
  ```

- [ ] **Step 6: 수동 검증**

  1. 학생 1명 AI 분석 실행 → 저장
  2. 스프레드시트 M열에 연관성 분석 텍스트, N열에 교사의견, O열에 점수가 저장되었는지 확인
  3. 도구를 닫고 다시 열어 같은 학생으로 이동 → 연관성 분석 내용이 화면에 복원되는지 확인

---

## Task 4: 채점 완료 여부 표시

**Files:**
- Modify: `Evaluator.html` — 헤더 HTML + `load()` JS

점수(O열)가 있는 학생을 로드했을 때 헤더에 `✅ 채점완료` 배지를 표시한다.

- [ ] **Step 1: 헤더 HTML에 채점완료 배지 추가**

  기존 `.student-info` 블록:
  ```html
  <div class="student-info">
    <h2 id="disp_name">학생 성명</h2>
    <span id="disp_id">학번: ----- | 로딩 중...</span>
  </div>
  ```

  교체 후:
  ```html
  <div class="student-info">
    <h2 id="disp_name">학생 성명</h2>
    <span id="disp_id">학번: ----- | 로딩 중...</span>
    <span id="scored_badge" style="display:none; background:#28a745; color:white;
      font-size:12px; font-weight:bold; padding:3px 10px; border-radius:12px; margin-left:8px;">
      ✅ 채점완료
    </span>
  </div>
  ```

- [ ] **Step 2: CSS에 scored-badge 호버 스타일 추가 (선택)**

  `<style>` 블록 내 아무 위치:
  ```css
  #scored_badge { vertical-align: middle; }
  ```

- [ ] **Step 3: load() 에서 배지 표시/숨김 처리**

  `load()` 의 `withSuccessHandler` 내부, 학생 이름을 렌더링하는 블록 바로 아래에 추가:
  ```javascript
  document.getElementById('scored_badge').style.display =
    data.score ? 'inline-block' : 'none';
  ```

- [ ] **Step 4: 수동 검증**

  채점된 학생으로 이동 시 헤더에 `✅ 채점완료` 배지가 초록색으로 표시되고, 미채점 학생에서는 표시되지 않으면 정상.

---

## Task 5: 미채점 학생 바로가기 버튼

**Files:**
- Modify: `Code.gs` — `getNextUnscored()` 함수 추가
- Modify: `Evaluator.html` — 헤더 버튼 추가 + `moveToUnscored()` 함수 추가

현재 미채점 학생을 찾으려면 수동으로 이전/다음을 눌러야 한다. `⏭ 미채점` 버튼 클릭 시 현재 위치 이후의 첫 번째 미채점 학생(O열 비어있는 행)으로 이동한다.

- [ ] **Step 1: Code.gs — getNextUnscored() 추가**

  `testAnalyze()` 함수 위에 삽입:
  ```javascript
  function getNextUnscored(fromRowIndex) {
    var sheet   = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return null;

    for (var i = fromRowIndex + 1; i <= lastRow - 1; i++) {
      var score = sheet.getRange(i + 1, 15).getValue(); // O열(15번째)
      if (!score && score !== 0) return i;
    }
    return null;
  }
  ```

- [ ] **Step 2: Evaluator.html — 헤더 네비게이션에 미채점 버튼 추가**

  기존 `.nav-btns` 내 `다음 ▶` 버튼 바로 뒤에 추가:
  ```html
  <button class="btn-nav" id="btn_unscored"
    style="background:#e67e22; color:white; border:none;"
    onclick="moveToUnscored()">⏭ 미채점</button>
  ```

- [ ] **Step 3: Evaluator.html — moveToUnscored() 함수 추가**

  `jump()` 함수 바로 아래에 추가:
  ```javascript
  function moveToUnscored() {
    const btn     = document.getElementById('btn_unscored');
    btn.disabled  = true;
    btn.innerText = '⏳ 검색 중...';
    google.script.run
      .withSuccessHandler(function(nextIdx) {
        btn.disabled  = false;
        btn.innerText = '⏭ 미채점';
        if (nextIdx) {
          load(nextIdx);
        } else {
          alert('🎉 현재 위치 이후 미채점 학생이 없습니다. 처음으로 이동합니다.');
          load(1);
        }
      })
      .withFailureHandler(function(err) {
        btn.disabled  = false;
        btn.innerText = '⏭ 미채점';
        alert('오류: ' + err.message);
      })
      .getNextUnscored(currentIdx);
  }
  ```

- [ ] **Step 4: 수동 검증**

  1. 몇 명을 채점한 상태에서 `⏭ 미채점` 버튼 클릭
  2. O열이 비어있는 첫 번째 학생으로 이동하면 정상
  3. 모두 채점된 상태에서 클릭 시 알림창 후 1번 학생으로 이동하면 정상

---

## Self-Review

**Spec coverage 체크:**
- [x] Task 7: Drive 파일 크기 초과 대응 → `analyzeCode()` 절단 로직
- [x] Task 1: API 호출 3→2번 → combinedPrompt + 파싱
- [x] Task 2: 배지 판단 개선 → `[높음]`/`[보통]`/`[낮음]` 강제 + 파싱
- [x] Task 3: O열 저장 → `saveResult()` 시그니처 변경 + `getStudentData()` 확장
- [x] Task 4: 채점완료 배지 → `scored_badge` HTML + `load()` 업데이트
- [x] Task 5: 미채점 바로가기 → `getNextUnscored()` + `moveToUnscored()`

**⚠️ 주의: Task 3 적용 전 스프레드시트 열 조정 필수**

기존 시트에 M열(교사의견), N열(점수) 데이터가 있다면:
1. GAS 에디터 업데이트 **전에** 스프레드시트에서 L열 오른쪽(M열 위치)에 빈 열 삽입
2. 그러면 기존 교사의견 → N열, 기존 점수 → O열로 자동으로 밀림
3. 새 M열이 연관성분석 열이 됨
4. 이후 GAS 코드 적용

**Placeholder scan:** 모든 Step에 실제 코드 포함 확인 ✅

**Type consistency:**
- `saveResult(rowIndex, codeStatus, codeAnalysis, relevanceAnalysis, comment, score)` — Task 3 Step 2, Step 4에서 동일 시그니처 사용 ✅
- `getNextUnscored(fromRowIndex)` — Task 5 Step 1(서버), Step 3(클라이언트 호출) 동일 이름 ✅
- `data.score` → Task 3 Step 3에서 O열(row[14])로 변경, Task 4 Step 3에서 동일 `data.score` 참조 ✅
