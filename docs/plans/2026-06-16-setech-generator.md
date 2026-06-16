# 교과세특 자동 생성 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 채점 도구에 교과세특 자동 생성 버튼을 추가한다. 3차(선행연구)·4차(웹앱 개발) 수행평가 내용을 OpenRouter AI로 분석해 350~400자 교과세특을 생성하고, P열(16번째)에 저장하며 재생성도 지원한다.

**Architecture:** `Code.gs`에 `generateSetech(rowIndex)` 함수 추가 — 시트1(4차 데이터)과 시트2(3차 데이터)를 통합해 OpenRouter에 생성 요청, P열 자동 저장. `Evaluator.html`에 세특 섹션(버튼+텍스트 표시) 추가 및 `load()`에서 기존 세특 복원. 블룸 분류학 4~6단계 동사·생활기록부 규칙·5단계 전개를 프롬프트에 내장.

**Tech Stack:** Google Apps Script, SpreadsheetApp, OpenRouter (`google/gemini-2.5-flash-lite`), HTML/CSS/JS

---

## 참조 규칙 (프롬프트 설계 기준)

**블룸 분류학 4~6단계 동사** (`D:\000_temp\000-clacostudy\bloom_taxonomy_verbs.md`)
- 4단계(분석): 분석·분류·추론·조사·논의·연관짓다
- 5단계(평가): 평가·검토·결정·논증·비판
- 6단계(창조): 창조·통합·설계·구성·제안·해결

**생활기록부 작성 규칙**
- 영문 약어 금지: API→외부 응용 프로그래밍 인터페이스, GAS→구글 앱스 스크립트
- 교과명 직접 나열 금지 → "범교과적으로", "융합적 관점에서"
- "~를 원용하여" 표현 적극 활용
- 수치 단순 나열 금지 → 서술형으로 맥락화
- 5단계 전개: 문제 인식 → 탐구 방법 → 구현 결과 → 분석·통찰 → 성장·마무리
- 350~400자 (한글 기준)

---

## 수정 파일 목록

| 파일 | 변경 내용 |
|------|-----------|
| `Code.gs` | `getStudentData()` P열 확장 + `generateSetech()` 추가 |
| `Evaluator.html` | 세특 섹션 CSS + HTML + `runGenerateSetech()` + `load()` 복원 로직 |

---

## 열 구조 (변경 후)

| 열 | 번호 | 내용 |
|----|------|------|
| K | 11 | 코드상태 (AI) |
| L | 12 | 코드분석 (AI) |
| M | 13 | 연관성분석 (AI) |
| N | 14 | 교사의견 |
| O | 15 | 점수 |
| **P** | **16** | **교과세특 (AI 생성)** ← 신규 |

---

## Task 1: Code.gs — getStudentData P열 확장 + generateSetech 추가

**Files:**
- Modify: `social-webapp-evaluator/Code.gs`

- [ ] **Step 1: getStudentData() 에서 16열로 확장 및 setech 필드 추가**

  기존:
  ```javascript
  var row = sheet.getRange(targetRow + 1, 1, 1, 15).getValues()[0];

  return {
    ...
    score:             row[14] || ''   // O(15)
  };
  ```

  교체 후:
  ```javascript
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
    codeStatus:        row[10] || '',
    codeAnalysis:      row[11] || '',
    relevanceAnalysis: row[12] || '',
    comment:           row[13] || '',
    score:             row[14] || '',
    setech:            row[15] || ''   // P(16): 교과세특
  };
  ```

- [ ] **Step 2: generateSetech() 함수 추가 — testAnalyze() 위에 삽입**

  ```javascript
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
  ```

- [ ] **Step 3: 수동 검증 — GAS 에디터에서 임시 테스트 함수 실행**

  ```javascript
  function testSetech() {
    var result = generateSetech(1);
    Logger.log(result.setech);
    Logger.log('글자수: ' + result.setech.length);
  }
  ```

  실행 로그에서:
  - 세특 텍스트가 자연스러운 한국어 3인칭 서술로 출력됨
  - 글자 수가 350~400 범위 내에 있음
  - 스프레드시트 P열에 텍스트가 저장됨

---

## Task 2: Evaluator.html — 세특 섹션 UI

**Files:**
- Modify: `social-webapp-evaluator/Evaluator.html`

- [ ] **Step 1: CSS — 세특 섹션 스타일 추가 (`</style>` 바로 위에 삽입)**

  ```css
  /* ── 교과세특 섹션 ── */
  .setech-section {
    border-top: 2px solid #e0e4ea;
    background: #fffdf0;
    flex-shrink: 0;
    padding: 10px 20px;
    display: flex;
    align-items: flex-start;
    gap: 14px;
    min-height: 58px;
  }
  .setech-btn-wrap {
    display: flex; flex-direction: column; gap: 4px; flex-shrink: 0;
  }
  .btn-setech {
    padding: 8px 14px; font-size: 13px; font-weight: bold;
    border: none; border-radius: 6px; cursor: pointer;
    background: #8B6914; color: white; white-space: nowrap;
  }
  .btn-setech:hover { background: #6f5310; }
  .btn-setech:disabled { background: #aaa; cursor: default; }
  .setech-charcount {
    font-size: 11px; text-align: center; font-weight: bold;
  }
  .setech-charcount.ok   { color: #155724; }
  .setech-charcount.over { color: #721c24; }
  .setech-label {
    font-size: 12px; font-weight: bold; color: #666;
    flex-shrink: 0; padding-top: 3px; min-width: 72px;
    line-height: 1.6;
  }
  .setech-text {
    font-size: 13px; color: #333; flex: 1;
    line-height: 1.7; max-height: 72px; overflow-y: auto;
    white-space: pre-wrap; word-break: keep-all;
  }
  .setech-text.placeholder { color: #aaa; font-style: italic; }
  ```

- [ ] **Step 2: HTML — 세특 섹션 마크업 추가 (`.ai-section` 닫는 태그 바로 아래)**

  ```html
  <!-- 교과세특 섹션 -->
  <div class="setech-section">
    <div class="setech-btn-wrap">
      <button class="btn-setech" id="btn_setech" onclick="runGenerateSetech()">✏️ 교과세특 작성</button>
      <span class="setech-charcount" id="setech_charcount"></span>
    </div>
    <span class="setech-label">교과세특<br>(350~400자)</span>
    <span class="setech-text placeholder" id="setech_text">
      교과세특 작성 버튼을 클릭하면 3차·4차 수행평가를 바탕으로 자동 생성됩니다.
    </span>
  </div>
  ```

- [ ] **Step 3: JS — load() 초기화 블록에 세특 리셋 추가**

  `load()` 함수 상단, 기존 초기화 블록(`document.querySelectorAll...` 이하)에 추가:

  ```javascript
  document.getElementById('btn_setech').disabled  = true;
  document.getElementById('btn_setech').innerText  = '⏳ 로딩 중...';
  document.getElementById('setech_text').innerText = '교과세특 작성 버튼을 클릭하면 3차·4차 수행평가를 바탕으로 자동 생성됩니다.';
  document.getElementById('setech_text').className = 'setech-text placeholder';
  document.getElementById('setech_charcount').innerText = '';
  document.getElementById('setech_charcount').className = 'setech-charcount';
  ```

- [ ] **Step 4: JS — load() withSuccessHandler에 세특 복원 블록 추가**

  `// 기존 점수 복원` 블록 바로 아래에 추가:

  ```javascript
  // 기존 세특 복원
  const setechEl    = document.getElementById('setech_text');
  const setechCount = document.getElementById('setech_charcount');
  const setechBtn   = document.getElementById('btn_setech');
  if (data.setech) {
    setechEl.innerText  = data.setech;
    setechEl.className  = 'setech-text';
    const len = data.setech.length;
    setechCount.innerText = len + '자';
    setechCount.className = 'setech-charcount ' + (len >= 350 && len <= 400 ? 'ok' : 'over');
    setechBtn.innerText = '🔄 재생성';
  } else {
    setechEl.innerText  = '교과세특 작성 버튼을 클릭하면 3차·4차 수행평가를 바탕으로 자동 생성됩니다.';
    setechEl.className  = 'setech-text placeholder';
    setechCount.innerText = '';
    setechBtn.innerText = '✏️ 교과세특 작성';
  }
  setechBtn.disabled = false;
  ```

- [ ] **Step 5: JS — runGenerateSetech() 함수 추가 (moveToUnscored() 아래에 삽입)**

  ```javascript
  function runGenerateSetech() {
    const btn      = document.getElementById('btn_setech');
    const textEl   = document.getElementById('setech_text');
    const countEl  = document.getElementById('setech_charcount');
    btn.disabled   = true;
    btn.innerText  = '⏳ 생성 중...';
    textEl.innerText = '3차·4차 수행평가 내용을 분석하고 교과세특을 작성 중입니다...';
    textEl.className = 'setech-text placeholder';
    countEl.innerText = '';

    google.script.run
      .withSuccessHandler(function(result) {
        const text = result.setech || '';
        textEl.innerText  = text;
        textEl.className  = 'setech-text';
        const len = text.length;
        countEl.innerText = len + '자';
        countEl.className = 'setech-charcount ' + (len >= 350 && len <= 400 ? 'ok' : 'over');
        btn.disabled  = false;
        btn.innerText = '🔄 재생성';
      })
      .withFailureHandler(function(err) {
        textEl.innerText  = '오류: ' + err.message;
        textEl.className  = 'setech-text';
        btn.disabled  = false;
        btn.innerText = '✏️ 교과세특 작성';
      })
      .generateSetech(currentIdx);
  }
  ```

- [ ] **Step 6: 수동 검증 체크리스트**

  GAS 에디터 업데이트 후 도구 열어 확인:
  - [ ] `✏️ 교과세특 작성` 버튼이 AI 분석 영역 아래, 푸터 위에 표시됨
  - [ ] 버튼 클릭 시 "⏳ 생성 중..." 상태로 변환됨
  - [ ] 생성 완료 후 세특 텍스트가 표시되고 글자 수 배지가 나타남
  - [ ] 350~400자 범위이면 초록색, 범위 벗어나면 빨간색으로 글자 수 표시
  - [ ] 스프레드시트 P열에 세특 텍스트가 저장됨
  - [ ] 다른 학생으로 이동 후 돌아왔을 때 기존 세특이 자동 복원됨
  - [ ] 복원 시 버튼이 `🔄 재생성`으로 표시됨
  - [ ] 재생성 버튼 클릭 시 새 텍스트로 덮어씀

---

## Self-Review

**Spec coverage 체크:**
- [x] 3차+4차 연계 세특 작성 → `generateSetech()` 프롬프트에 두 차시 데이터 통합
- [x] 블룸 동사 4~6단계 → 프롬프트 규칙에 명시
- [x] 350~400자 → 프롬프트 지침 + 글자 수 배지로 시각 확인
- [x] P열(점수 오른쪽) 저장 → `sheet.getRange(row, 16).setValue(setech)`
- [x] 버튼 별도 생성 → `✏️ 교과세특 작성` 버튼
- [x] 재생성 → 동일 버튼이 `🔄 재생성`으로 전환, P열 덮어쓰기
- [x] 기존 세특 복원 → `load()` withSuccessHandler에서 `data.setech` 확인

**Placeholder scan:** 모든 Step에 실제 코드 포함 ✅

**Type consistency:**
- `generateSetech(rowIndex)` — Task 1 Step 2(서버), Task 2 Step 5(클라이언트 호출) 동일 이름 ✅
- `result.setech` — Task 1 반환 `{ setech: setech }`, Task 2 Step 5에서 `result.setech` 참조 ✅
- `data.setech` — Task 1 Step 1에서 `getStudentData()` 반환에 `setech: row[15]` 추가, Task 2 Step 4에서 `data.setech` 참조 ✅
- `setechBtn.disabled = false` — Task 2 Step 3(로딩 중 비활성), Step 4(로드 완료 후 활성), Step 5(생성 완료 후 활성) 일관 ✅

**생활기록부 규칙 프롬프트 반영 확인:**
- 영문 약어 금지 → 프롬프트에 명시 (API·GAS·AI 예시 포함) ✅
- 교과명 나열 금지 → 프롬프트에 명시 ✅
- '원용하다' 표현 → 프롬프트에 명시 ✅
- 5단계 전개 → 프롬프트에 명시 ✅
- 3인칭 서술 → 학생 이름 포함하여 명시 ✅
