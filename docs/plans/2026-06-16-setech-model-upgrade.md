# 교과세특 생성 모델 분리 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 교과세특 생성 호출만 고성능 모델(Claude Sonnet)로 분리하고, 코드 분석은 기존 Gemini Flash Lite를 유지한다.

**Architecture:** `callOpenRouter(prompt)` 함수에 선택적 `model` 파라미터를 추가해 기본값은 기존 모델로 유지하고, `generateSetech()`에서만 Sonnet 모델을 명시적으로 지정한다. HTML은 변경 없음.

**Tech Stack:** Google Apps Script, OpenRouter API (`anthropic/claude-sonnet-4-5`, `google/gemini-2.5-flash-lite`)

---

## 수정 파일

| 파일 | 변경 내용 |
|------|-----------|
| `social-webapp-evaluator/Code.gs` | `callOpenRouter()` model 파라미터화 + `generateSetech()` 모델 지정 |

---

## Task 1: `callOpenRouter()` — model 파라미터 추가

**Files:**
- Modify: `social-webapp-evaluator/Code.gs` (line 107~131)

현재 코드:
```javascript
function callOpenRouter(prompt) {
  ...
  payload: JSON.stringify({
    model:      'google/gemini-2.5-flash-lite',
    messages:   [{ role: 'user', content: prompt }],
    max_tokens: 800
  }),
  ...
}
```

- [ ] **Step 1: `callOpenRouter` 시그니처에 model 파라미터 추가**

  아래 코드로 교체한다 (`callOpenRouter` 함수 전체):

  ```javascript
  function callOpenRouter(prompt, model) {
    var apiKey = PropertiesService.getScriptProperties()
                   .getProperty('OPENROUTER_API_KEY');
    if (!apiKey) throw new Error('OPENROUTER_API_KEY가 Script Properties에 설정되지 않았습니다.');

    var usedModel = model || 'google/gemini-2.5-flash-lite';

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
  ```

- [ ] **Step 2: 기존 호출부 영향 없음 확인**

  `analyzeCode()` 내의 두 호출:
  ```javascript
  var combinedRaw = callOpenRouter(combinedPrompt);        // 모델 인자 없음 → Gemini Flash Lite 유지
  relevanceAnalysis = callOpenRouter(relevancePrompt);     // 모델 인자 없음 → Gemini Flash Lite 유지
  ```
  변경 없이 기존 모델로 동작함을 확인. 두 줄 모두 그대로 둔다.

---

## Task 2: `generateSetech()` — Sonnet 모델 지정

**Files:**
- Modify: `social-webapp-evaluator/Code.gs` (generateSetech 함수 내 callOpenRouter 호출부)

현재 코드 (`generateSetech()` 하단):
```javascript
var setech = callOpenRouter(setechPrompt);
```

- [ ] **Step 1: 모델 상수를 함수 상단에 선언 후 호출에 전달**

  `generateSetech()` 함수 내 `var setechPrompt = ...` 블록 **위에** 상수를 추가하고, 호출부를 수정한다:

  ```javascript
  // 교과세특은 고품질 모델 사용 (코드 분석은 기본 Gemini Flash Lite 유지)
  var SETECH_MODEL = 'anthropic/claude-sonnet-4-5';
  ```

  그리고 기존 호출:
  ```javascript
  var setech = callOpenRouter(setechPrompt);
  ```
  를 아래로 교체:
  ```javascript
  var setech = callOpenRouter(setechPrompt, SETECH_MODEL);
  ```

- [ ] **Step 2: GAS 에디터에서 저장 후 testSetech() 함수로 동작 확인**

  GAS 에디터 하단에 아래 임시 함수를 추가해 실행:
  ```javascript
  function testSetech() {
    var result = generateSetech(1);
    Logger.log(result.setech);
    Logger.log('글자수: ' + result.setech.length);
  }
  ```

  실행 결과 확인:
  - 로그에 350~400자 범위의 한국어 세특이 출력됨
  - 스프레드시트 P열(16번째 열)에 텍스트 저장됨
  - `3차 연관성 분석 (AI 분석 버튼)`은 여전히 Gemini Flash Lite로 동작함

- [ ] **Step 3: 커밋**

  ```bash
  cd "D:\000_temp\000-clacostudy\social-webapp-evaluator"
  git add Code.gs
  git commit -m "feat: 교과세특 생성 모델을 claude-sonnet-4-5로 분리"
  git push
  ```

---

## Self-Review

**Spec coverage:**
- [x] 교과세특 생성만 Sonnet으로 분리 → Task 2 Step 1
- [x] 코드 분석은 Gemini Flash Lite 유지 → Task 1 Step 2 (기존 호출부 변경 없음)
- [x] HTML 변경 없음 → 파일 목록에 Evaluator.html 없음

**Placeholder scan:** 없음. 모든 Step에 실제 코드 포함.

**Type consistency:**
- `callOpenRouter(prompt, model)` — Task 1 Step 1(정의), Task 2 Step 1(호출) 시그니처 일치 ✅
- `SETECH_MODEL` — Task 2 Step 1에서 선언, 동일 Task에서 사용 ✅
- `generateSetech()` 반환값 `{ setech: setech }` — 변경 없음 ✅

**비용 참고:**
- Gemini 2.5 Flash Lite: ~$0.10/1M tokens
- Claude Sonnet 4.5: ~$3/1M input, $15/1M output (약 30배)
- 세특 1건 ≈ 입력 ~800토큰 + 출력 ~200토큰 → 약 $0.005/건
- 학급 30명 기준 세특 전체 생성 비용 ≈ $0.15 수준
