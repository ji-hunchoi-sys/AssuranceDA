# AI/LLM 모델 추천: 자금일보 검증기 웹 애플리케이션

| 항목 | 내용 |
|------|------|
| **작성일** | 2026-05-05 |
| **대상 프로젝트** | 자금일보(Daily Fund Report) PDF 검증기 - 단일 HTML 웹앱 |
| **작성 목적** | LLM 모델 선택, 규칙 기반 vs AI 기반 기능 분리, 아키텍처 권고 |

---

## 1. 핵심 결론 (Executive Summary)

### 권장 접근법: 하이브리드 (규칙 기반 80% + 선택적 LLM 20%)

자금일보 검증의 대부분은 **산술 검증, 합산 검증, 형식 검증**으로 구성되며, 이는 JavaScript 규칙 기반 로직으로 충분히 구현 가능하다. LLM은 **비정형 PDF 레이아웃 파싱**과 **이상치 판단의 맥락적 설명 생성**에만 선택적으로 활용하는 것이 비용, 보안, 성능 모든 측면에서 최적이다.

---

## 2. LLM 없이 JavaScript 규칙 기반으로 구현 가능한 영역

### 2.1 산술 검증 (Arithmetic Validation)
```
예) 전일잔액 + 당일입금 - 당일출금 = 당일잔액
```
- 단순 사칙연산 검증 → JavaScript로 100% 처리 가능
- 소수점 반올림 오차 허용 범위 설정 (예: ±1원)

### 2.2 합산 검증 (Summation Validation)
```
예) 각 계좌별 잔액의 합 = 총계 항목
    유동성 자금 합계 = 보통예금 + MMF + CMA + ...
```
- 행/열 단위 합계 교차 검증 → 배열 순회 합산으로 구현

### 2.3 형식 검증 (Format Validation)
- 날짜 형식 일관성 (YYYY-MM-DD, YYYY.MM.DD 등)
- 금액 단위 일관성 (천원, 백만원, 원)
- 음수/양수 부호 확인
- 빈 셀(누락 데이터) 탐지

### 2.4 기본 이상치 탐지 (Rule-based Anomaly Detection)
```javascript
// 예시: 전일 대비 변동률 기반 이상치
const changeRate = Math.abs((today - yesterday) / yesterday);
if (changeRate > 0.5) flag("전일 대비 50% 이상 변동");
```
- 임계값 기반 변동률 검증
- 마이너스 잔액 검출
- 전일 잔액 ≠ 금일 전일잔액 불일치 검출
- 특정 계정과목별 상/하한 범위 검증

### 2.5 교차 검증 (Cross-Validation)
- 자금일보 내 표 간 일관성 (예: 요약표 합계 = 상세표 합계)
- 기간 연속성 검증 (전일 당일잔액 = 금일 전일잔액)

---

## 3. LLM이 필요한 영역

### 3.1 비정형 PDF 텍스트 구조화 (높은 필요도)
- 자금일보 PDF는 회사마다 양식이 다름
- 표의 행/열 구조를 정형 데이터(JSON/배열)로 변환
- 병합 셀, 다단 레이아웃, 주석 처리

### 3.2 맥락적 이상치 판단 (중간 필요도)
- "이 변동이 정상인지"에 대한 맥락 설명 생성
- 계절적 패턴 인식 (월말 자금 집중, 급여일 등)
- 복합 조건의 이상치 판단 (단순 임계값으로 잡기 어려운 케이스)

### 3.3 자연어 검증 결과 요약 (낮은 필요도)
- 검증 결과를 사람이 읽기 쉬운 한국어로 요약
- 수정 제안 생성

---

## 4. LLM 모델 추천

### 4.1 1순위 추천: DeepSeek Chat (via OpenRouter)

| 항목 | 내용 |
|------|------|
| **모델 ID** | `deepseek/deepseek-chat` |
| **접근 방식** | OpenRouter API 경유 |
| **비용** | 입력 $0.14/M tokens, 출력 $0.28/M tokens (업계 최저 수준) |
| **컨텍스트** | 64K tokens |
| **한국어 성능** | 우수 (중국어 기반 모델로 CJK 토크나이저 효율적) |
| **적합 이유** | 구조화된 데이터 추출에 강함, 비용 효율 극대화, 한국어 금융 용어 이해도 양호 |

**OpenRouter 경유 장점:**
- 단일 API 엔드포인트로 여러 모델 접근 가능
- DeepSeek 직접 호출 대비 안정적인 가용성
- 사용량 모니터링 및 비용 제어 용이

### 4.2 2순위 추천: Claude 3.5 Haiku (via OpenRouter)

| 항목 | 내용 |
|------|------|
| **모델 ID** | `anthropic/claude-3.5-haiku` |
| **비용** | 입력 $0.80/M tokens, 출력 $4.00/M tokens |
| **컨텍스트** | 200K tokens |
| **적합 이유** | 복잡한 표 구조 파싱 정확도 높음, 긴 PDF 문서 처리 가능 |

### 4.3 3순위 추천: GPT-4o-mini (via OpenRouter)

| 항목 | 내용 |
|------|------|
| **모델 ID** | `openai/gpt-4o-mini` |
| **비용** | 입력 $0.15/M tokens, 출력 $0.60/M tokens |
| **컨텍스트** | 128K tokens |
| **적합 이유** | 범용적 성능, 안정적인 JSON 출력 포맷팅 |

### 4.4 비용 비교 (자금일보 1건 처리 기준 추정)

| 모델 | 입력 ~2K tokens | 출력 ~1K tokens | 건당 비용 | 월 20건 기준 |
|------|----------------|----------------|-----------|-------------|
| DeepSeek Chat | $0.00028 | $0.00028 | ~$0.0006 | ~$0.012 |
| Claude 3.5 Haiku | $0.0016 | $0.004 | ~$0.006 | ~$0.12 |
| GPT-4o-mini | $0.0003 | $0.0006 | ~$0.001 | ~$0.02 |

---

## 5. 클라이언트 사이드 직접 호출: 현실성 평가

### 5.1 기술적으로 가능한가?

**가능하다.** 브라우저에서 `fetch()`로 OpenRouter/OpenAI/Anthropic API를 직접 호출할 수 있다.

```javascript
// 단일 HTML 파일에서 직접 호출 예시
const response = await fetch('https://openrouter.ai/api/v1/chat/completions', {
  method: 'POST',
  headers: {
    'Authorization': `Bearer ${apiKey}`,
    'Content-Type': 'application/json',
    'HTTP-Referer': window.location.href,
    'X-Title': '자금일보 검증기'
  },
  body: JSON.stringify({
    model: 'deepseek/deepseek-chat',
    messages: [
      { role: 'system', content: SYSTEM_PROMPT },
      { role: 'user', content: extractedPdfText }
    ],
    temperature: 0.1,
    max_tokens: 4096
  })
});
```

### 5.2 API 키 보안 문제

| 접근법 | 보안 수준 | 적합 시나리오 |
|--------|----------|--------------|
| **A. 사용자 직접 입력** | 중 | 개인/소규모 팀 내부 도구 |
| **B. 환경변수 + 백엔드 프록시** | 높 | 외부 배포 서비스 |
| **C. localStorage 저장** | 낮 | 개인 PC 전용 |

#### 추천: 방식 A (사용자 API 키 직접 입력)

단일 HTML 파일 + 내부 팀 사용이라는 제약 조건에서는 **사용자가 본인의 API 키를 UI에 직접 입력**하는 방식이 가장 현실적이다.

```html
<!-- API 키 입력 UI -->
<div class="api-key-section">
  <label>OpenRouter API Key (선택 - AI 파싱 사용 시 필요)</label>
  <input type="password" id="apiKeyInput" placeholder="sk-or-v1-...">
  <small>키는 브라우저 메모리에만 유지되며, 외부로 전송되지 않습니다 (OpenRouter만 해당)</small>
</div>
```

**보안 강화 조치:**
- API 키를 변수에만 보관, localStorage에 저장하지 않음 (옵션으로 제공 가능)
- 페이지 새로고침 시 키 초기화
- OpenRouter의 "key spending limit" 기능으로 월 상한 설정 ($5~$10)
- HTTPS 환경에서만 동작하도록 권고

### 5.3 각 API 제공자별 클라이언트 사이드 호출 가능 여부

| 제공자 | 클라이언트 호출 | CORS 지원 | 비고 |
|--------|---------------|-----------|------|
| **OpenRouter** | O | O | 클라이언트 호출 공식 지원, `HTTP-Referer` 헤더 필요 |
| **OpenAI** | O | O | 가능하나 API 키 노출 리스크 |
| **Anthropic (직접)** | X | X | CORS 미지원, 서버 필수 |
| **DeepSeek (직접)** | O | O | 가능하나 OpenRouter 경유가 더 안정적 |

**결론: OpenRouter가 단일 HTML 파일 아키텍처에 가장 적합하다.**

---

## 6. 무료/로컬 모델 옵션

### 6.1 WebLLM (브라우저 내 로컬 실행)

| 항목 | 내용 |
|------|------|
| **기술** | WebGPU 기반 브라우저 내 LLM 실행 |
| **모델** | Llama 3.2 3B, Phi-3.5 mini 등 |
| **장점** | API 키 불필요, 완전 오프라인, 무료 |
| **단점** | 초기 모델 다운로드 (1~3GB), 한국어 성능 제한적, 추론 속도 느림 |
| **적합도** | 낮음 - 한국어 금융 용어 파싱에는 부족 |

### 6.2 Ollama (로컬 서버)

| 항목 | 내용 |
|------|------|
| **기술** | 로컬 LLM 서버 (별도 설치 필요) |
| **모델** | deepseek-r1:8b, llama3.2:latest 등 |
| **장점** | 무료, 데이터 외부 유출 없음 |
| **단점** | 단일 HTML 제약에 맞지 않음 (별도 프로세스 필요), GPU 필요 |
| **적합도** | 중간 - 보안이 극도로 중요한 경우 고려 |

### 6.3 평가

단일 HTML 파일이라는 제약 조건과 한국어 금융 문서 파싱 품질을 고려하면, **무료/로컬 모델보다 유료 API (DeepSeek via OpenRouter)가 월등히 적합**하다. 월 비용이 $1 미만으로 추정되므로 경제성도 충분하다.

---

## 7. 최종 권장 아키텍처: 하이브리드 접근법

```
┌─────────────────────────────────────────────────────────────────┐
│                    자금일보 검증기 (index2_pdf.html)               │
├─────────────────────────────────────────────────────────────────┤
│                                                                 │
│  [1단계] PDF 텍스트 추출 (JavaScript, 로컬)                       │
│  └─ pdf.js 라이브러리로 PDF → 텍스트 추출                          │
│                                                                 │
│  [2단계] 구조화 파싱 (분기 처리)                                   │
│  ├─ Case A: 정형 양식 → JavaScript 정규식/규칙 기반 파싱            │
│  │   (회사 표준 양식이 있는 경우, LLM 불필요)                       │
│  └─ Case B: 비정형 양식 → LLM API 호출 (DeepSeek via OpenRouter)  │
│      (다양한 양식을 처리해야 하는 경우)                              │
│                                                                 │
│  [3단계] 데이터 검증 (JavaScript, 로컬, LLM 불필요)                │
│  ├─ 산술 검증: 전일잔액 ± 입출금 = 당일잔액                        │
│  ├─ 합산 검증: 계정별 합계 = 총계                                  │
│  ├─ 형식 검증: 날짜/금액/단위 일관성                               │
│  ├─ 연속성 검증: 전일 마감 = 금일 시작                             │
│  └─ 이상치 탐지: 임계값 기반 변동률 플래그                          │
│                                                                 │
│  [4단계] 결과 리포팅 (JavaScript, 로컬)                            │
│  ├─ 검증 결과 테이블 (Pass/Fail/Warning)                          │
│  ├─ 오류 항목 하이라이트                                           │
│  └─ [선택] LLM으로 자연어 요약 생성                                │
│                                                                 │
└─────────────────────────────────────────────────────────────────┘
```

### 7.1 LLM 호출이 필요한 시점 (선택적)

| 시점 | 용도 | LLM 필요 조건 |
|------|------|--------------|
| PDF 파싱 | 비정형 표 → JSON 변환 | 회사별 양식이 표준화되지 않은 경우 |
| 이상치 설명 | "왜 비정상인지" 설명 | 사용자가 설명 요청 시 |
| 양식 미인식 | 파싱 실패 시 폴백 | 정규식 파싱이 실패한 경우 |

### 7.2 LLM 없이 동작하는 기본 모드

**LLM API 키가 없어도 기본 기능은 100% 동작해야 한다.**

- PDF 텍스트 추출 → pdf.js (로컬)
- 정형 양식 파싱 → 정규식/패턴 매칭 (로컬)
- 모든 검증 로직 → JavaScript 산술 연산 (로컬)
- 결과 표시 → DOM 렌더링 (로컬)

LLM은 "AI 보조 파싱" 버튼으로 분리하여, 사용자가 필요 시에만 활용.

---

## 8. 구현 시 프롬프트 설계 예시

### 8.1 PDF 텍스트 → 구조화 JSON 변환 프롬프트

```
당신은 한국 기업 자금팀의 자금일보(Daily Fund Report) 전문 파서입니다.
아래의 PDF에서 추출한 텍스트를 분석하여 아래 JSON 구조로 변환하세요.

## 출력 형식
{
  "report_date": "YYYY-MM-DD",
  "currency_unit": "원|천원|백만원",
  "accounts": [
    {
      "category": "계정 분류명",
      "account_name": "계좌/계정명",
      "bank": "은행명",
      "prev_balance": 숫자,
      "deposit": 숫자,
      "withdrawal": 숫자,
      "curr_balance": 숫자
    }
  ],
  "totals": {
    "prev_balance": 숫자,
    "total_deposit": 숫자,
    "total_withdrawal": 숫자,
    "curr_balance": 숫자
  }
}

## 규칙
1. 금액은 단위를 통일하여 숫자만 기재 (쉼표 제거)
2. 빈 셀은 0으로 처리
3. 확실하지 않은 값은 null로 표시
4. 병합된 셀의 경우 해당하는 모든 하위 행에 값을 복제

## 입력 텍스트:
{extracted_text}
```

### 8.2 파라미터 권장값

| 파라미터 | 값 | 이유 |
|---------|-----|------|
| `temperature` | 0.0 ~ 0.1 | 정형 데이터 추출은 결정적이어야 함 |
| `max_tokens` | 4096 | 자금일보 JSON 구조 충분 |
| `top_p` | 1.0 | temperature가 낮으므로 기본값 유지 |
| `response_format` | JSON mode (지원 시) | 파싱 안정성 확보 |

---

## 9. 리스크 및 완화 방안

| 리스크 | 영향 | 완화 방안 |
|--------|------|-----------|
| API 키가 브라우저에 노출 | 중 | 메모리만 보관, spending limit 설정, 페이지 닫으면 소멸 |
| LLM 파싱 오류 (할루시네이션) | 높 | 파싱 결과를 산술 검증으로 이중 확인, 오류 시 수동 확인 유도 |
| API 서비스 장애 | 중 | LLM 없는 기본 모드로 폴백, 정규식 파싱 우선 시도 |
| 비용 초과 | 낮 | OpenRouter spending limit, 호출 횟수 카운터 표시 |
| 한국어 금융 용어 오인식 | 중 | Few-shot 예시 포함, 계정과목 사전 제공 |

---

## 10. 최종 추천 요약

| 결정 항목 | 추천 |
|----------|------|
| **LLM 필요 여부** | 선택적 (비정형 PDF 파싱 시에만) |
| **1순위 모델** | DeepSeek Chat (`deepseek/deepseek-chat`) via OpenRouter |
| **2순위 모델** | Claude 3.5 Haiku via OpenRouter |
| **API 플랫폼** | OpenRouter (클라이언트 사이드 CORS 지원) |
| **API 키 관리** | 사용자 직접 입력 + 메모리 유지 + spending limit |
| **기본 동작** | LLM 없이 규칙 기반 100% 동작 |
| **AI 기능** | 선택적 "AI 보조 파싱" 모드로 분리 |
| **월 예상 비용** | $0.01 ~ $1.00 (사용량에 따라) |
| **PDF 파싱 라이브러리** | pdf.js (Mozilla, 무료, 로컬 실행) |

---

## 부록: 기술 참조

### OpenRouter API 기본 설정
```
Base URL: https://openrouter.ai/api/v1
Auth: Bearer token
필수 헤더: HTTP-Referer, X-Title
비용 확인: 응답의 usage.prompt_tokens, usage.completion_tokens
```

### DeepSeek 모델 비교
| 모델 | 용도 | 컨텍스트 | 비용(입력) |
|------|------|----------|-----------|
| deepseek/deepseek-chat | 범용 대화/파싱 | 64K | $0.14/M |
| deepseek/deepseek-reasoner | 복잡한 추론 | 64K | $0.55/M |

자금일보 파싱에는 `deepseek-chat`으로 충분하며, `deepseek-reasoner`는 불필요하다.

---

*본 문서는 단일 HTML 파일 아키텍처 + 한국어 자금 보고서 + 비용 효율이라는 3가지 제약 조건을 기반으로 작성되었습니다.*
