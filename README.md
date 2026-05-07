# DART 감사보고서 검색기

DART(전자공시시스템)에 공시된 감사보고서·사업보고서를 **본문 키워드**와 **감사인(회계법인)별**로 동시에 필터링해 검색하고, 주석 본문을 앱 안에서 바로 열람할 수 있는 웹 애플리케이션입니다.

> DART 사이트는 감사인별 필터링 기능을 제공하지 않습니다. 본 도구는 사업보고서 XML에서 감사인명을 자동 추출하여 회계감사 실무자가 **타사의 주석 기재 사례**를 빠르게 비교 참고할 수 있도록 합니다.

---

## 주요 기능

- **본문 키워드 검색** — DART `dsab007/search.ax` 기반, AND/OR/NOT·동의어 검색 지원
- **회사명 자동 매핑** — 입력한 회사명을 DART 회사코드(CIK)로 변환해 정확도 향상
- **보고서 유형 필터** — 정기공시(A) / 외부감사(F) 등 `dsp_type` 단위 다중 선택
- **감사인 자동 추출** — 사업보고서 XML에서 "외부감사에 관한 사항" 섹션을 파싱해 감사인명 인식
- **감사인별 클릭 필터링** — 검색 결과 상단의 감사인 태그 클릭으로 즉시 필터
- **주석 본문 인앱 렌더링** — DART 뷰어 링크 + 섹션 단위 HTML 추출 제공
- **세션·캐시 최적화** — DART 웹 세션 재사용, 회사코드/감사인/XML 캐시로 응답 속도 단축

---

## 스크린샷

> _캡처 이미지를 `docs/` 디렉터리에 추가한 뒤 아래 경로를 교체해 주세요._

![검색 화면](./docs/screenshot-search.png)

---

## 기술 스택

| 구분 | 사용 기술 |
|------|-----------|
| Backend | Python 3, Flask, flask-cors, requests, BeautifulSoup4, python-dotenv |
| Frontend | 단일 HTML(`index_pdf.html`) + Tailwind CSS (CDN) + Vanilla JS |
| 외부 연동 | DART OpenAPI, DART 본문검색 웹 엔드포인트 |

---

## 설치 및 실행

```bash
# 1) 의존성 설치
pip install -r requirements.txt

# 2) 환경변수 설정 (.env 생성)
cp .env.example .env
#   → .env 파일을 열어 DART_API_KEY 값 입력

# 3) 서버 실행
python app.py
#   → http://localhost:5000 접속
```

### 환경변수

| 키 | 설명 |
|----|------|
| `DART_API_KEY` | OpenDART API 키. [opendart.fss.or.kr](https://opendart.fss.or.kr) 에서 무료 발급 |

---

## API 명세

| 메서드 | 경로 | 설명 |
|--------|------|------|
| GET  | `/api/health` | 서버 상태 및 API 키 등록 여부 확인 |
| GET  | `/api/search` | 키워드/회사명 기반 보고서 검색 (`keyword`, `corp_name`, `start_date`, `end_date`, `dsp_type`, `page`, `max_results`) |
| POST | `/api/extract-auditors` | 접수번호 배열에 대해 감사인명 일괄 추출 (`rcept_nos`, 최대 5건) |
| GET  | `/api/extract-auditor/<rcept_no>` | 단일 보고서의 감사인명 추출 |
| GET  | `/api/document/<rcept_no>` | 보고서 내 섹션(목차) 목록 |
| GET  | `/api/document/<rcept_no>/content?section=<idx>` | 특정 섹션의 HTML 본문 |

---

## 프로젝트 구조

```
.
├── app.py              # Flask 백엔드 (검색·감사인 추출·문서 조회)
├── index_pdf.html      # 단일 페이지 UI
├── requirements.txt    # Python 의존성
├── .env.example        # 환경변수 템플릿
├── PRD.md              # 제품 요구사항 정의서 (상세 명세)
└── README.md
```

---

## 참고

- 상세 요구사항·화면 설계·아키텍처는 [`PRD.md`](./PRD.md) 참고
- 본 도구는 DART OpenAPI 이용약관을 준수하며, 과도한 호출을 막기 위해 요청 간 0.35초 간격의 rate limiting 을 적용합니다.
- 자금일보 검증기는 [`자금일보 검증기.md`](./자금일보%20검증기.md) 참고.
