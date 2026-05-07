# -*- coding: utf-8 -*-
"""
DART 감사보고서 본문검색 Flask 백엔드
- DART 본문검색(dsab007/search.ax) 기반 키워드 검색
- 사업보고서(A-type) 내 감사인 자동 추출
- 감사인별 필터링 기능
"""

import os
import io
import time
import logging
import re
import zipfile
import threading
from datetime import datetime, timedelta

import requests
from bs4 import BeautifulSoup
from dotenv import load_dotenv
from flask import Flask, request as flask_request, jsonify, send_from_directory
from flask_cors import CORS

load_dotenv()
API_KEY = os.environ.get("DART_API_KEY", "")
DART_API_BASE = "https://opendart.fss.or.kr/api"
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

app = Flask(__name__, static_folder=BASE_DIR)
CORS(app)
app.config["JSON_AS_ASCII"] = False

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger(__name__)

# ── 캐시 ──
_auditor_cache = {}
_xml_cache = {}
_XML_CACHE_MAX = 30
_XML_CACHE_TTL = 600

# 감사보고서/연결감사보고서 섹션 본문 텍스트 캐시 (per rcept_no)
_audit_text_cache = {}
_AUDIT_TEXT_CACHE_MAX = 200

# ── DART 웹 세션 ──
_dart_session = None
_dart_session_lock = threading.Lock()
_dart_session_time = 0
_DART_SESSION_TTL = 600

# ── Rate limiting ──
_last_api_call = 0.0
_last_dart_call = 0.0
API_CALL_INTERVAL = 0.35
_api_rate_lock = threading.Lock()
_dart_rate_lock = threading.Lock()


def _rate_limit_api():
    global _last_api_call
    with _api_rate_lock:
        elapsed = time.time() - _last_api_call
        if elapsed < API_CALL_INTERVAL:
            time.sleep(API_CALL_INTERVAL - elapsed)
        _last_api_call = time.time()


def _rate_limit_dart():
    global _last_dart_call
    with _dart_rate_lock:
        elapsed = time.time() - _last_dart_call
        if elapsed < API_CALL_INTERVAL:
            time.sleep(API_CALL_INTERVAL - elapsed)
        _last_dart_call = time.time()


def _create_dart_session():
    """새 DART 웹 세션을 생성한다."""
    s = requests.Session()
    s.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Accept-Language": "ko-KR,ko;q=0.9,en;q=0.8",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Referer": "https://dart.fss.or.kr/dsab007/main.do",
    })
    s.get("https://dart.fss.or.kr/", timeout=15)
    s.post("https://dart.fss.or.kr/dsab007/detailSearchMain2.do",
           data={"autoSearch": "Y", "option": "contents"}, timeout=15)
    return s


def _get_dart_session(force_new=False):
    """DART 웹 세션을 반환한다."""
    global _dart_session, _dart_session_time
    with _dart_session_lock:
        now = time.time()
        if not force_new and _dart_session and (now - _dart_session_time) < _DART_SESSION_TTL:
            return _dart_session
        try:
            _dart_session = _create_dart_session()
            _dart_session_time = now
            logger.info("DART 웹 세션 생성 완료")
        except Exception as exc:
            logger.warning("DART 세션 초기화 실패: %s", exc)
            if _dart_session is None:
                raise
        return _dart_session


_corp_cik_cache = {}


def _lookup_corp_cik(corp_name):
    """DART 회사코드(CIK)를 조회한다."""
    if not corp_name:
        return ""
    if corp_name in _corp_cik_cache:
        return _corp_cik_cache[corp_name]
    try:
        session = _get_dart_session()
        _rate_limit_dart()
        resp = session.post(
            "https://dart.fss.or.kr/corp/searchExistAll.ax",
            data={"textCrpNm": corp_name},
            timeout=10,
        )
        cik = resp.text.strip().rstrip(",").strip()
        if cik and cik != "null":
            _corp_cik_cache[corp_name] = cik
            logger.info("회사코드 조회: %s → %s", corp_name, cik)
            return cik
        _corp_cik_cache[corp_name] = ""
        logger.info("회사코드 없음 (다중 매칭 또는 미존재): %s", corp_name)
        return ""
    except Exception as exc:
        logger.warning("회사코드 조회 실패 (%s): %s", corp_name, exc)
        return ""


def _search_dart(keyword, start_date, end_date, page=1, max_results=15,
                 dsp_types=None, corp_name=""):
    """DART 본문검색을 수행한다. 실패 시 세션을 재생성하여 1회 재시도."""
    corp_cik = _lookup_corp_cik(corp_name) if corp_name else ""

    # 멀티 키워드(공백으로 분리된 2개 이상)일 때는 synonym 확장을 끄고 body-only 옵션을 명시해
    # DART가 더 strict한 AND 매칭을 하도록 유도한다. 단일 키워드일 때는 기존 동작(synonym=Y) 유지.
    keyword_parts = keyword.strip().split()
    is_multi_kw = len(keyword_parts) > 1

    form_data = {
        "currentPage": str(page),
        "maxResults": str(max_results),
        "maxLinks": "10",
        "sort": "",
        "sortType": "",
        "keyword": keyword,
        "synonym": "N" if is_multi_kw else "Y",
        "textCrpNm": corp_name,
        "textCrpCik": corp_cik,
        "textPresenterNm": "",
        "startDate": start_date.replace("-", ""),
        "endDate": end_date.replace("-", ""),
        "docType": "",
        "reportName": "",
        "autoSearch": "N",
        "option": "contents" if is_multi_kw else "",
        "tocSrch": "",
    }
    if dsp_types:
        form_data["dspType"] = dsp_types
    else:
        form_data["dspType"] = ["A", "F"]

    timeouts = [45, 30]
    for attempt in range(2):
        try:
            session = _get_dart_session(force_new=(attempt > 0))
            _rate_limit_dart()
            resp = session.post(
                "https://dart.fss.or.kr/dsab007/search.ax",
                data=form_data, timeout=timeouts[attempt],
            )
            resp.raise_for_status()
            resp.encoding = "utf-8"

            if len(resp.text) < 100 or "tbWideList" not in resp.text:
                if attempt == 0:
                    logger.warning("DART 응답이 비정상입니다, 세션 재생성 후 재시도")
                    continue
                return {"results": [], "page_info": {"total": 0, "has_next": False, "has_prev": False, "total_pages": 0}}

            return _parse_search_results(resp.text)

        except requests.exceptions.Timeout:
            if attempt == 0:
                logger.warning("DART 요청 타임아웃 (%ds), 세션 재생성 후 재시도", timeouts[attempt])
                continue
            raise Exception("DART 서버 응답 시간이 초과되었습니다. 검색 기간을 줄이거나 잠시 후 다시 시도해주세요.")
        except requests.exceptions.ConnectionError:
            if attempt == 0:
                logger.warning("DART 연결 오류, 재시도")
                time.sleep(1)
                continue
            raise Exception("DART 서버에 연결할 수 없습니다. 네트워크 상태를 확인해주세요.")
        except Exception as exc:
            if attempt == 0 and "세션" not in str(exc):
                logger.warning("DART 검색 실패 (%s), 세션 재생성 후 재시도", exc)
                continue
            raise

    return {"results": [], "page_info": {"total": 0, "has_next": False, "has_prev": False, "total_pages": 0}}


def _parse_search_results(html):
    """본문검색 HTML 응답을 파싱한다."""
    soup = BeautifulSoup(html, "html.parser")

    total = 0
    count_elem = soup.select_one("#searchCnt")
    if count_elem:
        m = re.search(r"([\d,]+)", count_elem.get_text())
        if m:
            total = int(m.group(1).replace(",", ""))

    results = []
    rows = soup.select("table.tbWideList tbody tr")

    for row in rows:
        th = row.find("th")
        if not th:
            continue

        company_a = th.find("a", class_="company")
        company_name = company_a.get_text(strip=True) if company_a else ""
        corp_code = ""
        if company_a:
            m = re.search(r"'(\d+)'", company_a.get("href", ""))
            if m:
                corp_code = m.group(1)

        market_span = th.find("span", class_=re.compile("tagCom"))
        market = ""
        if market_span:
            title = market_span.get("title", "")
            if "코스피" in title:
                market = "코스피"
            elif "코스닥" in title:
                market = "코스닥"
            elif "코넥스" in title:
                market = "코넥스"
            else:
                market = title or market_span.get_text(strip=True)

        report_a = th.find("a", class_="second")
        report_name = ""
        rcept_no = ""
        dcm_no = ""
        if report_a:
            report_name = re.sub(r"\s+", " ", report_a.get_text(strip=True))
            href = report_a.get("href", "")
            m = re.search(r"rcpNo=(\d+)", href)
            if m:
                rcept_no = m.group(1)
            m2 = re.search(r"dcmNo=(\d+)", href)
            if m2:
                dcm_no = m2.group(1)

        info_td = row.find("td", class_="info")
        info_text = re.sub(r"\s+", " ", info_td.get_text(strip=True)) if info_td else ""

        date_td = row.find("td", class_="date")
        date_text = date_td.get_text(strip=True) if date_td else ""

        filing_type = ""
        submitter = ""
        if info_text:
            ft_match = re.search(r"\[([^\]]+)\]", info_text)
            if ft_match:
                filing_type = ft_match.group(1)
            sub_match = re.search(r"제출인\s*:\s*(.+?)(?:\s*$)", info_text)
            if sub_match:
                submitter = sub_match.group(1).strip()

        auditor = ""
        if "외부감사" in filing_type:
            auditor = submitter
        elif rcept_no in _auditor_cache:
            auditor = _auditor_cache[rcept_no]

        snippet = ""
        for td in row.find_all("td"):
            cls = td.get("class")
            if cls and ("info" in cls or "date" in cls):
                continue
            if td.find("strong"):
                snippet = re.sub(r"\s+", " ", td.get_text(strip=True))[:200]
                break

        results.append({
            "rcept_no": rcept_no,
            "dcm_no": dcm_no,
            "corp_code": corp_code,
            "company_name": company_name,
            "market": market,
            "report_name": report_name,
            "filing_type": filing_type,
            "submitter": submitter,
            "auditor": auditor,
            "date": date_text,
            "snippet": snippet,
            "dart_url": f"https://dart.fss.or.kr/dsaf001/main.do?rcpNo={rcept_no}" if rcept_no else "",
        })

    current_page = 1
    total_pages = 0
    page_info = {"total": total, "has_next": False, "has_prev": False, "total_pages": 0}

    info_div = soup.select_one(".pageInfo")
    if info_div:
        m = re.search(r"\[(\d+)/(\d+)\]", info_div.get_text())
        if m:
            current_page = int(m.group(1))
            total_pages = int(m.group(2))
            page_info["total_pages"] = total_pages
            page_info["has_prev"] = current_page > 1
            page_info["has_next"] = current_page < total_pages
        m2 = re.search(r"총\s*([\d,]+)건", info_div.get_text())
        if m2:
            page_info["total"] = int(m2.group(1).replace(",", ""))

    return {"results": results, "page_info": page_info}


# ── 감사인 추출 ──

def _download_document_xml(rcept_no):
    now = time.time()
    if rcept_no in _xml_cache:
        cached_time, cached_content = _xml_cache[rcept_no]
        if now - cached_time < _XML_CACHE_TTL:
            return cached_content

    _rate_limit_api()
    url = f"{DART_API_BASE}/document.xml"
    resp = requests.get(url, params={"crtfc_key": API_KEY, "rcept_no": rcept_no}, timeout=60)
    resp.raise_for_status()
    z = zipfile.ZipFile(io.BytesIO(resp.content))
    xml_name = z.namelist()[0]
    xml_content = z.read(xml_name).decode("utf-8", errors="replace")

    if len(_xml_cache) >= _XML_CACHE_MAX:
        oldest_key = min(_xml_cache, key=lambda k: _xml_cache[k][0])
        del _xml_cache[oldest_key]
    _xml_cache[rcept_no] = (now, xml_content)
    return xml_content


AUDITOR_PATTERNS = [
    re.compile(r"([\w가-힣]+회계법인)", re.UNICODE),
    re.compile(r"(삼정\s*KPMG)", re.UNICODE),
    re.compile(r"([\w가-힣]+공인회계사감사반[^<\s]*)", re.UNICODE),
    re.compile(r"([\w가-힣]+감사반[^<\s]*)", re.UNICODE),
]


def _extract_auditor_from_xml(xml_content):
    title_pattern = re.compile(r"<TITLE[^>]*>(.*?)</TITLE>", re.IGNORECASE | re.DOTALL)
    titles = list(title_pattern.finditer(xml_content))

    search_keywords = [
        "독립된감사인",
        "외부감사에관한사항",
        "회계감사인의감사의견",
        "회계감사인",
        "감사보고서",
        "감사인",
    ]

    candidate_sections = []
    for keyword in search_keywords:
        for i, match in enumerate(titles):
            raw_title = re.sub(r"<[^>]+>", "", match.group(1)).strip()
            clean = raw_title.replace(" ", "")
            if keyword in clean:
                start = match.start()
                end = titles[i + 1].start() if i + 1 < len(titles) else len(xml_content)
                candidate_sections.append(xml_content[start:end])
        if candidate_sections:
            break

    if not candidate_sections:
        candidate_sections = [xml_content[:min(80000, len(xml_content))]]

    for section in candidate_sections:
        section_text = re.sub(r"<[^>]+>", " ", section)
        # First try matching on original text (handles "안진회계법인" etc.)
        for pattern in AUDITOR_PATTERNS:
            matches = pattern.findall(section_text)
            if matches:
                return matches[0].strip()
        # Try matching spaced auditor names (e.g. "삼 일 회 계 법 인")
        spaced = re.findall(
            r"([가-힣]\s+[가-힣]\s+회\s*계\s*법\s*인)", section_text
        )
        if spaced:
            return re.sub(r"\s+", "", spaced[0])

    return ""


def _extract_auditor(rcept_no):
    if rcept_no in _auditor_cache:
        return _auditor_cache[rcept_no]
    try:
        xml_content = _download_document_xml(rcept_no)
        auditor = _extract_auditor_from_xml(xml_content)
        _auditor_cache[rcept_no] = auditor
        return auditor
    except Exception as exc:
        logger.warning("감사인 추출 실패 (rcept_no=%s): %s", rcept_no, exc)
        _auditor_cache[rcept_no] = ""
        return ""


# ── 문서 섹션 ──

def _extract_sections_from_xml(xml_content):
    sections = []
    title_pattern = re.compile(r"<TITLE[^>]*>(.*?)</TITLE>", re.IGNORECASE | re.DOTALL)
    for match in title_pattern.finditer(xml_content):
        raw_title = match.group(1)
        clean_title = re.sub(r"<[^>]+>", "", raw_title).strip()
        clean_title = re.sub(r"\s+", " ", clean_title).strip("[] ")
        if not clean_title:
            continue
        sections.append({
            "title": clean_title,
            "index": len(sections),
            "start": match.start(),
        })
    return sections


def _extract_section_html(xml_content, sections, section_index):
    if section_index < 0 or section_index >= len(sections):
        return ""
    start = sections[section_index]["start"]
    end = sections[section_index + 1]["start"] if section_index + 1 < len(sections) else len(xml_content)
    return xml_content[start:end]


# ── 감사보고서 본문 추출 ──

def _extract_audit_section_text(xml_content):
    """XML에서 감사보고서/연결감사보고서 섹션의 텍스트만 분리 추출한다."""
    sections = _extract_sections_from_xml(xml_content)

    audit_html_chunks = []
    consolidated_html_chunks = []

    for sec in sections:
        title_clean = sec["title"].replace(" ", "")
        if "연결감사보고서" in title_clean:
            consolidated_html_chunks.append(_extract_section_html(xml_content, sections, sec["index"]))
        elif "감사보고서" in title_clean and "내부회계" not in title_clean and "감사의감사" not in title_clean:
            audit_html_chunks.append(_extract_section_html(xml_content, sections, sec["index"]))

    def _to_text(chunks):
        if not chunks:
            return ""
        joined = " ".join(chunks)
        # 태그 제거 → 공백 정규화
        text = re.sub(r"<[^>]+>", " ", joined)
        text = re.sub(r"\s+", " ", text)
        return text.strip()

    return {
        "audit": _to_text(audit_html_chunks),
        "consolidated": _to_text(consolidated_html_chunks),
    }


def _get_audit_text(rcept_no):
    """감사/연결감사 섹션 텍스트를 캐시 사용해 반환."""
    if rcept_no in _audit_text_cache:
        return _audit_text_cache[rcept_no]
    try:
        xml_content = _download_document_xml(rcept_no)
        result = _extract_audit_section_text(xml_content)
    except Exception as exc:
        logger.warning("감사보고서 본문 추출 실패 (rcept_no=%s): %s", rcept_no, exc)
        result = {"audit": "", "consolidated": ""}

    if len(_audit_text_cache) >= _AUDIT_TEXT_CACHE_MAX:
        _audit_text_cache.pop(next(iter(_audit_text_cache)))
    _audit_text_cache[rcept_no] = result
    return result


# ── Flask 라우트 ──

@app.errorhandler(400)
def bad_request(e):
    return jsonify({"error": "잘못된 요청입니다.", "detail": str(e)}), 400

@app.errorhandler(404)
def not_found(e):
    return jsonify({"error": "요청한 리소스를 찾을 수 없습니다."}), 404

@app.errorhandler(500)
def internal_error(e):
    logger.exception("서버 내부 오류 발생")
    return jsonify({"error": "서버 내부 오류가 발생했습니다.", "detail": str(e)}), 500


@app.route("/")
def index():
    index_path = os.path.join(BASE_DIR, "index_pdf.html")
    if os.path.exists(index_path):
        return send_from_directory(BASE_DIR, "index_pdf.html")
    return jsonify({"message": "DART 감사보고서 본문검색 API"})


@app.route("/api/health")
def health():
    return jsonify({"status": "ok", "timestamp": datetime.now().isoformat()})


_SUB_DOC_SUFFIXES = ["감사보고서", "연결감사보고서", "영업보고서", "정관",
                     "분기검토보고서", "분기연결검토보고서", "반기검토보고서",
                     "반기연결검토보고서", "내부회계관리제도운영보고서",
                     "감사의감사보고서"]


def _deduplicate_results(results):
    """같은 rcept_no의 하위문서를 제거하고, 기재정정이 있으면 원본을 제거한다."""
    # Step 1: 같은 rcept_no 하위문서 제거
    seen_rcept = {}
    deduped = []
    for r in results:
        rcept = r.get("rcept_no", "")
        name = r.get("report_name", "")
        is_sub = any(name.endswith(s) for s in _SUB_DOC_SUFFIXES)
        if rcept not in seen_rcept:
            seen_rcept[rcept] = len(deduped)
            deduped.append(r)
        elif not is_sub:
            existing = deduped[seen_rcept[rcept]]
            existing_is_sub = any(existing.get("report_name", "").endswith(s) for s in _SUB_DOC_SUFFIXES)
            if existing_is_sub:
                deduped[seen_rcept[rcept]] = r

    # Step 2: 기재정정 있으면 같은 기간 원본 제거 (최종보고서만 남기기)
    import re as _re
    period_map = {}
    for r in deduped:
        name = r.get("report_name", "")
        m = _re.search(r"(사업보고서|반기보고서|분기보고서)\s*\((\d{4}\.\d{2})\)", name)
        if m:
            key = f"{r.get('company_name', '')}_{m.group(1)}_{m.group(2)}"
            is_correction = "[기재정정]" in name or "[첨부정정]" in name
            if key not in period_map:
                period_map[key] = {"correction": None, "original": None}
            if is_correction:
                period_map[key]["correction"] = r.get("rcept_no", "")
            else:
                period_map[key]["original"] = r.get("rcept_no", "")

    remove_rcepts = set()
    for info in period_map.values():
        if info["correction"] and info["original"]:
            remove_rcepts.add(info["original"])

    return [r for r in deduped if r.get("rcept_no", "") not in remove_rcepts]


def _list_filings(corp_name, start_date, end_date, page=1, max_results=15):
    """키워드 없이 회사명으로 OpenDART 공시목록을 조회한다."""
    corp_cik = _lookup_corp_cik(corp_name) if corp_name else ""
    if not corp_cik:
        return {"results": [], "page_info": {"total": 0, "has_next": False, "has_prev": False, "total_pages": 0}}

    _rate_limit_api()
    try:
        resp = requests.get(
            f"{DART_API_BASE}/list.json",
            params={
                "crtfc_key": API_KEY,
                "corp_code": corp_cik,
                "bgn_de": start_date.replace("-", ""),
                "end_de": end_date.replace("-", ""),
                "pblntf_ty": "A",
                "page_no": str(page),
                "page_count": str(max_results),
            },
            timeout=30,
        )
        resp.raise_for_status()
        data = resp.json()
    except Exception as exc:
        logger.error("OpenDART 공시목록 조회 실패: %s", exc)
        raise Exception("DART 공시목록 조회에 실패했습니다.")

    if data.get("status") != "000":
        if data.get("status") == "013":
            return {"results": [], "page_info": {"total": 0, "has_next": False, "has_prev": False, "total_pages": 0}}
        raise Exception(f"DART 오류: {data.get('message', '알 수 없는 오류')}")

    market_map = {"Y": "유가증권시장", "K": "코스닥", "N": "코넥스", "E": "기타"}
    results = []
    for item in data.get("list", []):
        rcept_no = item.get("rcept_no", "")
        date_raw = item.get("rcept_dt", "")
        date_text = f"{date_raw[:4]}.{date_raw[4:6]}.{date_raw[6:8]}" if len(date_raw) == 8 else date_raw
        results.append({
            "rcept_no": rcept_no,
            "dcm_no": "",
            "corp_code": item.get("corp_code", ""),
            "company_name": item.get("corp_name", ""),
            "market": market_map.get(item.get("corp_cls", ""), ""),
            "report_name": item.get("report_nm", ""),
            "filing_type": "정기공시",
            "submitter": item.get("flr_nm", ""),
            "auditor": _auditor_cache.get(rcept_no, ""),
            "date": date_text,
            "snippet": "",
            "dart_url": f"https://dart.fss.or.kr/dsaf001/main.do?rcpNo={rcept_no}" if rcept_no else "",
        })

    total_count = data.get("total_count", 0)
    total_pages = data.get("total_page", 1)
    page_info = {
        "total": total_count,
        "has_next": page < total_pages,
        "has_prev": page > 1,
        "total_pages": total_pages,
    }

    return {"results": results, "page_info": page_info}


@app.route("/api/search")
def search():
    keyword = flask_request.args.get("keyword", "").strip()
    corp_name = flask_request.args.get("corp_name", "").strip()

    if not keyword and not corp_name:
        return jsonify({"error": "검색 키워드 또는 회사명을 입력해주세요."}), 400
    if keyword and len(keyword) < 2:
        return jsonify({"error": "검색어는 2자 이상 입력해주세요."}), 400

    start_date = flask_request.args.get("start_date", "")
    end_date = flask_request.args.get("end_date", "")
    page = flask_request.args.get("page", "1")
    max_results = flask_request.args.get("max_results", "15")
    dsp_type = flask_request.args.get("dsp_type", "A,F")
    scan_mode = flask_request.args.get("scan", "").lower() in ("1", "true", "yes")

    if not start_date:
        start_date = (datetime.now() - timedelta(days=1095)).strftime("%Y%m%d")
    else:
        start_date = start_date.replace("-", "")
    if not end_date:
        end_date = datetime.now().strftime("%Y%m%d")
    else:
        end_date = end_date.replace("-", "")

    try:
        page_num = max(1, int(page))
        max_num = min(100, max(10, int(max_results)))
    except ValueError:
        page_num = 1
        max_num = 15

    try:
        if keyword:
            dsp_types = [d.strip() for d in dsp_type.split(",") if d.strip()]
            # DART는 페이지당 10건 고정이므로, dedup 후 원하는 건수를 채우기 위해
            # 여러 페이지를 가져와서 합침
            all_results = []
            dart_page = (page_num - 1) * 5 + 1  # 각 우리 페이지당 DART 5페이지씩
            total_info = None
            for i in range(5):
                data = _search_dart(
                    keyword=keyword,
                    start_date=start_date,
                    end_date=end_date,
                    page=dart_page + i,
                    max_results=100,
                    dsp_types=dsp_types,
                    corp_name=corp_name,
                )
                all_results.extend(data["results"])
                if total_info is None:
                    total_info = data["page_info"]
                if not data["page_info"]["has_next"]:
                    break

            deduped = _deduplicate_results(all_results)
            if scan_mode:
                # 스캔 모드: 한 호출에 5 DART 페이지(최대 ~500건) 전부 반환
                # 페이지네이션은 우리 page 단위(=DART 5페이지) 그대로 유지
                last_page_info = data["page_info"]
                data = {
                    "results": deduped,
                    "page_info": {
                        "total": last_page_info.get("total", len(deduped)),
                        "has_next": last_page_info.get("has_next", False),
                        "has_prev": page_num > 1,
                        "total_pages": total_info.get("total_pages", 0) if total_info else 0,
                    }
                }
            else:
                has_more = len(deduped) > max_num or data["page_info"]["has_next"]
                data = {
                    "results": deduped[:max_num],
                    "page_info": {
                        "total": len(deduped),
                        "has_next": has_more,
                        "has_prev": page_num > 1,
                        "total_pages": total_info.get("total_pages", 0) if total_info else 0,
                    }
                }
        else:
            data = _list_filings(
                corp_name=corp_name,
                start_date=start_date,
                end_date=end_date,
                page=page_num,
                max_results=max_num,
            )
    except Exception as exc:
        logger.error("DART 검색 실패: %s", exc)
        return jsonify({"error": str(exc)}), 502

    auditors_found = set()
    need_extraction = set()
    for r in data["results"]:
        if r["auditor"]:
            auditors_found.add(r["auditor"])
        elif "정기공시" in r.get("filing_type", "") and r["rcept_no"]:
            need_extraction.add(r["rcept_no"])

    corp_cik_resolved = bool(_corp_cik_cache.get(corp_name)) if corp_name else None

    return jsonify({
        "results": data["results"],
        "total": data["page_info"]["total"],
        "total_pages": data["page_info"].get("total_pages", 0),
        "has_next": data["page_info"]["has_next"],
        "has_prev": data["page_info"]["has_prev"],
        "page": page_num,
        "auditors_in_results": sorted(auditors_found),
        "need_auditor_extraction": sorted(need_extraction),
        "corp_name_resolved": corp_cik_resolved,
    })


@app.route("/api/extract-auditors", methods=["POST"])
def extract_auditors():
    body = flask_request.get_json(silent=True) or {}
    rcept_nos = body.get("rcept_nos", [])

    if not rcept_nos or not isinstance(rcept_nos, list):
        return jsonify({"error": "rcept_nos 배열이 필요합니다."}), 400

    rcept_nos = rcept_nos[:20]
    result = {}
    for rcept_no in rcept_nos:
        if not rcept_no or not str(rcept_no).isdigit():
            continue
        try:
            auditor = _extract_auditor(str(rcept_no))
        except Exception:
            auditor = ""
        result[str(rcept_no)] = auditor

    return jsonify({"auditors": result})


@app.route("/api/extract-auditor/<rcept_no>")
def extract_auditor_single(rcept_no):
    if not rcept_no or not rcept_no.isdigit():
        return jsonify({"error": "유효하지 않은 접수번호입니다."}), 400
    auditor = _extract_auditor(rcept_no)
    return jsonify({"rcept_no": rcept_no, "auditor": auditor})


@app.route("/api/check-audit-content", methods=["POST"])
def check_audit_content():
    """주어진 키워드들이 (감사보고서|연결감사보고서) 섹션 본문에 모두 포함되는지 검증."""
    body = flask_request.get_json(silent=True) or {}
    rcept_nos = body.get("rcept_nos", [])
    keywords = body.get("keywords", [])

    if not rcept_nos or not isinstance(rcept_nos, list):
        return jsonify({"error": "rcept_nos 배열이 필요합니다."}), 400
    if not isinstance(keywords, list):
        keywords = []

    rcept_nos = rcept_nos[:20]
    keywords = [str(k).strip() for k in keywords if str(k).strip()]

    result = {}
    for rcept_no in rcept_nos:
        rcept_str = str(rcept_no)
        if not rcept_str.isdigit():
            continue
        text_dict = _get_audit_text(rcept_str)
        audit_text = text_dict.get("audit", "")
        consol_text = text_dict.get("consolidated", "")

        # 키워드 모두 포함 여부 (대소문자 무시)
        def _all_in(text, kws):
            if not kws:
                return bool(text)
            tl = text.lower()
            return all(kw.lower() in tl for kw in kws)

        in_audit = _all_in(audit_text, keywords) if audit_text else False
        in_consolidated = _all_in(consol_text, keywords) if consol_text else False

        result[rcept_str] = {
            "in_audit": in_audit,
            "in_consolidated": in_consolidated,
            "matched": in_audit or in_consolidated,
            "has_audit_section": bool(audit_text),
            "has_consolidated_section": bool(consol_text),
        }

    return jsonify({"results": result})


@app.route("/api/document/<rcept_no>")
def document_sub_docs(rcept_no):
    if not rcept_no or not rcept_no.isdigit():
        return jsonify({"error": "유효하지 않은 접수번호입니다."}), 400
    try:
        xml_content = _download_document_xml(rcept_no)
        sections = _extract_sections_from_xml(xml_content)
        documents = [{"title": s["title"], "index": s["index"]} for s in sections]
        return jsonify({"documents": documents, "total": len(documents)})
    except Exception as exc:
        logger.error("문서 목록 조회 실패: %s", exc)
        return jsonify({"error": "문서 목록 조회에 실패했습니다.", "detail": str(exc)}), 502


@app.route("/api/document/<rcept_no>/content")
def document_content(rcept_no):
    if not rcept_no or not rcept_no.isdigit():
        return jsonify({"error": "유효하지 않은 접수번호입니다."}), 400

    section_idx = flask_request.args.get("section", "").strip()

    try:
        xml_content = _download_document_xml(rcept_no)
        sections = _extract_sections_from_xml(xml_content)

        target_index = None
        if section_idx and section_idx.isdigit():
            target_index = int(section_idx)
        else:
            for sec in sections:
                title_clean = sec["title"].replace(" ", "")
                if "감사보고서" in title_clean or "주석" in title_clean:
                    target_index = sec["index"]
                    break

        if target_index is None and sections:
            target_index = 0

        if target_index is None:
            return jsonify({"error": "문서를 찾을 수 없습니다."}), 404

        html_content = _extract_section_html(xml_content, sections, target_index)
        html_content = f'<div style="font-family: \'Malgun Gothic\', sans-serif; padding: 20px; line-height: 1.8;">{html_content}</div>'

        return jsonify({
            "rcept_no": rcept_no,
            "section_index": target_index,
            "section_title": sections[target_index]["title"] if target_index < len(sections) else "",
            "html": html_content,
            "dart_url": f"https://dart.fss.or.kr/dsaf001/main.do?rcpNo={rcept_no}",
        })
    except Exception as exc:
        logger.error("문서 내용 조회 실패: %s", exc)
        return jsonify({"error": "문서 본문을 가져오는 데 실패했습니다.", "detail": str(exc)}), 502


if __name__ == "__main__":
    logger.info("DART 감사보고서 본문검색 서버를 시작합니다 (http://localhost:5000)")
    app.run(host="0.0.0.0", port=5000, debug=True, threaded=True)
