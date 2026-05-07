# -*- coding: utf-8 -*-
"""
PDF 입력 경로(parseDailyFundPlan + parseBankPDF) 시뮬레이션.
pdfplumber로 텍스트 추출 → JS 파서 로직 재현.
"""
import re
from pathlib import Path

import pdfplumber
from simulate_validator import cross_validate

OUT = Path(__file__).resolve().parent


# ─────────────────────────────────────────────────────────────────
# parseDailyFundPlan 재현 (index2_pdf.html에서 직접 옮김)
# ─────────────────────────────────────────────────────────────────
def extract_valid_amounts(line):
    out = []
    for m in re.finditer(r"[\d,]+", line):
        raw = m.group(0)
        if raw == "," or raw.replace(",", "") == "":
            continue
        digits = raw.replace(",", "")
        if "," in raw:
            if not re.match(r"^\d{1,3}(,\d{3})+$", raw):
                continue
        else:
            if not (4 <= len(digits) <= 12):
                continue
        try:
            n = int(digits)
        except ValueError:
            continue
        if n < 1000:
            continue
        out.append(n)
    return out


def classify_fund_item(line):
    t = re.sub(r"\s+", "", line)
    if re.search(r"예금입출", t): return "funding"
    if re.search(r"정부지원자금", t): return "funding"
    if re.search(r"조달자금|조달후", t): return "funding"
    if re.search(r"차입금", t) and not re.search(r"차입금상환|상환", t):
        return "funding"
    if re.search(r"차입금상환|외상대결제|설비투자|연구개발", t): return "expense"
    if re.search(r"급/?상여|회직급여|급상여", t): return "expense"
    if re.search(r"일반경비|전도금", t): return "expense"
    if re.search(r"지급|결제|상환|납부|이체출금|출금|지출|급여|경비|투자|개발비", t): return "expense"
    if re.search(r"외상대수금|매출수금", t): return "income"
    if re.search(r"수금|회수|이체입금|매출입금|매출대금|입금|수입", t): return "income"
    return None


def parse_daily_fund_plan(text):
    result = {
        "transactions": [], "date": None, "totals": {},
        "openingBalance": 0, "closingBalance": 0,
    }
    # 날짜
    m = re.search(r"(\d{4})[.\-/년]\s*(\d{1,2})[.\-/월]\s*(\d{1,2})", text)
    if m:
        result["date"] = f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"

    lines = [l.strip() for l in text.split("\n") if l.strip()]
    section = "none"
    income_items, expense_items, funding_items = [], [], []
    opening = closing = 0

    for line in lines:
        clean = re.sub(r"\s+", "", line)

        # 기초/마감 잔액 (같은 줄에 둘 다 있을 수 있음)
        has_opening = bool(re.search(r"전일이월|전일\s*잔액|기초\s*시재", line))
        has_closing = bool(re.search(r"조달\s*후|마감\s*시재|당일\s*잔액|기말\s*시재|마감\s*잔액", line))
        if has_opening or has_closing:
            cands = extract_valid_amounts(line)
            if cands:
                if has_opening and has_closing and len(cands) >= 2:
                    opening = cands[0]
                    closing = cands[-1]
                elif has_closing:
                    closing = max(cands)
                else:
                    opening = max(cands)
            continue

        if re.search(r"과\s*부\s*족", line):
            continue
        # 메타/주석 줄 스킵
        if re.match(r"^\s*[※*]", line):
            continue
        if re.search(r"회\s*사\s*명|예\s*금\s*주|계\s*좌\s*번\s*호|조\s*회\s*기\s*간|계\s*획\s*일\s*자|단\s*위\s*[:：]|문\s*서\s*번\s*호|법\s*인\s*명|작\s*성\s*자|결\s*재", line):
            continue

        if re.match(r"^[\|\s]*수\s*입[\|\s]*$", clean) or re.search(r"수입\s*항목|입금\s*내역", clean):
            section = "income"; continue
        if re.match(r"^[\|\s]*지\s*출[\|\s]*$", clean) or re.search(r"지출\s*항목|출금\s*내역", clean):
            section = "expense"; continue
        if re.match(r"^[\|\s]*조\s*달[\|\s]*$", clean) or re.search(r"조달\s*항목|조달\s*내역", clean):
            section = "funding"; continue

        if re.search(r"소\s*계|^\s*합\s*계\s*$|^\s*총\s*계\s*$", line):
            continue

        if re.search(r"구\s*분|과\s*목|^\s*금\s*액\s*$|세\s*부\s*내\s*역|^\s*항\s*목\s*$", line) and not re.search(r"[\d,]{4,}", line):
            continue

        if re.match(r"^\d{2,4}[\/.\-]\d{1,2}[\/.\-]\d{1,2}$", re.sub(r"\s", "", line)):
            continue

        valid_nums = extract_valid_amounts(line)
        if not valid_nums:
            continue

        item_name = re.sub(r"[\d,]+", "", line)
        item_name = re.sub(r"[|│┃()（）]", " ", item_name).strip()
        item_name = re.sub(r"^\s*[○●■□▪▸►·\-]\s*", "", item_name)
        item_name = re.sub(r"\s+", " ", item_name).strip()

        korean_chars = len(re.findall(r"[가-힣]", item_name))
        if korean_chars < 2 or not item_name or len(item_name) > 60:
            continue

        amount = max(valid_nums)
        if amount == 0:
            continue

        item_class = classify_fund_item(line)
        final_section = item_class or (section if section != "none" else "expense")
        if item_class:
            section = item_class

        item = {"name": item_name, "amount": amount, "description": line[:80]}
        if final_section == "income":
            income_items.append(item)
        elif final_section == "funding":
            funding_items.append(item)
        else:
            expense_items.append(item)

    for it in income_items:
        result["transactions"].append({
            "account": it["name"], "deposit": it["amount"], "withdrawal": 0,
            "balance": 0, "description": it["description"],
        })
    for it in expense_items:
        result["transactions"].append({
            "account": it["name"], "deposit": 0, "withdrawal": it["amount"],
            "balance": 0, "description": it["description"],
        })
    for it in funding_items:
        result["transactions"].append({
            "account": it["name"] + " (조달)", "deposit": it["amount"], "withdrawal": 0,
            "balance": 0, "description": it["description"],
        })

    result["openingBalance"] = opening
    result["closingBalance"] = closing
    result["totals"] = {
        "openingBalance": opening, "closingBalance": closing,
        "totalIncome": sum(i["amount"] for i in income_items),
        "totalExpense": sum(i["amount"] for i in expense_items),
        "totalFunding": sum(i["amount"] for i in funding_items),
    }
    return result


# ─────────────────────────────────────────────────────────────────
# parseBankPDF 재현 (날짜+시간 추출, 잔액 체인 기반)
# ─────────────────────────────────────────────────────────────────
def parse_bank_pdf(text):
    headers = ["거래일자", "입금액", "출금액", "잔액", "적요"]
    lines = text.split("\n")
    date_pat = re.compile(r"(\d{4}[.\-/]\d{1,2}[.\-/]\d{1,2})(\s*\d{1,2}[:.\s]?\d{2}[:.\s]?\d{2})?")
    stuck_ts = re.compile(r"^(\d{6})(?=\D|$)")
    summary_pat = re.compile(r"합\s*계|총\s*계|소\s*계|이월|이전잔액|page|페이지|조회기간|계좌번호|예금주|거래건수", re.IGNORECASE)
    header_pat = re.compile(r"번호\s*(거래|일자|날짜)|거래일\s*시?\s*(입금|출금|적요)", re.IGNORECASE)

    raw_entries = []
    for line in lines:
        trimmed = line.strip()
        if not trimmed:
            continue
        if summary_pat.search(trimmed) or header_pat.search(trimmed):
            continue
        m = date_pat.search(trimmed)
        if not m:
            continue
        date_str = m.group(1).replace(".", "-").replace("/", "-")
        if re.match(r"^\d{2}-", date_str):
            date_str = "20" + date_str
        yc = re.match(r"^(\d{4})-", date_str)
        if yc and not (1990 <= int(yc.group(1)) <= 2030):
            continue
        mc = re.search(r"-(\d{1,2})-(\d{1,2})$", date_str)
        if mc and not (1 <= int(mc.group(1)) <= 12 and 1 <= int(mc.group(2)) <= 31):
            continue

        full_match_len = len(m.group(1)) + (len(m.group(2)) if m.group(2) else 0)
        idx = trimmed.find(m.group(1))
        after = trimmed[idx + full_match_len:]

        ts_match = stuck_ts.match(after.lstrip())
        if ts_match:
            ts = ts_match.group(1)
            hh, mm_, ss = int(ts[:2]), int(ts[2:4]), int(ts[4:6])
            if 0 <= hh <= 23 and 0 <= mm_ <= 59 and 0 <= ss <= 59:
                pos = after.find(ts)
                after = after[pos + len(ts):]

        amounts = []
        for n in re.finditer(r"\d[\d,.]*\d|\d", after):
            raw = n.group(0)
            digits = raw.replace(",", "").replace(".", "")
            has_sep = "," in raw or "." in raw
            if len(digits) < 4:
                continue
            if has_sep:
                if len(digits) > 12:
                    continue
            else:
                if len(digits) > 9:
                    continue
                if digits.startswith("0"):
                    continue
            try:
                amounts.append(int(digits))
            except ValueError:
                continue

        desc_parts = re.sub(r"[\d,.]+", "|", after).split("|")
        desc_parts = [s.strip() for s in desc_parts if s.strip() and len(s.strip()) > 1 and not re.match(r"^\d+$", s.strip())]
        description = " ".join(desc_parts)[:60]

        if amounts:
            raw_entries.append({"dateStr": date_str, "amounts": amounts, "description": description, "line": trimmed})

    if not raw_entries:
        return {"headers": headers, "rows": []}

    # 잔액 체인 점수 계산
    def score_pos(get_balance):
        s = 0
        for i in range(1, len(raw_entries)):
            pb = get_balance(raw_entries[i - 1])
            cb = get_balance(raw_entries[i])
            if pb is None or cb is None:
                continue
            diff = abs(pb - cb)
            if diff > 0:
                others = [a for a in raw_entries[i]["amounts"] if a != cb]
                if any(abs(a - diff) <= 1 for a in others):
                    s += 2
                elif diff < pb:
                    s += 0.5
        return s

    score_last = score_pos(lambda e: e["amounts"][-1] if e["amounts"] else None)
    score_max = score_pos(lambda e: max(e["amounts"]) if e["amounts"] else None)
    score_2nd = score_pos(lambda e: e["amounts"][1] if len(e["amounts"]) >= 2 else (e["amounts"][0] if e["amounts"] else None))

    best = max(score_last, score_max, score_2nd)
    if best <= 0:
        get_balance = lambda e: max(e["amounts"]) if e["amounts"] else None
    elif score_last >= score_max and score_last >= score_2nd:
        get_balance = lambda e: e["amounts"][-1]
    elif score_max >= score_2nd:
        get_balance = lambda e: max(e["amounts"])
    else:
        get_balance = lambda e: e["amounts"][1] if len(e["amounts"]) >= 2 else e["amounts"][0]

    use_chain = best >= max(1, (len(raw_entries) - 1) * 0.3)
    rows = []
    MAX_PLAUS = 1e12

    if use_chain and len(raw_entries) >= 2:
        for i, entry in enumerate(raw_entries):
            balance = get_balance(entry)
            if balance is None:
                continue
            if i == 0:
                if i + 1 < len(raw_entries):
                    nb = get_balance(raw_entries[i + 1])
                    if nb is not None:
                        diff = balance - nb
                        if abs(diff) > MAX_PLAUS:
                            continue
                        others = [a for a in entry["amounts"] if a != balance]
                        closest = next((a for a in others if abs(a - abs(diff)) <= 1), None)
                        first_is_deposit = diff > 0
                        if closest:
                            if first_is_deposit:
                                rows.append([entry["dateStr"], closest, 0, balance, entry["description"]])
                            else:
                                rows.append([entry["dateStr"], 0, closest, balance, entry["description"]])
                        elif others:
                            line_hint = entry.get("line", "")
                            txt_dep = bool(re.search(r"입금|수입|수금|회수", line_hint))
                            txt_wd = bool(re.search(r"출금|지출|이체|결제|지급|상환|납부", line_hint))
                            is_dep = True if txt_dep else (False if txt_wd else first_is_deposit)
                            if is_dep:
                                rows.append([entry["dateStr"], others[0], 0, balance, entry["description"]])
                            else:
                                rows.append([entry["dateStr"], 0, others[0], balance, entry["description"]])
                continue
            pb = get_balance(raw_entries[i - 1])
            if pb is None:
                continue
            diff = pb - balance
            abs_diff = abs(diff)
            if diff == 0 or abs_diff > MAX_PLAUS:
                continue
            others = [a for a in entry["amounts"] if a != balance]
            matched = any(abs(a - abs_diff) <= 1 for a in others) or any(
                abs(a - abs_diff) / max(abs_diff, 1) < 0.01 for a in others
            )
            tx = abs_diff
            if not matched:
                if abs_diff >= 1000 and abs_diff < MAX_PLAUS:
                    tx = abs_diff
                elif len(others) == 1:
                    tx = others[0]
                else:
                    continue
            is_wd = diff > 0
            rows.append([entry["dateStr"], 0 if is_wd else tx, tx if is_wd else 0, balance, entry["description"]])
    return {"headers": headers, "rows": rows}


def to_bank_data(parsed):
    rows = parsed["rows"]
    return {"transactions": [{
        "date": r[0], "deposit": r[1], "withdrawal": r[2], "balance": r[3], "description": r[4],
    } for r in rows], "headers": parsed["headers"]}


def extract_pdf_text(path):
    with pdfplumber.open(path) as pdf:
        return "\n".join(page.extract_text() for page in pdf.pages)


if __name__ == "__main__":
    print("=" * 60)
    print("PDF 경로 시뮬레이션 - 자금일보_샘플2.pdf + 은행거래내역_샘플2.pdf")
    print("=" * 60)

    fund_text = extract_pdf_text(OUT / "자금일보_샘플2.pdf")
    fund_data = parse_daily_fund_plan(fund_text)
    print(f"[Fund] date={fund_data['date']}")
    print(f"[Fund] opening={fund_data['totals']['openingBalance']:,}  closing={fund_data['totals']['closingBalance']:,}")
    print(f"[Fund] totals: income={fund_data['totals']['totalIncome']:,}  "
          f"expense={fund_data['totals']['totalExpense']:,}  "
          f"funding={fund_data['totals']['totalFunding']:,}")
    print(f"[Fund] {len(fund_data['transactions'])}건의 거래:")
    for t in fund_data["transactions"]:
        print(f"  - {t['account']:30}  수입={t['deposit']:>12,}  지출={t['withdrawal']:>12,}")

    print()
    bank_text = extract_pdf_text(OUT / "은행거래내역_샘플2.pdf")
    bank_data = to_bank_data(parse_bank_pdf(bank_text))
    print(f"[Bank] {len(bank_data['transactions'])}건의 거래:")
    for t in bank_data["transactions"]:
        print(f"  - {t['date']:19} 입금={t['deposit']:>12,}  출금={t['withdrawal']:>12,}  {t['description']}")

    cross_validate(fund_data, bank_data)
