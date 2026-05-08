# -*- coding: utf-8 -*-
"""
추가 샘플(샘플3 양식A, 샘플4 양식B)을 Excel/PDF 양쪽 경로로 검증.

기존 simulate_validator.py(엑셀 파서) / simulate_pdf_path.py(PDF 파서)의
함수를 그대로 import해서 4개 시나리오를 일괄 실행.
"""
from pathlib import Path

from simulate_validator import (
    parse_fund_excel, parse_bank_excel, cross_validate, load_xlsx,
)
from simulate_pdf_path import (
    parse_daily_fund_plan, parse_bank_pdf, to_bank_data, extract_pdf_text,
)

OUT = Path(__file__).resolve().parent


def banner(title):
    print()
    print("=" * 70)
    print(f"  {title}")
    print("=" * 70)


def run_excel(fund_xlsx, bank_xlsx, label):
    banner(f"[Excel] {label}")
    fund_raw = load_xlsx(OUT / fund_xlsx)
    fund_data = parse_fund_excel(fund_raw)
    print(f"[Fund] date={fund_data['date']}  transactions={len(fund_data['transactions'])}")
    for t in fund_data["transactions"]:
        print(f"  - {t['account']:30} 수입={t['deposit']:>13,.0f}  지출={t['withdrawal']:>13,.0f}")

    bank_raw = load_xlsx(OUT / bank_xlsx)
    bank_data = parse_bank_excel(bank_raw)
    print(f"[Bank] transactions={len(bank_data['transactions'])}")
    for t in bank_data["transactions"]:
        print(f"  - {t['date']:20} 입금={t['deposit']:>13,.0f}  "
              f"출금={t['withdrawal']:>13,.0f}  {t['description']}")

    cross_validate(fund_data, bank_data)


def run_pdf(fund_pdf, bank_pdf, label):
    banner(f"[PDF] {label}")
    fund_text = extract_pdf_text(OUT / fund_pdf)
    fund_data = parse_daily_fund_plan(fund_text)
    print(f"[Fund] date={fund_data['date']}  transactions={len(fund_data['transactions'])}")
    print(f"[Fund] totals: income={fund_data['totals']['totalIncome']:,}  "
          f"expense={fund_data['totals']['totalExpense']:,}  "
          f"funding={fund_data['totals']['totalFunding']:,}")
    for t in fund_data["transactions"]:
        print(f"  - {t['account']:30} 수입={t['deposit']:>13,}  지출={t['withdrawal']:>13,}")

    bank_text = extract_pdf_text(OUT / bank_pdf)
    bank_data = to_bank_data(parse_bank_pdf(bank_text))
    print(f"[Bank] transactions={len(bank_data['transactions'])}")
    for t in bank_data["transactions"]:
        print(f"  - {t['date']:19} 입금={t['deposit']:>13,}  "
              f"출금={t['withdrawal']:>13,}  {t['description']}")

    cross_validate(fund_data, bank_data)


if __name__ == "__main__":
    # 양식 A — Excel
    run_excel("자금일보_샘플3.xlsx", "은행거래내역_샘플3.xlsx",
              "양식 A · 자금일계표 Ledger 시계열형 (샘플3)")
    # 양식 A — PDF
    run_pdf("자금일보_샘플3.pdf", "은행거래내역_샘플3.pdf",
            "양식 A · 자금일계표 Ledger 시계열형 (샘플3)")
    # 양식 B — Excel
    run_excel("자금일보_샘플4.xlsx", "은행거래내역_샘플4.xlsx",
              "양식 B · 분류형 (대분류/결제수단/거래처별, 샘플4)")
    # 양식 B — PDF
    run_pdf("자금일보_샘플4.pdf", "은행거래내역_샘플4.pdf",
            "양식 B · 분류형 (대분류/결제수단/거래처별, 샘플4)")
    # 샘플5 — 다중계좌 (LS마린솔루션 사례2 참고)
    run_excel("자금일보_샘플5.xlsx", "은행거래내역_샘플5.xlsx",
              "샘플5 · 다중계좌 (4개 은행/증권사 통합)")
    run_pdf("자금일보_샘플5.pdf", "은행거래내역_샘플5.pdf",
            "샘플5 · 다중계좌 (4개 은행/증권사 통합)")
