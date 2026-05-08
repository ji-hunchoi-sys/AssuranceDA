# -*- coding: utf-8 -*-
"""
샘플6 검증기 경로 시뮬레이션 — 자금일보 1개 + 은행 4개 누적 매칭.

자금일보 검증기는 좌측에 자금일보 1개, 우측에 은행 거래내역 N개를
누적해서 한꺼번에 매칭. 이 스크립트는 그 흐름을 그대로 재현:
  1. 자금일보_샘플6.{xlsx,pdf} 파싱 → fund transactions
  2. 은행거래내역_샘플6_{신한,국민,하나,우리}.{xlsx,pdf} 파싱 → bank transactions
     4개 파일을 합쳐서 통합 거래 리스트 생성
  3. cross_validate(fund, merged_bank) → 합계 검증 + 1:1/N:1 매칭

Excel/PDF 두 경로 모두 100% 매칭 검증.
"""
from pathlib import Path

from simulate_validator import (
    parse_fund_excel, parse_bank_excel, cross_validate, load_xlsx,
)
from simulate_pdf_path import (
    parse_daily_fund_plan, parse_bank_pdf, to_bank_data, extract_pdf_text,
)

OUT = Path(__file__).resolve().parent
BANK_KEYS = ["신한", "국민", "하나", "우리"]


def merge_bank_data(bank_data_list):
    """여러 은행의 transactions를 시간순 병합."""
    merged_txns = []
    for bd in bank_data_list:
        merged_txns.extend(bd["transactions"])
    merged_txns.sort(key=lambda t: t.get("date", ""))
    return {"transactions": merged_txns, "headers": bank_data_list[0]["headers"] if bank_data_list else []}


def banner(title):
    print()
    print("=" * 72)
    print(f"  {title}")
    print("=" * 72)


def run_excel():
    banner("[Excel] 자금일보 1개 + 은행 4개 누적 → cross-validation")

    fund_raw = load_xlsx(OUT / "자금일보_샘플6.xlsx")
    fund_data = parse_fund_excel(fund_raw)
    print(f"[Fund] date={fund_data['date']}  transactions={len(fund_data['transactions'])}")

    bank_data_list = []
    for key in BANK_KEYS:
        path = OUT / f"은행거래내역_샘플6_{key}.xlsx"
        bd = parse_bank_excel(load_xlsx(path))
        print(f"[Bank-{key}] transactions={len(bd['transactions'])}")
        bank_data_list.append(bd)

    merged = merge_bank_data(bank_data_list)
    print(f"[Bank-merged] 총 transactions={len(merged['transactions'])}")
    print()
    cross_validate(fund_data, merged)


def run_pdf():
    banner("[PDF] 자금일보 1개 + 은행 4개 누적 → cross-validation")

    fund_text = extract_pdf_text(OUT / "자금일보_샘플6.pdf")
    fund_data = parse_daily_fund_plan(fund_text)
    print(f"[Fund] date={fund_data['date']}  transactions={len(fund_data['transactions'])}")
    print(f"[Fund] totals: income={fund_data['totals']['totalIncome']:,}  "
          f"expense={fund_data['totals']['totalExpense']:,}  "
          f"funding={fund_data['totals']['totalFunding']:,}")

    bank_data_list = []
    for key in BANK_KEYS:
        path = OUT / f"은행거래내역_샘플6_{key}.pdf"
        text = extract_pdf_text(path)
        bd = to_bank_data(parse_bank_pdf(text))
        print(f"[Bank-{key}] transactions={len(bd['transactions'])}")
        bank_data_list.append(bd)

    merged = merge_bank_data(bank_data_list)
    print(f"[Bank-merged] 총 transactions={len(merged['transactions'])}")
    print()
    cross_validate(fund_data, merged)


if __name__ == "__main__":
    run_excel()
    run_pdf()
