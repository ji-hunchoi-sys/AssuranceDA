# -*- coding: utf-8 -*-
"""
자금일보 검증기 샘플 통합 생성기 (양식 선택 지원).

사용법:
  python generate_sample.py                     # 인터랙티브 메뉴 (양식 선택)
  python generate_sample.py --format A          # 양식 A만 (샘플3 - 시계열 Ledger)
  python generate_sample.py --format B          # 양식 B만 (샘플4 - 분류형)
  python generate_sample.py --format C          # 양식 C만 (샘플5 - 다중계좌)
  python generate_sample.py --format all        # 모든 양식
  python generate_sample.py --format A,C        # 여러 개 선택
  python generate_sample.py --list              # 양식 목록만 출력

양식 종류:
  A) 시계열 Ledger (자금일계표)
     · 한 거래 한 행 시계열 정렬, 시간/구분/계정과목/거래처/적요/수입/지출
     · 출력 → 자금일보_샘플3.{xlsx,pdf} + 은행거래내역_샘플3.{xlsx,pdf}

  B) 분류형 (대분류 + 결제수단)
     · 대분류·계정과목·거래처별 그룹핑, 결제수단 컬럼 추가
     · N:1 합산 매칭 케이스 포함
     · 출력 → 자금일보_샘플4.{xlsx,pdf} + 은행거래내역_샘플4.{xlsx,pdf}

  C) 다중계좌 (LS마린솔루션 자금일일결산 사례 참고)
     · 4개 금융기관(은행+증권사) 통합, 운영자금 메타 + 입출금내역 표
     · 출력 → 자금일보_샘플5.{xlsx,pdf} + 은행거래내역_샘플5.{xlsx,pdf}
"""
import argparse
import sys
from pathlib import Path

OUT_DIR = Path(__file__).resolve().parent

# 양식 정의 (코드, 라벨, 모듈, 함수 목록)
FORMATS = {
    "A": {
        "label": "시계열 Ledger (자금일계표)",
        "files": [
            "자금일보_샘플3.xlsx", "자금일보_샘플3.pdf",
            "은행거래내역_샘플3.xlsx", "은행거래내역_샘플3.pdf",
        ],
        "module": "generate_extra_samples",
        "functions": [
            "build_form_a_fund_excel",
            "build_form_a_bank_excel",
            "build_form_a_fund_pdf",
            "build_form_a_bank_pdf",
        ],
    },
    "B": {
        "label": "분류형 (대분류·결제수단·거래처별)",
        "files": [
            "자금일보_샘플4.xlsx", "자금일보_샘플4.pdf",
            "은행거래내역_샘플4.xlsx", "은행거래내역_샘플4.pdf",
        ],
        "module": "generate_extra_samples",
        "functions": [
            "build_form_b_fund_excel",
            "build_form_b_bank_excel",
            "build_form_b_fund_pdf",
            "build_form_b_bank_pdf",
        ],
    },
    "C": {
        "label": "다중계좌 (4개 은행/증권사 통합)",
        "files": [
            "자금일보_샘플5.xlsx", "자금일보_샘플5.pdf",
            "은행거래내역_샘플5.xlsx", "은행거래내역_샘플5.pdf",
        ],
        "module": "generate_multi_account_sample",
        "functions": [
            "build_fund_excel",
            "build_bank_excel",
            "build_fund_pdf",
            "build_bank_pdf",
        ],
    },
}


def list_formats():
    print()
    print("=" * 70)
    print("  자금일보 양식 목록")
    print("=" * 70)
    for code, info in FORMATS.items():
        print(f"  [{code}] {info['label']}")
        for f in info["files"]:
            print(f"        · {f}")
        print()


def run_format(code):
    info = FORMATS[code]
    print(f"\n[양식 {code}] {info['label']}")
    print("-" * 70)
    mod = __import__(info["module"])
    for fn in info["functions"]:
        getattr(mod, fn)()


def interactive_menu():
    print()
    print("=" * 70)
    print("  자금일보 검증기 — 샘플 양식 생성기")
    print("=" * 70)
    print()
    for code, info in FORMATS.items():
        print(f"  [{code}] {info['label']}")
    print(f"  [all] 모든 양식 일괄 생성")
    print(f"  [q]   종료")
    print()
    while True:
        choice = input("생성할 양식을 선택하세요 (예: A / B / C / all / A,C): ").strip()
        if not choice:
            continue
        if choice.lower() in ("q", "quit", "exit"):
            print("취소됨.")
            return []
        if choice.lower() == "all":
            return list(FORMATS.keys())
        codes = [c.strip().upper() for c in choice.split(",") if c.strip()]
        unknown = [c for c in codes if c not in FORMATS]
        if unknown:
            print(f"  ⚠ 알 수 없는 양식: {', '.join(unknown)} — 다시 입력하세요.")
            continue
        return codes


def parse_format_arg(arg):
    if arg.lower() == "all":
        return list(FORMATS.keys())
    codes = [c.strip().upper() for c in arg.split(",") if c.strip()]
    unknown = [c for c in codes if c not in FORMATS]
    if unknown:
        print(f"오류: 알 수 없는 양식 '{', '.join(unknown)}'", file=sys.stderr)
        print(f"사용 가능: {', '.join(FORMATS.keys())} 또는 all", file=sys.stderr)
        sys.exit(2)
    return codes


def main():
    parser = argparse.ArgumentParser(
        description="자금일보 검증기 샘플 양식 생성기 (양식 A/B/C 또는 일괄)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument(
        "--format", "-f",
        help="생성할 양식 (A | B | C | all | 콤마 결합 예: A,C)",
    )
    parser.add_argument(
        "--list", "-l", action="store_true",
        help="사용 가능한 양식 목록 출력 후 종료",
    )
    args = parser.parse_args()

    if args.list:
        list_formats()
        return

    if args.format:
        codes = parse_format_arg(args.format)
    else:
        codes = interactive_menu()
        if not codes:
            return

    for code in codes:
        run_format(code)

    print()
    print("=" * 70)
    print(f"생성 완료 ({len(codes)}개 양식, {sum(len(FORMATS[c]['files']) for c in codes)}개 파일)")
    print(f"파일 위치: {OUT_DIR}")
    print("=" * 70)


if __name__ == "__main__":
    main()
