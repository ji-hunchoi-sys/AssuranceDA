# -*- coding: utf-8 -*-
"""
자금일보 자동 생성 페이지의 PDF 처리 시뮬레이션.

index2_pdf.html의 자동 생성 페이지(handleGenBankFile)는 두 가지 방식으로
PDF 텍스트를 추출:
  · BEFORE (버그): tc.items.map(it => it.str).join(' ') — 모든 텍스트를
    한 페이지당 한 줄로 합쳐버려 표 구조가 뭉개짐 → parseBankPDF가 라인
    패턴 매칭 실패 → 0건 인식 (사용자 캡쳐 재현)
  · AFTER (수정): extractPdfTextLineAware — Y좌표 기반 줄 단위 그룹핑

이 스크립트로 BEFORE/AFTER 두 경로를 시뮬레이션해서:
  1. 버그 재현 (BEFORE에서 0건 또는 매우 적은 건 인식)
  2. 수정 효과 (AFTER에서 정상적으로 모든 거래 인식)
"""
from pathlib import Path

import pdfplumber
from simulate_pdf_path import parse_bank_pdf

OUT = Path(__file__).resolve().parent


def extract_text_buggy(pdf_path):
    """기존 자동 생성 페이지의 단순 join 방식 (수정 전 버그 재현).
    pdf.js의 tc.items.map(...).join(' ')와 동등한 동작 — Y좌표 무시,
    왼쪽-오른쪽-위-아래 순서로 모두 한 줄로 합침.
    """
    with pdfplumber.open(pdf_path) as pdf:
        all_words = []
        for page in pdf.pages:
            for word in page.extract_words():
                all_words.append(word["text"])
        return " ".join(all_words)


def extract_text_fixed(pdf_path):
    """수정된 줄 단위 추출 (extractPdfTextLineAware 시뮬레이션).
    pdfplumber.extract_text()는 Y좌표 기반으로 자동 줄 그룹핑.
    """
    with pdfplumber.open(pdf_path) as pdf:
        return "\n".join(page.extract_text() for page in pdf.pages)


def simulate(bank_pdf, extract_fn):
    text = extract_fn(bank_pdf)
    parsed = parse_bank_pdf(text)
    return parsed["rows"]


def run_check(pdf_filename):
    path = OUT / pdf_filename
    if not path.exists():
        print(f"  ⚠ 파일 없음: {pdf_filename} — 건너뜀")
        return None
    before_rows = simulate(path, extract_text_buggy)
    after_rows = simulate(path, extract_text_fixed)
    return {
        "file": pdf_filename,
        "before": len(before_rows),
        "after": len(after_rows),
        "after_sample": after_rows[:3],
    }


if __name__ == "__main__":
    print("=" * 70)
    print("  자금일보 자동 생성 페이지 — PDF 텍스트 추출 BEFORE/AFTER 비교")
    print("=" * 70)
    print()
    print("BEFORE: 단순 join (페이지당 한 줄로 합침) — 사용자 캡쳐 재현")
    print("AFTER : 줄 단위 그룹핑 (extractPdfTextLineAware) — 수정안")
    print()
    print(f"{'파일':<35} {'BEFORE':>8} {'AFTER':>8}  결과")
    print("-" * 70)

    results = []
    for pdf in [
        "은행거래내역_샘플3.pdf",
        "은행거래내역_샘플4.pdf",
        "은행거래내역_샘플5.pdf",
        "은행거래내역_샘플2.pdf",  # 기존 샘플도 비교
    ]:
        r = run_check(pdf)
        if r is None:
            continue
        results.append(r)
        ok = r["after"] > 0 and r["after"] >= r["before"]
        mark = "✓ 정상" if ok else "✗ 실패"
        print(f"  {r['file']:<33} {r['before']:>6}건 {r['after']:>6}건  {mark}")

    print()
    # 상세
    for r in results:
        print(f"\n[{r['file']}] AFTER 거래 샘플 (최대 3건):")
        for row in r["after_sample"]:
            print(f"  · {row}")

    print()
    print("=" * 70)
    all_ok = all(r["after"] > 0 for r in results)
    if all_ok:
        print("  ✓ 모든 샘플에서 수정 후 거래 인식 정상")
    else:
        print("  ✗ 일부 샘플에서 거래 인식 실패 — 추가 수정 필요")
    print("=" * 70)
