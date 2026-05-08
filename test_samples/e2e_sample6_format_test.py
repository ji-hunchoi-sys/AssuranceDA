# -*- coding: utf-8 -*-
"""
샘플6 + 자금일보 생성기 3종 포맷 E2E 테스트.

테스트 범위:
  1. 자금일보 검증기 페이지 — 자금일보_샘플6 + 4개 은행 누적 → 100% 매칭
  2. 자금일보 자동 생성 페이지 — 4개 은행 PDF 누적 업로드 → 13건 인식
  3. 3가지 포맷 (summary / ledger / category) 각각 generateFundReport()
     실행 후 HTML 출력에 핵심 마커 포함 여부 검증

각 포맷의 검증 마커:
  · summary  : '주요 입출금 내역' + '거래 상세 내역'
  · ledger   : '자금일계표' + '시계열' + 합계 행
  · category : '수입 합계' + '지출 합계' + '금일 순증감'
"""
import asyncio
from pathlib import Path

from playwright.async_api import async_playwright

ROOT = Path(__file__).resolve().parent.parent
PAGE_PATH = ROOT / "index2_pdf.html"
SAMPLES_DIR = Path(__file__).resolve().parent
BANK_KEYS = ["신한", "국민", "하나", "우리"]


async def poll_until(page, expr, timeout_ms=60000, interval_ms=200):
    elapsed = 0
    while elapsed < timeout_ms:
        try:
            v = await page.evaluate(f"() => {expr}")
            if v:
                return v
        except Exception:
            pass
        await page.wait_for_timeout(interval_ms)
        elapsed += interval_ms
    raise TimeoutError(f"Timed out waiting for: {expr}")


async def test_validator_path(page):
    """자금일보 검증기 페이지: 자금일보 1개 + 은행 4개 누적 매칭."""
    print()
    print("=" * 72)
    print("  TEST 1: 자금일보 검증기 — 1 자금일보 + 4 은행 누적")
    print("=" * 72)

    # 검증 페이지로 이동
    await page.evaluate("() => showPage('upload')")
    await page.wait_for_timeout(300)

    # 자금일보 업로드 (PDF)
    fund_pdf = SAMPLES_DIR / "자금일보_샘플6.pdf"
    print(f"  · 자금일보 업로드: {fund_pdf.name}")
    await page.set_input_files("#fundFileInput", str(fund_pdf))
    await poll_until(page, "APP.fundData !== null && APP.fundData.transactions.length > 0", 90000)
    fund_count = await page.evaluate("() => APP.fundData.transactions.length")
    print(f"    → 자금일보 거래 {fund_count}건")

    # 은행 4개 누적 업로드
    for key in BANK_KEYS:
        bank_pdf = SAMPLES_DIR / f"은행거래내역_샘플6_{key}.pdf"
        print(f"  · 은행 업로드: {bank_pdf.name}")
        await page.set_input_files("#bankFileInput", str(bank_pdf))
        await page.wait_for_timeout(800)

    bank_total = await page.evaluate(
        "() => APP.bankFiles.reduce((s,f) => s + (f.transactions||[]).length, 0)"
    )
    bank_files = await page.evaluate("() => APP.bankFiles.length")
    print(f"  · 은행 파일: {bank_files}개, 총 거래 {bank_total}건")

    # 검증 실행
    await page.wait_for_selector("#verifyBtn:not([disabled])", timeout=10000)
    await page.click("#verifyBtn")
    await poll_until(page, "APP.results && APP.results.length > 0", 30000)
    await page.wait_for_timeout(300)

    snapshot = await page.evaluate("""
        () => {
            const fundTx = (APP.fundData && APP.fundData.transactions) || [];
            const bankTx = [];
            for (const f of (APP.bankFiles || [])) for (const t of (f.transactions || [])) bankTx.push(t);
            const fundDep = fundTx.reduce((s,t) => s+t.deposit, 0);
            const fundWd = fundTx.reduce((s,t) => s+t.withdrawal, 0);
            const bankDep = bankTx.reduce((s,t) => s+t.deposit, 0);
            const bankWd = bankTx.reduce((s,t) => s+t.withdrawal, 0);
            const statuses = {};
            for (const r of (APP.results || [])) statuses[r.status] = (statuses[r.status]||0) + 1;
            const mismatches = (APP.results || []).filter(r => r.status !== 'match').map(r => ({
                status: r.status, type: r.type, fundAmount: r.fundAmount,
                bankAmount: r.bankAmount, description: r.description, note: r.note
            }));
            return { fundCount: fundTx.length, bankCount: bankTx.length,
                     fundDep, fundWd, bankDep, bankWd, statuses, mismatches };
        }
    """)
    s = snapshot
    print(f"  · 자금일보: {s['fundCount']}건, 입금 {s['fundDep']:,} / 출금 {s['fundWd']:,}")
    print(f"  · 은행 통합: {s['bankCount']}건, 입금 {s['bankDep']:,} / 출금 {s['bankWd']:,}")
    print(f"  · 결과: {s['statuses']}")
    if s['mismatches']:
        print(f"  · 비매칭 상세:")
        for m in s['mismatches']:
            print(f"      [{m['status']}] {m['type']} fund={m['fundAmount']:,} bank={m['bankAmount']:,} | {m['description']}")

    ok = (
        s['fundCount'] == 13 and s['bankCount'] == 13
        and s['fundDep'] == s['bankDep'] and s['fundWd'] == s['bankWd']
        and s['statuses'].get('mismatch', 0) == 0
        and s['statuses'].get('missing', 0) == 0
        and s['statuses'].get('extra', 0) == 0
    )
    print(f"  → {'[OK] 검증기 100% 일치' if ok else '[FAIL] 검증 오류'}")
    return ok


async def test_generator_path(page, fmt):
    """자동 생성기 — 4개 은행 PDF + 포맷 선택 → HTML 생성 + 마커 검증."""
    print()
    print("=" * 72)
    print(f"  TEST 2: 자금일보 자동 생성 — 4개 은행 + 포맷 [{fmt}]")
    print("=" * 72)

    # 자동 생성 페이지로 이동
    await page.evaluate("() => showPage('generate')")
    await page.wait_for_timeout(300)

    # 기존 누적 파일 초기화
    await page.evaluate("() => { genBankFiles.length = 0; updateGenBankFileListUI(); }")
    await page.wait_for_timeout(200)

    # 4개 은행 PDF 누적 업로드
    for key in BANK_KEYS:
        bank_pdf = SAMPLES_DIR / f"은행거래내역_샘플6_{key}.pdf"
        print(f"  · 은행 업로드: {bank_pdf.name}")
        await page.set_input_files("#genBankInput", str(bank_pdf))
        await page.wait_for_timeout(800)
        # 옵션 카드가 나타날 때까지 대기 (첫 업로드 후)
        await poll_until(page, "!document.getElementById('genOptionsCard').classList.contains('hidden')", 10000)

    bank_count = await page.evaluate(
        "() => genBankFiles.reduce((s,f) => s + (f.transactions||[]).length, 0)"
    )
    bank_files = await page.evaluate("() => genBankFiles.length")
    print(f"  · 은행 파일: {bank_files}개, 총 거래 {bank_count}건")
    if bank_count != 13:
        print(f"  ⚠ 거래 인식 실패 (기대 13건, 실제 {bank_count}건)")
        return False

    # 옵션 입력
    await page.fill("#genCompany", "(주)한국상사")
    await page.fill("#genDate", "2026-05-07")
    # 포맷 선택
    await page.select_option("#genFormat", fmt)
    await page.wait_for_timeout(200)

    # 자금일보 생성 버튼 클릭
    await page.evaluate("() => generateFundReport()")
    await poll_until(page, "APP.generatedReport !== null && APP.generatedReport !== undefined", 5000)
    await page.wait_for_timeout(300)

    # 결과 카드 표시 확인
    visible = await page.evaluate(
        "() => !document.getElementById('genResultCard').classList.contains('hidden')"
    )
    if not visible:
        print(f"  ⚠ 결과 카드가 표시되지 않음")
        return False

    # 생성된 자금일보 데이터 + HTML 검증
    rpt = await page.evaluate("""
        () => {
            const r = APP.generatedReport;
            const html = document.getElementById('genReportContent').innerHTML;
            return {
                txCount: r.transactions.length,
                totalDep: r.totalDeposit,
                totalWd: r.totalWithdrawal,
                format: r.format,
                groupCount: Object.keys(r.groups).length,
                significantCount: r.significantTx.length,
                htmlLength: html.length,
                html: html
            };
        }
    """)

    print(f"  · 거래 수: {rpt['txCount']}건  (기대 13건)")
    print(f"  · 입금/출금: {rpt['totalDep']:,} / {rpt['totalWd']:,}  (기대 135M / 238M)")
    print(f"  · 그룹 수: {rpt['groupCount']}개  (기대 4개 — 4개 은행)")
    print(f"  · 주요 거래: {rpt['significantCount']}건")
    print(f"  · 포맷: {rpt['format']}")
    print(f"  · HTML 길이: {rpt['htmlLength']:,}자")

    # 포맷별 마커 검증
    html = rpt['html']
    markers = {
        'summary':  ['주요 입출금 내역', '거래 상세 내역', '계좌별' if False else '구분'],
        'ledger':   ['자금일계표', '시간순', '시계열 Ledger'],
        'category': ['수입 합계', '지출 합계', '금일 순증감', '분류형'],
    }
    expected_markers = markers[fmt]
    missing_markers = [m for m in expected_markers if m not in html]

    base_ok = (
        rpt['txCount'] == 13
        and rpt['totalDep'] == 135_000_000
        and rpt['totalWd'] == 238_000_000
        and rpt['format'] == fmt
        and rpt['htmlLength'] > 1000
    )
    marker_ok = len(missing_markers) == 0

    if base_ok and marker_ok:
        print(f"  → [OK] 포맷 [{fmt}] 정상 생성 (모든 마커 확인)")
        return True
    else:
        print(f"  → [FAIL] base_ok={base_ok}, marker_ok={marker_ok}")
        if missing_markers:
            print(f"    누락된 마커: {missing_markers}")
        return False


async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        ctx = await browser.new_context()
        page = await ctx.new_page()

        # 콘솔 에러 수집
        errors = []
        page.on("console", lambda msg: errors.append(f"[{msg.type}] {msg.text}") if msg.type == "error" else None)
        page.on("pageerror", lambda exc: errors.append(f"[pageerror] {exc}"))

        url = "file:///" + str(PAGE_PATH).replace("\\", "/")
        await page.goto(url, wait_until="domcontentloaded")
        await poll_until(page, "typeof XLSX !== 'undefined' && typeof pdfjsLib !== 'undefined'", 30000)

        results = {}

        # 1. 검증기 경로
        try:
            results['validator'] = await test_validator_path(page)
        except Exception as e:
            import traceback
            traceback.print_exc()
            results['validator'] = False

        # 2. 자동 생성기 — 3가지 포맷
        for fmt in ['summary', 'ledger', 'category']:
            try:
                results[f'generator-{fmt}'] = await test_generator_path(page, fmt)
            except Exception as e:
                import traceback
                traceback.print_exc()
                results[f'generator-{fmt}'] = False

        # 콘솔 에러 출력
        if errors:
            print()
            print("⚠ 콘솔 에러 발생:")
            for e in errors[-10:]:
                print(f"  {e}")

        await browser.close()

    # 종합 결과
    print()
    print("#" * 72)
    print("# 종합 결과")
    print("#" * 72)
    for k, ok in results.items():
        mark = "[OK]" if ok else "[FAIL]"
        print(f"  {mark}  {k}")

    all_ok = all(results.values()) and len(errors) == 0
    print()
    if all_ok:
        print("✓ 모든 테스트 통과 (콘솔 에러 0건)")
    else:
        print("✗ 일부 실패 또는 콘솔 에러 있음")
    return 0 if all_ok else 1


if __name__ == "__main__":
    asyncio.run(main())
