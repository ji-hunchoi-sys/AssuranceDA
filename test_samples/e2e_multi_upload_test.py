# -*- coding: utf-8 -*-
"""
다중 파일 동시 업로드 E2E 테스트.

이전엔 은행 파일을 한 번에 하나씩만 업로드 가능 → 4개 은행 파일을 4번
클릭해서 누적해야 했음. 이제 multiple 속성 + 시퀀셜 핸들러로 한 번 클릭에
여러 파일을 선택해 일괄 업로드 가능.

검증:
  TEST 1) 검증 페이지: 자금일보 1개 + 은행 4개를 한 번에 업로드 (set_input_files
          에 list 전달) → 13건 인식 + 100% 매칭
  TEST 2) 자동 생성 페이지: 은행 4개 한 번에 업로드 → 13건 인식 + 자금일보 생성
"""
import asyncio
from pathlib import Path

from playwright.async_api import async_playwright

ROOT = Path(__file__).resolve().parent.parent
PAGE_PATH = ROOT / "index2_pdf.html"
SAMPLES_DIR = Path(__file__).resolve().parent
BANK_KEYS = ["신한", "국민", "하나", "우리"]


async def poll_until(page, expr, timeout_ms=120000, interval_ms=200):
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


async def test_validator_multi(page):
    """검증 페이지 — 은행 4개 한 번에 업로드."""
    print()
    print("=" * 72)
    print("  TEST 1: 검증 페이지 — 자금일보 1 + 은행 4 한 번에 업로드")
    print("=" * 72)

    await page.evaluate("() => showPage('upload')")
    await page.wait_for_timeout(300)

    # 자금일보 (단일)
    fund_pdf = SAMPLES_DIR / "자금일보_샘플6.pdf"
    print(f"  · 자금일보 업로드: {fund_pdf.name}")
    await page.set_input_files("#fundFileInput", str(fund_pdf))
    await poll_until(page, "APP.fundData !== null && APP.fundData.transactions.length > 0", 90000)

    # 은행 4개 한 번에 업로드 — multiple 동작 핵심 부분
    bank_files = [str(SAMPLES_DIR / f"은행거래내역_샘플6_{k}.pdf") for k in BANK_KEYS]
    print(f"  · 은행 4개 한 번에 업로드 (set_input_files에 list 전달)")
    await page.set_input_files("#bankFileInput", bank_files)

    # 4개 파일 모두 처리 완료 대기 (genBankFiles 길이 = 4)
    await poll_until(page, "APP.bankFiles && APP.bankFiles.length === 4", 240000)
    bank_total = await page.evaluate(
        "() => APP.bankFiles.reduce((s,f) => s + (f.transactions||[]).length, 0)"
    )
    bank_files_count = await page.evaluate("() => APP.bankFiles.length")
    print(f"  · 결과: {bank_files_count}개 파일, 총 거래 {bank_total}건")

    if bank_total != 13 or bank_files_count != 4:
        print(f"  → [FAIL] 기대 4파일/13건, 실제 {bank_files_count}파일/{bank_total}건")
        return False

    # 검증 실행
    await page.wait_for_selector("#verifyBtn:not([disabled])", timeout=10000)
    await page.click("#verifyBtn")
    await poll_until(page, "APP.results && APP.results.length > 0", 30000)
    await page.wait_for_timeout(300)

    s = await page.evaluate("""
        () => {
            const statuses = {};
            for (const r of (APP.results || [])) statuses[r.status] = (statuses[r.status]||0) + 1;
            return { statuses };
        }
    """)
    print(f"  · 검증 결과: {s['statuses']}")
    ok = (
        s['statuses'].get('mismatch', 0) == 0
        and s['statuses'].get('missing', 0) == 0
        and s['statuses'].get('extra', 0) == 0
    )
    print(f"  → {'[OK] 다중 업로드 + 검증 100%' if ok else '[FAIL] 검증 오류'}")
    return ok


async def test_generator_multi(page):
    """자동 생성 페이지 — 은행 4개 한 번에 업로드."""
    print()
    print("=" * 72)
    print("  TEST 2: 자동 생성 페이지 — 은행 4개 한 번에 업로드")
    print("=" * 72)

    await page.evaluate("() => showPage('generate')")
    await page.wait_for_timeout(300)
    await page.evaluate("() => { genBankFiles.length = 0; updateGenBankFileListUI(); }")
    await page.wait_for_timeout(200)

    # 4개 파일 한 번에 업로드
    bank_files = [str(SAMPLES_DIR / f"은행거래내역_샘플6_{k}.pdf") for k in BANK_KEYS]
    print(f"  · 은행 4개 한 번에 업로드 (set_input_files에 list 전달)")
    await page.set_input_files("#genBankInput", bank_files)

    # 4개 파일 모두 처리 완료 대기
    await poll_until(page, "genBankFiles.length === 4", 240000)
    bank_count = await page.evaluate(
        "() => genBankFiles.reduce((s,f) => s + (f.transactions||[]).length, 0)"
    )
    print(f"  · 결과: {await page.evaluate('() => genBankFiles.length')}개 파일, 총 거래 {bank_count}건")

    if bank_count != 13:
        print(f"  → [FAIL] 기대 13건, 실제 {bank_count}건")
        return False

    # 자금일보 생성 (요약형)
    await page.fill("#genCompany", "(주)한국상사")
    await page.fill("#genDate", "2026-05-07")
    await page.select_option("#genFormat", "summary")
    await page.evaluate("() => generateFundReport()")
    await poll_until(page, "APP.generatedReport !== null && APP.generatedReport !== undefined", 5000)

    rpt = await page.evaluate("""
        () => ({
            txCount: APP.generatedReport.transactions.length,
            totalDep: APP.generatedReport.totalDeposit,
            totalWd: APP.generatedReport.totalWithdrawal,
            groupCount: Object.keys(APP.generatedReport.groups).length
        })
    """)
    print(f"  · 자금일보 생성: {rpt['txCount']}건, 입금 {rpt['totalDep']:,} / 출금 {rpt['totalWd']:,}, "
          f"그룹 {rpt['groupCount']}개")

    ok = (
        rpt['txCount'] == 13
        and rpt['totalDep'] == 135_000_000
        and rpt['totalWd'] == 238_000_000
        and rpt['groupCount'] == 4
    )
    print(f"  → {'[OK] 다중 업로드 + 자금일보 생성 정상' if ok else '[FAIL] 생성 오류'}")
    return ok


async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        page = await browser.new_context().new_page() if False else await (await browser.new_context()).new_page()

        errors = []
        page.on("console", lambda msg: errors.append(f"[{msg.type}] {msg.text}") if msg.type == "error" else None)
        page.on("pageerror", lambda exc: errors.append(f"[pageerror] {exc}"))

        url = "file:///" + str(PAGE_PATH).replace("\\", "/")
        await page.goto(url, wait_until="domcontentloaded")
        await poll_until(page, "typeof XLSX !== 'undefined' && typeof pdfjsLib !== 'undefined'", 30000)

        results = {}
        try:
            results['validator-multi'] = await test_validator_multi(page)
        except Exception as e:
            import traceback
            traceback.print_exc()
            results['validator-multi'] = False

        try:
            results['generator-multi'] = await test_generator_multi(page)
        except Exception as e:
            import traceback
            traceback.print_exc()
            results['generator-multi'] = False

        if errors:
            print()
            print("⚠ 콘솔 에러:")
            for e in errors[-10:]:
                print(f"  {e}")

        await browser.close()

    print()
    print("#" * 72)
    print("# 종합 결과")
    print("#" * 72)
    for k, ok in results.items():
        print(f"  {'[OK]' if ok else '[FAIL]'}  {k}")
    print()
    all_ok = all(results.values()) and len(errors) == 0
    print("✓ 모든 테스트 통과 (콘솔 에러 0건)" if all_ok else "✗ 일부 실패 또는 콘솔 에러")
    return 0 if all_ok else 1


if __name__ == "__main__":
    asyncio.run(main())
