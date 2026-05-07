# -*- coding: utf-8 -*-
"""
실제 헤드리스 브라우저(Chromium)에서 index2_pdf.html을 띄우고
파일을 업로드한 뒤 검증을 실행해 결과 화면의 정확성을 확인한다.

PDF.js, Tesseract.js, OpenCV.js, xlsx.js 등이 모두 실제로 동작한다.
Python 시뮬레이션이 놓치는 브라우저-특정 동작도 잡을 수 있다.
"""
import asyncio
from pathlib import Path

from playwright.async_api import async_playwright

ROOT = Path(__file__).resolve().parent.parent
PAGE_PATH = ROOT / "index2_pdf.html"
SAMPLES_DIR = Path(__file__).resolve().parent


async def poll_until(page, expr: str, timeout_ms: int = 60000, interval_ms: int = 200):
    """page.evaluate를 폴링해서 expr가 truthy가 될 때까지 대기."""
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


async def upload_and_verify(fund_path: Path, bank_path: Path, label: str):
    """단일 케이스 실행 → 파싱 결과/검증 결과 dict 반환."""
    print(f"\n{'=' * 70}")
    print(f"  E2E 테스트: {label}")
    print(f"  자금일보: {fund_path.name}")
    print(f"  은행:     {bank_path.name}")
    print('=' * 70)

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        ctx = await browser.new_context()
        page = await ctx.new_page()

        # 콘솔 메시지 수집 (디버깅용)
        console_logs = []
        page.on("console", lambda msg: console_logs.append(f"[{msg.type}] {msg.text}"))
        page.on("pageerror", lambda exc: console_logs.append(f"[error] {exc}"))

        # 페이지 로드 (file:// 프로토콜)
        url = "file:///" + str(PAGE_PATH).replace("\\", "/")
        await page.goto(url, wait_until="domcontentloaded")
        # 페이지 안의 의존 스크립트 로드 대기
        await poll_until(page, "typeof XLSX !== 'undefined' && typeof pdfjsLib !== 'undefined'", 30000)

        # 자금일보 업로드
        print("  · 자금일보 업로드...")
        await page.set_input_files("#fundFileInput", str(fund_path))
        await poll_until(page, "typeof APP !== 'undefined' && APP.fundData !== null", 180000)
        await page.wait_for_timeout(500)
        fund_count = await page.evaluate("() => APP.fundData.transactions.length")
        print(f"    → 자금일보 거래 {fund_count}건")
        if fund_count == 0:
            # 자금일보 OCR/PDF 원문 캡처
            debug_text = await page.evaluate("() => document.getElementById('fundOcrDebugText')?.textContent || ''")
            if debug_text:
                # OCR 원문 부분만 보기 좋게 출력
                lines = debug_text.split('\n')
                ocr_start = next((i for i, l in enumerate(lines) if 'OCR/PDF 원문' in l), -1)
                if ocr_start >= 0:
                    print('    [DEBUG] OCR/PDF 원문:')
                    for line in lines[ocr_start+1:ocr_start+40]:
                        print(f'      {line}')

        # 은행 업로드
        print("  · 은행 업로드...")
        await page.set_input_files("#bankFileInput", str(bank_path))
        await poll_until(page, "typeof APP !== 'undefined' && APP.bankFiles && APP.bankFiles.length > 0", 180000)
        await page.wait_for_timeout(500)
        bank_count = await page.evaluate("() => APP.bankFiles.reduce((s,f) => s + (f.transactions||[]).length, 0)")
        print(f"    → 은행 거래 {bank_count}건")

        # 검증 버튼 활성화 대기 후 클릭
        await page.wait_for_selector("#verifyBtn:not([disabled])", timeout=10000)
        print("  · 검증 실행...")
        await page.click("#verifyBtn")
        try:
            await poll_until(page, "typeof APP !== 'undefined' && APP.results && APP.results.length > 0", 30000)
        except TimeoutError:
            print("    → 결과 미생성. 콘솔 로그 마지막 10줄:")
            for log in console_logs[-10:]:
                print(f"      {log}")
            # 토스트 메시지 캡쳐
            toast = await page.evaluate("() => document.getElementById('toast')?.textContent || '(no toast)'")
            print(f"    → 토스트: {toast}")
            return False, {"results": [], "fundTx": [], "bankTx": []}
        await page.wait_for_timeout(300)

        # APP.fundData / APP.results 추출
        snapshot = await page.evaluate("""
            () => {
                const fundTx = (APP.fundData && APP.fundData.transactions) || [];
                const bankTx = [];
                for (const f of (APP.bankFiles || [])) {
                    for (const t of (f.transactions || [])) bankTx.push(t);
                }
                const results = APP.results || [];
                return {
                    fundDate: APP.fundData && APP.fundData.date,
                    fundOpening: APP.fundData && APP.fundData.totals && APP.fundData.totals.openingBalance || 0,
                    fundClosing: APP.fundData && APP.fundData.totals && APP.fundData.totals.closingBalance || 0,
                    fundTx: fundTx.map(t => ({account: t.account, deposit: t.deposit, withdrawal: t.withdrawal})),
                    bankTx: bankTx.map(t => ({date: t.date, deposit: t.deposit, withdrawal: t.withdrawal, description: t.description})),
                    results: results.map(r => ({
                        status: r.status, type: r.type, fundAmount: r.fundAmount,
                        bankAmount: r.bankAmount, diff: r.diff, description: r.description, note: r.note,
                    })),
                };
            }
        """)

        await browser.close()

    # 화면에 결과 출력
    f = snapshot
    print(f"\n[자금일보 파싱]  날짜={f['fundDate']}  기초={f['fundOpening']:,}  마감={f['fundClosing']:,}")
    print(f"[자금일보 거래 {len(f['fundTx'])}건]")
    for t in f['fundTx']:
        print(f"  - {t['account'][:50]:50}  수입={t['deposit']:>13,}  지출={t['withdrawal']:>13,}")

    print(f"\n[은행 거래 {len(f['bankTx'])}건]")
    for t in f['bankTx'][:15]:
        print(f"  - {(t['date'] or '')[:20]:20}  입금={t['deposit']:>13,}  출금={t['withdrawal']:>13,}  {t['description'] or ''}")

    fund_dep = sum(t['deposit'] for t in f['fundTx'])
    fund_wd = sum(t['withdrawal'] for t in f['fundTx'])
    bank_dep = sum(t['deposit'] for t in f['bankTx'])
    bank_wd = sum(t['withdrawal'] for t in f['bankTx'])

    print(f"\n[합계]")
    print(f"  자금일보: 수입 {fund_dep:>13,}원   지출 {fund_wd:>13,}원")
    print(f"  은행:     입금 {bank_dep:>13,}원   출금 {bank_wd:>13,}원")
    ok_dep = abs(fund_dep - bank_dep) <= 100
    ok_wd = abs(fund_wd - bank_wd) <= 100
    print(f"  합계 일치: 수입 {'[OK]' if ok_dep else '[X]'}, 지출 {'[OK]' if ok_wd else '[X]'}")

    # 결과 분류
    statuses = {}
    for r in f['results']:
        statuses[r['status']] = statuses.get(r['status'], 0) + 1

    n_match = statuses.get('match', 0)
    n_mismatch = statuses.get('mismatch', 0)
    n_missing = statuses.get('missing', 0)
    n_extra = statuses.get('extra', 0)

    print(f"\n[검증 결과 분류]")
    print(f"  match    : {n_match}")
    print(f"  mismatch : {n_mismatch}")
    print(f"  missing  : {n_missing}")
    print(f"  extra    : {n_extra}")

    print(f"\n[상세 결과]")
    for r in f['results'][:20]:
        emoji = {'match': '[OK]', 'mismatch': '[X]', 'missing': '[누락]', 'extra': '[초과]'}.get(r['status'], '[?]')
        desc = (r['description'] or '')[:60]
        print(f"  {emoji:6} {r['type']:8} fund={r['fundAmount']:>11,} vs bank={r['bankAmount']:>11,}  {desc}")

    overall_ok = ok_dep and ok_wd and n_mismatch == 0 and n_missing == 0 and n_extra == 0
    print(f"\n{'=' * 70}")
    print(f"  최종: {'[성공] 100% 일치' if overall_ok else '[실패] 검증 오류 있음'}")
    print('=' * 70)

    return overall_ok, snapshot


async def main():
    cases = [
        (SAMPLES_DIR / "자금일보_샘플.xlsx", SAMPLES_DIR / "은행거래내역_샘플.xlsx", "Excel 경로 (1차 샘플)"),
        (SAMPLES_DIR / "자금일보_샘플2.pdf", SAMPLES_DIR / "은행거래내역_샘플2.pdf", "PDF 경로 (2차 샘플)"),
        (SAMPLES_DIR / "자금일보_샘플2.png", SAMPLES_DIR / "은행거래내역_샘플2.png", "PNG/OCR 경로 (2차 샘플)"),
    ]
    summary = []
    for fund, bank, label in cases:
        try:
            ok, _ = await upload_and_verify(fund, bank, label)
            summary.append((label, ok))
        except Exception as e:
            import traceback
            traceback.print_exc()
            summary.append((label, False))

    print(f"\n\n{'#' * 70}")
    print("# 종합 결과")
    print('#' * 70)
    for label, ok in summary:
        print(f"  {('[성공]' if ok else '[실패]')}  {label}")


if __name__ == "__main__":
    asyncio.run(main())
