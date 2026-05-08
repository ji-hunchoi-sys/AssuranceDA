# -*- coding: utf-8 -*-
"""
샘플6(다중계좌) 검증 후 audit 패널의 '기초/마감 시재 불일치' 메시지가
사라졌는지 확인하는 빠른 회귀 테스트.
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


async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        ctx = await browser.new_context()
        page = await ctx.new_page()
        url = "file:///" + str(PAGE_PATH).replace("\\", "/")
        await page.goto(url, wait_until="domcontentloaded")
        await poll_until(page, "typeof XLSX !== 'undefined' && typeof pdfjsLib !== 'undefined'", 30000)

        await page.evaluate("() => showPage('upload')")

        # 자금일보 + 은행 4개 한 번에 업로드
        await page.set_input_files("#fundFileInput", str(SAMPLES_DIR / "자금일보_샘플6.pdf"))
        await poll_until(page, "APP.fundData !== null && APP.fundData.transactions.length > 0", 90000)
        bank_files = [str(SAMPLES_DIR / f"은행거래내역_샘플6_{k}.pdf") for k in BANK_KEYS]
        await page.set_input_files("#bankFileInput", bank_files)
        await poll_until(page, "APP.bankFiles && APP.bankFiles.length === 4", 240000)

        # 검증 실행
        await page.wait_for_selector("#verifyBtn:not([disabled])", timeout=10000)
        await page.click("#verifyBtn")
        await poll_until(page, "APP.results && APP.results.length > 0", 30000)
        await page.wait_for_timeout(500)

        # audit 패널 텍스트 추출
        audit_html = await page.evaluate(
            "() => document.getElementById('auditPointsContent')?.innerHTML || ''"
        )

        # 결과 매칭/불일치 메시지
        results = await page.evaluate("""
            () => {
                const balanceResults = (APP.results || []).filter(r => r.type === '기초시재' || r.type === '마감시재');
                return balanceResults.map(r => ({
                    type: r.type, status: r.status, fundAmount: r.fundAmount,
                    bankAmount: r.bankAmount, diff: r.diff, note: r.note
                }));
            }
        """)

        await browser.close()

    print("=" * 70)
    print("  시재 검증 결과 (APP.results)")
    print("=" * 70)
    for r in results:
        mark = "✓" if r['status'] == 'match' else "✗"
        print(f"  {mark} [{r['type']}] {r['status']}  자금일보={r['fundAmount']:,}  은행={r['bankAmount']:,}  차이={r['diff']:,}")
        print(f"       note: {r['note']}")

    print()
    print("=" * 70)
    print("  Audit 패널 시재 불일치 메시지 검사")
    print("=" * 70)
    has_opening_audit = '기초시재 불일치' in audit_html
    has_closing_audit = '마감시재 불일치' in audit_html
    print(f"  · '기초시재 불일치' 메시지: {'있음 ✗' if has_opening_audit else '없음 ✓'}")
    print(f"  · '마감시재 불일치' 메시지: {'있음 ✗' if has_closing_audit else '없음 ✓'}")

    all_match = all(r['status'] == 'match' for r in results)
    no_audit_warning = not has_opening_audit and not has_closing_audit
    print()
    if all_match and no_audit_warning:
        print("✓ 시재 검증 통과 — 메인 검증 + audit 모두 일치")
    else:
        print("✗ 여전히 시재 불일치 메시지 존재")


if __name__ == "__main__":
    asyncio.run(main())
