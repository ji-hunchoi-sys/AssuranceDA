# -*- coding: utf-8 -*-
"""
자금일보 검증기 두번째 테스트 세트 (PDF/PNG 형식)

생성 파일:
  - 자금일보_샘플2.pdf       / 자금일보_샘플2.png
  - 은행거래내역_샘플2.pdf   / 은행거래내역_샘플2.png

테스트 시나리오 (계획일자 2024-12-15):
  전일이월 현금시재:  200,000,000원
  수입:               25,000,000원  (외상대 수금)
  조달:               50,000,000원  (정부지원자금)
  지출:               58,000,000원
    · 외상대 결제:    35,000,000원  (= 10M + 12M + 13M, 3건 합산매칭)
    · 일반경비:        8,000,000원  (= 2M + 1M + 5M, 3건 합산매칭)
    · 급/상여:        15,000,000원  (1:1 매칭)
  조달 후 시재:      217,000,000원

매칭 검증 포인트:
  · 1:1 매칭     : 외상대 수금(25M), 정부지원자금(50M, 조달), 급/상여(15M)
  · N:1 합산매칭 : 외상대 결제(35M = 10+12+13), 일반경비(8M = 2+1+5)
  · 조달 섹션   : 정부지원자금이 조달로 분류되어 deposit으로 합산되는지
"""
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak,
)

OUT_DIR = Path(__file__).resolve().parent

# ─────────────────────────────────────────────────────────────────
# 한글 폰트 등록 (Windows 기본 폰트 우선)
# ─────────────────────────────────────────────────────────────────
def register_korean_font():
    candidates = [
        ("MalgunGothic", r"C:\Windows\Fonts\malgun.ttf"),
        ("MalgunGothicBold", r"C:\Windows\Fonts\malgunbd.ttf"),
        ("NanumGothic", r"C:\Windows\Fonts\NanumGothic.ttf"),
    ]
    registered = []
    for name, path in candidates:
        if Path(path).exists():
            try:
                pdfmetrics.registerFont(TTFont(name, path))
                registered.append(name)
            except Exception as e:
                print(f"  [WARN] font register failed: {name}: {e}")
    if not registered:
        raise RuntimeError("한글 폰트를 찾지 못했습니다 (Malgun/Nanum)")
    return registered[0], (registered[1] if len(registered) > 1 else registered[0])


KOR_FONT, KOR_FONT_BOLD = register_korean_font()


def styled(font=KOR_FONT, size=10, bold=False, align="LEFT"):
    return ParagraphStyle(
        "k", fontName=KOR_FONT_BOLD if bold else font, fontSize=size,
        leading=size * 1.4, alignment={"LEFT": 0, "CENTER": 1, "RIGHT": 2}[align],
    )


# ─────────────────────────────────────────────────────────────────
# 1) 자금일보 PDF
# ─────────────────────────────────────────────────────────────────
def build_fund_pdf():
    out = OUT_DIR / "자금일보_샘플2.pdf"
    doc = SimpleDocTemplate(
        str(out), pagesize=A4,
        leftMargin=18 * mm, rightMargin=18 * mm,
        topMargin=18 * mm, bottomMargin=18 * mm,
    )

    title_style = styled(size=18, bold=True, align="CENTER")
    sub_style = styled(size=10, align="CENTER")
    note_style = styled(size=9, align="LEFT")

    story = []
    story.append(Paragraph("일일 자금 계획", title_style))
    story.append(Spacer(1, 4 * mm))
    story.append(Paragraph(
        "회사명: (주)테스트회사   |   계획일자: 2024-12-15   |   단위: 원",
        sub_style,
    ))
    story.append(Spacer(1, 6 * mm))

    # 메타 잔액 박스 (parseDailyFundPlan이 키워드로 인식)
    meta_data = [
        ["전일이월 현금시재", "200,000,000원", "조달 후 현금시재", "217,000,000원"],
    ]
    meta = Table(meta_data, colWidths=[45 * mm, 45 * mm, 45 * mm, 45 * mm])
    meta.setStyle(TableStyle([
        ("FONT", (0, 0), (-1, -1), KOR_FONT_BOLD, 11),
        ("BACKGROUND", (0, 0), (0, 0), colors.HexColor("#E8EEF5")),
        ("BACKGROUND", (2, 0), (2, 0), colors.HexColor("#E8EEF5")),
        ("ALIGN", (1, 0), (1, 0), "RIGHT"),
        ("ALIGN", (3, 0), (3, 0), "RIGHT"),
        ("BOX", (0, 0), (-1, -1), 0.6, colors.grey),
        ("INNERGRID", (0, 0), (-1, -1), 0.4, colors.lightgrey),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))
    story.append(meta)
    story.append(Spacer(1, 6 * mm))

    # 본 표
    fmt = lambda n: "" if n in (None, "") else (f"{n:,}" if isinstance(n, (int, float)) else str(n))
    rows = [
        ["구분", "항목", "금액", "세부내역"],
        ["수입", "외상대 수금(원화)", 25_000_000, "거래처A 매출채권 회수"],
        ["수입", "외상대 수금(외화)", 0, ""],
        ["수입", "매출 수금", 0, ""],
        ["수입", "소계", 25_000_000, ""],
        ["지출", "차입금 상환", 0, ""],
        ["지출", "외상대 결제(원화)", 35_000_000, "거래처B 10M + 거래처C 12M + 거래처D 13M"],
        ["지출", "외상대 결제(외화)", 0, ""],
        ["지출", "설비투자", 0, ""],
        ["지출", "연구개발비", 0, ""],
        ["지출", "급/상여(회직급여)", 15_000_000, "12월 급여 지급"],
        ["지출", "일반경비", 8_000_000, "사무용품 2M + 통신비 1M + 임대료 5M"],
        ["지출", "전도금", 0, ""],
        ["지출", "소계", 58_000_000, ""],
        ["과부족", "", 17_000_000, ""],
        ["조달", "예금 입출(원화)", 0, ""],
        ["조달", "예금 입출(외화)", 0, ""],
        ["조달", "정부지원자금", 50_000_000, "중소기업 R&D 지원금 입금"],
        ["조달", "차입금", 0, ""],
        ["조달", "소계", 50_000_000, ""],
    ]
    table_rows = [[Paragraph(str(c), styled(size=9, bold=(i == 0), align="CENTER" if j != 3 else "LEFT")) if j != 2 else Paragraph(fmt(c), styled(size=9, bold=(i == 0), align="RIGHT")) for j, c in enumerate(r)] for i, r in enumerate(rows)]

    tbl = Table(table_rows, colWidths=[18 * mm, 50 * mm, 32 * mm, 78 * mm])
    style_cmds = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E8EEF5")),
        ("BOX", (0, 0), (-1, -1), 0.7, colors.grey),
        ("INNERGRID", (0, 0), (-1, -1), 0.3, colors.lightgrey),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]
    # 소계/과부족 행 강조
    for i, r in enumerate(rows):
        if "소계" in str(r[1]) or r[0] == "과부족":
            style_cmds.append(("BACKGROUND", (0, i), (-1, i), colors.HexColor("#F4F6F9")))
    tbl.setStyle(TableStyle(style_cmds))
    story.append(tbl)
    story.append(Spacer(1, 6 * mm))
    story.append(Paragraph("※ 본 자금일보의 수입 합계 = 25,000,000원 / 지출 합계 = 58,000,000원 / 조달 합계 = 50,000,000원", note_style))

    doc.build(story)
    print(f"  [OK] {out.name}")
    return out


# ─────────────────────────────────────────────────────────────────
# 2) 은행 거래내역 PDF
# ─────────────────────────────────────────────────────────────────
def build_bank_pdf():
    out = OUT_DIR / "은행거래내역_샘플2.pdf"
    doc = SimpleDocTemplate(
        str(out), pagesize=A4,
        leftMargin=14 * mm, rightMargin=14 * mm,
        topMargin=15 * mm, bottomMargin=15 * mm,
    )

    title_style = styled(size=16, bold=True, align="CENTER")
    sub_style = styled(size=9, align="CENTER")

    story = []
    story.append(Paragraph("거래내역 조회", title_style))
    story.append(Spacer(1, 3 * mm))
    story.append(Paragraph(
        "계좌번호: 987-654321-02-002   |   예금주: (주)테스트회사   |   조회기간: 2024-12-15 ~ 2024-12-15",
        sub_style,
    ))
    story.append(Spacer(1, 5 * mm))

    txns = []
    bal = 200_000_000

    def add(time, deposit, withdrawal, desc, vendor):
        nonlocal bal
        bal = bal + deposit - withdrawal
        txns.append((f"2024-12-15 {time}", deposit, withdrawal, bal, desc, vendor))

    # 수입 (외상대 수금)
    add("09:30:11", 25_000_000, 0, "외상대 수금", "거래처A 주식회사")
    # 조달 (정부지원자금)
    add("10:15:33", 50_000_000, 0, "정부지원금 입금", "중소기업청")
    # 지출: 일반경비 8M = 2M + 1M + 5M
    add("11:02:45", 0, 2_000_000, "사무용품 구입",   "오피스월드")
    add("11:18:09", 0, 1_000_000, "통신비 자동이체", "KT")
    add("13:30:22", 0, 5_000_000, "임대료 지급",     "강남빌딩")
    # 지출: 외상대 결제 35M = 10M + 12M + 13M
    add("14:05:18", 0, 10_000_000, "외상대 결제", "거래처B 상사")
    add("14:06:02", 0, 12_000_000, "외상대 결제", "거래처C 상사")
    add("14:06:47", 0, 13_000_000, "외상대 결제", "거래처D 상사")
    # 지출: 급/상여 15M (1:1 매칭, 일괄지급)
    add("16:00:55", 0, 15_000_000, "12월 급여 지급", "임직원 일괄")

    headers = ["거래일시", "입금액", "출금액", "잔액", "적요", "거래처"]
    fmt = lambda n: f"{n:,}" if isinstance(n, (int, float)) and n != 0 else ("0" if n == 0 else str(n))
    data = [headers]
    for t in txns:
        data.append([
            t[0],
            f"{t[1]:,}" if t[1] else "0",
            f"{t[2]:,}" if t[2] else "0",
            f"{t[3]:,}",
            t[4],
            t[5],
        ])

    # 합계 행도 추가 (실제 은행 PDF처럼)
    data.append([
        "합계",
        f"{sum(t[1] for t in txns):,}",
        f"{sum(t[2] for t in txns):,}",
        "", "", "",
    ])

    table_rows = []
    for i, row in enumerate(data):
        cells = []
        for j, c in enumerate(row):
            align = "RIGHT" if j in (1, 2, 3) else ("CENTER" if j == 0 else "LEFT")
            cells.append(Paragraph(str(c), styled(size=8, bold=(i == 0 or i == len(data) - 1), align=align)))
        table_rows.append(cells)

    tbl = Table(table_rows, colWidths=[34 * mm, 23 * mm, 23 * mm, 28 * mm, 35 * mm, 39 * mm])
    style_cmds = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E8EEF5")),
        ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#F4F6F9")),
        ("BOX", (0, 0), (-1, -1), 0.7, colors.grey),
        ("INNERGRID", (0, 0), (-1, -1), 0.3, colors.lightgrey),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]
    tbl.setStyle(TableStyle(style_cmds))
    story.append(tbl)

    doc.build(story)
    print(f"  [OK] {out.name}")
    return out


# ─────────────────────────────────────────────────────────────────
# 3) PDF → PNG 변환 (PIL 직접 렌더링 — pdf2image/poppler 없이)
# ─────────────────────────────────────────────────────────────────
def render_png_from_data():
    """PIL로 깔끔한 표를 직접 그려 PNG로 저장. OCR 친화적으로 큰 폰트·고해상도 사용."""
    from PIL import Image, ImageDraw, ImageFont

    win_font_dir = Path(r"C:\Windows\Fonts")
    font_path = win_font_dir / "malgun.ttf"
    font_bold_path = win_font_dir / "malgunbd.ttf"
    if not font_path.exists():
        raise RuntimeError("malgun.ttf 한글 폰트를 찾지 못했습니다")

    def fnt(size, bold=False):
        return ImageFont.truetype(str(font_bold_path if bold and font_bold_path.exists() else font_path), size)

    # ── 자금일보 PNG (OCR 친화적 레이아웃 — 큰 글씨, 흰색 배경, 0 항목 제외) ──
    rows = [
        ("수입", "외상대 수금", "25,000,000", "거래처A 매출채권 회수"),
        ("지출", "외상대 결제", "35,000,000", "거래처B C D 합계"),
        ("지출", "급/상여", "15,000,000", "12월 급여 지급"),
        ("지출", "일반경비", "8,000,000", "사무용품 통신비 임대료"),
        ("조달", "정부지원자금", "50,000,000", "중소기업 R&D 지원금"),
    ]

    W = 2400
    line_h = 110
    top = 360
    H = top + line_h * (len(rows) + 2) + 200
    img = Image.new("RGB", (W, H), "white")
    d = ImageDraw.Draw(img)

    # 제목
    d.text((W // 2 - 230, 50), "일일 자금 계획", font=fnt(64, bold=True), fill="black")
    d.text((W // 2 - 470, 160), "회사명 (주)테스트회사    계획일자 2024-12-15    단위 원",
           font=fnt(34), fill="black")

    # 메타 (전일이월/조달후) — 단순 텍스트로
    y = 270
    d.text((140, y), "전일이월 현금시재", font=fnt(36, bold=True), fill="black")
    d.text((900, y), "200,000,000 원", font=fnt(36, bold=True), fill="black")
    d.text((1500, y), "조달 후 현금시재", font=fnt(36, bold=True), fill="black")
    d.text((2050, y), "217,000,000 원", font=fnt(36, bold=True), fill="black")

    # 표 헤더 — 배경색 없이 단순 줄로
    y = top
    cols = [120, 360, 1000, 1640, W - 120]
    headers = ("구분", "항목", "금액", "세부내역")
    fh = fnt(36, bold=True)
    for j, h in enumerate(headers):
        x = cols[j] + (30 if j != 2 else 0)
        if j == 2:
            tw = d.textlength(h, font=fh)
            x = cols[3] - tw - 30
        d.text((x, y + 30), h, font=fh, fill="black")
    # 헤더 하단 굵은 줄
    d.line((cols[0], y + line_h - 5, cols[-1], y + line_h - 5), fill="black", width=3)
    y += line_h

    fr = fnt(34)
    for r in rows:
        # 각 셀 텍스트
        d.text((cols[0] + 30, y + 30), r[0], font=fr, fill="black")
        d.text((cols[1] + 30, y + 30), r[1], font=fr, fill="black")
        # 금액 우측정렬
        tw = d.textlength(r[2], font=fr)
        d.text((cols[3] - tw - 30, y + 30), r[2], font=fr, fill="black")
        d.text((cols[3] + 30, y + 30), r[3], font=fr, fill="black")
        # 행 구분선 (얇은 회색)
        d.line((cols[0], y + line_h - 1, cols[-1], y + line_h - 1), fill="#bbb", width=1)
        y += line_h

    out = OUT_DIR / "자금일보_샘플2.png"
    img.save(out, "PNG", dpi=(200, 200))
    print(f"  [OK] {out.name}")

    # ── 은행 거래내역 PNG ──
    W2, H2 = 2400, 1600
    img2 = Image.new("RGB", (W2, H2), "white")
    d2 = ImageDraw.Draw(img2)

    d2.text((W2 // 2 - 180, 40), "거래내역 조회", font=fnt(54, bold=True), fill="black")
    d2.text((W2 // 2 - 600, 120), "계좌번호: 987-654321-02-002   |   예금주: (주)테스트회사",
            font=fnt(26), fill="#444")
    d2.text((W2 // 2 - 380, 165), "조회기간: 2024-12-15 ~ 2024-12-15",
            font=fnt(26), fill="#444")

    txns = [
        ("2024-12-15 09:30:11", "25,000,000", "0",          "225,000,000", "외상대 수금",     "거래처A 주식회사"),
        ("2024-12-15 10:15:33", "50,000,000", "0",          "275,000,000", "정부지원금 입금", "중소기업청"),
        ("2024-12-15 11:02:45", "0",          "2,000,000",  "273,000,000", "사무용품 구입",   "오피스월드"),
        ("2024-12-15 11:18:09", "0",          "1,000,000",  "272,000,000", "통신비 자동이체", "KT"),
        ("2024-12-15 13:30:22", "0",          "5,000,000",  "267,000,000", "임대료 지급",     "강남빌딩"),
        ("2024-12-15 14:05:18", "0",          "10,000,000", "257,000,000", "외상대 결제",     "거래처B 상사"),
        ("2024-12-15 14:06:02", "0",          "12,000,000", "245,000,000", "외상대 결제",     "거래처C 상사"),
        ("2024-12-15 14:06:47", "0",          "13,000,000", "232,000,000", "외상대 결제",     "거래처D 상사"),
        ("2024-12-15 16:00:55", "0",          "15,000,000", "217,000,000", "12월 급여 지급",  "임직원 일괄"),
    ]
    headers = ("거래일시", "입금액", "출금액", "잔액", "적요", "거래처")

    cols = [50, 510, 730, 950, 1230, 1640, W2 - 50]
    y = 240
    row_h = 75

    # 헤더
    d2.rectangle((cols[0], y, cols[-1], y + row_h), fill="#E8EEF5", outline="#888", width=1)
    for j in range(1, len(cols) - 1):
        d2.line((cols[j], y, cols[j], y + row_h), fill="#888", width=1)
    f = fnt(26, bold=True)
    for j, h in enumerate(headers):
        cell_x_left = cols[j]
        cell_x_right = cols[j + 1]
        text_w = d2.textlength(h, font=f)
        if j == 0 or j >= 4:
            tx = cell_x_left + 20
        else:
            tx = cell_x_right - text_w - 20
        d2.text((tx, y + 22), h, font=f, fill="black")
    y += row_h

    f_data = fnt(24)
    for t in txns:
        d2.rectangle((cols[0], y, cols[-1], y + row_h), outline="#bbb", width=1)
        for j in range(1, len(cols) - 1):
            d2.line((cols[j], y, cols[j], y + row_h), fill="#bbb", width=1)
        for j, val in enumerate(t):
            cell_x_left = cols[j]
            cell_x_right = cols[j + 1]
            tw = d2.textlength(val, font=f_data)
            if j in (1, 2, 3):
                tx = cell_x_right - tw - 20
            elif j == 0:
                tx = cell_x_left + 15
            else:
                tx = cell_x_left + 15
            d2.text((tx, y + 22), val, font=f_data, fill="black")
        y += row_h

    # 합계
    d2.rectangle((cols[0], y, cols[-1], y + row_h), fill="#F4F6F9", outline="#888", width=1)
    for j in range(1, len(cols) - 1):
        d2.line((cols[j], y, cols[j], y + row_h), fill="#888", width=1)
    fb = fnt(26, bold=True)
    d2.text((cols[0] + 15, y + 22), "합계", font=fb, fill="black")
    sum_dep = "75,000,000"
    sum_wd = "58,000,000"
    tw = d2.textlength(sum_dep, font=fb)
    d2.text((cols[2] - tw - 20, y + 22), sum_dep, font=fb, fill="black")
    tw = d2.textlength(sum_wd, font=fb)
    d2.text((cols[3] - tw - 20, y + 22), sum_wd, font=fb, fill="black")

    out2 = OUT_DIR / "은행거래내역_샘플2.png"
    img2.save(out2, "PNG", dpi=(200, 200))
    print(f"  [OK] {out2.name}")


if __name__ == "__main__":
    print("PDF/PNG 두번째 샘플 생성 중...")
    build_fund_pdf()
    build_bank_pdf()
    render_png_from_data()
    print()
    print("생성 위치: test_samples/ 폴더")
    print("  - 자금일보_샘플2.pdf, .png")
    print("  - 은행거래내역_샘플2.pdf, .png")
