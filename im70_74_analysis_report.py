from __future__ import annotations

import sys
from datetime import datetime
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont
from reportlab.lib.pagesizes import landscape, A4
from reportlab.pdfgen import canvas


SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPT_DIR))

import im70_74_excel_com as core  # noqa: E402


W, H = 3300, 2200
BLUE = "#174A7C"
INK = "#102A43"
MUTED = "#52616F"
GRID = "#D9E2EC"
PALE = "#F4F7FA"
ALT = "#F7FAFC"
COLORS = ["#1B9E77", "#D95F02", "#7570B3", "#2C7FB8", "#66A61E", "#E7298A", "#A6761D", "#1F78B4"]


def font(size: int, bold: bool = False):
    for p in [
        r"C:\Windows\Fonts\arialbd.ttf" if bold else r"C:\Windows\Fonts\arial.ttf",
        r"C:\Windows\Fonts\calibrib.ttf" if bold else r"C:\Windows\Fonts\calibri.ttf",
    ]:
        if Path(p).exists():
            return ImageFont.truetype(p, size)
    return ImageFont.load_default()


def fmt_num(v: float) -> str:
    if abs(float(v)) < 0.005:
        return "0"
    return f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", " ")


def fmt_int(v: float) -> str:
    return f"{int(v):,}".replace(",", " ")


def trunc(text, n: int) -> str:
    text = core.to_latin(str(text))
    return text if len(text) <= n else text[: max(0, n - 1)] + "..."


def new_page(title: str, subtitle: str = ""):
    img = Image.new("RGB", (W, H), "white")
    draw = ImageDraw.Draw(img)
    draw.text((60, 34), title, fill=INK, font=font(50, True))
    if subtitle:
        draw.text((62, 98), subtitle, fill=MUTED, font=font(28))
    draw.line((60, 145, W - 60, 145), fill=BLUE, width=5)
    return img, draw


def draw_kpis(draw, kpis: list[tuple[str, str]]):
    x = 60
    box_w = 600
    for label, value in kpis:
        draw.rounded_rectangle((x, 185, x + box_w, 340), radius=12, fill=PALE, outline="#B8C7D9", width=2)
        draw.text((x + 24, 210), label, fill=MUTED, font=font(28, True))
        draw.text((x + 24, 258), value, fill=INK, font=font(42, True))
        x += box_w + 28


def draw_table(draw, x: int, y: int, title: str, headers: list[str], rows: list[list], widths: list[int], row_h: int = 52):
    draw.text((x, y), title, fill=INK, font=font(34, True))
    y += 58
    cx = x
    for i, head in enumerate(headers):
        draw.rectangle((cx, y, cx + widths[i], y + row_h), fill=BLUE)
        draw.text((cx + 10, y + 13), str(head), fill="white", font=font(22, True))
        cx += widths[i]
    y += row_h
    for ri, row in enumerate(rows):
        cx = x
        fill = ALT if ri % 2 == 0 else "white"
        for i, val in enumerate(row):
            draw.rectangle((cx, y, cx + widths[i], y + row_h), fill=fill, outline=GRID)
            draw.text((cx + 10, y + 13), str(val), fill="#1B1F23", font=font(22))
            cx += widths[i]
        y += row_h


def draw_barh(draw, x: int, y: int, title: str, data: list[tuple[str, float]], width: int, label_w: int = 360):
    draw.text((x, y), title, fill=INK, font=font(34, True))
    y += 62
    max_v = max([v for _, v in data] or [1])
    for i, (label, value) in enumerate(data):
        yy = y + i * 62
        bar_w = int((value / max_v) * (width - label_w - 110))
        draw.text((x, yy + 5), trunc(label, 34), fill="#1B1F23", font=font(23))
        draw.rectangle((x + label_w, yy + 10, x + label_w + bar_w, yy + 44), fill=COLORS[i % len(COLORS)])
        draw.text((x + label_w + bar_w + 14, yy + 6), fmt_int(value), fill="#1B1F23", font=font(23, True))


def draw_donut(draw, cx: int, cy: int, radius: int, title: str, data: list[tuple[str, float]]):
    draw.text((cx - radius, cy - radius - 70), title, fill=INK, font=font(34, True))
    total = sum(v for _, v in data) or 1
    start = -90
    for i, (_, value) in enumerate(data):
        angle = value / total * 360
        draw.pieslice((cx - radius, cy - radius, cx + radius, cy + radius), start, start + angle, fill=COLORS[i % len(COLORS)])
        start += angle
    draw.ellipse((cx - radius // 2, cy - radius // 2, cx + radius // 2, cy + radius // 2), fill="white")
    ly = cy - radius + 10
    for i, (label, value) in enumerate(data[:8]):
        draw.rectangle((cx + radius + 40, ly + i * 44, cx + radius + 70, ly + i * 44 + 30), fill=COLORS[i % len(COLORS)])
        draw.text((cx + radius + 82, ly + i * 44 - 2), f"{trunc(label, 28)}: {fmt_int(value)}", fill="#1B1F23", font=font(22))


def draw_donut_metrics(draw, cx: int, cy: int, radius: int, title: str, data: list[tuple[str, float, float]], unit: str = ""):
    draw.text((cx - radius, cy - radius - 70), title, fill=INK, font=font(34, True))
    total = sum(partiya for _, partiya, _ in data) or 1
    start = -90
    for i, (_, partiya, _) in enumerate(data):
        angle = partiya / total * 360
        draw.pieslice((cx - radius, cy - radius, cx + radius, cy + radius), start, start + angle, fill=COLORS[i % len(COLORS)])
        start += angle
    draw.ellipse((cx - radius // 2, cy - radius // 2, cx + radius // 2, cy + radius // 2), fill="white")
    ly = cy - radius + 10
    for i, (label, partiya, qiymat) in enumerate(data[:8]):
        draw.rectangle((cx + radius + 40, ly + i * 50, cx + radius + 72, ly + i * 50 + 32), fill=COLORS[i % len(COLORS)])
        suffix = f" {unit}" if unit else ""
        text = f"{trunc(label, 24)}: {fmt_int(partiya)} | {fmt_num(qiymat)}{suffix}"
        draw.text((cx + radius + 84, ly + i * 50 - 2), text, fill="#1B1F23", font=font(22))


def draw_reason_bars(draw, x: int, y: int, title: str, data: list[tuple[str, float]], width: int):
    draw.text((x, y), title, fill=INK, font=font(34, True))
    y += 62
    max_v = max([v for _, v in data] or [1])
    for i, (label, value) in enumerate(data[:8]):
        yy = y + i * 56
        bar_w = int((value / max_v) * (width - 520))
        draw.text((x, yy + 4), trunc(label, 44), fill="#1B1F23", font=font(21))
        draw.rectangle((x + 540, yy + 8, x + 540 + bar_w, yy + 40), fill=COLORS[i % len(COLORS)])
        draw.text((x + 540 + bar_w + 12, yy + 5), fmt_int(value), fill="#1B1F23", font=font(22, True))


def build(source: Path, output_png: Path, output_pdf: Path, report_date: datetime, deposit_path: Path | None = None):
    df = core.read_source(source)
    df["_regime_code"] = df[core.SRC["regime"]].map(core.regime_code)
    deposits, deposit_date, deposit_total = core.read_deposit(deposit_path)
    expired = core.with_expiry(df, report_date)
    expired = expired[expired["_expired"]].copy()

    pages: list[Path] = []
    output_png.parent.mkdir(parents=True, exist_ok=True)
    base_name = output_png.name[:-4] if output_png.name.lower().endswith(".png") else output_png.name

    title = f"IM70-74 tahliliy ma'lumotlar | {report_date:%d.%m.%Y} holatiga"
    subtitle = f"Depozit fayli: {deposit_date:%d.%m.%Y} holatiga" if deposit_date else ""
    age_rows = []
    age_total = []
    for label, cutoff in [
        ("3 oy+", report_date - core.pd.DateOffset(months=3)),
        ("6 oy+", report_date - core.pd.DateOffset(months=6)),
        ("1 yil+", report_date - core.pd.DateOffset(years=1)),
    ]:
        d = df[df["_date"].le(cutoff)].copy()
        partiya = int(d["_partiya"].sum())
        qiymat = float(d["_value_usd_k"].sum())
        vazn = float(d["_weight_tn"].sum())
        tolov = float(d["_pay_total_mln"].sum())
        age_total.append((label, partiya))
        age_rows.append([label, partiya, fmt_num(vazn), fmt_num(qiymat), fmt_num(tolov)])

    img, draw = new_page(title, subtitle)
    draw_kpis(draw, [
        ("Partiya", fmt_int(df["_partiya"].sum())),
        ("Qiymat, ming $", fmt_num(float(df["_value_usd_k"].sum()))),
        ("Kutilayotgan, mln so'm", fmt_num(float(df["_pay_total_mln"].sum()))),
        ("Depozit, mln so'm", fmt_num(float(deposit_total))),
        ("Muddati o'tgan", fmt_int(expired["_partiya"].sum())),
    ])
    goods = df.groupby("_tnved_name").agg(partiya=("_partiya", "sum"), qiymat=("_value_usd_k", "sum")).sort_values("qiymat", ascending=False).head(10).reset_index()
    goods_rows = [[i + 1, trunc(r._tnved_name, 38), int(r.partiya), fmt_num(r.qiymat)] for i, r in goods.iterrows()]
    draw_table(draw, 60, 340, "TOP 10 tovar guruhlari", ["N", "Tovar guruhi", "Partiya", "Qiymat"], goods_rows, [50, 450, 105, 150])
    companies_all = df.groupby(core.SRC["stir"]).agg(company=(core.SRC["company"], "first"), qiymat=("_value_usd_k", "sum"), tolov=("_pay_total_mln", "sum"), partiya=("_partiya", "sum")).reset_index()
    companies = companies_all.sort_values("qiymat", ascending=False).head(20).reset_index(drop=True)
    company_rows = []
    for i, r in companies.iterrows():
        stir = str(r[core.SRC["stir"]])
        company_rows.append([i + 1, trunc(r.company, 28), stir, int(r.partiya), fmt_num(r.qiymat), fmt_num(r.tolov), fmt_num(float(deposits.get(stir, 0.0)))])
    draw_table(draw, 60, 1030, "TOP 20 korxona (qiymat)", ["N", "Korxona", "STIR", "Partiya", "Qiymat", "To'lov", "Depozit"], company_rows, [50, 330, 125, 90, 135, 135, 135], row_h=42)
    regime = df.groupby("_regime_code").agg(partiya=("_partiya", "sum"), qiymat=("_value_usd_k", "sum")).sort_values("partiya", ascending=False).reset_index()
    draw_donut_metrics(draw, 2050, 650, 250, "Rejimlar kesimida partiya va qiymat", [(r["_regime_code"], float(r.partiya), float(r.qiymat)) for _, r in regime.iterrows()], unit="ming $")
    deposit_rank = companies_all.copy()
    deposit_rank["deposit"] = deposit_rank[core.SRC["stir"]].map(lambda s: float(deposits.get(str(s), 0.0)))
    deposit_rank = deposit_rank.sort_values("deposit", ascending=False).head(20)
    deposit_rows = []
    for i, r in deposit_rank.iterrows():
        stir = str(r[core.SRC["stir"]])
        deposit_rows.append([len(deposit_rows) + 1, trunc(r.company, 28), stir, int(r.partiya), fmt_num(r.qiymat), fmt_num(r.tolov), fmt_num(float(deposits.get(stir, 0.0)))])
    draw_table(draw, 1650, 1030, "TOP 20 korxonalar (depozit mablag'lari bo'yicha)", ["N", "Korxona", "STIR", "Partiya", "Qiymat", "To'lov", "Depozit"], deposit_rows, [50, 330, 125, 90, 135, 135, 135], row_h=42)
    page1 = output_png
    img.save(page1)
    pages.append(page1)

    img, draw = new_page(f"Muddati o'tgan tovarlar | {report_date:%d.%m.%Y} holatiga")
    exp_comp = expired.groupby([core.SRC["stir"], "_regime_code", "_post_group"], dropna=False).agg(company=(core.SRC["company"], "first"), partiya=("_partiya", "sum"), qiymat=("_value_usd_k", "sum"), tolov=("_pay_total_mln", "sum"), expiry=("_expiry", "min")).sort_values(["partiya", "qiymat"], ascending=False).reset_index()
    exp_comp_rows = []
    for i, r in exp_comp.iterrows():
        stir = str(r[core.SRC["stir"]])
        overdue_days = max(0, int((report_date - r.expiry).days)) if core.pd.notna(r.expiry) else 0
        exp_comp_rows.append([i + 1, trunc(r.company, 27), stir, r["_regime_code"], trunc(r["_post_group"], 22), int(r.partiya), overdue_days, fmt_num(r.qiymat), fmt_num(r.tolov)])
    draw_table(draw, 60, 180, "Muddati o'tgan korxonalar", ["N", "Korxona", "STIR", "Rejim", "Nazorat posti", "Partiya", "Kun", "Qiymat", "To'lov"], exp_comp_rows, [45, 285, 120, 75, 215, 85, 75, 125, 125], row_h=42)
    exp_goods = expired.groupby("_tnved_name").agg(partiya=("_partiya", "sum"), qiymat=("_value_usd_k", "sum")).sort_values(["partiya", "qiymat"], ascending=False).head(10).reset_index()
    exp_goods_rows = [[i + 1, trunc(r._tnved_name, 48), int(r.partiya), fmt_num(r.qiymat)] for i, r in exp_goods.iterrows()]
    draw_table(draw, 1600, 660, "Muddati o'tgan tovar guruhlari TOP 10", ["N", "Tovar guruhi", "Partiya", "Qiymat"], exp_goods_rows, [50, 540, 105, 150])
    exp_post = expired.groupby("_post_group")["_partiya"].sum().sort_values(ascending=False)
    draw_barh(draw, 1600, 210, "Muddati o'tgan partiyalar postlar kesimida", list(exp_post.items()), 1350)
    exp_reason = expired.groupby(core.SRC["reason"])["_partiya"].sum().sort_values(ascending=False).head(8)
    draw_reason_bars(draw, 1600, 1510, "Muddati o'tgan saqlanish sabablari", [(core.clean_reason_display(k), v) for k, v in exp_reason.items()], 1350)
    page2 = output_png.with_name(base_name + "_2_muddati_otgan.png")
    img.save(page2)
    pages.append(page2)

    img, draw = new_page("Davlatlar va qo'shimcha tahlillar", f"{report_date:%d.%m.%Y} holatiga")
    countries = df.groupby(core.SRC["country"]).agg(partiya=("_partiya", "sum"), qiymat=("_value_usd_k", "sum"), tolov=("_pay_total_mln", "sum"), korxona=(core.SRC["stir"], "nunique")).sort_values("qiymat", ascending=False).head(12).reset_index()
    country_rows = [[i + 1, trunc(r[core.SRC["country"]], 30), int(r.partiya), int(r.korxona), fmt_num(r.qiymat), fmt_num(r.tolov)] for i, r in countries.iterrows()]
    draw_table(draw, 60, 180, "Yuk jo'natuvchi davlatlar TOP 12", ["N", "Davlat", "Partiya", "Korxona", "Qiymat", "To'lov"], country_rows, [50, 320, 100, 100, 155, 155])
    reasons = df.groupby(core.SRC["reason"]).agg(partiya=("_partiya", "sum"), qiymat=("_value_usd_k", "sum")).sort_values("partiya", ascending=False).head(10).reset_index()
    reason_rows = [[i + 1, trunc(core.clean_reason_display(r[core.SRC["reason"]]), 55), int(r.partiya), fmt_num(r.qiymat)] for i, r in reasons.iterrows()]
    draw_table(draw, 60, 1000, "Saqlanish sabablari TOP 10", ["N", "Sabab", "Partiya", "Qiymat"], reason_rows, [50, 560, 105, 150])
    draw_barh(draw, 1600, 190, "Davlatlar kesimida qiymat", [(r[core.SRC["country"]], r.qiymat) for _, r in countries.head(8).iterrows()], 1350)
    draw_table(draw, 1600, 900, "Muddatlar kesimida", ["Muddat", "Partiya", "Vazn", "Qiymat", "To'lov"], age_rows, [150, 130, 170, 170, 170])
    draw_barh(draw, 1600, 1260, "Muddatlar diagrammasi", age_total, 1250, label_w=190)
    page3 = output_png.with_name(base_name + "_3_davlatlar.png")
    img.save(page3)
    pages.append(page3)

    own = df[df[core.SRC["warehouse"]].map(core.clean).eq("")].copy()
    img, draw = new_page("O'z ombor bo'yicha tahlil", f"{report_date:%d.%m.%Y} holatiga")
    own_comp = own.groupby(core.SRC["stir"], dropna=False).agg(
        company=(core.SRC["company"], "first"),
        partiya=("_partiya", "sum"),
        vazn=("_weight_tn", "sum"),
        qiymat=("_value_usd_k", "sum"),
        tolov=("_pay_total_mln", "sum"),
    ).sort_values(["partiya", "qiymat"], ascending=False).reset_index()
    own_rows = []
    for i, r in own_comp.iterrows():
        own_rows.append([i + 1, trunc(r.company, 32), str(r[core.SRC["stir"]]), int(r.partiya), fmt_num(r.vazn), fmt_num(r.qiymat), fmt_num(r.tolov)])
    draw_table(draw, 60, 180, "O'z ombor korxonalar kesimida", ["N", "Korxona", "STIR", "Partiya", "Vazn", "Qiymat", "To'lov"], own_rows, [45, 330, 120, 90, 130, 130, 130], row_h=42)
    own_3m = own[own["_date"].le(report_date - core.pd.DateOffset(months=3))].copy()
    own_3m_comp = own_3m.groupby(core.SRC["stir"], dropna=False).agg(
        company=(core.SRC["company"], "first"),
        partiya=("_partiya", "sum"),
        vazn=("_weight_tn", "sum"),
        qiymat=("_value_usd_k", "sum"),
        tolov=("_pay_total_mln", "sum"),
    ).sort_values(["partiya", "qiymat"], ascending=False).reset_index()
    own_3m_rows = []
    for i, r in own_3m_comp.iterrows():
        own_3m_rows.append([i + 1, trunc(r.company, 32), str(r[core.SRC["stir"]]), int(r.partiya), fmt_num(r.vazn), fmt_num(r.qiymat), fmt_num(r.tolov)])
    draw_table(draw, 1600, 1050, "O'z ombor 3 oy+ korxonalar kesimida", ["N", "Korxona", "STIR", "Partiya", "Vazn", "Qiymat", "To'lov"], own_3m_rows, [45, 330, 120, 90, 130, 130, 130], row_h=42)
    own_age = []
    for label, cutoff in [
        ("3 oy+", report_date - core.pd.DateOffset(months=3)),
        ("6 oy+", report_date - core.pd.DateOffset(months=6)),
        ("1 yil+", report_date - core.pd.DateOffset(years=1)),
    ]:
        d = own[own["_date"].le(cutoff)].copy()
        own_age.append((label, float(d["_partiya"].sum()), float(d["_value_usd_k"].sum())))
    draw_donut_metrics(draw, 2050, 630, 250, "O'z ombor muddatlari: partiya va qiymat", own_age)
    page5 = output_png.with_name(base_name + "_5_oz_ombor.png")
    img.save(page5)
    pages.append(page5)

    c = canvas.Canvas(str(output_pdf), pagesize=landscape(A4))
    page_w, page_h = landscape(A4)
    for page in pages:
        c.drawImage(str(page), 18, 18, width=page_w - 36, height=page_h - 36, preserveAspectRatio=True, anchor="c")
        c.showPage()
    c.save()


if __name__ == "__main__":
    dep = Path(sys.argv[5]) if len(sys.argv) > 5 and sys.argv[5] else None
    build(Path(sys.argv[1]), Path(sys.argv[2]), Path(sys.argv[3]), datetime.strptime(sys.argv[4], "%d.%m.%Y"), dep)
