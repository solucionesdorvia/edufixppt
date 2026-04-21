#!/usr/bin/env python3
"""Genera EduFix_presentacion_datos.pptx a partir de respuestas.csv (solo lectura)."""
from __future__ import annotations

import csv
from collections import Counter
from datetime import date
from pathlib import Path

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.util import Emu, Inches, Pt

BASE = Path(__file__).resolve().parent
CSV_PATH = BASE / "respuestas.csv"
OUT_PATH = BASE / "EduFix_presentacion_datos.pptx"

# Paleta slide EduFix
C_NAVY = RGBColor(0x0A, 0x1C, 0x36)
C_ORANGE = RGBColor(0xD9, 0x53, 0x1E)
C_TEAL = RGBColor(0x1F, 0x8B, 0x63)
C_BLUE = RGBColor(0x1F, 0x57, 0x8B)
C_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
C_MUTED = RGBColor(0xB8, 0xC5, 0xD9)
C_CARD = RGBColor(0x14, 0x28, 0x48)


def load_rows(path: Path) -> list[dict]:
    with open(path, newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))


def mean(nums: list[float]) -> float:
    return sum(nums) / len(nums) if nums else 0.0


def pct(part: int, total: int) -> float:
    return 100.0 * part / total if total else 0.0


def fmt_num(x: float, dec: int = 1) -> str:
    s = f"{x:.{dec}f}".replace(".", ",")
    return s


def add_deck_chrome(slide, prs: Presentation) -> None:
    """Barra lateral naranja + franja inferior tricolor."""
    h = prs.slide_height
    w = prs.slide_width
    stripe = Emu(int(w * 0.012))
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, stripe, h)
    bar.fill.solid()
    bar.fill.fore_color.rgb = C_ORANGE
    bar.line.fill.background()

    fh = Emu(int(h * 0.012))
    third = w // 3
    for i, col in enumerate((C_ORANGE, C_TEAL, C_BLUE)):
        seg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, third * i, h - fh, third, fh)
        seg.fill.solid()
        seg.fill.fore_color.rgb = col
        seg.line.fill.background()


def textbox(
    slide,
    left,
    top,
    width,
    height,
    text: str,
    *,
    size_pt: float = 14,
    bold: bool = False,
    color: RGBColor = C_WHITE,
    align=PP_ALIGN.LEFT,
    font_name: str = "Calibri",
) -> None:
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.vertical_anchor = MSO_ANCHOR.TOP
    p = tf.paragraphs[0]
    p.text = text
    p.alignment = align
    p.font.size = Pt(size_pt)
    p.font.bold = bold
    p.font.color.rgb = color
    p.font.name = font_name


def slide_dark(prs: Presentation):
    layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(layout)
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = C_NAVY
    return slide


def add_finding_cards(slide, prs, findings: list[tuple[str, str, RGBColor]]) -> None:
    """Grid 2×3 de hallazgos."""
    w = prs.slide_width
    h = prs.slide_height
    pad = Emu(int(w * 0.055))
    top0 = Emu(int(h * 0.22))
    gap_x = Emu(int(w * 0.028))
    gap_y = Emu(int(h * 0.04))
    cell_w = (w - 2 * pad - gap_x) // 2
    cell_h = Emu(int(h * 0.2))

    for i, (stat, desc, accent) in enumerate(findings):
        row, col = divmod(i, 2)
        left = pad + col * (cell_w + gap_x)
        top = top0 + row * (cell_h + gap_y)
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, cell_w, cell_h)
        card.adjustments[0] = 0.06
        card.fill.solid()
        card.fill.fore_color.rgb = C_CARD
        card.line.color.rgb = RGBColor(0x30, 0x45, 0x66)
        card.line.width = Pt(0.5)

        accent_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, Emu(int(w * 0.01)), cell_h)
        accent_bar.fill.solid()
        accent_bar.fill.fore_color.rgb = accent
        accent_bar.line.fill.background()

        textbox(
            slide,
            left + Emu(int(w * 0.035)),
            top + Emu(int(h * 0.018)),
            cell_w - Emu(int(w * 0.045)),
            Emu(int(cell_h * 0.42)),
            stat,
            size_pt=26,
            bold=True,
            color=C_WHITE,
        )
        textbox(
            slide,
            left + Emu(int(w * 0.035)),
            top + Emu(int(cell_h * 0.48)),
            cell_w - Emu(int(w * 0.045)),
            Emu(int(cell_h * 0.45)),
            desc,
            size_pt=11,
            bold=False,
            color=C_MUTED,
        )


def build_presentation(rows: list[dict]) -> Presentation:
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    n = len(rows)
    if n == 0:
        raise ValueError("No hay filas en respuestas.csv")

    def pred_count(fn) -> int:
        return sum(1 for r in rows if fn(r))

    pct_no_supo = pct(pred_count(lambda r: r.get("detectaste_problema_y_no_supiste_a_quien_reportar") == "Si"), n)
    pct_ignora = pct(pred_count(lambda r: r.get("ante_desperfecto_que_sueles_hacer") == "Ignorarlo"), n)
    pct_no_canal = pct(pred_count(lambda r: r.get("tiempo_reportar_canales_oficiales") == "No se como hacerlo"), n)
    avg_impacto = mean([float(r["afecta_entorno_descuidado_1_10"]) for r in rows])
    pct_1a3 = pct(pred_count(lambda r: r.get("veces_por_mes_notas_desperfecto") == "1 a 3 veces"), n)
    pct_edufix = pct(pred_count(lambda r: r.get("dispuesto_foto_si_notificacion_al_cerrar") == "Si"), n)
    avg_util = mean([float(r["utilidad_app_foto_estado_1_10"]) for r in rows])

    roles = Counter(r.get("rol_principal", "") for r in rows)
    frust = Counter(r.get("frustracion_reportes_actual", "") for r in rows)

    hist_u = Counter()
    hist_a = Counter()
    for r in rows:
        try:
            hist_u[int(float(r["utilidad_app_foto_estado_1_10"]))] += 1
            hist_a[int(float(r["afecta_entorno_descuidado_1_10"]))] += 1
        except (ValueError, TypeError):
            pass

    today = date.today().strftime("%d/%m/%Y")

    # --- 1 Portada ---
    s = slide_dark(prs)
    add_deck_chrome(s, prs)
    textbox(s, Inches(0.85), Inches(1.35), Inches(11.5), Inches(1.2), "EduFix", size_pt=54, bold=True, color=C_ORANGE)
    textbox(
        s,
        Inches(0.85),
        Inches(2.55),
        Inches(11.5),
        Inches(0.9),
        "Validación con datos",
        size_pt=28,
        bold=True,
        color=C_WHITE,
    )
    textbox(
        s,
        Inches(0.85),
        Inches(3.35),
        Inches(11),
        Inches(0.7),
        f"{n} respuestas analizadas · {today}",
        size_pt=16,
        color=C_MUTED,
    )
    textbox(
        s,
        Inches(0.85),
        Inches(4.0),
        Inches(11),
        Inches(1.2),
        "Proceso de reporte de desperfectos en el campus: percepción,\nfricciones y disposición a una herramienta rápida (EduFix).",
        size_pt=14,
        color=C_MUTED,
    )

    # --- 2 Problema validado ---
    s = slide_dark(prs)
    add_deck_chrome(s, prs)
    textbox(s, Inches(0.85), Inches(0.65), Inches(3), Inches(0.5), "01", size_pt=36, bold=True, color=C_ORANGE)
    textbox(s, Inches(0.85), Inches(1.15), Inches(11), Inches(0.8), "Problema validado", size_pt=36, bold=True)
    lead = (
        "Más de 400 encuestas confirman que el proceso actual es ineficiente e informal."
        if n >= 400
        else f"{n} respuestas en este set confirman un patrón claro: el flujo actual se percibe como ineficiente e informal."
    )
    body = (
        f"{lead}\n\n"
        "• Poca claridad sobre a quién escalar y cómo usar canales oficiales.\n"
        "• Comportamiento frecuente de ignorar el desperfecto.\n"
        "• Impacto notable en la experiencia en el entorno físico del campus."
    )
    textbox(s, Inches(0.85), Inches(2.0), Inches(11.2), Inches(4.2), body, size_pt=15, color=C_MUTED)

    # --- 3 Hallazgos 6 tarjetas ---
    s = slide_dark(prs)
    add_deck_chrome(s, prs)
    textbox(s, Inches(0.85), Inches(0.5), Inches(2.5), Inches(0.45), "03", size_pt=32, bold=True, color=C_ORANGE)
    textbox(s, Inches(0.85), Inches(0.95), Inches(11), Inches(0.65), "Hallazgos principales", size_pt=32, bold=True)
    textbox(
        s,
        Inches(0.85),
        Inches(1.45),
        Inches(11),
        Inches(0.4),
        "Validación — indicadores calculados sobre el dataset actual (sin alterar los datos fuente).",
        size_pt=12,
        color=C_MUTED,
    )
    findings = [
        (f"{fmt_num(pct_no_supo)}%", "No supo a quién reportar", C_ORANGE),
        (f"{fmt_num(pct_ignora)}%", "Ignoró el incidente directamente", RGBColor(0xE8, 0x5D, 0x5D)),
        (f"{fmt_num(pct_no_canal)}%", "No sabe cómo usar el canal oficial", RGBColor(0x4A, 0x9F, 0xE8)),
        (f"{fmt_num(avg_impacto)}/10", "Impacto en la experiencia académica", C_TEAL),
        (f"{fmt_num(pct_1a3)}%", "Detecta 1–3 desperfectos por mes", C_BLUE),
        (f"{fmt_num(pct_edufix)}%", "Usaría EduFix si tardara 30 segundos", RGBColor(0x4E, 0xD4, 0x9E)),
    ]
    add_finding_cards(s, prs, findings)

    # --- 4 Rol ---
    s = slide_dark(prs)
    add_deck_chrome(s, prs)
    textbox(s, Inches(0.85), Inches(0.55), Inches(11), Inches(0.7), "Composición por rol", size_pt=30, bold=True)
    chart_data = CategoryChartData()
    cats = [k for k, _ in roles.most_common()]
    chart_data.categories = cats
    chart_data.add_series("Respuestas", tuple(roles[c] for c in cats))
    x, y, cx, cy = Inches(0.9), Inches(1.45), Inches(11.4), Inches(5.2)
    ch = s.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, x, y, cx, cy, chart_data).chart
    ch.has_legend = False
    plot = ch.plots[0]
    plot.has_data_labels = True
    ser = ch.series[0]
    ser.format.fill.solid()
    ser.format.fill.fore_color.rgb = C_ORANGE

    # --- 5 Frustración ---
    s = slide_dark(prs)
    add_deck_chrome(s, prs)
    textbox(s, Inches(0.85), Inches(0.55), Inches(11), Inches(0.7), "Mayor frustración con los reportes", size_pt=28, bold=True)
    chart_data = CategoryChartData()
    cats = [k for k, _ in frust.most_common()]
    chart_data.categories = cats
    chart_data.add_series("Menciones", tuple(frust[c] for c in cats))
    ch = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.9), Inches(1.45), Inches(11.4), Inches(5.2), chart_data).chart
    ch.has_legend = False
    ch.plots[0].has_data_labels = True
    ch.series[0].format.fill.solid()
    ch.series[0].format.fill.fore_color.rgb = C_TEAL

    # --- 6 Utilidad (1–10) ---
    s = slide_dark(prs)
    add_deck_chrome(s, prs)
    textbox(
        s,
        Inches(0.85),
        Inches(0.55),
        Inches(11),
        Inches(0.7),
        f"Utilidad percibida de la app (media {fmt_num(avg_util)})",
        size_pt=28,
        bold=True,
    )
    chart_data = CategoryChartData()
    chart_data.categories = [str(i) for i in range(1, 11)]
    chart_data.add_series("Respuestas", tuple(hist_u.get(i, 0) for i in range(1, 11)))
    ch = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.9), Inches(1.45), Inches(11.4), Inches(5.2), chart_data).chart
    ch.has_legend = False
    ch.plots[0].has_data_labels = True
    ch.series[0].format.fill.solid()
    ch.series[0].format.fill.fore_color.rgb = C_ORANGE

    # --- 7 Impacto entorno ---
    s = slide_dark(prs)
    add_deck_chrome(s, prs)
    textbox(
        s,
        Inches(0.85),
        Inches(0.55),
        Inches(11),
        Inches(0.7),
        f"Impacto del entorno descuidado (media {fmt_num(avg_impacto)})",
        size_pt=28,
        bold=True,
    )
    chart_data = CategoryChartData()
    chart_data.categories = [str(i) for i in range(1, 11)]
    chart_data.add_series("Respuestas", tuple(hist_a.get(i, 0) for i in range(1, 11)))
    ch = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.9), Inches(1.45), Inches(11.4), Inches(5.2), chart_data).chart
    ch.has_legend = False
    ch.plots[0].has_data_labels = True
    ch.series[0].format.fill.solid()
    ch.series[0].format.fill.fore_color.rgb = C_TEAL

    # --- 8 Síntesis + próximo paso ---
    s = slide_dark(prs)
    add_deck_chrome(s, prs)
    textbox(s, Inches(0.85), Inches(0.65), Inches(11), Inches(0.8), "Síntesis", size_pt=34, bold=True)
    syn = (
        f"• {fmt_num(pct_no_supo)}% no supo a quién reportar.\n"
        f"• {fmt_num(pct_ignora)}% suele ignorar el desperfecto.\n"
        f"• {fmt_num(pct_no_canal)}% no sabe usar el canal oficial.\n"
        f"• Impacto medio en experiencia: {fmt_num(avg_impacto)}/10.\n"
        f"• {fmt_num(pct_edufix)}% usaría EduFix si el reporte tomara ~30 s.\n\n"
        "EduFix concentra foto + contexto + notificación para cerrar el ciclo "
        "con trazabilidad y menor fricción para alumnos, docentes y mantenimiento."
    )
    textbox(s, Inches(0.85), Inches(1.45), Inches(11.2), Inches(5.0), syn, size_pt=15, color=C_MUTED)

    # --- 9 Cierre ---
    s = slide_dark(prs)
    add_deck_chrome(s, prs)
    textbox(s, Inches(0.85), Inches(2.2), Inches(11.5), Inches(1.0), "Gracias", size_pt=44, bold=True, color=C_WHITE)
    textbox(
        s,
        Inches(0.85),
        Inches(3.35),
        Inches(11),
        Inches(1.5),
        "EduFix — reportes claros, rápidos y accionables.\n\n"
        f"Dataset: {n} respuestas · Generado {today}",
        size_pt=16,
        color=C_MUTED,
        align=PP_ALIGN.CENTER,
    )

    return prs


def main():
    if not CSV_PATH.is_file():
        raise SystemExit(f"No se encontró {CSV_PATH}")
    rows = load_rows(CSV_PATH)
    prs = build_presentation(rows)
    prs.save(OUT_PATH)
    print(f"Guardado: {OUT_PATH} ({len(rows)} filas → {len(prs.slides)} diapositivas)")


if __name__ == "__main__":
    main()
