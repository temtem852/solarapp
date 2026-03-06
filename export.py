# =========================================================
# export.py — IEEE Engineering Paper PDF Export
# =========================================================

import io
from datetime import datetime

from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics


# =========================================================
# FONT REGISTRATION (once at import time)
# =========================================================
if "TH" not in pdfmetrics.getRegisteredFontNames():
    pdfmetrics.registerFont(TTFont("TH",   "THSarabunNew.ttf"))
    pdfmetrics.registerFont(TTFont("TH-B", "THSarabunNew-Bold.ttf"))


# =========================================================
# BUILD IEEE PAPER PDF
# =========================================================
def build_ieee_paper(
    dc_capacity,
    dc_ac_ratio,
    H_sun,
    PR,
    area,
    E_day,
    E_est_day,
    panels_per_string,
    strings_used,
    simple_payback,
    irr_val,
    project_life,
    ai_result="No AI result available.",
) -> bytes:
    """
    Builds and returns the IEEE paper as PDF bytes.
    All variables passed explicitly — no st.session_state access.
    """
    buffer = io.BytesIO()

    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=40,
        leftMargin=40,
        topMargin=40,
        bottomMargin=40,
    )

    styles = getSampleStyleSheet()

    styles.add(ParagraphStyle(
        name="IEEE_Title",
        fontName="TH-B",
        fontSize=18,
        alignment=TA_CENTER,
        spaceAfter=14,
    ))
    styles.add(ParagraphStyle(
        name="IEEE_Section",
        fontName="TH-B",
        fontSize=14,
        spaceBefore=12,
        spaceAfter=6,
    ))
    styles.add(ParagraphStyle(
        name="IEEE_Body",
        fontName="TH",
        fontSize=12,
        leading=16,
        alignment=TA_JUSTIFY,
    ))

    story = []

    # --------------------------------------------------
    # TITLE
    # --------------------------------------------------
    story.append(Paragraph(
        "Design and Optimization of Rooftop Solar PV System "
        "with AI-Assisted Component Selection",
        styles["IEEE_Title"],
    ))
    story.append(Spacer(1, 8))

    # --------------------------------------------------
    # ABSTRACT
    # --------------------------------------------------
    story.append(Paragraph("Abstract", styles["IEEE_Section"]))
    story.append(Paragraph(
        f"This paper presents the engineering design and economic evaluation of a rooftop "
        f"solar photovoltaic (PV) system sized at {dc_capacity:.2f} kWp. "
        f"The system is designed based on peak sun hours ({H_sun:.2f} h/day), "
        f"performance ratio ({PR:.2f}), and rooftop constraints ({area:.1f} m²). "
        f"A deterministic calculation approach is applied for system sizing, "
        f"while a large language model (LLM) is utilized for database-assisted "
        f"component selection. Financial feasibility including IRR and payback "
        f"period is evaluated to determine project viability.",
        styles["IEEE_Body"],
    ))

    # --------------------------------------------------
    # I. INTRODUCTION
    # --------------------------------------------------
    story.append(Paragraph("I. INTRODUCTION", styles["IEEE_Section"]))
    story.append(Paragraph(
        "Rooftop solar photovoltaic systems are increasingly adopted "
        "for residential and commercial applications. "
        "Proper engineering design is essential to ensure electrical safety, "
        "performance optimization, and financial feasibility.",
        styles["IEEE_Body"],
    ))

    # --------------------------------------------------
    # II. SYSTEM DESIGN METHODOLOGY
    # --------------------------------------------------
    story.append(Paragraph("II. SYSTEM DESIGN METHODOLOGY", styles["IEEE_Section"]))
    story.append(Paragraph(
        f"The required PV capacity is calculated using the daily energy demand "
        f"({E_day:.2f} kWh/day), peak sun hours, and performance ratio. "
        f"The DC/AC ratio is maintained at {dc_ac_ratio:.2f} to ensure inverter "
        f"loading optimization and clipping control.",
        styles["IEEE_Body"],
    ))

    # --------------------------------------------------
    # III. ENGINEERING RESULTS
    # --------------------------------------------------
    story.append(Paragraph("III. ENGINEERING RESULTS", styles["IEEE_Section"]))

    results_table = Table([
        ["Parameter",          "Value"],
        ["PV Capacity (kWp)",  f"{dc_capacity:.2f}"],
        ["DC/AC Ratio",        f"{dc_ac_ratio:.2f}"],
        ["Panels per String",  str(panels_per_string)],
        ["Number of Strings",  str(strings_used)],
    ], colWidths=[230, 230])

    results_table.setStyle(TableStyle([
        ("FONT",       (0, 0), (-1, -1), "TH"),
        ("GRID",       (0, 0), (-1, -1), 0.5, colors.grey),
        ("BACKGROUND", (0, 0), (-1,  0), colors.lightgrey),
    ]))
    story.append(results_table)

    # --------------------------------------------------
    # IV. AI-ASSISTED COMPONENT SELECTION
    # --------------------------------------------------
    story.append(Paragraph("IV. AI-ASSISTED COMPONENT SELECTION", styles["IEEE_Section"]))
    ai_text = ai_result if ai_result else "AI recommendation not generated. Run AI Analysis to populate this section."
    story.append(Paragraph(
        ai_text.replace("\n", "<br/>"),
        styles["IEEE_Body"],
    ))

    # --------------------------------------------------
    # V. FINANCIAL ANALYSIS
    # --------------------------------------------------
    story.append(Paragraph("V. FINANCIAL ANALYSIS", styles["IEEE_Section"]))

    payback_str = (
        f"{simple_payback} years"
        if simple_payback
        else f">{project_life} years"
    )

    story.append(Paragraph(
        f"The financial evaluation indicates a simple payback period of "
        f"{payback_str} and an internal rate of return (IRR) "
        f"of {irr_val * 100:.2f}%.",
        styles["IEEE_Body"],
    ))

    # --------------------------------------------------
    # VI. CONCLUSION
    # --------------------------------------------------
    story.append(Paragraph("VI. CONCLUSION", styles["IEEE_Section"]))
    story.append(Paragraph(
        "The designed solar PV system satisfies engineering constraints "
        "and demonstrates economic feasibility. "
        "The integration of deterministic calculation with AI-assisted "
        "database selection enhances engineering workflow efficiency "
        "while maintaining technical reliability.",
        styles["IEEE_Body"],
    ))

    doc.build(story)
    return buffer.getvalue()
