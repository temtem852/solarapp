# =========================================================
# export.py — Solar Rooftop Engineering Report PDF
# Layout: สูตรออกแบบจำนวนแผงโซลาร์และการจัด String
# =========================================================

import io
from datetime import datetime

from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
)
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.units import mm
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics

# =========================================================
# FONT
# =========================================================
if "TH" not in pdfmetrics.getRegisteredFontNames():
    pdfmetrics.registerFont(TTFont("TH",   "THSarabunNew.ttf"))
    pdfmetrics.registerFont(TTFont("TH-B", "THSarabunNew-Bold.ttf"))

# =========================================================
# COLORS
# =========================================================
C_HEADER  = colors.HexColor("#1F5C8B")
C_SUBHDR  = colors.HexColor("#2E75B6")
C_ORANGE  = colors.HexColor("#F4B942")
C_YELLOW  = colors.HexColor("#FFF2CC")
C_GREEN   = colors.HexColor("#E2EFDA")
C_GRNTXT  = colors.HexColor("#375623")
C_REDTXT  = colors.HexColor("#C00000")
C_LGREY   = colors.HexColor("#F2F2F2")
C_BORDER  = colors.HexColor("#9DC3E6")
C_WHITE   = colors.white


def _s(v, fmt=None):
    if v is None:
        return "-"
    try:
        return fmt.format(v) if fmt else str(v)
    except Exception:
        return str(v)


def _ok_fail(cond):
    return "OK" if cond else "FAIL"


def _base_ts(extra=None):
    cmds = [
        ("FONTNAME",      (0,0), (-1,-1), "TH"),
        ("FONTSIZE",      (0,0), (-1,-1), 9),
        ("GRID",          (0,0), (-1,-1), 0.4, C_BORDER),
        ("VALIGN",        (0,0), (-1,-1), "MIDDLE"),
        ("TOPPADDING",    (0,0), (-1,-1), 3),
        ("BOTTOMPADDING", (0,0), (-1,-1), 3),
        ("LEFTPADDING",   (0,0), (-1,-1), 5),
        ("RIGHTPADDING",  (0,0), (-1,-1), 5),
    ]
    if extra:
        cmds += extra
    return TableStyle(cmds)


# =========================================================
# MAIN
# =========================================================
def build_ieee_paper(
    dc_capacity,
    dc_ac_ratio,
    H_sun,
    PR,
    area,
    E_day,
    E_est_day,
    d=None,
    inv_ac=None,
    inv_v=None,
    inv_i=None,
    inv_pv=None,
    v_mppt_min=None,
    v_mppt_max=None,
    mppt_count=None,
    Pm=None, Vmp=None, Voc=None, Imp=None, Isc=None,
    CAPEX=None,
    simple_payback=None,
    discounted_payback=None,
    npv=None,
    irr_val=None,
    project_life=25,
    tariff_self=4.0,
    tariff_export=0.0,
    E_year_1=None,
    panels_per_string=None,
    strings_used=None,
    ai_result=None,
) -> bytes:

    if d is None:
        d = {}

    pps       = panels_per_string or d.get("panels_per_string", "-")
    n_str     = strings_used      or d.get("strings_used", "-")
    n_req     = d.get("panels_required", "-")
    n_min_mppt= d.get("n_min_mppt", 0)
    Voc_str   = d.get("Voc_string", None)
    Vmp_str   = d.get("Vmp_string", None)
    I_str     = d.get("I_string",   None)
    mppt_alloc= d.get("mppt_allocation", [])
    dc_cap    = d.get("dc_capacity", dc_capacity) or dc_capacity
    dc_ratio  = d.get("dc_ac_ratio", dc_ac_ratio)  or dc_ac_ratio

    # checks
    voc_ok  = bool(Voc_str and inv_v  and Voc_str <= inv_v)
    isc_ok  = bool(I_str   and inv_i  and I_str   <= inv_i * 1.25)
    pps_ok  = bool(isinstance(pps, (int,float)) and isinstance(n_min_mppt, (int,float)) and pps >= n_min_mppt)
    vmpp_ok = bool(Vmp_str and v_mppt_min and Vmp_str >= v_mppt_min)
    dc_total_w = (dc_cap or 0) * 1000

    # CO2 (Thailand grid 0.4715 kgCO2/kWh)
    e_yr    = E_year_1 or (E_est_day * 365)
    co2_yr  = e_yr * 0.4715 / 1000
    co2_25  = co2_yr * 25

    # ---- document ----
    buffer = io.BytesIO()
    W, H = A4
    M = 14 * mm
    doc = SimpleDocTemplate(buffer, pagesize=A4,
                            rightMargin=M, leftMargin=M,
                            topMargin=M, bottomMargin=M)
    CW = W - 2 * M

    # ---- paragraph styles ----
    def ps(name, **kw):
        defaults = dict(fontName="TH", fontSize=9, leading=12)
        defaults.update(kw)
        return ParagraphStyle(name, **defaults)

    S = {
        "WHDR":  ps("WHDR",  fontName="TH-B", fontSize=11, alignment=TA_CENTER, textColor=C_WHITE),
        "WBOLD": ps("WBOLD", fontName="TH-B", fontSize=9,  alignment=TA_CENTER, textColor=C_WHITE),
        "LBL":   ps("LBL",   fontSize=8, leading=10),
        "VAL":   ps("VAL",   fontName="TH-B", fontSize=10, alignment=TA_CENTER),
        "UNIT":  ps("UNIT",  fontSize=8, alignment=TA_CENTER, textColor=colors.HexColor("#555")),
        "OK":    ps("OK",    fontName="TH-B", fontSize=9, alignment=TA_CENTER, textColor=C_GRNTXT),
        "FAIL":  ps("FAIL",  fontName="TH-B", fontSize=9, alignment=TA_CENTER, textColor=C_REDTXT),
        "BODY":  ps("BODY",  fontSize=9, leading=13),
        "TINY":  ps("TINY",  fontSize=7, textColor=colors.HexColor("#666")),
        "SEC":   ps("SEC",   fontName="TH-B", fontSize=10, textColor=C_HEADER, spaceBefore=5, spaceAfter=2),
    }

    def ok_p(cond):
        return Paragraph("OK" if cond else "FAIL", S["OK"] if cond else S["FAIL"])

    story = []
    half  = (CW - 4) / 2

    # =========================================================
    # 1. TITLE
    # =========================================================
    title = Table([[Paragraph(
        "สูตรออกแบบจำนวนแผงโซลาร์และการจัด String ให้ระบบมีประสิทธิภาพ",
        S["WHDR"]
    )]], colWidths=[CW])
    title.setStyle(_base_ts([
        ("BACKGROUND",    (0,0),(-1,-1), C_HEADER),
        ("TOPPADDING",    (0,0),(-1,-1), 7),
        ("BOTTOMPADDING", (0,0),(-1,-1), 7),
    ]))
    story.append(title)
    story.append(Spacer(1, 4))

    # =========================================================
    # 2. SPEC — Inverter | PV Module (2 columns)
    # =========================================================
    def spec_col(title_txt, rows_data, hdr_color):
        lw, vw, uw = half*0.60, half*0.27, half*0.13
        hdr = [Paragraph(title_txt, S["WHDR"]), Paragraph("", S["WHDR"]), Paragraph("", S["WHDR"])]
        rows = [hdr] + [
            [Paragraph(r[0], S["LBL"]),
             Paragraph(_s(r[1]), ParagraphStyle("_v", fontName="TH-B", fontSize=9, alignment=TA_CENTER)),
             Paragraph(r[2], S["UNIT"])]
            for r in rows_data
        ]
        t = Table(rows, colWidths=[lw, vw, uw])
        t.setStyle(_base_ts([
            ("BACKGROUND", (0,0), (-1,0), hdr_color),
            ("SPAN",       (0,0), (-1,0)),
            ("BACKGROUND", (0,1), (-1,-1), C_ORANGE),
            ("FONTNAME",   (0,0), (-1,0),  "TH-B"),
        ]))
        return t

    inv_data = [
        ("แรงดันเริ่มทำงาน (Start-up voltage)",                    v_mppt_min,         "V"),
        ("กำลังอินเวอร์เตอร์ (kW)",                                inv_ac/1000 if inv_ac else None, "kW"),
        ("แรงดันต่ำสุดอินเวอร์เตอร์ทำงาน (MPPT min)",              v_mppt_min,         "V"),
        ("แรงดันสูงสุดที่อินเวอร์เตอร์รับได้ (Max. input voltage)", inv_v,              "V"),
        ("Max. PV power",                                           inv_pv/1000 if inv_pv else None, "kW"),
        ("กระแสลัดวงจรสูงสุด (Max. short-circuit current)",         inv_i,              "A"),
    ]
    pv_data = [
        ("Maximum Power Voltage (Vmp)",   Vmp,  "V"),
        ("Open Circuit Voltage (Voc)",    Voc,  "V"),
        ("Maximum Power Current (Imp)",   Imp,  "A"),
        ("Short Circuit Current (Isc)",   Isc,  "A"),
        ("วัตต์แผง (Pm)",                 Pm,   "W"),
        ("Temperature Coefficient of Voc", "-0.260", "%/°C"),
    ]

    spec_tbl = Table([[spec_col("Inverter (อินเวอร์เตอร์)", inv_data, C_HEADER),
                       spec_col("PV Module (แผงโซลาร์) at STC", pv_data, C_SUBHDR)]],
                     colWidths=[half, half])
    spec_tbl.setStyle(TableStyle([
        ("LEFTPADDING",  (0,0),(-1,-1), 0),
        ("RIGHTPADDING", (0,0),(-1,-1), 2),
        ("TOPPADDING",   (0,0),(-1,-1), 0),
        ("BOTTOMPADDING",(0,0),(-1,-1), 0),
    ]))
    story.append(spec_tbl)
    story.append(Spacer(1, 5))

    # =========================================================
    # 3. SUMMARY TABLE + CHECK TABLE
    # =========================================================
    SW  = CW * 0.54
    CKW = CW * 0.44

    sum_rows = [
        [Paragraph("ตารางสรุป (ไม่ต้องแก้ไขตารางนี้)", S["WBOLD"]), "", ""],
        [Paragraph("จำนวนแผงโซลาร์ใน 1 String", S["LBL"]), Paragraph(_s(pps), S["VAL"]), Paragraph("แผง", S["UNIT"])],
        [Paragraph("Vmp รวมของแผงโซลาร์ (STC)",   S["LBL"]), Paragraph(_s(Vmp_str,"{:.2f}") if Vmp_str else "-", S["VAL"]), Paragraph("V", S["UNIT"])],
        [Paragraph("Voc รวมของแผงโซลาร์ (STC)",   S["LBL"]), Paragraph(_s(Voc_str,"{:.2f}") if Voc_str else "-", S["VAL"]), Paragraph("V", S["UNIT"])],
        [Paragraph("Isc ของแผงโซลาร์ (STC)",      S["LBL"]), Paragraph(_s(I_str, "{:.2f}") if I_str else "-",   S["VAL"]), Paragraph("A", S["UNIT"])],
        [Paragraph("W รวมของแผง",                  S["LBL"]), Paragraph(_s(dc_total_w,"{:.0f}"), S["VAL"]),                  Paragraph("W", S["UNIT"])],
    ]
    st_sum = Table(sum_rows, colWidths=[SW*0.64, SW*0.24, SW*0.12])
    st_sum.setStyle(_base_ts([
        ("BACKGROUND", (0,0), (-1,0), C_SUBHDR),
        ("SPAN",       (0,0), (-1,0)),
        ("BACKGROUND", (0,1), (-1,-1), C_YELLOW),
    ]))

    chk_rows = [
        [Paragraph("ตรวจสอบเงื่อนไข", S["WBOLD"]), Paragraph("สถานะ", S["WBOLD"])],
        [Paragraph("จำนวนแผงโซลาร์ใน 1 String", S["LBL"]), ok_p(pps_ok)],
        [Paragraph("Vmp ของแผงโซลาร์ (STC)",    S["LBL"]), ok_p(vmpp_ok)],
        [Paragraph("Voc ของแผงโซลาร์ (STC)",    S["LBL"]), ok_p(voc_ok)],
        [Paragraph("กระแสของแผงโซลาร์ (STC)",   S["LBL"]), ok_p(isc_ok)],
        [Paragraph("W รวมของแผง",                S["LBL"]), ok_p(True)],
    ]
    st_chk = Table(chk_rows, colWidths=[CKW*0.72, CKW*0.28])
    st_chk.setStyle(_base_ts([
        ("BACKGROUND", (0,0), (-1,0), C_SUBHDR),
        ("BACKGROUND", (0,1), (-1,-1), C_LGREY),
        ("BACKGROUND", (1,1), (1,-1), C_GREEN),
    ]))

    sc = Table([[st_sum, st_chk]], colWidths=[SW, CKW])
    sc.setStyle(TableStyle([
        ("LEFTPADDING",  (0,0),(-1,-1), 0),
        ("RIGHTPADDING", (0,0),(-1,-1), 3),
        ("TOPPADDING",   (0,0),(-1,-1), 0),
        ("BOTTOMPADDING",(0,0),(-1,-1), 0),
    ]))
    story.append(sc)
    story.append(Spacer(1, 5))

    # =========================================================
    # 4. MPPT ALLOCATION TABLE
    # =========================================================
    n_mppt   = int(mppt_count or len(mppt_alloc) or 4)
    n_mppt   = min(n_mppt, 4)
    alloc    = list(mppt_alloc) + [0]*n_mppt
    alloc    = alloc[:n_mppt]
    dc_ratio_color = C_GREEN if 1.10 <= (dc_ratio or 0) <= 1.25 else colors.HexColor("#FCE4D6")

    mppt_hdr = (["Inverter"] +
                [f"MPPT.{i+1}" for i in range(n_mppt)] +
                ["จำนวนแผง\nทั้งหมด (Module)", "วัตต์แผง\n(Wp/Module)",
                 "Inverter\n(W)", "วัตต์รวมของแผง\n(W)", "DC to AC\nRatio"])
    mppt_row = (["1"] + [str(v) for v in alloc] +
                [_s(n_req), _s(Pm), _s(inv_ac), _s(dc_total_w,"{:.0f}"),
                 _s(dc_ratio,"{:.2f}")])

    last = len(mppt_hdr) - 1
    base_cw = 30
    rest_cw = (CW - base_cw - n_mppt*32 - 5) / 5
    m_cw = [base_cw] + [32]*n_mppt + [rest_cw]*5

    mppt_t = Table([mppt_hdr, mppt_row], colWidths=m_cw)
    mppt_t.setStyle(_base_ts([
        ("BACKGROUND", (0,0), (-1,0), C_SUBHDR),
        ("FONTNAME",   (0,0), (-1,0), "TH-B"),
        ("TEXTCOLOR",  (0,0), (-1,0), C_WHITE),
        ("ALIGN",      (0,0), (-1,-1), "CENTER"),
        ("FONTSIZE",   (0,0), (-1,-1), 8),
        ("BACKGROUND", (0,1), (-1,1), C_LGREY),
        ("BACKGROUND", (last,1), (last,1), dc_ratio_color),
        ("FONTNAME",   (last,1), (last,1), "TH-B"),
    ]))

    story.append(Paragraph("ออกแบบ String", S["SEC"]))
    story.append(Paragraph(
        "ข้อมูลอินเวอร์เตอร์ตามใน Datasheet ของอินเวอร์เตอร์",
        ParagraphStyle("lnk", fontName="TH", fontSize=8, textColor=C_REDTXT, spaceAfter=2)
    ))
    story.append(mppt_t)
    story.append(Spacer(1, 4))

    # actual install verify rows
    act_hdr = ["จำนวนแผงโซลาร์ใน 1 String",
               "Vmp รวมของแผงโซลาร์ (STC)",
               "Voc รวมของแผงโซลาร์ (STC)",
               "Isc ของแผงโซลาร์ (STC)",
               "W รวมของแผง"]
    act_val = [f"{pps} แผง",
               f"{_s(Vmp_str,'{:.2f}')} V" if Vmp_str else "-",
               f"{_s(Voc_str,'{:.2f}')} V" if Voc_str else "-",
               f"{_s(I_str,'{:.2f}')} A"   if I_str   else "-",
               f"{dc_total_w:.0f} W"]
    act_chk = [ok_p(pps_ok), ok_p(vmpp_ok), ok_p(voc_ok), ok_p(isc_ok), ok_p(True)]

    act_cw  = [CW/5]*5
    act_t   = Table([act_hdr, act_val, act_chk], colWidths=act_cw)
    act_t.setStyle(_base_ts([
        ("BACKGROUND", (0,0), (-1,0), C_SUBHDR),
        ("FONTNAME",   (0,0), (-1,0), "TH-B"),
        ("TEXTCOLOR",  (0,0), (-1,0), C_WHITE),
        ("ALIGN",      (0,0), (-1,-1), "CENTER"),
        ("FONTSIZE",   (0,0), (-1,-1), 8),
        ("BACKGROUND", (0,1), (-1,1), C_YELLOW),
        ("BACKGROUND", (0,2), (-1,2), C_GREEN),
    ]))
    story.append(Paragraph(
        "ทดลองใส่แผง (เพื่อยืนยันในการใช้งานจริง)",
        ParagraphStyle("tl", fontName="TH-B", fontSize=9,
                       textColor=C_HEADER, spaceAfter=2)
    ))
    story.append(act_t)
    story.append(Spacer(1, 5))

    # =========================================================
    # 5. DC/AC RATIO BANNER
    # =========================================================
    if 1.10 <= (dc_ratio or 0) <= 1.25:
        dc_status = "✅ อยู่ในช่วงที่แนะนำ (1.10 – 1.25)  ระบบมีประสิทธิภาพดี"
    elif (dc_ratio or 0) > 1.25:
        dc_status = "⚠️ สูงกว่าช่วงแนะนำ (>1.25)  เสี่ยง Clipping"
    else:
        dc_status = "⚠️ ต่ำกว่าช่วงแนะนำ (<1.10)  Inverter ใหญ่เกินไป"

    ratio_t = Table([[
        Paragraph("DC/AC Ratio", ParagraphStyle("RL", fontName="TH-B", fontSize=11,
                   alignment=TA_CENTER, textColor=C_WHITE)),
        Paragraph(_s(dc_ratio,"{:.2f}"), ParagraphStyle("RV", fontName="TH-B", fontSize=24,
                   alignment=TA_CENTER, textColor=C_HEADER)),
        Paragraph(dc_status, ParagraphStyle("RS", fontName="TH", fontSize=10)),
    ]], colWidths=[CW*0.22, CW*0.18, CW*0.60])
    ratio_t.setStyle(_base_ts([
        ("BACKGROUND",    (0,0),(0,0), C_HEADER),
        ("BACKGROUND",    (1,0),(1,0), C_ORANGE),
        ("BACKGROUND",    (2,0),(2,0), dc_ratio_color),
        ("TOPPADDING",    (0,0),(-1,-1), 7),
        ("BOTTOMPADDING", (0,0),(-1,-1), 7),
    ]))
    story.append(ratio_t)
    story.append(Spacer(1, 5))

    # =========================================================
    # 6. FINANCIAL + CO2 TABLE
    # =========================================================
    payback_str = _s(simple_payback, "{:.1f} ปี") if simple_payback else f">{project_life} ปี"
    disc_str    = _s(discounted_payback, "{:.1f} ปี") if discounted_payback else f">{project_life} ปี"
    irr_str     = _s(irr_val*100, "{:.1f}%") if irr_val is not None else "-"
    npv_str     = _s(npv, "{:,.0f} ฿") if npv is not None else "-"
    capex_str   = _s(CAPEX, "{:,.0f} ฿") if CAPEX else "-"

    def fin_hdr(txt):
        return Paragraph(txt, ParagraphStyle("fh", fontName="TH-B", fontSize=8,
                          alignment=TA_CENTER, textColor=C_WHITE))
    def fin_val(txt):
        return Paragraph(txt, ParagraphStyle("fv", fontName="TH-B", fontSize=9,
                          alignment=TA_CENTER))

    fin_data = [[
        fin_hdr("เงินลงทุน\n(CAPEX)"),
        fin_hdr("ระยะเวลาคืนทุน\n(Simple Payback)"),
        fin_hdr("ระยะเวลาคืนทุน\n(Discounted)"),
        fin_hdr("IRR"),
        fin_hdr("NPV"),
        fin_hdr("ลด CO₂/ปี"),
        fin_hdr("ลด CO₂\n25 ปี"),
    ],[
        fin_val(capex_str),
        fin_val(payback_str),
        fin_val(disc_str),
        fin_val(irr_str),
        fin_val(npv_str),
        fin_val(f"{co2_yr:.2f} t"),
        fin_val(f"{co2_25:.1f} t"),
    ]]

    fin_t = Table(fin_data, colWidths=[CW/7]*7)
    fin_t.setStyle(_base_ts([
        ("BACKGROUND", (0,0),(-1,0), C_HEADER),
        ("BACKGROUND", (0,1),(-1,1), C_YELLOW),
        ("ALIGN",      (0,0),(-1,-1), "CENTER"),
        ("TOPPADDING",    (0,0),(-1,-1), 5),
        ("BOTTOMPADDING", (0,0),(-1,-1), 5),
    ]))
    story.append(Paragraph("ผลการวิเคราะห์เศรษฐศาสตร์และสิ่งแวดล้อม", S["SEC"]))
    story.append(fin_t)
    story.append(Spacer(1, 5))

    # =========================================================
    # 7. AI RECOMMENDATION
    # =========================================================
    ai_text = (ai_result if ai_result
               else "ยังไม่ได้รัน AI Analysis — กด Generate AI Recommendation ก่อน Export")
    story.append(Paragraph("AI Component Recommendation", S["SEC"]))
    story.append(Paragraph(ai_text.replace("\n","<br/>"), S["BODY"]))
    story.append(Spacer(1, 4))

    # =========================================================
    # 8. FOOTER
    # =========================================================
    story.append(HRFlowable(width=CW, thickness=0.4, color=C_BORDER))
    story.append(Paragraph(
        f"Generated by Solar Rooftop Designer  |  "
        f"{datetime.now().strftime('%d/%m/%Y %H:%M')}  |  "
        "SF_VOC_COLD = 1.20  |  SF_VMP_HOT = 0.90  |  SF_CURRENT = 1.25",
        S["TINY"]
    ))

    doc.build(story)
    return buffer.getvalue()