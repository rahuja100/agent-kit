"""PowerPoint report generator for Sales Enablement Agent.

Uses a single user-provided .pptx template for branding.
Reads slide layouts from the template and generates content slides.
"""
# pyright: basic  # python-pptx has no type stubs

import logging
from typing import Any

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

logger = logging.getLogger(__name__)

# Colors
COLOR_HEADER_BG = RGBColor(0x2C, 0x3E, 0x50)
COLOR_ALT_ROW = RGBColor(0xF5, 0xF5, 0xF5)
COLOR_DARK = RGBColor(0x33, 0x33, 0x33)
COLOR_BODY = RGBColor(0x55, 0x55, 0x55)
COLOR_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
COLOR_IDENTICAL = RGBColor(0x27, 0xAE, 0x60)
COLOR_SIMILAR = RGBColor(0x52, 0xBE, 0x80)  # lighter green for "High"
COLOR_LESS_SIMILAR = RGBColor(0xE7, 0x4C, 0x3C)
COLOR_NO_MATCH = RGBColor(0xBD, 0xC3, 0xC7)

SIMILARITY_COLORS = {"identical": COLOR_IDENTICAL, "similar": COLOR_SIMILAR, "less_similar": COLOR_LESS_SIMILAR}
SIMILARITY_LABELS = {"identical": "Identical", "similar": "Similar", "less_similar": "Less Similar"}

# Portco report uses different labels
COMPARABILITY_MAP = {"identical": ("Very High", COLOR_IDENTICAL), "similar": ("High", COLOR_SIMILAR), "less_similar": ("Medium", COLOR_LESS_SIMILAR)}


FONT_NAME = "Avenir"
TITLE_COLOR = RGBColor(0x39, 0x39, 0x41)
COLOR_TREYA_ORANGE = RGBColor(0xF6, 0x6C, 0x2E)


def _layout(prs, idx: int = 1):
    """Get slide layout by index, fallback to last available."""
    layouts = prs.slide_layouts
    return layouts[idx] if idx < len(layouts) else layouts[len(layouts) - 1]


def _add_slide(prs, title: str, layout_idx: int = 1):
    """Add a content slide with a title text box matching the template style."""
    slide = prs.slides.add_slide(_layout(prs, layout_idx))
    # Add title as text box (template has no title placeholder)
    title_box = slide.shapes.add_textbox(Inches(0.6), Inches(0.15), Inches(8.0), Inches(0.6))
    tf = title_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = title
    run.font.size = Pt(26)
    run.font.bold = True
    run.font.color.rgb = TITLE_COLOR
    run.font.name = FONT_NAME

    # Page number bottom right
    num_box = slide.shapes.add_textbox(Inches(9.0), Inches(5.2), Inches(0.8), Inches(0.3))
    ntf = num_box.text_frame
    np = ntf.paragraphs[0]
    np.alignment = PP_ALIGN.RIGHT
    nr = np.add_run()
    nr.text = str(len(prs.slides))
    nr.font.size = Pt(9)
    nr.font.color.rgb = COLOR_BODY
    nr.font.name = FONT_NAME

    return slide


def _textbox(slide, text, left, top, width, height, size=11, bold=False, color=None):
    """Add a text box (dimensions in inches)."""
    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color or COLOR_BODY
    run.font.name = FONT_NAME
    return box


def _table(slide, rows, cols, left, top, width, height):
    """Add a table (dimensions in inches)."""
    return slide.shapes.add_table(
        rows, cols, Inches(left), Inches(top), Inches(width), Inches(height)
    ).table


def _cell(cell, text, size=10, bold=False, color=None, align=PP_ALIGN.LEFT):
    """Format a table cell."""
    cell.text = ""
    p = cell.text_frame.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = str(text)
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color or COLOR_DARK
    run.font.name = FONT_NAME
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE


def _fill(cell, color):
    """Set cell background."""
    cell.fill.solid()
    cell.fill.fore_color.rgb = color


def _header_row(tbl, headers):
    """Format header row with dark background. Font size larger than body text."""
    for i, h in enumerate(headers):
        _cell(tbl.cell(0, i), h, size=11, bold=True, color=COLOR_WHITE)
        _fill(tbl.cell(0, i), COLOR_HEADER_BG)


def _alt_row(tbl, row, cols):
    """Apply alternating row shading."""
    if row % 2 == 0:
        for c in range(cols):
            _fill(tbl.cell(row, c), COLOR_ALT_ROW)


# ── PE Report Slide generators ────────────────────────────────────

def _slide_title(prs, data):
    """Update the cover slide with PE firm name."""
    name = _title_case(data["pe_firm"]["name"])
    if len(prs.slides) > 0:
        slide = prs.slides[0]
        for shape in slide.shapes:
            if shape.has_text_frame and "unlock" in shape.text_frame.text.lower():
                from pptx.util import Emu
                tf = shape.text_frame
                tf.clear()
                p1 = tf.paragraphs[0]
                run1 = p1.add_run()
                run1.text = "Procurement Savings Opportunity"
                run1.font.name = FONT_NAME
                run1.font.size = Pt(40)
                run1.font.bold = True
                run1.font.color.rgb = COLOR_WHITE
                p2 = tf.add_paragraph()
                p2.space_before = Emu(200000)
                run2 = p2.add_run()
                run2.text = name
                run2.font.name = FONT_NAME
                run2.font.size = Pt(28)
                run2.font.bold = False
                run2.font.color.rgb = COLOR_WHITE
                break


def _slide_exec_summary(prs, data):
    from pptx.util import Emu

    slide = _add_slide(prs, "Executive Summary")
    box = slide.shapes.add_textbox(Inches(0.6), Inches(1.0), Inches(8.8), Inches(4.0))
    tf = box.text_frame
    tf.word_wrap = True

    pe_name = _title_case(data.get("pe_firm", {}).get("name", ""))
    companies = data.get("portfolio_companies", [])
    assessments = data.get("savings_assessment", [])
    total_sav = sum(a.get("total_estimated_savings_mm", 0) for a in assessments)
    total_rev = sum(a.get("estimated_revenue_mm", 0) for a in assessments)
    pct = (total_sav / total_rev * 100) if total_rev > 0 else 0

    def _section(tf, title, body, first=False):
        if not first:
            p = tf.add_paragraph()
            p.space_before = Emu(200000)
        else:
            p = tf.paragraphs[0]
        run = p.add_run()
        run.text = title
        run.font.size = Pt(12)
        run.font.bold = True
        run.font.color.rgb = COLOR_TREYA_ORANGE
        run.font.name = FONT_NAME
        p2 = tf.add_paragraph()
        p2.space_before = Emu(50000)
        run2 = p2.add_run()
        run2.text = body
        run2.font.size = Pt(10)
        run2.font.color.rgb = COLOR_BODY
        run2.font.name = FONT_NAME

    _section(tf, "Portfolio Overview",
             f"{pe_name} has {len(companies)} active portfolio companies with combined estimated revenue "
             f"of ${total_rev:,.0f}MM.", first=True)

    _section(tf, "Savings Opportunity",
             f"We estimate a total procurement savings opportunity of ${total_sav:,.1f}MM ({pct:.1f}% of revenue) "
             f"across the portfolio through strategic sourcing of addressable direct and indirect spend.")

    # Top targets
    ranked = sorted(assessments, key=lambda a: a.get("total_estimated_savings_mm", 0), reverse=True)[:3]
    if ranked:
        targets = ", ".join(f"{a['portfolio_company']} (${a.get('total_estimated_savings_mm', 0):,.1f}MM)" for a in ranked)
        _section(tf, "Priority Targets", f"Top opportunities: {targets}.")

    # Comps
    comps = data.get("comps_analysis", [])
    strong_count = sum(1 for c in comps if any(m.get("similarity") == "identical" for m in c.get("matching_clients", [])))
    if strong_count:
        _section(tf, "Treya Experience",
                 f"Treya has direct comparable experience with {strong_count} of {len(companies)} portfolio companies, "
                 f"enabling faster engagement and higher-confidence savings delivery.")

    _section(tf, "Recommended Next Steps",
             "Prioritize outreach to top savings targets, prepare tailored case studies based on comparable engagements, "
             "and propose detailed spend assessments for the highest-value opportunities.")

    _textbox(slide, "Note: All estimates based on outside-in analysis of publicly available revenue and spend data.",
             0.6, 5.25, 8.8, 0.3, size=7, color=COLOR_BODY)


def _slide_portfolio(prs, data):
    companies = data.get("portfolio_companies", [])
    savings_map = {a["portfolio_company"]: a.get("total_estimated_savings_mm", 0) for a in data.get("savings_assessment", [])}
    if not companies:
        return
    slide = _add_slide(prs, "Portfolio Overview")
    rows = len(companies) + 2
    tbl = slide.shapes.add_table(rows, 4, Inches(0.3), Inches(1.1), Inches(9.4), Inches(min(4.0, 0.25 * rows))).table

    tbl.columns[0].width = Inches(3.2)
    tbl.columns[1].width = Inches(3.0)
    tbl.columns[2].width = Inches(1.6)
    tbl.columns[3].width = Inches(1.6)

    _header_row(tbl, ["Portfolio Company", "Industry", "Est. Revenue\n($MM)", "Savings Opp.\n($MM)"])

    total_rev = total_sav = 0.0
    for r, co in enumerate(companies, 1):
        rev = co.get("estimated_revenue_mm", 0)
        sav = savings_map.get(co.get("name", ""), 0)
        total_rev += rev
        total_sav += sav
        _cell(tbl.cell(r, 0), co.get("name", ""), size=9, bold=True)
        _cell(tbl.cell(r, 1), co.get("industry", ""), size=9)
        _cell(tbl.cell(r, 2), f"${rev:,.0f}", size=9, align=PP_ALIGN.RIGHT)
        _cell(tbl.cell(r, 3), f"${sav:,.1f}", size=9, align=PP_ALIGN.RIGHT)
        _alt_row(tbl, r, 4)

    tr = len(companies) + 1
    for i, val in enumerate(["TOTAL", "", f"${total_rev:,.0f}", f"${total_sav:,.1f}"]):
        _cell(tbl.cell(tr, i), val, size=9, bold=True, color=COLOR_WHITE, align=PP_ALIGN.RIGHT if i >= 2 else PP_ALIGN.LEFT)
        _fill(tbl.cell(tr, i), COLOR_HEADER_BG)

    _textbox(slide, "Note: All estimates based on outside-in analysis of publicly available data.",
             0.3, 5.25, 9.4, 0.3, size=7, color=COLOR_BODY)


def _slide_comps_highlights(prs, data):
    """PE comps slide using Comparability labels and colors."""
    comps = data.get("comps_analysis", [])
    if not comps:
        return

    slide = _add_slide(prs, "Treya Experience: Comparable Engagements")

    rows_data: list[tuple[str, str, str, str]] = []
    for entry in comps:
        pc = entry["portfolio_company"]
        strong = [m for m in entry.get("matching_clients", []) if m["similarity"] in ("identical", "similar")]
        if not strong:
            continue
        for m in strong[:2]:
            rows_data.append((pc, m["client_name"], m["similarity"], m.get("rationale", "")))

    if not rows_data:
        _textbox(slide, "No strong comparable engagements identified.", 0.5, 1.5, 9.0, 1.0)
        return

    rows = len(rows_data) + 1
    tbl = slide.shapes.add_table(rows, 4, Inches(0.5), Inches(1.2), Inches(9.0), Inches(min(4.2, 0.30 * rows))).table

    tbl.columns[0].width = Inches(2.2)
    tbl.columns[1].width = Inches(2.0)
    tbl.columns[2].width = Inches(1.2)
    tbl.columns[3].width = Inches(3.6)

    _header_row(tbl, ["Portfolio Company", "Treya Client", "Comparability", "Company Description"])

    for r, (pc, client, sim, rationale) in enumerate(rows_data, 1):
        _cell(tbl.cell(r, 0), pc, size=9, bold=True)
        _cell(tbl.cell(r, 1), client, size=9)
        if sim in COMPARABILITY_MAP:
            label, color = COMPARABILITY_MAP[sim]
            _cell(tbl.cell(r, 2), label, size=9, color=COLOR_WHITE, align=PP_ALIGN.CENTER)
            _fill(tbl.cell(r, 2), color)
        else:
            _cell(tbl.cell(r, 2), sim, size=9, align=PP_ALIGN.CENTER)
        _cell(tbl.cell(r, 3), rationale, size=8)


COGS_MATERIALS_KEYWORDS = {"raw materials", "direct materials", "food ingredients", "commodities", "pharmaceuticals", "lab supplies", "medical supplies", "devices", "medical equipment"}
COGS_SEPARATE_KEYWORDS = {"packaging", "containers", "freight", "logistics", "parcel", "shipping", "fuel", "energy", "mro", "maintenance", "janitorial"}
COGS_ADDRESSABLE_OVERRIDE = 0.20  # consolidated COGS materials line uses 20% addressability


def _consolidate_for_display(categories: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Consolidate categories for PPTX display:
    - Merge COGS material categories into one line at 20% addressability
    - Keep packaging, freight, fuel, medical supplies, MRO separate
    - Show top 5-6 indirect individually, lump rest as 'Other Indirect'
    """
    cogs_materials: list[dict[str, Any]] = []
    cogs_separate: list[dict[str, Any]] = []
    indirects: list[dict[str, Any]] = []

    for cat in categories:
        name_lower = cat.get("category", "").lower()
        is_cogs_material = any(kw in name_lower for kw in COGS_MATERIALS_KEYWORDS)
        is_cogs_separate = any(kw in name_lower for kw in COGS_SEPARATE_KEYWORDS)

        if is_cogs_material and not is_cogs_separate:
            cogs_materials.append(cat)
        elif is_cogs_separate:
            cogs_separate.append(cat)
        else:
            indirects.append(cat)

    display_rows: list[dict[str, Any]] = []

    # Consolidated COGS materials line -- pass through LLM numbers as-is, just sum them
    if cogs_materials:
        component_names = [c["category"].split("(")[0].strip() for c in cogs_materials]
        total_spend = sum(c.get("estimated_spend_mm", 0) for c in cogs_materials)
        total_savings = sum(c.get("estimated_savings_mm", 0) for c in cogs_materials)
        avg_pct = (total_savings / total_spend * 100) if total_spend > 0 else 0

        label = f"COGS ({', '.join(component_names)})"
        display_rows.append({"category": label, "estimated_spend_mm": total_spend, "savings_opportunity_pct": avg_pct, "estimated_savings_mm": total_savings})

    # Separate COGS lines (packaging, freight, fuel, medical, MRO)
    for cat in sorted(cogs_separate, key=lambda c: c.get("estimated_savings_mm", 0), reverse=True):
        display_rows.append(cat)

    # Target 10 total rows: fill remaining slots with top indirect, lump rest
    used_rows = len(display_rows)  # COGS consolidated + separate lines so far
    indirect_slots = max(3, 10 - used_rows - 1)  # -1 for "Other Indirect" row
    indirects_sorted = sorted(indirects, key=lambda c: c.get("estimated_savings_mm", 0), reverse=True)
    top_indirect = indirects_sorted[:indirect_slots]
    other_indirect = indirects_sorted[indirect_slots:]

    for cat in top_indirect:
        display_rows.append(cat)

    if other_indirect:
        other_spend = sum(c.get("estimated_spend_mm", 0) for c in other_indirect)
        other_savings = sum(c.get("estimated_savings_mm", 0) for c in other_indirect)
        other_pct = (other_savings / other_spend * 100) if other_spend > 0 else 0
        display_rows.append({"category": "Other Indirects", "estimated_spend_mm": other_spend, "savings_opportunity_pct": other_pct, "estimated_savings_mm": other_savings})

    return display_rows


def _slide_company_detail(prs, assessment):
    """Per-company savings table -- same 4-column format as portco."""
    categories = assessment.get("spend_categories", [])
    if not categories:
        return
    pc = assessment["portfolio_company"]
    rev = assessment.get("estimated_revenue_mm", 0)
    total_sav = assessment.get("total_estimated_savings_mm", 0)

    slide = _add_slide(prs, f"Savings Detail: {pc}")
    _textbox(slide, f"Est. Revenue: ${rev:,.0f}MM  |  Est. Savings: ${total_sav:,.1f}MM",
             0.5, 0.85, 9.0, 0.3, size=10, bold=True, color=TITLE_COLOR)

    display = _consolidate_for_display(categories)

    rows = len(display) + 2
    cols = 4
    table_height = 0.25 * rows
    tbl = slide.shapes.add_table(rows, cols, Inches(0.3), Inches(1.1), Inches(9.4), Inches(table_height)).table

    tbl.columns[0].width = Inches(4.5)
    tbl.columns[1].width = Inches(2.0)
    tbl.columns[2].width = Inches(1.2)
    tbl.columns[3].width = Inches(1.7)

    _header_row(tbl, ["Category", "Est. Addressable\nSpend ($MM)", "Savings\n(%)", "Est. Savings\n($MM)"])

    tot_addr = tot_sav_disp = 0.0
    for r, cat in enumerate(display, 1):
        addr_spend = cat.get("estimated_spend_mm", 0)
        pct = cat.get("savings_opportunity_pct", 0)
        sav = cat.get("estimated_savings_mm", 0)
        tot_addr += addr_spend
        tot_sav_disp += sav
        _cell(tbl.cell(r, 0), cat["category"], size=9)
        _cell(tbl.cell(r, 1), f"${addr_spend:,.1f}", size=9, align=PP_ALIGN.RIGHT)
        _cell(tbl.cell(r, 2), f"{pct:.0f}%", size=9, align=PP_ALIGN.RIGHT)
        _cell(tbl.cell(r, 3), f"${sav:,.1f}", size=9, align=PP_ALIGN.RIGHT)
        _alt_row(tbl, r, cols)

    tr = len(display) + 1
    for i, val in enumerate(["TOTAL", f"${tot_addr:,.1f}", "", f"${tot_sav_disp:,.1f}"]):
        _cell(tbl.cell(tr, i), val, size=9, bold=True, color=COLOR_WHITE, align=PP_ALIGN.RIGHT if i > 0 else PP_ALIGN.LEFT)
        _fill(tbl.cell(tr, i), COLOR_HEADER_BG)

    _textbox(slide, "Note: Estimates based on outside-in analysis of publicly available data.",
             0.3, 5.25, 9.4, 0.3, size=7, color=COLOR_BODY)


def _slide_next_steps(prs, data):
    from pptx.util import Emu

    slide = _add_slide(prs, "Recommended Next Steps")

    box = slide.shapes.add_textbox(Inches(0.6), Inches(1.0), Inches(8.8), Inches(4.2))
    tf = box.text_frame
    tf.word_wrap = True

    def _section(tf, title, first=False):
        if not first:
            p = tf.add_paragraph()
            p.space_before = Emu(200000)
        else:
            p = tf.paragraphs[0]
        run = p.add_run()
        run.text = title
        run.font.size = Pt(12)
        run.font.bold = True
        run.font.color.rgb = COLOR_TREYA_ORANGE
        run.font.name = FONT_NAME
        return tf

    def _bullet(tf, text):
        p = tf.add_paragraph()
        p.space_before = Emu(50000)
        run = p.add_run()
        run.text = f"\u2022  {text}"
        run.font.size = Pt(10)
        run.font.color.rgb = COLOR_BODY
        run.font.name = FONT_NAME

    _section(tf, "Priority Targets", first=True)
    ranked = sorted(data.get("savings_assessment", []),
                    key=lambda a: a.get("total_estimated_savings_mm", 0), reverse=True)
    for a in ranked[:5]:
        sav = a.get("total_estimated_savings_mm", 0)
        _bullet(tf, f"{a['portfolio_company']} \u2013 Est. ${sav:,.1f}MM savings opportunity")

    _section(tf, "Suggested Actions")
    _bullet(tf, "Schedule introductory calls with PE operating partners")
    _bullet(tf, "Prepare tailored case studies based on comparable engagements")
    _bullet(tf, "Develop detailed spend assessment proposals for top targets")
    _bullet(tf, "Propose phased implementation with 90-day quick-win milestones")


def _slide_appendix(prs):
    from pptx.util import Emu

    slide = _add_slide(prs, "Appendix: Assumptions & Data Sources")

    box = slide.shapes.add_textbox(Inches(0.6), Inches(1.0), Inches(8.8), Inches(4.2))
    tf = box.text_frame
    tf.word_wrap = True

    def _section(tf, title, first=False):
        if not first:
            p = tf.add_paragraph()
            p.space_before = Emu(200000)
        else:
            p = tf.paragraphs[0]
        run = p.add_run()
        run.text = title
        run.font.size = Pt(12)
        run.font.bold = True
        run.font.color.rgb = COLOR_TREYA_ORANGE
        run.font.name = FONT_NAME

    def _bullet(tf, text):
        p = tf.add_paragraph()
        p.space_before = Emu(50000)
        run = p.add_run()
        run.text = f"\u2022  {text}"
        run.font.size = Pt(10)
        run.font.color.rgb = COLOR_BODY
        run.font.name = FONT_NAME

    _section(tf, "Data Sources", first=True)
    _bullet(tf, "Portfolio company information from PE firm website and public records")
    _bullet(tf, "Revenue estimates triangulated from employee count, PE fund size, and public filings")
    _bullet(tf, "Industry classifications based on primary business activity")

    _section(tf, "Key Assumptions")
    _bullet(tf, "Addressable spend estimated using outside-in analysis of company operations")
    _bullet(tf, "Savings percentages reflect typical opportunities at moderate procurement maturity")
    _bullet(tf, "Actual savings depend on current contract structures, supplier landscape, and implementation timeline")
    _bullet(tf, "Mid-market PE portfolio companies typically range $50MM\u2013$500MM revenue")


# ── Main entry point ─────────────────────────────────────────────

def _clear_slides_except_first(prs):
    """Remove all slides except the first (cover slide)."""
    while len(prs.slides) > 1:
        rId = prs.slides._sldIdLst[1].rId
        prs.part.drop_rel(rId)
        del prs.slides._sldIdLst[1]


def generate_report(data: dict[str, Any], template_path: str, output_path: str) -> int:
    """Generate the full PPTX report. Returns slide count."""
    prs = Presentation(template_path)

    # Keep slide 0 (cover slide with "Unlock Hidden Value" graphic), remove the rest
    _clear_slides_except_first(prs)

    # Slide 1: Title (updates the kept cover slide)
    _slide_title(prs, data)
    # Slide 2: Executive Summary (structured with Treya orange headers)
    _slide_exec_summary(prs, data)
    # Slide 3: Portfolio Overview
    _slide_portfolio(prs, data)
    # Slide 4: Comps Highlights
    _slide_comps_highlights(prs, data)
    # Slides 5+: Per-company savings detail (top 5 by savings, 4-column format)
    for a in sorted(data.get("savings_assessment", []),
                    key=lambda x: x.get("total_estimated_savings_mm", 0), reverse=True)[:5]:
        _slide_company_detail(prs, a)
    # Next Steps
    _slide_next_steps(prs, data)
    # Appendix
    _slide_appendix(prs)

    prs.save(output_path)
    count = len(prs.slides)
    logger.info(f"Saved {count} slides to {output_path}")
    return count


# ── Portco report ─────────────────────────────────────────────────

def _title_case(name: str) -> str:
    """Capitalize initial letters of each word."""
    return " ".join(w.capitalize() if w.islower() else w for w in name.split())


def _portco_slide_title(prs, data):
    """Update the cover slide text. Keeps the 'Unlock Hidden Value' graphic."""
    name = _title_case(data.get("company_name", "Company"))
    if len(prs.slides) > 0:
        slide = prs.slides[0]
        for shape in slide.shapes:
            if shape.has_text_frame and "unlock" in shape.text_frame.text.lower():
                # Clear existing paragraphs and rebuild
                tf = shape.text_frame
                tf.clear()
                p1 = tf.paragraphs[0]
                run1 = p1.add_run()
                run1.text = "Procurement Savings Opportunity"
                run1.font.name = FONT_NAME
                run1.font.size = Pt(40)
                run1.font.bold = True
                run1.font.color.rgb = COLOR_WHITE

                from pptx.util import Emu
                p2 = tf.add_paragraph()
                p2.space_before = Emu(200000)
                run2 = p2.add_run()
                run2.text = name
                run2.font.name = FONT_NAME
                run2.font.size = Pt(28)
                run2.font.bold = False
                run2.font.color.rgb = COLOR_WHITE
                break


def _portco_slide_exec_summary(prs, data):
    from pptx.util import Emu

    slide = _add_slide(prs, "Executive Summary")

    top = Inches(1.0)
    left = Inches(0.6)
    width = Inches(8.8)

    box = slide.shapes.add_textbox(left, top, width, Inches(4.0))
    tf = box.text_frame
    tf.word_wrap = True

    name = _title_case(data.get("company_name", ""))
    industry = data.get("industry", "")
    desc = data.get("description", "")
    rev = data.get("estimated_revenue_mm", 0)
    total_sav = data.get("total_estimated_savings_mm", 0)
    pct = (total_sav / rev * 100) if rev > 0 else 0

    # Company overview
    def _add_section(tf, title_text, body_text, first=False):
        if not first:
            p = tf.add_paragraph()
            p.space_before = Emu(200000)
        else:
            p = tf.paragraphs[0]
        run = p.add_run()
        run.text = title_text
        run.font.size = Pt(12)
        run.font.bold = True
        run.font.color.rgb = COLOR_TREYA_ORANGE
        run.font.name = FONT_NAME

        p2 = tf.add_paragraph()
        p2.space_before = Emu(50000)
        run2 = p2.add_run()
        run2.text = body_text
        run2.font.size = Pt(10)
        run2.font.color.rgb = COLOR_BODY
        run2.font.name = FONT_NAME

    _add_section(tf, "Company Overview",
                 f"{name} is a {industry.lower()} company. {desc}" + (f" Estimated revenue: ${rev:,.0f}MM." if rev else ""),
                 first=True)

    _add_section(tf, "Savings Opportunity",
                 f"Based on our outside-in assessment, we estimate a total procurement savings opportunity "
                 f"of ${total_sav:,.1f}MM ({pct:.1f}% of revenue) across addressable direct and indirect spend categories.")

    # Comps mention
    clients = data.get("matching_clients", [])
    strong = [m for m in clients if m.get("similarity") in ("identical", "similar")]
    if strong:
        client_names = ", ".join(m["client_name"] for m in strong[:3])
        _add_section(tf, "Treya Experience",
                     f"Treya has completed similar engagements with {len(strong)} comparable companies "
                     f"including {client_names}. Our direct experience in this sector enables faster "
                     f"time-to-value and higher confidence in savings delivery.")
    else:
        _add_section(tf, "Treya Experience",
                     "Treya brings procurement expertise across a broad range of industries and spend categories "
                     "applicable to this engagement.")

    _add_section(tf, "Recommended Next Steps",
                 "Conduct a detailed spend assessment with actual data to validate savings estimates, "
                 "identify quick-win categories for immediate sourcing events, and develop a phased "
                 "implementation roadmap.")

    # Disclaimer
    _textbox(slide, "Note: All estimates based on outside-in analysis of publicly available revenue and spend data. Actual results will vary.",
             0.6, 5.2, 8.8, 0.3, size=7, color=COLOR_BODY)


def _portco_slide_comps(prs, data):
    all_clients = data.get("matching_clients", [])
    # Filter out less_similar
    clients = [m for m in all_clients if m.get("similarity") in ("identical", "similar")]
    if not clients:
        return

    slide = _add_slide(prs, "Treya Experience: Comparable Engagements")

    rows = len(clients) + 1
    # 3 columns: client name (2.0"), comparability (1.2"), description (6.2")
    tbl = slide.shapes.add_table(rows, 3, Inches(0.5), Inches(1.2), Inches(9.0), Inches(min(4.5, 0.35 * rows))).table

    # Set column widths
    tbl.columns[0].width = Inches(2.0)
    tbl.columns[1].width = Inches(1.2)
    tbl.columns[2].width = Inches(5.8)

    _header_row(tbl, ["Treya Client", "Comparability", "Company Description"])

    for r, m in enumerate(clients, 1):
        _cell(tbl.cell(r, 0), m.get("client_name", ""), size=10, bold=True)

        sim = m.get("similarity", "")
        if sim in COMPARABILITY_MAP:
            label, color = COMPARABILITY_MAP[sim]
            _cell(tbl.cell(r, 1), label, size=9, color=COLOR_WHITE, align=PP_ALIGN.CENTER)
            _fill(tbl.cell(r, 1), color)
        else:
            _cell(tbl.cell(r, 1), sim, size=9, align=PP_ALIGN.CENTER)

        _cell(tbl.cell(r, 2), m.get("rationale", ""), size=9)


def _portco_slide_savings(prs, data):
    """Savings detail table sized to fit on one page."""
    categories = data.get("spend_categories", [])
    if not categories:
        return

    slide = _add_slide(prs, "Savings Opportunity Detail")

    rev = data.get("estimated_revenue_mm", 0)
    total_sav = data.get("total_estimated_savings_mm", 0)
    _textbox(slide, f"Est. Revenue: ${rev:,.0f}MM  |  Est. Savings: ${total_sav:,.1f}MM",
             0.5, 0.85, 9.0, 0.3, size=10, bold=True, color=TITLE_COLOR)

    display = _consolidate_for_display(categories)

    rows = len(display) + 2  # +1 header +1 total row
    cols = 4
    table_height = 0.25 * rows
    tbl = slide.shapes.add_table(rows, cols, Inches(0.3), Inches(1.1), Inches(9.4), Inches(table_height)).table

    tbl.columns[0].width = Inches(4.5)
    tbl.columns[1].width = Inches(2.0)
    tbl.columns[2].width = Inches(1.2)
    tbl.columns[3].width = Inches(1.7)

    _header_row(tbl, ["Category", "Est. Addressable\nSpend ($MM)", "Savings\n(%)", "Est. Savings\n($MM)"])

    tot_addr = tot_sav = 0.0
    for r, cat in enumerate(display, 1):
        addr_spend = cat.get("estimated_spend_mm", 0)
        pct = cat.get("savings_opportunity_pct", 0)
        sav = cat.get("estimated_savings_mm", 0)
        tot_addr += addr_spend
        tot_sav += sav
        _cell(tbl.cell(r, 0), cat["category"], size=9)
        _cell(tbl.cell(r, 1), f"${addr_spend:,.1f}", size=9, align=PP_ALIGN.RIGHT)
        _cell(tbl.cell(r, 2), f"{pct:.0f}%", size=9, align=PP_ALIGN.RIGHT)
        _cell(tbl.cell(r, 3), f"${sav:,.1f}", size=9, align=PP_ALIGN.RIGHT)
        _alt_row(tbl, r, cols)

    tr = len(display) + 1
    for i, val in enumerate(["TOTAL", f"${tot_addr:,.1f}", "", f"${tot_sav:,.1f}"]):
        _cell(tbl.cell(tr, i), val, size=9, bold=True, color=COLOR_WHITE, align=PP_ALIGN.RIGHT if i > 0 else PP_ALIGN.LEFT)
        _fill(tbl.cell(tr, i), COLOR_HEADER_BG)

    # Disclaimer -- fixed at bottom of slide
    _textbox(slide, "Note: Estimates based on outside-in analysis of publicly available data. Actual spend and savings will vary based on detailed assessment.",
             0.3, 5.25, 9.4, 0.3, size=7, color=COLOR_BODY)


def _portco_slide_savings_levers(prs, data):
    slide = _add_slide(prs, "Key Savings Levers")

    categories = data.get("spend_categories", [])
    # Take more than needed before dedup to ensure at least 5 unique levers
    top_cats = sorted(categories, key=lambda c: c.get("estimated_savings_mm", 0), reverse=True)[:12]

    if not top_cats:
        return

    # Savings lever descriptions by category
    lever_descriptions = {
        "freight": "Consolidate carriers, renegotiate lane rates, optimize mode mix, implement competitive bidding",
        "parcel": "Carrier diversification, zone-skip strategies, dimensional weight optimization, volume aggregation",
        "shipping": "Consolidate carriers, renegotiate lane rates, optimize mode mix, implement competitive bidding",
        "logistics": "Consolidate carriers, renegotiate lane rates, optimize mode mix, implement competitive bidding",
        "it": "License rationalization, competitive rebid of managed services, consolidate vendors, renegotiate renewals",
        "software": "License rationalization, competitive rebid of managed services, consolidate vendors, renegotiate renewals",
        "telecom": "Audit circuits for unused lines, competitive rebid, consolidate carriers, renegotiate rates",
        "insurance": "Broker remarketing, deductible optimization, loss control programs, carrier diversification",
        "benefits": "Plan design review, broker remarketing, pharmacy carve-out, voluntary benefit optimization",
        "fleet": "Lifecycle optimization, fuel program negotiation, maintenance consolidation, telematics-driven savings",
        "vehicle": "Lifecycle optimization, fuel program negotiation, maintenance consolidation, telematics-driven savings",
        "mro": "Consolidate suppliers, implement catalog programs, standardize specifications, competitive bidding",
        "maintenance": "Consolidate suppliers, implement catalog programs, standardize specifications, competitive bidding",
        "marketing": "Agency fee benchmarking, media buying consolidation, production cost reduction, digital optimization",
        "travel": "Policy enforcement, preferred supplier programs, booking tool optimization, meeting consolidation",
        "waste": "Competitive rebid, right-size containers, optimize pickup frequency, recycling revenue capture",
        "temp": "Rate card renegotiation, supplier consolidation, conversion-to-hire programs, MSP implementation",
        "contingent": "Rate card renegotiation, supplier consolidation, conversion-to-hire programs, MSP implementation",
        "professional": "Competitive RFP for retained services, rate benchmarking, scope optimization, panel consolidation",
        "printing": "Print management program, digital migration, supplier consolidation, specification standardization",
        "uniform": "Competitive rebid, program standardization, inventory optimization, direct purchasing",
        "packaging": "Specification optimization, supplier consolidation, volume aggregation, material substitution",
        "raw material": "Strategic sourcing, supplier diversification, specification review, volume leverage",
        "direct material": "Strategic sourcing, supplier diversification, specification review, volume leverage",
        "medical": "GPO optimization, formulary standardization, competitive bidding, physician preference alignment",
        "food": "Menu engineering, supplier consolidation, specification standardization, seasonal buying strategies",
        "fuel": "Fleet card programs, bulk purchasing, route optimization, alternative fuel evaluation",
        "energy": "Rate negotiation, demand management, renewable procurement, efficiency programs",
        "payroll": "Competitive rebid, platform consolidation, self-service optimization, scope review",
    }

    def _get_lever(cat_name: str) -> str:
        name_lower = cat_name.lower()
        for key, desc in lever_descriptions.items():
            if key in name_lower:
                return desc
        return "Competitive sourcing, supplier consolidation, specification review, demand management"

    # Deduplicate: skip categories that would produce the same lever text, keep at least 5
    seen_levers: set[str] = set()
    deduped: list[dict[str, Any]] = []
    for cat in top_cats:
        lever = _get_lever(cat.get("category", ""))
        if lever not in seen_levers:
            seen_levers.add(lever)
            deduped.append(cat)
        if len(deduped) >= 6:
            break

    rows = len(deduped) + 1
    tbl = slide.shapes.add_table(rows, 2, Inches(0.5), Inches(1.2), Inches(9.0), Inches(min(4.0, 0.35 * rows))).table
    tbl.columns[0].width = Inches(3.0)
    tbl.columns[1].width = Inches(6.0)

    _header_row(tbl, ["Savings Category", "Key Savings Levers"])

    for r, cat in enumerate(deduped, 1):
        _cell(tbl.cell(r, 0), cat.get("category", ""), size=10, bold=True)
        _cell(tbl.cell(r, 1), _get_lever(cat.get("category", "")), size=9)
        _alt_row(tbl, r, 2)


def generate_portco_report(data: dict[str, Any], template_path: str, output_path: str) -> int:
    """Generate a portco-level PPTX report (5 slides). Returns slide count."""
    prs = Presentation(template_path)

    _clear_slides_except_first(prs)

    _portco_slide_title(prs, data)
    _portco_slide_exec_summary(prs, data)
    _portco_slide_comps(prs, data)
    _portco_slide_savings(prs, data)
    _portco_slide_savings_levers(prs, data)

    prs.save(output_path)
    count = len(prs.slides)
    logger.info(f"Saved portco report: {count} slides to {output_path}")
    return count
