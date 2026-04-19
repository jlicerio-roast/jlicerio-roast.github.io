"""
GUIA Report Automation Tool
MVP — Early Diagnostic Insight Generator
Generates structured slide insights (heading, subheading, bullets, callouts) from CSV data.
"""

import streamlit as st
import pandas as pd
import numpy as np
import json
import io
import os
from docx import Document
from docx.shared import Pt, RGBColor
from pptx import Presentation as PPTXPres
from pptx.util import Emu, Pt as PPTXPt
from pptx.dml.color import RGBColor as RC
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
import anthropic
from dotenv import load_dotenv

load_dotenv()

# ─── SLIDE CONFIGURATION ─────────────────────────────────────────────────────

SUBJECT_DISPLAY_MAP = {
    "all":         ["Science", "Math", "Language (ENG)", "Reading (ENG)", "Language (FIL)", "Reading (FIL)"],
    "science":     ["Science"],
    "math":        ["Math"],
    "english":     ["Language (ENG)", "Reading (ENG)"],
    "english_lp":  ["Language (ENG)"],
    "english_rc":  ["Reading (ENG)"],
    "filipino":    ["Language (FIL)", "Reading (FIL)"],
    "filipino_lp": ["Language (FIL)"],
    "filipino_rc": ["Reading (FIL)"],
}

PREPAREDNESS_ORDER = [
    "[1]Not Yet Prepared",
    "[2]Partially Prepared",
    "[3]Adequately Prepared",
    "[4]Well Prepared",
]

SLIDES = [
    # ── Overall ──────────────────────────────────────────────────────────────
    {"id":  1, "section": "Overall",  "title": "Cohort-level Preparedness Level",
     "subject_key": "all",         "slide_type": "preparedness_overall"},
    {"id":  2, "section": "Overall",  "title": "Overall Accuracy per Subject and Difficulty",
     "subject_key": "all",         "slide_type": "accuracy_overall"},
    # ── Science ──────────────────────────────────────────────────────────────
    {"id":  3, "section": "Science",  "title": "Science Preparedness Level Distribution",
     "subject_key": "science",     "slide_type": "preparedness_subject"},
    {"id":  4, "section": "Science",  "title": "Science Accuracy per Subtopic and Difficulty",
     "subject_key": "science",     "slide_type": "accuracy_subtopic"},
    {"id":  5, "section": "Science",  "title": "Science Error Types",
     "subject_key": "science",     "slide_type": "error_types"},
    {"id":  6, "section": "Science",  "title": "Science Time Management and Accuracy",
     "subject_key": "science",     "slide_type": "time_management"},
    # ── Math ─────────────────────────────────────────────────────────────────
    {"id":  7, "section": "Math",     "title": "Math Preparedness Level Distribution",
     "subject_key": "math",        "slide_type": "preparedness_subject"},
    {"id":  8, "section": "Math",     "title": "Math Accuracy per Subtopic and Difficulty",
     "subject_key": "math",        "slide_type": "accuracy_subtopic"},
    {"id":  9, "section": "Math",     "title": "Math Error Types",
     "subject_key": "math",        "slide_type": "error_types"},
    {"id": 10, "section": "Math",     "title": "Math Time Management and Accuracy",
     "subject_key": "math",        "slide_type": "time_management"},
    # ── English ──────────────────────────────────────────────────────────────
    {"id": 11, "section": "English",  "title": "LP and RC in English — Preparedness Level Distribution",
     "subject_key": "english",     "slide_type": "preparedness_subject"},
    {"id": 12, "section": "English",  "title": "Language Proficiency in English — Accuracy per Subtopic and Difficulty",
     "subject_key": "english_lp",  "slide_type": "accuracy_subtopic"},
    {"id": 13, "section": "English",  "title": "Reading Comprehension in English — Accuracy per Subtopic and Difficulty",
     "subject_key": "english_rc",  "slide_type": "accuracy_subtopic"},
    {"id": 14, "section": "English",  "title": "Language Proficiency in English — Error Types",
     "subject_key": "english_lp",  "slide_type": "error_types"},
    {"id": 15, "section": "English",  "title": "Reading Comprehension in English — Error Types",
     "subject_key": "english_rc",  "slide_type": "error_types"},
    {"id": 16, "section": "English",  "title": "LP and RC in English — Time Management and Accuracy",
     "subject_key": "english",     "slide_type": "time_management"},
    # ── Filipino ─────────────────────────────────────────────────────────────
    {"id": 17, "section": "Filipino", "title": "LP and RC in Filipino — Preparedness Level Distribution",
     "subject_key": "filipino",    "slide_type": "preparedness_subject"},
    {"id": 18, "section": "Filipino", "title": "Language Proficiency in Filipino — Accuracy per Subtopic and Difficulty",
     "subject_key": "filipino_lp", "slide_type": "accuracy_subtopic"},
    {"id": 19, "section": "Filipino", "title": "Reading Comprehension in Filipino — Accuracy per Subtopic and Difficulty",
     "subject_key": "filipino_rc", "slide_type": "accuracy_subtopic"},
    {"id": 20, "section": "Filipino", "title": "Language Proficiency in Filipino — Error Types",
     "subject_key": "filipino_lp", "slide_type": "error_types"},
    {"id": 21, "section": "Filipino", "title": "Reading Comprehension in Filipino — Error Types",
     "subject_key": "filipino_rc", "slide_type": "error_types"},
    {"id": 22, "section": "Filipino", "title": "LP and RC in Filipino — Time Management and Accuracy",
     "subject_key": "filipino",    "slide_type": "time_management"},
]


# ─── DATA LOADING ─────────────────────────────────────────────────────────────

def load_csv(file_obj):
    """Load a CSV file object into a DataFrame."""
    if file_obj is None:
        return None
    try:
        file_obj.seek(0)
        return pd.read_csv(file_obj)
    except Exception as e:
        st.warning(f"Could not read file: {e}")
        return None


def filter_entity(df, entity_ids):
    """Filter DataFrame to one or more entity_ids."""
    if df is None or "entity_id" not in df.columns:
        return df
    return df[df["entity_id"].isin(entity_ids)].copy()


def filter_subjects(df, subject_key):
    """Filter DataFrame to the subjects relevant for a slide."""
    if df is None or df.empty or "subject_display" not in df.columns:
        return df
    subjects = SUBJECT_DISPLAY_MAP.get(subject_key, [])
    if not subjects:
        return df
    return df[df["subject_display"].isin(subjects)].copy()


# ─── METRICS COMPUTATION ──────────────────────────────────────────────────────

def score_stats_text(series, label="Score"):
    """Return a concise stats string for a numeric series."""
    if series.empty:
        return f"{label}: no data"
    return (
        f"{label}: Mean={series.mean():.1f}%, Median={series.median():.1f}%, "
        f"Min={series.min():.1f}%, Max={series.max():.1f}%, N={len(series)}"
    )


def fmt_preparedness(df_sq13, subjects=None):
    """
    Preparedness distribution from SQ1.3.
    subjects=None → all subjects; otherwise filter.
    """
    df = df_sq13.copy() if df_sq13 is not None else pd.DataFrame()
    if subjects:
        df = df[df["subject_display"].isin(subjects)]
    if df.empty:
        return "No preparedness data available."

    lines = []

    # Per-subject distribution
    for subj in df["subject_display"].unique():
        sdf = df[df["subject_display"] == subj]
        dist = sdf["preparedness_level"].value_counts()
        total = len(sdf)
        lines.append(f"\n{subj} (n={total}):")
        for level in PREPAREDNESS_ORDER:
            cnt = dist.get(level, 0)
            pct = cnt / total * 100
            lines.append(f"  {level}: {cnt} ({pct:.1f}%)")
        lines.append("  " + score_stats_text(sdf["first_mock_score"], "First Mock Score"))

    return "\n".join(lines)


def fmt_accuracy_overall(df_spv2):
    """Overall accuracy by subject × difficulty from SubPerf V2."""
    if df_spv2 is None or df_spv2.empty:
        return "No Subject Performance V2 data available."

    # Per-subject summary
    by_subj = (
        df_spv2.groupby("subject_display")["score_pct"]
        .agg(mean="mean", median="median")
        .round(1)
        .reset_index()
        .sort_values("median", ascending=False)
    )
    lines = ["PER-SUBJECT ACCURACY (all difficulties combined):"]
    for _, r in by_subj.iterrows():
        lines.append(f"  {r['subject_display']}: Mean={r['mean']}%, Median={r['median']}%")

    # Subject × Difficulty pivot
    if "difficulty_name" in df_spv2.columns:
        pivot = (
            df_spv2.groupby(["subject_display", "difficulty_name"])["score_pct"]
            .mean()
            .round(1)
            .unstack(fill_value=None)
        )
        lines.append("\nACCURACY BY SUBJECT × DIFFICULTY (mean %):")
        lines.append(pivot.to_string())

    return "\n".join(lines)


def fmt_accuracy_subtopic(df_spv2, subjects):
    """Accuracy by subtopic × difficulty for a specific subject group."""
    df = filter_subjects(df_spv2, None)  # already entity-filtered
    if subjects:
        df = df[df["subject_display"].isin(subjects)]
    if df is None or df.empty:
        return "No Subject Performance V2 data available for this subject."

    lines = []

    # By subtopic
    by_sub = (
        df.groupby("subtopic_name")["score_pct"]
        .agg(mean="mean", median="median", n="count")
        .round(1)
        .reset_index()
        .sort_values("mean")
    )
    lines.append("ACCURACY BY SUBTOPIC (sorted ascending — weakest first):")
    for _, r in by_sub.iterrows():
        lines.append(f"  {r['subtopic_name']}: Mean={r['mean']}%, Median={r['median']}%, N={r['n']}")

    # Subtopic × Difficulty
    if "difficulty_name" in df.columns and "subtopic_name" in df.columns:
        pivot = (
            df.groupby(["subtopic_name", "difficulty_name"])["score_pct"]
            .mean()
            .round(1)
            .unstack(fill_value=None)
        )
        lines.append("\nACCURACY BY SUBTOPIC × DIFFICULTY (mean %):")
        lines.append(pivot.to_string())

    return "\n".join(lines)


def fmt_error_types(df_sq15, subjects):
    """Struggle score and error breakdown from SQ1.5."""
    df = df_sq15.copy() if df_sq15 is not None else pd.DataFrame()
    if subjects:
        df = df[df["subject_display"].isin(subjects)]
    if df.empty:
        return "No Question Difficulty Analysis data available for this subject."

    lines = []

    # By subtopic — aggregate struggle
    by_sub = (
        df.groupby("subtopic_name")
        .agg(
            avg_struggle=("struggle_score", "mean"),
            avg_wrong_pct=("pct_students_wrong", "mean"),
            avg_correct_pct=("pct_students_correct", "mean"),
            avg_blank_pct=("pct_students_blank", "mean"),
            n_questions=("question_id", "count"),
        )
        .round(1)
        .reset_index()
        .sort_values("avg_struggle", ascending=False)
    )
    lines.append("ERROR TYPES BY SUBTOPIC (sorted by struggle score — higher = harder):")
    for _, r in by_sub.iterrows():
        lines.append(
            f"  {r['subtopic_name']} (n={r['n_questions']} Qs): "
            f"Struggle={r['avg_struggle']}, Wrong={r['avg_wrong_pct']}%, "
            f"Correct={r['avg_correct_pct']}%, Blank={r['avg_blank_pct']}%"
        )

    # Top 5 hardest individual questions
    top_q = df.nlargest(5, "struggle_score")[
        ["subtopic_name", "difficulty_name", "struggle_score",
         "pct_students_wrong", "pct_students_correct", "pct_students_blank"]
    ]
    lines.append("\nTOP 5 HARDEST QUESTIONS:")
    for _, r in top_q.iterrows():
        lines.append(
            f"  [{r['difficulty_name']}] {r['subtopic_name']}: "
            f"Struggle={r['struggle_score']}, Wrong={r['pct_students_wrong']}%, "
            f"Correct={r['pct_students_correct']}%"
        )

    return "\n".join(lines)


def fmt_time_management(df_tm, subjects):
    """Time management metrics. Returns placeholder if data is missing."""
    if df_tm is None or df_tm.empty:
        return (
            "Time Management data: NOT PROVIDED. "
            "Generate placeholder insights acknowledging this data is pending. "
            "Note that this slide covers allotted time vs. used time and median accuracy per subject."
        )
    df = df_tm.copy()
    if subjects and "subject_display" in df.columns:
        df = df[df["subject_display"].isin(subjects)]
    if df.empty:
        return "No Time Management data for this subject."
    return df.to_string(index=False)


def fmt_cross_subject_summary(df_sq13):
    """Cross-subject comparison context for all subject-level slides."""
    if df_sq13 is None or df_sq13.empty:
        return "No cross-subject data available."

    by_subj = (
        df_sq13.groupby("subject_display")
        .agg(
            mean_score=("first_mock_score", "mean"),
            median_score=("first_mock_score", "median"),
            pct_well_prepared=("preparedness_level",
                               lambda x: (x == "[4]Well Prepared").mean() * 100),
            pct_not_prepared=("preparedness_level",
                              lambda x: (x == "[1]Not Yet Prepared").mean() * 100),
            n_students=("user_email", "nunique"),
        )
        .round(1)
        .reset_index()
        .sort_values("median_score", ascending=False)
    )
    lines = ["CROSS-SUBJECT COMPARISON (use selectively for comparisons):"]
    for _, r in by_subj.iterrows():
        lines.append(
            f"  {r['subject_display']} (n={r['n_students']}): "
            f"Median={r['median_score']}%, Mean={r['mean_score']}%, "
            f"Well Prepared={r['pct_well_prepared']}%, Not Prepared={r['pct_not_prepared']}%"
        )
    return "\n".join(lines)


def compute_slide_data(slide, dfs_entity):
    """Route each slide to its metrics function and return formatted text."""
    sq13  = dfs_entity.get("sq13")
    spv2  = dfs_entity.get("spv2")
    sq15  = dfs_entity.get("sq15")
    tm    = dfs_entity.get("tm")
    subjects = SUBJECT_DISPLAY_MAP.get(slide["subject_key"], [])

    stype = slide["slide_type"]

    if stype == "preparedness_overall":
        return fmt_preparedness(sq13)

    elif stype == "accuracy_overall":
        return fmt_accuracy_overall(spv2)

    elif stype == "preparedness_subject":
        return fmt_preparedness(sq13, subjects)

    elif stype == "accuracy_subtopic":
        return fmt_accuracy_subtopic(spv2, subjects)

    elif stype == "error_types":
        return fmt_error_types(sq15, subjects)

    elif stype == "time_management":
        return fmt_time_management(tm, subjects)

    return "No data function defined for this slide type."


# ─── CLAUDE INSIGHT GENERATION ────────────────────────────────────────────────

SYSTEM_PROMPT = """\
You are an educational data analyst generating slide insights for GUIA, a Philippine edtech company \
helping students prepare for the College Entrance Test (CET).

You write structured performance insights for diagnostic reports delivered to partner institutions. \
Each report has 22 slides across 5 sections: Overall, Science, Math, English, and Filipino.

STRICT WRITING RULES:
1. Academic and concise — no filler words or padding.
2. Bullet format: **Bold key term** → insight  (arrow →, NOT an em dash)
3. Bold ONLY the single most critical term per bullet.
4. Heading must be a SPECIFIC, data-backed claim — not a vague label.
   ✓ Good: "Math records the lowest median accuracy at 38.2% across all subjects"
   ✗ Bad:  "Math Performance Overview"
5. Keep insights brief — they must fit on a presentation slide.
6. Cross-subject comparisons: include ONLY where a subject ranks clearly highest or lowest.
7. For Time Management slides: cite MEDIAN accuracy — never mean.
8. English/Filipino slides: compare LP and RC subcomponents where relevant.
9. Flag statistical outliers in callouts when present.

OUTPUT FORMAT — return ONLY valid JSON, no markdown fences:
{
  "heading": "Specific, data-backed claim",
  "subheading": "Context or nuance, or null",
  "bullets": [
    "**Bold Term** → insight",
    "**Bold Term** → insight"
  ],
  "callouts": [
    "Short so-what takeaway"
  ]
}

Rules for the JSON:
- bullets: 1–3 items max
- callouts: 0–3 items max; use empty array [] if none needed
- subheading: use null (not empty string) if not needed
"""


def generate_insights(client, slide, slide_data_text, cross_subject_text):
    """Call Claude API and return parsed JSON insights for one slide."""
    user_msg = (
        f"Generate insights for this slide.\n\n"
        f"SLIDE {slide['id']}: {slide['title']}\n"
        f"SECTION: {slide['section']}\n\n"
        f"SLIDE DATA:\n{slide_data_text}\n\n"
        f"CROSS-SUBJECT CONTEXT (use selectively):\n{cross_subject_text}\n\n"
        f"Return JSON only."
    )

    response = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=1024,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": user_msg}],
    )

    raw = response.content[0].text.strip()
    # Strip accidental markdown code fences
    if raw.startswith("```"):
        parts = raw.split("```")
        raw = parts[1] if len(parts) > 1 else raw
        if raw.startswith("json"):
            raw = raw[4:]
    return json.loads(raw.strip())


# ─── DOCX EXPORT ─────────────────────────────────────────────────────────────

def create_docx(all_results, batch_name):
    """Build and return a BytesIO DOCX with all 22 slides' insights."""
    doc = Document()

    # Cover
    doc.add_heading(f"GUIA Early Diagnostic Report — {batch_name}", level=0)
    doc.add_paragraph(
        f"Automated insight output — copy and paste into Canva slides.\n"
        f"Generated on {pd.Timestamp.now().strftime('%B %d, %Y')}."
    )

    current_section = None

    for result in all_results:
        slide   = result["slide"]
        insights = result["insights"]

        # Section divider
        if slide["section"] != current_section:
            current_section = slide["section"]
            doc.add_page_break()
            doc.add_heading(f"SECTION: {current_section.upper()}", level=1)

        # Slide heading
        doc.add_heading(f"Slide {slide['id']}: {slide['title']}", level=2)

        if "error" in insights:
            doc.add_paragraph(f"⚠ Error: {insights['error']}")
            doc.add_paragraph("─" * 60)
            continue

        # Heading
        p = doc.add_paragraph()
        p.add_run("HEADING:  ").bold = True
        p.add_run(insights.get("heading") or "")

        # Subheading
        if insights.get("subheading"):
            p = doc.add_paragraph()
            p.add_run("SUBHEADING:  ").bold = True
            p.add_run(insights["subheading"])

        # Bullets
        doc.add_paragraph("BULLETS:", style="Heading 3")
        for bullet in insights.get("bullets") or []:
            doc.add_paragraph(bullet, style="List Bullet")

        # Callouts
        callouts = [c for c in (insights.get("callouts") or []) if c]
        if callouts:
            doc.add_paragraph("CALLOUTS:", style="Heading 3")
            for c in callouts:
                doc.add_paragraph(f"⚡ {c}", style="List Bullet")

        doc.add_paragraph("─" * 60)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


# ─── PPTX EXPORT ─────────────────────────────────────────────────────────────

# Brand constants extracted from Pathways Final Diagnostic Review.pptx
_TEAL      = RC(0x1F, 0xAB, 0xCB)   # #1FABCB  accent bar + cover bg
_TEAL_HEAD = RC(0x1A, 0x93, 0xAF)   # #1A93AF  main insight heading
_GREEN     = RC(0x00, 0xBF, 0x63)   # #00BF63  callout boxes
_GRAY_BOX  = RC(0xF5, 0xF2, 0xF2)   # #F5F2F2  chart placeholder bg
_BLACK     = RC(0x00, 0x00, 0x00)
_WHITE     = RC(0xFF, 0xFF, 0xFF)
_GRAY_FT   = RC(0xA6, 0xA6, 0xA6)   # #A6A6A6  footer text

_FONT_H = "Poppins"    # headings (Poppins Bold via bold=True)
_FONT_B = "Work Sans"  # body text

# Slide dimensions — matches the source deck (20" × 11.25", 1920×1080 at 96dpi)
_SW, _SH = 18288000, 10287000  # EMU


def _no_border(shape):
    try:
        shape.line.fill.background()
    except Exception:
        pass


def _add_box(slide, x, y, w, h, fill=None):
    """Borderless rectangle. fill=None → transparent."""
    s = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.RECTANGLE, Emu(x), Emu(y), Emu(w), Emu(h)
    )
    _no_border(s)
    if fill:
        s.fill.solid()
        s.fill.fore_color.rgb = fill
    else:
        s.fill.background()
    return s


def _add_text(slide, x, y, w, h, text,
              font=_FONT_B, size=14, color=_BLACK,
              bold=False, italic=False,
              align=PP_ALIGN.LEFT, wrap=True):
    """Styled text box."""
    tb = slide.shapes.add_textbox(Emu(x), Emu(y), Emu(w), Emu(h))
    tf = tb.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name = font
    run.font.size = PPTXPt(size)
    run.font.color.rgb = color
    run.font.bold = bold
    run.font.italic = italic
    return tf


def _add_bullets(slide, x, y, w, h, bullets):
    """Multi-paragraph bullet text box."""
    tb = slide.shapes.add_textbox(Emu(x), Emu(y), Emu(w), Emu(h))
    tf = tb.text_frame
    tf.word_wrap = True
    for i, bullet in enumerate(bullets[:3]):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_before = PPTXPt(6)
        run = p.add_run()
        run.text = "• " + bullet.replace("**", "")   # strip markdown bold markers
        run.font.name = _FONT_B
        run.font.size = PPTXPt(13)
        run.font.color.rgb = _BLACK


def _build_cover(slide, batch_name):
    _add_box(slide, 0, 0, _SW, _SH, _TEAL)
    _add_text(slide, 914400, _SH // 3, _SW - 1828800, 1828800,
              batch_name, _FONT_H, 44, _WHITE, bold=True, align=PP_ALIGN.CENTER)
    _add_text(slide, 914400, _SH // 3 + 1828800, _SW - 1828800, 548640,
              "Early Diagnostic Review", _FONT_B, 24, _WHITE, align=PP_ALIGN.CENTER)
    _add_text(slide, 914400, _SH - 457200, _SW - 1828800, 274320,
              "Confidential & Proprietary — GUIA Learning Solutions, Inc.",
              _FONT_H, 11, _WHITE, italic=True, align=PP_ALIGN.CENTER)


def _build_divider(slide, section_name):
    _add_box(slide, 0, 0, _SW, _SH, _TEAL_HEAD)
    _add_text(slide, 914400, _SH // 2 - 548640, _SW - 1828800, 1097280,
              section_name.upper(), _FONT_H, 56, _WHITE, bold=True,
              align=PP_ALIGN.CENTER)
    _add_text(slide, 914400, _SH - 457200, _SW - 1828800, 274320,
              "Confidential & Proprietary — GUIA Learning Solutions, Inc.",
              _FONT_H, 11, _WHITE, italic=True, align=PP_ALIGN.CENTER)


def _build_insight_slide(slide, cfg, ins):
    BAR_W  = 228600    # 0.25in  left teal bar
    TX     = 365760    # 0.4in   text left edge
    TW     = 9144000   # 10in    text width → ends at 10.4in
    CX     = 9875520   # 10.8in  chart area left edge
    CW     = 8047872   # 8.8in   chart area width → ends at ~19.6in
    TOP    = 274320    # 0.3in   top margin
    FOOT_Y = 9830400   # 10.75in footer y

    # Left teal accent bar
    _add_box(slide, 0, 0, BAR_W, _SH, _TEAL)

    # Chart placeholder (right column)
    _add_box(slide, CX, TOP, CW, _SH - TOP - 457200, _GRAY_BOX)
    _add_text(slide, CX + 457200, _SH // 2 - 274320, CW - 914400, 548640,
              "[ INSERT CHART ]", _FONT_B, 16, _GRAY_FT, align=PP_ALIGN.CENTER)

    # Section · Slide tag (small teal label at top)
    _add_text(slide, TX, TOP, TW, 228600,
              f"{cfg['section'].upper()}  ·  SLIDE {cfg['id']}",
              _FONT_H, 11, _TEAL, bold=True)

    # Slide title (chart / slide name)
    _add_text(slide, TX, 640080, TW, 457200,
              cfg["title"], _FONT_H, 18, _BLACK, bold=True)

    if "error" in ins:
        _add_text(slide, TX, 1280160, TW, 457200,
                  f"Error: {ins['error']}", _FONT_B, 13, RC(0xCC, 0, 0))
        return

    # Main heading — teal, large (Poppins Bold ~27pt)
    _add_text(slide, TX, 1280160, TW, 1737360,
              ins.get("heading") or "",
              _FONT_H, 27, _TEAL_HEAD, bold=True)

    # Subheading — black bold
    sub = ins.get("subheading") or ""
    bullet_y = 3200400
    if sub:
        _add_text(slide, TX, 3108960, TW, 548640,
                  sub, _FONT_H, 16, _BLACK, bold=True)
        bullet_y = 3749040

    # Bullet points
    bullets = ins.get("bullets") or []
    if bullets:
        _add_bullets(slide, TX, bullet_y, TW, 3200400, bullets)

    # Callout boxes (green rounded)
    callouts = [c for c in (ins.get("callouts") or []) if c]
    if callouts:
        CAL_Y = 7772160   # 8.5in
        CAL_H = 1005840   # 1.1in
        CAL_W = 2834640   # 3.1in
        GAP   = 182880    # 0.2in
        for i, txt in enumerate(callouts[:3]):
            cx = TX + i * (CAL_W + GAP)
            _add_box(slide, cx, CAL_Y, CAL_W, CAL_H, _GREEN)
            _add_text(slide, cx + 91440, CAL_Y + 91440,
                      CAL_W - 182880, CAL_H - 182880,
                      txt, _FONT_B, 11, _WHITE, bold=True)

    # Footer
    _add_text(slide, TX, FOOT_Y, _SW - TX - 182880, 274320,
              "Confidential & Proprietary — GUIA Learning Solutions, Inc.",
              _FONT_H, 11, _GRAY_FT, italic=True)


def create_pptx(all_results, batch_name):
    """Build and return a BytesIO branded PPTX matching GUIA visual identity."""
    prs = PPTXPres()
    prs.slide_width  = Emu(_SW)
    prs.slide_height = Emu(_SH)
    blank = prs.slide_layouts[6]  # truly blank layout

    # Cover slide
    _build_cover(prs.slides.add_slide(blank), batch_name)

    current_section = None
    for result in all_results:
        cfg = result["slide"]
        ins = result["insights"]

        # Section divider slide
        if cfg["section"] != current_section:
            current_section = cfg["section"]
            _build_divider(prs.slides.add_slide(blank), current_section)

        # Insight slide
        _build_insight_slide(prs.slides.add_slide(blank), cfg, ins)

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf


# ─── STREAMLIT UI ─────────────────────────────────────────────────────────────

def main():
    st.set_page_config(
        page_title="GUIA Report Automation",
        page_icon="📊",
        layout="wide",
    )

    st.title("📊 GUIA Report Automation Tool")
    st.caption("MVP · Early Diagnostic Insight Generator · v1.0")

    # ── Sidebar ──────────────────────────────────────────────────────────────
    with st.sidebar:
        st.header("⚙ Settings")
        api_key = st.text_input(
            "Anthropic API Key",
            type="password",
            value=os.getenv("ANTHROPIC_API_KEY", ""),
            help="Your API key is never stored beyond this session.",
        )
        st.divider()
        st.markdown(
            "**Required files:**\n"
            "- SQ1.3 Subject Performance Gains\n"
            "- Subject Performance Version 2\n"
            "- SQ1.5 Question Difficulty Analysis\n\n"
            "**Optional:**\n"
            "- Subject Performance Version 1\n"
            "- Time Management"
        )

    # ── File Uploads ──────────────────────────────────────────────────────────
    st.header("1 · Upload CSV Files")
    c1, c2 = st.columns(2)
    with c1:
        f_sq13 = st.file_uploader(
            "SQ1.3 — Subject Performance Gains ⭐",
            type="csv", key="sq13",
            help="Provides preparedness levels and mock scores for all subjects.",
        )
        f_spv2 = st.file_uploader(
            "Subject Performance Version 2 ⭐",
            type="csv", key="spv2",
            help="Accuracy per student × subtopic × difficulty.",
        )
        f_sq15 = st.file_uploader(
            "SQ1.5 — Question Difficulty Analysis ⭐",
            type="csv", key="sq15",
            help="Struggle score and error rate per question.",
        )
    with c2:
        f_spv1 = st.file_uploader(
            "Subject Performance Version 1 (optional)",
            type="csv", key="spv1",
            help="Overall accuracy per student × subject.",
        )
        f_tm = st.file_uploader(
            "Time Management (optional)",
            type="csv", key="tm",
            help="Allotted vs. used time per subject. Slides 6, 10, 16, 22 will show a placeholder if missing.",
        )

    # ── Batch configuration (populated from SQ1.3) ───────────────────────────
    st.header("2 · Configure Batch")

    batch_name    = st.text_input("Batch Name", placeholder="e.g., PATHWAYS 2026 Batch 22")
    selected_ids  = []

    if f_sq13:
        try:
            f_sq13.seek(0)
            id_df = pd.read_csv(f_sq13)[["entity_id", "school_school_name"]].drop_duplicates().sort_values("entity_id")
            f_sq13.seek(0)
            options = {
                f"{int(r['entity_id'])} — {r['school_school_name']}": int(r["entity_id"])
                for _, r in id_df.iterrows()
            }
            chosen = st.multiselect(
                "Entity IDs to include in this batch",
                options=list(options.keys()),
                help="Select all entity IDs that belong to this batch. Data from all selected IDs will be aggregated.",
            )
            selected_ids = [options[k] for k in chosen]
            if selected_ids:
                st.caption(f"Aggregating data from entity IDs: {selected_ids}")
        except Exception as e:
            st.warning(f"Could not read entity IDs from SQ1.3: {e}")
            selected_ids = []
    else:
        st.info("Upload SQ1.3 above to see available entity IDs.")

    # ── Generate ──────────────────────────────────────────────────────────────
    st.header("3 · Generate Insights")
    go = st.button("🚀 Generate All 22 Slides", type="primary", use_container_width=True)

    if go:
        # Validation
        errors = []
        if not api_key:
            errors.append("API Key is required.")
        if not batch_name:
            errors.append("Batch Name is required.")
        if not selected_ids:
            errors.append("Select at least one Entity ID.")
        if not f_sq13:
            errors.append("SQ1.3 file is required.")
        if not f_spv2:
            errors.append("Subject Performance V2 file is required.")
        if not f_sq15:
            errors.append("SQ1.5 file is required.")
        if errors:
            for e in errors:
                st.error(e)
            st.stop()

        client = anthropic.Anthropic(api_key=api_key)

        # Load & filter data
        with st.spinner("Loading and filtering data…"):
            raw = {
                "sq13": load_csv(f_sq13),
                "spv1": load_csv(f_spv1),
                "spv2": load_csv(f_spv2),
                "sq15": load_csv(f_sq15),
                "tm":   load_csv(f_tm),
            }
            dfs = {k: filter_entity(v, selected_ids) for k, v in raw.items()}

        # Validate we have data after filtering
        sq13_filtered = dfs.get("sq13")
        if sq13_filtered is None or sq13_filtered.empty:
            st.error(
                f"No data found for entity IDs **{selected_ids}** in SQ1.3. "
                "Check that you selected the correct entity IDs."
            )
            st.stop()

        n_students = sq13_filtered["user_email"].nunique() if "user_email" in sq13_filtered.columns else "?"
        st.success(f"Data loaded — Entity IDs {selected_ids} · {n_students} unique students aggregated.")

        # Cross-subject context (computed once)
        cross_context = fmt_cross_subject_summary(sq13_filtered)

        # ── Slide generation loop ─────────────────────────────────────────────
        st.subheader("Generating…")
        progress_bar = st.progress(0)
        status_text  = st.empty()
        all_results  = []

        for i, slide in enumerate(SLIDES):
            status_text.markdown(
                f"**Slide {slide['id']} / {len(SLIDES)}** — {slide['title']}"
            )

            try:
                slide_data = compute_slide_data(slide, dfs)
                insights   = generate_insights(client, slide, slide_data, cross_context)
            except Exception as exc:
                insights = {"error": str(exc)}

            all_results.append({"slide": slide, "insights": insights})
            progress_bar.progress((i + 1) / len(SLIDES))

        status_text.markdown("✅ **All 22 slides generated!**")

        # ── Results preview ───────────────────────────────────────────────────
        st.header("4 · Review Insights")
        current_section = None
        for result in all_results:
            slide    = result["slide"]
            insights = result["insights"]

            if slide["section"] != current_section:
                current_section = slide["section"]
                st.subheader(f"— {current_section} —")

            with st.expander(f"Slide {slide['id']}: {slide['title']}"):
                if "error" in insights:
                    st.error(f"Error: {insights['error']}")
                else:
                    st.markdown(f"**Heading:** {insights.get('heading', '')}")
                    if insights.get("subheading"):
                        st.markdown(f"**Subheading:** {insights['subheading']}")
                    st.markdown("**Bullets:**")
                    for b in insights.get("bullets") or []:
                        st.markdown(f"• {b}")
                    callouts = [c for c in (insights.get("callouts") or []) if c]
                    if callouts:
                        st.markdown("**Callouts:**")
                        for c in callouts:
                            st.info(f"⚡ {c}")

        # ── Download ──────────────────────────────────────────────────────────
        st.header("5 · Download")
        safe_name = batch_name.replace(" ", "_")
        dl1, dl2 = st.columns(2)

        with dl1:
            pptx_buf = create_pptx(all_results, batch_name)
            st.download_button(
                label="📊 Download Branded PPTX",
                data=pptx_buf,
                file_name=f"GUIA_{safe_name}_Insights.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True,
            )
            st.caption(
                "All 22 slides with GUIA branding applied. "
                "Drop chart images into the gray placeholder boxes in PowerPoint or Google Slides."
            )

        with dl2:
            docx_buf = create_docx(all_results, batch_name)
            st.download_button(
                label="📄 Download Plain DOCX",
                data=docx_buf,
                file_name=f"GUIA_{safe_name}_Insights.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )
            st.caption("Plain text version — copy and paste into Canva manually.")


if __name__ == "__main__":
    main()
