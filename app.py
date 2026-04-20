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

# Difficulty levels to include (D3=Easy, D4=Medium, D5=Hard)
VALID_DIFFICULTIES = {"D3", "D4", "D5"}

# Science subtopics to include
SCIENCE_SUBTOPICS = {
    "Astronomy", "Biology", "Chemistry", "Earth Science", "General Science", "Physics"
}


def _assign_preparedness(score):
    """Classify a first_mock_score (%) into a preparedness level."""
    if pd.isna(score):
        return None
    if score >= 76:
        return "[4]Well Prepared"
    elif score >= 51:
        return "[3]Adequately Prepared"
    elif score >= 26:
        return "[2]Partially Prepared"
    else:
        return "[1]Not Yet Prepared"

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

def fmt_preparedness(df_sq13, subjects=None):
    """
    Preparedness distribution computed from first_mock_score thresholds.
    subjects=None → all subjects; otherwise filter.
    Deduplicates to one row per student per subject.
    """
    df = df_sq13.copy() if df_sq13 is not None else pd.DataFrame()
    if subjects:
        df = df[df["subject_display"].isin(subjects)]
    if df.empty:
        return "No preparedness data available."

    # Deduplicate: one row per student per subject
    if "user_email" in df.columns:
        df = df.drop_duplicates(subset=["user_email", "subject_display"])

    # Compute preparedness from first_mock_score using defined thresholds
    df = df.copy()
    df["prep_computed"] = df["first_mock_score"].apply(_assign_preparedness)

    lines = []
    for subj in df["subject_display"].unique():
        sdf = df[df["subject_display"] == subj].dropna(subset=["prep_computed"])
        total = len(sdf)
        if total == 0:
            continue
        dist = sdf["prep_computed"].value_counts()
        lines.append(f"\n{subj} (n={total}):")
        for level in PREPAREDNESS_ORDER:
            cnt = dist.get(level, 0)
            pct = cnt / total * 100
            lines.append(f"  {level}: {cnt} ({pct:.1f}%)")
        scores = sdf["first_mock_score"]
        lines.append(
            f"  Score: Mean={scores.mean():.1f}%, "
            f"Min={scores.min():.1f}%, Max={scores.max():.1f}%, N={total}"
        )

    return "\n".join(lines)


def fmt_accuracy_overall(df_spv2):
    """Overall average accuracy by subject × difficulty from SubPerf V2. D3/D4/D5 only."""
    if df_spv2 is None or df_spv2.empty:
        return "No Subject Performance V2 data available."

    df = df_spv2.copy()
    if "difficulty_name" in df.columns:
        df = df[df["difficulty_name"].isin(VALID_DIFFICULTIES)]
    if df.empty:
        return "No data for difficulty levels D3/D4/D5."

    # Per-subject average accuracy
    by_subj = (
        df.groupby("subject_display")["score_pct"]
        .mean()
        .round(1)
        .reset_index()
        .rename(columns={"score_pct": "avg_accuracy"})
        .sort_values("avg_accuracy", ascending=False)
    )
    lines = ["PER-SUBJECT AVERAGE ACCURACY (D3/D4/D5 combined):"]
    for _, r in by_subj.iterrows():
        lines.append(f"  {r['subject_display']}: {r['avg_accuracy']}%")

    # Subject × Difficulty pivot (mean, D3/D4/D5 only)
    if "difficulty_name" in df.columns:
        pivot = (
            df.groupby(["subject_display", "difficulty_name"])["score_pct"]
            .mean()
            .round(1)
            .unstack(fill_value=None)
        )
        ordered_cols = [c for c in ["D3", "D4", "D5"] if c in pivot.columns]
        pivot = pivot[ordered_cols]
        lines.append("\nAVERAGE ACCURACY BY SUBJECT × DIFFICULTY (D3=Easy, D4=Medium, D5=Hard):")
        lines.append(pivot.to_string())

    return "\n".join(lines)


def _classify_strength(pct):
    if pct >= 75:
        return "STRENGTH"
    elif pct <= 25:
        return "WEAKNESS"
    return "NEUTRAL"


def fmt_accuracy_subtopic(df_spv2, subjects):
    """Average accuracy by subtopic × difficulty. D3/D4/D5 only. Science subtopics filtered."""
    df = df_spv2.copy() if df_spv2 is not None else pd.DataFrame()
    if subjects:
        df = df[df["subject_display"].isin(subjects)]
    if df.empty:
        return "No Subject Performance V2 data available for this subject."

    # Filter to valid difficulty levels
    if "difficulty_name" in df.columns:
        df = df[df["difficulty_name"].isin(VALID_DIFFICULTIES)]

    # Filter Science subtopics
    if subjects and "Science" in subjects and "subtopic_name" in df.columns:
        df = df[df["subtopic_name"].isin(SCIENCE_SUBTOPICS)]

    if df.empty:
        return "No data after applying D3/D4/D5 difficulty filter."

    lines = []

    # By subtopic — average accuracy with strength/weakness classification
    by_sub = (
        df.groupby("subtopic_name")["score_pct"]
        .agg(mean="mean", n="count")
        .round(1)
        .reset_index()
        .sort_values("mean")
    )
    by_sub["category"] = by_sub["mean"].apply(_classify_strength)

    lines.append("AVERAGE ACCURACY BY SUBTOPIC (weakest first, D3=Easy/D4=Medium/D5=Hard only):")
    lines.append("Thresholds: STRENGTH ≥75% | NEUTRAL 26–74% | WEAKNESS ≤25%")
    for _, r in by_sub.iterrows():
        lines.append(f"  [{r['category']}] {r['subtopic_name']}: {r['mean']}% avg (n={r['n']})")

    # Subtopic × Difficulty pivot
    if "difficulty_name" in df.columns and "subtopic_name" in df.columns:
        pivot = (
            df.groupby(["subtopic_name", "difficulty_name"])["score_pct"]
            .mean()
            .round(1)
            .unstack(fill_value=None)
        )
        ordered_cols = [c for c in ["D3", "D4", "D5"] if c in pivot.columns]
        pivot = pivot[ordered_cols]
        lines.append("\nAVERAGE ACCURACY BY SUBTOPIC × DIFFICULTY (D3=Easy, D4=Medium, D5=Hard):")
        lines.append(pivot.to_string())

    return "\n".join(lines)


def fmt_error_types(df_sq15, subjects, df_spv2=None):
    """
    Top 4 common error types by struggle score from SQ1.5.
    Includes question prompts as sample items.
    For ENG/FIL slides, also includes subtopic strength/weakness from SubjV2.
    """
    df = df_sq15.copy() if df_sq15 is not None else pd.DataFrame()
    if subjects:
        df = df[df["subject_display"].isin(subjects)]
    if df.empty:
        return "No Question Difficulty Analysis data available for this subject."

    # Filter to valid difficulty levels and Science subtopics
    if "difficulty_name" in df.columns:
        df = df[df["difficulty_name"].isin(VALID_DIFFICULTIES)]
    if subjects and "Science" in subjects and "subtopic_name" in df.columns:
        df = df[df["subtopic_name"].isin(SCIENCE_SUBTOPICS)]
    if df.empty:
        return "No data after applying difficulty/subtopic filters."

    is_math_science = subjects and all(s in {"Math", "Science"} for s in subjects)
    lines = []

    # Top 4 questions by struggle score
    top4 = df.nlargest(4, "struggle_score")
    lines.append("TOP 4 COMMON ERROR TYPES (highest struggle score = most students struggled):")
    for i, (_, r) in enumerate(top4.iterrows(), 1):
        prompt = str(r.get("question_prompt", "")).strip().replace("\r\n", " ").replace("\n", " ")
        if len(prompt) > 250:
            prompt = prompt[:250] + "…"
        lines.append(f"\n  Error #{i}:")
        lines.append(f"    Subtopic: {r.get('subtopic_name', 'N/A')}")
        lines.append(f"    Difficulty: {r.get('difficulty_name', 'N/A')}")
        lines.append(f"    Struggle Score: {r.get('struggle_score', 'N/A')}")
        lines.append(f"    % Wrong: {r.get('pct_students_wrong', 'N/A')}%")
        lines.append(f"    % Correct: {r.get('pct_students_correct', 'N/A')}%")
        if prompt:
            lines.append(f"    Sample Item: {prompt}")

    # For ENG/FIL: include subtopic strength/weakness context from SubjV2
    if not is_math_science and df_spv2 is not None and not df_spv2.empty:
        spv2 = df_spv2.copy()
        if subjects:
            spv2 = spv2[spv2["subject_display"].isin(subjects)]
        if "difficulty_name" in spv2.columns:
            spv2 = spv2[spv2["difficulty_name"].isin(VALID_DIFFICULTIES)]
        if not spv2.empty:
            sub_avg = (
                spv2.groupby(["subject_display", "subtopic_name"])["score_pct"]
                .mean().round(1).reset_index()
            )
            sub_avg["category"] = sub_avg["score_pct"].apply(_classify_strength)
            lines.append("\nSUBTOPIC STRENGTH/WEAKNESS CONTEXT (from SubjV2, for ENG/FIL table):")
            for subj in sub_avg["subject_display"].unique():
                sdf = sub_avg[sub_avg["subject_display"] == subj].sort_values("score_pct", ascending=False)
                lines.append(f"\n  {subj}:")
                for _, r in sdf.iterrows():
                    lines.append(f"    [{r['category']}] {r['subtopic_name']}: {r['score_pct']}%")

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
    """Cross-subject comparison context. Preparedness computed from first_mock_score. Deduplicated."""
    if df_sq13 is None or df_sq13.empty:
        return "No cross-subject data available."

    df = df_sq13.copy()
    if "user_email" in df.columns:
        df = df.drop_duplicates(subset=["user_email", "subject_display"])
    df["prep_computed"] = df["first_mock_score"].apply(_assign_preparedness)

    by_subj = (
        df.groupby("subject_display")
        .agg(
            mean_score=("first_mock_score", "mean"),
            pct_well=("prep_computed",    lambda x: (x == "[4]Well Prepared").mean() * 100),
            pct_adeq=("prep_computed",    lambda x: (x == "[3]Adequately Prepared").mean() * 100),
            pct_part=("prep_computed",    lambda x: (x == "[2]Partially Prepared").mean() * 100),
            pct_not=("prep_computed",     lambda x: (x == "[1]Not Yet Prepared").mean() * 100),
            n_students=("user_email", "nunique"),
        )
        .round(1)
        .reset_index()
        .sort_values("mean_score", ascending=False)
    )
    lines = ["CROSS-SUBJECT COMPARISON (use selectively):"]
    for _, r in by_subj.iterrows():
        lines.append(
            f"  {r['subject_display']} (n={r['n_students']}): Avg={r['mean_score']}% | "
            f"Well Prepared={r['pct_well']}% | Adequately Prepared={r['pct_adeq']}% | "
            f"Partially Prepared={r['pct_part']}% | Not Yet Prepared={r['pct_not']}%"
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
        return fmt_error_types(sq15, subjects, spv2)

    elif stype == "time_management":
        return fmt_time_management(tm, subjects)

    return "No data function defined for this slide type."


# ─── CLAUDE INSIGHT GENERATION ────────────────────────────────────────────────

SYSTEM_PROMPT = """\
You are an educational data analyst generating slide insights for GUIA, a Philippine edtech company \
helping students prepare for the College Entrance Test (CET).

You write structured performance insights for diagnostic reports delivered to partner institutions. \
Each report has 22 slides across 5 sections: Overall, Science, Math, English, and Filipino.

PREPAREDNESS LEVEL DEFINITIONS (based on first mock exam score):
- Well Prepared: 76%–100%
- Adequately Prepared: 51%–75%
- Partially Prepared: 26%–50%
- Not Yet Prepared: 0%–25% (also called "Underprepared")
- "Moderately prepared" = Partially + Adequately Prepared combined

STRENGTH / WEAKNESS THRESHOLDS (for subtopic accuracy):
- Strength: 75% and above
- Neutral: 26%–74%
- Weakness: 25% and below
- If no clear strength/weakness exists, identify relative strengths/weaknesses (highest vs. lowest).

STRICT WRITING RULES:
1. Academic and concise — no filler words or padding.
2. Bullet format: **Bold key term** → observation. Implication.
   Each bullet must state BOTH the observation AND a 1-sentence implication of what it means for students or the program.
   Example: "**Math** → 50% of students are Partially Prepared. This suggests foundational gaps that may limit performance on higher-difficulty items."
3. Bold ONLY the single most critical term per bullet. Use arrow →, NOT an em dash.
4. HEADING vs SUBHEADING rules:
   - Heading: General finding — e.g., which subject is strongest/weakest, or the overall preparedness trend.
   - Subheading: Specific numbers supporting the heading — e.g., exact averages or distribution percentages.
   - For OVERALL slides: Heading = strongest and weakest subjects + notable difficulty-level pattern. Subheading = exact averages for the strongest and weakest.
   - For PER-SUBJECT slides: Heading = general preparedness or performance observation. Subheading = specific data points.
5. Keep insights brief — they must fit on a presentation slide.
6. Use AVERAGE (mean) accuracy for all accuracy-related observations. NEVER use median for accuracy — EXCEPT on Time Management slides where median accuracy is used.
7. English/Filipino slides: compare LP and RC subcomponents where relevant.
8. Flag statistical outliers in callouts when present.
9. Time Management slides: add an implication for each observation (e.g., more unused time + lower accuracy → premature guessing, not lack of time).

ERROR TYPES FORMAT:
- For Math and Science slides, present each of the top 4 common errors as:
  • Error type name (subtopic/skill)
  • 1–2 sentence description of the error type
  • 1–2 sample items (limit to 1 if the question is long)
  • 1–2 bullet points explaining why students likely struggled
- For English and Filipino slides, format per subcomponent (Reading and Language):
  • State strengths (subtopics ≥75%) and weaknesses (subtopics ≤25%)
  • Present a table: Common error type | 1–2 sentence description | 1 sample question
  • No bullet points needed for ENG/FIL error slides

OUTPUT FORMAT — return ONLY valid JSON, no markdown fences:
{
  "heading": "General finding",
  "subheading": "Specific supporting data, or null",
  "bullets": [
    "**Bold Term** → observation. Implication.",
    "**Bold Term** → observation. Implication."
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
    stype = slide["slide_type"]
    subjects = SUBJECT_DISPLAY_MAP.get(slide["subject_key"], [])

    # Slide-type-specific instructions injected into the user message
    extra = ""
    if stype == "error_types":
        is_ms = all(s in {"Math", "Science"} for s in subjects)
        if is_ms:
            extra = (
                "\nFORMAT REMINDER: Use Math/Science error types format — "
                "for each of the 4 errors: name, description, sample item(s), why students struggled."
            )
        else:
            extra = (
                "\nFORMAT REMINDER: Use English/Filipino error types format — "
                "per subcomponent (Reading, Language): strengths, weaknesses, then a table of "
                "common error type | description | sample question. No bullet points."
            )
    elif stype == "time_management":
        extra = (
            "\nREMINDER: Add a short implication for each time management observation "
            "(e.g., high unused time + low accuracy → premature guessing, not lack of time)."
        )
    elif stype == "preparedness_subject":
        extra = (
            "\nREMINDER: Heading = general observation (e.g., overall preparedness trend). "
            "Subheading = specific data (e.g., exact % per level or average score)."
        )
    elif stype == "accuracy_overall":
        extra = (
            "\nREMINDER: Heading = strongest and weakest subjects + notable difficulty pattern. "
            "Subheading = exact average accuracies for the strongest and weakest subjects."
        )

    user_msg = (
        f"Generate insights for this slide.\n\n"
        f"SLIDE {slide['id']}: {slide['title']}\n"
        f"SECTION: {slide['section']}\n"
        f"SLIDE TYPE: {stype}\n"
        f"{extra}\n\n"
        f"SLIDE DATA:\n{slide_data_text}\n\n"
        f"CROSS-SUBJECT CONTEXT (use selectively for comparisons):\n{cross_subject_text}\n\n"
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
