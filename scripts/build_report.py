"""Build the research report (DOCX) from the results produced by the analysis notebook.

Every number in the text is read from results/results.json or results/tables/*.csv,
so the report always matches the executed notebook.

Usage (from the repository root):
    python scripts/build_report.py            # writes report/Student_Depression_Report.docx
    python scripts/export_pdf.py              # updates the table of contents and exports the PDF (needs MS Word)
"""
import json
import re
from pathlib import Path

import pandas as pd
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

ROOT = Path(__file__).resolve().parents[1]
RES = ROOT / "results"
TAB = RES / "tables"
FIG = RES / "figures"
OUT = ROOT / "report" / "Student_Depression_Report.docx"

R = json.loads((RES / "results.json").read_text())
QS_FRACTION = re.search(r"A fraction ([0-9.]+) of observations",
                        (RES / "model_output" / "logit_health_summary.txt").read_text()).group(1)
T = {p.stem: pd.read_csv(p) for p in TAB.glob("*.csv")}

MID, HIGH = "Mid living expense", "High living expense"
CRAMER = "Cramer's V"
TITLE = "Determinants of Depression among University Students: An Econometric Analysis using Logit and Probit Models"
BODY_FONT = "Cambria"
AUTHORS = ["Atharv Lodhi", "Sourav Choudhary", "Piyush Bisht"]

# ----------------------------------------------------------------------------- helpers to read results


def row(table, label, col="Variable"):
    return T[table].set_index(col).loc[label]


def ame(label, model="logit"):
    return row(f"{model}_marginal_effects_ame", label)


def coef(label, model="logit"):
    return row(f"{model}_coefficients", label)


def odds(label):
    return row("logit_odds_ratios", label)


def rate(characteristic, group):
    t = T["depression_rates_by_group"]
    return t[(t["Characteristic"] == characteristic) & (t["Group"] == group)].iloc[0]


def chi(var):
    return row("chi_square_tests", var)


def ttest(var):
    return row("t_tests", var)


def fmt_p(p):
    return "p < 0.001" if p < 0.001 else f"p = {p:.3f}"


def pp(x):
    """Probability change as percentage points with sign."""
    return f"{100 * x:+.1f} percentage points"


# ----------------------------------------------------------------------------- document helpers
doc = Document()
CAPTIONS = {"Figure": {}, "Table": {}}


def set_cell_border(cell, **edges):
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = tc_pr.find(qn("w:tcBorders"))
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tc_pr.append(borders)
    for edge, val in edges.items():
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), val.get("val", "single"))
        el.set(qn("w:sz"), str(val.get("sz", 8)))
        el.set(qn("w:color"), val.get("color", "000000"))
        borders.append(el)


def add_field(paragraph, instruction):
    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = instruction
    sep = OxmlElement("w:fldChar")
    sep.set(qn("w:fldCharType"), "separate")
    txt = OxmlElement("w:t")
    txt.text = "Right-click and choose Update Field."
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    for el in (begin, instr, sep, txt, end):
        run._r.append(el)
    return run


def setup_styles():
    styles = doc.styles
    normal = styles["Normal"]
    normal.font.name = BODY_FONT
    normal.font.size = Pt(12)
    normal.element.rPr.rFonts.set(qn("w:eastAsia"), BODY_FONT)
    pf = normal.paragraph_format
    pf.space_after = Pt(6)
    pf.line_spacing = 1.3
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    for name, size, before, after in [("Heading 1", 20, 24, 14), ("Heading 2", 15, 18, 8), ("Heading 3", 12.5, 12, 6)]:
        st = styles[name]
        st.font.name = BODY_FONT
        st.font.size = Pt(size)
        st.font.bold = True
        st.font.color.rgb = RGBColor(0, 0, 0)
        st.element.rPr.rFonts.set(qn("w:eastAsia"), BODY_FONT)
        st.element.rPr.rFonts.set(qn("w:ascii"), BODY_FONT)
        st.element.rPr.rFonts.set(qn("w:hAnsi"), BODY_FONT)
        st.paragraph_format.space_before = Pt(before)
        st.paragraph_format.space_after = Pt(after)
        st.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        st.paragraph_format.keep_with_next = True

    for name in ("Figure Caption", "Table Caption"):
        st = styles.add_style(name, 1)
        st.base_style = styles["Normal"]
        st.font.size = Pt(11)
        st.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        st.paragraph_format.space_before = Pt(4)
        st.paragraph_format.space_after = Pt(14)

    code = styles.add_style("Code Block", 1)
    code.base_style = styles["Normal"]
    code.font.name = "Consolas"
    code.font.size = Pt(7.5)
    code.element.rPr.rFonts.set(qn("w:eastAsia"), "Consolas")
    code.paragraph_format.line_spacing = 1.0
    code.paragraph_format.space_after = Pt(0)
    code.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT

    eq = styles.add_style("Equation", 1)
    eq.base_style = styles["Normal"]
    eq.font.italic = True
    eq.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    eq.paragraph_format.space_before = Pt(6)
    eq.paragraph_format.space_after = Pt(10)

    for st in ("TOC 1", "TOC 2", "TOC 3"):
        try:
            styles[st].font.name = BODY_FONT
        except KeyError:
            pass


def para(text="", bold=False, italic=False, align=None, size=None, space_after=None, style=None):
    p = doc.add_paragraph(style=style)
    if text:
        add_rich(p, text, bold=bold, italic=italic, size=size)
    if align is not None:
        p.alignment = align
    if space_after is not None:
        p.paragraph_format.space_after = Pt(space_after)
    return p


def add_rich(p, text, bold=False, italic=False, size=None):
    """Supports **bold** spans inside text."""
    parts = text.split("**")
    for i, part in enumerate(parts):
        if not part:
            continue
        run = p.add_run(part)
        run.bold = bold or (i % 2 == 1)
        run.italic = italic
        if size:
            run.font.size = Pt(size)
    return p


def note(text):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.font.size = Pt(10)
    return p


def bullets(items):
    for item in items:
        p = doc.add_paragraph(style="List Bullet")
        add_rich(p, item)
        p.paragraph_format.space_after = Pt(3)


def numbered(items):
    for i, item in enumerate(items, 1):
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.9)
        p.paragraph_format.first_line_indent = Cm(-0.6)
        p.paragraph_format.space_after = Pt(3)
        add_rich(p, f"{i}.  {item}")


CHAPTER = {"n": 0}
SECTION = {"n": 0, "sub": 0}


def chapter(title, numbered_chapter=True):
    p = doc.add_paragraph(style="Heading 1")
    if numbered_chapter:
        CHAPTER["n"] += 1
        SECTION["n"] = 0
        CAPTIONS["Figure"][CHAPTER["n"]] = 0
        CAPTIONS["Table"][CHAPTER["n"]] = 0
        p.add_run(f"Chapter {CHAPTER['n']}")
        p.add_run().add_break()
        p.add_run(title)
    else:
        p.add_run(title)
        CHAPTER["n"] = title[0]  # e.g. "A" for Appendix captions
        CAPTIONS["Figure"][CHAPTER["n"]] = 0
        CAPTIONS["Table"][CHAPTER["n"]] = 0
    return p


def section(title):
    SECTION["n"] += 1
    SECTION["sub"] = 0
    return doc.add_heading(f"{CHAPTER['n']}.{SECTION['n']}   {title}", level=2)


def subsection(title):
    SECTION["sub"] += 1
    return doc.add_heading(f"{CHAPTER['n']}.{SECTION['n']}.{SECTION['sub']}   {title}", level=3)


def caption(kind, text):
    CAPTIONS[kind][CHAPTER["n"]] = CAPTIONS[kind].get(CHAPTER["n"], 0) + 1
    label = f"{kind} {CHAPTER['n']}.{CAPTIONS[kind][CHAPTER['n']]}"
    p = doc.add_paragraph(style=f"{kind} Caption")
    p.add_run(f"{label}: ").bold = True
    p.add_run(text)
    return label


def next_label(kind):
    return f"{kind} {CHAPTER['n']}.{CAPTIONS[kind].get(CHAPTER['n'], 0) + 1}"


def figure(name, text, width_cm=15.5):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.keep_with_next = True
    p.add_run().add_picture(str(FIG / f"{name}.png"), width=Cm(width_cm))
    return caption("Figure", text)


def table(df, text, col_widths=None, font_size=9.5, decimals=3, first_col_bold=False):
    df = df.copy()
    for c in df.columns:
        if not pd.api.types.is_float_dtype(df[c]):
            continue
        is_p = any(k in c for k in ("P>|z|", "p-value", " p", "Prob >"))
        df[c] = df[c].map(lambda v: "" if pd.isna(v) else
                          ("<0.001" if is_p and v < 0.001 else f"{v:.{decimals}f}"))
    df = df.fillna("").astype(str).replace("nan", "")
    t = doc.add_table(rows=1, cols=len(df.columns))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for j, c in enumerate(df.columns):
        cell = t.rows[0].cells[j]
        cell.text = ""
        run = cell.paragraphs[0].add_run(str(c))
        run.bold = True
        run.font.size = Pt(font_size)
    for _, r in df.iterrows():
        cells = t.add_row().cells
        for j, v in enumerate(r):
            cells[j].text = ""
            run = cells[j].paragraphs[0].add_run(v)
            run.font.size = Pt(font_size)
            if first_col_bold and j == 0:
                run.bold = True
    for i, tr in enumerate(t.rows):
        tr_pr = tr._tr.get_or_add_trPr()
        cant = OxmlElement("w:cantSplit")
        tr_pr.append(cant)
        for j, cell in enumerate(tr.cells):
            pf = cell.paragraphs[0].paragraph_format
            pf.space_after = Pt(1)
            pf.space_before = Pt(1)
            pf.line_spacing = 1.0
            pf.keep_with_next = True
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT if j == 0 else WD_ALIGN_PARAGRAPH.CENTER
            edges = {}
            if i == 0:
                edges = {"top": {"sz": 12}, "bottom": {"sz": 6}}
            if i == len(t.rows) - 1:
                edges["bottom"] = {"sz": 12}
            if edges:
                set_cell_border(cell, **edges)
    if col_widths:
        t.autofit = False
        tbl_pr = t._tbl.tblPr
        layout = OxmlElement("w:tblLayout")
        layout.set(qn("w:type"), "fixed")
        tbl_pr.append(layout)
        margins = OxmlElement("w:tblCellMar")
        for side in ("left", "right"):
            m = OxmlElement(f"w:{side}")
            m.set(qn("w:w"), str(int(Cm(0.06).twips)))
            m.set(qn("w:type"), "dxa")
            margins.append(m)
        tbl_pr.append(margins)
        for j, (grid_col, w) in enumerate(zip(t._tbl.tblGrid.findall(qn("w:gridCol")), col_widths)):
            grid_col.set(qn("w:w"), str(int(Cm(w).twips)))
            for tr in t.rows:
                tr.cells[j].width = Cm(w)
    return caption("Table", text)


def page_break():
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


def code_block(text):
    for line in text.rstrip().splitlines():
        p = doc.add_paragraph(style="Code Block")
        p.add_run(line if line.strip() else " ")
    doc.add_paragraph()


def set_page_numbering(sect, fmt, start=None):
    sect_pr = sect._sectPr
    pg = sect_pr.find(qn("w:pgNumType"))
    if pg is None:
        pg = OxmlElement("w:pgNumType")
        sect_pr.append(pg)
    pg.set(qn("w:fmt"), fmt)
    if start is not None:
        pg.set(qn("w:start"), str(start))


def footer_page_number(sect):
    sect.footer.is_linked_to_previous = False
    p = sect.footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_field(p, "PAGE")


# ============================================================================= build
setup_styles()
sec0 = doc.sections[0]
for s in (sec0,):
    s.page_height, s.page_width = Cm(29.7), Cm(21.0)
    s.left_margin = s.right_margin = Cm(2.5)
    s.top_margin = s.bottom_margin = Cm(2.5)
sec0.different_first_page_header_footer = True
set_page_numbering(sec0, "lowerRoman", 1)
footer_page_number(sec0)

# ----------------------------------------------------------------------------- title page
para("", space_after=30)
para(TITLE, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, size=20, space_after=30)
para("Report submitted to", align=WD_ALIGN_PARAGRAPH.CENTER, size=13, space_after=4)
para("Dr. Prantik Bagchi", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, size=14, space_after=2)
para("Department of Humanities and Social Sciences", align=WD_ALIGN_PARAGRAPH.CENTER, size=12, space_after=2)
para("Indian Institute of Technology Kharagpur", align=WD_ALIGN_PARAGRAPH.CENTER, size=12, space_after=30)
para("Submitted by", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, size=13, space_after=6)
for name in AUTHORS:
    para(name, align=WD_ALIGN_PARAGRAPH.CENTER, size=12, space_after=2)
logo = para("", align=WD_ALIGN_PARAGRAPH.CENTER, space_after=30)
logo.paragraph_format.space_before = Pt(36)
logo.add_run().add_picture(str(ROOT / "assets" / "iitkgp_logo.png"), width=Cm(4.6))
para("Department of Humanities and Social Sciences", align=WD_ALIGN_PARAGRAPH.CENTER, size=12, space_after=2)
para("Indian Institute of Technology Kharagpur", align=WD_ALIGN_PARAGRAPH.CENTER, size=12, space_after=2)
para("Kharagpur, West Bengal 721302", align=WD_ALIGN_PARAGRAPH.CENTER, size=12, space_after=2)
para("© 2026. All rights reserved", align=WD_ALIGN_PARAGRAPH.CENTER, size=11)
page_break()

# ----------------------------------------------------------------------------- abstract
single_ame, rich_ame, low_ame = ame("Single"), ame("Rich class"), ame("Lower-middle class")
mid_ame, high_ame, male_ame = ame(MID), ame(HIGH), ame("Male")
cgpa_ame, friend_ame = ame("CGPA"), ame("Has friend circle")

p = doc.add_paragraph(style="Heading 1")
p.add_run("Abstract")
para(
    "Depression is one of the most common mental health problems among university students, yet the personal, "
    "social and economic factors that separate students who are depressed from those who are not remain unclear. "
    "This study examines which student characteristics are associated with a higher or lower probability of "
    "depression, with special attention to relationship status, family background, living costs, gender, academic "
    "performance and social support."
)
para(
    f"Using a quantitative cross-sectional design, survey data on {R['n']} university students aged "
    f"{R['age_min']}–{R['age_max']} with a CGPA above 6.00 on a 10-point scale were analysed, of whom "
    f"{R['n_depressed']} ({R['pct_depressed']}%) reported depression. After bivariate chi-square and t-tests, a Binary Logistic Regression (logit) model was estimated, "
    "followed by a Binary Probit model, average marginal effects, goodness-of-fit diagnostics and a set of robustness checks."
)
para(
    f"The results show that being **single** is the strongest risk factor: single students are "
    f"{100 * single_ame['dy/dx']:.1f} percentage points more likely to be depressed than students in a relationship "
    f"({fmt_p(single_ame['P>|z|'])}). Compared with students whose monthly **living expense** is low, students with a "
    f"mid expense ({pp(mid_ame['dy/dx'])}, {fmt_p(mid_ame['P>|z|'])}) or a high expense ({pp(high_ame['dy/dx'])}, "
    f"{fmt_p(high_ame['P>|z|'])}) are more likely to be depressed. **Male** students are "
    f"{abs(100 * male_ame['dy/dx']):.1f} percentage points less likely to be depressed than female students "
    f"({fmt_p(male_ame['P>|z|'])}). Students from **rich families** are less likely to be depressed, but only at the "
    f"10% level ({pp(rich_ame['dy/dx'])}, {fmt_p(rich_ame['P>|z|'])}). Surprisingly, CGPA, semester, age, living "
    "arrangement and having a friend circle show no significant association once other factors are held constant."
)
para(
    f"The logit model is jointly significant (LR χ²({R['logit_lr_df']}) = {R['logit_lr_chi2']:.2f}, "
    f"{fmt_p(R['logit_lr_p'])}), fits the data well (Hosmer–Lemeshow p = {R['logit_hl_p']:.3f}) and discriminates "
    f"reasonably between depressed and non-depressed students (AUC = {R['logit_auc']:.3f}). The probit model gives "
    "almost identical marginal effects, and a linear probability model confirms the relationship and gender effects. "
    "The study concludes that university mental "
    "health support should focus on social and relational wellbeing and financial pressure, rather than on academic "
    "performance alone."
)
para("**Keywords:** Depression, University students, Binary logit, Probit, Marginal effects, Relationship status, Socioeconomic status")
page_break()

# ----------------------------------------------------------------------------- table of contents
p = doc.add_paragraph(style="Heading 1")
p.add_run("Table of Contents")
add_field(doc.add_paragraph(), 'TOC \\o "1-3" \\h \\z \\u')
page_break()
p = doc.add_paragraph(style="Heading 1")
p.add_run("List of Figures")
add_field(doc.add_paragraph(), 'TOC \\h \\z \\t "Figure Caption,1"')
p = doc.add_paragraph(style="Heading 2")
p.add_run("List of Tables")
add_field(doc.add_paragraph(), 'TOC \\h \\z \\t "Table Caption,1"')

# body section with arabic page numbers
body = doc.add_section(WD_SECTION.NEW_PAGE)
body.different_first_page_header_footer = False
set_page_numbering(body, "decimal", 1)
footer_page_number(body)

# ============================================================================= Chapter 1
chapter("Introduction")
section("Background of the Study")
para(
    "University life is a period of rapid change. Within a few years, students move away from home, build new "
    "friendships and romantic relationships, manage their own money, and face continuous academic evaluation. "
    "For many students these transitions are exciting, but for a large share they are also a source of persistent "
    "stress. Depression — a condition marked by low mood, loss of interest, fatigue and feelings of worthlessness — is "
    "among the most frequent mental health problems in this age group. A systematic review of university students "
    "worldwide reported an average depression prevalence of around 30 percent, far above the rate in the general "
    "population (Ibrahim et al., 2013)."
)
para(
    "Depression among students is not only a personal health issue. It lowers academic productivity, increases the "
    "risk of dropping out, and has lasting effects on later employment and wellbeing (Hysenbegasi et al., 2005; "
    "Eisenberg et al., 2007). Understanding which students are most at risk therefore matters to universities, "
    "families and policymakers who have to decide where to direct limited counselling and support resources."
)
para(
    "The factors linked to depression are many and interrelated. Psychological theories stress the role of stressful "
    "life events acting on individual vulnerability (Monroe & Simons, 1991). Social theories emphasise the protective "
    "role of close relationships and social support (Cohen & Wills, 1985). Economic research shows a clear "
    "socioeconomic gradient in depression (Lorant et al., 2003), and a large body of work documents higher depression "
    "rates among women (Nolen-Hoeksema, 1987). This study brings these strands together and tests them on a single, "
    "well-defined group of university students."
)

section("Problem Statement")
para(
    "Although the high prevalence of depression among university students is well documented, there is less clarity "
    "about which student characteristics actually matter once they are considered together. Many studies examine one "
    "factor at a time — for example, academic performance or family income — and therefore cannot tell whether an "
    "apparent effect is real or simply reflects other, correlated characteristics. For instance, if single students "
    "are also more likely to live away from family, a simple comparison cannot separate the effect of relationship "
    "status from that of living arrangement."
)
para(
    "There is also a common assumption, among students and institutions alike, that poor grades are a main driver of "
    "depression. If this assumption is wrong, support programmes that focus mainly on academic help may miss the "
    "students who need support most. This study addresses these gaps by estimating multivariate probability models "
    "that isolate the contribution of each factor while holding the others constant."
)

section("Research Objectives")
para("The main objectives of this study are:")
numbered([
    "To measure the prevalence of depression in a sample of university students aged 18–24.",
    "To examine how depression rates differ across gender, family background, relationship status, living style, "
    "living expense, social support (friend circle) and self-reported health.",
    "To estimate the effect of each factor on the probability of depression using binary logit and probit models, "
    "holding all other factors constant.",
    "To express these effects as marginal effects (changes in probability) that can be interpreted directly.",
    "To test the robustness of the findings across alternative model specifications and across male and female students.",
    "To derive practical recommendations for university mental health support.",
])

section("Research Questions")
numbered([
    "What share of university students in the sample report depression?",
    "Is academic performance (CGPA) associated with the probability of depression?",
    "Does relationship status affect the probability of depression?",
    "Do family socioeconomic background and the level of monthly living expense influence depression?",
    "Are female students more likely to be depressed than male students?",
    "Do social support (having a friend circle) and living with family protect students against depression?",
])

section("Research Hypotheses")
para("Based on the literature reviewed in Chapter 2, the following hypotheses are tested:")
hyp = T["hypothesis_tests"]
bullets([f"**{r.Hypothesis}:** {r.Statement}" for r in hyp.itertuples()])

section("Significance of the Study")
para(
    "The study contributes in three ways. First, it provides evidence on the relative importance of social, economic "
    "and academic factors, measured on the same students and estimated jointly. Second, by reporting marginal effects "
    "rather than only coefficients, it translates statistical results into easily interpreted changes in the "
    "probability of depression. Third, it offers practical guidance for universities on which groups of students to "
    "prioritise for screening and support, and highlights that common assumptions — such as the central role of "
    "grades — may not hold."
)

section("Scope and Limitations")
para(
    f"The study is based on {R['n']} students aged {R['age_min']}–{R['age_max']} with a CGPA above 6.00 (10-point "
    "scale). Depression is "
    "measured as a self-reported yes/no indicator rather than a clinical diagnosis. The data are cross-sectional, so "
    "the estimated relationships describe associations rather than proven causal effects. These limitations are "
    "discussed further in Section 5.5."
)

section("Organization of the Paper")
para(
    "The remainder of the report is organised as follows. Chapter 2 reviews the theoretical and empirical literature. "
    "Chapter 3 describes the data, the transformation of variables and the econometric methodology. Chapter 4 presents "
    "the descriptive analysis, bivariate tests, logit and probit estimates, marginal effects, diagnostics and "
    "robustness checks. Chapter 5 summarises the findings, discusses their implications, and concludes. The appendix "
    "reports the full model output, and the Stata code used to replicate the analysis is provided with the report."
)

# ============================================================================= Chapter 2
# (no page break: chapters run on continuously)
chapter("Literature Review")
section("Introduction to the Theoretical Framework")
para(
    "Depression has no single cause. Modern research views it as the result of an interaction between individual "
    "vulnerability and environmental stressors, moderated by the resources a person can draw on. This chapter "
    "reviews the main theories that guide the choice of variables in this study and summarises the empirical evidence "
    "on each group of factors: academic, social, economic, demographic and health-related."
)

section("Theoretical Underpinnings")
subsection("The Diathesis–Stress Model")
para(
    "The diathesis–stress model proposes that depression develops when stressful life events act on a pre-existing "
    "vulnerability (Monroe & Simons, 1991). University students face a dense cluster of potential stressors — exams, "
    "financial obligations, relationship changes and living away from home — so the model predicts that students "
    "exposed to more of these stressors, or with fewer buffers, will show higher rates of depression. Beck's (1967) "
    "cognitive theory adds that negative interpretations of such events (for example, viewing a poor grade or a "
    "breakup as a personal failure) turn stress into depressive symptoms."
)
subsection("Social Support and the Stress-Buffering Hypothesis")
para(
    "Cohen and Wills (1985) showed that social support protects mental health mainly by 'buffering' the impact of "
    "stress: people with close, supportive relationships cope better with the same level of adversity. Kawachi and "
    "Berkman (2001) further document that social ties — friendships, partners and family — reduce the risk of "
    "depression through emotional support, a sense of belonging and practical help. For students, friend circles, "
    "romantic partners and living with family are the most direct sources of this support, which motivates "
    "hypotheses H2, H3 and H6."
)

section("Academic Pressure and Performance")
para(
    "Academic stress is widely cited as a driver of student depression. Poor results can threaten future career "
    "plans and self-esteem, while depression in turn reduces concentration and productivity (Hysenbegasi et al., "
    "2005). Eisenberg et al. (2007) found that depression is associated with lower grade point averages. However, the "
    "direction of the relationship is unclear — depression may lower grades rather than grades causing depression — "
    "and several studies find only weak links once social and economic variables are included. Hypothesis H1 tests "
    "whether CGPA retains an independent association in this sample."
)

section("Socioeconomic Status and Financial Strain")
para(
    "A meta-analysis by Lorant et al. (2003) found that people with lower socioeconomic status have significantly "
    "higher odds of depression, with financial strain being one of the main channels. Richardson et al. (2013) show "
    "that debt and money worries are strongly linked to poor mental health, including among students. For students, "
    "family background shapes both the money available and the expectations placed on them, while the level of "
    "monthly living expense is a direct measure of financial pressure. These ideas motivate hypotheses H4a, H4b and H5."
)

section("Relationships and Loneliness")
para(
    "Close relationships are among the strongest protective factors against depression. Being in a stable "
    "relationship provides companionship and emotional support, while being single can be associated with loneliness, "
    "particularly in early adulthood when peers are forming couples. Whitton and Whisman (2010) show that the quality "
    "and stability of romantic relationships are closely tied to depressive symptoms, and Bulloch et al. (2009) find "
    "that the link between relationship disruption and depression runs in both directions. Hypothesis H3 tests whether single "
    "students face a higher risk."
)

section("Gender Differences in Depression")
para(
    "One of the most consistent findings in psychiatric epidemiology is that women experience depression roughly twice "
    "as often as men (Nolen-Hoeksema, 1987; Kessler, 2003). Explanations include differences in stress exposure, "
    "coping styles — women are more likely to ruminate, men to distract themselves — and hormonal factors. Among "
    "students, the gap is often smaller but usually still present (Eisenberg et al., 2007). Hypothesis H7 tests "
    "whether female students in this sample are more likely to be depressed after other factors are controlled for."
)

section("Physical Health and Depression")
para(
    "Physical and mental health are closely intertwined. Using World Health Survey data from 60 countries, Moussavi et "
    "al. (2007) showed that chronic physical conditions sharply increase the probability of depression, and that "
    "depression worsens health outcomes more than many chronic diseases. Self-reported health is therefore expected to "
    "be strongly related to depression. However, when self-reported health partly reflects the same symptoms as "
    "depression, it becomes difficult to treat it as an independent cause — an issue that turns out to be important "
    "in this study (Section 3.2)."
)

section("Conclusion of Literature Review")
para(
    "The literature suggests that depression among students is shaped by academic pressure, social support, "
    "relationships, financial conditions, gender and health. However, most studies examine these factors separately or "
    "in very different populations. By estimating them jointly in one sample, this study aims to identify which "
    "factors carry independent explanatory power and which only appear important when considered in isolation."
)

# ============================================================================= Chapter 3
# (no page break: chapters run on continuously)
chapter("Research Methodology")
section("Data Overview")
para(
    f"The study uses survey data on {R['n']} university students aged {R['age_min']}–{R['age_max']} "
    f"({R['n_male']} male and {R['n_female']} female). The 18–24 range is the typical undergraduate age range. CGPA "
    "is measured on a 10-point scale, and every student in the sample has a CGPA above 6.00."
)
para(
    "For each student, the survey records gender, age, family background, relationship status, living style, current "
    "semester, CGPA, whether the student has a friend circle, self-reported health condition, monthly living expense "
    "excluding study costs (in three bands: low, mid, high), and whether the student is depressed. "
    f"{next_label('Table')} shows the first five records of the raw data."
)
prev = T["raw_data_preview"].drop(columns="Student_ID").rename(columns={
    "Family_background": "Family background", "Relationship_status": "Relationship",
    "Living_style": "Living", "Current_semester": "Semester", "friend_circle": "Friends",
    "Health_condition": "Health", "Living_expense": "Expense"})
table(prev, "Raw data (first five records)", font_size=7, decimals=2,
      col_widths=[1.3, 0.75, 1.75, 2.0, 1.4, 1.4, 1.0, 1.2, 1.4, 1.3, 1.5])

section("Data Transformation")
para("Several preprocessing steps were carried out to make the data accurate, consistent and ready for modelling:")
para("**1. Data quality checks.** "
     f"The data contain {R['missing_total']} missing values and {R['duplicates']} duplicate records, so no imputation or "
     "deletion was needed. All variables were checked for valid ranges (for example, CGPA between 6 and 10).")
para("**2. Encoding of binary variables.** Text answers were converted into 0/1 indicators: depression (Yes = 1), "
     "gender (Male = 1), living style (with family = 1) and friend circle (Yes = 1).")
para("**3. Dummy variables for family background.** Family background has three categories. Two dummy variables were "
     "created — lower-middle class and rich class — with middle class (the largest group) as the reference category.")
para("**4. Merging a very small category.** The 'Complicated' relationship group contains only "
     f"{int(rate('Relationship status', 'Complicated')['Students'])} students, none of whom is depressed, which would "
     "make its coefficient impossible to estimate (perfect prediction). It was merged with 'In a relationship', so "
     "relationship status becomes **Single (1) vs Not single (0)**.")
para("**5. Living expense bands.** Monthly living expense (excluding study costs) is grouped into three bands cut at "
     "the rounded sample thirds: **Low** (below 10,000), **Mid** (10,000–14,999) and **High** (15,000 or more). Two "
     "dummy variables (mid, high) enter the models, with low expense as the reference category.")
para(
    "**6. Treatment of self-reported health condition.** Health condition is almost perfectly aligned with depression "
    f"in the data: {R['health_critical_rate']:.0f}% of students in 'critical' health and {R['health_good_rate']:.0f}% "
    f"of students in 'good' health are depressed (Section 4.2.3). Including it in the main model would cause "
    "quasi-complete separation, and it very likely captures the same underlying condition as the outcome rather than "
    "an independent cause. It is therefore analysed descriptively and entered only in one robustness model "
    "(Section 4.12), as a single 'poor health' indicator."
)
para(f"After these steps the final dataset has {R['n']} observations and the variables described in "
     f"{next_label('Table')}.")

section("Data and Types of Variables")
dic = T["variable_dictionary"].rename(columns={"Variable": "Variable name"})
table(dic[["Variable name", "Description and coding", "Data type"]], "Description of variables",
      col_widths=[3.2, 9.0, 3.8], font_size=9)

section("Data Analysis")
para(
    "The analysis proceeds in four steps. First, descriptive statistics and charts summarise the sample and compare "
    "depression rates across groups. Second, bivariate tests check whether each factor is associated with depression "
    "on its own: Pearson chi-square tests (with Cramér's V as a measure of strength) for categorical variables, and "
    "Welch t-tests and Mann–Whitney U tests for numerical variables. Third, multivariate binary logit and probit "
    "models estimate the effect of each factor while holding the others constant. Fourth, diagnostics and robustness "
    "checks assess the reliability of the results."
)
subsection("Graphs Overview")
para(
    f"{next_label('Figure')} shows the distribution of the dependent variable. Of the {R['n']} students, "
    f"{R['n_depressed']} ({R['pct_depressed']}%) report depression and {R['n_not_depressed']} "
    f"({100 - R['pct_depressed']:.1f}%) do not. The outcome is therefore reasonably balanced: depression is common "
    "enough to be studied with a binary choice model, and there are enough depressed students (events) relative to "
    "the number of explanatory variables for stable estimation."
)
figure("fig01_depression_distribution", "Distribution of the dependent variable (depression status)", width_cm=10)

section("Econometric Model Specification")
para(
    "Because the dependent variable is binary (depressed = 1, not depressed = 0), ordinary least squares is not ideal: "
    "its predicted values can fall outside the 0–1 range and its error term is heteroskedastic by construction. The "
    "study therefore uses binary choice models, which model the probability of depression directly (Wooldridge, 2010; "
    "Long & Freese, 2014)."
)
para("**Binary logit model.** The probability that student i is depressed is modelled as:")
para("Pr(Depressionᵢ = 1 | xᵢ) = Λ(xᵢ′β) = exp(xᵢ′β) / [1 + exp(xᵢ′β)]", style="Equation")
para("where the vector xᵢ′β is:")
para("β₀ + β₁Maleᵢ + β₂Ageᵢ + β₃CGPAᵢ + β₄Semesterᵢ + β₅MidExpenseᵢ + β₆HighExpenseᵢ + β₇LowerMidᵢ + β₈Richᵢ + "
     "β₉Singleᵢ + β₁₀LiveFamilyᵢ + β₁₁FriendCircleᵢ", style="Equation")
para(
    "The coefficients are estimated by maximum likelihood. A positive coefficient means the factor raises the "
    "probability of depression. Exponentiated coefficients, exp(β), are odds ratios: an odds ratio above 1 raises "
    "the odds of depression, and below 1 lowers them."
)
para("**Binary probit model.** As an alternative, the probit model replaces the logistic function with the standard "
     "normal cumulative distribution function Φ:")
para("Pr(Depressionᵢ = 1 | xᵢ) = Φ(xᵢ′β)", style="Equation")
para(
    "**Marginal effects.** Logit and probit coefficients are not directly comparable, so the main results are reported "
    "as average marginal effects (AME): the average change in the probability of depression when a continuous variable "
    "rises by one unit, or when a 0/1 variable changes from 0 to 1, computed for every student and then averaged "
    "(equivalent to Stata's margins, dydx(*) command)."
)
para(
    "**Diagnostics.** Model fit is assessed with the likelihood-ratio (LR) test, McFadden's pseudo-R², the AIC and BIC, "
    "classification tables, the area under the ROC curve (AUC) and the Hosmer–Lemeshow test (Hosmer et al., 2013). "
    "Multicollinearity is checked using variance inflation factors (VIF). Robustness is assessed with robust standard "
    "errors, a linear probability model (LPM), a model that adds health status, and separate models for male and "
    "female students."
)
para(
    "All estimation was carried out in Python using the pandas, SciPy and statsmodels libraries (Seabold & Perktold, "
    "2010). An equivalent Stata do-file that reproduces every step is provided with this report (StataCorp, 2021)."
)

section("Ethical Considerations")
para(
    "The study uses anonymised secondary survey data. The data contain no names, contact details or other personal "
    "identifiers; each student is identified only by a sequential number. Because the topic is mental health, results are reported only at group level, and no attempt is made "
    "to identify or profile individual students."
)

# ============================================================================= Chapter 4
# (no page break: chapters run on continuously)
chapter("Empirical Results")
section("Introduction")
para(
    "This chapter presents the results. It begins with a descriptive analysis of depression across student groups, "
    "followed by bivariate tests, the multicollinearity check, the logit and probit estimates with their marginal "
    "effects, model diagnostics, the hypothesis summary and the robustness checks."
)

section("Descriptive Analysis of Student Characteristics")
subsection("Prevalence of Depression")
para(
    f"Depression is common in the sample: {R['pct_depressed']}% of students — almost three in ten — report being "
    "depressed. This is in line with the international estimates of around 30% reported by Ibrahim et al. (2013), "
    "and confirms that student depression is a widespread rather than a marginal problem."
)
subsection("Social and Demographic Characteristics")
s_rate, r_rate, c_rate = rate("Relationship status", "Single"), rate("Relationship status", "In a Relationship"), rate("Relationship status", "Complicated")
mid, low, rich = rate("Family background", "Middle class"), rate("Family background", "Lower Middle class"), rate("Family background", "Rich class")
bach, fam = rate("Living style", "Bachelor"), rate("Living style", "Family")
fem, mal = rate("Gender", "Female"), rate("Gender", "Male")
fr_y, fr_n = rate("Friend circle", "Yes"), rate("Friend circle", "No")
ex_l, ex_m, ex_h = rate("Living expense", "Low"), rate("Living expense", "Mid"), rate("Living expense", "High")
para(f"{next_label('Figure')} compares depression rates across the main social and demographic groups. Several patterns stand out:")
bullets([
    f"**Relationship status** shows the largest gap: {s_rate['Depression rate (%)']:.1f}% of single students are depressed, "
    f"compared with only {r_rate['Depression rate (%)']:.1f}% of students in a relationship and "
    f"{c_rate['Depression rate (%)']:.0f}% of the {int(c_rate['Students'])} students in a complicated relationship.",
    f"**Family background:** students from middle-class families have the highest rate ({mid['Depression rate (%)']:.1f}%), "
    f"well above lower-middle class ({low['Depression rate (%)']:.1f}%) and rich ({rich['Depression rate (%)']:.1f}%) students.",
    f"**Living expense:** students with a low monthly living expense are less often depressed "
    f"({ex_l['Depression rate (%)']:.1f}%) than those with a mid ({ex_m['Depression rate (%)']:.1f}%) or high "
    f"({ex_h['Depression rate (%)']:.1f}%) expense.",
    f"**Living style:** students living as bachelors (hostels, messes or shared housing) are more often depressed "
    f"({bach['Depression rate (%)']:.1f}%) than those living with family ({fam['Depression rate (%)']:.1f}%).",
    f"**Gender and friend circle** show small raw differences: {mal['Depression rate (%)']:.1f}% of male and "
    f"{fem['Depression rate (%)']:.1f}% of female students are depressed, and having a friend circle makes no "
    f"difference ({fr_y['Depression rate (%)']:.1f}% vs {fr_n['Depression rate (%)']:.1f}%).",
])
figure("fig02_rates_by_group", "Depression rate by student characteristics")
para(
    "These raw comparisons do not control for other factors. For example, single students are more often male and less "
    "often live with family, and living expense overlaps with family background, so part of the gap between groups may "
    "be due to these overlaps. The regression models in Sections 4.6–4.9 separate these influences."
)
subsection("Self-Reported Health Condition")
para(
    f"{next_label('Figure')} shows a very strong relationship between self-reported health and depression. None of the "
    f"students in good health and only {R['health_normal_rate']:.1f}% of those in normal health are depressed, compared "
    f"with {R['health_abnormal_rate']:.1f}% of students in abnormal health and {R['health_critical_rate']:.0f}% of those in "
    "critical health. The relationship is so close that health condition almost perfectly predicts depression. As "
    "explained in Section 3.2, this suggests that the health variable captures the same underlying condition as the "
    "outcome, so it is not used as an explanatory variable in the main model."
)
figure("fig03_health_condition", "Depression by self-reported health condition", width_cm=13)
subsection("Numerical Variables: CGPA, Age and Semester")
cg, ag, se = ttest("CGPA"), ttest("Age (years)"), ttest("Current semester")
para(
    f"{next_label('Figure')} compares the distribution of the numerical variables between depressed and non-depressed "
    f"students. Mean CGPA on the 10-point scale is almost identical in the two groups ({cg['Mean (not depressed)']:.2f} vs "
    f"{cg['Mean (depressed)']:.2f}), and so are mean age ({ag['Mean (not depressed)']:.1f} vs "
    f"{ag['Mean (depressed)']:.1f} years) and mean semester ({se['Mean (not depressed)']:.1f} vs "
    f"{se['Mean (depressed)']:.1f})."
)
figure("fig04_numeric_by_depression", "Distribution of CGPA, age and semester by depression status")
subsection("Age and Semester")
sem = T["depression_rate_by_semester"]
sem_big = sem[sem["Students"] >= 20]
small_sem = sem[sem["Students"] < 20]
para(
    f"{next_label('Figure')} shows depression rates by age and by semester group. Rates vary between ages without a "
    "clear upward or downward trend, and the very small groups at ages 18 and 19 should be read with caution. Across "
    f"the larger semester groups the rate ranges from {100 * sem_big['Depression rate'].min():.0f}% to "
    f"{100 * sem_big['Depression rate'].max():.0f}% without a steady trend"
    + (f" (the {', '.join(small_sem['semester_group'].astype(str))} group has only "
       f"{', '.join(small_sem['Students'].astype(str))} students)" if len(small_sem) else "")
    + ", suggesting that depression is not concentrated at a particular stage of the degree."
)
figure("fig05_rate_by_age_semester", "Depression rate by age and semester group")

section("Descriptive Statistics of Variables")
para(f"{next_label('Table')} reports summary statistics for the coded variables (equivalent to Stata's summarize command).")
summ = T["summary_statistics"]
table(summ, "Summary statistics of coded variables", decimals=3)
sv = summ.set_index("Variable")
para(
    f"Students are on average {sv.loc['Age (years)', 'Mean']:.1f} years old, with a mean CGPA of "
    f"{sv.loc['CGPA', 'Mean']:.2f} out of 10. {100 * sv.loc[MID, 'Mean']:.0f}% have a mid and "
    f"{100 * sv.loc[HIGH, 'Mean']:.0f}% a high monthly living expense. "
    f"{100 * sv.loc['Single', 'Mean']:.0f}% are single, "
    f"{100 * sv.loc['Has friend circle', 'Mean']:.0f}% have a friend circle and "
    f"{100 * sv.loc['Lives with family', 'Mean']:.0f}% live with their family."
)
cm = T["correlation_matrix"].set_index(T["correlation_matrix"].columns[0])
para(
    f"{next_label('Figure')} shows the correlation matrix. Apart from the very high correlation between depression and "
    "health level discussed above, correlations among the explanatory variables are modest. The largest are between "
    f"the two expense dummies ({cm.loc[MID, HIGH]:.2f}, which follows from their construction), between being male and "
    f"living with family ({cm.loc['Male', 'Lives with family']:.2f}) and between high expense and lower-middle class "
    f"({cm.loc[HIGH, 'Lower-middle class']:.2f}), none of which is high enough to cause estimation problems."
    .replace("-0.", "−0.")
)
figure("fig06_correlation_matrix", "Correlation matrix of coded variables", width_cm=14.5)

section("Bivariate Analysis")
subsection("Chi-Square Tests of Independence")
ch_rel, ch_fam, ch_liv, ch_hea = chi("Relationship status"), chi("Family background"), chi("Living style"), chi("Health condition")
ch_exp = chi("Living expense")
para(
    f"{next_label('Table')} tests whether each categorical variable is independent of depression. Relationship status "
    f"(χ² = {ch_rel['Chi-square']:.2f}, {fmt_p(ch_rel['p-value'])}, Cramér's V = {ch_rel[CRAMER]:.2f}), "
    f"family background (χ² = {ch_fam['Chi-square']:.2f}, {fmt_p(ch_fam['p-value'])}), living expense "
    f"(χ² = {ch_exp['Chi-square']:.2f}, {fmt_p(ch_exp['p-value'])}) and living style "
    f"(χ² = {ch_liv['Chi-square']:.2f}, {fmt_p(ch_liv['p-value'])}) are significantly associated with depression, as "
    f"is health condition (Cramér's V = {ch_hea[CRAMER]:.2f}, a very strong association). Gender and friend circle "
    "show no significant association on their own. For relationship status some expected cell counts are below 5, "
    "which is one more reason the small 'complicated' group was merged before modelling."
)
chi_t = T["chi_square_tests"].copy()
chi_t["p-value"] = chi_t["p-value"].map(lambda v: "<0.001" if v < 0.001 else f"{v:.3f}")
table(chi_t, "Chi-square tests of independence with depression", font_size=9)
subsection("Difference in Means Tests")
para(
    f"{next_label('Table')} compares the means of the numerical variables. None of the differences is significant: "
    f"CGPA ({fmt_p(cg['t-test p-value'])}), age ({fmt_p(ag['t-test p-value'])}) and semester "
    f"({fmt_p(se['t-test p-value'])}) show no difference between depressed and non-depressed students, and the "
    "Mann–Whitney tests agree."
)
table(T["t_tests"], "Difference in means between depressed and non-depressed students", font_size=8.5, decimals=3)

section("Multicollinearity Check")
vif_tab = T["vif"].set_index("Variable")
para(
    f"{next_label('Table')} reports the variance inflation factors for the explanatory variables of the main model. "
    f"The largest VIF is {R['vif_max']:.2f}, below the common warning threshold of 10 and the stricter threshold of 5. "
    f"The highest values belong to the two living-expense dummies ({vif_tab.loc[MID, 'VIF']:.2f} and "
    f"{vif_tab.loc[HIGH, 'VIF']:.2f}), as expected for dummies of the same variable that also overlap with family "
    "background. Multicollinearity is therefore not a concern."
)
table(T["vif"], "Variance inflation factors", decimals=3)

section("Inferential Analysis: Binary Logistic Regression")
para(
    f"{next_label('Table')} presents the logit estimates. The model is highly significant overall: the likelihood-ratio "
    f"test rejects the hypothesis that all coefficients are zero (LR χ²({R['logit_lr_df']}) = {R['logit_lr_chi2']:.2f}, "
    f"{fmt_p(R['logit_lr_p'])}). McFadden's pseudo-R² is {R['logit_pseudo_r2']:.3f}, which is a reasonable value for "
    "cross-sectional individual-level data (values of 0.2–0.4 are generally considered a good fit for logit models)."
)
lt = T["logit_coefficients"].copy()
lt["Sig."] = lt["Sig."].fillna("")
table(lt, f"Binary logit estimates (N = {R['n']}; dependent variable: depression)", font_size=8.5, decimals=4)
note("Significance levels: *** p < 0.01, ** p < 0.05, * p < 0.10.")
subsection("Interpretation of Coefficients")
c_single, c_mid, c_high, c_rich, c_male = coef("Single"), coef(MID), coef(HIGH), coef("Rich class"), coef("Male")
bullets([
    f"**Single (β = {c_single['Coefficient']:.3f}, {fmt_p(c_single['P>|z|'])}):** the strongest and most significant "
    "effect. Holding all other factors constant, single students are much more likely to be depressed than students in a relationship.",
    f"**Mid living expense (β = {c_mid['Coefficient']:.3f}, {fmt_p(c_mid['P>|z|'])}) and high living expense "
    f"(β = {c_high['Coefficient']:.3f}, {fmt_p(c_high['P>|z|'])}):** students whose monthly living expense is mid or "
    "high are significantly more likely to be depressed than students with a low expense. The two coefficients are "
    "almost equal.",
    f"**Male (β = {c_male['Coefficient']:.3f}, {fmt_p(c_male['P>|z|'])}):** male students are significantly less "
    "likely to be depressed than female students once other factors are controlled.",
    f"**Rich class (β = {c_rich['Coefficient']:.3f}, {fmt_p(c_rich['P>|z|'])}):** the coefficient is negative, as "
    "expected, but not significant once living expense is controlled for.",
    "**CGPA, age, semester, lower-middle class, living with family and friend circle** are not statistically "
    "significant: once the other variables are held constant, they do not explain differences in depression.",
])
subsection("Odds Ratios")
o_single, o_mid, o_high, o_male = odds("Single"), odds(MID), odds(HIGH), odds("Male")
para(
    f"{next_label('Table')} and {next_label('Figure')} report the odds ratios. The odds of depression for single students "
    f"are {o_single['Odds ratio']:.1f} times those of students in a relationship (95% CI {o_single['95% CI lower']:.1f}–"
    f"{o_single['95% CI upper']:.1f}). Students with a mid or high living expense have {o_mid['Odds ratio']:.1f} and "
    f"{o_high['Odds ratio']:.1f} times the odds of students with a low expense, although the confidence intervals are "
    f"wide. Male students have {100 * (1 - o_male['Odds ratio']):.0f}% lower odds than female students."
)
ot = T["logit_odds_ratios"].copy()
ot["Sig."] = ot["Sig."].fillna("")
table(ot, "Odds ratios from the logit model", font_size=9, decimals=3)
figure("fig07_logit_odds_ratios", "Odds ratios with 95% confidence intervals (log scale)", width_cm=14)

section("Marginal Effects Analysis (Logit)")
para(
    f"Coefficients and odds ratios show the direction and relative strength of effects, but marginal effects are easier "
    f"to interpret because they are expressed in probability terms. {next_label('Table')} and {next_label('Figure')} "
    f"report the average marginal effects. For reference, the predicted probability of depression for a student with "
    f"average characteristics is {100 * R['logit_pred_at_means']:.1f}%."
)
la = T["logit_marginal_effects_ame"].copy()
la["Sig."] = la["Sig."].fillna("")
table(la, "Average marginal effects after logit (dy/dx)", font_size=8.5, decimals=4)
figure("fig08_logit_marginal_effects", "Average marginal effects after logit with 95% confidence intervals", width_cm=14)
subsection("Interpretation of Marginal Effects")
bullets([
    f"Being **single** raises the probability of depression by {100 * single_ame['dy/dx']:.1f} percentage points "
    f"({fmt_p(single_ame['P>|z|'])}) — by far the largest effect in the model.",
    f"Compared with a low living expense, a **mid expense** raises the probability by {100 * mid_ame['dy/dx']:.1f} "
    f"percentage points ({fmt_p(mid_ame['P>|z|'])}) and a **high expense** by {100 * high_ame['dy/dx']:.1f} points "
    f"({fmt_p(high_ame['P>|z|'])}). Because the two effects are almost equal, the risk rises in one step from low to "
    "any higher expense rather than increasing steadily.",
    f"**Male** students are {abs(100 * male_ame['dy/dx']):.1f} percentage points less likely to be depressed than "
    f"female students ({fmt_p(male_ame['P>|z|'])}).",
    f"Coming from a **rich family** lowers the probability by {abs(100 * rich_ame['dy/dx']):.1f} percentage points "
    f"relative to a middle-class family, significant only at the 10% level ({fmt_p(rich_ame['P>|z|'])}); a "
    f"**lower-middle class** background makes almost no difference ({pp(low_ame['dy/dx'])}, {fmt_p(low_ame['P>|z|'])}).",
    f"A one-point increase in **CGPA** (on the 10-point scale) changes the probability by only "
    f"{100 * cgpa_ame['dy/dx']:+.1f} percentage points and is not significant ({fmt_p(cgpa_ame['P>|z|'])}); having a "
    f"**friend circle** is also not significant ({fmt_p(friend_ame['P>|z|'])}).",
])
para(
    f"Marginal effects at the means of all variables (equivalent to Stata's older mfx command) give very similar "
    "conclusions and are reported in the Excel workbook accompanying this report."
)

section("Binary Probit Regression")
para(
    f"{next_label('Table')} reports the probit estimates. The model is again highly significant "
    f"(LR χ²({R['logit_lr_df']}) = {R['probit_lr_chi2']:.2f}, {fmt_p(R['probit_lr_p'])}) with a pseudo-R² of "
    f"{R['probit_pseudo_r2']:.3f}, almost identical to the logit model. Probit coefficients are on a different scale "
    "(roughly the logit coefficients divided by 1.6–1.8), but the signs and significance pattern are the same."
)
pt = T["probit_coefficients"].copy()
pt["Sig."] = pt["Sig."].fillna("")
table(pt, f"Binary probit estimates (N = {R['n']}; dependent variable: depression)", font_size=8.5, decimals=4)
subsection("Interpretation of Coefficients")
pc_male, pc_mid, pc_high, pc_rich = coef("Male", "probit"), coef(MID, "probit"), coef(HIGH, "probit"), coef("Rich class", "probit")
para(
    "Single status remains strongly positive and highly significant. Mid and high living expense remain positive and "
    f"significant ({fmt_p(pc_mid['P>|z|'])} and {fmt_p(pc_high['P>|z|'])}), and the male coefficient remains negative "
    f"and significant ({fmt_p(pc_male['P>|z|'])}). Rich family background is again negative but not significant "
    f"({fmt_p(pc_rich['P>|z|'])}). CGPA, age, semester, lower-middle class, living with family and friend circle "
    "remain insignificant."
)

section("Marginal Effects (Probit)")
para(f"{next_label('Table')} reports the average marginal effects after probit.")
pa = T["probit_marginal_effects_ame"].copy()
pa["Sig."] = pa["Sig."].fillna("")
table(pa, "Average marginal effects after probit (dy/dx)", font_size=8.5, decimals=4)
ps, pmid, phigh, pm = ame("Single", "probit"), ame(MID, "probit"), ame(HIGH, "probit"), ame("Male", "probit")
para(
    f"The probit marginal effects are almost identical to the logit ones: being single raises the probability of "
    f"depression by {100 * ps['dy/dx']:.1f} percentage points (logit: {100 * single_ame['dy/dx']:.1f}), a mid expense "
    f"by {100 * pmid['dy/dx']:.1f} points (logit: {100 * mid_ame['dy/dx']:.1f}), a high expense by "
    f"{100 * phigh['dy/dx']:.1f} points (logit: {100 * high_ame['dy/dx']:.1f}), and being male lowers it by "
    f"{abs(100 * pm['dy/dx']):.1f} points (logit: {abs(100 * male_ame['dy/dx']):.1f}). The choice of link function "
    "therefore makes no practical difference to the conclusions."
)

section("Goodness of Fit and Model Diagnostics")
para(f"{next_label('Table')} summarises the fit statistics of both models.")
fs = T["model_fit_statistics"].set_index("Model").T.reset_index().rename(columns={"index": "Statistic"})
table(fs, "Goodness-of-fit statistics", font_size=9, decimals=4)
para(
    f"Both models have almost the same likelihood, AIC and AUC. The Hosmer–Lemeshow test does not reject the logit "
    f"model (χ² = {R['logit_hl_chi2']:.2f}, p = {R['logit_hl_p']:.3f}), meaning that its predicted and observed "
    f"depression rates agree across risk groups. For the probit model the test is rejected "
    f"(p = {T['model_fit_statistics'].set_index('Model').loc['Probit', 'HL p-value']:.3f}), which supports keeping the "
    "logit as the main model. The AIC values of the two models ({R['logit_aic']:.2f} for logit, "
    f"{R['probit_aic']:.2f} for probit) differ only negligibly."
)
para(f"{next_label('Table')} shows the classification performance (equivalent to Stata's estat classification).")
cl = T["classification_tables"].copy()
table(cl, "Classification tables at the 0.5 cut-off and at the sample share of depressed students", font_size=8, decimals=3)
para(
    f"At the conventional 0.5 cut-off the logit model classifies {100 * R['logit_correct_05']:.1f}% of students "
    f"correctly, with a specificity of {100 * R['logit_spec_05']:.1f}% but a sensitivity of only "
    f"{100 * R['logit_sens_05']:.1f}%. Because depression affects fewer than half of the students, a lower cut-off "
    f"equal to the sample share of depressed students ({R['pct_depressed'] / 100:.3f}) is more appropriate for "
    f"screening: it identifies {100 * cl.loc[1, 'Sensitivity']:.1f}% of depressed students while still correctly "
    f"classifying {100 * cl.loc[1, 'Specificity']:.1f}% of non-depressed students."
)
para(
    f"{next_label('Figure')} plots the ROC curves. The area under the curve is {R['logit_auc']:.3f} for the logit and "
    f"{R['probit_auc']:.3f} for the probit model, indicating acceptable discrimination: a randomly chosen depressed "
    f"student has a {100 * R['logit_auc']:.0f}% chance of receiving a higher predicted probability than a randomly "
    "chosen non-depressed student."
)
figure("fig09_roc_curves", "ROC curves of the logit and probit models", width_cm=10.5)
prof = T["predicted_probability_profiles"]
prof_s, prof_n = (prof.loc[prof["Single"] == s, "Predicted probability"] for s in (1, 0))
para(
    f"{next_label('Figure')} illustrates the size of the relationship-status effect. Holding all other variables at their "
    f"means, the predicted probability of depression for single students ranges from {100 * prof_s.min():.0f}% to "
    f"{100 * prof_s.max():.0f}% across the CGPA range, compared with {100 * prof_n.min():.0f}–{100 * prof_n.max():.0f}% "
    "for students who are not single. Both lines slope gently downward, but the CGPA effect is not statistically "
    "significant."
)
figure("fig10_predicted_probability_cgpa", "Predicted probability of depression by CGPA and relationship status", width_cm=12.5)

section("Hypothesis Testing Summary")
para(f"{next_label('Table')} summarises the outcome of each hypothesis using the logit model.")
table(hyp[["Hypothesis", "Variable", "Expected sign", "Logit coefficient", "AME", "P>|z|", "Result"]],
      "Summary of hypothesis tests", font_size=8.5, decimals=4)
sup = hyp[hyp["Result"] == "Supported"]["Hypothesis"].tolist()
weak = hyp[hyp["Result"].str.startswith("Weakly")]["Hypothesis"].tolist()
nots = hyp[hyp["Result"].str.startswith("Not")]["Hypothesis"].tolist()
assert sup == ["H3", "H5", "H7"] and not weak, "text below describes H3, H5 and H7 as the supported hypotheses"
para(
    "Hypotheses H3 (single students), H5 (high living expense) and H7 (female students) are supported at the 5% level. "
    f"Hypotheses {', '.join(nots)} are not supported: academic performance, friend circle, family background and "
    "living with family do not have a significant independent effect on depression at the 5% level. The rich-family "
    "effect (H4b) has the expected sign and its marginal effect is significant at the 10% level."
)

section("Robustness Checks")
para(
    f"{next_label('Table')} compares the main logit model with six alternative specifications. Coefficients are shown "
    "with standard errors in parentheses."
)
rb = T["robustness_models"].fillna("")
rb_short = rb.rename(columns=lambda c: c.replace("(", "").replace(")", "").replace(" robust SE", " rob.").replace("Logit + poor health", "Logit+health"))
table(rb_short, "Robustness checks across model specifications", font_size=7,
      col_widths=[2.6, 1.9, 1.9, 1.9, 1.9, 1.9, 1.9, 1.9])
note("Significance levels: *** p < 0.01, ** p < 0.05, * p < 0.10. Standard errors in parentheses.")
para(
    f"{next_label('Figure')} compares the marginal effects of the logit, probit and linear probability models side by "
    "side. The logit and probit effects are almost identical for every variable; the linear probability model agrees "
    "on relationship status and gender but gives smaller living-expense effects."
)
figure("fig11_robustness_comparison", "Marginal effects across logit, probit and linear probability models", width_cm=14)
sub = T["subsample_marginal_effects"]


def sub_me(sample, var):
    r = sub[(sub["Sample"] == sample) & (sub["Variable"] == var)].iloc[0]
    return r["dy/dx"], r["P>|z|"]


m_s, f_s = sub_me("Male students", "Single"), sub_me("Female students", "Single")
m_r, f_r = sub_me("Male students", "Rich class"), sub_me("Female students", "Rich class")
m_mid, f_mid = sub_me("Male students", MID), sub_me("Female students", MID)
m_high, f_high = sub_me("Male students", HIGH), sub_me("Female students", HIGH)
f_g = sub_me("Female students", "CGPA")
cmp_ = T["marginal_effects_comparison"].set_index("Variable")
bullets([
    "**Robust standard errors (model 2):** using heteroskedasticity-robust standard errors changes the standard errors "
    "only slightly; every variable keeps its sign and its significance at the 5% level.",
    f"**Linear probability model (model 4):** the OLS coefficients, which are directly marginal effects "
    f"(R² = {R['lpm_r2']:.3f}), are very close to the logit and probit AMEs for being single "
    f"({pp(cmp_.loc['Single', 'LPM coefficient'])}) and male ({pp(cmp_.loc['Male', 'LPM coefficient'])}). The "
    f"living-expense effects are smaller ({pp(cmp_.loc[MID, 'LPM coefficient'])} for mid and "
    f"{pp(cmp_.loc[HIGH, 'LPM coefficient'])} for high) and significant only at the 10% level "
    f"(p = {cmp_.loc[MID, 'LPM p (robust)']:.3f} and {cmp_.loc[HIGH, 'LPM p (robust)']:.3f}), so the expense result "
    "is somewhat less robust than the relationship and gender results.",
    f"**Adding poor health (model 5):** including the health indicator raises the pseudo-R² from "
    f"{R['logit_pseudo_r2']:.3f} to {R['health_model_pseudo_r2']:.3f} and the AUC to {R['health_model_auc']:.3f} "
    f"(LR χ²(1) = {R['lr_health_chi2']:.1f}, {fmt_p(R['lr_health_p'])}). Health status alone explains most of the "
    "variation in depression, and the estimation software flags quasi-complete separation (a fraction "
    f"{QS_FRACTION} of the observations is predicted perfectly), which confirms that it largely measures the same condition as the "
    "outcome. Being single and being male remain significant even in this model; the living-expense and rich-family "
    "coefficients lose significance.",
    f"**Male and female students (models 6 and 7):** a separate logit cannot be estimated for male students because "
    f"none of the {R['male_not_single_n']} male students who are not single is depressed — being single predicts "
    "depression perfectly within this group. The gender subsamples are therefore estimated as linear probability "
    f"models. Being single is a strong risk factor for both men ({pp(m_s[0])}, {fmt_p(m_s[1])}) and women "
    f"({pp(f_s[0])}, {fmt_p(f_s[1])}). Among men, a mid ({pp(m_mid[0])}, {fmt_p(m_mid[1])}) or high living expense "
    f"({pp(m_high[0])}, {fmt_p(m_high[1])}) raises the risk and a rich family lowers it ({pp(m_r[0])}, "
    f"{fmt_p(m_r[1])}). Among women these effects are smaller and not significant (p = {f_mid[1]:.2f}, "
    f"{f_high[1]:.2f} and {f_r[1]:.2f}), while a higher CGPA is associated with a lower risk "
    f"({100 * f_g[0]:+.1f} points per grade point, {fmt_p(f_g[1])}). Each subsample has fewer than 50 depressed "
    "students, so these differences are indicative only.",
])

# ============================================================================= Chapter 5
# (no page break: chapters run on continuously)
chapter("Findings, Conclusion, and Recommendations")
section("Introduction")
para(
    "This chapter summarises the main findings, discusses what they mean for universities and for theory, sets out "
    "the limitations of the study, and suggests directions for future research."
)
section("Summary of Major Findings")
subsection("Relationships Matter More than Grades")
para(
    f"The single most important factor in this study is relationship status. Single students are "
    f"{100 * single_ame['dy/dx']:.0f} percentage points more likely to be depressed than students in a relationship, "
    "an effect that is large, highly significant and stable across every model specification and for both men and "
    "women. This strongly supports the stress-buffering view of close relationships (Cohen & Wills, 1985)."
)
subsection("Financial Pressure: Living Expenses and Family Background")
para(
    f"Students whose monthly living expense is mid or high are about "
    f"{100 * (mid_ame['dy/dx'] + high_ame['dy/dx']) / 2:.0f} percentage points more likely to be depressed than "
    "students with a low expense, after controlling for family background and the other factors. The risk rises in one "
    "step from low to any higher expense rather than increasing steadily. Family background matters less once "
    "expenses are taken into account: middle-class students have the highest raw rate, and rich students are less "
    "likely to be depressed, but only at the 10% level. A possible explanation is that higher day-to-day costs create "
    "financial pressure that students from middle-class families cannot easily absorb, consistent with research "
    "linking financial strain to poor mental health (Lorant et al., 2003; Richardson et al., 2013)."
)
subsection("The Gender Gap")
para(
    f"Once other factors are held constant, female students are about {abs(100 * male_ame['dy/dx']):.0f} percentage "
    "points more likely to be depressed than male students. This gap is hidden in the raw data, where male students "
    "actually show a slightly higher rate, because male students are more often single. The adjusted result is "
    "consistent with the well-established gender difference in depression (Nolen-Hoeksema, 1987; Kessler, 2003)."
)
subsection("What Did Not Matter: Grades, Friends and Living Arrangement")
para(
    "Contrary to common expectations, CGPA has no significant relationship with depression, and neither does the stage "
    "of study (semester) or age. Because every student in the sample has a CGPA above 6.00, this means that grades "
    "make little difference within this range; it does not show that very low grades are harmless. Having a friend "
    "circle and living with family also show no independent effect. Living with family appears protective in the raw "
    "data, but this disappears once relationship status and gender are controlled for. The kind of social support "
    "that matters most in this sample seems to be a close, intimate relationship rather than a wider friend group."
)
subsection("Health and Depression")
para(
    "Self-reported health condition is almost a mirror image of depression: all students in critical health are "
    "depressed and none in good health are. This shows how closely physical and mental health are linked, but also "
    "warns that self-reported health measures can overlap with the outcome they are meant to explain."
)

section("Policy Implications for Universities")
subsection("Target Social Isolation, not only Academic Performance")
para(
    "Since academic performance does not predict depression in this sample, support services that focus mainly on "
    "academic difficulties will miss many depressed students. Universities should invest in programmes that reduce "
    "loneliness and build close connections — peer mentoring, small-group activities, residential community programmes "
    "and relationship and emotional-skills workshops."
)
subsection("Address Financial Pressure")
para(
    "Because students with higher living expenses are more often depressed, universities should combine mental health "
    "support with financial guidance: budgeting support, affordable housing and meal options, and need-based "
    "assistance that also reaches middle-income students, who may fall outside traditional aid schemes."
)
subsection("Gender-Sensitive Screening and Support")
para(
    "Female students face a higher adjusted risk, while male students may be less likely to seek help. Screening "
    "campaigns and counselling services should be designed to reach both groups — for example, confidential online "
    "screening, counsellors of both genders, and awareness campaigns that reduce stigma among men."
)
subsection("Simple Screening Tools")
para(
    "The model shows that a few easily collected characteristics can identify a large share of students at risk: using "
    f"a cut-off equal to the sample share of depressed students, the logit model identifies "
    f"{100 * cl.loc[1, 'Sensitivity']:.0f}% of depressed students. Such risk scores should never replace clinical "
    "assessment, but they can help universities decide whom to invite for further screening."
)

section("Theoretical Implications")
para(
    "The results support the diathesis–stress and stress-buffering frameworks: stressors such as financial pressure "
    "raise the probability of depression, while close relationships protect against it. They also suggest a refinement "
    "of social support theory for young adults — the protective effect appears to come from intimate relationships "
    "rather than from the size of the wider social network. Finally, the step pattern of living expenses, with mid and "
    "high expenses carrying the same extra risk, suggests that moving beyond a low cost of living — rather than the "
    "exact amount spent — is what adds pressure."
)

section("Limitations of the Study")
bullets([
    "**Cross-sectional design:** the data capture one point in time, so the results show associations, not causal "
    "effects. Depression may itself affect relationship status (for example, by making it harder to form relationships).",
    "**Self-reported measures:** depression and health condition are self-reported yes/no or categorical indicators, "
    "not clinical diagnoses or validated symptom scales.",
    f"**Sample size:** with {R['n']} students and {R['n_depressed']} depressed cases, the small 'complicated' "
    "relationship group had to be merged, and the gender subsample results are indicative only.",
    "**Banded living expense:** living expense is recorded in three bands, which loses detail within each band, and "
    "the expense effects are weaker in the linear probability model.",
    "**Restricted sample:** the sample covers only students with a CGPA above 6.00 and is not representative of all "
    "university students, which limits how far the results can be generalised.",
    "**Omitted variables:** factors such as family history of mental illness, substance use, sleep, physical activity "
    "and exposure to traumatic events are not available in the data.",
])

section("Recommendations for Future Research")
bullets([
    "**Longitudinal studies** following students over several semesters to establish whether changes in relationships "
    "and finances precede changes in depression.",
    "**Validated depression scales** such as the PHQ-9, which would allow depression severity to be modelled with "
    "ordered or count models rather than a simple yes/no outcome.",
    "**Larger and more diverse samples** covering the full range of grades, to estimate separate effects for small "
    "groups.",
    "**Richer measures of social support**, distinguishing the quality of friendships and family relationships from "
    "their mere presence.",
])

section("Conclusion")
para(
    f"Depression affects almost three in ten students in this sample. The analysis shows that it is driven less by "
    "grades and more by students' relationships and financial circumstances. Being single is the strongest risk "
    "factor; a mid or high living expense raises the risk compared with a low one; and male students are less likely "
    "to be depressed than female students, while a rich family background is only weakly protective. These findings "
    "are consistent across the logit and probit models, and the relationship and gender effects also hold in the "
    "linear probability model. "
    "For universities, the message is clear: protecting student mental health requires attention to social "
    "connection and financial wellbeing, not only to academic support."
)

# ============================================================================= Appendix
# (no page break: chapters run on continuously)
chapter("Appendix", numbered_chapter=False)
para(
    "This appendix reports the full estimation output produced by the analysis notebook "
    "(notebooks/student_depression_analysis.ipynb), computed with Python statsmodels. The Stata do-file "
    "stata/student_depression_analysis.do contains the equivalent Stata commands (summarize, tabulate, logit, "
    "probit, margins, estat classification, estat gof, lroc, regress) and reproduces the same models. All tables are "
    "also available in the Excel workbook results/Student_Depression_Analysis_Results.xlsx."
)
for title, fname in [
    ("A.1  Binary logit model", "logit_summary"),
    ("A.2  Binary probit model", "probit_summary"),
    ("A.3  Logit model with robust standard errors", "logit_robust_summary"),
    ("A.4  Linear probability model (robust standard errors)", "lpm_summary"),
    ("A.5  Logit model including poor health", "logit_health_summary"),
    ("A.6  Linear probability model: male students", "lpm_male_summary"),
    ("A.7  Linear probability model: female students", "lpm_female_summary"),
]:
    h = doc.add_paragraph(style="Heading 3")
    h.add_run(title)
    code_block((RES / "model_output" / f"{fname}.txt").read_text())

h = doc.add_paragraph(style="Heading 3")
h.add_run("A.8  Marginal effects at the means (logit)")
mem = T["logit_marginal_effects_mem"].copy()
mem["Sig."] = mem["Sig."].fillna("")
table(mem, "Marginal effects at the means after logit", font_size=8.5, decimals=4)
h = doc.add_paragraph(style="Heading 3")
h.add_run("A.9  Subsample marginal effects (linear probability model)")
sp = sub.pivot(index="Variable", columns="Sample", values=["dy/dx", "P>|z|"])
sp.columns = [f"{b}: {a}" for a, b in sp.columns]
table(sp.reset_index(), "Marginal effects for male and female students (LPM, robust SE)", font_size=8.5, decimals=4)

# ============================================================================= References
# (no page break: chapters run on continuously)
chapter("References", numbered_chapter=False)
REFS = [
    "Beck, A. T. (1967). Depression: Clinical, experimental, and theoretical aspects. New York: Harper & Row.",
    "Bulloch, A. G., Williams, J. V., Lavorato, D. H., & Patten, S. B. (2009). The relationship between major depression "
    "and marital disruption is bidirectional. Depression and Anxiety, 26(12), 1172–1177.",
    "Cohen, S., & Wills, T. A. (1985). Stress, social support, and the buffering hypothesis. Psychological Bulletin, "
    "98(2), 310–357.",
    "Eisenberg, D., Gollust, S. E., Golberstein, E., & Hefner, J. L. (2007). Prevalence and correlates of depression, "
    "anxiety, and suicidality among university students. American Journal of Orthopsychiatry, 77(4), 534–542.",
    "Hosmer, D. W., Lemeshow, S., & Sturdivant, R. X. (2013). Applied logistic regression (3rd ed.). Hoboken, NJ: Wiley.",
    "Hysenbegasi, A., Hass, S. L., & Rowland, C. R. (2005). The impact of depression on the academic productivity of "
    "university students. Journal of Mental Health Policy and Economics, 8(3), 145–151.",
    "Ibrahim, A. K., Kelly, S. J., Adams, C. E., & Glazebrook, C. (2013). A systematic review of studies of depression "
    "prevalence in university students. Journal of Psychiatric Research, 47(3), 391–400.",
    "Kawachi, I., & Berkman, L. F. (2001). Social ties and mental health. Journal of Urban Health, 78(3), 458–467.",
    "Kessler, R. C. (2003). Epidemiology of women and depression. Journal of Affective Disorders, 74(1), 5–13.",
    "Long, J. S., & Freese, J. (2014). Regression models for categorical dependent variables using Stata (3rd ed.). "
    "College Station, TX: Stata Press.",
    "Lorant, V., Deliège, D., Eaton, W., Robert, A., Philippot, P., & Ansseau, M. (2003). Socioeconomic inequalities in "
    "depression: A meta-analysis. American Journal of Epidemiology, 157(2), 98–112.",
    "Monroe, S. M., & Simons, A. D. (1991). Diathesis–stress theories in the context of life stress research: "
    "Implications for the depressive disorders. Psychological Bulletin, 110(3), 406–425.",
    "Moussavi, S., Chatterji, S., Verdes, E., Tandon, A., Patel, V., & Ustun, B. (2007). Depression, chronic diseases, "
    "and decrements in health: Results from the World Health Surveys. The Lancet, 370(9590), 851–858.",
    "Nolen-Hoeksema, S. (1987). Sex differences in unipolar depression: Evidence and theory. Psychological Bulletin, "
    "101(2), 259–282.",
    "Richardson, T., Elliott, P., & Roberts, R. (2013). The relationship between personal unsecured debt and mental and "
    "physical health: A systematic review and meta-analysis. Clinical Psychology Review, 33(8), 1148–1162.",
    "Seabold, S., & Perktold, J. (2010). Statsmodels: Econometric and statistical modeling with Python. Proceedings of "
    "the 9th Python in Science Conference, 92–96.",
    "StataCorp. (2021). Stata statistical software: Release 17. College Station, TX: StataCorp LLC.",
    "Whitton, S. W., & Whisman, M. A. (2010). Relationship satisfaction instability and depression. Journal of Family "
    "Psychology, 24(6), 791–794.",
    "Wooldridge, J. M. (2010). Econometric analysis of cross section and panel data (2nd ed.). Cambridge, MA: MIT Press.",
]
for ref in REFS:
    p = para(ref, size=11)
    p.paragraph_format.left_indent = Cm(1)
    p.paragraph_format.first_line_indent = Cm(-1)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT

doc.core_properties.title = TITLE
doc.core_properties.author = ", ".join(AUTHORS)
doc.core_properties.last_modified_by = ""
doc.core_properties.comments = ""
OUT.parent.mkdir(exist_ok=True)
doc.save(OUT)
print("Saved", OUT.relative_to(ROOT))
