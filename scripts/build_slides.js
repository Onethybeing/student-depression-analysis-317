/*
 * Build the 10-slide presentation from results/slides_data.json (made by scripts/export_slide_data.py).
 * Charts are native PowerPoint charts, so every number stays editable.
 *
 * Usage (from the repository root):
 *   python scripts/export_slide_data.py
 *   npm install pptxgenjs          # once
 *   node scripts/build_slides.js   # writes report/Student_Depression_Presentation.pptx
 */
const path = require("path");
const fs = require("fs");
const pptxgen = require("pptxgenjs");

const ROOT = path.resolve(__dirname, "..");
const D = JSON.parse(fs.readFileSync(path.join(ROOT, "results", "slides_data.json"), "utf8"));
const R = D.R;
const OUT = path.join(ROOT, "report", "Student_Depression_Presentation.pptx");
const LOGO = path.join(ROOT, "assets", "iitkgp_logo.png");
const APPLY_THEME = process.env.APPLY_THEME_JS; // optional path to the pptx skill's apply_theme.js

const THEME = {
  name: "Student Wellbeing",
  headFontFace: "Cambria",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "1B1F24", lt1: "FFFFFF", dk2: "12355B", lt2: "EEF4F6",
    accent1: "1B7F8C", accent2: "E07A5F", accent3: "12355B", accent4: "8AB6BF",
    accent5: "F2CC8F", accent6: "5C6B73", hlink: "1B7F8C", folHlink: "5C6B73",
  },
};
const HEX = THEME.colors;
const MUTED = "5C6B73";
const GRID = "DDE5E8";

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5 in
pres.title = "Determinants of Depression among University Students";
pres.author = "Atharv Lodhi, Sourav Choudhary, Piyush Bisht";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
const C = pres.SchemeColor;
const W = 13.333;

// ------------------------------------------------------------------------- layouts
pres.defineSlideMaster({
  title: "TITLE_DARK",
  background: { color: HEX.dk2 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.7, y: 1.55, w: 7.6, h: 2.3, fontSize: 38, bold: true,
      color: C.background1, valign: "bottom", align: "left", fontFace: THEME.headFontFace }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: 0.7, y: 3.95, w: 7.6, h: 0.8, fontSize: 18,
      color: "BFE3E8", valign: "top" }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: HEX.lt1 },
  margin: [0.5, 0.6, 0.6, 0.6],
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.35, w: 12.1, h: 0.85, fontSize: 30, bold: true,
      color: C.text2, valign: "middle", fontFace: THEME.headFontFace }, text: "" } },
    { text: { text: "Determinants of Depression among University Students  ·  IIT Kharagpur",
      options: { x: 0.6, y: 7.0, w: 9, h: 0.3, fontSize: 10, color: MUTED, margin: 0 } } },
  ],
  slideNumber: { x: 12.2, y: 7.0, w: 0.5, h: 0.3, fontSize: 10, color: MUTED, align: "right" },
});
pres.defineSlideMaster({
  title: "CLOSING_DARK",
  background: { color: HEX.dk2 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.35, w: 12.1, h: 0.85, fontSize: 30, bold: true,
      color: C.background1, valign: "middle", fontFace: THEME.headFontFace }, text: "" } },
  ],
  slideNumber: { x: 12.2, y: 7.0, w: 0.5, h: 0.3, fontSize: 10, color: "BFE3E8", align: "right" },
});

// ------------------------------------------------------------------------- helpers
const fmtP = (p) => (p < 0.001 ? "p < 0.001" : `p = ${p.toFixed(3)}`);
const signed = (x, d = 1) => (x > 0 ? `+${x.toFixed(d)}` : `−${Math.abs(x).toFixed(d)}`);
const shadow = () => ({ type: "outer", color: "000000", opacity: 0.12, blur: 6, offset: 2, angle: 90 });

function card(slide, name, x, y, w, h, fill = C.background2) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.12, fill: { color: fill },
    line: { color: fill }, shadow: shadow(), objectName: name });
}

function stat(slide, name, x, y, w, value, label, valueColor = C.accent1) {
  card(slide, `${name} card`, x, y, w, 1.55);
  slide.addText(value, { x: x + 0.25, y: y + 0.15, w: w - 0.5, h: 0.8, fontSize: 40, bold: true, color: valueColor,
    fontFace: THEME.headFontFace, margin: 0, isTextBox: true, objectName: `${name} value` });
  slide.addText(label, { x: x + 0.25, y: y + 0.95, w: w - 0.5, h: 0.5, fontSize: 14, color: C.text1, margin: 0,
    isTextBox: true, objectName: `${name} label` });
}

const axisText = { catAxisLabelColor: MUTED, valAxisLabelColor: MUTED, catAxisLabelFontFace: "+mn-lt",
  valAxisLabelFontFace: "+mn-lt", catAxisLabelFontSize: 12, valAxisLabelFontSize: 11,
  dataLabelFontFace: "+mn-lt", dataLabelFontSize: 12, dataLabelColor: "1B1F24",
  titleFontFace: "+mj-lt", titleFontSize: 14, titleColor: HEX.dk2 };

function rateChart(slide, name, group, title, x, y, w, h, color) {
  slide.addChart(pres.charts.BAR, [{ name: "Depressed (%)", labels: group.labels.map((l, i) => `${l} (n=${group.n[i]})`),
    values: group.values }], {
    x, y, w, h, barDir: "col", objectName: name, showTitle: true, title, chartColors: [color],
    showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: '0"%"', valAxisMaxVal: 60, valAxisMinVal: 0,
    valAxisMajorUnit: 20, valAxisLabelFormatCode: '0"%"', valGridLine: { color: GRID, size: 0.75 },
    catGridLine: { style: "none" }, showLegend: false, barGapWidthPct: 60, ...axisText,
  });
}

// ========================================================================= slide 1: title
pres.addSection({ title: "Introduction" });
let s = pres.addSlide({ masterName: "TITLE_DARK", sectionTitle: "Introduction" });
s.addText("Determinants of Depression among University Students", { placeholder: "title" });
s.addText("An econometric analysis using logit and probit models", { placeholder: "body" });
s.addText([
  { text: "Atharv Lodhi  ·  Sourav Choudhary  ·  Piyush Bisht", options: { fontSize: 18, bold: true, color: "FFFFFF", breakLine: true } },
  { text: "Submitted to Dr. Prantik Bagchi", options: { fontSize: 14, color: "BFE3E8", breakLine: true } },
  { text: "Department of Humanities and Social Sciences, IIT Kharagpur", options: { fontSize: 14, color: "BFE3E8" } },
], { x: 0.7, y: 5.2, w: 8.0, h: 1.3, margin: 0, valign: "top", isTextBox: true, objectName: "team and course" });
s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 9.15, y: 1.9, w: 3.5, h: 3.5, rectRadius: 0.25, fill: { color: "FFFFFF" },
  line: { color: "FFFFFF" }, objectName: "logo card" });
s.addImage({ path: LOGO, x: 9.4, y: 2.55, w: 3.0, h: 3.0 * 253 / 360, objectName: "IIT Kharagpur logo" });
s.addNotes("Introduce the team and the question: which student characteristics are associated with depression once they are considered together? We use a 317-student survey sample (CGPA above 6 on a 10-point scale) and binary logit and probit models.");

// ========================================================================= slide 2: overview
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Introduction" });
s.addText("Almost three in ten students report depression", { placeholder: "title" });
stat(s, "students", 0.6, 1.5, 3.4, `${R.n}`, `university students aged ${R.age_min}–${R.age_max}`);
stat(s, "prevalence", 4.2, 1.5, 3.4, `${R.pct_depressed}%`, `report depression (${R.n_depressed} students)`, C.accent2);
stat(s, "gender", 0.6, 3.3, 3.4, `${R.n_male} / ${R.n_female}`, "male / female students");
stat(s, "variables", 4.2, 3.3, 3.4, "11", "explanatory variables in the model");
card(s, "questions card", 8.0, 1.5, 4.7, 5.1);
s.addText("Research questions", { x: 8.3, y: 1.7, w: 4.1, h: 0.5, fontSize: 20, bold: true, color: C.text2, margin: 0,
  fontFace: THEME.headFontFace, isTextBox: true, objectName: "questions heading" });
s.addText([
  "Does relationship status matter?",
  "Do family background and living expenses matter?",
  "Is low CGPA a driver of depression?",
  "Do gender, friends and living with family make a difference?",
  "Which factors remain once all are considered together?",
].map((t, i, a) => ({ text: t, options: { bullet: true, breakLine: i < a.length - 1 } })),
{ x: 8.3, y: 2.3, w: 4.1, h: 4.0, fontSize: 16, color: C.text1, paraSpaceAfter: 10, valign: "top", margin: 0,
  isTextBox: true, objectName: "questions list" });
s.addText(`Sample: survey data on ${R.n} students with CGPA above 6 (10-point scale); no personal identifiers.`,
  { x: 0.6, y: 5.15, w: 7.0, h: 0.9, fontSize: 14, italic: true, color: MUTED, margin: 0, isTextBox: true, objectName: "sample note" });
s.addNotes(`${R.n_depressed} of ${R.n} students (${R.pct_depressed}%) report depression, close to the ~30% found internationally for university students. The outcome is balanced enough for a binary choice model.`);

// ========================================================================= slide 3: method
pres.addSection({ title: "Method" });
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Method" });
s.addText("Method: from survey answers to probability effects", { placeholder: "title" });
const steps = [
  ["Clean & code", "Dummies; tiny groups merged to avoid perfect prediction"],
  ["Describe", "Depression rates and distributions by group"],
  ["Test", "Chi-square, Cramér's V, t-test, Mann–Whitney"],
  ["Model", "Binary logit (main) and probit, 11 regressors"],
  ["Interpret", "Average marginal effects in percentage points"],
  ["Check", "Fit, ROC, Hosmer–Lemeshow, VIF, robustness"],
];
steps.forEach(([head, body], i) => {
  const x = 0.6 + i * 2.05;
  s.addShape(pres.shapes.OVAL, { x: x + 0.55, y: 1.55, w: 0.8, h: 0.8, fill: { color: i === 3 || i === 4 ? C.accent2 : C.accent1 },
    line: { color: "FFFFFF" }, objectName: `step ${i + 1} circle` });
  s.addText(`${i + 1}`, { x: x + 0.55, y: 1.55, w: 0.8, h: 0.8, fontSize: 24, bold: true, color: C.background1, align: "center",
    valign: "middle", margin: 0, isTextBox: true, objectName: `step ${i + 1} number` });
  if (i < steps.length - 1) {
    s.addShape(pres.shapes.LINE, { x: x + 1.45, y: 1.95, w: 1.15, h: 0, line: { color: HEX.accent4, width: 2, endArrowType: "triangle" },
      objectName: `step ${i + 1} arrow` });
  }
  s.addText(head, { x, y: 2.5, w: 1.9, h: 0.45, fontSize: 18, bold: true, color: C.text2, align: "center", margin: 0,
    isTextBox: true, objectName: `step ${i + 1} title` });
  s.addText(body, { x, y: 2.95, w: 1.9, h: 1.1, fontSize: 14, color: C.text1, align: "center", valign: "top", margin: 0,
    isTextBox: true, objectName: `step ${i + 1} text` });
});
card(s, "model card", 0.6, 4.45, 12.1, 2.2);
s.addText("Main model (binary logit)", { x: 0.9, y: 4.6, w: 6, h: 0.45, fontSize: 18, bold: true, color: C.text2, margin: 0,
  isTextBox: true, objectName: "model heading" });
s.addText("Pr(Depression = 1) = Λ(β₀ + β₁Male + β₂Age + β₃CGPA + β₄Semester + β₅MidExpense + β₆HighExpense + β₇LowerMid + β₈Rich + β₉Single + β₁₀LiveFamily + β₁₁FriendCircle)",
  { x: 0.9, y: 5.1, w: 11.5, h: 0.75, fontSize: 15, italic: true, color: C.text1, margin: 0, isTextBox: true, objectName: "model equation" });
s.addText("Living expense in bands: Low < 10,000, Mid 10,000–14,999, High ≥ 15,000 (base: Low). Health is shown descriptively only — it almost restates the outcome. Probit and LPM are checks.",
  { x: 0.9, y: 5.85, w: 11.5, h: 0.65, fontSize: 14, color: MUTED, margin: 0, isTextBox: true, objectName: "model note" });
s.addNotes("Six steps. The main estimates are logit average marginal effects: the change in the probability of depression, in percentage points, when a factor changes. Stata code replicating every step is in the repository.");

// ========================================================================= slide 4: descriptive rates
pres.addSection({ title: "Results" });
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Results" });
const rel = D.rates.relationship;
const ratio = (rel.values[0] / rel.values[1]).toFixed(1);
s.addText(`Single students are ${ratio}× as likely to be depressed`, { placeholder: "title" });
rateChart(s, "relationship chart", D.rates.relationship, "Relationship status", 0.5, 1.4, 4.4, 2.75, HEX.accent2);
rateChart(s, "family chart", D.rates.family, "Family background", 4.95, 1.4, 4.4, 2.75, HEX.accent1);
rateChart(s, "living chart", D.rates.living, "Living style", 0.5, 4.2, 4.4, 2.7, HEX.accent1);
rateChart(s, "expense chart", D.rates.expense, "Living expense", 4.95, 4.2, 4.4, 2.7, HEX.accent2);
card(s, "takeaway card", 9.65, 1.5, 3.05, 5.2);
s.addText("What the raw data show", { x: 9.9, y: 1.7, w: 2.6, h: 0.8, fontSize: 18, bold: true, color: C.text2, margin: 0,
  isTextBox: true, objectName: "takeaway heading" });
s.addText([
  `${rel.values[0]}% of single students vs ${rel.values[1]}% in a relationship`,
  `Middle class highest at ${D.rates.family.values[1]}%`,
  `Low expense ${D.rates.expense.values[0]}% vs ${D.rates.expense.values[1]}% mid and ${D.rates.expense.values[2]}% high`,
  `Bachelors ${D.rates.living.values[0]}% vs ${D.rates.living.values[1]}% with family`,
  `Gender gap small before controls (male ${D.rates.gender.values[1]}%, female ${D.rates.gender.values[0]}%)`,
].map((t, i, a) => ({ text: t, options: { bullet: true, breakLine: i < a.length - 1 } })),
{ x: 9.9, y: 2.55, w: 2.65, h: 4.0, fontSize: 15, color: C.text1, paraSpaceAfter: 10, valign: "top", margin: 0,
  isTextBox: true, objectName: "takeaway list" });
s.addNotes("Raw depression rates by group. The relationship gap is the largest. These comparisons do not yet control for other factors; the regression does that.");

// ========================================================================= slide 5: health
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Results" });
s.addText("Self-reported health almost determines depression", { placeholder: "title" });
const h = D.health;
s.addChart(pres.charts.BAR, [
  { name: "Depressed", labels: h.labels.map((l, i) => `${l} (n=${h.n[i]})`), values: h.depressed },
  { name: "Not depressed", labels: h.labels.map((l, i) => `${l} (n=${h.n[i]})`), values: h.depressed.map((v) => +(100 - v).toFixed(1)) },
], {
  x: 0.5, y: 1.45, w: 7.6, h: 5.3, barDir: "bar", barGrouping: "percentStacked", objectName: "health chart",
  chartColors: [HEX.accent2, HEX.accent1], showValue: true, dataLabelPosition: "ctr", dataLabelFormatCode: '[<5]"";0"%"',
  dataLabelColor: "FFFFFF", showLegend: true, legendPos: "b", legendFontFace: "+mn-lt", legendFontSize: 12,
  valGridLine: { color: GRID, size: 0.75 }, catGridLine: { style: "none" }, catAxisOrientation: "maxMin",
  valAxisLabelFormatCode: "0%", barGapWidthPct: 50, showTitle: true, title: "Share of students by depression status",
  ...axisText, dataLabelColor: "FFFFFF",
});
stat(s, "critical", 8.5, 1.55, 4.2, `${h.depressed[3]}%`, "of students in critical health are depressed", C.accent2);
stat(s, "good", 8.5, 3.3, 4.2, `${h.depressed[0]}%`, "of students in good health are depressed");
s.addText("Health nearly restates the outcome (perfect prediction), so it is left out of the main model and tested only in a robustness check.",
  { x: 8.5, y: 5.1, w: 4.2, h: 1.5, fontSize: 15, color: C.text1, margin: 0, valign: "top", isTextBox: true, objectName: "health note" });
s.addNotes(`Every student in critical health is depressed and none in good health is. Adding health to the logit raises pseudo-R² from ${R.logit_pseudo_r2.toFixed(3)} to ${R.health_model_pseudo_r2.toFixed(3)}, which shows it measures the outcome itself rather than a separate cause.`);

// ========================================================================= slide 6: bivariate tests
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Results" });
s.addText("Before controls, relationships, family and expenses stand out", { placeholder: "title" });
const hdr = (t) => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: HEX.dk2 } } });
const chiRows = [[hdr("Variable"), hdr("χ²"), hdr("p-value"), hdr("Cramér's V")]].concat(
  D.chi.map((r) => {
    const sig = r.p < 0.05;
    const o = { color: sig ? HEX.dk2 : MUTED, bold: sig, fill: { color: sig ? "FBE9E3" : "FFFFFF" } };
    return [{ text: r.var, options: o }, { text: r.chi2.toFixed(2), options: o },
      { text: r.p < 0.001 ? "< 0.001" : r.p.toFixed(3), options: o }, { text: r.v.toFixed(2), options: o }];
  }));
s.addTable(chiRows, { x: 0.6, y: 1.5, w: 7.3, colW: [2.8, 1.5, 1.5, 1.5], fontSize: 15, fontFace: THEME.bodyFontFace,
  border: { type: "solid", pt: 0.75, color: GRID }, rowH: 0.5, align: "left", valign: "middle", objectName: "chi-square table" });
s.addText("Chi-square tests of independence with depression (shaded: significant at 5%)",
  { x: 0.6, y: 5.85, w: 7.3, h: 0.4, fontSize: 12, italic: true, color: MUTED, margin: 0, isTextBox: true, objectName: "table caption" });
const cg = D.ttests["CGPA"], cs = D.ttests["Current semester"];
stat(s, "cgpa", 8.4, 1.5, 4.3, `${cg.no} vs ${cg.yes}`, `mean CGPA (out of 10), not depressed vs depressed (${fmtP(cg.p)})`);
stat(s, "semester", 8.4, 3.3, 4.3, `${cs.no.toFixed(2)} vs ${cs.yes.toFixed(2)}`, `mean semester, not depressed vs depressed (${fmtP(cs.p)})`);
s.addText("Grades and stage of study are almost identical between depressed and non-depressed students.",
  { x: 8.4, y: 5.1, w: 4.3, h: 1.2, fontSize: 15, color: C.text1, margin: 0, valign: "top", isTextBox: true, objectName: "tests note" });
s.addNotes("Bivariate tests: relationship status, family background, living expense, living style and health are associated with depression; gender and friend circle are not. CGPA and semester show no difference.");

// ========================================================================= slide 7: marginal effects
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Results" });
const ame = D.ame;
const single = ame.find((r) => r.var === "Single");
s.addText(`Being single raises the probability of depression by ${Math.round(single.pp)} points`, { placeholder: "title" });
const labels = ame.map((r) => r.var);
s.addChart(pres.charts.BAR, [
  { name: "Significant at 5%", labels, values: ame.map((r) => (r.p < 0.05 ? r.pp : 0)) },
  { name: "Not significant", labels, values: ame.map((r) => (r.p < 0.05 ? 0 : r.pp)) },
], {
  x: 0.5, y: 1.4, w: 7.9, h: 5.45, barDir: "bar", barGrouping: "stacked", objectName: "marginal effects chart",
  chartColors: [HEX.accent2, HEX.accent4], showValue: true, dataLabelPosition: "ctr",
  dataLabelFormatCode: '+0.0;-0.0;;', dataLabelColor: "1B1F24", catAxisOrientation: "maxMin",
  valAxisLabelFormatCode: '+0;-0;0', valGridLine: { color: GRID, size: 0.75 }, catGridLine: { style: "none" },
  showLegend: true, legendPos: "b", legendFontFace: "+mn-lt", legendFontSize: 12, barGapWidthPct: 45,
  showTitle: true, title: "Logit average marginal effects (percentage points)", catAxisLabelPos: "low", ...axisText,
});
const pick = (name) => ame.find((r) => r.var === name);
const cards = [
  ["Single", `${signed(single.pp)} pp`, fmtP(single.p), C.accent2],
  ["Mid expense", `${signed(pick("Mid living expense").pp)} pp`, fmtP(pick("Mid living expense").p), C.accent2],
  ["High expense", `${signed(pick("High living expense").pp)} pp`, fmtP(pick("High living expense").p), C.accent2],
  ["Male", `${signed(pick("Male").pp)} pp`, fmtP(pick("Male").p), C.accent1],
  ["Rich family", `${signed(pick("Rich class").pp)} pp`, fmtP(pick("Rich class").p), MUTED],
];
cards.forEach(([label, value, p, color], i) => {
  const y = 1.45 + i * 1.08;
  card(s, `${label} card`, 8.75, y, 3.95, 0.95);
  s.addText(value, { x: 8.95, y: y + 0.1, w: 1.75, h: 0.75, fontSize: 26, bold: true, color, valign: "middle", margin: 0,
    fontFace: THEME.headFontFace, isTextBox: true, objectName: `${label} value` });
  s.addText([{ text: label, options: { bold: true, breakLine: true } }, { text: p, options: { color: MUTED } }],
    { x: 10.75, y: y + 0.1, w: 1.85, h: 0.75, fontSize: 14, color: C.text1, valign: "middle", margin: 0, isTextBox: true,
      objectName: `${label} label` });
});
s.addNotes("Average marginal effects from the logit model. Coral bars are significant at 5%. Mid and high living expense are measured against low expense. Rich family is significant only at 10%. CGPA, age, semester, friend circle and living with family are not significant once the other factors are held constant.");

// ========================================================================= slide 8: fit and ROC
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Results" });
s.addText("The model fits well and separates students reasonably", { placeholder: "title" });
stat(s, "lr", 0.6, 1.45, 2.85, R.logit_lr_chi2.toFixed(1), `LR χ²(${R.logit_lr_df}), ${fmtP(R.logit_lr_p)}`);
stat(s, "pseudo", 3.65, 1.45, 2.85, R.logit_pseudo_r2.toFixed(3), "McFadden pseudo-R²");
stat(s, "auc", 6.7, 1.45, 2.85, R.logit_auc.toFixed(3), "area under ROC curve", C.accent2);
stat(s, "hl", 9.85, 1.45, 2.85, R.logit_hl_p.toFixed(3), "Hosmer–Lemeshow p (good fit)");
const grid = D.roc.grid.map((g) => g.toFixed(2));
s.addChart(pres.charts.LINE, [
  { name: `Logit (AUC ${R.logit_auc.toFixed(3)})`, labels: grid, values: D.roc.logit },
  { name: `Probit (AUC ${R.probit_auc.toFixed(3)})`, labels: grid, values: D.roc.probit },
  { name: "Random guess", labels: grid, values: D.roc.grid },
], {
  x: 0.5, y: 3.2, w: 7.4, h: 3.65, objectName: "ROC chart", chartColors: [HEX.accent2, HEX.accent1, "B8C2C7"],
  lineSize: 2, lineDataSymbol: "none", showLegend: true, legendPos: "r", legendFontFace: "+mn-lt", legendFontSize: 12,
  valAxisMaxVal: 1, valAxisMinVal: 0, valAxisMajorUnit: 0.25, valAxisLabelFormatCode: "0.00", catAxisLabelFrequency: 10,
  valGridLine: { color: GRID, size: 0.75 }, catGridLine: { style: "none" }, showTitle: true,
  title: "ROC curve: true positive rate vs false positive rate", ...axisText,
});
card(s, "classification card", 8.3, 3.3, 4.4, 3.45);
s.addText("Classification", { x: 8.55, y: 3.45, w: 3.9, h: 0.45, fontSize: 18, bold: true, color: C.text2, margin: 0,
  isTextBox: true, objectName: "classification heading" });
s.addText([
  `${(100 * R.logit_correct_05).toFixed(1)}% classified correctly at cut-off 0.5`,
  `Specificity ${(100 * R.logit_spec_05).toFixed(1)}%, sensitivity ${(100 * R.logit_sens_05).toFixed(1)}%`,
  `At cut-off ${D.classification[1].Cutoff.toFixed(3)} (sample share) the model finds ${(100 * D.classification[1].Sensitivity).toFixed(1)}% of depressed students`,
].map((t, i, a) => ({ text: t, options: { bullet: true, breakLine: i < a.length - 1 } })),
{ x: 8.55, y: 3.95, w: 3.95, h: 2.7, fontSize: 15, color: C.text1, paraSpaceAfter: 8, valign: "top", margin: 0,
  isTextBox: true, objectName: "classification list" });
s.addNotes(`The logit model is jointly significant, calibrated (Hosmer–Lemeshow p = ${R.logit_hl_p.toFixed(3)}) and has AUC ${R.logit_auc.toFixed(3)}. Probit gives almost the same fit.`);

// ========================================================================= slide 9: robustness and hypotheses
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Results" });
s.addText("Results largely hold across logit, probit and LPM", { placeholder: "title" });
const cmp = D.compare;
const cmpLabels = cmp.vars.map((v) => (v === "Rich class" ? "Rich family" : v.replace(" living expense", " expense")));
const scale = cmp.vars.map(() => 1);
s.addChart(pres.charts.BAR, [
  { name: "Logit AME", labels: cmpLabels, values: cmp.logit.map((v, i) => +(v * scale[i]).toFixed(1)) },
  { name: "Probit AME", labels: cmpLabels, values: cmp.probit.map((v, i) => +(v * scale[i]).toFixed(1)) },
  { name: "LPM", labels: cmpLabels, values: cmp.lpm.map((v, i) => +(v * scale[i]).toFixed(1)) },
], {
  x: 0.5, y: 1.4, w: 6.9, h: 5.45, barDir: "bar", objectName: "robustness chart",
  chartColors: [HEX.accent2, HEX.accent1, HEX.accent4], showValue: true, dataLabelPosition: "outEnd",
  dataLabelFormatCode: '+0;-0;0', dataLabelFontSize: 11, catAxisOrientation: "maxMin", catAxisLabelPos: "low",
  valAxisLabelFormatCode: '+0;-0;0', valGridLine: { color: GRID, size: 0.75 }, catGridLine: { style: "none" },
  showLegend: true, legendPos: "b", legendFontFace: "+mn-lt", legendFontSize: 12, barGapWidthPct: 40,
  showTitle: true, title: "Effect on probability of depression (percentage points)", ...axisText, dataLabelFontSize: 11,
});
const resColor = (r) => (r === "Supported" ? HEX.accent1 : r.startsWith("Weakly") ? "B7791F" : MUTED);
const resText = (r) => (r === "Supported" ? "Supported" : r.startsWith("Weakly") ? "Weakly (10%)" : "Not supported");
const hRows = [[hdr("H"), hdr("Variable"), hdr("Result")]].concat(D.hypotheses.map((r) => [
  { text: r.h, options: { bold: true, color: HEX.dk2 } }, { text: r.var, options: { color: "1B1F24" } },
  { text: resText(r.result), options: { bold: true, color: resColor(r.result) } }]));
s.addTable(hRows, { x: 7.75, y: 1.5, w: 4.95, colW: [0.75, 2.4, 1.8], fontSize: 14, fontFace: THEME.bodyFontFace,
  border: { type: "solid", pt: 0.75, color: GRID }, rowH: 0.46, valign: "middle", objectName: "hypothesis table" });
s.addText("Gender subsamples: being single is a strong risk factor for both men and women; the expense effect appears mainly among men.",
  { x: 7.75, y: 5.95, w: 4.95, h: 0.8, fontSize: 14, italic: true, color: MUTED, margin: 0, valign: "top", isTextBox: true,
    objectName: "subsample note" });
s.addNotes("Logit and probit give nearly identical effects; the LPM gives the same single and male effects but smaller expense effects, significant only at 10%. H3 (single), H5 (high living expense) and H7 (gender) are supported at 5%. CGPA, friend circle, family background and living with family are not supported at 5%.");

// ========================================================================= slide 10: conclusion
pres.addSection({ title: "Conclusion" });
s = pres.addSlide({ masterName: "CLOSING_DARK", sectionTitle: "Conclusion" });
s.addText("Relationships and money matter more than grades", { placeholder: "title" });
const cols = [
  ["Key findings", [
    `Single students: ${signed(single.pp)} points`,
    `Mid / high living expense: ${signed(pick("Mid living expense").pp)} / ${signed(pick("High living expense").pp)} points`,
    `Women more at risk after controls (${signed(-pick("Male").pp)} points)`,
    "CGPA, friend circle, living style: no effect",
  ]],
  ["For universities", [
    "Programmes that reduce loneliness and build close ties",
    "Financial guidance that also reaches middle-income students",
    "Gender-sensitive outreach and screening",
    "Use simple risk scores to prioritise screening",
  ]],
  ["Limits & next steps", [
    "Cross-sectional: associations, not causes",
    "Self-reported yes/no depression measure",
    "Follow students over time; use PHQ-9 scale",
    "Larger, more diverse samples",
  ]],
];
cols.forEach(([head, items], i) => {
  const x = 0.6 + i * 4.1;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.5, w: 3.85, h: 5.1, rectRadius: 0.12, fill: { color: "1D4A75" },
    line: { color: "1D4A75" }, objectName: `${head} card` });
  s.addText(head, { x: x + 0.3, y: 1.7, w: 3.3, h: 0.55, fontSize: 20, bold: true, color: i === 0 ? "F4A98F" : "BFE3E8",
    margin: 0, fontFace: THEME.headFontFace, isTextBox: true, objectName: `${head} heading` });
  s.addText(items.map((t, j) => ({ text: t, options: { bullet: true, breakLine: j < items.length - 1 } })),
    { x: x + 0.3, y: 2.4, w: 3.3, h: 4.0, fontSize: 16, color: "FFFFFF", paraSpaceAfter: 12, valign: "top", margin: 0,
      isTextBox: true, objectName: `${head} list` });
});
s.addNotes("Close by stressing the main message: student mental health support should focus on social connection and financial wellbeing, not only on academic help. Thank the audience and invite questions.");

(async () => {
  await pres.writeFile({ fileName: OUT });
  if (APPLY_THEME) {
    const { applyTheme } = require(APPLY_THEME);
    await applyTheme(OUT, THEME);
  }
  console.log("Saved", path.relative(ROOT, OUT));
})();
