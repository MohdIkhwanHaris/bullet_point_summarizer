const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");

// Icons
const { FaBrain, FaFilter, FaCogs, FaChartBar, FaShieldAlt, FaCode, FaRocket, FaDatabase, FaCheck, FaTimes, FaArrowRight, FaSearch, FaLayerGroup } = require("react-icons/fa");
const { MdAnalytics } = require("react-icons/md");

// Color palette — dark enterprise SaaS
const C = {
  bg:       "0D1117",  // near-black
  bg2:      "161B22",  // card bg
  bg3:      "1F2937",  // lighter card
  accent:   "00D4FF",  // cyan glow
  accent2:  "7C3AED",  // purple
  accent3:  "10B981",  // green
  accent4:  "F59E0B",  // amber
  accent5:  "EF4444",  // red
  white:    "F0F6FC",
  muted:    "8B949E",
  border:   "30363D",
};

async function iconPng(IconComp, color, size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(
    React.createElement(IconComp, { color: "#" + color, size: String(size) })
  );
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

function makeShadow() {
  return { type: "outer", blur: 12, offset: 3, angle: 135, color: "000000", opacity: 0.4 };
}

// Draw a glowing card
function card(slide, x, y, w, h, opts = {}) {
  slide.addShape(pres.shapes.RECTANGLE, {
    x, y, w, h,
    fill: { color: opts.fill || C.bg2 },
    line: { color: opts.border || C.border, width: 1 },
    shadow: makeShadow(),
  });
  if (opts.accent) {
    slide.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 0.05, h,
      fill: { color: opts.accent },
      line: { color: opts.accent, width: 0 },
    });
  }
}

let pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "Enterprise Extractive Summarization Engine";
pres.author = "Muhammad Ikhwan Haris bin Mohd Ezman";

// ─────────────────────────────────────────────
// SLIDE 1: Title
// ─────────────────────────────────────────────
async function slide1() {
  let sl = pres.addSlide();
  sl.background = { color: C.bg };

  // Decorative grid dots
  for (let r = 0; r < 6; r++) {
    for (let c = 0; c < 14; c++) {
      sl.addShape(pres.shapes.OVAL, {
        x: 0.4 + c * 0.7, y: 0.3 + r * 0.9, w: 0.04, h: 0.04,
        fill: { color: C.border }, line: { color: C.border, width: 0 },
      });
    }
  }

  // Glowing geometric nodes top-right
  const nodeData = await iconPng(FaBrain, C.accent, 256);
  sl.addImage({ data: nodeData, x: 7.8, y: 0.3, w: 1.8, h: 1.8 });

  // Top accent line
  sl.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 0.25, w: 4.5, h: 0.04,
    fill: { color: C.accent }, line: { color: C.accent, width: 0 },
  });

  // Tag label
  sl.addText("UNIVERSITI MALAYA  ·  NLP COURSEWORK", {
    x: 0.5, y: 0.35, w: 6, h: 0.3,
    fontSize: 9, color: C.accent, bold: true, charSpacing: 3, margin: 0,
  });

  // Main title
  sl.addText("Enterprise Extractive", {
    x: 0.5, y: 1.0, w: 9, h: 0.85,
    fontSize: 48, color: C.white, bold: true, fontFace: "Arial Black", margin: 0,
  });
  sl.addText("Summarization Engine", {
    x: 0.5, y: 1.8, w: 9, h: 0.85,
    fontSize: 48, color: C.accent, bold: true, fontFace: "Arial Black", margin: 0,
  });

  // Subtitle
  sl.addText("A Mathematical NLP Pipeline with MMR Redundancy Filtering", {
    x: 0.5, y: 2.75, w: 8.5, h: 0.45,
    fontSize: 16, color: C.muted, italic: true, margin: 0,
  });

  // Divider
  sl.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 3.3, w: 9, h: 0.015,
    fill: { color: C.border }, line: { color: C.border, width: 0 },
  });

  // Presenter card
  card(sl, 0.5, 3.5, 5.8, 1.7, { fill: C.bg2, accent: C.accent2 });
  sl.addText("PRESENTER", {
    x: 0.7, y: 3.6, w: 5, h: 0.25,
    fontSize: 8, color: C.accent2, bold: true, charSpacing: 3, margin: 0,
  });
  sl.addText("Muhammad Ikhwan Haris bin Mohd Ezman", {
    x: 0.7, y: 3.88, w: 5.5, h: 0.35,
    fontSize: 15, color: C.white, bold: true, margin: 0,
  });
  sl.addText("Group Members: [To be added]", {
    x: 0.7, y: 4.28, w: 5.5, h: 0.28,
    fontSize: 11, color: C.muted, margin: 0,
  });
  sl.addText("Universiti Malaya  ·  2025", {
    x: 0.7, y: 4.6, w: 5.5, h: 0.25,
    fontSize: 10, color: C.muted, italic: true, margin: 0,
  });

  // Stat badges right
  const stats = [
    { val: "4", label: "Scoring Factors" },
    { val: "MMR", label: "Redundancy Filter" },
    { val: "100%", label: "Factual Accuracy" },
  ];
  stats.forEach((s, i) => {
    const sx = 7.0, sy = 3.5 + i * 0.6;
    card(sl, sx, sy, 2.7, 0.52, { fill: C.bg3, border: C.border });
    sl.addText(s.val, { x: sx + 0.12, y: sy + 0.06, w: 0.9, h: 0.4, fontSize: 18, color: C.accent, bold: true, margin: 0 });
    sl.addText(s.label, { x: sx + 1.0, y: sy + 0.13, w: 1.6, h: 0.28, fontSize: 10, color: C.muted, margin: 0 });
  });
}

// ─────────────────────────────────────────────
// SLIDE 2: Problem Statement
// ─────────────────────────────────────────────
async function slide2() {
  let sl = pres.addSlide();
  sl.background = { color: C.bg };

  // Header
  sl.addText("02", { x: 0.4, y: 0.25, w: 1, h: 0.5, fontSize: 11, color: C.accent, bold: true, margin: 0, charSpacing: 2 });
  sl.addText("The Problem Statement", {
    x: 0.4, y: 0.55, w: 9, h: 0.55,
    fontSize: 28, color: C.white, bold: true, fontFace: "Arial Black", margin: 0,
  });
  sl.addText("Black-Box AI vs. Mathematical Transparency", {
    x: 0.4, y: 1.1, w: 9, h: 0.3,
    fontSize: 13, color: C.muted, italic: true, margin: 0,
  });

  // LEFT PANEL — Black Box
  card(sl, 0.4, 1.5, 4.25, 3.7, { fill: C.bg2, accent: C.accent5 });
  sl.addText("⚠  BLACK-BOX AI (ChatGPT / LLMs)", {
    x: 0.6, y: 1.6, w: 3.9, h: 0.35,
    fontSize: 10.5, color: C.accent5, bold: true, charSpacing: 1, margin: 0,
  });
  sl.addShape(pres.shapes.RECTANGLE, {
    x: 0.55, y: 2.0, w: 3.9, h: 0.015,
    fill: { color: C.border }, line: { color: C.border, width: 0 },
  });

  const leftProblems = [
    { icon: "🎭", title: "Hallucination Problem", body: "Generative models fabricate or alter meaning. The output may sound fluent but is not verifiably traceable to the source." },
    { icon: "🔁", title: "Echo Chamber Effect", body: "Without diversity filters, models repeat the same high-scoring point with slight rephrasing — wasting summary space." },
    { icon: "🔒", title: "Zero Explainability", body: "No score breakdown. Users must blindly trust the AI's selection — unacceptable in enterprise data curation." },
  ];
  leftProblems.forEach((p, i) => {
    const py = 2.1 + i * 1.05;
    sl.addText(p.icon + "  " + p.title, {
      x: 0.6, y: py, w: 3.9, h: 0.3,
      fontSize: 11, color: C.white, bold: true, margin: 0,
    });
    sl.addText(p.body, {
      x: 0.6, y: py + 0.3, w: 3.9, h: 0.6,
      fontSize: 9.5, color: C.muted, margin: 0,
    });
  });

  // Arrow
  sl.addShape(pres.shapes.RECTANGLE, {
    x: 4.75, y: 3.05, w: 0.5, h: 0.04,
    fill: { color: C.accent }, line: { color: C.accent, width: 0 },
  });
  sl.addText("→", { x: 4.72, y: 2.9, w: 0.6, h: 0.4, fontSize: 18, color: C.accent, align: "center", margin: 0 });

  // RIGHT PANEL — Our Solution
  card(sl, 5.35, 1.5, 4.25, 3.7, { fill: C.bg2, accent: C.accent3 });
  sl.addText("✓  OUR EXTRACTIVE APPROACH", {
    x: 5.55, y: 1.6, w: 3.9, h: 0.35,
    fontSize: 10.5, color: C.accent3, bold: true, charSpacing: 1, margin: 0,
  });
  sl.addShape(pres.shapes.RECTANGLE, {
    x: 5.4, y: 2.0, w: 3.9, h: 0.015,
    fill: { color: C.border }, line: { color: C.border, width: 0 },
  });

  const rightSolutions = [
    { icon: "✅", title: "100% Factual Accuracy", body: "We select original sentences verbatim. No words are changed, generated, or hallucinated — ever." },
    { icon: "🎯", title: "MMR Diversity Filter", body: "Maximal Marginal Relevance actively penalises redundant sentences using cosine similarity scoring." },
    { icon: "📊", title: "Full Audit Trail", body: "Every sentence receives an exact TF-IDF, POS, Semantic & Structural score — all visible to the user." },
  ];
  rightSolutions.forEach((p, i) => {
    const py = 2.1 + i * 1.05;
    sl.addText(p.icon + "  " + p.title, {
      x: 5.55, y: py, w: 3.9, h: 0.3,
      fontSize: 11, color: C.white, bold: true, margin: 0,
    });
    sl.addText(p.body, {
      x: 5.55, y: py + 0.3, w: 3.9, h: 0.6,
      fontSize: 9.5, color: C.muted, margin: 0,
    });
  });

  // Bottom goal banner
  card(sl, 0.4, 5.2, 9.2, 0.25, { fill: C.bg3, border: C.accent });
  sl.addText("GOAL: Score importance mathematically + actively filter redundancy → verifiable, diverse, enterprise-grade summaries", {
    x: 0.55, y: 5.22, w: 8.9, h: 0.22,
    fontSize: 9, color: C.accent, bold: true, margin: 0, align: "center",
  });
}

// ─────────────────────────────────────────────
// SLIDE 3: System Architecture
// ─────────────────────────────────────────────
async function slide3() {
  let sl = pres.addSlide();
  sl.background = { color: C.bg };

  sl.addText("03", { x: 0.4, y: 0.25, w: 1, h: 0.35, fontSize: 11, color: C.accent, bold: true, margin: 0, charSpacing: 2 });
  sl.addText("System Architecture", {
    x: 0.4, y: 0.55, w: 9, h: 0.55,
    fontSize: 28, color: C.white, bold: true, fontFace: "Arial Black", margin: 0,
  });
  sl.addText("The 4-Stage Modular Pipeline Overview", {
    x: 0.4, y: 1.1, w: 9, h: 0.3,
    fontSize: 13, color: C.muted, italic: true, margin: 0,
  });

  // Pipeline boxes
  const stages = [
    { num: "01", title: "Raw Text Input", sub: "User paste / upload", color: C.accent2 },
    { num: "02", title: "Linguistic Preprocessing", sub: "spaCy · Lemmatization · Stop-word removal", color: C.accent },
    { num: "03", title: "Multi-Factor Heuristic Scoring", sub: "TF-IDF · POS · Semantics · Structural", color: C.accent4 },
    { num: "04", title: "MMR Redundancy Filter", sub: "Cosine similarity diversity enforcement", color: C.accent3 },
  ];

  const boxW = 1.95, boxH = 2.2, startX = 0.4, y = 1.6, gap = 0.5;

  stages.forEach((s, i) => {
    const x = startX + i * (boxW + gap);
    card(sl, x, y, boxW, boxH, { fill: C.bg2, border: s.color });
    // Top accent strip
    sl.addShape(pres.shapes.RECTANGLE, {
      x, y, w: boxW, h: 0.08,
      fill: { color: s.color }, line: { color: s.color, width: 0 },
    });
    sl.addText(s.num, {
      x: x + 0.12, y: y + 0.15, w: 0.5, h: 0.4,
      fontSize: 20, color: s.color, bold: true, fontFace: "Arial Black", margin: 0,
    });
    sl.addText(s.title, {
      x: x + 0.12, y: y + 0.6, w: boxW - 0.2, h: 0.7,
      fontSize: 11, color: C.white, bold: true, margin: 0,
    });
    sl.addText(s.sub, {
      x: x + 0.12, y: y + 1.35, w: boxW - 0.2, h: 0.65,
      fontSize: 9, color: C.muted, margin: 0,
    });

    // Arrow between boxes
    if (i < stages.length - 1) {
      const ax = x + boxW + 0.05;
      sl.addText("›", {
        x: ax, y: y + 0.85, w: gap, h: 0.5,
        fontSize: 22, color: s.color, align: "center", bold: true, margin: 0,
      });
    }
  });

  // Output box
  const outX = 0.4 + 4 * (boxW + gap) - gap * 0.5;
  card(sl, 8.1, y, 1.7, boxH, { fill: C.bg3, border: C.accent3 });
  sl.addShape(pres.shapes.RECTANGLE, {
    x: 8.1, y, w: 1.7, h: 0.08,
    fill: { color: C.accent3 }, line: { color: C.accent3, width: 0 },
  });
  sl.addText("OUTPUT", { x: 8.15, y: y + 0.2, w: 1.6, h: 0.3, fontSize: 8, color: C.accent3, bold: true, charSpacing: 2, margin: 0 });
  sl.addText([
    { text: "•", options: { bullet: true, breakLine: true } },
    { text: " Ranked bullet", options: { breakLine: true } },
    { text: "•", options: { bullet: true, breakLine: true } },
    { text: " points", options: { breakLine: true } },
    { text: "•", options: { bullet: true, breakLine: true } },
    { text: " Diverse", options: { breakLine: true } },
    { text: "•", options: { bullet: true } },
    { text: " Verified" },
  ], { x: 8.2, y: y + 0.6, w: 1.5, h: 1.5, fontSize: 9, color: C.muted, margin: 0 });

  // Bottom description
  card(sl, 0.4, 4.0, 9.2, 1.3, { fill: C.bg2, accent: C.accent });
  sl.addText("Pipeline Philosophy", {
    x: 0.6, y: 4.08, w: 8.8, h: 0.3,
    fontSize: 12, color: C.accent, bold: true, margin: 0,
  });
  sl.addText([
    { text: "Each stage acts as a dedicated filter. ", options: { bold: true, color: C.white } },
    { text: "Stage 1 cleans linguistic noise. Stage 2 assigns mathematical importance across 4 independent lenses. Stage 3 enforces diversity using cosine similarity. No stage can be bypassed — the architecture is modular and each component is independently testable.", options: { color: C.muted } }
  ], { x: 0.6, y: 4.4, w: 9.0, h: 0.8, fontSize: 10.5, margin: 0 });
}

// ─────────────────────────────────────────────
// SLIDE 4: Linguistic Preprocessing
// ─────────────────────────────────────────────
async function slide4() {
  let sl = pres.addSlide();
  sl.background = { color: C.bg };

  sl.addText("04", { x: 0.4, y: 0.25, w: 1, h: 0.35, fontSize: 11, color: C.accent, bold: true, margin: 0, charSpacing: 2 });
  sl.addText("Phase 1 — Linguistic Preprocessing", {
    x: 0.4, y: 0.55, w: 9.2, h: 0.55,
    fontSize: 26, color: C.white, bold: true, fontFace: "Arial Black", margin: 0,
  });
  sl.addText("Powered by spaCy  en_core_web_sm  ·  Noise elimination before math", {
    x: 0.4, y: 1.1, w: 9, h: 0.3,
    fontSize: 12, color: C.muted, italic: true, margin: 0,
  });

  // Terminal / code block
  card(sl, 0.4, 1.55, 5.7, 3.6, { fill: "0D1117", border: C.accent });
  // Terminal header bar
  sl.addShape(pres.shapes.RECTANGLE, {
    x: 0.4, y: 1.55, w: 5.7, h: 0.38,
    fill: { color: "161B22" }, line: { color: C.border, width: 0 },
  });
  sl.addText("● ● ●", { x: 0.6, y: 1.6, w: 1, h: 0.28, fontSize: 12, color: C.muted, margin: 0 });
  sl.addText("preprocess.py  —  nlp_pipeline", { x: 1.5, y: 1.62, w: 3.5, h: 0.25, fontSize: 9, color: C.muted, align: "center", margin: 0 });

  const codeLines = [
    { txt: "import spacy", col: C.accent2 },
    { txt: "nlp = spacy.load('en_core_web_sm')", col: C.white },
    { txt: "", col: C.muted },
    { txt: "def preprocess(text):", col: C.accent4 },
    { txt: "    doc = nlp(text)", col: C.white },
    { txt: "    tokens = [", col: C.white },
    { txt: "        token.lemma_.lower()", col: C.accent3 },
    { txt: "        for token in doc", col: C.white },
    { txt: "        if not token.is_stop", col: C.accent },
    { txt: "        and not token.is_punct", col: C.accent },
    { txt: "    ]", col: C.white },
    { txt: "    return tokens  # cleaned!", col: C.muted },
  ];
  codeLines.forEach((line, i) => {
    sl.addText(line.txt, {
      x: 0.55, y: 2.05 + i * 0.24, w: 5.4, h: 0.24,
      fontSize: 9.5, color: line.col, fontFace: "Consolas", margin: 0,
    });
  });

  // Right side — transformation steps
  const steps = [
    { icon: "①", label: "Tokenisation", desc: "Split raw text into individual word tokens using spaCy's built-in tokenizer.", color: C.accent },
    { icon: "②", label: "Stop-Word Removal", desc: "Remove grammatical 'noise' (the, is, and…) that would distort scoring frequency.", color: C.accent4 },
    { icon: "③", label: "Lemmatisation", desc: "Reduce all inflections to root form: running, ran, runs → run", color: C.accent3 },
    { icon: "④", label: "POS Annotation", desc: "Tag each surviving token as NOUN, VERB, ADJ etc. Used in Phase 2 scoring.", color: C.accent2 },
  ];

  steps.forEach((s, i) => {
    const sy = 1.55 + i * 0.92;
    card(sl, 6.35, sy, 3.3, 0.82, { fill: C.bg2, accent: s.color });
    sl.addText(s.icon, { x: 6.45, y: sy + 0.08, w: 0.45, h: 0.55, fontSize: 18, color: s.color, bold: true, margin: 0 });
    sl.addText(s.label, { x: 6.92, y: sy + 0.08, w: 2.6, h: 0.28, fontSize: 11, color: C.white, bold: true, margin: 0 });
    sl.addText(s.desc, { x: 6.92, y: sy + 0.38, w: 2.65, h: 0.38, fontSize: 8.5, color: C.muted, margin: 0 });
  });

  // Why spaCy banner
  card(sl, 0.4, 5.22, 9.2, 0.25, { fill: C.bg3, border: C.accent });
  sl.addText("Why spaCy?  Industrial-strength NLP · Pre-trained transformer model · 50x faster than NLTK on lemmatisation benchmarks", {
    x: 0.55, y: 5.24, w: 8.9, h: 0.2,
    fontSize: 8.5, color: C.accent, bold: true, align: "center", margin: 0,
  });
}

// ─────────────────────────────────────────────
// SLIDE 5: Heuristic Scoring Matrix
// ─────────────────────────────────────────────
async function slide5() {
  let sl = pres.addSlide();
  sl.background = { color: C.bg };

  sl.addText("05", { x: 0.4, y: 0.25, w: 1, h: 0.35, fontSize: 11, color: C.accent, bold: true, margin: 0, charSpacing: 2 });
  sl.addText("Phase 2 — Heuristic Scoring Matrix", {
    x: 0.4, y: 0.55, w: 9.2, h: 0.55,
    fontSize: 26, color: C.white, bold: true, fontFace: "Arial Black", margin: 0,
  });
  sl.addText("The 4 Lenses: α · β · γ · δ  —  Every sentence scored across four independent dimensions", {
    x: 0.4, y: 1.1, w: 9.2, h: 0.3,
    fontSize: 12, color: C.muted, italic: true, margin: 0,
  });

  // Formula
  card(sl, 0.4, 1.5, 9.2, 0.72, { fill: C.bg3, border: C.accent });
  sl.addText("Score(s)  =  α · TF-IDF(s)  +  β · POS(s)  +  γ · Semantic(s)  +  δ · Structural(s)", {
    x: 0.6, y: 1.58, w: 8.9, h: 0.55,
    fontSize: 15, color: C.accent, bold: true, fontFace: "Consolas", align: "center", margin: 0,
  });

  // 4 quadrant cards
  const lenses = [
    {
      greek: "α", weight: "0.40", label: "TF-IDF", sub: "Statistical Baseline",
      body: "Term Frequency × Inverse Document Frequency identifies the statistically rarest and therefore most informative keywords in the document. Sentences rich in these terms score highest.",
      badge: "40% weight", color: C.accent,
    },
    {
      greek: "β", weight: "0.20", label: "POS Tagging", sub: "Grammatical Filter",
      body: "spaCy assigns Part-of-Speech tags. Sentences densely packed with NOUN and ACTION VERB tokens receive bonus scores — they carry factual, actionable content.",
      badge: "20% weight", color: C.accent2,
    },
    {
      greek: "γ", weight: "0.20", label: "Semantic / WordNet", sub: "Lexical Gap Bridging",
      body: "Wu-Palmer taxonomy similarity from WordNet bridges synonyms: 'car' ≈ 'automobile'. Prevents penalising semantically identical sentences that use different vocabulary.",
      badge: "20% weight", color: C.accent4,
    },
    {
      greek: "δ", weight: "0.20", label: "Structural Bonus", sub: "Positional Heuristic",
      body: "Opening sentences (title, first paragraph) statistically carry topic-defining information. A positional multiplier rewards them — matching human reading behaviour.",
      badge: "20% weight", color: C.accent3,
    },
  ];

  const qW = 4.55, qH = 1.9;
  lenses.forEach((l, i) => {
    const x = 0.4 + (i % 2) * (qW + 0.1);
    const y = 2.35 + Math.floor(i / 2) * (qH + 0.12);
    card(sl, x, y, qW, qH, { fill: C.bg2, accent: l.color });

    // Weight badge
    card(sl, x + qW - 1.2, y + 0.1, 1.1, 0.38, { fill: C.bg3, border: l.color });
    sl.addText(l.badge, { x: x + qW - 1.18, y: y + 0.17, w: 1.06, h: 0.25, fontSize: 8.5, color: l.color, align: "center", bold: true, margin: 0 });

    sl.addText(l.greek, { x: x + 0.12, y: y + 0.08, w: 0.55, h: 0.55, fontSize: 28, color: l.color, bold: true, fontFace: "Arial Black", margin: 0 });
    sl.addText(l.label, { x: x + 0.7, y: y + 0.1, w: 2.5, h: 0.3, fontSize: 14, color: C.white, bold: true, margin: 0 });
    sl.addText(l.sub, { x: x + 0.7, y: y + 0.4, w: 2.5, h: 0.25, fontSize: 9, color: l.color, margin: 0 });
    sl.addText(l.body, { x: x + 0.12, y: y + 0.72, w: qW - 0.22, h: 1.05, fontSize: 9, color: C.muted, margin: 0 });
  });
}

// ─────────────────────────────────────────────
// SLIDE 6: MMR Diversity Engine
// ─────────────────────────────────────────────
async function slide6() {
  let sl = pres.addSlide();
  sl.background = { color: C.bg };

  sl.addText("06", { x: 0.4, y: 0.25, w: 1, h: 0.35, fontSize: 11, color: C.accent, bold: true, margin: 0, charSpacing: 2 });
  sl.addText("Phase 3 — The Diversity Engine", {
    x: 0.4, y: 0.55, w: 9.2, h: 0.55,
    fontSize: 26, color: C.white, bold: true, fontFace: "Arial Black", margin: 0,
  });
  sl.addText("Maximal Marginal Relevance (MMR)  ·  The Anti-Redundancy Bouncer", {
    x: 0.4, y: 1.1, w: 9.2, h: 0.3,
    fontSize: 12, color: C.muted, italic: true, margin: 0,
  });

  // MMR formula
  card(sl, 0.4, 1.5, 9.2, 0.62, { fill: C.bg3, border: C.accent3 });
  sl.addText("MMR(sᵢ)  =  λ · Score(sᵢ)  –  (1 – λ) · max  sim(sᵢ, sⱼ)   ∀sⱼ ∈ Selected", {
    x: 0.6, y: 1.56, w: 8.9, h: 0.5,
    fontSize: 13, color: C.accent3, bold: true, fontFace: "Consolas", align: "center", margin: 0,
  });

  // Visual flow of MMR loop
  // Candidate pool
  card(sl, 0.4, 2.25, 2.8, 2.9, { fill: C.bg2, border: C.border });
  sl.addText("CANDIDATE POOL", {
    x: 0.5, y: 2.32, w: 2.6, h: 0.28,
    fontSize: 8, color: C.muted, bold: true, charSpacing: 2, margin: 0,
  });
  const sentences = ["S1 — High score", "S2 — High score", "S3 — Similar to S1", "S4 — Unique content", "S5 — Moderate"];
  const sColors = [C.accent3, C.accent2, C.accent5, C.accent3, C.muted];
  sentences.forEach((s, i) => {
    const sy = 2.7 + i * 0.44;
    sl.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y: sy, w: 2.55, h: 0.35,
      fill: { color: C.bg3 }, line: { color: sColors[i], width: 1 },
    });
    sl.addText(s, { x: 0.55, y: sy + 0.05, w: 2.45, h: 0.26, fontSize: 9, color: sColors[i], margin: 0 });
  });

  // Filter / MMR Shield  
  card(sl, 3.5, 2.8, 1.5, 1.7, { fill: C.bg3, border: C.accent3 });
  sl.addText("MMR\nFILTER", {
    x: 3.55, y: 3.05, w: 1.4, h: 0.9,
    fontSize: 14, color: C.accent3, bold: true, align: "center", fontFace: "Arial Black", margin: 0,
  });
  sl.addText("🛡", { x: 3.72, y: 2.85, w: 1.0, h: 0.45, fontSize: 22, align: "center", margin: 0 });
  sl.addText("cosine_similarity", { x: 3.5, y: 4.1, w: 1.5, h: 0.3, fontSize: 7.5, color: C.muted, align: "center", margin: 0 });

  // Arrows
  sl.addText("→", { x: 3.1, y: 3.3, w: 0.4, h: 0.4, fontSize: 18, color: C.accent3, align: "center", bold: true, margin: 0 });
  sl.addText("→", { x: 5.1, y: 3.0, w: 0.4, h: 0.4, fontSize: 18, color: C.accent3, align: "center", bold: true, margin: 0 });
  sl.addText("✗", { x: 5.1, y: 3.6, w: 0.4, h: 0.4, fontSize: 18, color: C.accent5, align: "center", bold: true, margin: 0 });

  // ACCEPTED
  card(sl, 5.6, 2.5, 2.6, 1.2, { fill: C.bg2, border: C.accent3 });
  sl.addText("✓  ACCEPTED", { x: 5.7, y: 2.58, w: 2.4, h: 0.3, fontSize: 10, color: C.accent3, bold: true, margin: 0 });
  sl.addText("S1 — High score, unique", { x: 5.7, y: 2.9, w: 2.4, h: 0.25, fontSize: 9.5, color: C.white, margin: 0 });
  sl.addText("S4 — Unique content", { x: 5.7, y: 3.2, w: 2.4, h: 0.25, fontSize: 9.5, color: C.white, margin: 0 });

  // REJECTED
  card(sl, 5.6, 3.85, 2.6, 0.9, { fill: C.bg2, border: C.accent5 });
  sl.addText("✗  PENALISED / SKIPPED", { x: 5.7, y: 3.92, w: 2.4, h: 0.3, fontSize: 10, color: C.accent5, bold: true, margin: 0 });
  sl.addText("S3 — sim(S3,S1) too high", { x: 5.7, y: 4.22, w: 2.4, h: 0.25, fontSize: 9.5, color: C.muted, margin: 0 });

  // Right: How it works
  const steps = [
    { n: "1", t: "Select Best", d: "Pick the sentence with the highest raw score → slot #1 confirmed." },
    { n: "2", t: "Cosine Check", d: "For each remaining candidate, compute cosine similarity against all already-selected sentences." },
    { n: "3", t: "Penalty Applied", d: "Subtract the similarity score from the relevance score. High similarity = heavy penalty." },
    { n: "4", t: "Fill Next Slot", d: "Select the highest-scoring candidate after penalty → guaranteed diverse." },
  ];
  steps.forEach((s, i) => {
    const sy = 2.2 + i * 0.85;
    card(sl, 8.45, sy, 1.4, 0.75, { fill: C.bg2, border: C.accent });
    sl.addText(s.n, { x: 8.5, y: sy + 0.05, w: 0.3, h: 0.4, fontSize: 18, color: C.accent, bold: true, fontFace: "Arial Black", margin: 0 });
    sl.addText(s.t, { x: 8.82, y: sy + 0.07, w: 0.95, h: 0.28, fontSize: 9, color: C.white, bold: true, margin: 0 });
    sl.addText(s.d, { x: 8.5, y: sy + 0.38, w: 1.3, h: 0.35, fontSize: 7.5, color: C.muted, margin: 0 });
  });

  // Bottom banner
  card(sl, 0.4, 5.22, 9.2, 0.25, { fill: C.bg3, border: C.accent3 });
  sl.addText("Result: A summary where every sentence is both highly relevant AND maximally different from every other sentence.", {
    x: 0.55, y: 5.24, w: 8.9, h: 0.2,
    fontSize: 9, color: C.accent3, bold: true, align: "center", margin: 0,
  });
}

// ─────────────────────────────────────────────
// SLIDE 7: Enterprise UI & Analytics
// ─────────────────────────────────────────────
async function slide7() {
  let sl = pres.addSlide();
  sl.background = { color: C.bg };

  sl.addText("07", { x: 0.4, y: 0.25, w: 1, h: 0.35, fontSize: 11, color: C.accent, bold: true, margin: 0, charSpacing: 2 });
  sl.addText("Enterprise UI & Analytics Transparency", {
    x: 0.4, y: 0.55, w: 9.2, h: 0.55,
    fontSize: 24, color: C.white, bold: true, fontFace: "Arial Black", margin: 0,
  });
  sl.addText("Flask SaaS Application  ·  Split-Panel Dashboard  ·  Chart.js Score Visualisation", {
    x: 0.4, y: 1.1, w: 9.2, h: 0.3,
    fontSize: 12, color: C.muted, italic: true, margin: 0,
  });

  // Left — mockup area (Flask UI representation)
  card(sl, 0.4, 1.5, 5.5, 3.8, { fill: "0D1117", border: C.accent });
  // Browser chrome bar
  sl.addShape(pres.shapes.RECTANGLE, {
    x: 0.4, y: 1.5, w: 5.5, h: 0.38,
    fill: { color: "161B22" }, line: { color: C.border, width: 0 },
  });
  sl.addText("● ● ●", { x: 0.55, y: 1.55, w: 1.2, h: 0.28, fontSize: 11, color: C.muted, margin: 0 });
  sl.addText("http://localhost:5000  /  summarizer", { x: 1.5, y: 1.57, w: 3.0, h: 0.24, fontSize: 8, color: C.muted, align: "center", margin: 0 });
  sl.addText("Enterprise Extractive Summarizer", { x: 4.6, y: 1.57, w: 1.2, h: 0.24, fontSize: 7.5, color: C.accent, align: "right", margin: 0 });

  // Left panel (input)
  sl.addShape(pres.shapes.RECTANGLE, {
    x: 0.45, y: 1.95, w: 2.6, h: 3.2,
    fill: { color: "161B22" }, line: { color: C.border, width: 1 },
  });
  sl.addText("INPUT TEXT", { x: 0.55, y: 2.02, w: 2.4, h: 0.25, fontSize: 7.5, color: C.accent, bold: true, charSpacing: 2, margin: 0 });
  sl.addShape(pres.shapes.RECTANGLE, {
    x: 0.52, y: 2.32, w: 2.45, h: 2.1,
    fill: { color: C.bg }, line: { color: C.border, width: 1 },
  });
  sl.addText("Paste your document text here...", { x: 0.58, y: 2.4, w: 2.33, h: 0.3, fontSize: 7.5, color: C.border, italic: true, margin: 0 });
  sl.addShape(pres.shapes.RECTANGLE, {
    x: 0.52, y: 4.47, w: 2.45, h: 0.5,
    fill: { color: C.accent }, line: { color: C.accent, width: 0 },
  });
  sl.addText("SUMMARIZE →", { x: 0.55, y: 4.55, w: 2.39, h: 0.32, fontSize: 9, color: "0D1117", bold: true, align: "center", margin: 0 });

  // Right panel (output)
  sl.addShape(pres.shapes.RECTANGLE, {
    x: 3.12, y: 1.95, w: 2.65, h: 3.2,
    fill: { color: "161B22" }, line: { color: C.border, width: 1 },
  });
  sl.addText("SUMMARY OUTPUT", { x: 3.22, y: 2.02, w: 2.45, h: 0.25, fontSize: 7.5, color: C.accent3, bold: true, charSpacing: 2, margin: 0 });
  const bullets = ["Key finding about performance...", "Statistical model shows...", "Results indicate that...", "Conclusion: approach is valid..."];
  bullets.forEach((b, i) => {
    sl.addShape(pres.shapes.OVAL, {
      x: 3.18, y: 2.38 + i * 0.55, w: 0.14, h: 0.14,
      fill: { color: C.accent3 }, line: { color: C.accent3, width: 0 },
    });
    sl.addText(b, { x: 3.38, y: 2.35 + i * 0.55, w: 2.3, h: 0.45, fontSize: 8, color: C.muted, margin: 0 });
  });

  // Right side — feature highlights
  const features = [
    { icon: "📊", title: "Chart.js Pie Visualisation", desc: "Score distribution (TF-IDF, POS, Semantic, Structural) rendered as interactive pie chart per sentence." },
    { icon: "📋", title: "Sentence Score Data Table", desc: "Full audit table: every sentence with exact α, β, γ, δ sub-scores and final composite score — ranked." },
    { icon: "🔍", title: "Why-Did-It-Pick-This?", desc: "User can inspect exactly why any sentence was selected. Zero black-box trust required." },
    { icon: "⚡", title: "Flask REST Backend", desc: "Lightweight Python API serving the NLP pipeline. Stateless, scalable, ready for enterprise deployment." },
  ];
  features.forEach((f, i) => {
    const fy = 1.5 + i * 0.95;
    card(sl, 6.1, fy, 3.65, 0.82, { fill: C.bg2, accent: C.accent });
    sl.addText(f.icon, { x: 6.18, y: fy + 0.1, w: 0.45, h: 0.55, fontSize: 18, margin: 0 });
    sl.addText(f.title, { x: 6.65, y: fy + 0.08, w: 3.0, h: 0.28, fontSize: 10.5, color: C.white, bold: true, margin: 0 });
    sl.addText(f.desc, { x: 6.65, y: fy + 0.38, w: 3.0, h: 0.42, fontSize: 8.5, color: C.muted, margin: 0 });
  });

  // Bottom
  card(sl, 0.4, 5.22, 9.2, 0.25, { fill: C.bg3, border: C.accent });
  sl.addText("\"The user never has to trust the AI. The dashboard proves every decision mathematically.\"", {
    x: 0.55, y: 5.24, w: 8.9, h: 0.2,
    fontSize: 9, color: C.accent, bold: true, align: "center", italic: true, margin: 0,
  });
}

// ─────────────────────────────────────────────
// SLIDE 8: Conclusion & Future Work
// ─────────────────────────────────────────────
async function slide8() {
  let sl = pres.addSlide();
  sl.background = { color: C.bg };

  sl.addText("08", { x: 0.4, y: 0.25, w: 1, h: 0.35, fontSize: 11, color: C.accent, bold: true, margin: 0, charSpacing: 2 });
  sl.addText("Conclusion & Future Enhancements", {
    x: 0.4, y: 0.55, w: 9.2, h: 0.55,
    fontSize: 26, color: C.white, bold: true, fontFace: "Arial Black", margin: 0,
  });
  sl.addText("What we built · What we learned · Where we're going", {
    x: 0.4, y: 1.1, w: 9.2, h: 0.3,
    fontSize: 12, color: C.muted, italic: true, margin: 0,
  });

  // Achievements section
  card(sl, 0.4, 1.5, 5.6, 2.8, { fill: C.bg2, border: C.accent3 });
  sl.addText("✅  ACHIEVEMENTS", {
    x: 0.6, y: 1.6, w: 5.2, h: 0.3,
    fontSize: 11, color: C.accent3, bold: true, charSpacing: 2, margin: 0,
  });
  sl.addShape(pres.shapes.RECTANGLE, {
    x: 0.55, y: 1.96, w: 5.3, h: 0.02,
    fill: { color: C.border }, line: { color: C.border, width: 0 },
  });

  const achievements = [
    "Built a fully modular, mathematically transparent 4-stage NLP pipeline",
    "Balanced relevance (Heuristic Score) with diversity (MMR) in a single architecture",
    "Bridged the lexical gap using Wu-Palmer semantic similarity via WordNet",
    "Delivered an enterprise Flask SaaS product with full audit transparency",
    "Achieved 100% factual accuracy — zero hallucination by design",
  ];
  achievements.forEach((a, i) => {
    sl.addShape(pres.shapes.OVAL, {
      x: 0.58, y: 2.1 + i * 0.42, w: 0.14, h: 0.14,
      fill: { color: C.accent3 }, line: { color: C.accent3, width: 0 },
    });
    sl.addText(a, {
      x: 0.82, y: 2.07 + i * 0.42, w: 5.1, h: 0.38,
      fontSize: 9.5, color: C.white, margin: 0,
    });
  });

  // Limitation
  card(sl, 0.4, 4.45, 5.6, 0.95, { fill: C.bg2, border: C.accent5 });
  sl.addText("⚠  CURRENT LIMITATION — Greedy Algorithm", {
    x: 0.6, y: 4.52, w: 5.2, h: 0.28,
    fontSize: 10, color: C.accent5, bold: true, margin: 0,
  });
  sl.addText("The pipeline always fills the requested N slots, even if late-ranked candidates have very low quality scores. A user requesting 10 bullets always gets 10, regardless of quality threshold.", {
    x: 0.6, y: 4.83, w: 5.2, h: 0.5,
    fontSize: 8.5, color: C.muted, margin: 0,
  });

  // Roadmap / future
  card(sl, 6.15, 1.5, 3.65, 3.9, { fill: C.bg2, border: C.accent2 });
  sl.addText("🚀  FUTURE ROADMAP", {
    x: 6.3, y: 1.6, w: 3.3, h: 0.3,
    fontSize: 11, color: C.accent2, bold: true, charSpacing: 2, margin: 0,
  });
  sl.addShape(pres.shapes.RECTANGLE, {
    x: 6.2, y: 1.96, w: 3.5, h: 0.02,
    fill: { color: C.border }, line: { color: C.border, width: 0 },
  });

  const roadmap = [
    { v: "v2.0", t: "Dynamic Quality Gate", d: "Automatically stop extraction when candidate score falls below minimum quality threshold — smart termination." },
    { v: "v2.1", t: "Transformer Embeddings", d: "Replace TF-IDF with BERT sentence embeddings for semantic-level importance scoring." },
    { v: "v2.2", t: "Multi-Document Mode", d: "Cross-document summarization with source citation tracking for research workflows." },
    { v: "v3.0", t: "API + Enterprise Auth", d: "REST API with JWT auth, rate limiting, and org-level usage dashboards for B2B SaaS." },
  ];
  roadmap.forEach((r, i) => {
    const ry = 2.1 + i * 0.82;
    sl.addShape(pres.shapes.RECTANGLE, {
      x: 6.22, y: ry, w: 0.55, h: 0.3,
      fill: { color: C.accent2 }, line: { color: C.accent2, width: 0 },
    });
    sl.addText(r.v, { x: 6.22, y: ry + 0.03, w: 0.55, h: 0.26, fontSize: 7.5, color: "0D1117", bold: true, align: "center", margin: 0 });
    sl.addText(r.t, { x: 6.85, y: ry, w: 2.85, h: 0.28, fontSize: 10, color: C.white, bold: true, margin: 0 });
    sl.addText(r.d, { x: 6.22, y: ry + 0.33, w: 3.48, h: 0.42, fontSize: 8, color: C.muted, margin: 0 });
  });

  // Final bottom banner
  card(sl, 0.4, 5.22, 9.2, 0.25, { fill: C.bg3, border: C.accent });
  sl.addText("This project proves that transparency and intelligence are not in conflict — they are the same thing, engineered correctly.", {
    x: 0.55, y: 5.24, w: 8.9, h: 0.2,
    fontSize: 9, color: C.accent, bold: true, align: "center", italic: true, margin: 0,
  });
}

async function main() {
  console.log("Building slides...");
  await slide1();
  console.log("Slide 1 done");
  await slide2();
  console.log("Slide 2 done");
  await slide3();
  console.log("Slide 3 done");
  await slide4();
  console.log("Slide 4 done");
  await slide5();
  console.log("Slide 5 done");
  await slide6();
  console.log("Slide 6 done");
  await slide7();
  console.log("Slide 7 done");
  await slide8();
  console.log("Slide 8 done");

  await pres.writeFile({ fileName: "nlp_summarizer_deck.pptx" });
  console.log("DONE: nlp_summarizer_deck.pptx");
}

main().catch(console.error);