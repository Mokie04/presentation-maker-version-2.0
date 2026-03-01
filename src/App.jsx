import { useState } from "react";
import PptxGenJS from "pptxgenjs";

const P = {
  purple: "7C3AED",
  purpleL: "EDE9FE",
  blue: "2563EB",
  blueL: "DBEAFE",
  teal: "0891B2",
  tealL: "CFFAFE",
  green: "16A34A",
  greenL: "DCFCE7",
  orange: "EA580C",
  orangeL: "FFEDD5",
  pink: "DB2777",
  pinkL: "FCE7F3",
  yellow: "CA8A04",
  yellowL: "FEF9C3",
  red: "DC2626",
  redL: "FEE2E2",
  dark: "1E1B4B",
  white: "FFFFFF",
  offWhite: "F8F7FF",
  black: "111111",
};

const PALETTES = {
  rainbow: {
    label: "Rainbow",
    preview: ["#7C3AED", "#2563EB", "#0891B2", "#EA580C"],
    colors: {
      ...P,
      dark: "1E1B4B",
      offWhite: "F8F7FF",
    },
  },
  ocean: {
    label: "Ocean",
    preview: ["#0F4C81", "#0891B2", "#06B6D4", "#14B8A6"],
    colors: {
      purple: "0F4C81",
      purpleL: "DBEAFE",
      blue: "0369A1",
      blueL: "E0F2FE",
      teal: "0F766E",
      tealL: "CCFBF1",
      green: "0F766E",
      greenL: "D1FAE5",
      orange: "0284C7",
      orangeL: "E0F2FE",
      pink: "14B8A6",
      pinkL: "CCFBF1",
      yellow: "0EA5E9",
      yellowL: "E0F2FE",
      red: "155E75",
      redL: "CFFAFE",
      dark: "082F49",
      white: "FFFFFF",
      offWhite: "F3FBFF",
      black: "0F172A",
    },
  },
  forest: {
    label: "Forest",
    preview: ["#166534", "#65A30D", "#0F766E", "#CA8A04"],
    colors: {
      purple: "166534",
      purpleL: "DCFCE7",
      blue: "3F6212",
      blueL: "ECFCCB",
      teal: "0F766E",
      tealL: "CCFBF1",
      green: "15803D",
      greenL: "D1FAE5",
      orange: "CA8A04",
      orangeL: "FEF9C3",
      pink: "4D7C0F",
      pinkL: "ECFCCB",
      yellow: "A16207",
      yellowL: "FEF3C7",
      red: "B45309",
      redL: "FFEDD5",
      dark: "052E16",
      white: "FFFFFF",
      offWhite: "F7FDF7",
      black: "111827",
    },
  },
  sunset: {
    label: "Sunset",
    preview: ["#B91C1C", "#EA580C", "#F59E0B", "#EC4899"],
    colors: {
      purple: "B91C1C",
      purpleL: "FEE2E2",
      blue: "C2410C",
      blueL: "FFEDD5",
      teal: "EA580C",
      tealL: "FFEDD5",
      green: "D97706",
      greenL: "FEF3C7",
      orange: "F59E0B",
      orangeL: "FEF3C7",
      pink: "EC4899",
      pinkL: "FCE7F3",
      yellow: "FB7185",
      yellowL: "FFE4E6",
      red: "BE123C",
      redL: "FFE4E6",
      dark: "431407",
      white: "FFFFFF",
      offWhite: "FFF7ED",
      black: "1F2937",
    },
  },
};

const TEMPLATE_OPTIONS = {
  classroom: {
    label: "Classroom",
    description: "Bright lesson slides with balanced cards and clear labels.",
    badge: "Balanced",
    heroTone: "AUTO-BUILT",
    titleWords: ["SPARK", "LEARN", "GROW"],
    cardTransparency: 82,
    titleQuoteLabel: "Quote",
  },
  playful: {
    label: "Playful",
    description: "More energetic accents and a more lively title slide.",
    badge: "Energetic",
    heroTone: "FUN MODE",
    titleWords: ["PLAY", "MAKE", "SHINE"],
    cardTransparency: 76,
    titleQuoteLabel: "Class Motto",
  },
  formal: {
    label: "Formal",
    description: "Cleaner academic styling for reports and polished lessons.",
    badge: "Polished",
    heroTone: "PRESENTATION",
    titleWords: ["FOCUS", "STRUCTURE", "RESULT"],
    cardTransparency: 88,
    titleQuoteLabel: "Key Thought",
  },
};

function buildPptx(data) {
  const pres = new PptxGenJS();
  pres.layout = "LAYOUT_16x9";
  pres.title = data.topic;
  const template = TEMPLATE_OPTIONS[data.template] || TEMPLATE_OPTIONS.classroom;
  const theme = PALETTES[data.palette] || PALETTES.rainbow;
  const P = theme.colors;
  const ACCENTS = [P.purple, P.blue, P.teal, P.green, P.orange, P.pink, P.yellow, P.red];
  const LIGHTS = [P.purpleL, P.blueL, P.tealL, P.greenL, P.orangeL, P.pinkL, P.yellowL, P.redL];

  function hdr(slide, title, sub, bg) {
    const c = bg || P.purple;
    slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 1.05, fill: { color: c }, line: { color: c } });
    [[9.3, 0.2], [9.6, 0.5], [9.1, 0.7]].forEach(([x, y]) =>
      slide.addShape(pres.shapes.OVAL, { x, y, w: 0.3, h: 0.3, fill: { color: P.white, transparency: 30 }, line: { color: P.white, transparency: 30 } })
    );
    slide.addText(title, { x: 0.35, y: 0.08, w: 8.8, h: 0.65, fontSize: 26, bold: true, color: P.white, valign: "middle", margin: 0 });
    if (sub) {
      slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 1.05, w: 10, h: 0.34, fill: { color: P.offWhite }, line: { color: P.offWhite } });
      slide.addText(sub, { x: 0.35, y: 1.07, w: 9.3, h: 0.3, fontSize: 10.5, color: c, italic: true, valign: "middle", margin: 0 });
    }
    slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.33, w: 10, h: 0.295, fill: { color: c }, line: { color: c } });
    slide.addText(`Grade ${data.gradeLevel} | ${data.subject}`, { x: 0.3, y: 5.33, w: 5, h: 0.295, fontSize: 9, color: P.white, valign: "middle", margin: 0 });
    slide.addText(data.topic, { x: 5, y: 5.33, w: 4.7, h: 0.295, fontSize: 9, color: P.white, valign: "middle", align: "right", margin: 0 });
    slide.background = { color: P.offWhite };
  }

  function card(slide, x, y, w, h, col, light) {
    slide.addShape(pres.shapes.RECTANGLE, { x: x + 0.06, y: y + 0.06, w, h, fill: { color: col, transparency: template.cardTransparency }, line: { color: col, transparency: template.cardTransparency } });
    slide.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: light || P.white }, line: { color: col, pt: 2.5 } });
    slide.addShape(pres.shapes.RECTANGLE, { x, y, w, h: 0.09, fill: { color: col }, line: { color: col } });
  }

  function badge(slide, x, y, lbl, col) {
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 1.7, h: 0.32, fill: { color: col }, line: { color: col }, rectRadius: 0.08 });
    slide.addText(lbl, { x, y, w: 1.7, h: 0.32, fontSize: 9, bold: true, color: P.white, align: "center", valign: "middle", margin: 0, charSpacing: 2 });
  }

  {
    const sl = pres.addSlide();
    sl.background = { color: P.dark };
    [[0, 0, 3.5, P.purple, 55], [6.5, 3.5, 4, P.blue, 60], [4, 1.5, 2.5, P.teal, 65], [-0.5, 3.8, 3, P.pink, 65], [7.5, -0.5, 2.5, P.orange, 65]]
      .forEach(([x, y, s, c, t]) => sl.addShape(pres.shapes.OVAL, { x, y, w: s, h: s, fill: { color: c, transparency: t }, line: { color: c, transparency: t } }));
    sl.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 0.7, w: 9, h: 2.6, fill: { color: P.white, transparency: 10 }, line: { color: P.white, transparency: 30 } });
    sl.addText(data.topic, { x: 0.6, y: 0.8, w: 8.8, h: 1.3, fontSize: 44, bold: true, color: P.white, valign: "middle", margin: 0 });
    sl.addText(data.subject, { x: 0.6, y: 2.1, w: 8.8, h: 0.52, fontSize: 20, color: P.yellowL, valign: "middle", margin: 0 });
    [[data.subject, "purple"], [`Grade ${data.gradeLevel}`, "blue"], [data.quarter || "Q1", "orange"]].forEach(([t, k], i) => {
      const col = k === "purple" ? P.purple : k === "blue" ? P.blue : P.orange;
      sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6 + i * 2.1, y: 3.48, w: 1.85, h: 0.38, fill: { color: col }, line: { color: col }, rectRadius: 0.1 });
      sl.addText(t, { x: 0.6 + i * 2.1, y: 3.48, w: 1.85, h: 0.38, fontSize: 10, bold: true, color: P.white, align: "center", valign: "middle", margin: 0, charSpacing: 2 });
    });
    sl.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.1, w: 9, h: 0.7, fill: { color: P.white, transparency: 85 }, line: { color: P.white, transparency: 80 } });
    sl.addText(`"${data.tagline || "Words are bridges - they carry meaning from one mind to another."}"`, { x: 0.65, y: 4.14, w: 8.7, h: 0.62, fontSize: 12.5, color: P.white, italic: true, valign: "middle", margin: 0 });
    template.titleWords.forEach((s, i) => sl.addText(s, { x: 0.3 + i * 3.1, y: 5.0, w: 1.8, h: 0.3, fontSize: 9, color: P.white, bold: true, align: "center", valign: "middle", margin: 0 }));
  }

  {
    const sl = pres.addSlide();
    hdr(sl, "Learning Objectives", "By the end of this lesson, YOU will be able to...", P.purple);
    [
      { letter: "A", label: "KNOWLEDGE", icon: "BRAIN", text: data.objectives.knowledge, col: P.purple, lt: P.purpleL },
      { letter: "B", label: "SKILLS", icon: "WRITE", text: data.objectives.skills, col: P.blue, lt: P.blueL },
      { letter: "C", label: "ATTITUDE", icon: "IDEA", text: data.objectives.attitude, col: P.green, lt: P.greenL },
    ].forEach((o, i) => {
      const y = 1.5 + i * 1.17;
      card(sl, 0.35, y, 9.3, 1.05, o.col, o.lt);
      sl.addShape(pres.shapes.OVAL, { x: 0.5, y: y + 0.18, w: 0.68, h: 0.68, fill: { color: o.col }, line: { color: o.col } });
      sl.addText(o.letter, { x: 0.5, y: y + 0.18, w: 0.68, h: 0.68, fontSize: 20, bold: true, color: P.white, align: "center", valign: "middle", margin: 0 });
      sl.addText(o.icon, { x: 1.28, y: y + 0.18, w: 0.75, h: 0.2, fontSize: 8, bold: true, color: o.col, margin: 0 });
      sl.addText(o.label, { x: 1.9, y: y + 0.13, w: 2.5, h: 0.32, fontSize: 11, bold: true, color: o.col, margin: 0, charSpacing: 2 });
      sl.addText(o.text, { x: 1.9, y: y + 0.45, w: 7.55, h: 0.52, fontSize: 12, color: P.black, margin: 0 });
    });
    sl.addText(`Competency: ${data.competency}`, { x: 0.35, y: 5.0, w: 9.3, h: 0.25, fontSize: 9, color: P.purple, italic: true, margin: 0 });
  }

  {
    const sl = pres.addSlide();
    const act = data.activity;
    hdr(sl, `Activity: ${act.title}`, `4A's - Phase 1: ACTIVITY | ${act.duration} | Group Work`, P.orange);
    badge(sl, 0.35, 1.47, `TIME ${act.duration}`, P.orange);
    badge(sl, 2.18, 1.47, "GROUP WORK", P.purple);
    card(sl, 0.35, 1.9, 5.65, 3.2, P.orange, P.orangeL);
    sl.addText("What To Do:", { x: 0.52, y: 2.02, w: 5.3, h: 0.38, fontSize: 14, bold: true, color: P.orange, margin: 0 });
    (act.steps || []).forEach((step, i) => {
      const y = 2.48 + i * 0.5;
      sl.addText(`${i + 1}.`, { x: 0.5, y, w: 0.5, h: 0.44, fontSize: 16, align: "center", valign: "middle", margin: 0 });
      sl.addText(step, { x: 1.05, y: y + 0.04, w: 4.82, h: 0.36, fontSize: 11, color: P.black, margin: 0 });
    });
    card(sl, 6.18, 1.9, 3.48, 1.42, P.purple, P.purpleL);
    sl.addText("Guide Question", { x: 6.32, y: 2.0, w: 3.22, h: 0.38, fontSize: 12, bold: true, color: P.purple, margin: 0 });
    sl.addText(`"${act.guideQuestion}"`, { x: 6.32, y: 2.42, w: 3.22, h: 0.82, fontSize: 11.5, color: P.black, italic: true, margin: 0 });
    card(sl, 6.18, 3.45, 3.48, 1.65, P.teal, P.tealL);
    sl.addText("Materials:", { x: 6.32, y: 3.55, w: 3.22, h: 0.35, fontSize: 12, bold: true, color: P.teal, margin: 0 });
    sl.addText((act.materials || []).map((m, i) => ({ text: m, options: { bullet: true, breakLine: i < act.materials.length - 1, fontSize: 11, color: P.black } })), { x: 6.32, y: 3.96, w: 3.22, h: 1.1 });
  }

  {
    const sl = pres.addSlide();
    const ana = data.analysis;
    hdr(sl, ana.title, "4A's - Phase 2: ANALYSIS | Compare and observe!", P.teal);
    sl.addText(ana.prompt || "Read both versions carefully. What's different?", { x: 0.35, y: 1.45, w: 9.3, h: 0.28, fontSize: 12, color: P.teal, italic: true, margin: 0 });
    card(sl, 0.35, 1.8, 4.42, 3.26, P.red, P.redL);
    sl.addShape(pres.shapes.RECTANGLE, { x: 0.35, y: 1.8, w: 4.42, h: 0.52, fill: { color: P.red }, line: { color: P.red } });
    sl.addText(`Bad Example: ${ana.versionA.label}`, { x: 0.45, y: 1.8, w: 4.25, h: 0.52, fontSize: 11.5, bold: true, color: P.white, valign: "middle", margin: 0 });
    sl.addText(ana.versionA.text, { x: 0.5, y: 2.4, w: 4.15, h: 1.35, fontSize: 12, color: P.black, margin: 0 });
    sl.addShape(pres.shapes.RECTANGLE, { x: 0.35, y: 3.78, w: 4.42, h: 0.38, fill: { color: P.redL }, line: { color: P.red } });
    sl.addText(`Note: ${ana.versionA.note}`, { x: 0.48, y: 3.8, w: 4.15, h: 0.35, fontSize: 10.5, color: P.red, italic: true, valign: "middle", margin: 0 });
    sl.addShape(pres.shapes.OVAL, { x: 4.64, y: 3.0, w: 0.72, h: 0.72, fill: { color: P.yellow }, line: { color: P.yellow } });
    sl.addText("VS", { x: 4.64, y: 3.0, w: 0.72, h: 0.72, fontSize: 14, bold: true, color: P.white, align: "center", valign: "middle", margin: 0 });
    card(sl, 5.23, 1.8, 4.42, 3.26, P.green, P.greenL);
    sl.addShape(pres.shapes.RECTANGLE, { x: 5.23, y: 1.8, w: 4.42, h: 0.52, fill: { color: P.green }, line: { color: P.green } });
    sl.addText(`Strong Example: ${ana.versionB.label}`, { x: 5.33, y: 1.8, w: 4.25, h: 0.52, fontSize: 11.5, bold: true, color: P.white, valign: "middle", margin: 0 });
    sl.addText(ana.versionB.text, { x: 5.35, y: 2.4, w: 4.15, h: 1.35, fontSize: 12, color: P.black, margin: 0 });
    sl.addShape(pres.shapes.RECTANGLE, { x: 5.23, y: 3.78, w: 4.42, h: 0.38, fill: { color: P.greenL }, line: { color: P.green } });
    sl.addText(`Note: ${ana.versionB.note}`, { x: 5.36, y: 3.8, w: 4.2, h: 0.35, fontSize: 10.5, color: P.green, italic: true, valign: "middle", margin: 0 });
    sl.addText(`Discussion: ${ana.discussion}`, { x: 0.35, y: 5.07, w: 9.3, h: 0.22, fontSize: 9.5, color: P.teal, italic: true, margin: 0 });
  }

  {
    const sl = pres.addSlide();
    const def = data.definition;
    hdr(sl, def.title, "4A's - Phase 3: ABSTRACTION | The Key Concept", P.purple);
    sl.addShape(pres.shapes.RECTANGLE, { x: 0.35, y: 1.5, w: 9.3, h: 1.62, fill: { color: P.purple }, line: { color: P.purple } });
    sl.addText("Definition", { x: 0.55, y: 1.6, w: 9, h: 0.35, fontSize: 12, bold: true, color: P.purpleL, margin: 0 });
    sl.addText(def.text, { x: 0.55, y: 1.95, w: 9, h: 1.05, fontSize: 13, color: P.white, margin: 0 });
    (def.purposes || []).slice(0, 3).forEach((p, i) => {
      const cols = [[P.blue, P.blueL], [P.orange, P.orangeL], [P.green, P.greenL]];
      const x = 0.35 + i * 3.17;
      card(sl, x, 3.25, 3.02, 1.82, cols[i][0], cols[i][1]);
      sl.addText(p.icon, { x, y: 3.38, w: 3.02, h: 0.55, fontSize: 28, align: "center", valign: "middle", margin: 0 });
      sl.addText(p.title, { x: x + 0.15, y: 3.96, w: 2.72, h: 0.34, fontSize: 14, bold: true, color: cols[i][0], margin: 0 });
      sl.addText(p.desc, { x: x + 0.15, y: 4.32, w: 2.72, h: 0.68, fontSize: 10.5, color: P.black, margin: 0 });
    });
  }

  const concepts = data.concepts || [];
  [[0, 4], [4, 8]].forEach(([start, end], slideIdx) => {
    const sl = pres.addSlide();
    const colSet = slideIdx === 0 ? P.blue : P.pink;
    hdr(sl, `Key Concepts - Part ${slideIdx + 1}`, `4A's - Phase 3: ABSTRACTION | (${slideIdx + 1} of 2)`, colSet);
    concepts.slice(start, end).forEach((c, i) => {
      const col = i % 2;
      const row = Math.floor(i / 2);
      const x = 0.35 + col * 4.85;
      const y = 1.5 + row * 2.0;
      const ci = (start + i) % 8;
      card(sl, x, y, 4.62, 1.85, ACCENTS[ci], LIGHTS[ci]);
      sl.addShape(pres.shapes.OVAL, { x: x + 0.15, y: y + 0.2, w: 0.58, h: 0.58, fill: { color: ACCENTS[ci] }, line: { color: ACCENTS[ci] } });
      sl.addText(c.icon || "TAG", { x: x + 0.15, y: y + 0.2, w: 0.58, h: 0.58, fontSize: 18, align: "center", valign: "middle", margin: 0 });
      sl.addText((c.label || "").toUpperCase(), { x: x + 0.83, y: y + 0.22, w: 3.65, h: 0.35, fontSize: 11.5, bold: true, color: ACCENTS[ci], margin: 0, charSpacing: 1 });
      sl.addText((c.words || []).join(", "), { x: x + 0.15, y: y + 0.86, w: 4.35, h: 0.45, fontSize: 10, color: ACCENTS[ci], italic: true, margin: 0 });
      sl.addShape(pres.shapes.LINE, { x: x + 0.15, y: y + 1.35, w: 4.2, h: 0, line: { color: ACCENTS[ci], width: 1, dashType: "dash" } });
      sl.addText(`e.g. ${c.example || ""}`, { x: x + 0.15, y: y + 1.42, w: 4.35, h: 0.36, fontSize: 9.5, color: P.black, margin: 0 });
    });
  });

  {
    const sl = pres.addSlide();
    const app = data.application;
    hdr(sl, app.title, "4A's - Phase 4: APPLICATION | Individual Activity | 5 Minutes", P.orange);
    badge(sl, 0.35, 1.47, "INDIVIDUAL", P.orange);
    badge(sl, 2.18, 1.47, "5 MINUTES", P.red);
    const wb = app.wordBox || [];
    sl.addShape(pres.shapes.RECTANGLE, { x: 0.35, y: 1.88, w: 9.3, h: 0.6, fill: { color: P.dark }, line: { color: P.dark } });
    sl.addText("Word Box:", { x: 0.5, y: 1.92, w: 1.5, h: 0.5, fontSize: 11.5, bold: true, color: P.yellowL, valign: "middle", margin: 0 });
    const ww = Math.min(7.0 / Math.max(wb.length, 1), 1.45);
    wb.forEach((w, i) => {
      sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 2.08 + i * (ww + 0.08), y: 2.0, w: ww, h: 0.36, fill: { color: ACCENTS[i % 8] }, line: { color: ACCENTS[i % 8] }, rectRadius: 0.07 });
      sl.addText(w, { x: 2.08 + i * (ww + 0.08), y: 2.0, w: ww, h: 0.36, fontSize: 10, bold: true, color: P.white, align: "center", valign: "middle", margin: 0 });
    });
    (app.items || []).forEach((item, i) => {
      const y = 2.6 + i * 0.54;
      sl.addShape(pres.shapes.OVAL, { x: 0.35, y: y + 0.04, w: 0.42, h: 0.42, fill: { color: ACCENTS[i % 8] }, line: { color: ACCENTS[i % 8] } });
      sl.addText(`${i + 1}`, { x: 0.35, y: y + 0.04, w: 0.42, h: 0.42, fontSize: 13, bold: true, color: P.white, align: "center", valign: "middle", margin: 0 });
      sl.addText(item, { x: 0.88, y: y + 0.08, w: 8.65, h: 0.38, fontSize: 11.5, color: P.black, margin: 0 });
    });
    sl.addText("Write your answers on your activity sheet.", { x: 0.35, y: 5.08, w: 9.3, h: 0.22, fontSize: 9.5, color: P.orange, italic: true, margin: 0 });
  }

  {
    const sl = pres.addSlide();
    const qs = data.assessment || [];
    hdr(sl, "Formative Assessment", "5 Minutes | Answer independently!", P.red);
    sl.addText("Answer on your own. No looking at seatmates.", { x: 0.35, y: 1.47, w: 9.3, h: 0.3, fontSize: 12, color: P.red, italic: true, margin: 0 });
    qs.forEach((q, i) => {
      const y = 1.85 + i * 0.88;
      const col = ACCENTS[i % 8];
      const lt = LIGHTS[i % 8];
      card(sl, 0.35, y, 8.52, 0.78, col, lt);
      sl.addText(`${i + 1}.`, { x: 0.5, y: y + 0.1, w: 0.5, h: 0.58, fontSize: 18, align: "center", valign: "middle", margin: 0 });
      sl.addText(q.question, { x: 1.1, y: y + 0.1, w: 7.58, h: 0.6, fontSize: 11.5, color: P.black, margin: 0 });
      sl.addShape(pres.shapes.OVAL, { x: 8.75, y: y + 0.1, w: 0.8, h: 0.58, fill: { color: col }, line: { color: col } });
      sl.addText(`${q.points} pt${q.points > 1 ? "s" : ""}`, { x: 8.75, y: y + 0.1, w: 0.8, h: 0.58, fontSize: 9.5, bold: true, color: P.white, align: "center", valign: "middle", margin: 0 });
    });
    const total = qs.reduce((s, q) => s + (q.points || 0), 0);
    sl.addShape(pres.shapes.RECTANGLE, { x: 0.35, y: 5.04, w: 9.3, h: 0.28, fill: { color: P.dark }, line: { color: P.dark } });
    sl.addText(`Total: ${total} points | Pass mark: ${Math.round(total * 0.8)} points (80%)`, { x: 0.35, y: 5.04, w: 9.3, h: 0.28, fontSize: 11, bold: true, color: P.white, align: "center", valign: "middle", margin: 0 });
  }

  {
    const sl = pres.addSlide();
    const asgn = data.assignment;
    hdr(sl, "Assignment / Agreement", "To be submitted next meeting", P.purple);
    sl.addShape(pres.shapes.RECTANGLE, { x: 0.35, y: 1.5, w: 9.3, h: 2.3, fill: { color: P.purple }, line: { color: P.purple } });
    sl.addText("Your Task:", { x: 0.55, y: 1.6, w: 9, h: 0.35, fontSize: 13, bold: true, color: P.purpleL, margin: 0 });
    sl.addText(asgn.task, { x: 0.55, y: 1.98, w: 9, h: 1.72, fontSize: 12.5, color: P.white, margin: 0 });
    card(sl, 0.35, 3.92, 4.4, 1.28, P.green, P.greenL);
    sl.addText("Checklist", { x: 0.52, y: 4.02, w: 4.1, h: 0.35, fontSize: 12, bold: true, color: P.green, margin: 0 });
    sl.addText((asgn.checklist || []).map((c, i) => ({ text: c, options: { bullet: true, breakLine: i < asgn.checklist.length - 1, fontSize: 11, color: P.black } })), { x: 0.52, y: 4.4, w: 4.12, h: 0.72 });
    card(sl, 5, 3.92, 4.65, 1.28, P.orange, P.orangeL);
    sl.addText("Suggested Topics:", { x: 5.15, y: 4.02, w: 4.35, h: 0.35, fontSize: 12, bold: true, color: P.orange, margin: 0 });
    sl.addText((asgn.topics || []).map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < asgn.topics.length - 1, fontSize: 11, color: P.black } })), { x: 5.15, y: 4.4, w: 4.4, h: 0.72 });
  }

  {
    const sl = pres.addSlide();
    sl.background = { color: P.dark };
    ACCENTS.forEach((c, i) => sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: i * 0.7, w: 0.22, h: 0.72, fill: { color: c }, line: { color: c } }));
    [[9.5, 0.2, P.yellow], [8.8, 0.8, P.pink], [9.2, 1.5, P.teal], [8.5, 2.1, P.orange], [9.4, 3.0, P.purple], [8.7, 3.6, P.green], [9.1, 4.3, P.blue], [8.5, 5.0, P.red]]
      .forEach(([x, y, c]) => sl.addShape(pres.shapes.OVAL, { x, y, w: 0.28, h: 0.28, fill: { color: c, transparency: 40 }, line: { color: c, transparency: 40 } }));
    sl.addShape(pres.shapes.RECTANGLE, { x: 0.38, y: 0.35, w: 8.88, h: 2.35, fill: { color: P.purple, transparency: 82 }, line: { color: P.purple, transparency: 60 } });
    sl.addText(template.titleQuoteLabel, { x: 0.42, y: 0.38, w: 1.3, h: 1.2, fontSize: 22, color: P.purpleL, bold: true, margin: 0 });
    sl.addText(data.closingQuote || "Great writing is not accidental - it is built, bridge by bridge, with the right words.", { x: 1.05, y: 0.5, w: 8, h: 1.55, fontSize: 22, color: P.white, italic: true, margin: 0 });
    sl.addShape(pres.shapes.LINE, { x: 0.38, y: 2.9, w: 8.88, h: 0, line: { color: "888888", width: 1 } });
    sl.addText("Key Takeaways:", { x: 0.38, y: 3.05, w: 8.88, h: 0.38, fontSize: 15, bold: true, color: P.yellowL, margin: 0 });
    (data.takeaways || []).slice(0, 3).forEach((tk, i) => {
      sl.addShape(pres.shapes.OVAL, { x: 0.38 + i * 3, y: 3.52, w: 0.42, h: 0.42, fill: { color: ACCENTS[i] }, line: { color: ACCENTS[i] } });
      sl.addText(`${i + 1}`, { x: 0.38 + i * 3, y: 3.52, w: 0.42, h: 0.42, fontSize: 14, color: P.white, align: "center", valign: "middle", margin: 0 });
      sl.addText(tk, { x: 0.88 + i * 3, y: 3.55, w: 2.6, h: 0.38, fontSize: 10.5, color: P.white, margin: 0 });
    });
    sl.addText("Great work today, class!", { x: 0.38, y: 4.22, w: 8.88, h: 0.42, fontSize: 14, bold: true, color: P.yellowL, align: "center", margin: 0 });
    sl.addText(`Grade ${data.gradeLevel} ${data.subject} | MATATAG Curriculum`, { x: 0.38, y: 4.8, w: 8.88, h: 0.3, fontSize: 10, color: "999999", align: "center", margin: 0 });
  }

  return pres.writeFile({ fileName: `${data.topic.replace(/[^a-z0-9]/gi, "_")}_Presentation.pptx` });
}

function cleanTopicTokens(topic) {
  return topic
    .split(/[^A-Za-z0-9]+/)
    .map((token) => token.trim())
    .filter((token) => token.length > 2)
    .filter((token) => !["the", "and", "for", "with", "from", "into", "using", "lesson", "grade"].includes(token.toLowerCase()));
}

function buildTopicKeywords(topic) {
  const unique = [];
  for (const token of cleanTopicTokens(topic)) {
    const normalized = token.toLowerCase();
    if (!unique.some((item) => item.toLowerCase() === normalized)) {
      unique.push(token);
    }
  }
  const fallback = ["idea", "example", "reason", "process", "application"];
  return [...unique.slice(0, 5), ...fallback].slice(0, 5);
}

function createLocalLessonData(form) {
  const topic = form.topic.trim();
  const subject = form.subject || "English";
  const gradeLevel = form.gradeLevel || "8";
  const quarter = form.quarter || "First Quarter";
  const template = form.template || "classroom";
  const palette = form.palette || "rainbow";
  const templateConfig = TEMPLATE_OPTIONS[template] || TEMPLATE_OPTIONS.classroom;
  const paletteConfig = PALETTES[palette] || PALETTES.rainbow;
  const keywords = buildTopicKeywords(topic);
  const lead = keywords[0] || "Concept";
  const second = keywords[1] || "Process";
  const third = keywords[2] || "Application";

  return {
    topic,
    subject,
    gradeLevel,
    quarter,
    template,
    palette,
    themeLabel: paletteConfig.label,
    competency: form.competency || `Students build understanding and practical use of ${topic} through explanation, examples, and application tasks.`,
    tagline: `${topic} becomes stronger when students understand it, practice it, and apply it with purpose in the ${templateConfig.label.toLowerCase()} template.`,
    objectives: {
      knowledge: form.objKnowledge || `Explain the meaning, purpose, and important ideas related to ${topic}.`,
      skills: form.objSkills || `Apply what was learned about ${topic} in guided and independent activities.`,
      attitude: form.objAttitude || `Show confidence, curiosity, and responsibility while learning about ${topic}.`,
    },
    activity: {
      title: `Discovering ${lead}`,
      duration: "10 Minutes",
      steps: [
        `Observe the sample about ${topic}.`,
        "List the important words or ideas you notice.",
        `Discuss why ${topic} matters in the lesson.`,
        "Work with your group to organize your ideas.",
        "Share one key insight with the class.",
      ],
      guideQuestion: `What makes ${topic} clear, useful, and meaningful for learners?`,
      materials: ["activity sheet", "pen or pencil", "sample prompt or text"],
    },
    analysis: {
      title: `Looking Closely at ${lead}`,
      prompt: `Read both versions carefully. Which one explains ${topic} more clearly?`,
      versionA: {
        label: "Without clear support",
        text: `${topic} is introduced, but the explanation is short and lacks examples. Some ideas feel disconnected. The audience may understand only part of the lesson.`,
        note: "The explanation is weak because the ideas are too limited and loosely connected.",
      },
      versionB: {
        label: "With clear support",
        text: `${topic} is easier to understand when the explanation includes the main idea, a clear example, and a practical classroom application.`,
        note: "This version is stronger because it gives a clearer flow of ideas and a useful example.",
      },
      discussion: `How did the clearer explanation improve understanding of ${topic}? What details made it stronger?`,
    },
    definition: {
      title: `Understanding ${topic}`,
      text: `${topic} is a lesson focus that helps learners build knowledge, practice a skill, and connect ideas to real situations. It becomes meaningful when students can explain it clearly, identify examples, and use it in tasks.`,
      purposes: [
        { icon: "BOOK", title: "Build Knowledge", desc: `Students identify the main ideas behind ${topic}.` },
        { icon: "TOOLS", title: "Use the Skill", desc: `Learners practice ${topic} through guided activities.` },
        { icon: "STAR", title: "See the Value", desc: `${topic} becomes useful when linked to real-life situations.` },
      ],
    },
    concepts: [
      { label: "KEY IDEA", icon: "IDEA", words: [lead, "meaning", "focus"], example: `${topic} gives the class a clear focus for learning.` },
      { label: "PURPOSE", icon: "TARGET", words: [second, "goal", "outcome"], example: `${topic} helps learners work toward a clear goal.` },
      { label: "PROCESS", icon: "PATH", words: [third, "steps", "sequence"], example: `Students can follow ${topic} step by step.` },
      { label: "EXAMPLES", icon: "TEST", words: [lead, "model", "sample"], example: `A strong example makes ${topic} easier to understand.` },
      { label: "APPLICATION", icon: "BUILD", words: [second, "task", "practice"], example: `${topic} becomes meaningful when used in a task.` },
      { label: "REAL LIFE", icon: "WORLD", words: [third, "daily life", "community"], example: `${topic} can be connected to home, school, and community situations.` },
      { label: "COMMON ERRORS", icon: "WARN", words: ["confusion", "missing details", "weak explanation"], example: `Students improve ${topic} by noticing common mistakes.` },
      { label: "REFLECTION", icon: "VIEW", words: ["insight", "growth", "improvement"], example: `Reflection strengthens understanding of ${topic}.` },
    ],
    application: {
      title: "Let's Practice!",
      wordBox: keywords,
      items: [
        `${topic} helps students understand __________ more clearly.`,
        `A strong example of ${topic} should include a clear __________.`,
        `Learners use ${topic} during guided __________.`,
        `${topic} becomes useful when connected to real-life __________.`,
        `Reflection helps students improve their understanding of __________.`,
      ],
    },
    assessment: [
      { points: 2, question: `What is the main idea of ${topic}?` },
      { points: 2, question: `Give one classroom example related to ${topic}.` },
      { points: 3, question: `Explain why ${topic} is important for ${subject} learners.` },
      { points: 3, question: `Write one short response showing what you learned about ${topic}.` },
    ],
    assignment: {
      task: `Create a short output that explains ${topic} in your own words. Include one example, one practical use, and one reflection about why the lesson matters.`,
      checklist: ["Use your own words", "Give one clear example", "Write neatly and completely"],
      topics: [
        `${topic} at school`,
        `${topic} at home`,
        `${topic} in the community`,
        `${topic} in everyday life`,
      ],
    },
    closingQuote: `${topic} grows stronger when students understand it, practice it, and use it with confidence.`,
    takeaways: [
      `Know the key idea of ${topic}`,
      `Use ${topic} in meaningful tasks`,
      `Reflect on how ${topic} helps learning`,
    ],
  };
}

export default function App() {
  const [form, setForm] = useState({
    topic: "",
    subject: "English",
    gradeLevel: "8",
    quarter: "First Quarter",
    template: "classroom",
    palette: "rainbow",
    competency: "",
    objKnowledge: "",
    objSkills: "",
    objAttitude: "",
    extraContext: "",
  });

  const [status, setStatus] = useState("idle");
  const [progress, setProgress] = useState("");
  const [generatedData, setGeneratedData] = useState(null);
  const [errorMsg, setErrorMsg] = useState("");

  const subjectOptions = ["English", "Filipino", "Science", "Mathematics", "Araling Panlipunan", "EAPP", "Literature", "History"];
  const gradeOptions = ["7", "8", "9", "10", "11", "12"];
  const quarterOptions = ["First Quarter", "Second Quarter", "Third Quarter", "Fourth Quarter"];

  const set = (k, v) => setForm((f) => ({ ...f, [k]: v }));

  async function handleGenerate() {
    if (!form.topic.trim()) {
      return;
    }

    setStatus("loading");
    setErrorMsg("");
    setProgress("Preparing your lesson details...");

    try {
      setProgress("Generating your lesson content...");
      const data = createLocalLessonData(form);
      setProgress("Building your PowerPoint presentation...");
      setGeneratedData(data);
      await buildPptx(data);
      setStatus("done");
      setProgress("");
    } catch (e) {
      console.error(e);
      setErrorMsg(`Something went wrong: ${e.message || "Unknown error. Please try again."}`);
      setStatus("error");
    }
  }

  async function handleRedownload() {
    if (!generatedData) {
      return;
    }
    setStatus("loading");
    setProgress("Re-generating your PowerPoint...");
    try {
      await buildPptx(generatedData);
      setStatus("done");
      setProgress("");
    } catch (e) {
      setStatus("error");
      setErrorMsg(e.message);
    }
  }

  const styles = `
    @import url('https://fonts.googleapis.com/css2?family=Nunito:wght@400;600;700;800;900&family=Quicksand:wght@500;600;700&display=swap');
    *{box-sizing:border-box;margin:0;padding:0}
    body{font-family:'Nunito',sans-serif;background:#f0f4ff;min-height:100vh}
    .app{min-height:100vh;background:linear-gradient(135deg,#667eea22 0%,#764ba222 50%,#f093fb22 100%);padding:24px 16px;display:flex;flex-direction:column;align-items:center}
    .hero{text-align:center;margin-bottom:32px;padding-top:8px}
    .hero-badge{display:inline-flex;align-items:center;gap:6px;background:#7C3AED;color:#fff;border-radius:999px;padding:5px 16px;font-size:12px;font-weight:800;letter-spacing:2px;margin-bottom:14px;text-transform:uppercase}
    .hero-title{font-size:clamp(28px,5vw,46px);font-weight:900;color:#1E1B4B;line-height:1.1;margin-bottom:10px;font-family:'Quicksand',sans-serif}
    .hero-title span{background:linear-gradient(135deg,#7C3AED,#DB2777,#EA580C);-webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text}
    .hero-sub{font-size:15px;color:#6B7280;max-width:520px;margin:0 auto;font-weight:600}
    .card{background:#fff;border-radius:24px;box-shadow:0 8px 40px #7C3AED18;padding:32px;width:100%;max-width:780px;margin-bottom:20px}
    .section-title{font-size:13px;font-weight:800;color:#7C3AED;letter-spacing:2px;text-transform:uppercase;margin-bottom:16px;display:flex;align-items:center;gap:8px}
    .section-title::after{content:'';flex:1;height:2px;background:linear-gradient(90deg,#7C3AED33,transparent)}
    .grid-2{display:grid;grid-template-columns:1fr 1fr;gap:16px}
    .field{display:flex;flex-direction:column;gap:6px;margin-bottom:16px}
    .field label{font-size:13px;font-weight:700;color:#374151}
    .field input,.field select,.field textarea{
      border:2px solid #E5E7EB;border-radius:12px;padding:12px 14px;font-size:14px;font-family:'Nunito',sans-serif;
      color:#111;background:#FAFAFA;transition:all .2s;outline:none;resize:vertical;font-weight:600;
    }
    .field input:focus,.field select:focus,.field textarea:focus{border-color:#7C3AED;background:#fff;box-shadow:0 0 0 4px #7C3AED15}
    .field textarea{min-height:72px}
    .btn-generate{
      width:100%;padding:18px;border-radius:16px;border:none;cursor:pointer;
      background:linear-gradient(135deg,#7C3AED,#DB2777);
      color:#fff;font-size:17px;font-weight:800;font-family:'Nunito',sans-serif;
      letter-spacing:.5px;transition:all .2s;position:relative;overflow:hidden;
    }
    .btn-generate:hover:not(:disabled){transform:translateY(-2px);box-shadow:0 8px 30px #7C3AED55}
    .btn-generate:disabled{opacity:.7;cursor:not-allowed}
    .loading-box{text-align:center;padding:32px;background:#7C3AED08;border-radius:20px;border:2px dashed #7C3AED44}
    .spinner{width:52px;height:52px;border:5px solid #E5E7EB;border-top-color:#7C3AED;border-radius:50%;animation:spin 1s linear infinite;margin:0 auto 16px}
    @keyframes spin{to{transform:rotate(360deg)}}
    .progress-txt{font-size:15px;font-weight:700;color:#7C3AED;margin-bottom:4px}
    .progress-sub{font-size:13px;color:#9CA3AF}
    .steps-row{display:flex;gap:8px;justify-content:center;margin-top:16px;flex-wrap:wrap}
    .step-dot{width:10px;height:10px;border-radius:50%;background:#E5E7EB;transition:all .4s}
    .step-dot.active{background:#7C3AED;transform:scale(1.3)}
    .success-box{text-align:center;padding:28px;background:linear-gradient(135deg,#16A34A11,#0891B211);border-radius:20px;border:2px solid #16A34A33}
    .success-icon{font-size:52px;margin-bottom:12px}
    .success-title{font-size:20px;font-weight:900;color:#16A34A;margin-bottom:6px}
    .success-sub{font-size:14px;color:#6B7280;margin-bottom:20px;font-weight:600}
    .preview-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(120px,1fr));gap:8px;margin:16px 0}
    .preview-pill{background:#fff;border:2px solid #E5E7EB;border-radius:10px;padding:8px 10px;font-size:11px;font-weight:700;color:#374151;text-align:center}
    .preview-pill span{display:block;font-size:16px;margin-bottom:2px}
    .btn-dl{
      display:inline-flex;align-items:center;gap:8px;
      background:linear-gradient(135deg,#16A34A,#0891B2);color:#fff;
      border:none;border-radius:14px;padding:14px 28px;font-size:15px;font-weight:800;
      cursor:pointer;font-family:'Nunito',sans-serif;transition:all .2s;
    }
    .btn-dl:hover{transform:translateY(-2px);box-shadow:0 8px 24px #16A34A44}
    .btn-new{
      display:inline-flex;align-items:center;gap:8px;margin-top:10px;
      background:#F3F4F6;color:#374151;border:2px solid #E5E7EB;
      border-radius:14px;padding:12px 22px;font-size:14px;font-weight:700;
      cursor:pointer;font-family:'Nunito',sans-serif;transition:all .2s;
    }
    .btn-new:hover{background:#E5E7EB}
    .error-box{padding:20px;background:#FEE2E2;border-radius:16px;border:2px solid #DC262633}
    .error-title{font-size:15px;font-weight:800;color:#DC2626;margin-bottom:4px}
    .error-msg{font-size:13px;color:#7F1D1D;font-weight:600}
    .tag{display:inline-block;background:#7C3AED18;color:#7C3AED;border-radius:999px;padding:2px 10px;font-size:11px;font-weight:800}
    .choice-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:12px}
    .choice-btn{border:2px solid #E5E7EB;border-radius:18px;padding:14px;background:#FFFFFF;cursor:pointer;text-align:left;transition:all .2s}
    .choice-btn.active{border-color:#7C3AED;box-shadow:0 0 0 4px #7C3AED15;transform:translateY(-1px)}
    .choice-title{font-size:14px;font-weight:800;color:#111827;margin-bottom:4px}
    .choice-text{font-size:12px;color:#6B7280;font-weight:600;line-height:1.4}
    .swatch-row{display:flex;gap:6px;margin-top:10px}
    .swatch{width:18px;height:18px;border-radius:999px;border:1px solid #ffffffaa;box-shadow:0 0 0 1px #E5E7EB}
    @media(max-width:600px){.grid-2{grid-template-columns:1fr}}
    @media(max-width:600px){.choice-grid{grid-template-columns:1fr}}
  `;

  const slideLabels = [
    { icon: "1", label: "Title" },
    { icon: "2", label: "Objectives" },
    { icon: "3", label: "Activity" },
    { icon: "4", label: "Analysis" },
    { icon: "5", label: "Definition" },
    { icon: "6", label: "Concepts 1" },
    { icon: "7", label: "Concepts 2" },
    { icon: "8", label: "Practice" },
    { icon: "9", label: "Assessment" },
    { icon: "10", label: "Assignment" },
    { icon: "11", label: "Closing" },
  ];

  return (
    <>
      <style>{styles}</style>
      <div className="app">
        <div className="hero">
          <div className="hero-badge">{TEMPLATE_OPTIONS[form.template].heroTone}</div>
          <h1 className="hero-title">Lesson <span>Presentation</span> Generator</h1>
          <p className="hero-sub">Built for the MATATAG Curriculum. Fill in your lesson details and get a colorful, ready-to-use PowerPoint in seconds.</p>
        </div>

        {status === "idle" || status === "error" ? (
          <div className="card">
            <div className="section-title">Lesson Information</div>
            <div className="grid-2">
              <div className="field">
                <label>Subject</label>
                <select value={form.subject} onChange={(e) => set("subject", e.target.value)}>
                  {subjectOptions.map((subject) => <option key={subject}>{subject}</option>)}
                </select>
              </div>
              <div className="field">
                <label>Grade Level</label>
                <select value={form.gradeLevel} onChange={(e) => set("gradeLevel", e.target.value)}>
                  {gradeOptions.map((grade) => <option key={grade} value={grade}>Grade {grade}</option>)}
                </select>
              </div>
            </div>
            <div className="field">
              <label>Quarter</label>
              <select value={form.quarter} onChange={(e) => set("quarter", e.target.value)}>
                {quarterOptions.map((quarter) => <option key={quarter}>{quarter}</option>)}
              </select>
            </div>
            <div className="field">
              <label>Topic / Lesson Title <span className="tag">Required</span></label>
              <input
                placeholder="e.g. Expository Essay on the Use of Transitional Devices"
                value={form.topic}
                onChange={(e) => set("topic", e.target.value)}
              />
            </div>
            <div className="field">
              <label>Competency</label>
              <textarea
                placeholder="e.g. Examine linguistic features as tools to achieve organizational efficiency in informational texts."
                value={form.competency}
                onChange={(e) => set("competency", e.target.value)}
              />
            </div>

            <div className="section-title" style={{ marginTop: 8 }}>Design Setup</div>
            <div className="field">
              <label>Template</label>
              <div className="choice-grid">
                {Object.entries(TEMPLATE_OPTIONS).map(([key, option]) => (
                  <button
                    key={key}
                    type="button"
                    className={`choice-btn ${form.template === key ? "active" : ""}`}
                    onClick={() => set("template", key)}
                  >
                    <div className="choice-title">{option.label}</div>
                    <div className="choice-text">{option.description}</div>
                    <div className="tag" style={{ marginTop: 10 }}>{option.badge}</div>
                  </button>
                ))}
              </div>
            </div>
            <div className="field">
              <label>Color Palette</label>
              <div className="choice-grid">
                {Object.entries(PALETTES).map(([key, option]) => (
                  <button
                    key={key}
                    type="button"
                    className={`choice-btn ${form.palette === key ? "active" : ""}`}
                    onClick={() => set("palette", key)}
                  >
                    <div className="choice-title">{option.label}</div>
                    <div className="choice-text">Applies this palette to title, content cards, and closing slides.</div>
                    <div className="swatch-row">
                      {option.preview.map((color) => <span key={color} className="swatch" style={{ background: color }} />)}
                    </div>
                  </button>
                ))}
              </div>
            </div>

            <div className="section-title" style={{ marginTop: 8 }}>Learning Objectives</div>
            <div className="field">
              <label>Knowledge Objective <span className="tag">Optional - generator will draft if blank</span></label>
              <input
                placeholder="e.g. Define transitional devices and explain their role in organizational efficiency"
                value={form.objKnowledge}
                onChange={(e) => set("objKnowledge", e.target.value)}
              />
            </div>
            <div className="field">
              <label>Skills Objective</label>
              <input
                placeholder="e.g. Identify, classify, and use transitional devices to improve coherence of written texts"
                value={form.objSkills}
                onChange={(e) => set("objSkills", e.target.value)}
              />
            </div>
            <div className="field">
              <label>Attitude Objective</label>
              <input
                placeholder="e.g. Appreciate the value of transitional devices in producing coherent informational texts"
                value={form.objAttitude}
                onChange={(e) => set("objAttitude", e.target.value)}
              />
            </div>

            <div className="section-title" style={{ marginTop: 8 }}>Additional Context</div>
            <div className="field">
              <label>Any extra notes for the generator? (optional)</label>
              <textarea
                placeholder="e.g. Focus on examples from environmental issues. Make it fun and interactive. Include Filipino sample sentences."
                value={form.extraContext}
                onChange={(e) => set("extraContext", e.target.value)}
              />
            </div>

            {status === "error" && (
              <div className="error-box" style={{ marginBottom: 16 }}>
                <div className="error-title">Oops! Something went wrong</div>
                <div className="error-msg">{errorMsg}</div>
              </div>
            )}

            <button
              className="btn-generate"
              onClick={handleGenerate}
              disabled={!form.topic.trim()}
            >
              Generate My Presentation
            </button>
          </div>
        ) : status === "loading" ? (
          <div className="card">
            <div className="loading-box">
              <div className="spinner" />
              <div className="progress-txt">{progress}</div>
              <div className="progress-sub">This takes about 20-40 seconds. Hang tight.</div>
              <div className="steps-row">
                {["Preparing", "Generating", "Building"].map((step) => (
                  <div key={step} style={{ textAlign: "center" }}>
                    <div className={`step-dot ${progress.includes(step) ? "active" : ""}`} style={{ margin: "0 auto 4px" }} />
                    <div style={{ fontSize: 10, color: "#9CA3AF", fontWeight: 700 }}>{step}</div>
                  </div>
                ))}
              </div>
            </div>
          </div>
        ) : (
          <div className="card">
            <div className="success-box">
              <div className="success-icon">DONE</div>
              <div className="success-title">Your Presentation is Ready</div>
              <div className="success-sub">
                <strong>{generatedData?.topic}</strong> - Grade {generatedData?.gradeLevel} {generatedData?.subject} - {generatedData?.themeLabel} palette - {slideLabels.length} Slides
              </div>
              <div className="preview-grid">
                {slideLabels.map((slide, index) => (
                  <div className="preview-pill" key={index}>
                    <span>{slide.icon}</span>{slide.label}
                  </div>
                ))}
              </div>
              <div style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 10, marginTop: 8 }}>
                <button className="btn-dl" onClick={handleRedownload}>
                  Download PowerPoint (.pptx)
                </button>
                <button className="btn-new" onClick={() => { setStatus("idle"); setGeneratedData(null); }}>
                  Create Another Presentation
                </button>
              </div>
            </div>
          </div>
        )}

        <div style={{ fontSize: 12, color: "#9CA3AF", fontWeight: 600, textAlign: "center", maxWidth: 480 }}>
          Powered by built-in lesson generation - MATATAG Curriculum - Generates 11 colorful slides following the 4A&apos;s framework
        </div>
      </div>
    </>
  );
}
