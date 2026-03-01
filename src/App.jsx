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
    layoutVariant: "balanced",
  },
  playful: {
    label: "Playful",
    description: "More energetic accents and a more lively title slide.",
    badge: "Energetic",
    heroTone: "FUN MODE",
    titleWords: ["PLAY", "MAKE", "SHINE"],
    cardTransparency: 76,
    titleQuoteLabel: "Class Motto",
    layoutVariant: "expressive",
  },
  formal: {
    label: "Formal",
    description: "Cleaner academic styling for reports and polished lessons.",
    badge: "Polished",
    heroTone: "PRESENTATION",
    titleWords: ["FOCUS", "STRUCTURE", "RESULT"],
    cardTransparency: 88,
    titleQuoteLabel: "Key Thought",
    layoutVariant: "split",
  },
};

const FRAMEWORK_OPTIONS = {
  fourAs: {
    label: "4A's",
    description: "Activity, Analysis, Abstraction, Application",
  },
  fiveEs: {
    label: "5E's",
    description: "Engage, Explore, Explain, Elaborate, Evaluate",
  },
  sevenEs: {
    label: "7E's",
    description: "Elicit, Engage, Explore, Explain, Elaborate, Evaluate, Extend",
  },
};

const DETAIL_OPTIONS = {
  minimal: {
    label: "Minimal",
    description: "Cleaner slides with shorter text blocks.",
  },
  balanced: {
    label: "Balanced",
    description: "Default level with enough detail for teaching.",
  },
  detailed: {
    label: "Detailed",
    description: "More guidance, more explanation, more teaching prompts.",
  },
};

function clampWords(text, maxWords) {
  const words = `${text}`.trim().split(/\s+/).filter(Boolean);
  if (words.length <= maxWords) {
    return text;
  }
  return `${words.slice(0, maxWords).join(" ")}...`;
}

function pickByDetail(detailLevel, minimal, balanced, detailed = balanced) {
  if (detailLevel === "minimal") {
    return minimal;
  }
  if (detailLevel === "detailed") {
    return detailed;
  }
  return balanced;
}

function getFrameworkContent(frameworkKey) {
  switch (frameworkKey) {
    case "fiveEs":
      return {
        objectivesSub: "5E's - ENGAGE | Set purpose and learning targets",
        activitySub: (duration) => `5E's - EXPLORE | ${duration} | Hands-on discovery`,
        analysisSub: "5E's - EXPLORE | Observe, compare, and investigate",
        definitionSub: "5E's - EXPLAIN | Clarify the key concept",
        conceptsSub: (part) => `5E's - EXPLAIN | Deepen understanding (${part} of 2)`,
        applicationSub: "5E's - ELABORATE | Independent practice | 5 Minutes",
        assessmentTitle: "Evaluate Understanding",
        assessmentSub: "5E's - EVALUATE | 5 Minutes | Answer independently!",
      };
    case "sevenEs":
      return {
        objectivesSub: "7E's - ELICIT | Surface prior ideas and targets",
        activitySub: (duration) => `7E's - ENGAGE | ${duration} | Spark curiosity`,
        analysisSub: "7E's - EXPLORE | Compare, investigate, and discuss",
        definitionSub: "7E's - EXPLAIN | Name and clarify the concept",
        conceptsSub: (part) => `7E's - ELABORATE | Build understanding (${part} of 2)`,
        applicationSub: "7E's - EXTEND | Apply learning in a new task | 5 Minutes",
        assessmentTitle: "Evaluate and Reflect",
        assessmentSub: "7E's - EVALUATE | 5 Minutes | Show what you learned!",
      };
    case "fourAs":
    default:
      return {
        objectivesSub: "By the end of this lesson, YOU will be able to...",
        activitySub: (duration) => `4A's - Phase 1: ACTIVITY | ${duration} | Group Work`,
        analysisSub: "4A's - Phase 2: ANALYSIS | Compare and observe!",
        definitionSub: "4A's - Phase 3: ABSTRACTION | The Key Concept",
        conceptsSub: (part) => `4A's - Phase 3: ABSTRACTION | (${part} of 2)`,
        applicationSub: "4A's - Phase 4: APPLICATION | Individual Activity | 5 Minutes",
        assessmentTitle: "Formative Assessment",
        assessmentSub: "4A's checkpoint | 5 Minutes | Answer independently!",
      };
  }
}

function getSlideLabels(frameworkKey) {
  switch (frameworkKey) {
    case "fiveEs":
      return [
        { icon: "1", label: "Title" },
        { icon: "2", label: "Engage" },
        { icon: "3", label: "Engage 2" },
        { icon: "4", label: "Explore" },
        { icon: "5", label: "Explore 2" },
        { icon: "6", label: "Explain" },
        { icon: "7", label: "Explain 2" },
        { icon: "8", label: "Explain 3" },
        { icon: "9", label: "Elaborate" },
        { icon: "10", label: "Evaluate" },
        { icon: "11", label: "Assignment" },
        { icon: "12", label: "Closing" },
      ];
    case "sevenEs":
      return [
        { icon: "1", label: "Title" },
        { icon: "2", label: "Elicit" },
        { icon: "3", label: "Engage" },
        { icon: "4", label: "Explore" },
        { icon: "5", label: "Explore 2" },
        { icon: "6", label: "Explain" },
        { icon: "7", label: "Elaborate" },
        { icon: "8", label: "Elaborate 2" },
        { icon: "9", label: "Extend" },
        { icon: "10", label: "Extend 2" },
        { icon: "11", label: "Evaluate" },
        { icon: "12", label: "Assignment" },
        { icon: "13", label: "Closing" },
      ];
    case "fourAs":
    default:
      return [
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
  }
}

function buildPptx(data) {
  const pres = new PptxGenJS();
  pres.layout = "LAYOUT_16x9";
  pres.title = data.topic;
  const template = TEMPLATE_OPTIONS[data.template] || TEMPLATE_OPTIONS.classroom;
  const theme = PALETTES[data.palette] || PALETTES.rainbow;
  const framework = getFrameworkContent(data.framework);
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

  function addVisualFrame(slide, opts) {
    const { x, y, w, h, col, light, label, imageData, accent } = opts;
    card(slide, x, y, w, h, col || P.blue, light || P.blueL);
    if (imageData) {
      slide.addImage({ data: imageData, x: x + 0.12, y: y + 0.12, w: w - 0.24, h: h - 0.24, rounding: true });
      return;
    }
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: x + 0.2,
      y: y + 0.22,
      w: w - 0.4,
      h: h - 0.44,
      rectRadius: 0.08,
      fill: { color: P.white, transparency: 28 },
      line: { color: accent || col || P.blue, dashType: "dash", pt: 1.5 },
    });
    slide.addShape(pres.shapes.OVAL, {
      x: x + w / 2 - 0.28,
      y: y + 0.42,
      w: 0.56,
      h: 0.56,
      fill: { color: accent || col || P.blue },
      line: { color: accent || col || P.blue },
    });
    slide.addText("IMG", {
      x: x + w / 2 - 0.28,
      y: y + 0.42,
      w: 0.56,
      h: 0.56,
      fontSize: 11,
      bold: true,
      color: P.white,
      align: "center",
      valign: "middle",
      margin: 0,
    });
    slide.addText(label || "Insert image here", {
      x: x + 0.35,
      y: y + h / 2 - 0.05,
      w: w - 0.7,
      h: 0.6,
      fontSize: 12,
      bold: true,
      color: accent || col || P.blue,
      align: "center",
      valign: "middle",
      margin: 0,
    });
    slide.addText("Teacher can replace this with a local image or diagram.", {
      x: x + 0.4,
      y: y + h / 2 + 0.35,
      w: w - 0.8,
      h: 0.42,
      fontSize: 9.5,
      color: P.black,
      italic: true,
      align: "center",
      valign: "middle",
      margin: 0,
    });
  }

  function addFrameworkSpotlight(slide, title, body, col, light) {
    card(slide, 0.35, 1.58, 9.3, 3.5, col, light);
    slide.addText(title, { x: 0.6, y: 1.8, w: 8.8, h: 0.45, fontSize: 24, bold: true, color: col, margin: 0 });
    slide.addText(body, { x: 0.6, y: 2.4, w: 5.0, h: 1.6, fontSize: 15, color: P.black, margin: 0 });
    addVisualFrame(slide, {
      x: 6.05,
      y: 2.02,
      w: 3.2,
      h: 2.45,
      col,
      light,
      label: "Add supporting visual",
      imageData: data.images?.concept,
      accent: col,
    });
  }

  {
    const sl = pres.addSlide();
    sl.background = { color: P.dark };
    [[0, 0, 3.5, P.purple, 55], [6.5, 3.5, 4, P.blue, 60], [4, 1.5, 2.5, P.teal, 65], [-0.5, 3.8, 3, P.pink, 65], [7.5, -0.5, 2.5, P.orange, 65]]
      .forEach(([x, y, s, c, t]) => sl.addShape(pres.shapes.OVAL, { x, y, w: s, h: s, fill: { color: c, transparency: t }, line: { color: c, transparency: t } }));
    if (template.layoutVariant === "split") {
      sl.addShape(pres.shapes.RECTANGLE, { x: 0.45, y: 0.62, w: 4.9, h: 3.45, fill: { color: P.white, transparency: 10 }, line: { color: P.white, transparency: 30 } });
      addVisualFrame(sl, { x: 5.7, y: 0.72, w: 3.65, h: 3.2, col: P.blue, light: P.blueL, label: "Add title image", imageData: data.images?.title, accent: P.blue });
      sl.addText(data.topic, { x: 0.65, y: 0.9, w: 4.4, h: 1.35, fontSize: 34, bold: true, color: P.white, valign: "middle", margin: 0 });
      sl.addText(data.subject, { x: 0.65, y: 2.28, w: 4.4, h: 0.45, fontSize: 18, color: P.yellowL, valign: "middle", margin: 0 });
    } else {
      sl.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 0.7, w: 9, h: 2.6, fill: { color: P.white, transparency: 10 }, line: { color: P.white, transparency: 30 } });
      sl.addText(data.topic, { x: 0.6, y: 0.8, w: 8.8, h: 1.3, fontSize: template.layoutVariant === "expressive" ? 40 : 44, bold: true, color: P.white, valign: "middle", margin: 0 });
      sl.addText(data.subject, { x: 0.6, y: 2.1, w: 8.8, h: 0.52, fontSize: 20, color: P.yellowL, valign: "middle", margin: 0 });
      if (template.layoutVariant === "expressive") {
        addVisualFrame(sl, { x: 6.3, y: 0.95, w: 2.5, h: 1.8, col: P.orange, light: P.orangeL, label: "Add cover visual", imageData: data.images?.title, accent: P.orange });
      }
    }
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
    hdr(sl, "Learning Objectives", framework.objectivesSub, P.purple);
    [
      { letter: "A", label: "KNOWLEDGE", icon: "🧠", text: data.objectives.knowledge, col: P.purple, lt: P.purpleL },
      { letter: "B", label: "SKILLS", icon: "✍", text: data.objectives.skills, col: P.blue, lt: P.blueL },
      { letter: "C", label: "ATTITUDE", icon: "💡", text: data.objectives.attitude, col: P.green, lt: P.greenL },
    ].forEach((o, i) => {
      const y = 1.5 + i * 1.17;
      card(sl, 0.35, y, 9.3, 1.05, o.col, o.lt);
      sl.addShape(pres.shapes.OVAL, { x: 0.5, y: y + 0.18, w: 0.68, h: 0.68, fill: { color: o.col }, line: { color: o.col } });
      sl.addText(o.letter, { x: 0.5, y: y + 0.18, w: 0.68, h: 0.68, fontSize: 20, bold: true, color: P.white, align: "center", valign: "middle", margin: 0 });
      sl.addText(o.icon, { x: 1.24, y: y + 0.18, w: 0.55, h: 0.55, fontSize: 18, margin: 0, align: "center", valign: "middle" });
      sl.addText(o.label, { x: 1.9, y: y + 0.13, w: 2.5, h: 0.32, fontSize: 11, bold: true, color: o.col, margin: 0, charSpacing: 2 });
      sl.addText(o.text, { x: 1.9, y: y + 0.45, w: 7.55, h: 0.52, fontSize: 12, color: P.black, margin: 0 });
    });
    sl.addText(`Competency: ${data.competency}`, { x: 0.35, y: 5.0, w: 9.3, h: 0.25, fontSize: 9, color: P.purple, italic: true, margin: 0 });
  }

  if (data.framework === "fiveEs" || data.framework === "sevenEs") {
    const sl = pres.addSlide();
    const hookColor = data.framework === "sevenEs" ? P.blue : P.orange;
    const hookLight = data.framework === "sevenEs" ? P.blueL : P.orangeL;
    hdr(sl, data.hook.title, data.hook.subtitle, hookColor);
    addFrameworkSpotlight(sl, data.hook.prompt, data.hook.body, hookColor, hookLight);
  }

  {
    const sl = pres.addSlide();
    const act = data.activity;
    hdr(sl, `Activity: ${act.title}`, framework.activitySub(act.duration), P.orange);
    badge(sl, 0.35, 1.47, `TIME ${act.duration}`, P.orange);
    badge(sl, 2.18, 1.47, "GROUP WORK", P.purple);
    card(sl, 0.35, 1.9, template.layoutVariant === "split" ? 4.95 : 5.65, 3.2, P.orange, P.orangeL);
    sl.addText("What To Do:", { x: 0.52, y: 2.02, w: 5.3, h: 0.38, fontSize: 14, bold: true, color: P.orange, margin: 0 });
    (act.steps || []).forEach((step, i) => {
      const y = 2.48 + i * 0.5;
      sl.addText(`${i + 1}.`, { x: 0.5, y, w: 0.5, h: 0.44, fontSize: 16, align: "center", valign: "middle", margin: 0 });
      sl.addText(step, { x: 1.05, y: y + 0.04, w: template.layoutVariant === "split" ? 4.15 : 4.82, h: 0.36, fontSize: 11, color: P.black, margin: 0 });
    });
    if (template.layoutVariant === "split") {
      addVisualFrame(sl, { x: 5.5, y: 1.9, w: 4.15, h: 2.1, col: P.teal, light: P.tealL, label: "Add activity visual", imageData: data.images?.activity, accent: P.teal });
      card(sl, 5.5, 4.12, 4.15, 0.98, P.purple, P.purpleL);
      sl.addText(`Guide: "${act.guideQuestion}"`, { x: 5.7, y: 4.3, w: 3.8, h: 0.25, fontSize: 10.5, color: P.black, italic: true, margin: 0 });
      sl.addText((act.materials || []).map((m, i) => ({ text: m, options: { bullet: true, breakLine: i < act.materials.length - 1, fontSize: 10, color: P.black } })), { x: 5.7, y: 4.58, w: 3.7, h: 0.34 });
    } else {
      card(sl, 6.18, 1.9, 3.48, 1.42, P.purple, P.purpleL);
      sl.addText("Guide Question", { x: 6.32, y: 2.0, w: 3.22, h: 0.38, fontSize: 12, bold: true, color: P.purple, margin: 0 });
      sl.addText(`"${act.guideQuestion}"`, { x: 6.32, y: 2.42, w: 3.22, h: 0.82, fontSize: 11.5, color: P.black, italic: true, margin: 0 });
      card(sl, 6.18, 3.45, 3.48, 1.65, P.teal, P.tealL);
      sl.addText("Materials:", { x: 6.32, y: 3.55, w: 3.22, h: 0.35, fontSize: 12, bold: true, color: P.teal, margin: 0 });
      sl.addText((act.materials || []).map((m, i) => ({ text: m, options: { bullet: true, breakLine: i < act.materials.length - 1, fontSize: 11, color: P.black } })), { x: 6.32, y: 3.96, w: 3.22, h: 1.1 });
    }
  }

  {
    const sl = pres.addSlide();
    const ana = data.analysis;
    hdr(sl, ana.title, framework.analysisSub, P.teal);
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
    hdr(sl, def.title, framework.definitionSub, P.purple);
    if (template.layoutVariant === "split") {
      card(sl, 0.35, 1.52, 5.25, 3.55, P.purple, P.purpleL);
      sl.addText("Definition", { x: 0.55, y: 1.68, w: 4.7, h: 0.35, fontSize: 12, bold: true, color: P.purple, margin: 0 });
      sl.addText(def.text, { x: 0.55, y: 2.08, w: 4.7, h: 1.5, fontSize: 12.5, color: P.black, margin: 0 });
      addVisualFrame(sl, { x: 5.88, y: 1.52, w: 3.77, h: 2.28, col: P.blue, light: P.blueL, label: "Add diagram or model", imageData: data.images?.concept, accent: P.blue });
    } else {
      sl.addShape(pres.shapes.RECTANGLE, { x: 0.35, y: 1.5, w: 9.3, h: 1.62, fill: { color: P.purple }, line: { color: P.purple } });
      sl.addText("Definition", { x: 0.55, y: 1.6, w: 9, h: 0.35, fontSize: 12, bold: true, color: P.purpleL, margin: 0 });
      sl.addText(def.text, { x: 0.55, y: 1.95, w: 9, h: 1.05, fontSize: 13, color: P.white, margin: 0 });
    }
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
    hdr(sl, `Key Concepts - Part ${slideIdx + 1}`, framework.conceptsSub(slideIdx + 1), colSet);
    concepts.slice(start, end).forEach((c, i) => {
      const col = i % 2;
      const row = Math.floor(i / 2);
      const x = 0.35 + col * 4.85;
      const y = 1.5 + row * 2.0;
      const ci = (start + i) % 8;
      card(sl, x, y, 4.62, 1.85, ACCENTS[ci], LIGHTS[ci]);
      sl.addShape(pres.shapes.OVAL, { x: x + 0.15, y: y + 0.2, w: 0.58, h: 0.58, fill: { color: ACCENTS[ci] }, line: { color: ACCENTS[ci] } });
      sl.addText(c.icon || "•", { x: x + 0.15, y: y + 0.2, w: 0.58, h: 0.58, fontSize: 16, align: "center", valign: "middle", margin: 0 });
      sl.addText((c.label || "").toUpperCase(), { x: x + 0.83, y: y + 0.22, w: 3.65, h: 0.35, fontSize: 11.5, bold: true, color: ACCENTS[ci], margin: 0, charSpacing: 1 });
      sl.addText((c.words || []).join(", "), { x: x + 0.15, y: y + 0.86, w: 4.35, h: 0.45, fontSize: 10, color: ACCENTS[ci], italic: true, margin: 0 });
      sl.addShape(pres.shapes.LINE, { x: x + 0.15, y: y + 1.35, w: 4.2, h: 0, line: { color: ACCENTS[ci], width: 1, dashType: "dash" } });
      sl.addText(`e.g. ${c.example || ""}`, { x: x + 0.15, y: y + 1.42, w: 4.35, h: 0.36, fontSize: 9.5, color: P.black, margin: 0 });
    });
  });

  {
    const sl = pres.addSlide();
    const app = data.application;
    hdr(sl, app.title, framework.applicationSub, P.orange);
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

  if (data.framework === "sevenEs") {
    const sl = pres.addSlide();
    hdr(sl, data.extend.title, data.extend.subtitle, P.green);
    addFrameworkSpotlight(sl, data.extend.prompt, data.extend.body, P.green, P.greenL);
  }

  {
    const sl = pres.addSlide();
    const qs = data.assessment || [];
    hdr(sl, framework.assessmentTitle, framework.assessmentSub, P.red);
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
    addVisualFrame(sl, { x: 7.25, y: 3.55, w: 2.0, h: 1.2, col: P.blue, light: P.blueL, label: "Add closing image", imageData: data.images?.closing, accent: P.blue });
    sl.addText("Great work today, class!", { x: 0.38, y: 4.22, w: 6.55, h: 0.42, fontSize: 14, bold: true, color: P.yellowL, align: "center", margin: 0 });
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

function normalizeSubject(subject) {
  const value = `${subject}`.toLowerCase();
  if (value.includes("english")) return "english";
  if (value.includes("filipino")) return "filipino";
  if (value.includes("science")) return "science";
  if (value.includes("math")) return "math";
  if (value.includes("araling")) return "ap";
  if (value.includes("history")) return "history";
  if (value.includes("literature")) return "literature";
  if (value.includes("eapp")) return "eapp";
  return "general";
}

function getSubjectProfile(subject) {
  const normalized = normalizeSubject(subject);
  const common = {
    partnerTask: "Talk with a partner",
    studentProduct: "student output",
    realWorld: "school and daily life",
    misconception: "a shallow understanding of the topic",
    strongMove: "a clearer, more purposeful use of the topic",
    extendTask: "apply the concept in a new context",
  };

  const profiles = {
    english: {
      ...common,
      focus: "communication, structure, and meaning",
      partnerTask: "Compare how the language choices affect meaning",
      studentProduct: "short paragraph or spoken response",
      realWorld: "reading, writing, and communication",
      misconception: "a sentence or paragraph that names the topic but does not use it effectively",
      strongMove: "a revised paragraph that uses the topic purposefully and clearly",
      extendTask: "use the topic in a fresh writing or speaking situation",
      concepts: ["purpose", "language choice", "organization", "effect on the reader"],
    },
    filipino: {
      ...common,
      focus: "pakikipagtalastasan, kahulugan, at organisasyon ng ideya",
      partnerTask: "Pag-usapan kung paano nakatutulong ang paksa sa mas malinaw na pagpapahayag",
      studentProduct: "maikling talata o sagot na pasalita",
      realWorld: "pakikipag-usap, pagbasa, at pagsulat",
      misconception: "isang pahayag na binabanggit ang paksa ngunit hindi malinaw ang gamit nito",
      strongMove: "isang inayos na pahayag na malinaw at makabuluhan ang gamit ng paksa",
      extendTask: "ilapat ang paksa sa panibagong sitwasyon o teksto",
      concepts: ["layunin", "salitang ginamit", "ayos ng ideya", "epekto sa tagapakinig o mambabasa"],
    },
    science: {
      ...common,
      focus: "evidence, explanation, and observable patterns",
      partnerTask: "Use observations and evidence to explain the pattern you notice",
      studentProduct: "claim-evidence-explanation response",
      realWorld: "natural phenomena and everyday observations",
      misconception: "an explanation that names the idea but gives no evidence or cause",
      strongMove: "an explanation that connects the topic to evidence, cause, and effect",
      extendTask: "apply the idea to a new phenomenon or classroom investigation",
      concepts: ["observation", "cause", "evidence", "application"],
    },
    math: {
      ...common,
      focus: "reasoning, representation, and problem solving",
      partnerTask: "Compare two solution paths and justify which one is more efficient",
      studentProduct: "worked solution with explanation",
      realWorld: "quantitative decisions and problem solving",
      misconception: "a solution that jumps to an answer without showing the reasoning",
      strongMove: "a solution that models the topic step by step and explains the reasoning",
      extendTask: "solve a new problem using the same concept in a different form",
      concepts: ["pattern", "strategy", "representation", "justification"],
    },
    ap: {
      ...common,
      focus: "context, perspective, and real-life connections",
      partnerTask: "Connect the topic to a familiar issue in community or national life",
      studentProduct: "short explanation with evidence or examples",
      realWorld: "community life, citizenship, and social issues",
      misconception: "a response that names the topic but ignores context or evidence",
      strongMove: "a response that uses examples, context, and explanation to support the idea",
      extendTask: "apply the concept to a new historical or community case",
      concepts: ["context", "perspective", "evidence", "action"],
    },
    history: {
      ...common,
      focus: "context, chronology, and interpretation",
      partnerTask: "Discuss how the topic changes when you consider time, place, and perspective",
      studentProduct: "timeline note or short explanation",
      realWorld: "historical interpretation and civic understanding",
      misconception: "a statement that names the event but ignores causes or context",
      strongMove: "an explanation that connects event, cause, and consequence",
      extendTask: "compare the topic to a different event or time period",
      concepts: ["cause", "effect", "context", "significance"],
    },
    literature: {
      ...common,
      focus: "interpretation, theme, and textual meaning",
      partnerTask: "Compare how the topic changes the reader's understanding of the text",
      studentProduct: "interpretive paragraph or spoken insight",
      realWorld: "reading, interpretation, and personal connection",
      misconception: "a response that repeats the text without interpreting the topic",
      strongMove: "an interpretation that explains how the topic shapes meaning",
      extendTask: "apply the same interpretation move to a new text or excerpt",
      concepts: ["theme", "character", "symbol", "reader response"],
    },
    eapp: {
      ...common,
      focus: "academic communication, structure, and purpose",
      partnerTask: "Judge how effectively the topic helps the speaker or writer communicate",
      studentProduct: "organized academic response",
      realWorld: "academic tasks, reporting, and presentations",
      misconception: "a response that mentions the concept but does not use it purposefully",
      strongMove: "a response that uses the concept to improve structure and clarity",
      extendTask: "apply the concept to a new academic communication task",
      concepts: ["purpose", "organization", "clarity", "audience"],
    },
    general: {
      ...common,
      focus: "understanding, application, and reflection",
      concepts: ["main idea", "purpose", "application", "reflection"],
    },
  };

  return profiles[normalized];
}

function buildTopicProfile(topic, subject) {
  const keywords = buildTopicKeywords(topic);
  const core = keywords[0] || topic.split(/\s+/)[0] || "concept";
  const phrase = keywords.slice(0, 3).join(", ");
  const subjectProfile = getSubjectProfile(subject);
  return {
    keywords,
    core,
    phrase,
    subjectProfile,
  };
}

function buildStudentObjectives(topic, profile, detailLevel, custom = {}) {
  const knowledge = clampWords(
    custom.knowledge || pickByDetail(
      detailLevel,
      `I can explain the key idea behind ${topic}.`,
      `I can explain how ${topic} works and why it matters in ${profile.subjectProfile.focus}.`,
      `I can explain how ${topic} works, identify its important parts, and connect it to ${profile.subjectProfile.focus}.`
    ),
    detailLevel === "minimal" ? 14 : detailLevel === "detailed" ? 28 : 20
  );
  const skills = clampWords(
    custom.skills || pickByDetail(
      detailLevel,
      `I can use ${topic} in a short task.`,
      `I can use ${topic} in a task, discussion, or response with a partner or group.`,
      `I can use ${topic} in a task, discussion, or response and explain my choices clearly to others.`
    ),
    detailLevel === "minimal" ? 14 : detailLevel === "detailed" ? 28 : 20
  );
  const attitude = clampWords(
    custom.attitude || pickByDetail(
      detailLevel,
      `I can stay curious and reflective while learning ${topic}.`,
      `I can stay curious, listen to others, and reflect on how ${topic} helps me learn.`,
      `I can stay curious, listen to others, and reflect on how ${topic} helps me learn, communicate, and solve problems in new situations.`
    ),
    detailLevel === "minimal" ? 14 : detailLevel === "detailed" ? 28 : 20
  );
  return { knowledge, skills, attitude };
}

function buildConceptSet(topic, profile) {
  const [k1 = "idea", k2 = "focus", k3 = "example", k4 = "process"] = profile.keywords;
  const focus = profile.subjectProfile.concepts;
  return [
    { label: "TOPIC FOCUS", icon: "💡", words: [k1, "core idea", "central meaning"], example: `${topic} asks learners to focus on the central idea, not just mention the term.` },
    { label: "PURPOSE", icon: "🎯", words: [focus[0], "goal", "reason"], example: `${topic} becomes clearer when students know its purpose.` },
    { label: "EVIDENCE OR SUPPORT", icon: "🧩", words: [focus[1], "support", "details"], example: `Strong work on ${topic} includes support, not just a short answer.` },
    { label: "PROCESS", icon: "🧭", words: [k4, "steps", "flow"], example: `Students can work through ${topic} step by step.` },
    { label: "EXAMPLE", icon: "🧪", words: [k3, "model", "sample"], example: `A clear example helps students understand ${topic} faster.` },
    { label: "REAL-LIFE LINK", icon: "🌍", words: [profile.subjectProfile.realWorld, k2], example: `${topic} matters because it connects to ${profile.subjectProfile.realWorld}.` },
    { label: "MISCONCEPTION", icon: "⚠", words: ["common error", "confusion", "surface answer"], example: `Students improve ${topic} when they notice surface-level mistakes early.` },
    { label: "REFLECTION", icon: "🔍", words: ["self-check", "insight", "growth"], example: `Reflection helps students revise and improve how they use ${topic}.` },
  ];
}

function createLocalLessonData(form) {
  const topic = form.topic.trim();
  const subject = form.subject || "English";
  const gradeLevel = form.gradeLevel || "8";
  const quarter = form.quarter || "First Quarter";
  const framework = form.framework || "fourAs";
  const frameworkLabel = FRAMEWORK_OPTIONS[framework]?.label || "4A's";
  const detailLevel = form.detailLevel || "balanced";
  const template = form.template || "classroom";
  const palette = form.palette || "rainbow";
  const templateConfig = TEMPLATE_OPTIONS[template] || TEMPLATE_OPTIONS.classroom;
  const paletteConfig = PALETTES[palette] || PALETTES.rainbow;
  const profile = buildTopicProfile(topic, subject);
  const keywords = profile.keywords;
  const lead = profile.core;
  const second = keywords[1] || "focus";
  const third = keywords[2] || "example";
  const studentObjectives = buildStudentObjectives(topic, profile, detailLevel, {
    knowledge: form.objKnowledge,
    skills: form.objSkills,
    attitude: form.objAttitude,
  });
  const concepts = buildConceptSet(topic, profile);
  const competency = clampWords(
    form.competency || `Students explore ${topic} through inquiry, collaboration, and student-created responses connected to ${profile.subjectProfile.realWorld}.`,
    detailLevel === "minimal" ? 14 : detailLevel === "detailed" ? 28 : 20
  );
  const activitySteps = pickByDetail(
    detailLevel,
    [
      `Look at the prompt, example, or problem about ${topic}.`,
      `Notice one idea, pattern, or choice connected to ${topic}.`,
      profile.subjectProfile.partnerTask,
      "Share one student insight with the class.",
    ],
    [
      `Look closely at the prompt, example, or problem about ${topic}.`,
      `List the important words, moves, or patterns you notice in ${topic}.`,
      `Discuss how ${topic} connects to ${profile.subjectProfile.realWorld}.`,
      "Work with your group to organize your ideas and prepare a short share-out.",
      "Share one key student insight with the class.",
    ],
    [
      `Look closely at the prompt, example, or problem about ${topic}.`,
      `List the important words, moves, or patterns you notice in ${topic}.`,
      `Discuss how ${topic} connects to ${profile.subjectProfile.realWorld}.`,
      "Compare your observations with a classmate.",
      "Organize your ideas as a group and prepare a claim, explanation, or solution.",
      "Share one key student insight with the class.",
    ]
  );

  return {
    topic,
    subject,
    gradeLevel,
    quarter,
    framework,
    frameworkLabel,
    detailLevel,
    template,
    palette,
    themeLabel: paletteConfig.label,
    competency,
    tagline: pickByDetail(
      detailLevel,
      `${topic} becomes powerful when students test ideas, explain them, and make them their own.`,
      `${topic} becomes meaningful when students explore it, talk through it, and apply it with purpose in the ${templateConfig.label.toLowerCase()} template.`,
      `${topic} becomes meaningful when students explore it, talk through it, revise their thinking, and apply it with purpose in the ${templateConfig.label.toLowerCase()} template.`
    ),
    images: form.images || {},
    objectives: {
      knowledge: studentObjectives.knowledge,
      skills: studentObjectives.skills,
      attitude: studentObjectives.attitude,
    },
    hook: {
      title: framework === "sevenEs" ? "Elicit and Engage" : "Engage the Lesson",
      subtitle: framework === "sevenEs" ? "7E's - ELICIT/ENGAGE | Connect prior knowledge to the new lesson" : "5E's - ENGAGE | Activate prior knowledge and curiosity",
      prompt: pickByDetail(
        detailLevel,
        `What do you already know about ${topic}?`,
        `What do you already know about ${topic}, and where have you already seen it in ${profile.subjectProfile.realWorld}?`,
        `What do you already know about ${topic}, where have you already seen it in ${profile.subjectProfile.realWorld}, and what do you still want to figure out?`
      ),
      body: pickByDetail(
        detailLevel,
        `Start with a student prompt, image, or quick challenge that gets everyone thinking about ${topic}.`,
        `Start with a student prompt, image, or quick challenge that gets everyone thinking about ${topic} before they explore the lesson more deeply.`,
        `Start with a student prompt, image, or quick challenge that gets everyone thinking about ${topic}. Let students connect it to prior knowledge, predict what they will discover, and raise their own questions.`
      ),
    },
    activity: {
      title: `Exploring ${lead}`,
      duration: "10 Minutes",
      steps: activitySteps,
      guideQuestion: `How can students show that they understand and can use ${topic}, not just define it?`,
      materials: pickByDetail(
        detailLevel,
        ["activity sheet", "pen or pencil"],
        ["activity sheet", "pen or pencil", "sample prompt or text"],
        ["activity sheet", "pen or pencil", "sample prompt or text", "optional visual reference"]
      ),
    },
    analysis: {
      title: `Comparing Student Work on ${lead}`,
      prompt: `Study both examples. Which one helps a student understand or use ${topic} more effectively?`,
      versionA: {
        label: "Surface-level response",
        text: pickByDetail(
          detailLevel,
          `The student mentions ${topic}, but the idea stays brief and unclear.`,
          `The student mentions ${topic}, but the response stays short and lacks evidence, explanation, or a strong example. The audience may understand only part of the idea.`,
          `The student mentions ${topic}, but the response stays short and lacks evidence, explanation, or a strong example. The audience may understand only part of the idea, and the work does not yet show how the concept functions in practice.`
        ),
        note: `This response stays at the surface and does not yet show strong ${profile.subjectProfile.focus}.`,
      },
      versionB: {
        label: "Stronger student response",
        text: pickByDetail(
          detailLevel,
          `${topic} becomes clearer when the student gives a focused idea and a useful example.`,
          `${topic} becomes clearer when the student gives a focused idea, a useful example, and a clear connection to the task or real situation.`,
          `${topic} becomes clearer when the student gives a focused idea, a useful example, and a clear connection to the task or real situation. The response shows reasoning, not just recall.`
        ),
        note: `This response is stronger because the student makes thinking visible and uses ${topic} purposefully.`,
      },
      discussion: `What makes the second student response stronger? Which choices made ${topic} easier to notice, understand, or apply?`,
    },
    definition: {
      title: `Understanding ${topic}`,
      text: pickByDetail(
        detailLevel,
        `${topic} helps students understand an idea and use it in a real task.`,
        `${topic} helps students build knowledge, practice a skill, and connect ideas to real situations. It matters most when learners can explain it clearly, identify examples, and use it in their own work.`,
        `${topic} helps students build knowledge, practice a skill, and connect ideas to real situations. It matters most when learners can explain it clearly, identify examples, use it in their own work, and reflect on how their understanding changed.`
      ),
      purposes: [
        { icon: "📘", title: "Build Understanding", desc: `Students identify the main ideas behind ${topic}.` },
        { icon: "🛠", title: "Try It Out", desc: `Learners use ${topic} in discussion, problem solving, or creation.` },
        { icon: "⭐", title: "Make It Matter", desc: `${topic} becomes meaningful when learners connect it to ${profile.subjectProfile.realWorld}.` },
      ],
    },
    concepts,
    application: {
      title: "Student Practice",
      wordBox: keywords.slice(0, 5),
      items: pickByDetail(
        detailLevel,
        [
          `Write one sentence that uses ${topic} to show clear thinking.`,
          `Give one example that shows ${topic} in ${profile.subjectProfile.realWorld}.`,
          `Explain how a student can improve their work on ${topic}.`,
          `Reflect: what part of ${topic} is easier for you now?`,
        ],
        [
          `Write one short response, solution, or explanation that shows ${topic} clearly.`,
          `Use one example or detail to support your thinking about ${topic}.`,
          `Explain how ${topic} connects to ${profile.subjectProfile.realWorld}.`,
          `Revise one part of your work so ${topic} becomes clearer or stronger.`,
          `Reflect on what helped you understand ${topic} today.`,
        ],
        [
          `Write one short response, solution, or explanation that shows ${topic} clearly.`,
          `Use one example or detail to support your thinking about ${topic}.`,
          `Explain how ${topic} connects to ${profile.subjectProfile.realWorld}.`,
          `Revise one part of your work so ${topic} becomes clearer or stronger.`,
          `Give feedback to a classmate about how they used ${topic}.`,
        ]
      ),
    },
    extend: {
      title: "Extend the Learning",
      subtitle: "7E's - EXTEND | Transfer the concept to a new situation",
      prompt: `How can learners apply ${topic} beyond today's activity?`,
      body: pickByDetail(
        detailLevel,
        `Invite students to use ${topic} in a different subject, home situation, or real-world example.`,
        `Invite students to use ${topic} in a different subject, home situation, or real-world example so they can see how the concept transfers beyond the first task.`,
        `Invite students to use ${topic} in a different subject, home situation, or real-world example. This helps them transfer the idea, compare contexts, and explain why the concept still matters.`
      ),
    },
    assessment: [
      { points: 2, question: `What is one key idea every student should understand about ${topic}?` },
      { points: 2, question: `Give one example that shows ${topic} clearly.` },
      { points: 3, question: `Explain how ${topic} helps a learner in ${subject}.` },
      { points: 3, question: `Create a short response, solution, or explanation that applies ${topic}.` },
    ],
    assignment: {
      task: `Create your own ${profile.subjectProfile.studentProduct} that uses ${topic}. Include an example, a clear explanation, and a short reflection on what choice you made as a learner.`,
      checklist: ["Use your own words or reasoning", "Include one clear example or support", "Add one reflection or student voice note"],
      topics: [
        `${topic} in school life`,
        `${topic} in home situations`,
        `${topic} in the community`,
        `${topic} in a situation you care about`,
      ],
    },
    closingQuote: pickByDetail(
      detailLevel,
      `${topic} grows when students question, test, and use it for themselves.`,
      `${topic} grows when students question, test, explain, and use it for themselves.`,
      `${topic} grows when students question, test, explain, revise, and use it for themselves in new contexts.`
    ),
    takeaways: [
      `Students can explain ${topic}`,
      `Students can apply ${topic}`,
      `Students can reflect on ${topic}`,
    ],
  };
}

export default function App() {
  const [form, setForm] = useState({
    topic: "",
    subject: "English",
    gradeLevel: "8",
    quarter: "First Quarter",
    framework: "fourAs",
    detailLevel: "balanced",
    template: "classroom",
    palette: "rainbow",
    competency: "",
    objKnowledge: "",
    objSkills: "",
    objAttitude: "",
    extraContext: "",
    images: {
      title: "",
      activity: "",
      concept: "",
      closing: "",
    },
  });

  const [status, setStatus] = useState("idle");
  const [progress, setProgress] = useState("");
  const [generatedData, setGeneratedData] = useState(null);
  const [errorMsg, setErrorMsg] = useState("");
  const [currentStep, setCurrentStep] = useState(0);

  const subjectOptions = ["English", "Filipino", "Science", "Mathematics", "Araling Panlipunan", "EAPP", "Literature", "History"];
  const gradeOptions = ["7", "8", "9", "10", "11", "12"];
  const quarterOptions = ["First Quarter", "Second Quarter", "Third Quarter", "Fourth Quarter"];
  const wizardSteps = [
    { id: "lesson", number: "01", title: "Lesson Setup", subtitle: "Topic, subject, grade, and competency" },
    { id: "design", number: "02", title: "Design Style", subtitle: "Template and color palette" },
    { id: "objectives", number: "03", title: "Objectives", subtitle: "Learning goals and extra notes" },
    { id: "review", number: "04", title: "Review & Generate", subtitle: "Check your choices before export" },
  ];

  const set = (k, v) => setForm((f) => ({ ...f, [k]: v }));
  const setImage = (slot, value) => setForm((f) => ({ ...f, images: { ...f.images, [slot]: value } }));
  const canGoNext = currentStep === 0
    ? Boolean(form.topic.trim() && form.subject.trim())
    : currentStep === 1
      ? Boolean(form.template && form.palette)
      : true;
  const nextStep = () => {
    if (!canGoNext) {
      setErrorMsg("Complete the required fields in this step before continuing.");
      setStatus("error");
      return;
    }
    setStatus("idle");
    setErrorMsg("");
    setCurrentStep((step) => Math.min(step + 1, wizardSteps.length - 1));
  };
  const prevStep = () => {
    setStatus("idle");
    setErrorMsg("");
    setCurrentStep((step) => Math.max(step - 1, 0));
  };

  async function handleImageChange(slot, event) {
    const file = event.target.files?.[0];
    if (!file) {
      setImage(slot, "");
      return;
    }
    const reader = new FileReader();
    reader.onload = () => setImage(slot, `${reader.result}`);
    reader.readAsDataURL(file);
  }

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
    :root{
      --surface:#eef3ff;
      --surface-2:#f6f8ff;
      --surface-3:#e8eefc;
      --ink:#111827;
      --muted:#6B7280;
      --accent:#7C3AED;
      --accent-2:#DB2777;
      --warm:#EA580C;
      --shadow-hi:rgba(255,255,255,.92);
      --shadow-lo:rgba(135,152,190,.22);
      --shadow-mid:rgba(123,97,255,.14);
    }
    *{box-sizing:border-box;margin:0;padding:0}
    body{
      font-family:'Nunito',sans-serif;
      background:
        radial-gradient(circle at top left,#ffffff 0%,transparent 35%),
        radial-gradient(circle at bottom right,#ffe7ef 0%,transparent 28%),
        linear-gradient(145deg,#eef3ff 0%,#e8eefc 48%,#f5f7ff 100%);
      min-height:100vh;
      color:var(--ink)
    }
    .app{min-height:100vh;padding:24px 16px 36px;display:flex;flex-direction:column;align-items:center}
    .hero{text-align:center;margin-bottom:32px;padding-top:8px}
    .hero-badge{
      display:inline-flex;align-items:center;gap:6px;color:#fff;border-radius:999px;padding:8px 18px;font-size:12px;font-weight:800;
      letter-spacing:2px;margin-bottom:14px;text-transform:uppercase;
      background:linear-gradient(135deg,#7C3AED,#DB2777);
      box-shadow:8px 8px 16px rgba(139,155,190,.22),-8px -8px 16px rgba(255,255,255,.95)
    }
    .hero-title{font-size:clamp(28px,5vw,46px);font-weight:900;color:#1E1B4B;line-height:1.1;margin-bottom:10px;font-family:'Quicksand',sans-serif}
    .hero-title span{background:linear-gradient(135deg,#7C3AED,#DB2777,#EA580C);-webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text}
    .hero-sub{font-size:15px;color:#6B7280;max-width:520px;margin:0 auto;font-weight:600}
    .card{
      background:linear-gradient(145deg,var(--surface-2),var(--surface));
      border-radius:30px;
      box-shadow:18px 18px 36px var(--shadow-lo),-18px -18px 36px var(--shadow-hi),0 20px 50px rgba(124,58,237,.08);
      padding:32px;width:100%;max-width:940px;margin-bottom:20px;border:1px solid rgba(255,255,255,.65)
    }
    .section-title{font-size:13px;font-weight:800;color:#7C3AED;letter-spacing:2px;text-transform:uppercase;margin-bottom:16px;display:flex;align-items:center;gap:8px}
    .section-title::after{content:'';flex:1;height:2px;background:linear-gradient(90deg,#7C3AED33,transparent)}
    .grid-2{display:grid;grid-template-columns:1fr 1fr;gap:16px}
    .field{display:flex;flex-direction:column;gap:6px;margin-bottom:16px}
    .field label{font-size:13px;font-weight:700;color:#374151}
    .field input,.field select,.field textarea{
      border:1px solid rgba(255,255,255,.7);border-radius:18px;padding:14px 16px;font-size:14px;font-family:'Nunito',sans-serif;
      color:#111;background:linear-gradient(145deg,#edf2ff,#f8faff);transition:all .2s;outline:none;resize:vertical;font-weight:600;
      box-shadow:inset 6px 6px 12px rgba(189,199,225,.45),inset -6px -6px 12px rgba(255,255,255,.92);
    }
    .field input:focus,.field select:focus,.field textarea:focus{
      border-color:#c4b5fd;background:linear-gradient(145deg,#f7f4ff,#ffffff);
      box-shadow:inset 4px 4px 10px rgba(189,199,225,.35),inset -4px -4px 10px rgba(255,255,255,.98),0 0 0 4px #7C3AED15
    }
    .field textarea{min-height:72px}
    .btn-generate{
      width:100%;padding:18px;border-radius:16px;border:none;cursor:pointer;
      background:linear-gradient(135deg,#7C3AED,#DB2777);
      color:#fff;font-size:17px;font-weight:800;font-family:'Nunito',sans-serif;
      letter-spacing:.5px;transition:all .2s;position:relative;overflow:hidden;
      box-shadow:10px 10px 20px rgba(151,103,180,.28),-8px -8px 18px rgba(255,255,255,.72)
    }
    .btn-generate:hover:not(:disabled){transform:translateY(-2px);box-shadow:14px 14px 24px rgba(151,103,180,.32),-10px -10px 18px rgba(255,255,255,.78)}
    .btn-generate:active:not(:disabled){transform:translateY(0);box-shadow:inset 5px 5px 12px rgba(94,40,140,.32),inset -5px -5px 12px rgba(255,255,255,.18)}
    .btn-generate:disabled{opacity:.7;cursor:not-allowed}
    .loading-box{
      text-align:center;padding:32px;background:linear-gradient(145deg,#eef2ff,#f8faff);border-radius:26px;
      box-shadow:inset 8px 8px 18px rgba(191,200,226,.34),inset -10px -10px 20px rgba(255,255,255,.94)
    }
    .spinner{
      width:52px;height:52px;border:5px solid rgba(229,231,235,.8);border-top-color:#7C3AED;border-radius:50%;animation:spin 1s linear infinite;margin:0 auto 16px;
      box-shadow:6px 6px 16px rgba(148,163,184,.18),-6px -6px 16px rgba(255,255,255,.92)
    }
    @keyframes spin{to{transform:rotate(360deg)}}
    .progress-txt{font-size:15px;font-weight:700;color:#7C3AED;margin-bottom:4px}
    .progress-sub{font-size:13px;color:#9CA3AF}
    .steps-row{display:flex;gap:8px;justify-content:center;margin-top:16px;flex-wrap:wrap}
    .step-dot{width:10px;height:10px;border-radius:50%;background:#E5E7EB;transition:all .4s;box-shadow:4px 4px 8px rgba(148,163,184,.18),-4px -4px 8px rgba(255,255,255,.95)}
    .step-dot.active{background:#7C3AED;transform:scale(1.3)}
    .success-box{
      text-align:center;padding:28px;background:linear-gradient(145deg,#eef8f2,#f6fdf8);border-radius:26px;
      box-shadow:18px 18px 36px rgba(163,177,198,.18),-18px -18px 30px rgba(255,255,255,.92),inset 0 0 0 1px rgba(255,255,255,.68)
    }
    .success-icon{font-size:52px;margin-bottom:12px}
    .success-title{font-size:20px;font-weight:900;color:#16A34A;margin-bottom:6px}
    .success-sub{font-size:14px;color:#6B7280;margin-bottom:20px;font-weight:600}
    .wizard-shell{display:grid;grid-template-columns:240px minmax(0,1fr);gap:24px;align-items:start}
    .wizard-rail{
      background:linear-gradient(145deg,#f3f0ff,#fbfcff);
      border:1px solid rgba(255,255,255,.78);
      border-radius:28px;padding:18px;
      box-shadow:16px 16px 28px rgba(169,179,204,.2),-14px -14px 24px rgba(255,255,255,.96)
    }
    .wizard-rail-title{font-size:12px;font-weight:800;letter-spacing:.16em;color:#7C3AED;text-transform:uppercase;margin-bottom:14px}
    .wizard-step{
      padding:14px 14px 14px 16px;border-radius:22px;border:1px solid transparent;transition:all .2s;
      background:linear-gradient(145deg,#eff3ff,#f9fbff);
      box-shadow:8px 8px 16px rgba(175,184,210,.12),-8px -8px 16px rgba(255,255,255,.8);margin-bottom:10px
    }
    .wizard-step.active{
      background:linear-gradient(145deg,#f8f2ff,#ffffff);border-color:#D8B4FE;
      box-shadow:inset 1px 1px 0 rgba(255,255,255,.8),12px 12px 26px rgba(159,139,201,.18),-10px -10px 22px rgba(255,255,255,.98)
    }
    .wizard-step.done{opacity:.8}
    .wizard-num{font-size:11px;font-weight:900;letter-spacing:.14em;color:#A78BFA;margin-bottom:6px}
    .wizard-step.active .wizard-num{color:#7C3AED}
    .wizard-name{font-size:14px;font-weight:800;color:#111827;margin-bottom:4px}
    .wizard-sub{font-size:12px;color:#6B7280;font-weight:600;line-height:1.35}
    .wizard-main{
      background:linear-gradient(145deg,#f4f7ff,#edf2ff);
      border:1px solid rgba(255,255,255,.78);
      border-radius:30px;padding:28px;
      box-shadow:18px 18px 32px rgba(164,176,205,.2),-16px -16px 30px rgba(255,255,255,.98)
    }
    .wizard-head{display:flex;align-items:flex-start;justify-content:space-between;gap:16px;margin-bottom:22px}
    .wizard-kicker{font-size:12px;font-weight:900;letter-spacing:.16em;color:#7C3AED;text-transform:uppercase;margin-bottom:8px}
    .wizard-title{font-size:28px;font-weight:900;color:#111827;line-height:1.05}
    .wizard-copy{font-size:14px;color:#6B7280;font-weight:600;max-width:440px;margin-top:8px}
    .wizard-panel{min-height:430px}
    .wizard-nav{display:flex;justify-content:space-between;gap:12px;margin-top:24px;padding-top:20px;border-top:1px solid rgba(227,220,248,.7)}
    .btn-secondary{
      padding:14px 20px;border-radius:16px;border:1px solid rgba(255,255,255,.75);background:linear-gradient(145deg,#eef3ff,#fdfdff);
      color:#374151;font-size:14px;font-weight:800;cursor:pointer;
      box-shadow:8px 8px 18px rgba(167,178,203,.18),-8px -8px 16px rgba(255,255,255,.95)
    }
    .btn-secondary:hover{background:linear-gradient(145deg,#f7f9ff,#ffffff)}
    .btn-secondary:active:not(:disabled){box-shadow:inset 5px 5px 12px rgba(175,184,210,.18),inset -5px -5px 12px rgba(255,255,255,.95)}
    .btn-secondary:disabled{opacity:.5;cursor:not-allowed}
    .summary-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:12px}
    .summary-card{
      border:1px solid rgba(255,255,255,.76);border-radius:22px;padding:16px;background:linear-gradient(145deg,#f1f5ff,#fbfcff);
      box-shadow:10px 10px 20px rgba(170,179,205,.14),-8px -8px 18px rgba(255,255,255,.96)
    }
    .summary-label{font-size:11px;font-weight:900;letter-spacing:.12em;color:#8B5CF6;text-transform:uppercase;margin-bottom:8px}
    .summary-value{font-size:15px;font-weight:700;color:#111827;line-height:1.35}
    .preview-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(120px,1fr));gap:8px;margin:16px 0}
    .preview-pill{
      background:linear-gradient(145deg,#f0f4ff,#ffffff);border:1px solid rgba(255,255,255,.82);border-radius:16px;padding:10px 12px;font-size:11px;font-weight:700;color:#374151;text-align:center;
      box-shadow:8px 8px 18px rgba(167,178,203,.16),-8px -8px 16px rgba(255,255,255,.96)
    }
    .preview-pill span{display:block;font-size:16px;margin-bottom:2px}
    .btn-dl{
      display:inline-flex;align-items:center;gap:8px;
      background:linear-gradient(135deg,#16A34A,#0891B2);color:#fff;
      border:none;border-radius:14px;padding:14px 28px;font-size:15px;font-weight:800;
      cursor:pointer;font-family:'Nunito',sans-serif;transition:all .2s;
      box-shadow:10px 10px 22px rgba(135,168,159,.24),-8px -8px 18px rgba(255,255,255,.78)
    }
    .btn-dl:hover{transform:translateY(-2px);box-shadow:12px 12px 24px rgba(135,168,159,.28),-9px -9px 18px rgba(255,255,255,.82)}
    .btn-new{
      display:inline-flex;align-items:center;gap:8px;margin-top:10px;
      background:linear-gradient(145deg,#eef3ff,#fbfcff);color:#374151;border:1px solid rgba(255,255,255,.75);
      border-radius:14px;padding:12px 22px;font-size:14px;font-weight:700;
      cursor:pointer;font-family:'Nunito',sans-serif;transition:all .2s;
      box-shadow:8px 8px 18px rgba(167,178,203,.16),-8px -8px 16px rgba(255,255,255,.96)
    }
    .btn-new:hover{background:linear-gradient(145deg,#f7f9ff,#ffffff)}
    .error-box{
      padding:20px;background:linear-gradient(145deg,#fff0f1,#ffe3e6);border-radius:20px;border:1px solid rgba(255,255,255,.72);
      box-shadow:10px 10px 20px rgba(220,38,38,.08),-8px -8px 18px rgba(255,255,255,.92)
    }
    .error-title{font-size:15px;font-weight:800;color:#DC2626;margin-bottom:4px}
    .error-msg{font-size:13px;color:#7F1D1D;font-weight:600}
    .tag{
      display:inline-block;background:linear-gradient(145deg,#efe8ff,#fcfbff);color:#7C3AED;border-radius:999px;padding:4px 12px;font-size:11px;font-weight:800;
      border:1px solid rgba(255,255,255,.75);box-shadow:6px 6px 12px rgba(167,178,203,.14),-6px -6px 12px rgba(255,255,255,.94)
    }
    .choice-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:12px}
    .choice-btn{
      border:1px solid rgba(255,255,255,.78);border-radius:22px;padding:16px;background:linear-gradient(145deg,#f0f4ff,#ffffff);cursor:pointer;text-align:left;transition:all .2s;
      box-shadow:10px 10px 18px rgba(167,178,203,.16),-10px -10px 18px rgba(255,255,255,.96)
    }
    .choice-btn.active{
      border-color:#d8b4fe;
      box-shadow:inset 1px 1px 0 rgba(255,255,255,.86),0 0 0 4px #7C3AED10,12px 12px 24px rgba(159,139,201,.18),-10px -10px 20px rgba(255,255,255,.98);
      transform:translateY(-1px)
    }
    .choice-title{font-size:14px;font-weight:800;color:#111827;margin-bottom:4px}
    .choice-text{font-size:12px;color:#6B7280;font-weight:600;line-height:1.4}
    .swatch-row{display:flex;gap:6px;margin-top:10px}
    .swatch{width:18px;height:18px;border-radius:999px;border:1px solid #ffffffaa;box-shadow:4px 4px 8px rgba(167,178,203,.15),-4px -4px 8px rgba(255,255,255,.94)}
    .upload-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:12px}
    .upload-card{
      border:1px solid rgba(255,255,255,.78);border-radius:22px;padding:16px;background:linear-gradient(145deg,#f0f4ff,#ffffff);
      box-shadow:10px 10px 18px rgba(167,178,203,.16),-10px -10px 18px rgba(255,255,255,.96)
    }
    .upload-head{display:flex;align-items:center;justify-content:space-between;gap:10px;margin-bottom:12px}
    .upload-title{font-size:14px;font-weight:800;color:#111827}
    .upload-note{font-size:12px;color:#6B7280;font-weight:600;line-height:1.4}
    .upload-card input[type="file"]{width:100%;margin-top:10px}
    .upload-preview{
      margin-top:12px;height:120px;border-radius:18px;background:linear-gradient(145deg,#edf2ff,#f9fbff);
      display:flex;align-items:center;justify-content:center;border:1px dashed rgba(124,58,237,.28);overflow:hidden;
      box-shadow:inset 6px 6px 12px rgba(189,199,225,.22),inset -6px -6px 12px rgba(255,255,255,.88)
    }
    .upload-preview img{width:100%;height:100%;object-fit:cover}
    .upload-placeholder{font-size:12px;font-weight:800;color:#7C3AED;letter-spacing:.06em;text-transform:uppercase}
    @media(max-width:600px){.grid-2{grid-template-columns:1fr}}
    @media(max-width:600px){.choice-grid{grid-template-columns:1fr}}
    @media(max-width:600px){.upload-grid{grid-template-columns:1fr}}
    @media(max-width:900px){.wizard-shell{grid-template-columns:1fr}.wizard-rail{padding:14px}.wizard-main{padding:22px}.summary-grid{grid-template-columns:1fr}}
  `;

  const slideLabels = getSlideLabels(generatedData?.framework || form.framework);

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
            <div className="wizard-shell">
              <div className="wizard-rail">
                <div className="wizard-rail-title">Workflow</div>
                {wizardSteps.map((step, index) => (
                  <div key={step.id} className={`wizard-step ${currentStep === index ? "active" : ""} ${currentStep > index ? "done" : ""}`}>
                    <div className="wizard-num">{step.number}</div>
                    <div className="wizard-name">{step.title}</div>
                    <div className="wizard-sub">{step.subtitle}</div>
                  </div>
                ))}
              </div>

              <div className="wizard-main">
                <div className="wizard-head">
                  <div>
                    <div className="wizard-kicker">Step {wizardSteps[currentStep].number}</div>
                    <div className="wizard-title">{wizardSteps[currentStep].title}</div>
                    <div className="wizard-copy">{wizardSteps[currentStep].subtitle}</div>
                  </div>
                  <div className="tag">{currentStep + 1} / {wizardSteps.length}</div>
                </div>

                <div className="wizard-panel">
                  {currentStep === 0 && (
                    <>
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
                      <div className="grid-2">
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
                      </div>
                      <div className="field">
                        <label>Competency</label>
                        <textarea
                          placeholder="e.g. Examine linguistic features as tools to achieve organizational efficiency in informational texts."
                          value={form.competency}
                          onChange={(e) => set("competency", e.target.value)}
                        />
                      </div>
                      <div className="field">
                        <label>Instructional Framework</label>
                        <div className="choice-grid">
                          {Object.entries(FRAMEWORK_OPTIONS).map(([key, option]) => (
                            <button
                              key={key}
                              type="button"
                              className={`choice-btn ${form.framework === key ? "active" : ""}`}
                              onClick={() => set("framework", key)}
                            >
                              <div className="choice-title">{option.label}</div>
                              <div className="choice-text">{option.description}</div>
                            </button>
                          ))}
                        </div>
                      </div>
                    </>
                  )}

                  {currentStep === 1 && (
                    <>
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
                      <div className="field">
                        <label>Optional Visuals</label>
                        <div className="upload-grid">
                          {[
                            ["title", "Title slide image"],
                            ["activity", "Activity slide image"],
                            ["concept", "Concept/definition image"],
                            ["closing", "Closing slide image"],
                          ].map(([slot, label]) => (
                            <div key={slot} className="upload-card">
                              <div className="upload-head">
                                <div className="upload-title">{label}</div>
                                <div className="tag">{form.images[slot] ? "Added" : "Placeholder"}</div>
                              </div>
                              <div className="upload-note">Upload an image for this slot or leave it blank and the slide will show a themed placeholder frame.</div>
                              <input type="file" accept="image/*" onChange={(e) => handleImageChange(slot, e)} />
                              <div className="upload-preview">
                                {form.images[slot] ? (
                                  <img src={form.images[slot]} alt={label} />
                                ) : (
                                  <div className="upload-placeholder">Placeholder</div>
                                )}
                              </div>
                            </div>
                          ))}
                        </div>
                      </div>
                    </>
                  )}

                  {currentStep === 2 && (
                    <>
                      <div className="field">
                        <label>Content Density</label>
                        <div className="choice-grid">
                          {Object.entries(DETAIL_OPTIONS).map(([key, option]) => (
                            <button
                              key={key}
                              type="button"
                              className={`choice-btn ${form.detailLevel === key ? "active" : ""}`}
                              onClick={() => set("detailLevel", key)}
                            >
                              <div className="choice-title">{option.label}</div>
                              <div className="choice-text">{option.description}</div>
                            </button>
                          ))}
                        </div>
                      </div>
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
                      <div className="field">
                        <label>Any extra notes for the generator? (optional)</label>
                        <textarea
                          placeholder="e.g. Focus on examples from environmental issues. Make it fun and interactive. Include Filipino sample sentences."
                          value={form.extraContext}
                          onChange={(e) => set("extraContext", e.target.value)}
                        />
                      </div>
                    </>
                  )}

                  {currentStep === 3 && (
                    <div className="summary-grid">
                      <div className="summary-card">
                        <div className="summary-label">Lesson</div>
                        <div className="summary-value">{form.topic || "No topic yet"}</div>
                      </div>
                      <div className="summary-card">
                        <div className="summary-label">Class Setup</div>
                        <div className="summary-value">{form.subject} • Grade {form.gradeLevel} • {form.quarter}</div>
                      </div>
                      <div className="summary-card">
                        <div className="summary-label">Template</div>
                        <div className="summary-value">{TEMPLATE_OPTIONS[form.template].label} • {PALETTES[form.palette].label}</div>
                      </div>
                      <div className="summary-card">
                        <div className="summary-label">Framework</div>
                        <div className="summary-value">{FRAMEWORK_OPTIONS[form.framework].label} • {FRAMEWORK_OPTIONS[form.framework].description}</div>
                      </div>
                      <div className="summary-card">
                        <div className="summary-label">Density</div>
                        <div className="summary-value">{DETAIL_OPTIONS[form.detailLevel].label}</div>
                      </div>
                      <div className="summary-card">
                        <div className="summary-label">Competency</div>
                        <div className="summary-value">{form.competency || "Generator will create a default competency note."}</div>
                      </div>
                      <div className="summary-card">
                        <div className="summary-label">Objectives</div>
                        <div className="summary-value">
                          {form.objKnowledge || "Knowledge auto"}
                          <br />
                          {form.objSkills || "Skills auto"}
                          <br />
                          {form.objAttitude || "Attitude auto"}
                        </div>
                      </div>
                      <div className="summary-card">
                        <div className="summary-label">Extra Context</div>
                        <div className="summary-value">{form.extraContext || "No extra notes."}</div>
                      </div>
                      <div className="summary-card">
                        <div className="summary-label">Visuals</div>
                        <div className="summary-value">{Object.values(form.images).filter(Boolean).length} uploaded image slot(s)</div>
                      </div>
                    </div>
                  )}
                </div>

                {status === "error" && (
                  <div className="error-box" style={{ marginBottom: 16 }}>
                    <div className="error-title">Oops! Something went wrong</div>
                    <div className="error-msg">{errorMsg}</div>
                  </div>
                )}

                <div className="wizard-nav">
                  <button type="button" className="btn-secondary" onClick={prevStep} disabled={currentStep === 0}>
                    Back
                  </button>
                  {currentStep < wizardSteps.length - 1 ? (
                    <button type="button" className="btn-generate" style={{ width: "auto", minWidth: 190 }} onClick={nextStep}>
                      Continue
                    </button>
                  ) : (
                    <button
                      className="btn-generate"
                      style={{ width: "auto", minWidth: 240 }}
                      onClick={handleGenerate}
                      disabled={!form.topic.trim()}
                    >
                      Generate My Presentation
                    </button>
                  )}
                </div>
              </div>
            </div>
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
                <strong>{generatedData?.topic}</strong> - {generatedData?.frameworkLabel} - Grade {generatedData?.gradeLevel} {generatedData?.subject} - {generatedData?.themeLabel} palette - {slideLabels.length} Slides
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
                <button className="btn-new" onClick={() => { setStatus("idle"); setGeneratedData(null); setCurrentStep(0); }}>
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
