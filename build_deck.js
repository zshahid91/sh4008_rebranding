/**
 * build_deck.js
 *
 * Generalized PptxGenJS renderer for the n8n "extract design system, then
 * restyle a target deck" automation. Takes two JSON files as input:
 *
 *   1. design_spec.json   - colours, typography, persistent elements, and
 *                           measurements for each of the 11 layout types
 *                           (Type A - Type K), as produced by the Claude
 *                           "design extraction" step.
 *   2. slide_content.json - one entry per output slide: which layout type
 *                           to use and the actual text/image content for
 *                           that slide, as produced by the Claude
 *                           "content mapping" step.
 *
 * Usage:
 *   node build_deck.js <design_spec.json> <slide_content.json> <output.pptx> [assetsDir]
 *
 * assetsDir (optional) is the folder containing extracted images
 * (logo, embedded pictures) referenced by relative path in the JSON.
 *
 * Known pptxgenjs quirk: shape type must be passed as the string 'rect',
 * never pres.ShapeType.rect (that property does not exist at runtime).
 */

const fs = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");

// ----------------------------------------------------------------------
// Defaults: the exact measurements reverse-engineered from the SH4008
// reference deck. Anything missing from design_spec.json's "layouts" or
// "persistentElements" falls back to these, so the script still produces
// a sensible deck even from a partial extraction of a *different*
// reference deck.
// ----------------------------------------------------------------------

const DEFAULTS = {
  canvas: { width: 10, height: 7.5 },
  colors: {
    steelBlue: "4472C4",
    paleBlue: "DCE6F1",
    navy: "1F3864",
    white: "FFFFFF",
    black: "000000",
    green: "70AD47",
    paleGreen: "E2EFDA",
    orange: "C55A11",
    amber: "FFC000",
    grey: "595959",
  },
  typography: {
    fontFace: "Calibri",
  },
  persistentElements: {
    breadcrumb: { x: 0.20, y: 0.07, w: 7.80, h: 0.20, size: 9, color: "595959" },
    logo: { x: 8.35, y: 0.05, w: 1.55, h: 0.52 },
    sectionTitle: { x: 0.20, y: 0.35, w: 8.00, h: 0.55, size: 28, color: "000000" },
    subtitle: { x: 0.20, y: 0.92, w: 8.00, h: 0.30, size: 14, color: "4472C4" },
    footerBar: { x: 0, y: 7.22, w: 10, h: 0.28, color: "1F3864" },
    footerText: { x: 0.15, y: 7.22, w: 9.70, h: 0.28, size: 9, color: "FFFFFF" },
  },
  layouts: {
    A: {
      cards: [
        { x: 0.20, y: 1.35, w: 4.60, h: 1.90 },
        { x: 5.10, y: 1.35, w: 4.70, h: 1.90 },
      ],
      headerH: 0.38,
    },
    B: {
      moduleLabel: { x: 0.20, y: 1.00, w: 8.00, h: 0.40, size: 16 },
      headline: { x: 0.20, y: 1.45, w: 9.00, h: 2.00, size: 56 },
      panel: { x: 0, y: 4.75, w: 10, h: 2.47, color: "1F3864" },
      tagline: { x: 0.5, y: 5.05, w: 9, h: 0.6, size: 20 },
      lecturer: { x: 0.5, y: 5.85, w: 9, h: 0.5, size: 13 },
    },
    C: {
      origin: { x: 0.20, y: 1.42 },
      cardW: 4.70,
      cardH: 1.55,
      gapX: 0.20,
      gapY: 0.17,
      columns: 2,
      numberColW: 0.55,
    },
    D: { image: { x: 0.20, y: 1.35, w: 9.60, h: 5.60 } },
    E: {
      startPoint: { x: 0.20, y: 1.35, w: 2.30, h: 0.35 },
      context: { x: 2.65, y: 1.35, w: 7.15, h: 0.40 },
      origin: { x: 0.20, y: 1.85 },
      cardW: 2.30,
      cardH: 1.20,
      gapX: 0.12,
      gapY: 0.12,
      columns: 4,
      numberColW: 0.42,
      finalCardCentered: true,
    },
    F: {
      startPoint: { x: 0.20, y: 1.35, w: 2.30, h: 0.35 },
      context: { x: 2.65, y: 1.35, w: 7.15, h: 0.55 },
      stagePill: { x: 0.20, y: 2.05, w: 2.30, h: 0.35 },
      origin: { x: 0.20, y: 2.55 },
      cardW: 4.70,
      cardH: 1.35,
      gapX: 0.20,
      gapY: 0.15,
      columns: 2,
      headerH: 0.32,
      annotation: { x: 0.20, y: 6.80, w: 9.60, h: 0.35, size: 11 },
    },
    G: {
      x: 0.20,
      y: 1.42,
      w: 9.70,
      headerH: 0.40,
      rowH: 0.55,
    },
    H: {
      columns: [
        { x: 0.20, y: 1.35 },
        { x: 3.47, y: 1.35 },
        { x: 6.74, y: 1.35 },
      ],
      colW: 3.10,
      headerH: 0.40,
      subtitleH: 0.28,
      bodyH: 4.65,
    },
    I: {
      citationBar: { x: 0.20, y: 1.35, w: 9.60, h: 0.45 },
      body: { x: 0.20, y: 1.95, w: 9.60, h: 3.45 },
      pills: { y: 5.85, h: 0.38, w: 2.90, gapX: 0.15, x0: 0.20 },
    },
    J: {
      image: { x: 0.20, y: 1.35, w: 9.60, h: 5.60 },
      links: { x: 0.20, y: 1.42, w: 9.60, h: 5.50, size: 14 },
    },
    K: {
      headline: { x: 0.20, y: 1.35, w: 9.60, h: 0.70, size: 40 },
      infoBoxes: { y: 2.20, h: 0.70, w: 4.70, gapX: 0.20, x0: 0.20 },
      panel: { x: 0, y: 3.30, w: 10, h: 3.92, color: "1F3864" },
      origin: { x: 0.40, y: 3.65 },
      cardW: 4.40,
      cardH: 1.55,
      gapX: 0.20,
      gapY: 0.15,
      columns: 2,
    },
  },
};

// ----------------------------------------------------------------------
// Helpers
// ----------------------------------------------------------------------

function deepMerge(base, override) {
  if (!override) return base;
  const out = Array.isArray(base) ? base.slice() : { ...base };
  for (const key of Object.keys(override)) {
    const bv = base ? base[key] : undefined;
    const ov = override[key];
    if (ov && typeof ov === "object" && !Array.isArray(ov) && bv && typeof bv === "object" && !Array.isArray(bv)) {
      out[key] = deepMerge(bv, ov);
    } else {
      out[key] = ov;
    }
  }
  return out;
}

function loadSpec(specPath) {
  const raw = JSON.parse(fs.readFileSync(specPath, "utf8"));
  const merged = deepMerge(DEFAULTS, raw);
  return merged;
}

function resolveAsset(assetsDir, ref) {
  if (!ref) return null;
  if (ref.startsWith("data:") || ref.startsWith("image/")) return { data: ref };
  if (/^https?:\/\//.test(ref)) return { path: ref };
  return { path: path.join(assetsDir || ".", ref) };
}

function grid(origin, cardW, cardH, gapX, gapY, columns, count) {
  const positions = [];
  for (let i = 0; i < count; i++) {
    const col = i % columns;
    const row = Math.floor(i / columns);
    positions.push({
      x: origin.x + col * (cardW + gapX),
      y: origin.y + row * (cardH + gapY),
    });
  }
  return positions;
}

// ----------------------------------------------------------------------
// Persistent structural elements (every slide)
// ----------------------------------------------------------------------

function addPersistentElements(slide, spec, slideData, assetsDir) {
  const pe = spec.persistentElements;
  const font = spec.typography.fontFace || "Calibri";

  if (pe.breadcrumb && (slideData.breadcrumb || pe.breadcrumb.text)) {
    slide.addText(slideData.breadcrumb || pe.breadcrumb.text || "", {
      x: pe.breadcrumb.x, y: pe.breadcrumb.y, w: pe.breadcrumb.w, h: pe.breadcrumb.h,
      fontSize: pe.breadcrumb.size, italic: true, color: pe.breadcrumb.color,
      fontFace: font, margin: 0,
    });
  }

  if (pe.logo) {
    const logoRef = slideData.logo || pe.logo.path || pe.logo.data;
    const asset = pe.logo.data ? { data: pe.logo.data } : resolveAsset(assetsDir, logoRef);
    if (asset) {
      slide.addImage({ ...asset, x: pe.logo.x, y: pe.logo.y, w: pe.logo.w, h: pe.logo.h });
    }
  }

  if (slideData.sectionTitle) {
    slide.addText(slideData.sectionTitle, {
      x: pe.sectionTitle.x, y: pe.sectionTitle.y, w: pe.sectionTitle.w, h: pe.sectionTitle.h,
      fontSize: pe.sectionTitle.size, bold: true, color: pe.sectionTitle.color,
      fontFace: font, margin: 0,
    });
  }

  if (slideData.subtitle) {
    slide.addText(slideData.subtitle, {
      x: pe.subtitle.x, y: pe.subtitle.y, w: pe.subtitle.w, h: pe.subtitle.h,
      fontSize: pe.subtitle.size, italic: true, color: pe.subtitle.color,
      fontFace: font, margin: 0,
    });
  }

  if (pe.footerBar) {
    slide.addShape("rect", {
      x: pe.footerBar.x, y: pe.footerBar.y, w: pe.footerBar.w, h: pe.footerBar.h,
      fill: { color: pe.footerBar.color }, line: { color: pe.footerBar.color },
    });
  }

  if (pe.footerText && (slideData.footerText || pe.footerText.text)) {
    slide.addText(slideData.footerText || pe.footerText.text || "", {
      x: pe.footerText.x, y: pe.footerText.y, w: pe.footerText.w, h: pe.footerText.h,
      fontSize: pe.footerText.size, color: pe.footerText.color,
      fontFace: font, margin: 0,
    });
  }
}

// ----------------------------------------------------------------------
// Shared card primitives
// ----------------------------------------------------------------------

function headerCard(slide, spec, font, x, y, w, h, headerH, header, body, colors) {
  const c = colors || spec.colors;
  slide.addShape("rect", { x, y, w, h, fill: { color: c.paleBlue }, line: { color: c.paleBlue } });
  slide.addShape("rect", { x, y, w, h: headerH, fill: { color: c.steelBlue }, line: { color: c.steelBlue } });
  slide.addText(header || "", {
    x: x + 0.10, y, w: w - 0.20, h: headerH,
    fontSize: 12, bold: true, color: c.white, fontFace: font, valign: "middle", margin: 0,
  });
  slide.addText(body || "", {
    x: x + 0.15, y: y + headerH + 0.08, w: w - 0.30, h: h - headerH - 0.16,
    fontSize: 12, color: c.black, fontFace: font, valign: "top", margin: 0,
  });
}

function numberCard(slide, spec, font, x, y, w, h, numColW, number, body, colors) {
  const c = colors || spec.colors;
  slide.addShape("rect", { x, y, w, h, fill: { color: c.paleBlue }, line: { color: c.paleBlue } });
  slide.addShape("rect", { x, y, w: numColW, h, fill: { color: c.steelBlue }, line: { color: c.steelBlue } });
  slide.addText(String(number != null ? number : ""), {
    x, y, w: numColW, h,
    fontSize: 16, bold: true, color: c.white, fontFace: font, align: "center", valign: "middle", margin: 0,
  });
  slide.addText(body || "", {
    x: x + numColW + 0.12, y, w: w - numColW - 0.24, h,
    fontSize: 12, color: c.black, fontFace: font, valign: "middle", margin: 0,
  });
}

function pill(slide, spec, font, x, y, w, h, text, color, textColor) {
  slide.addShape("rect", { x, y, w, h, fill: { color }, line: { color } });
  slide.addText(text || "", {
    x, y, w, h, fontSize: 11, bold: true, color: textColor || "FFFFFF",
    fontFace: font, align: "center", valign: "middle", margin: 0,
  });
}

// ----------------------------------------------------------------------
// Layout type renderers
// ----------------------------------------------------------------------

function renderA(slide, spec, font, L, content) {
  const cards = (content.cards || []).slice(0, L.cards.length || 2);
  cards.forEach((card, i) => {
    const pos = L.cards[i];
    if (!pos) return;
    headerCard(slide, spec, font, pos.x, pos.y, pos.w, pos.h, L.headerH, card.header, card.body);
  });
}

function renderB(slide, spec, font, L, content) {
  const c = spec.colors;
  if (content.moduleLabel) {
    slide.addText(content.moduleLabel, {
      ...L.moduleLabel, fontSize: L.moduleLabel.size, color: c.steelBlue, fontFace: font, margin: 0,
    });
  }
  slide.addText(content.headline || "", {
    ...L.headline, fontSize: L.headline.size, bold: true, color: c.black, fontFace: font, margin: 0,
  });
  slide.addShape("rect", { ...L.panel, fill: { color: L.panel.color }, line: { color: L.panel.color } });
  if (content.tagline) {
    slide.addText(content.tagline, {
      ...L.tagline, fontSize: L.tagline.size, italic: true, color: c.steelBlue, fontFace: font, margin: 0,
    });
  }
  if (content.lecturer) {
    slide.addText(content.lecturer, {
      ...L.lecturer, fontSize: L.lecturer.size, bold: true, color: c.white, fontFace: font, margin: 0,
    });
  }
}

function renderC(slide, spec, font, L, content) {
  const cards = content.cards || [];
  const positions = grid(L.origin, L.cardW, L.cardH, L.gapX, L.gapY, L.columns, cards.length);
  cards.forEach((card, i) => {
    const pos = positions[i];
    numberCard(slide, spec, font, pos.x, pos.y, L.cardW, L.cardH, L.numberColW, card.number != null ? card.number : i + 1, card.text);
  });
}

function renderD(slide, spec, font, L, content, assetsDir) {
  const asset = resolveAsset(assetsDir, content.image);
  if (asset) slide.addImage({ ...asset, ...L.image });
}

function renderE(slide, spec, font, L, content) {
  const c = spec.colors;
  if (content.startPoint) {
    pill(slide, spec, font, L.startPoint.x, L.startPoint.y, L.startPoint.w, L.startPoint.h, content.startPoint.label || "START POINT", c.green);
  }
  if (content.startPoint && content.startPoint.context) {
    slide.addShape("rect", { ...L.context, fill: { color: c.paleGreen }, line: { color: c.paleGreen } });
    slide.addText(content.startPoint.context, {
      x: L.context.x + 0.10, y: L.context.y, w: L.context.w - 0.20, h: L.context.h,
      fontSize: 12, color: c.black, fontFace: font, valign: "middle", margin: 0,
    });
  }
  const steps = content.steps || [];
  const mainSteps = L.finalCardCentered ? steps.slice(0, -1) : steps;
  const lastStep = L.finalCardCentered ? steps[steps.length - 1] : null;
  const positions = grid(L.origin, L.cardW, L.cardH, L.gapX, L.gapY, L.columns, mainSteps.length);
  mainSteps.forEach((step, i) => {
    const pos = positions[i];
    numberCard(slide, spec, font, pos.x, pos.y, L.cardW, L.cardH, L.numberColW, step.number != null ? step.number : i + 1, step.text);
  });
  if (lastStep) {
    const rows = Math.ceil(mainSteps.length / L.columns);
    const lastY = L.origin.y + rows * (L.cardH + L.gapY);
    const lastX = L.origin.x + ((L.columns * (L.cardW + L.gapX) - L.gapX) - L.cardW) / 2;
    numberCard(slide, spec, font, lastX, lastY, L.cardW, L.cardH, L.numberColW, lastStep.number != null ? lastStep.number : steps.length, lastStep.text);
  }
}

function renderF(slide, spec, font, L, content) {
  const c = spec.colors;
  if (content.startPoint) {
    pill(slide, spec, font, L.startPoint.x, L.startPoint.y, L.startPoint.w, L.startPoint.h, content.startPoint.label || "START POINT", c.steelBlue);
  }
  if (content.startPoint && content.startPoint.context) {
    slide.addShape("rect", { ...L.context, fill: { color: c.paleBlue }, line: { color: c.paleBlue } });
    slide.addText(content.startPoint.context, {
      x: L.context.x + 0.10, y: L.context.y, w: L.context.w - 0.20, h: L.context.h,
      fontSize: 12, color: c.black, fontFace: font, valign: "middle", margin: 0,
    });
  }
  if (content.stageLabel) {
    pill(slide, spec, font, L.stagePill.x, L.stagePill.y, L.stagePill.w, L.stagePill.h, content.stageLabel, c.steelBlue);
  }
  const cards = content.cards || [];
  const positions = grid(L.origin, L.cardW, L.cardH, L.gapX, L.gapY, L.columns, cards.length);
  cards.forEach((card, i) => {
    const pos = positions[i];
    headerCard(slide, spec, font, pos.x, pos.y, L.cardW, L.cardH, L.headerH, card.header, card.body);
  });
  if (content.annotation) {
    slide.addText(content.annotation, {
      ...L.annotation, fontSize: L.annotation.size, italic: true, color: c.navy, fontFace: font, margin: 0,
    });
  }
}

function renderG(slide, spec, font, L, content) {
  const c = spec.colors;
  const headers = content.headers || [];
  const rows = content.rows || [];
  const colW = L.w / Math.max(headers.length, 1);
  headers.forEach((h, i) => {
    slide.addShape("rect", { x: L.x + i * colW, y: L.y, w: colW, h: L.headerH, fill: { color: c.steelBlue }, line: { color: c.steelBlue } });
    slide.addText(h, {
      x: L.x + i * colW + 0.05, y: L.y, w: colW - 0.10, h: L.headerH,
      fontSize: 12, bold: true, color: c.white, fontFace: font, align: "center", valign: "middle", margin: 0,
    });
  });
  rows.forEach((row, r) => {
    const rowY = L.y + L.headerH + r * L.rowH;
    const fill = r % 2 === 0 ? c.paleBlue : c.white;
    row.forEach((cell, i) => {
      slide.addShape("rect", { x: L.x + i * colW, y: rowY, w: colW, h: L.rowH, fill: { color: fill }, line: { color: fill } });
      slide.addText(cell, {
        x: L.x + i * colW + 0.08, y: rowY, w: colW - 0.16, h: L.rowH,
        fontSize: 11, color: c.black, fontFace: font, valign: "middle", margin: 0,
      });
    });
  });
}

function renderH(slide, spec, font, L, content) {
  const c = spec.colors;
  const columns = content.columns || [];
  columns.slice(0, L.columns.length).forEach((col, i) => {
    const pos = L.columns[i];
    slide.addShape("rect", { x: pos.x, y: pos.y, w: L.colW, h: L.headerH, fill: { color: c.steelBlue }, line: { color: c.steelBlue } });
    slide.addText(col.header || "", {
      x: pos.x + 0.05, y: pos.y, w: L.colW - 0.10, h: L.headerH,
      fontSize: 12, bold: true, color: c.white, fontFace: font, align: "center", valign: "middle", margin: 0,
    });
    const subY = pos.y + L.headerH;
    slide.addShape("rect", { x: pos.x, y: subY, w: L.colW, h: L.subtitleH, fill: { color: c.paleBlue }, line: { color: c.paleBlue } });
    slide.addText(col.subtitle || "", {
      x: pos.x + 0.05, y: subY, w: L.colW - 0.10, h: L.subtitleH,
      fontSize: 10, italic: true, color: c.black, fontFace: font, align: "center", valign: "middle", margin: 0,
    });
    const bodyY = subY + L.subtitleH;
    slide.addShape("rect", { x: pos.x, y: bodyY, w: L.colW, h: L.bodyH, fill: { color: c.white }, line: { color: c.paleBlue } });
    slide.addText(col.body || "", {
      x: pos.x + 0.12, y: bodyY + 0.10, w: L.colW - 0.24, h: L.bodyH - 0.20,
      fontSize: 11, color: c.black, fontFace: font, valign: "top", margin: 0,
    });
  });
}

function renderI(slide, spec, font, L, content) {
  const c = spec.colors;
  if (content.citationHeader) {
    slide.addShape("rect", { ...L.citationBar, fill: { color: c.paleBlue }, line: { color: c.paleBlue } });
    slide.addText(content.citationHeader, {
      x: L.citationBar.x + 0.10, y: L.citationBar.y, w: L.citationBar.w - 0.20, h: L.citationBar.h,
      fontSize: 12, italic: true, color: c.black, fontFace: font, valign: "middle", margin: 0,
    });
  }
  slide.addText(content.body || "", {
    ...L.body, fontSize: 13, color: c.black, fontFace: font, valign: "top", margin: 0,
  });
  const pills = content.pills || [];
  pills.forEach((p, i) => {
    const x = L.pills.x0 + i * (L.pills.w + L.pills.gapX);
    pill(slide, spec, font, x, L.pills.y, L.pills.w, L.pills.h, p.label, p.color || c.steelBlue);
  });
}

function renderJ(slide, spec, font, L, content, assetsDir) {
  const c = spec.colors;
  if (content.image) {
    const asset = resolveAsset(assetsDir, content.image);
    if (asset) slide.addImage({ ...asset, ...L.image });
    return;
  }
  const links = content.links || [];
  slide.addText(
    links.map((l, i) => ({ text: l, options: { bullet: { type: "number" }, breakLine: i < links.length - 1 } })),
    { ...L.links, fontSize: L.links.size, color: c.black, fontFace: font, valign: "top", margin: 0 }
  );
}

function renderK(slide, spec, font, L, content) {
  const c = spec.colors;
  slide.addText(content.headline || "", {
    ...L.headline, fontSize: L.headline.size, bold: true, color: c.black, fontFace: font, margin: 0,
  });
  const boxes = content.infoBoxes || [];
  boxes.forEach((box, i) => {
    const x = L.infoBoxes.x0 + i * (L.infoBoxes.w + L.infoBoxes.gapX);
    slide.addShape("rect", { x, y: L.infoBoxes.y, w: L.infoBoxes.w, h: L.infoBoxes.h, fill: { color: c.white }, line: { color: c.steelBlue, width: 1 } });
    slide.addText(box, {
      x: x + 0.10, y: L.infoBoxes.y, w: L.infoBoxes.w - 0.20, h: L.infoBoxes.h,
      fontSize: 12, color: c.black, fontFace: font, valign: "middle", margin: 0,
    });
  });
  slide.addShape("rect", { ...L.panel, fill: { color: L.panel.color }, line: { color: L.panel.color } });
  const cards = content.cards || [];
  const positions = grid(L.origin, L.cardW, L.cardH, L.gapX, L.gapY, L.columns, cards.length);
  cards.forEach((card, i) => {
    const pos = positions[i];
    numberCard(slide, spec, font, pos.x, pos.y, L.cardW, L.cardH, 0.45, card.number != null ? card.number : i + 1, card.text, { ...c, paleBlue: "FFFFFF" });
  });
}

const RENDERERS = {
  A: renderA, B: renderB, C: renderC, D: renderD, E: renderE,
  F: renderF, G: renderG, H: renderH, I: renderI, J: renderJ, K: renderK,
};

// ----------------------------------------------------------------------
// Main
// ----------------------------------------------------------------------

function main() {
  const [specPath, contentPath, outPath, assetsDirArg] = process.argv.slice(2);
  if (!specPath || !contentPath || !outPath) {
    console.error("Usage: node build_deck.js <design_spec.json> <slide_content.json> <output.pptx> [assetsDir]");
    process.exit(1);
  }
  const assetsDir = assetsDirArg || path.dirname(contentPath);
  const spec = loadSpec(specPath);
  const slides = JSON.parse(fs.readFileSync(contentPath, "utf8"));
  const font = spec.typography.fontFace || "Calibri";

  const pres = new pptxgen();
  pres.defineLayout({ name: "CUSTOM", width: spec.canvas.width, height: spec.canvas.height });
  pres.layout = "CUSTOM";

  const list = Array.isArray(slides) ? slides : slides.slides;

  list.forEach((slideData, idx) => {
    const slide = pres.addSlide();
    const layoutType = slideData.layoutType;
    const renderer = RENDERERS[layoutType];
    if (!renderer) {
      console.error(`Slide ${idx + 1}: unknown layoutType "${layoutType}", skipping content render (persistent elements only).`);
    }
    addPersistentElements(slide, spec, slideData, assetsDir);
    if (renderer) {
      const L = (spec.layouts && spec.layouts[layoutType]) || {};
      renderer(slide, spec, font, L, slideData.content || {}, assetsDir);
    }
    if (slideData.notes) slide.addNotes(slideData.notes);
  });

  pres
    .writeFile({ fileName: outPath })
    .then(() => console.log(`SAVED OK: ${outPath} (${list.length} slides)`))
    .catch((e) => {
      console.error("ERROR:", e);
      process.exit(1);
    });
}

main();
