/**
 * build_deck.js
 *
 * Generalized PptxGenJS renderer for the n8n "extract design system, then
 * restyle a target deck" automation.
 *
 * The previous version of this file hardcoded one specific reference deck's
 * layout vocabulary (cards, headerH, moduleLabel, panel...). That only
 * worked for decks shaped exactly like the deck it was reverse-engineered
 * from. This version is a generic ELEMENT interpreter instead: every layout
 * type is just a list of positioned elements with a "kind", and content
 * maps element ids to values. Any reference deck's design spec, whatever
 * its own layout types look like, renders through the same five kinds:
 *
 *   - "rect"           decorative fill (panel, header bar, divider). No content.
 *   - "text"           a text box. Content: a string, or an array of
 *                       strings if the element has "bullets": true.
 *   - "image"          a picture placeholder. Content: an asset filename.
 *   - "table"          a simple grid. Content: { headers: [...], rows: [[...]] }.
 *   - "repeatingCards"  a variable-count card grid (numbered steps, question
 *                       cards, etc). Content: an array of { label, text }.
 *
 * Takes two JSON files as input:
 *
 *   1. design_spec.json   - colours, typography, persistent elements, and
 *                           an elements[] array per layout type, as produced
 *                           by the Claude "design extraction" step.
 *   2. slide_content.json - one entry per output slide: which layout type
 *                           to use, and a content map of elementId -> value,
 *                           as produced by the Claude "content mapping" step.
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
// Defaults: only for canvas/colours/typography/persistentElements, used
// when an extraction is partial. There is no default for "layouts" any
// more, every layout's elements always come from the extracted spec,
// since hardcoding one deck's layouts here would defeat the point.
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
  // Only merge the non-layout defaults; layouts always come straight from
  // the extracted spec (see comment on DEFAULTS above).
  const merged = deepMerge(DEFAULTS, { ...raw, layouts: undefined });
  merged.layouts = raw.layouts || {};
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
  const pe = spec.persistentElements || {};
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

  if (slideData.sectionTitle && pe.sectionTitle) {
    slide.addText(slideData.sectionTitle, {
      x: pe.sectionTitle.x, y: pe.sectionTitle.y, w: pe.sectionTitle.w, h: pe.sectionTitle.h,
      fontSize: pe.sectionTitle.size, bold: true, color: pe.sectionTitle.color,
      fontFace: font, margin: 0,
    });
  }

  if (slideData.subtitle && pe.subtitle) {
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
// Generic element renderers, one per "kind"
// ----------------------------------------------------------------------

function renderRect(slide, el, colors) {
  const color = el.color || colors.steelBlue;
  slide.addShape("rect", { x: el.x, y: el.y, w: el.w, h: el.h, fill: { color }, line: { color } });
}

// ----------------------------------------------------------------------
// Text-fit helpers: every text element gets a modest size bump for
// readability, but never beyond what its box can actually hold, so a
// long title shrinks back down automatically instead of overflowing.
// ----------------------------------------------------------------------

const FONT_SIZE_BUMP = 2; // "increase body text font size a bit"
const LINE_SPACING_MULTIPLE = 1.3; // "increase line/paragraph spacing"
const PARA_SPACE_AFTER = 6; // points between bullet items

function estimateLineCount(text, fontSizePt, boxWidthIn) {
  const charWidthIn = (fontSizePt * 0.52) / 72; // rough average glyph width for Calibri-ish fonts
  const charsPerLine = Math.max(1, Math.floor(boxWidthIn / charWidthIn));
  return Math.max(1, Math.ceil(String(text).length / charsPerLine));
}

// Shrinks fontSize down until the (estimated) wrapped text fits the box
// height, never going below ~55% of the requested size so legibility
// doesn't collapse on extreme cases.
function fitFontSize(lines, requestedSize, boxWidthIn, boxHeightIn) {
  const minSize = Math.max(9, Math.round(requestedSize * 0.55));
  let size = requestedSize;
  while (size > minSize) {
    const lineHeightIn = (size * LINE_SPACING_MULTIPLE) / 72;
    let totalLines = 0;
    lines.forEach((line) => { totalLines += estimateLineCount(line, size, boxWidthIn); });
    if (totalLines * lineHeightIn <= boxHeightIn) break;
    size -= 1;
  }
  return size;
}

// Short, label-style lines (all caps, or "A | B | C" style) get bolded
// as mini-headers instead of blending into surrounding paragraphs.
function looksLikeLabel(text) {
  const trimmed = String(text).trim();
  if (!trimmed || trimmed.length > 60) return false;
  return trimmed.includes("|") || trimmed === trimmed.toUpperCase();
}

function renderTextEl(slide, el, value, font) {
  if (value == null || value === "") return;
  const requestedSize = (el.fontSize || 14) + FONT_SIZE_BUMP;

  if (el.bullets && Array.isArray(value)) {
    const fittedSize = fitFontSize(value, requestedSize, el.w, el.h);
    const items = value.map((line, i) => ({
      text: String(line),
      options: {
        bullet: true,
        breakLine: i < value.length - 1,
        bold: !!el.bold || looksLikeLabel(line),
        paraSpaceAfter: PARA_SPACE_AFTER,
      },
    }));
    slide.addText(items, {
      x: el.x, y: el.y, w: el.w, h: el.h,
      fontSize: fittedSize,
      color: el.color || "000000",
      italic: !!el.italic,
      align: el.align || "left",
      valign: el.valign || "top",
      fontFace: font,
      margin: 0,
    });
  } else {
    const text = Array.isArray(value) ? value.join("\n") : String(value);
    const fittedSize = fitFontSize(text.split("\n"), requestedSize, el.w, el.h);
    slide.addText(text, {
      x: el.x, y: el.y, w: el.w, h: el.h,
      fontSize: fittedSize,
      color: el.color || "000000",
      bold: !!el.bold,
      italic: !!el.italic,
      align: el.align || "left",
      valign: el.valign || "top",
      fontFace: font,
      lineSpacingMultiple: LINE_SPACING_MULTIPLE,
      margin: 0,
    });
  }
}

function renderImageEl(slide, el, value, assetsDir) {
  if (!value) return;
  const asset = resolveAsset(assetsDir, value);
  if (asset) slide.addImage({ ...asset, x: el.x, y: el.y, w: el.w, h: el.h });
}

function renderTableEl(slide, el, value, colors, font) {
  if (!value) return;
  const headers = value.headers || [];
  const rows = value.rows || [];
  const headerColor = el.headerColor || colors.steelBlue;
  const headerTextColor = el.headerTextColor || colors.white;
  const rowColor = el.rowColor || colors.paleBlue;
  const altRowColor = el.altRowColor || colors.white;
  const headerHeight = el.headerHeight || 0.4;
  const rowHeight = el.rowHeight || 0.5;
  const fontSize = (el.fontSize || 12) + 1;
  const colW = el.w / Math.max(headers.length, 1);

  headers.forEach((h, i) => {
    slide.addShape("rect", { x: el.x + i * colW, y: el.y, w: colW, h: headerHeight, fill: { color: headerColor }, line: { color: headerColor } });
    slide.addText(String(h), {
      x: el.x + i * colW + 0.05, y: el.y, w: colW - 0.10, h: headerHeight,
      fontSize, bold: true, color: headerTextColor, fontFace: font, align: "center", valign: "middle", margin: 0,
    });
  });

  rows.forEach((row, r) => {
    const rowY = el.y + headerHeight + r * rowHeight;
    const fill = r % 2 === 0 ? rowColor : altRowColor;
    row.forEach((cell, i) => {
      slide.addShape("rect", { x: el.x + i * colW, y: rowY, w: colW, h: rowHeight, fill: { color: fill }, line: { color: fill } });
      slide.addText(String(cell), {
        x: el.x + i * colW + 0.08, y: rowY, w: colW - 0.16, h: rowHeight,
        fontSize: fontSize - 1, color: "000000", fontFace: font, valign: "middle", margin: 0,
      });
    });
  });
}

function renderRepeatingCardsEl(slide, el, value, colors, font) {
  const items = Array.isArray(value) ? value : [];
  if (!items.length) return;

  const style = el.style || "plain";
  const headerColor = el.headerColor || colors.steelBlue;
  const bodyColor = el.bodyColor || colors.paleBlue;
  const textColor = el.textColor || "000000";
  const headerTextColor = el.headerTextColor || colors.white;
  const numberTextColor = el.numberTextColor || colors.white;
  const headerHeight = el.headerHeight || 0.38;
  const numberColWidth = el.numberColWidth || 0.5;

  const mainItems = el.lastItemCentered ? items.slice(0, -1) : items;
  const lastItem = el.lastItemCentered ? items[items.length - 1] : null;

  const positions = grid(el.origin, el.cardW, el.cardH, el.gapX, el.gapY, el.columns, mainItems.length);

  function drawCard(x, y, item) {
    slide.addShape("rect", { x, y, w: el.cardW, h: el.cardH, fill: { color: bodyColor }, line: { color: bodyColor } });
    if (style === "header") {
      slide.addShape("rect", { x, y, w: el.cardW, h: headerHeight, fill: { color: headerColor }, line: { color: headerColor } });
      slide.addText(item.label || "", {
        x: x + 0.10, y, w: el.cardW - 0.20, h: headerHeight,
        fontSize: 13, bold: true, color: headerTextColor, fontFace: font, valign: "middle", margin: 0,
      });
      slide.addText(item.text || "", {
        x: x + 0.15, y: y + headerHeight + 0.08, w: el.cardW - 0.30, h: el.cardH - headerHeight - 0.16,
        fontSize: 13, color: textColor, fontFace: font, valign: "top", lineSpacingMultiple: LINE_SPACING_MULTIPLE, margin: 0,
      });
    } else if (style === "number") {
      slide.addShape("rect", { x, y, w: numberColWidth, h: el.cardH, fill: { color: headerColor }, line: { color: headerColor } });
      slide.addText(String(item.label != null ? item.label : ""), {
        x, y, w: numberColWidth, h: el.cardH,
        fontSize: 17, bold: true, color: numberTextColor, fontFace: font, align: "center", valign: "middle", margin: 0,
      });
      slide.addText(item.text || "", {
        x: x + numberColWidth + 0.12, y, w: el.cardW - numberColWidth - 0.24, h: el.cardH,
        fontSize: 13, color: textColor, fontFace: font, valign: "middle", lineSpacingMultiple: LINE_SPACING_MULTIPLE, margin: 0,
      });
    } else {
      slide.addText(item.text || item.label || "", {
        x: x + 0.12, y: y + 0.08, w: el.cardW - 0.24, h: el.cardH - 0.16,
        fontSize: 13, color: textColor, fontFace: font, valign: "top", lineSpacingMultiple: LINE_SPACING_MULTIPLE, margin: 0,
      });
    }
  }

  mainItems.forEach((item, i) => drawCard(positions[i].x, positions[i].y, item));

  if (lastItem) {
    const rows = Math.ceil(mainItems.length / el.columns);
    const lastY = el.origin.y + rows * (el.cardH + el.gapY);
    const lastX = el.origin.x + ((el.columns * (el.cardW + el.gapX) - el.gapX) - el.cardW) / 2;
    drawCard(lastX, lastY, lastItem);
  }
}

function renderElement(slide, spec, font, assetsDir, el, value) {
  switch (el.kind) {
    case "rect":
      return renderRect(slide, el, spec.colors);
    case "text":
      return renderTextEl(slide, el, value, font);
    case "image":
      return renderImageEl(slide, el, value, assetsDir);
    case "table":
      return renderTableEl(slide, el, value, spec.colors, font);
    case "repeatingCards":
      return renderRepeatingCardsEl(slide, el, value, spec.colors, font);
    default:
      return;
  }
}

function renderLayout(slide, spec, font, assetsDir, layout, content) {
  (layout.elements || []).forEach((el) => {
    const value = content ? content[el.id] : undefined;
    renderElement(slide, spec, font, assetsDir, el, value);
  });
}

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
    const layout = spec.layouts && spec.layouts[layoutType];
    if (!layout) {
      console.error(`Slide ${idx + 1}: unknown layoutType "${layoutType}", skipping content render (persistent elements only).`);
    }
    addPersistentElements(slide, spec, slideData, assetsDir);
    if (layout) {
      renderLayout(slide, spec, font, assetsDir, layout, slideData.content || {});
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
