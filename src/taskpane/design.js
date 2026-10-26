/* global PowerPoint */
// Draws consulting-style slides on a blank layout: insight headline, a structured
// body (points, columns or process) and a "so what" takeaway bar.
// Sizes are in points for PowerPoint's default 16:9 slide (960 x 540).

const W = 960;
const H = 540;
const M = 40; // side margin
const NAVY = "#051C2C";
const ACCENT = "#2251FF";
const TEXT = "#333F48";
const LIGHT = "#F2F4F7";
const LINE = "#C9D3DD";
const MUTED = "#8C96A0";
const TITLE_FONT = "Georgia";
const BODY_FONT = "Arial";

function box(slide, { left, top, width, height, fill, shape = PowerPoint.GeometricShapeType.rectangle }) {
  const s = slide.shapes.addGeometricShape(shape, { left, top, width, height });
  s.fill.setSolidColor(fill);
  s.lineFormat.visible = false;
  return s;
}

function style(s, { size, color = TEXT, bold = false, font = BODY_FONT, align = "Left", valign = "Top", shrink = false }) {
  const frame = s.textFrame;
  frame.wordWrap = true;
  frame.verticalAlignment = valign;
  if (shrink) frame.autoSizeSetting = "AutoSizeTextToFitShape";
  const range = frame.textRange;
  range.font.name = font;
  range.font.size = size;
  range.font.color = color;
  range.font.bold = bold;
  range.paragraphFormat.horizontalAlignment = align;
  return s;
}

function text(slide, value, area, look) {
  return style(slide.shapes.addTextBox(value, area), look);
}

// ---------- slide types ----------

export function drawTitleSlide(slide, { title, subtitle }) {
  box(slide, { left: 0, top: 0, width: W, height: H, fill: NAVY });
  text(slide, title, { left: M + 20, top: 120, width: 760, height: 130 }, {
    size: 38, color: "#FFFFFF", font: TITLE_FONT, valign: "Bottom", shrink: true,
  });
  box(slide, { left: M + 20, top: 266, width: 64, height: 4, fill: ACCENT });
  if (subtitle) {
    text(slide, subtitle, { left: M + 20, top: 284, width: 700, height: 70 }, { size: 17, color: "#C9D3DD" });
  }
}

export function drawContentSlide(slide, s, page) {
  // Headline and rule
  text(slide, s.title, { left: M, top: 22, width: W - 2 * M, height: 74 }, {
    size: 24, color: NAVY, font: TITLE_FONT, valign: "Bottom", shrink: true,
  });
  box(slide, { left: M, top: 104, width: W - 2 * M, height: 1.5, fill: NAVY });

  // Body
  const bottom = s.takeaway ? 420 : 490;
  const area = { left: M, top: 126, width: W - 2 * M, height: bottom - 126 };
  if (s.type === "columns") drawColumns(slide, s.items, area);
  else if (s.type === "process") drawProcess(slide, s.items, area);
  else drawPoints(slide, s.bullets, area);

  // So-what bar
  if (s.takeaway) {
    const bar = box(slide, { left: M, top: 436, width: W - 2 * M, height: 56, fill: NAVY });
    bar.textFrame.textRange.text = s.takeaway;
    bar.textFrame.leftMargin = 18;
    bar.textFrame.rightMargin = 18;
    style(bar, { size: 15, color: "#FFFFFF", bold: true, valign: "Middle", shrink: true });
  }

  text(slide, String(page), { left: W - M - 40, top: H - 30, width: 40, height: 18 }, {
    size: 9, color: MUTED, align: "Right",
  });
}

function drawPoints(slide, bullets, area) {
  // One row per point: accent marker, text, thin divider.
  const rowHeight = area.height / Math.max(bullets.length, 1);
  bullets.forEach((bullet, i) => {
    const top = area.top + i * rowHeight;
    box(slide, { left: area.left, top: top + rowHeight / 2 - 5, width: 10, height: 10, fill: ACCENT });
    text(slide, bullet, { left: area.left + 26, top, width: area.width - 26, height: rowHeight }, {
      size: 17, valign: "Middle", shrink: true,
    });
    if (i < bullets.length - 1) {
      box(slide, { left: area.left, top: top + rowHeight - 0.5, width: area.width, height: 1, fill: LINE });
    }
  });
}

function drawColumns(slide, items, area) {
  const gap = 16;
  const width = (area.width - gap * (items.length - 1)) / items.length;
  const height = Math.min(area.height, 210); // cards sit centred between headline and takeaway
  const top = area.top + (area.height - height) / 2;
  items.forEach((item, i) => {
    const left = area.left + i * (width + gap);
    box(slide, { left, top, width, height, fill: LIGHT });
    box(slide, { left, top, width, height: 4, fill: ACCENT });
    text(slide, item.heading, { left: left + 14, top: top + 18, width: width - 28, height: 40 }, {
      size: 16, bold: true, color: NAVY, shrink: true,
    });
    text(slide, item.text, { left: left + 14, top: top + 62, width: width - 28, height: height - 76 }, {
      size: 14, shrink: true,
    });
  });
}

function drawProcess(slide, items, area) {
  const gap = 16;
  const width = (area.width - gap * (items.length - 1)) / items.length;
  const d = 38; // step circle diameter
  const top = area.top + (area.height - 190) / 2; // the steps take about 190 points
  // Connector behind the circles, from the first centre to the last.
  box(slide, { left: area.left + width / 2, top: top + d / 2 - 1, width: area.width - width, height: 2, fill: LINE });
  items.forEach((item, i) => {
    const left = area.left + i * (width + gap);
    const circle = box(slide, {
      left: left + width / 2 - d / 2, top, width: d, height: d,
      fill: i === items.length - 1 ? ACCENT : NAVY, shape: PowerPoint.GeometricShapeType.ellipse,
    });
    circle.textFrame.textRange.text = String(i + 1);
    circle.textFrame.leftMargin = 0;
    circle.textFrame.rightMargin = 0;
    style(circle, { size: 15, color: "#FFFFFF", bold: true, align: "Center", valign: "Middle" });
    text(slide, item.heading, { left, top: top + d + 12, width, height: 36 }, {
      size: 16, bold: true, color: NAVY, align: "Center", shrink: true,
    });
    text(slide, item.text, { left: left + 6, top: top + d + 50, width: width - 12, height: 100 }, {
      size: 14, align: "Center", shrink: true,
    });
  });
}
