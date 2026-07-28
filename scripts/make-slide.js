/**
 * Ladder of Less Pain — presentation slide(s)
 * Screenshot + where to find it + Made with Grok Build
 */
const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const ROOT = path.join(__dirname, "..");
const HERO = path.join(ROOT, "assets", "ladder-of-less-pain.png");
const OUT = path.join(ROOT, "ladder-of-less-pain-slide.pptx");

const SITE = "https://fullergalway.github.io/ladder-of-less-pain/";
const REPO = "https://github.com/fullergalway/ladder-of-less-pain";

// Source screenshot 1680×1111
const IMG_AR = 1680 / 1111;

function fitImage(maxW, maxH) {
  let w = maxW;
  let h = w / IMG_AR;
  if (h > maxH) {
    h = maxH;
    w = h * IMG_AR;
  }
  return { w, h };
}

async function main() {
  if (!fs.existsSync(HERO)) throw new Error(`Missing screenshot: ${HERO}`);

  const pres = new pptxgen();
  // Standard 16:9 for decks / projectors
  pres.layout = "LAYOUT_16x9";
  pres.author = "Grok Build";
  pres.title = "Ladder of Less Pain";
  pres.subject = "70 Years of Developer Tooling (1956–2026)";
  pres.company = "xAI Grok Build";

  // ═══════════════════════════════════════════════════════════
  // Slide 1 — hero screenshot
  // ═══════════════════════════════════════════════════════════
  const s1 = pres.addSlide();
  s1.background = { color: "0B1020" };

  // Accent bar
  s1.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.06,
    fill: { color: "A855F7" },
    line: { color: "A855F7", width: 0 },
  });
  s1.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0.06, w: 10, h: 0.03,
    fill: { color: "22D3EE" },
    line: { color: "22D3EE", width: 0 },
  });

  // Title
  s1.addText("Ladder of Less Pain", {
    x: 0.4, y: 0.2, w: 6.8, h: 0.42,
    fontFace: "Arial", fontSize: 26, bold: true, color: "F8FAFC", margin: 0,
  });
  s1.addText("The Abstraction Ladder · 70 Years of Developer Tooling (1956–2026)", {
    x: 0.4, y: 0.58, w: 7.2, h: 0.28,
    fontFace: "Arial", fontSize: 12, color: "94A3B8", margin: 0,
  });

  // Grok Build badge
  s1.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 7.55, y: 0.22, w: 2.1, h: 0.58,
    fill: { color: "161F35" },
    line: { color: "475569", width: 1 },
    rectRadius: 0.08,
  });
  s1.addText([
    { text: "Made with", options: { breakLine: true, fontSize: 9, color: "94A3B8" } },
    { text: "Grok Build", options: { fontSize: 14, bold: true, color: "A5B4FC" } },
  ], {
    x: 7.6, y: 0.26, w: 2.0, h: 0.5,
    fontFace: "Arial", align: "center", valign: "middle", margin: 0,
  });

  // Screenshot
  const { w: imgW, h: imgH } = fitImage(9.2, 3.85);
  const imgX = (10 - imgW) / 2;
  const imgY = 0.98;

  s1.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: imgX - 0.05,
    y: imgY - 0.05,
    w: imgW + 0.1,
    h: imgH + 0.1,
    fill: { color: "1E293B" },
    line: { color: "334155", width: 1 },
    rectRadius: 0.06,
    shadow: { type: "outer", color: "000000", blur: 14, offset: 3, angle: 135, opacity: 0.4 },
  });

  s1.addImage({
    path: HERO,
    x: imgX,
    y: imgY,
    w: imgW,
    h: imgH,
    altText: "Screenshot of the Ladder of Less Pain interactive visualization",
    hyperlink: { url: SITE },
  });

  // Footer
  const fy = 5.0;
  s1.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: fy, w: 10, h: 0.625,
    fill: { color: "121A2E" },
    line: { color: "121A2E", width: 0 },
  });

  s1.addText([
    { text: "Live  ", options: { bold: true, color: "67E8F9" } },
    {
      text: "fullergalway.github.io/ladder-of-less-pain",
      options: { color: "F1F5F9", hyperlink: { url: SITE } },
    },
  ], {
    x: 0.4, y: fy + 0.08, w: 6.2, h: 0.24,
    fontFace: "Arial", fontSize: 12, margin: 0,
  });

  s1.addText([
    { text: "Repo  ", options: { bold: true, color: "C4B5FD" } },
    {
      text: "github.com/fullergalway/ladder-of-less-pain",
      options: { color: "CBD5E1", hyperlink: { url: REPO } },
    },
  ], {
    x: 0.4, y: fy + 0.34, w: 6.2, h: 0.22,
    fontFace: "Arial", fontSize: 11, margin: 0,
  });

  s1.addText("Yesterday’s expert tool is\ntoday’s baseline. The craft remains.", {
    x: 6.6, y: fy + 0.08, w: 3.1, h: 0.48,
    fontFace: "Georgia", fontSize: 10, italic: true, color: "A78BFA",
    align: "right", valign: "middle", margin: 0,
  });

  // ═══════════════════════════════════════════════════════════
  // Slide 2 — talking points + links
  // ═══════════════════════════════════════════════════════════
  const s2 = pres.addSlide();
  s2.background = { color: "0B1020" };

  s2.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.06,
    fill: { color: "A855F7" },
    line: { color: "A855F7", width: 0 },
  });
  s2.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0.06, w: 10, h: 0.03,
    fill: { color: "22D3EE" },
    line: { color: "22D3EE", width: 0 },
  });

  s2.addText("The message for students", {
    x: 0.45, y: 0.28, w: 9, h: 0.4,
    fontFace: "Arial", fontSize: 24, bold: true, color: "F8FAFC", margin: 0,
  });

  s2.addText("Tooling has always been changing. AI is a big change — and it rests on every prior rise of the ladder.", {
    x: 0.45, y: 0.72, w: 9.1, h: 0.4,
    fontFace: "Arial", fontSize: 13, color: "94A3B8", margin: 0,
  });

  const cards = [
    { title: "1956–1969", body: "Batch & languages\nPunch cards → FORTRAN,\nCOBOL, BASIC", accent: "64748B" },
    { title: "1970s–80s", body: "Interactive systems\nUnix, C, PCs,\nTurbo tools", accent: "0EA5E9" },
    { title: "1990s–2000s", body: "Network & collaborate\nWeb, Git, CI,\nopen source", accent: "22C55E" },
    { title: "2010s–2026", body: "Cloud → AI agents\nContainers, Copilot,\nintent as interface", accent: "EC4899" },
  ];

  cards.forEach((c, i) => {
    const x = 0.4 + i * 2.4;
    s2.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x, y: 1.25, w: 2.25, h: 2.15,
      fill: { color: "161F35" },
      line: { color: "243049", width: 1 },
      rectRadius: 0.1,
    });
    s2.addShape(pres.shapes.RECTANGLE, {
      x, y: 1.25, w: 2.25, h: 0.08,
      fill: { color: c.accent },
      line: { color: c.accent, width: 0 },
    });
    s2.addText(c.title, {
      x: x + 0.12, y: 1.48, w: 2.0, h: 0.32,
      fontFace: "Arial", fontSize: 13, bold: true, color: c.accent, margin: 0,
    });
    s2.addText(c.body, {
      x: x + 0.12, y: 1.85, w: 2.0, h: 1.3,
      fontFace: "Arial", fontSize: 12, color: "E2E8F0", margin: 0,
    });
  });

  // Where to find it
  s2.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.4, y: 3.6, w: 9.2, h: 1.2,
    fill: { color: "121A2E" },
    line: { color: "334155", width: 1 },
    rectRadius: 0.1,
  });

  s2.addText("Where to find it", {
    x: 0.6, y: 3.72, w: 4, h: 0.25,
    fontFace: "Arial", fontSize: 11, bold: true, color: "67E8F9", margin: 0,
  });

  s2.addText([
    {
      text: SITE,
      options: { breakLine: true, hyperlink: { url: SITE }, color: "F8FAFC", fontSize: 13 },
    },
    {
      text: REPO,
      options: { hyperlink: { url: REPO }, color: "CBD5E1", fontSize: 12 },
    },
  ], {
    x: 0.6, y: 4.05, w: 5.8, h: 0.55,
    fontFace: "Arial", margin: 0,
  });

  s2.addText([
    { text: "Made with Grok Build", options: { breakLine: true, bold: true, color: "A5B4FC", fontSize: 14 } },
    { text: "Interactive · single HTML file", options: { breakLine: true, color: "94A3B8", fontSize: 11 } },
    { text: "Open source on GitHub", options: { color: "94A3B8", fontSize: 11 } },
  ], {
    x: 6.5, y: 3.85, w: 2.9, h: 0.8,
    fontFace: "Arial", align: "right", margin: 0,
  });

  s2.addText("What never changed: Algorithms · Correctness · Trade-offs · Debugging · Empathy · Systems thinking", {
    x: 0.4, y: 5.0, w: 9.2, h: 0.25,
    fontFace: "Arial", fontSize: 11, color: "64748B", align: "center", margin: 0,
  });

  s2.addText("Yesterday’s expert tool is today’s baseline. The craft remains.", {
    x: 0.4, y: 5.28, w: 9.2, h: 0.25,
    fontFace: "Georgia", fontSize: 12, italic: true, color: "C4B5FD", align: "center", margin: 0,
  });

  await pres.writeFile({ fileName: OUT });
  console.log("Wrote", OUT);
  console.log("Image size on slide 1:", imgW.toFixed(2), "×", imgH.toFixed(2), "in");
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
