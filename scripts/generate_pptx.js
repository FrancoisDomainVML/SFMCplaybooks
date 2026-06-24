/**
 * Playbook PPTX Generator
 *
 * Usage:
 *   node scripts/generate_pptx.js content/playbook-example.yaml
 *   node scripts/generate_pptx.js content/        # regenerate all playbooks
 *
 * Reads a YAML content file and produces a .pptx in output/
 */

const pptxgen = require("pptxgenjs");
const yaml = require("js-yaml");
const fs = require("fs");
const path = require("path");

// ─── Colour Palette ──────────────────────────────────────────────────────────
const C = {
  navy:      "1E2A4A",   // primary dark (title slides, section slides)
  teal:      "0D7377",   // accent colour
  tealLight: "E8F4F4",   // tinted background for cards
  white:     "FFFFFF",
  offWhite:  "F7F8FA",
  slate:     "64748B",   // muted text
  charcoal:  "1E293B",   // body text
  silver:    "CBD5E1",   // dividers / borders
};

// ─── Helpers ─────────────────────────────────────────────────────────────────
function makeShadow() {
  return { type: "outer", color: "000000", blur: 6, offset: 2, angle: 45, opacity: 0.10 };
}

function addFooter(slide, version, lastUpdated) {
  slide.addText(`v${version}  ·  ${lastUpdated}`, {
    x: 0.5, y: 5.25, w: 9, h: 0.25,
    fontSize: 9, color: C.slate, align: "right", fontFace: "Calibri",
  });
}

// ─── Slide renderers ─────────────────────────────────────────────────────────

function renderTitle(pres, slide_data, meta) {
  const slide = pres.addSlide();
  slide.background = { color: C.navy };

  // Large title
  slide.addText(slide_data.title, {
    x: 0.8, y: 1.4, w: 8.4, h: 1.5,
    fontSize: 40, bold: true, color: C.white, fontFace: "Cambria", margin: 0,
  });

  // Teal rule
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.8, y: 3.05, w: 1.2, h: 0.06, fill: { color: C.teal }, line: { color: C.teal },
  });

  // Subtitle
  slide.addText(slide_data.subtitle || "", {
    x: 0.8, y: 3.2, w: 8.4, h: 0.7,
    fontSize: 16, color: "CADCFC", fontFace: "Calibri", margin: 0,
  });

  // Meta block
  const metaText = `Owner: ${meta.owner}  ·  ${meta.contact}  ·  v${meta.version}  ·  ${meta.last_updated}`;
  slide.addText(metaText, {
    x: 0.8, y: 5.0, w: 8.4, h: 0.35,
    fontSize: 10, color: C.slate, fontFace: "Calibri", margin: 0,
  });

  if (slide_data.notes) slide.addNotes(slide_data.notes);
  return slide;
}

function renderSection(pres, slide_data, meta) {
  const slide = pres.addSlide();
  slide.background = { color: C.teal };

  slide.addText(slide_data.heading, {
    x: 0.8, y: 1.8, w: 8.4, h: 1.2,
    fontSize: 36, bold: true, color: C.white, fontFace: "Cambria", margin: 0,
  });

  if (slide_data.subheading) {
    slide.addText(slide_data.subheading, {
      x: 0.8, y: 3.1, w: 8.4, h: 0.6,
      fontSize: 16, color: "D1EEF0", fontFace: "Calibri", margin: 0,
    });
  }

  addFooter(slide, meta.version, meta.last_updated);
  if (slide_data.notes) slide.addNotes(slide_data.notes);
  return slide;
}

function renderContent(pres, slide_data, meta) {
  const slide = pres.addSlide();
  slide.background = { color: C.offWhite };

  // Title bar background
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 1.0, fill: { color: C.navy }, line: { color: C.navy },
  });

  slide.addText(slide_data.title, {
    x: 0.5, y: 0.1, w: 9, h: 0.8,
    fontSize: 22, bold: true, color: C.white, fontFace: "Cambria", valign: "middle", margin: 0,
  });

  // Bullet card
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.5, y: 1.15, w: 9, h: 3.9,
    fill: { color: C.white }, rectRadius: 0.1,
    shadow: makeShadow(), line: { color: C.silver, width: 0.5 },
  });

  const bullets = (slide_data.bullets || []).map((b, i) => ({
    text: b,
    options: { bullet: true, breakLine: i < slide_data.bullets.length - 1 },
  }));

  slide.addText(bullets, {
    x: 0.9, y: 1.35, w: 8.2, h: 3.5,
    fontSize: 14, color: C.charcoal, fontFace: "Calibri",
    valign: "top", paraSpaceAfter: 8, margin: 0,
  });

  addFooter(slide, meta.version, meta.last_updated);
  if (slide_data.notes) slide.addNotes(slide_data.notes);
  return slide;
}

function renderProcess(pres, slide_data, meta) {
  const slide = pres.addSlide();
  slide.background = { color: C.offWhite };

  // Title bar
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 1.0, fill: { color: C.navy }, line: { color: C.navy },
  });
  slide.addText(slide_data.title, {
    x: 0.5, y: 0.1, w: 9, h: 0.8,
    fontSize: 22, bold: true, color: C.white, fontFace: "Cambria", valign: "middle", margin: 0,
  });

  const steps = slide_data.steps || [];
  const n = steps.length;
  const cardW = Math.min(1.7, (9.0 - (n - 1) * 0.25) / n);
  const totalW = n * cardW + (n - 1) * 0.25;
  const startX = (10 - totalW) / 2;

  steps.forEach((step, i) => {
    const x = startX + i * (cardW + 0.25);
    const y = 1.3;

    // Card
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x, y, w: cardW, h: 3.6,
      fill: { color: C.white }, rectRadius: 0.1,
      shadow: makeShadow(), line: { color: C.silver, width: 0.5 },
    });

    // Step number circle
    slide.addShape(pres.shapes.OVAL, {
      x: x + cardW / 2 - 0.28, y: y + 0.2, w: 0.56, h: 0.56,
      fill: { color: C.teal }, line: { color: C.teal },
    });
    slide.addText(String(i + 1), {
      x: x + cardW / 2 - 0.28, y: y + 0.2, w: 0.56, h: 0.56,
      fontSize: 13, bold: true, color: C.white, align: "center", valign: "middle",
      fontFace: "Calibri", margin: 0,
    });

    // Label
    slide.addText(step.label, {
      x: x + 0.1, y: y + 0.9, w: cardW - 0.2, h: 0.65,
      fontSize: 12, bold: true, color: C.navy, align: "center", fontFace: "Cambria",
      margin: 0,
    });

    // Description
    slide.addText(step.description, {
      x: x + 0.1, y: y + 1.65, w: cardW - 0.2, h: 1.7,
      fontSize: 11, color: C.slate, align: "center", fontFace: "Calibri",
      valign: "top", margin: 0,
    });

    // Arrow between cards
    if (i < n - 1) {
      slide.addShape(pres.shapes.LINE, {
        x: x + cardW + 0.04, y: y + 1.55, w: 0.17, h: 0,
        line: { color: C.teal, width: 2 },
      });
    }
  });

  addFooter(slide, meta.version, meta.last_updated);
  if (slide_data.notes) slide.addNotes(slide_data.notes);
  return slide;
}

function renderTable(pres, slide_data, meta) {
  const slide = pres.addSlide();
  slide.background = { color: C.offWhite };

  // Title bar
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 1.0, fill: { color: C.navy }, line: { color: C.navy },
  });
  slide.addText(slide_data.title, {
    x: 0.5, y: 0.1, w: 9, h: 0.8,
    fontSize: 22, bold: true, color: C.white, fontFace: "Cambria", valign: "middle", margin: 0,
  });

  const headers = (slide_data.headers || []).map(h => ({
    text: h,
    options: { bold: true, color: C.white, fill: { color: C.navy }, align: "center" },
  }));

  const rows = [headers, ...(slide_data.rows || []).map((row, ri) =>
    row.map(cell => ({
      text: cell,
      options: { color: C.charcoal, fill: { color: ri % 2 === 0 ? C.white : C.tealLight } },
    }))
  )];

  slide.addTable(rows, {
    x: 0.5, y: 1.2, w: 9, h: 4.0,
    border: { pt: 0.5, color: C.silver },
    fontSize: 13, fontFace: "Calibri",
    rowH: 0.5,
  });

  addFooter(slide, meta.version, meta.last_updated);
  if (slide_data.notes) slide.addNotes(slide_data.notes);
  return slide;
}

function renderContact(pres, slide_data, meta) {
  const slide = pres.addSlide();
  slide.background = { color: C.offWhite };

  // Title bar
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 1.0, fill: { color: C.navy }, line: { color: C.navy },
  });
  slide.addText(slide_data.title, {
    x: 0.5, y: 0.1, w: 9, h: 0.8,
    fontSize: 22, bold: true, color: C.white, fontFace: "Cambria", valign: "middle", margin: 0,
  });

  const contacts = slide_data.contacts || [];
  const cols = 2;
  const cardW = 4.3;
  const cardH = 0.75;
  const gapX = 0.4;
  const gapY = 0.2;
  const startX = 0.5;
  const startY = 1.2;

  contacts.forEach((c, i) => {
    const col = i % cols;
    const row = Math.floor(i / cols);
    const x = startX + col * (cardW + gapX);
    const y = startY + row * (cardH + gapY);

    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x, y, w: cardW, h: cardH,
      fill: { color: C.white }, rectRadius: 0.08,
      shadow: makeShadow(), line: { color: C.silver, width: 0.5 },
    });

    slide.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 0.08, h: cardH,
      fill: { color: C.teal }, line: { color: C.teal },
    });

    slide.addText(c.role, {
      x: x + 0.18, y: y + 0.06, w: cardW - 0.28, h: 0.28,
      fontSize: 11, bold: true, color: C.navy, fontFace: "Cambria", margin: 0,
    });
    slide.addText(c.detail, {
      x: x + 0.18, y: y + 0.38, w: cardW - 0.28, h: 0.28,
      fontSize: 11, color: C.slate, fontFace: "Calibri", margin: 0,
    });
  });

  addFooter(slide, meta.version, meta.last_updated);
  if (slide_data.notes) slide.addNotes(slide_data.notes);
  return slide;
}

// ─── Main ─────────────────────────────────────────────────────────────────────

async function generatePlaybook(yamlPath) {
  const raw = fs.readFileSync(yamlPath, "utf8");
  const data = yaml.load(raw);
  const { meta, slides } = data;

  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.title = meta.title;
  pres.author = meta.owner;

  for (const slide of slides) {
    switch (slide.type) {
      case "title":    renderTitle(pres, slide, meta);   break;
      case "section":  renderSection(pres, slide, meta); break;
      case "content":  renderContent(pres, slide, meta); break;
      case "process":  renderProcess(pres, slide, meta); break;
      case "table":    renderTable(pres, slide, meta);   break;
      case "contact":  renderContact(pres, slide, meta); break;
      default:
        console.warn(`⚠️  Unknown slide type: "${slide.type}" — skipped`);
    }
  }

  const baseName = path.basename(yamlPath, ".yaml");
  const outDir = path.join(path.dirname(yamlPath), "../output");
  fs.mkdirSync(outDir, { recursive: true });
  const outPath = path.join(outDir, `${baseName}.pptx`);

  await pres.writeFile({ fileName: outPath });
  console.log(`✅  Generated: ${outPath}`);
}

async function main() {
  const arg = process.argv[2];
  if (!arg) {
    console.error("Usage: node scripts/generate_pptx.js <file.yaml | content-dir/>");
    process.exit(1);
  }

  const stat = fs.statSync(arg);
  if (stat.isDirectory()) {
    const files = fs.readdirSync(arg).filter(f => f.endsWith(".yaml"));
    if (files.length === 0) { console.log("No .yaml files found."); return; }
    for (const f of files) await generatePlaybook(path.join(arg, f));
  } else {
    await generatePlaybook(arg);
  }
}

main().catch(err => { console.error(err); process.exit(1); });
