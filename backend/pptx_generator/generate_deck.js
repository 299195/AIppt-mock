#!/usr/bin/env node
"use strict";

const fs = require("fs");
const path = require("path");
const crypto = require("crypto");
const JSZip = require("jszip");
const PptxGenJS = require("pptxgenjs");
const {
  warnIfSlideHasOverlaps,
  warnIfSlideElementsOutOfBounds,
} = require("./helpers/layout");

function parseArgs(argv) {
  const args = { input: "", output: "" };
  for (let i = 2; i < argv.length; i += 1) {
    const token = argv[i];
    if (token === "--input") {
      args.input = argv[i + 1] || "";
      i += 1;
    } else if (token === "--output") {
      args.output = argv[i + 1] || "";
      i += 1;
    }
  }
  if (!args.input || !args.output) {
    throw new Error("Usage: node generate_deck.js --input <payload.json> --output <out.pptx>");
  }
  return args;
}

function mustReadJson(filePath) {
  const raw = fs.readFileSync(filePath, "utf8").replace(/^\uFEFF/, "");
  return JSON.parse(raw);
}

function toHexColor(rgb) {
  const r = Math.max(0, Math.min(255, Number(rgb[0] || 0)));
  const g = Math.max(0, Math.min(255, Number(rgb[1] || 0)));
  const b = Math.max(0, Math.min(255, Number(rgb[2] || 0)));
  return `${r.toString(16).padStart(2, "0")}${g.toString(16).padStart(2, "0")}${b.toString(16).padStart(2, "0")}`.toUpperCase();
}

function deriveTheme(templateId) {
  const digest = crypto.createHash("md5").update(String(templateId || "default")).digest();
  const pick = (seed, low, span) => low + (digest[seed] % span);

  const bg = [pick(0, 228, 24), pick(1, 233, 20), pick(2, 238, 17)];
  const header = [pick(3, 18, 44), pick(4, 45, 40), pick(5, 70, 35)];
  const text = [pick(6, 24, 40), pick(7, 38, 34), pick(8, 52, 30)];
  const muted = [Math.min(235, text[0] + 28), Math.min(235, text[1] + 24), Math.min(235, text[2] + 22)];
  const line = [pick(9, 188, 30), pick(10, 206, 24), pick(11, 217, 22)];
  const accent = [pick(12, 45, 130), pick(13, 95, 120), pick(14, 120, 115)];

  return {
    bg: toHexColor(bg),
    header: toHexColor(header),
    headerText: "F8FBFF",
    card: "FCFEFF",
    cardAlt: toHexColor([Math.max(225, bg[0] - 10), Math.max(228, bg[1] - 8), Math.max(232, bg[2] - 6)]),
    text: toHexColor(text),
    muted: toHexColor(muted),
    line: toHexColor(line),
    accent: toHexColor(accent),
  };
}

function truncate(text, limit) {
  const value = String(text || "").trim();
  if (value.length <= limit) return value;
  return `${value.slice(0, Math.max(0, limit - 3)).trim()}...`;
}

function normalizeTocItem(raw) {
  let txt = String(raw || "").trim();
  if (!txt) return "";
  txt = txt.replace(/^\u7b2c\s*\d+\s*\u9875[\uff1a:]\s*/u, "");
  txt = txt.replace(/^\d+\s*[\.\u3001\)\uff09]\s*/u, "");
  return txt.trim();
}

function contentSlides(slides) {
  const items = Array.isArray(slides) ? slides : [];
  return items.filter((s) => {
    const title = String((s && s.title) || "").toLowerCase();
    const slideType = String((s && s.slide_type) || "").toLowerCase();
    if (slideType === "title") return false;
    if (title.includes("cover") || title.includes("agenda")) return false;
    return true;
  });
}

function addWarnings(slide, pptx) {
  try {
    warnIfSlideHasOverlaps(slide, pptx);
    warnIfSlideElementsOutOfBounds(slide, pptx);
  } catch (err) {
    console.warn(`[pptx_generator] warning checks failed: ${String(err && err.message ? err.message : err)}`);
  }
}

function addTitleBar(slide, title, theme) {
  slide.addShape("roundRect", {
    x: 0.45,
    y: 0.25,
    w: 12.4,
    h: 0.9,
    radius: 0.08,
    fill: { color: theme.header },
    line: { color: theme.header, pt: 0 },
  });
  slide.addText(truncate(title, 78), {
    x: 0.75,
    y: 0.42,
    w: 11.8,
    h: 0.54,
    fontFace: "Microsoft YaHei",
    bold: true,
    color: theme.headerText,
    fontSize: 24,
    valign: "mid",
  });
}

function addImageIfExists(slide, imagePath, x, y, w, h) {
  if (!imagePath) return;
  if (!fs.existsSync(imagePath)) return;
  slide.addShape("roundRect", {
    x,
    y,
    w,
    h,
    radius: 0.06,
    fill: { color: "F8FBFF" },
    line: { color: "D6E0F0", pt: 1 },
  });
  slide.addImage({ path: imagePath, x: x + 0.08, y: y + 0.08, w: Math.max(0.3, w - 0.16), h: Math.max(0.3, h - 0.16) });
}

function addBulletList(slide, bullets, x, y, w, h, theme) {
  const clean = (Array.isArray(bullets) ? bullets : []).map((item) => truncate(item, 90)).filter(Boolean);
  const items = clean.length > 0 ? clean.slice(0, 5) : ["TBD point"];

  let cursorY = y;
  const lineHeight = Math.max(0.62, Math.min(0.95, (h - 0.2) / items.length));
  for (const item of items) {
    slide.addShape("ellipse", {
      x,
      y: cursorY + 0.16,
      w: 0.17,
      h: 0.17,
      fill: { color: theme.accent },
      line: { color: theme.accent, pt: 0 },
    });
    slide.addText(item, {
      x: x + 0.33,
      y: cursorY,
      w: Math.max(0.3, w - 0.35),
      h: lineHeight,
      fontFace: "Microsoft YaHei",
      fontSize: 17,
      color: theme.muted,
      valign: "top",
      breakLine: false,
      margin: 0,
    });
    cursorY += lineHeight;
  }
}

function renderCover(pptx, topic, theme, subtitle) {
  const slide = pptx.addSlide();
  slide.background = { color: theme.bg };

  slide.addShape("rect", {
    x: 0,
    y: 0,
    w: 13.333,
    h: 7.5,
    fill: { color: theme.header, transparency: 32 },
    line: { color: theme.header, pt: 0 },
  });

  slide.addShape("roundRect", {
    x: 0.9,
    y: 1.1,
    w: 11.6,
    h: 5.2,
    radius: 0.12,
    fill: { color: theme.cardAlt },
    line: { color: theme.line, pt: 1.2 },
  });

  slide.addText(truncate(topic || "Report", 60), {
    x: 1.35,
    y: 2.0,
    w: 10.7,
    h: 1.5,
    fontFace: "Microsoft YaHei",
    bold: true,
    color: theme.text,
    fontSize: 44,
    valign: "top",
  });

  const subtitleText = String(subtitle || "").trim();
  if (subtitleText) {
    slide.addText(subtitleText, {
      x: 1.35,
      y: 3.85,
      w: 10.7,
      h: 0.7,
      fontFace: "Microsoft YaHei",
      color: theme.muted,
      fontSize: 20,
      valign: "top",
    });
  }

  addWarnings(slide, pptx);
}

function renderToc(pptx, topic, outline, theme) {
  const slide = pptx.addSlide();
  slide.background = { color: theme.bg };
  addTitleBar(slide, `${topic} | Agenda`, theme);

  slide.addShape("roundRect", {
    x: 0.8,
    y: 1.45,
    w: 11.8,
    h: 5.7,
    radius: 0.08,
    fill: { color: theme.card },
    line: { color: theme.line, pt: 1.2 },
  });

  const items = (Array.isArray(outline) ? outline : []).map(normalizeTocItem).filter(Boolean).slice(0, 10);
  let y = 1.8;
  items.forEach((item, idx) => {
    slide.addShape("ellipse", {
      x: 1.15,
      y: y + 0.05,
      w: 0.35,
      h: 0.35,
      fill: { color: theme.accent },
      line: { color: theme.accent, pt: 0 },
    });
    slide.addText(String(idx + 1), {
      x: 1.22,
      y: y + 0.07,
      w: 0.2,
      h: 0.2,
      fontFace: "Microsoft YaHei",
      bold: true,
      color: "FFFFFF",
      fontSize: 10,
      valign: "mid",
      align: "center",
    });
    slide.addText(truncate(item, 68), {
      x: 1.6,
      y,
      w: 10.2,
      h: 0.52,
      fontFace: "Microsoft YaHei",
      color: theme.text,
      fontSize: 19,
      valign: "mid",
    });
    y += 0.63;
  });

  addWarnings(slide, pptx);
}

function renderSummarySlide(pptx, data, theme) {
  const slide = pptx.addSlide();
  slide.background = { color: theme.bg };
  addTitleBar(slide, String(data.title || "Content"), theme);

  slide.addShape("roundRect", {
    x: 0.85,
    y: 1.5,
    w: 11.6,
    h: 5.75,
    radius: 0.08,
    fill: { color: theme.card },
    line: { color: theme.line, pt: 1.2 },
  });

  const hasImage = Boolean(data.generated_image_path && fs.existsSync(data.generated_image_path));
  if (hasImage) {
    addBulletList(slide, data.bullets, 1.2, 1.95, 7.0, 4.9, theme);
    addImageIfExists(slide, data.generated_image_path, 8.35, 1.95, 3.95, 4.8);
  } else {
    addBulletList(slide, data.bullets, 1.2, 1.95, 10.9, 4.9, theme);
  }

  addWarnings(slide, pptx);
}

function renderRiskSlide(pptx, data, theme) {
  const slide = pptx.addSlide();
  slide.background = { color: theme.bg };
  addTitleBar(slide, String(data.title || "Risk"), theme);

  slide.addShape("roundRect", {
    x: 0.8,
    y: 1.5,
    w: 5.7,
    h: 5.7,
    radius: 0.08,
    fill: { color: theme.card },
    line: { color: theme.line, pt: 1.2 },
  });
  slide.addShape("roundRect", {
    x: 6.8,
    y: 1.5,
    w: 5.7,
    h: 5.7,
    radius: 0.08,
    fill: { color: theme.cardAlt },
    line: { color: theme.line, pt: 1.2 },
  });

  slide.addText("Top Risks", {
    x: 1.1,
    y: 1.75,
    w: 4.8,
    h: 0.45,
    fontFace: "Microsoft YaHei",
    bold: true,
    color: theme.text,
    fontSize: 20,
  });

  slide.addText("Mitigation", {
    x: 7.1,
    y: 1.75,
    w: 4.8,
    h: 0.45,
    fontFace: "Microsoft YaHei",
    bold: true,
    color: theme.text,
    fontSize: 20,
  });

  addBulletList(slide, data.bullets, 1.1, 2.25, 5.0, 4.6, theme);
  addBulletList(slide, Array.isArray(data.evidence) && data.evidence.length > 0 ? data.evidence : data.bullets, 7.1, 2.25, 5.0, 4.6, theme);

  if (data.generated_image_path && fs.existsSync(data.generated_image_path)) {
    addImageIfExists(slide, data.generated_image_path, 4.95, 5.15, 3.45, 1.85);
  }

  addWarnings(slide, pptx);
}

function renderTimelineSlide(pptx, data, theme) {
  const slide = pptx.addSlide();
  slide.background = { color: theme.bg };
  addTitleBar(slide, String(data.title || "Timeline"), theme);

  slide.addShape("line", {
    x: 1.1,
    y: 3.65,
    w: 11.0,
    h: 0,
    line: { color: theme.line, pt: 2 },
  });

  const points = (Array.isArray(data.bullets) ? data.bullets : []).map((x) => truncate(x, 54)).filter(Boolean);
  const items = points.length > 0 ? points.slice(0, 4) : ["Stage detail TBD"];
  const n = items.length;

  for (let idx = 0; idx < n; idx += 1) {
    const x = 1.2 + idx * (10.6 / Math.max(1, n - 1));

    slide.addShape("ellipse", {
      x,
      y: 3.42,
      w: 0.42,
      h: 0.42,
      fill: { color: theme.accent },
      line: { color: theme.accent, pt: 0 },
    });

    slide.addShape("roundRect", {
      x: x - 0.5,
      y: 1.95,
      w: 1.5,
      h: 1.2,
      radius: 0.06,
      fill: { color: idx % 2 === 1 ? theme.cardAlt : theme.card },
      line: { color: theme.line, pt: 1.2 },
    });

    slide.addText(`Stage ${idx + 1}`, {
      x: x - 0.45,
      y: 2.02,
      w: 1.4,
      h: 0.2,
      fontFace: "Microsoft YaHei",
      bold: true,
      color: theme.muted,
      fontSize: 11,
    });

    slide.addText(items[idx], {
      x: x - 0.45,
      y: 2.28,
      w: 1.4,
      h: 0.8,
      fontFace: "Microsoft YaHei",
      color: theme.text,
      fontSize: 12,
      valign: "top",
    });
  }

  if (data.generated_image_path && fs.existsSync(data.generated_image_path)) {
    addImageIfExists(slide, data.generated_image_path, 8.55, 4.35, 3.65, 2.55);
  }

  addWarnings(slide, pptx);
}

function chartPayload(data) {
  const chart = data && typeof data.chart_data === "object" ? data.chart_data : null;
  if (chart) {
    const labels = Array.isArray(chart.labels) ? chart.labels.map((x) => String(x).trim()).filter(Boolean) : [];
    const values = Array.isArray(chart.values)
      ? chart.values
          .map((x) => Number(x))
          .filter((x) => Number.isFinite(x))
      : [];
    const unit = String(chart.unit || "");
    if (labels.length > 0 && labels.length === values.length) {
      return { labels: labels.slice(0, 6), values: values.slice(0, 6), unit };
    }
  }
  return { labels: ["Metric 1", "Metric 2", "Metric 3"], values: [60, 72, 84], unit: "" };
}

function renderDataSlide(pptx, data, theme) {
  const slide = pptx.addSlide();
  slide.background = { color: theme.bg };
  addTitleBar(slide, String(data.title || "Data"), theme);

  slide.addShape("roundRect", {
    x: 0.8,
    y: 1.5,
    w: 5.3,
    h: 5.7,
    radius: 0.08,
    fill: { color: theme.card },
    line: { color: theme.line, pt: 1.2 },
  });
  slide.addShape("roundRect", {
    x: 6.35,
    y: 1.5,
    w: 6.0,
    h: 5.7,
    radius: 0.08,
    fill: { color: theme.cardAlt },
    line: { color: theme.line, pt: 1.2 },
  });

  addBulletList(slide, data.bullets, 1.1, 1.95, 4.7, 4.9, theme);

  const chart = chartPayload(data);
  const hasImage = Boolean(data.generated_image_path && fs.existsSync(data.generated_image_path));
  const labels = hasImage ? chart.labels.slice(0, 4) : chart.labels;
  const values = hasImage ? chart.values.slice(0, 4) : chart.values;

  slide.addChart(
    "bar",
    [{ name: "Metrics", labels, values }],
    {
      x: 6.75,
      y: 2.0,
      w: hasImage ? 4.65 : 5.2,
      h: hasImage ? 2.15 : 3.75,
      barDir: "col",
      catAxisLabelPos: "nextTo",
      valAxisTitle: chart.unit ? `Value (${chart.unit})` : "Value",
      showLegend: false,
      chartColors: [theme.accent],
      gapWidthPct: 30,
    }
  );

  if (hasImage) {
    addImageIfExists(slide, data.generated_image_path, 8.25, 4.35, 3.9, 2.55);
  }

  addWarnings(slide, pptx);
}

function renderConclusion(pptx, bodySlides, theme) {
  const slide = pptx.addSlide();
  slide.background = { color: theme.bg };
  addTitleBar(slide, "Summary", theme);
  slide.addShape("roundRect", {
    x: 0.9,
    y: 1.5,
    w: 11.5,
    h: 5.75,
    radius: 0.08,
    fill: { color: theme.card },
    line: { color: theme.line, pt: 1.2 },
  });

  const keyPoints = [];
  bodySlides.slice(0, 5).forEach((item) => {
    const title = truncate(item.title || "Untitled", 42);
    const firstBullet = Array.isArray(item.bullets) && item.bullets.length > 0 ? truncate(item.bullets[0], 54) : "Key takeaway";
    keyPoints.push(`${title}: ${firstBullet}`);
  });

  addBulletList(slide, keyPoints, 1.25, 2.0, 10.8, 4.9, theme);
  addWarnings(slide, pptx);
}

async function exportFromScratch(payload, outPath) {
  const slides = Array.isArray(payload.slides) ? payload.slides : [];
  const body = contentSlides(slides);
  const topic = String(payload.topic || (body[0] && body[0].title) || "Report");
  const subtitle = String(payload.subtitle || payload.coverSubtitle || "").trim();
  const outline = Array.isArray(payload.outline) ? payload.outline : body.map((s) => String(s.title || ""));
  const theme = deriveTheme(String(payload.templateId || "default"));

  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "AIppt";
  pptx.subject = "Generated by pptx-generator workflow";
  pptx.company = "AIppt";
  pptx.title = topic;
  pptx.theme = {
    lang: "zh-CN",
    headFontFace: "Microsoft YaHei",
    bodyFontFace: "Microsoft YaHei",
  };

  renderCover(pptx, topic, theme, subtitle);
  renderToc(pptx, topic, outline, theme);

  body.forEach((slideData) => {
    const slideType = String(slideData.slide_type || "summary").toLowerCase();
    if (slideType === "risk") {
      renderRiskSlide(pptx, slideData, theme);
    } else if (slideType === "timeline" || slideType === "status") {
      renderTimelineSlide(pptx, slideData, theme);
    } else if (slideType === "data") {
      renderDataSlide(pptx, slideData, theme);
    } else {
      renderSummarySlide(pptx, slideData, theme);
    }
  });

  renderConclusion(pptx, body, theme);

  await pptx.writeFile({ fileName: outPath });
}

function escapeXml(text) {
  return String(text || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function normalizeReplacementText(text) {
  return String(text || "")
    .replace(/\r\n/g, "\n")
    .replace(/^\s*#{1,6}\s+/gm, "")
    .replace(/<\/?[A-Za-z_][A-Za-z0-9._:-]*(?:\s[^>\n]*)?>/g, "")
    .replace(/&lt;\/?[A-Za-z_][^&]{0,120}&gt;/gi, "")
    .replace(/[ \t]+/g, " ")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function replaceTextRunsInXml(xml, replacements) {
  if (!Array.isArray(replacements) || replacements.length === 0) {
    return xml;
  }

  const cleanValues = replacements.map((item) => normalizeReplacementText(item));
  let cursor = 0;
  return xml.replace(/<a:t(?![A-Za-z0-9_:-])([^>]*)>([\s\S]*?)<\/a:t>/g, (match, attrs) => {
    if (cursor >= cleanValues.length) {
      return match;
    }
    const value = escapeXml(cleanValues[cursor]);
    cursor += 1;
    return `<a:t${attrs}>${value}</a:t>`;
  });
}

function buildTemplateReplacement(index, topic, subtitle, outline, bodySlides) {
  if (index === 0) {
    return [topic, subtitle];
  }
  if (index === 1) {
    const toc = ["Agenda"];
    const lines = (Array.isArray(outline) ? outline : []).map(normalizeTocItem).filter(Boolean).slice(0, 10);
    lines.forEach((item, i) => toc.push(`${i + 1}. ${item}`));
    return toc;
  }

  const bodyIndex = index - 2;
  if (bodyIndex < 0 || bodyIndex >= bodySlides.length) {
    return [];
  }

  const payload = bodySlides[bodyIndex] || {};
  const replacements = [String(payload.title || "")];
  (Array.isArray(payload.bullets) ? payload.bullets : []).slice(0, 8).forEach((line) => replacements.push(String(line || "")));
  return replacements;
}

async function updatePresentationSlides(zip, keepCount) {
  const presentationPath = "ppt/presentation.xml";
  const relsPath = "ppt/_rels/presentation.xml.rels";

  const presentationFile = zip.file(presentationPath);
  if (!presentationFile) {
    return;
  }

  const presentationXml = await presentationFile.async("string");
  const listMatch = presentationXml.match(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/);
  if (!listMatch) {
    return;
  }

  const slideTags = [...listMatch[0].matchAll(/<p:sldId\b[^>]*\/>/g)].map((m) => m[0]);
  if (slideTags.length <= keepCount) {
    return;
  }

  const keptTags = slideTags.slice(0, keepCount);
  const keptRelIds = new Set(
    keptTags
      .map((tag) => {
        const m = tag.match(/\br:id="([^"]+)"/);
        return m ? m[1] : "";
      })
      .filter(Boolean),
  );

  const nextList = `<p:sldIdLst>${keptTags.join("")}</p:sldIdLst>`;
  const nextPresentationXml = presentationXml.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/, nextList);
  zip.file(presentationPath, nextPresentationXml);

  const relsFile = zip.file(relsPath);
  if (!relsFile) {
    return;
  }

  const relsXml = await relsFile.async("string");
  const nextRelsXml = relsXml.replace(/<Relationship\b[^>]*\/>/g, (tag) => {
    const typeMatch = tag.match(/\bType="([^"]+)"/);
    const idMatch = tag.match(/\bId="([^"]+)"/);
    if (!typeMatch || !idMatch) {
      return tag;
    }
    const isSlideRel = /\/relationships\/slide$/i.test(typeMatch[1]);
    if (!isSlideRel) {
      return tag;
    }
    return keptRelIds.has(idMatch[1]) ? tag : "";
  });
  zip.file(relsPath, nextRelsXml);
}

function sortSlideXmlPaths(paths) {
  return [...paths].sort((a, b) => {
    const ai = Number((a.match(/slide(\d+)\.xml$/) || ["", "0"])[1]);
    const bi = Number((b.match(/slide(\d+)\.xml$/) || ["", "0"])[1]);
    return ai - bi;
  });
}

async function exportByXmlTemplate(payload, outPath) {
  const templatePath = String(payload.templatePptxPath || "");
  if (!templatePath || !fs.existsSync(templatePath)) {
    throw new Error("template path missing or does not exist");
  }

  const slides = Array.isArray(payload.slides) ? payload.slides : [];
  const body = contentSlides(slides);
  const topic = String(payload.topic || (body[0] && body[0].title) || "Report");
  const subtitle = String(payload.subtitle || payload.coverSubtitle || "").trim();
  const outline = Array.isArray(payload.outline) ? payload.outline : body.map((s) => String(s.title || ""));

  const raw = fs.readFileSync(templatePath);
  const zip = await JSZip.loadAsync(raw);
  const slideXmlPaths = sortSlideXmlPaths(Object.keys(zip.files).filter((name) => /^ppt\/slides\/slide\d+\.xml$/i.test(name)));

  const desiredSlideCount = Math.max(1, 2 + body.length);
  const effectiveSlideCount = Math.min(desiredSlideCount, slideXmlPaths.length);

  for (let i = 0; i < effectiveSlideCount; i += 1) {
    const replacements = buildTemplateReplacement(i, topic, subtitle, outline, body);
    if (replacements.length === 0) {
      continue;
    }
    const xml = await zip.file(slideXmlPaths[i]).async("string");
    const nextXml = replaceTextRunsInXml(xml, replacements);
    zip.file(slideXmlPaths[i], nextXml);
  }

  await updatePresentationSlides(zip, effectiveSlideCount);

  const outBuffer = await zip.generateAsync({ type: "nodebuffer" });
  fs.writeFileSync(outPath, outBuffer);
}

async function main() {
  const args = parseArgs(process.argv);
  const payload = mustReadJson(args.input);

  fs.mkdirSync(path.dirname(args.output), { recursive: true });

  const hasTemplate = Boolean(payload.templatePptxPath && fs.existsSync(String(payload.templatePptxPath)));
  if (hasTemplate) {
    await exportByXmlTemplate(payload, args.output);
    return;
  }

  await exportFromScratch(payload, args.output);
}

main().catch((err) => {
  console.error(`[pptx_generator] ${String(err && err.stack ? err.stack : err)}`);
  process.exit(1);
});


