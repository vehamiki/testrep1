// Naturtrade investor pitch — English, 16:9 wide
const path = require("path");
const pptxgen = require("pptxgenjs");
const sharp = require("sharp");
const React = require("react");
const RDS = require("react-dom/server");
const fa = require("react-icons/fa6");
const { applyTheme } = require("/root/.claude/skills/synced/b0fb6637-2810-4f2b-a6bb-7d66797f2dd1_1b0b8ba3-92b6-4628-804e-44ec9921b9e0/pptx/scripts/apply_theme.js");

const MEDIA = path.resolve(__dirname, "../unz/ppt/media");
const OUT = process.argv[2] || path.resolve(__dirname, "Naturtrade_Investor_Pitch_2026_EN.pptx");

const HEX = {
  ink: "1F2D2B", dark: "2A4A44", white: "FFFFFF", mist: "F5F9F8",
  teal: "86CFC8", gold: "F6C28B", green: "2F7F72", tealLt: "C4E8E3",
  grey: "61706D", terra: "EBA99B", line: "DDE6E4", darkCard: "FFFFFF", muteOnDark: "61706D",
  soft: "E6F4F0", warmTxt: "B5701F", sage: "6FB3A8", icon: "2A4A44",
};
const THEME = {
  name: "Naturtrade Pastel",
  headFontFace: "Cambria",
  bodyFontFace: "Calibri",
  colors: {
    dk1: HEX.ink, lt1: HEX.white, dk2: HEX.dark, lt2: HEX.mist,
    accent1: HEX.teal, accent2: HEX.gold, accent3: HEX.green, accent4: HEX.tealLt,
    accent5: HEX.grey, accent6: HEX.terra, hlink: HEX.green, folHlink: HEX.green,
  },
};

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "Naturtrade — Investor Presentation 2026";
pres.author = "Miklós Véha";
pres.company = "Naturtrade Hungary Kft.";
const C = pres.SchemeColor;
const W = 13.333, H = 7.5, MX = 0.6;
const fs = require("fs");
const AI = path.resolve(__dirname, "../ai");
const ai = (name, fallback) => (fs.existsSync(path.join(AI, name)) ? path.join(AI, name) : fallback);

// ---------- helpers ----------
async function photo(file, w, h, opts = {}) {
  const px = Math.round(w * 160), py = Math.round(h * 160);
  let img = sharp(path.isAbsolute(file) ? file : path.join(MEDIA, file));
  if (opts.extract) img = img.extract(opts.extract);
  const buf = await img.resize(px, py, { fit: "cover", position: opts.position || "centre" })
    .jpeg({ quality: 80, mozjpeg: true }).toBuffer();
  return "image/jpeg;base64," + buf.toString("base64");
}
async function logo(file, boxW, boxH, opts = {}) {
  let img = sharp(path.join(MEDIA, file));
  if (opts.extract) img = img.extract(opts.extract);
  const buf = await img.png().toBuffer();
  const trimmed = await sharp(buf).trim({ threshold: 10 }).png().toBuffer().catch(() => buf);
  const m = await sharp(trimmed).metadata();
  const r = Math.min(boxW / m.width, boxH / m.height);
  const w = m.width * r, h = m.height * r;
  const resized = await sharp(trimmed).resize(Math.round(w * 200), Math.round(h * 200), { fit: "fill" }).png().toBuffer();
  return { data: "image/png;base64," + resized.toString("base64"), w, h };
}
const iconCache = {};
async function icon(name, color = "FFFFFF") {
  const k = name + color;
  if (iconCache[k]) return iconCache[k];
  const svg = RDS.renderToStaticMarkup(React.createElement(fa[name], { color: "#" + color, size: 256 }));
  const buf = await sharp(Buffer.from(svg)).resize(256, 256).png().toBuffer();
  return (iconCache[k] = "image/png;base64," + buf.toString("base64"));
}
// circle badge with icon (motif)
async function badge(slide, x, y, d, iconName, fill = C.accent1, iconColor = HEX.icon, nm = "badge") {
  slide.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { type: "none" }, objectName: nm });
  const s = d * 0.5;
  slide.addImage({ data: await icon(iconName, iconColor), x: x + (d - s) / 2, y: y + (d - s) / 2, w: s, h: s, objectName: nm + " icon", altText: iconName.replace("Fa", "") + " icon" });
}
function numBadge(slide, x, y, d, n, fill = C.accent1, color = C.text2) {
  slide.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { type: "none" }, objectName: "number badge " + n });
  slide.addText(String(n), { x, y, w: d, h: d, align: "center", valign: "middle", fontSize: Math.round(d * 30), bold: true, color, margin: 0, isTextBox: true, objectName: "number " + n });
}
function heads(slide, eyebrow, title) {
  slide.addText(eyebrow.toUpperCase(), { placeholder: "eyebrow" });
  slide.addText(title, { placeholder: "title" });
}
function source(slide, text, dark = false) {
  slide.addText(text, { x: MX, y: 6.62, w: 9.6, h: 0.3, fontSize: 10, color: C.accent5, margin: 0, isTextBox: true, objectName: "source note" });
}
function card(slide, x, y, w, h, fill = C.background2, nm = "card") {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.12, fill: { color: fill }, line: fill === HEX.darkCard ? { color: "D3E7E2", width: 0.75 } : { type: "none" }, objectName: nm });
}
function shadowCard(slide, x, y, w, h, nm = "card") {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h, rectRadius: 0.12, fill: { color: C.background1 }, line: { color: HEX.line, width: 0.75 },
    shadow: { type: "outer", color: "0E2B27", opacity: 0.12, blur: 8, offset: 2, angle: 90 }, objectName: nm,
  });
}
const T = (text, o = {}) => ({ text, options: o });

// ---------- layouts ----------
const footerText = { text: { text: "NATURTRADE  ·  Investor Presentation 2026  ·  Confidential", options: { x: MX, y: 7.02, w: 8, h: 0.3, fontSize: 9, color: C.accent5, margin: 0 } } };
pres.defineSlideMaster({
  title: "CONTENT_LIGHT",
  background: { color: C.background1 },
  objects: [
    { placeholder: { options: { name: "eyebrow", type: "body", x: MX, y: 0.42, w: 9, h: 0.32, fontSize: 12, bold: true, color: C.accent3, charSpacing: 3, margin: 0 }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: MX, y: 0.74, w: 12.1, h: 0.8, fontFace: "Cambria", fontSize: 34, bold: true, color: C.text1, margin: 0, valign: "top", align: "left" }, text: "" } },
    footerText,
  ],
  slideNumber: { x: 12.2, y: 7.0, w: 0.55, h: 0.3, fontSize: 9, color: C.accent5, align: "right" },
});
pres.defineSlideMaster({
  title: "CONTENT_DARK",
  background: { color: HEX.soft },
  objects: [
    { placeholder: { options: { name: "eyebrow", type: "body", x: MX, y: 0.42, w: 9, h: 0.32, fontSize: 12, bold: true, color: C.accent3, charSpacing: 3, margin: 0 }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: MX, y: 0.74, w: 12.1, h: 0.8, fontFace: "Cambria", fontSize: 34, bold: true, color: C.text1, margin: 0, valign: "top", align: "left" }, text: "" } },
    { text: { text: "NATURTRADE  ·  Investor Presentation 2026  ·  Confidential", options: { x: MX, y: 7.02, w: 8, h: 0.3, fontSize: 9, color: C.accent3, margin: 0 } } },
  ],
  slideNumber: { x: 12.2, y: 7.0, w: 0.55, h: 0.3, fontSize: 9, color: C.accent3, align: "right" },
});
pres.defineSlideMaster({ title: "COVER_DARK", background: { color: HEX.soft }, objects: [] });

const chartBase = {
  catAxisLabelColor: HEX.grey, valAxisLabelColor: HEX.grey, catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt",
  catAxisLabelFontSize: 12, valAxisLabelFontSize: 11, dataLabelFontFace: "+mn-lt", dataLabelFontSize: 11, dataLabelColor: HEX.ink,
  valGridLine: { color: "E3EAE8", size: 0.75 }, catGridLine: { style: "none" }, legendFontFace: "+mn-lt", legendFontSize: 12, legendColor: HEX.ink,
  catAxisLineShow: false, valAxisLineShow: false,
};

async function main() {
  let s;
  // ======================= 1. COVER =======================
  pres.addSection({ title: "Opening" });
  s = pres.addSlide({ masterName: "COVER_DARK", sectionTitle: "Opening" });
  s.addImage({ data: await photo(ai("hero.jpg", "image17.jpg"), 5.6, 7.5, { position: "east" }), x: W - 5.6, y: 0, w: 5.6, h: 7.5, objectName: "hero grains image", altText: "Pseudo-cereal grains pouring onto fine flour" });
  s.addText("INVESTOR PRESENTATION  ·  2026", { x: MX + 0.1, y: 1.15, w: 6.8, h: 0.35, fontSize: 13, bold: true, color: C.accent3, charSpacing: 4, margin: 0, isTextBox: true, objectName: "cover eyebrow" });
  s.addText("Naturtrade", { x: MX + 0.1, y: 1.6, w: 7, h: 1.2, fontFace: "Cambria", fontSize: 66, bold: true, color: C.text1, margin: 0, isTextBox: true, objectName: "cover title" });
  s.addText([
    T("Four generations of milling.", { breakLine: true }),
    T("One gluten-free, plant-based future.", { color: HEX.warmTxt }),
  ], { x: MX + 0.1, y: 2.95, w: 7, h: 1.1, fontFace: "Cambria", fontSize: 26, color: C.text1, margin: 0, isTextBox: true, objectName: "cover subtitle" });
  s.addText("Raising HUF 800M–1bn to build our new production plant in Röszke.", { x: MX + 0.1, y: 4.25, w: 6.4, h: 0.7, fontSize: 16, color: HEX.muteOnDark, margin: 0, isTextBox: true, objectName: "cover ask" });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: MX + 0.1, y: 5.45, w: 1.15, h: 1.15, rectRadius: 0.1, fill: { color: C.background1 }, line: { type: "none" }, objectName: "logo tile" });
  const lg = await logo("image1.png", 0.95, 0.95);
  s.addImage({ data: lg.data, x: MX + 0.1 + (1.15 - lg.w) / 2, y: 5.45 + (1.15 - lg.h) / 2, w: lg.w, h: lg.h, objectName: "Naturtrade logo", altText: "Naturtrade logo" });
  s.addText([
    T("Miklós Véha", { bold: true, color: C.text1, breakLine: true }),
    T("Founder & CEO, Naturtrade Hungary Kft.", { color: HEX.muteOnDark }),
  ], { x: MX + 1.5, y: 5.6, w: 5, h: 0.85, fontSize: 15, margin: 0, valign: "middle", isTextBox: true, objectName: "presenter" });
  s.addNotes("Open with the family story: the Véha mill was founded in 1931 in Endrőd. Four generations later we turn premium pseudo-cereal flours into gluten-free, plant-based foods. Today we are raising HUF 800M–1bn to build our Röszke plant.");

  // ======================= 2. AT A GLANCE =======================
  s = pres.addSlide({ masterName: "CONTENT_DARK", sectionTitle: "Opening" });
  heads(s, "The opportunity at a glance", "A heritage miller building a free-from food platform");
  const glance = [
    ["800M–1bn", "HUF investment sought", "All of it goes into the greenfield plant in Röszke and its production technology", "FaMoneyBillTrendUp"],
    ["1931", "Milling heritage", "Four generations of millers. We make our own premium flour", "FaWheatAwn"],
    ["250+", "B2B partners", "Over 100 products in 10+ categories on Hungarian shelves and in export", "FaStore"],
    ["3–5×", "Target return", "Exit to a financial or strategic investor in 2030", "FaRocket"],
  ];
  for (let i = 0; i < 4; i++) {
    const x = MX + i * 3.08, y = 1.95, w = 2.85, h = 3.5;
    card(s, x, y, w, h, HEX.darkCard, "glance card " + (i + 1));
    await badge(s, x + 0.3, y + 0.3, 0.7, glance[i][3], C.accent1, HEX.icon, "glance badge " + (i + 1));
    s.addText(glance[i][0], { x: x + 0.3, y: y + 1.15, w: w - 0.5, h: 0.8, fontFace: "Cambria", fontSize: 34, bold: true, color: HEX.warmTxt, margin: 0, isTextBox: true, objectName: "glance stat " + (i + 1) });
    s.addText(glance[i][1], { x: x + 0.3, y: y + 1.95, w: w - 0.5, h: 0.4, fontSize: 16, bold: true, color: C.text1, margin: 0, isTextBox: true, objectName: "glance label " + (i + 1) });
    s.addText(glance[i][2], { x: x + 0.3, y: y + 2.38, w: w - 0.5, h: 0.95, fontSize: 13, color: HEX.muteOnDark, margin: 0, valign: "top", isTextBox: true, objectName: "glance desc " + (i + 1) });
  }
  s.addText([
    T("Business plan:  ", { color: C.accent3, bold: true }),
    T("revenue from HUF 732M (2025) to HUF 2.5bn (2030), with EBITDA of about HUF 525M in 2030", { color: C.text1 }),
  ], { x: MX, y: 5.85, w: 12.1, h: 0.5, fontSize: 16, margin: 0, isTextBox: true, objectName: "plan summary" });
  s.addNotes("One slide that answers 'why should I listen'. The amount, the heritage, the traction and the return target.");

  // ======================= 3. AGENDA =======================
  s = pres.addSlide({ masterName: "CONTENT_LIGHT", sectionTitle: "Opening" });
  heads(s, "Agenda", "What we will cover");
  const agenda = [
    ["Heritage & team", "From the 1931 Véha mill to a 27-person food company"],
    ["The market", "Why free-from food is moving into the mainstream"],
    ["Products & moat", "From grain to cup, all made in-house"],
    ["Traction", "Shelves, partners, export and the Saudi deal"],
    ["Competition", "Positioning, SWOT and growth levers"],
    ["The investment", "Use of funds, capacity, plan and exit"],
  ];
  for (let i = 0; i < 6; i++) {
    const col = i % 3, row = Math.floor(i / 3);
    const x = MX + col * 4.1, y = 1.95 + row * 2.3, w = 3.85, h = 2.0;
    shadowCard(s, x, y, w, h, "agenda card " + (i + 1));
    numBadge(s, x + 0.3, y + 0.3, 0.62, i + 1);
    s.addText(agenda[i][0], { x: x + 0.3, y: y + 1.02, w: w - 0.5, h: 0.4, fontFace: "Cambria", fontSize: 19, bold: true, color: C.text1, margin: 0, isTextBox: true, objectName: "agenda title " + (i + 1) });
    s.addText(agenda[i][1], { x: x + 0.3, y: y + 1.42, w: w - 0.5, h: 0.45, fontSize: 13, color: C.accent5, margin: 0, isTextBox: true, objectName: "agenda desc " + (i + 1) });
  }

  // ======================= 4. HERITAGE =======================
  pres.addSection({ title: "Heritage & team" });
  s = pres.addSlide({ masterName: "CONTENT_LIGHT", sectionTitle: "Heritage & team" });
  heads(s, "01 · Heritage", "Where tradition meets innovation");
  s.addImage({ data: await photo("image10.jpg", 6.1, 3.6), x: MX, y: 1.8, w: 6.1, h: 3.6, objectName: "Véha mill 1931 photo", altText: "Historic photo of the Véha mill in Endrőd, 1931" });
  s.addText("The Véha mill, Endrőd, 1931", { x: MX, y: 5.48, w: 6.1, h: 0.3, fontSize: 11, italic: true, color: C.accent5, margin: 0, isTextBox: true, objectName: "photo caption" });
  const tl = [
    ["1931", "The Véha mill opens in Endrőd. Flour milling has been in the family ever since."],
    ["4 gen.", "Four generations of millers, now applied to pseudo-cereals and legumes."],
    ["Today", "Szeged plant: own milling, mixing, pasta and instant lines. 27 people."],
    ["2026–27", "Greenfield plant in Röszke, five minutes from the current site."],
    ["2030", "International free-from food platform, ready for exit."],
  ];
  const tx = 7.15, ty0 = 1.85, step = 0.98;
  s.addShape(pres.shapes.LINE, { x: tx + 0.2, y: ty0 + 0.2, w: 0, h: step * 4, line: { color: HEX.line, width: 2 }, objectName: "timeline line" });
  for (let i = 0; i < tl.length; i++) {
    const y = ty0 + i * step;
    s.addShape(pres.shapes.OVAL, { x: tx + 0.05, y: y + 0.05, w: 0.3, h: 0.3, fill: { color: i === tl.length - 1 ? C.accent2 : C.accent1 }, line: { color: C.background1, width: 2 }, objectName: "timeline dot " + (i + 1) });
    s.addText(tl[i][0], { x: tx + 0.55, y: y - 0.02, w: 1.15, h: 0.4, fontFace: "Cambria", fontSize: 18, bold: true, color: C.accent3, margin: 0, isTextBox: true, objectName: "timeline year " + (i + 1) });
    s.addText(tl[i][1], { x: tx + 1.75, y: y - 0.02, w: 3.8, h: 0.85, fontSize: 14, color: C.text1, margin: 0, valign: "top", isTextBox: true, objectName: "timeline text " + (i + 1) });
  }
  s.addNotes("Flour milling is in our genes. That heritage is our edge: we control the most important raw material, premium fine-ground flour.");

  // ======================= 5. SITE =======================
  s = pres.addSlide({ masterName: "CONTENT_LIGHT", sectionTitle: "Heritage & team" });
  heads(s, "01 · Our base today", "A working food plant in Szeged");
  const gx = MX, gy = 1.8, gap = 0.12;
  s.addImage({ data: await photo("image18.jpg", 5.0, 4.75), x: gx, y: gy, w: 5.0, h: 4.75, objectName: "site photo export truck", altText: "Export container truck at the Szeged site" });
  s.addImage({ data: await photo("image12.jpg", 3.5, 2.3), x: gx + 5.0 + gap, y: gy, w: 3.5, h: 2.3, objectName: "site photo extruder line", altText: "Production line" });
  s.addImage({ data: await photo("image11.jpg", 3.5, 2.33), x: gx + 5.0 + gap, y: gy + 2.3 + gap, w: 3.5, h: 2.33, objectName: "site photo milling hall", altText: "Milling hall with flour sacks" });
  s.addImage({ data: await photo("image15.jpg", 3.48, 2.3), x: gx + 8.62, y: gy, w: 3.48, h: 2.3, objectName: "site photo warehouse", altText: "Finished goods warehouse" });
  s.addImage({ data: await photo("image8.jpg", 3.48, 2.33), x: gx + 8.62, y: gy + 2.3 + gap, w: 3.48, h: 2.33, objectName: "site photo pasta drying", altText: "Fresh gluten-free pasta drying" });
  s.addText([
    T("Own milling  ·  ", { bold: true, color: C.accent3 }), T("flour blending  ·  pasta  ·  instant cups  ·  R&D test kitchen  ·  export logistics"),
  ], { x: MX, y: 6.62, w: 12.1, h: 0.32, fontSize: 13, color: C.text1, margin: 0, isTextBox: true, objectName: "site caption" });

  // ======================= 6. MISSION & VISION =======================
  s = pres.addSlide({ masterName: "CONTENT_DARK", sectionTitle: "Heritage & team" });
  heads(s, "01 · Purpose", "Free-from food that anyone can enjoy");
  s.addImage({ data: await photo("image14.jpg", 4.2, 4.9), x: W - MX - 4.2, y: 1.75, w: 4.2, h: 4.9, objectName: "R&D kitchen photo", altText: "R&D test kitchen" });
  const mv = [
    ["FaBullseye", "Mission", "We make gluten-free products that anyone can eat with confidence, whatever their sensitivity or diet. We combine professional experience with youthful energy to keep improving them."],
    ["FaEye", "Vision", "We want to build a company that leads international markets, where tradition, innovation, respect for nature and a family atmosphere matter equally. Our products should set trends and lead their category."],
  ];
  for (let i = 0; i < 2; i++) {
    const y = 1.85 + i * 2.45;
    await badge(s, MX, y, 0.75, mv[i][0], C.accent2, HEX.dark, "purpose badge " + (i + 1));
    s.addText(mv[i][1], { x: MX + 1.05, y: y + 0.05, w: 6, h: 0.45, fontFace: "Cambria", fontSize: 22, bold: true, color: C.text1, margin: 0, isTextBox: true, objectName: "purpose head " + (i + 1) });
    s.addText(mv[i][2], { x: MX + 1.05, y: y + 0.55, w: 6.3, h: 1.6, fontSize: 15, color: HEX.muteOnDark, margin: 0, valign: "top", isTextBox: true, objectName: "purpose text " + (i + 1) });
  }

  // ======================= 7. TEAM =======================
  s = pres.addSlide({ masterName: "CONTENT_LIGHT", sectionTitle: "Heritage & team" });
  heads(s, "01 · Team", "A young team with deep scientific roots");
  const team = [
    ["image3.jpg", "Miklós Véha", "Founder-owner & CEO", "Fourth-generation miller. Leads strategy, export and investor relations."],
    ["image4.jpg", "Prof. Dr. Antal Véha, PhD", "Head of R&D", "Vice-dean, Faculty of Engineering, University of Szeged. 200+ publications."],
    ["image19.jpg", "Orgona Tóthné", "Chief Operating Officer", "Runs production, quality and the plant floor."],
    ["image2.jpg", "Anita Labát", "Head of Sales", "Domestic retail, wholesale and key accounts."],
  ];
  for (let i = 0; i < 4; i++) {
    const x = MX + i * 3.08, w = 2.85;
    shadowCard(s, x, 1.85, w, 4.2, "team card " + (i + 1));
    s.addImage({ data: await photo(team[i][0], 2.0, 2.0, { position: "north" }), x: x + (w - 2.0) / 2, y: 2.1, w: 2.0, h: 2.0, rounding: true, objectName: "headshot " + team[i][1], altText: "Portrait of " + team[i][1] });
    s.addText(team[i][1], { x: x + 0.2, y: 4.3, w: w - 0.4, h: 0.45, fontFace: "Cambria", fontSize: 16, bold: true, color: C.text1, align: "center", margin: 0, isTextBox: true, objectName: "team name " + (i + 1) });
    s.addText(team[i][2], { x: x + 0.2, y: 4.75, w: w - 0.4, h: 0.35, fontSize: 13, bold: true, color: C.accent3, align: "center", margin: 0, isTextBox: true, objectName: "team role " + (i + 1) });
    s.addText(team[i][3], { x: x + 0.25, y: 5.12, w: w - 0.5, h: 0.85, fontSize: 12, color: C.accent5, align: "center", valign: "top", margin: 0, isTextBox: true, objectName: "team bio " + (i + 1) });
  }
  s.addText([T("27", { bold: true, color: C.accent3, fontFace: "Cambria", fontSize: 20 }), T("  people in total: production, R&D, sales, logistics and finance", { fontSize: 14 })],
    { x: MX, y: 6.38, w: 12.1, h: 0.42, color: C.text1, margin: 0, isTextBox: true, objectName: "headcount" });

  // ======================= 8. THE NEED =======================
  pres.addSection({ title: "The market" });
  s = pres.addSlide({ masterName: "CONTENT_DARK", sectionTitle: "The market" });
  heads(s, "02 · The need", "Millions need a grain-conscious diet");
  s.addText("Share of people on a permanent special diet that affects grain consumption", { x: MX, y: 1.62, w: 8, h: 0.35, fontSize: 14, color: HEX.muteOnDark, margin: 0, isTextBox: true, objectName: "need subtitle" });
  const need = [["10–15%", "IBS / non-coeliac gluten sensitivity"], ["10–15%", "Diabetics"], ["7–10%", "Chronic kidney disease"], ["1–2%", "Gluten allergy"], ["~1%", "Coeliac disease"], ["0.4%", "Crohn's disease"]];
  for (let i = 0; i < 6; i++) {
    const col = i % 3, row = Math.floor(i / 3);
    const x = MX + col * 2.62, y = 2.15 + row * 1.85, w = 2.42, h = 1.65;
    card(s, x, y, w, h, HEX.darkCard, "need card " + (i + 1));
    s.addText(need[i][0], { x: x + 0.25, y: y + 0.2, w: w - 0.4, h: 0.7, fontFace: "Cambria", fontSize: 32, bold: true, color: i < 2 ? HEX.warmTxt : C.accent3, margin: 0, isTextBox: true, objectName: "need stat " + (i + 1) });
    s.addText(need[i][1], { x: x + 0.25, y: y + 0.9, w: w - 0.4, h: 0.6, fontSize: 13, color: C.text1, margin: 0, valign: "top", isTextBox: true, objectName: "need label " + (i + 1) });
  }
  const nx = 8.75;
  s.addImage({ data: await photo("image21.png", 3.98, 2.3), x: nx, y: 1.62, w: 3.98, h: 2.3, objectName: "bread crumb photo", altText: "Close-up of bread crumb" });
  s.addText([
    T("85 kg", { fontFace: "Cambria", fontSize: 30, bold: true, color: HEX.warmTxt, breakLine: true }),
    T("flour per person per year in Hungary, against an EU average of 60 kg", { fontSize: 13, color: C.text1, breakLine: true }),
    T(" ", { fontSize: 8, breakLine: true }),
    T("1.2M t", { fontFace: "Cambria", fontSize: 30, bold: true, color: HEX.warmTxt, breakLine: true }),
    T("of wheat milled every year in Hungary. Gliadin, about 25% of gluten protein, drives the autoimmune response.", { fontSize: 13, color: C.text1 }),
  ], { x: nx, y: 4.05, w: 3.98, h: 2.55, margin: 0, valign: "top", isTextBox: true, objectName: "need facts" });
  s.addNotes("Coeliac disease is only the tip of the iceberg. IBS, gluten sensitivity, diabetes and kidney disease all push people towards grain-conscious diets. That is why food retail is so open to new species, new varieties and functional grain products.");

  // ======================= 9. MARKET SIZE =======================
  s = pres.addSlide({ masterName: "CONTENT_LIGHT", sectionTitle: "The market" });
  heads(s, "02 · Market size", "Every market we serve is growing");
  s.addChart(pres.charts.BAR, [
    { name: "Base year (2024/25)", labels: ["Gluten-free, global", "Gluten-free, Europe", "Plant-based meat, global"], values: [8.34, 2.72, 9.57] },
    { name: "2030 forecast", labels: ["Gluten-free, global", "Gluten-free, Europe", "Plant-based meat, global"], values: [11.78, 4.72, 21.81] },
  ], {
    x: MX, y: 1.75, w: 7.2, h: 4.75, barDir: "col", barGapWidthPct: 60, chartColors: [HEX.tealLt, HEX.sage],
    showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: '"$"0.0"bn"', valAxisHidden: true, valGridLine: { style: "none" },
    showLegend: true, legendPos: "t", showTitle: true, title: "Market size, USD bn", titleFontFace: "+mn-lt", titleFontSize: 14, titleColor: HEX.ink,
    ...chartBase, valGridLine: { style: "none" }, objectName: "market size chart",
  });
  const mk = [
    ["+7–10%", "Gluten-free CAGR to 2030", "Global GF products grow 7.2%/yr and Europe 9.6%/yr. The UK is Europe's largest GF market."],
    ["+9.7%", "Saudi gluten-free CAGR", "The KSA market was worth $127M in 2025. It is the largest GCC market, and the GCC grows about 10%/yr."],
    ["+14.7%", "Plant-based meat CAGR", "Global plant-based meat grows from $9.6bn (2024) to $21.8bn (2030). This is Oosty's arena."],
  ];
  for (let i = 0; i < 3; i++) {
    const y = 1.75 + i * 1.6;
    shadowCard(s, 8.2, y, 4.53, 1.42, "market card " + (i + 1));
    s.addText(mk[i][0], { x: 8.42, y: y + 0.18, w: 1.6, h: 0.55, fontFace: "Cambria", fontSize: 24, bold: true, color: C.accent3, margin: 0, isTextBox: true, objectName: "market stat " + (i + 1) });
    s.addText(mk[i][1], { x: 10.0, y: y + 0.2, w: 2.6, h: 0.5, fontSize: 13, bold: true, color: C.text1, margin: 0, valign: "middle", isTextBox: true, objectName: "market label " + (i + 1) });
    s.addText(mk[i][2], { x: 8.42, y: y + 0.75, w: 4.15, h: 0.6, fontSize: 11, color: C.accent5, margin: 0, valign: "top", isTextBox: true, objectName: "market desc " + (i + 1) });
  }
  source(s, "Sources: MarketsandMarkets (GF global, 2025–30); Ken Research (GF Europe, 2024–30); IMARC (KSA GF, 2025); TechSci (GCC GF); Arizton (plant-based meat, 2024–30).");
  s.addNotes("Third-party forecasts. Definitions differ between research houses, so treat these as directional. The point is that all of our categories grow at high single to double digits.");

  // ======================= 10. WHY NOW =======================
  s = pres.addSlide({ masterName: "CONTENT_LIGHT", sectionTitle: "The market" });
  heads(s, "02 · Why now", "Four shifts working in our favour");
  const why = [
    ["FaLeaf", "Free-from goes mainstream", "Gluten-free and vegan are no longer medical niches but lifestyle choices. The Hungarian GF market alone was worth over HUF 30bn in 2023."],
    ["FaMugHot", "Healthy convenience gap", "Instant food means flavour enhancers and allergens. Clean-label instant is almost empty shelf space, which is where Cupster is the first mover."],
    ["FaPeopleGroup", "Gen Z & flexitarians", "Millennials, Gen Z and Gen Alpha buy online, try new brands and want food with a story. That is our core consumer for instant food."],
    ["FaEarthEurope", "Import-dependent markets", "The UK, Western Europe and the GCC rely on imported free-from food. A Central European producer wins on lead time, price and flexibility."],
  ];
  for (let i = 0; i < 4; i++) {
    const col = i % 2, row = Math.floor(i / 2);
    const x = MX + col * 6.15, y = 1.85 + row * 2.4, w = 5.95, h = 2.15;
    shadowCard(s, x, y, w, h, "why card " + (i + 1));
    await badge(s, x + 0.3, y + 0.35, 0.8, why[i][0], C.accent1, HEX.icon, "why badge " + (i + 1));
    s.addText(why[i][1], { x: x + 1.35, y: y + 0.3, w: w - 1.6, h: 0.45, fontFace: "Cambria", fontSize: 19, bold: true, color: C.text1, margin: 0, isTextBox: true, objectName: "why head " + (i + 1) });
    s.addText(why[i][2], { x: x + 1.35, y: y + 0.8, w: w - 1.6, h: 1.2, fontSize: 14, color: C.accent5, margin: 0, valign: "top", isTextBox: true, objectName: "why text " + (i + 1) });
  }

  // ======================= 11. VALUE CHAIN =======================
  pres.addSection({ title: "Products & moat" });
  s = pres.addSlide({ masterName: "CONTENT_DARK", sectionTitle: "Products & moat" });
  heads(s, "03 · Our moat", "From grain to cup, all in-house");
  const chain = [
    ["FaSeedling", "Raw grains", "Millet, buckwheat, sorghum, quinoa, pea, lentil"],
    ["FaWheatAwn", "Own milling", "Homogeneous fine flour, down to 160 µm"],
    ["FaBreadSlice", "Flour mixes", "Bake-Free: just add water and bake"],
    ["FaBowlFood", "Pasta", "Pseudo-cereal & legume pasta"],
    ["FaMugHot", "Instant & premix", "Cupster cups · Oosty plant-based premixes"],
  ];
  const cw = 2.18, cg = 0.3;
  for (let i = 0; i < 5; i++) {
    const x = MX + i * (cw + cg), y = 2.0;
    card(s, x, y, cw, 2.9, HEX.darkCard, "chain card " + (i + 1));
    await badge(s, x + (cw - 0.95) / 2, y + 0.3, 0.95, chain[i][0], i === 1 ? C.accent2 : C.accent1, HEX.icon, "chain badge " + (i + 1));
    s.addText(chain[i][1], { x: x + 0.15, y: y + 1.4, w: cw - 0.3, h: 0.45, fontFace: "Cambria", fontSize: 17, bold: true, color: C.text1, align: "center", margin: 0, isTextBox: true, objectName: "chain head " + (i + 1) });
    s.addText(chain[i][2], { x: x + 0.15, y: y + 1.88, w: cw - 0.3, h: 0.9, fontSize: 13, color: HEX.muteOnDark, align: "center", valign: "top", margin: 0, isTextBox: true, objectName: "chain text " + (i + 1) });
    if (i < 4) s.addImage({ data: await icon("FaArrowRight", HEX.gold), x: x + cw + 0.04, y: y + 1.3, w: 0.22, h: 0.22, objectName: "chain arrow " + (i + 1), altText: "arrow" });
  }
  const moat = [["Margin", "We don't pay a miller's margin on our most important input"], ["Speed", "From idea to shelf without outside suppliers"], ["Contract manufacturing", "We produce for competitors too, so the margin stays with us"]];
  for (let i = 0; i < 3; i++) {
    const x = MX + i * 4.1;
    s.addText([T(moat[i][0], { bold: true, color: HEX.warmTxt, breakLine: true }), T(moat[i][1], { color: C.text1 })],
      { x, y: 5.25, w: 3.85, h: 1.05, fontSize: 14, margin: 0, valign: "top", isTextBox: true, objectName: "moat point " + (i + 1) });
  }
  s.addNotes("This is the core of the investment case. Competitors buy flour and outsource production; we own the full chain.");

  // ======================= 12. PORTFOLIO =======================
  s = pres.addSlide({ masterName: "CONTENT_LIGHT", sectionTitle: "Products & moat" });
  heads(s, "03 · Portfolio", "Five categories, 75+ recipes, one platform");
  s.addImage({ data: await photo("image7.png", 4.6, 4.85), x: W - MX - 4.6, y: 1.75, w: 4.6, h: 4.85, objectName: "portfolio lifestyle photo", altText: "Product range lifestyle photo" });
  const port = [
    ["15", "Mono flours", "Premium fine-ground pseudo-cereals and legumes"],
    ["15", "Flour mixes", "Ready to bake: just add water"],
    ["30", "Gluten-free pastas", "Millet, buckwheat, sorghum, quinoa, lentil, pea"],
    ["15", "Instant soups & noodles", "10 soups and 5 noodle cups. A category creator"],
    ["2→7", "Plant-based premixes", "2 'cheese' mixes now, 5 'meat' mixes coming"],
  ];
  for (let i = 0; i < 5; i++) {
    const y = 1.78 + i * 0.97;
    s.addText(port[i][0], { x: MX, y, w: 1.35, h: 0.8, fontFace: "Cambria", fontSize: 32, bold: true, color: C.accent3, margin: 0, valign: "middle", isTextBox: true, objectName: "portfolio count " + (i + 1) });
    s.addText([T(port[i][1], { bold: true, fontSize: 16, color: C.text1, breakLine: true }), T(port[i][2], { fontSize: 13, color: C.accent5 })],
      { x: MX + 1.45, y, w: 5.6, h: 0.8, margin: 0, valign: "middle", isTextBox: true, objectName: "portfolio item " + (i + 1) });
    if (i < 4) s.addShape(pres.shapes.LINE, { x: MX, y: y + 0.89, w: 7.05, h: 0, line: { color: HEX.line, width: 0.75 }, objectName: "portfolio divider " + (i + 1) });
  }
  s.addText("Target: vegans, vegetarians, the health-conscious and the gluten-sensitive. For instant: Millennials, Gen Z and Gen Alpha.", { x: MX, y: 6.5, w: 7.1, h: 0.45, fontSize: 11, italic: true, color: C.accent5, margin: 0, valign: "top", isTextBox: true, objectName: "portfolio target" });

  // ======================= 13. BRANDS =======================
  s = pres.addSlide({ masterName: "CONTENT_LIGHT", sectionTitle: "Products & moat" });
  heads(s, "03 · Brand house", "Each brand has its own consumer and message");
  const brands = [
    ["Bake-Free", "Free-from, made for enjoyment", "Ready-to-bake gluten-, dairy- and egg-free mixes. Strong in organic and health-food stores.", "image29.jpg"],
    ["Cupster", "The healthy instant", "Vegan, gluten-free cups with no flavour enhancers, aimed at young urban adults.", "image26.jpg"],
    ["Oosty", "Plant-based cheese & meat", "Powdered premixes made at home. Long shelf life and affordable. Mostly sold for export.", "image40.jpg"],
    ["Mill & Folks", "Our international DTC brand", "Own webshop plus Amazon, for consumers in the UK and Western Europe.", "image16.jpg"],
  ];
  for (let i = 0; i < 4; i++) {
    const x = MX + i * 3.08, w = 2.85;
    shadowCard(s, x, 1.8, w, 4.85, "brand card " + (i + 1));
    s.addImage({ data: await photo(brands[i][3], w, 2.55), x, y: 1.8, w, h: 2.55, objectName: "brand photo " + brands[i][0], altText: brands[i][0] + " product photo" });
    s.addText(brands[i][0], { x: x + 0.25, y: 4.5, w: w - 0.5, h: 0.45, fontFace: "Cambria", fontSize: 20, bold: true, color: C.text1, margin: 0, isTextBox: true, objectName: "brand name " + (i + 1) });
    s.addText(brands[i][1], { x: x + 0.25, y: 4.95, w: w - 0.5, h: 0.35, fontSize: 13, bold: true, italic: true, color: C.accent3, margin: 0, isTextBox: true, objectName: "brand tagline " + (i + 1) });
    s.addText(brands[i][2], { x: x + 0.25, y: 5.35, w: w - 0.5, h: 1.2, fontSize: 12, color: C.accent5, margin: 0, valign: "top", isTextBox: true, objectName: "brand desc " + (i + 1) });
  }
  s.addText([T("Eden Premium", { bold: true, color: C.accent3 }), T(" is our domestic umbrella brand and multi-brand webshop, stocked by 250+ resellers.")], { x: MX, y: 6.72, w: 11, h: 0.28, fontSize: 11, color: C.accent5, margin: 0, isTextBox: true, objectName: "eden note" });

  // ======================= 14–18. PRODUCT SPOTLIGHTS =======================
  async function spotlight(eyebrow, title, big, bigLabel, bullets, imgs, notes) {
    const sl = pres.addSlide({ masterName: "CONTENT_LIGHT", sectionTitle: "Products & moat" });
    heads(sl, eyebrow, title);
    // left text column
    sl.addText(big, { x: MX, y: 1.85, w: 4.6, h: 1.0, fontFace: "Cambria", fontSize: 54, bold: true, color: C.accent3, margin: 0, isTextBox: true, objectName: "spotlight stat" });
    sl.addText(bigLabel, { x: MX, y: 2.85, w: 4.6, h: 0.45, fontSize: 16, bold: true, color: C.text1, margin: 0, isTextBox: true, objectName: "spotlight stat label" });
    for (let i = 0; i < bullets.length; i++) {
      const y = 3.78 + i * 0.93;
      sl.addShape(pres.shapes.OVAL, { x: MX, y, w: 0.34, h: 0.34, fill: { color: C.accent1 }, line: { type: "none" }, objectName: "check dot " + (i + 1) });
      sl.addImage({ data: await icon("FaCheck", HEX.icon), x: MX + 0.08, y: y + 0.08, w: 0.18, h: 0.18, objectName: "check icon " + (i + 1), altText: "check" });
      sl.addText([T(bullets[i][0], { bold: true, color: C.text1, breakLine: true }), T(bullets[i][1], { color: C.accent5, fontSize: 13 })],
        { x: MX + 0.5, y: y - 0.04, w: 4.1, h: 0.85, fontSize: 15, margin: 0, valign: "top", isTextBox: true, objectName: "spotlight point " + (i + 1) });
    }
    // right photo mosaic: one tall + two stacked
    const px = 5.65, pw1 = 3.6, pw2 = 3.32, ph = 4.85, g = 0.12;
    sl.addImage({ data: await photo(imgs[0], pw1, ph), x: px, y: 1.75, w: pw1, h: ph, objectName: "product photo 1", altText: "Product photo" });
    sl.addImage({ data: await photo(imgs[1], pw2, (ph - g) / 2), x: px + pw1 + g, y: 1.75, w: pw2, h: (ph - g) / 2, objectName: "product photo 2", altText: "Product photo" });
    sl.addImage({ data: await photo(imgs[2], pw2, (ph - g) / 2), x: px + pw1 + g, y: 1.75 + (ph + g) / 2, w: pw2, h: (ph - g) / 2, objectName: "product photo 3", altText: "Product photo" });
    if (notes) sl.addNotes(notes);
  }
  await spotlight("03 · Own-milled flour", "Premium flour is our foundation", "15", "single-ingredient flours from pseudo-cereals and legumes",
    [["Homogeneous particle size", "Ground as fine as 160 microns for better texture and baking"], ["Our own key ingredient", "Millet, buckwheat, sorghum, quinoa, green pea, red lentil"], ["Fully traceable", "Controlled sourcing and composition, in our own packaging"]],
    ["image16.jpg", "image5.jpg", "image17.jpg"]);
  await spotlight("03 · Bake-Free", "Bake-free flour mixes: just add water", "15", "ready-to-bake gluten-free flour mixes",
    [["Easy to make", "Just add water and bake: bread, cake, pizza"], ["Free from gluten, dairy and egg", "Taste and texture that hold up against the real thing"], ["Proven shelf presence", "Stocked in organic, health-food and drugstore chains"]],
    ["image27.jpg", "image20.jpg", "image29.jpg"]);
  await spotlight("03 · Pasta", "Gluten-free pasta from ancient grains", "30", "pasta varieties from pseudo-cereals and legumes",
    [["Pseudo-cereal pasta", "Millet, buckwheat, sorghum, quinoa, lentil, pea"], ["Made in-house from our own flour", "Not contract-made, unlike the main competitor"], ["Strong in export", "Long shelf life and easy to list with retailers"]],
    ["image23.jpg", "image24.jpg", "image25.jpg"]);
  await spotlight("03 · Cupster", "Cupster: the healthy instant category", "15", "instant soups (10) and noodle cups (5), and growing fast",
    [["Vegan, gluten-free, no flavour enhancers", "A clean-label alternative to mainstream instant brands"], ["New pack at a HUF 599 RRP", "With a material cost of only ~20%"], ["Our biggest breakout point", "More sub-categories are in development"]],
    ["image37.jpg", "image38.jpg", "image32.jpg"], "Cupster is the category creator. There is no direct domestic competitor in healthy instant food.");
  await spotlight("03 · Oosty", "Oosty: plant-based cheese & meat at home", "2→7", "'cheese' mixes today, with 5 'meat' mixes coming soon",
    [["Easy to make at home", "Powder format: just mix and cook"], ["Long shelf life, affordable", "Far cheaper than ready-made vegan analogues"], ["Built for export", "Most Oosty volume goes abroad"]],
    ["image41.jpg", "image28.jpg", "image51.jpg"]);

  // ======================= 19. TRACTION =======================
  pres.addSection({ title: "Traction" });
  s = pres.addSlide({ masterName: "CONTENT_DARK", sectionTitle: "Traction" });
  heads(s, "04 · Traction", "On shelves at home, growing abroad");
  const tr = [["100+", "products"], ["250+", "B2B resellers"], ["10+", "product categories"], ["4", "sales channels"]];
  for (let i = 0; i < 4; i++) {
    const x = MX + i * 3.08;
    s.addText(tr[i][0], { x, y: 1.75, w: 2.85, h: 0.95, fontFace: "Cambria", fontSize: 50, bold: true, color: HEX.warmTxt, margin: 0, isTextBox: true, objectName: "traction stat " + (i + 1) });
    s.addText(tr[i][1], { x, y: 2.68, w: 2.85, h: 0.4, fontSize: 16, color: C.text1, margin: 0, isTextBox: true, objectName: "traction label " + (i + 1) });
  }
  const ch = [
    ["FaStore", "Hungary · B2B", "We cover the health-food wholesalers, drugstores and organic shops. Multinational and discount chains are next."],
    ["FaLaptop", "Hungary · DTC", "Eden Premium multi-shop. We list partner brands to raise basket value."],
    ["FaHandshake", "International · B2B", "Distributors in the UK, Western Europe and KSA. We pivot to private label where a foreign brand doesn't fit."],
    ["FaCartShopping", "International · DTC", "Mill & Folks webshop and Amazon. Direct-to-consumer sales have started."],
  ];
  for (let i = 0; i < 4; i++) {
    const x = MX + i * 3.08, y = 3.45, w = 2.85, h = 3.2;
    card(s, x, y, w, h, HEX.darkCard, "channel card " + (i + 1));
    await badge(s, x + 0.25, y + 0.25, 0.65, ch[i][0], C.accent1, HEX.icon, "channel badge " + (i + 1));
    s.addText(ch[i][1], { x: x + 0.25, y: y + 1.05, w: w - 0.45, h: 0.4, fontSize: 16, bold: true, color: C.text1, margin: 0, isTextBox: true, objectName: "channel head " + (i + 1) });
    s.addText(ch[i][2], { x: x + 0.25, y: y + 1.48, w: w - 0.45, h: 1.6, fontSize: 13, color: HEX.muteOnDark, margin: 0, valign: "top", isTextBox: true, objectName: "channel text " + (i + 1) });
  }

  // ======================= 20. PARTNERS =======================
  s = pres.addSlide({ masterName: "CONTENT_LIGHT", sectionTitle: "Traction" });
  heads(s, "04 · Partners", "Trusted by retailers at home and abroad");
  const logos = [
    ["image44.png", "Amazon"], ["image70.png", "Whole Foods Market"], ["image36.png", "Müller"], ["image54.jpg", "Rossmann"], ["image42.jpg", "Interspar"],
    ["image52.png", "Spar", { left: 290, top: 90, width: 440, height: 520 }], ["image53.jpg", "CBA"], ["image30.jpg", "Ecofamily"], ["image39.png", "Kifli.hu"], ["image47.png", "Herbaház"],
    ["image34.png", "BioEgészség Biobolt"], ["image48.png", "Bijó"], ["image49.png", "Biocity"], ["image46.png", "Bio-Barát"], ["image43.png", "Plus Market"],
  ];
  const lw = 2.2, lh = 1.3, lgx = 0.275;
  for (let i = 0; i < logos.length; i++) {
    const col = i % 5, row = Math.floor(i / 5);
    const x = MX + col * (lw + lgx), y = 1.85 + row * (lh + 0.3);
    shadowCard(s, x, y, lw, lh, "logo card " + logos[i][1]);
    const L = await logo(logos[i][0], lw - 0.5, lh - 0.4, logos[i][2] ? { extract: logos[i][2] } : {});
    s.addImage({ data: L.data, x: x + (lw - L.w) / 2, y: y + (lh - L.h) / 2, w: L.w, h: L.h, objectName: "logo " + logos[i][1], altText: logos[i][1] + " logo" });
  }
  s.addText("A selection of retail, wholesale and marketplace partners carrying Naturtrade brands.", { x: MX, y: 6.68, w: 12, h: 0.3, fontSize: 11, italic: true, color: C.accent5, margin: 0, isTextBox: true, objectName: "partners note" });

  // ======================= 21. FOOTPRINT =======================
  s = pres.addSlide({ masterName: "CONTENT_LIGHT", sectionTitle: "Traction" });
  heads(s, "04 · Footprint", "From Hungary to the world");
  const mp = await logo("image56.png", 7.6, 4.9);
  s.addImage({ data: mp.data, x: MX, y: 1.75 + (4.9 - mp.h) / 2, w: mp.w, h: mp.h, objectName: "world map", altText: "World map highlighting current and target export markets" });
  const fp = [
    ["FaLandmark", C.accent4, "Home market", "Hungary", "Our widest coverage: organic, health-food, drugstore and retail chains."],
    ["FaPlaneDeparture", C.accent1, "Exporting today", "UK · Western Europe · Saudi Arabia", "Distributors, Amazon and a first industrial-scale GF bakery in KSA."],
    ["FaRocket", C.accent2, "Next markets", "USA · Canada · GCC (MENA)", "Plus deeper B2B and DTC reach in Western Europe."],
  ];
  for (let i = 0; i < 3; i++) {
    const y = 1.85 + i * 1.6;
    await badge(s, 8.55, y, 0.65, fp[i][0], fp[i][1], HEX.icon, "footprint badge " + (i + 1));
    s.addText([T(fp[i][2].toUpperCase(), { fontSize: 11, bold: true, color: C.accent5, charSpacing: 2, breakLine: true }), T(fp[i][3], { fontSize: 16, bold: true, color: C.text1, breakLine: true }), T(fp[i][4], { fontSize: 12, color: C.accent5 })],
      { x: 9.4, y: y - 0.05, w: 3.35, h: 1.45, margin: 0, valign: "top", isTextBox: true, objectName: "footprint text " + (i + 1) });
  }

  // ======================= 22. SALES BY CATEGORY =======================
  s = pres.addSlide({ masterName: "CONTENT_LIGHT", sectionTitle: "Traction" });
  heads(s, "04 · Momentum", "Cupster is our fastest-growing category");
  const cats = ["Flour", "Flour mixes", "Pasta", "Cupster", "Oosty"];
  s.addChart(pres.charts.BAR, [
    { name: "2024", labels: cats, values: [128, 144, 137, 30, 5] },
    { name: "2025", labels: cats, values: [150, 200, 160, 150, 50] },
    { name: "2026 plan", labels: cats, values: [200, 250, 200, 350, 80] },
  ], {
    x: MX, y: 1.75, w: 8.1, h: 4.85, barDir: "col", barGapWidthPct: 55, chartColors: [HEX.line, HEX.tealLt, HEX.sage],
    showValue: true, dataLabelPosition: "outEnd", showLegend: true, legendPos: "t", valAxisHidden: true,
    showTitle: true, title: "Sales by category, HUF million", titleFontFace: "+mn-lt", titleFontSize: 14, titleColor: HEX.ink,
    ...chartBase, valGridLine: { style: "none" }, dataLabelFontSize: 10, objectName: "category sales chart",
  });
  card(s, 9.05, 1.85, 3.68, 2.25, C.background2, "cupster callout");
  s.addText([T("11.7×", { fontFace: "Cambria", fontSize: 40, bold: true, color: C.accent3, breakLine: true }), T("Cupster growth from 2024 to the 2026 plan (HUF 30M → 350M)", { fontSize: 14, color: C.text1 })],
    { x: 9.3, y: 2.0, w: 3.25, h: 1.95, margin: 0, valign: "top", isTextBox: true, objectName: "cupster stat" });
  card(s, 9.05, 4.3, 3.68, 2.25, C.background2, "oosty callout");
  s.addText([T("16×", { fontFace: "Cambria", fontSize: 40, bold: true, color: C.accent3, breakLine: true }), T("Oosty, our newest category, is export-led (HUF 5M → 80M)", { fontSize: 14, color: C.text1 })],
    { x: 9.3, y: 4.45, w: 3.25, h: 1.95, margin: 0, valign: "top", isTextBox: true, objectName: "oosty stat" });

  // ======================= 23. KSA =======================
  s = pres.addSlide({ masterName: "CONTENT_DARK", sectionTitle: "Traction" });
  heads(s, "04 · Saudi Arabia", "The Saudi connection");
  s.addImage({ data: await photo("image67.jpg", 2.6, 4.85), x: MX, y: 1.75, w: 2.6, h: 4.85, objectName: "KSA bakery photo", altText: "Industrial gluten-free bakery in Saudi Arabia" });
  s.addImage({ data: await photo("image65.jpg", 2.6, 2.37, { position: "north" }), x: MX + 2.72, y: 1.75, w: 2.6, h: 2.37, objectName: "KSA partner photo", altText: "Meeting with Saudi partner" });
  s.addImage({ data: await photo("image68.jpg", 2.6, 2.37, { position: "north" }), x: MX + 2.72, y: 4.23, w: 2.6, h: 2.37, objectName: "KSA site photo", altText: "Site visit in Saudi Arabia" });
  const ksa = [["A first", "The first Hungarian company to export to KSA's first industrial-scale gluten-free bakery."], ["State-backed demand", "The Saudi state is the end buyer. Volumes are predictable and payment risk is low."], ["A gateway", "We aim to bring the partnership to Europe within 2–3 years."]];
  for (let i = 0; i < 3; i++) {
    s.addText([T(ksa[i][0], { bold: true, color: HEX.warmTxt, breakLine: true }), T(ksa[i][1], { color: C.text1 })],
      { x: 6.2, y: 1.8 + i * 0.92, w: 6.5, h: 0.85, fontSize: 14, margin: 0, valign: "top", isTextBox: true, objectName: "ksa point " + (i + 1) });
  }
  s.addChart(pres.charts.BAR, [{ name: "Export potential, EUR M", labels: ["Last 12 months", "2025", "2026", "2027"], values: [0.2, 0.4, 1.0, 1.5] }], {
    x: 6.1, y: 4.55, w: 6.6, h: 2.15, barDir: "col", barGapWidthPct: 45, chartColors: [HEX.gold], showValue: true, dataLabelPosition: "outEnd",
    dataLabelFormatCode: '"€"0.0"M"', valAxisHidden: true, showLegend: false,
    ...chartBase, valGridLine: { style: "none" }, catAxisLabelColor: HEX.grey, dataLabelColor: HEX.ink, objectName: "KSA export chart",
  });

  // ======================= 24. NOPEX =======================
  s = pres.addSlide({ masterName: "CONTENT_LIGHT", sectionTitle: "Traction" });
  heads(s, "04 · Visibility", "On the international stage: NOPEX London 2025");
  s.addImage({ data: await photo("image50.jpg", 3.9, 4.85), x: MX, y: 1.75, w: 3.9, h: 4.85, objectName: "NOPEX poster photo", altText: "Natural and Organic Products Europe 2025 poster" });
  s.addImage({ data: await photo("image69.jpg", 3.9, 4.85), x: MX + 4.1, y: 1.75, w: 3.9, h: 4.85, objectName: "Mill and Folks stand photo", altText: "Mill and Folks stand at NOPEX London" });
  s.addImage({ data: await photo("image66.jpg", 3.9, 4.85), x: MX + 8.2, y: 1.75, w: 3.9, h: 4.85, objectName: "Whole Foods photo", altText: "Whole Foods Market visit in London" });

  // ======================= 25. COMPETITION =======================
  pres.addSection({ title: "Competition" });
  s = pres.addSlide({ masterName: "CONTENT_LIGHT", sectionTitle: "Competition" });
  heads(s, "05 · Competition", "Only one player owns the whole chain");
  const hdr = (t) => ({ text: t, options: { bold: true, color: C.text1, fill: { color: C.accent4 }, align: "center", valign: "middle", fontSize: 12 } });
  const yes = { text: "●", options: { color: C.accent3, align: "center", fontSize: 18 } };
  const no = { text: "○", options: { color: C.accent5, align: "center", fontSize: 18 } };
  const part = { text: "◐", options: { color: HEX.warmTxt, align: "center", fontSize: 18 } };
  const rowL = (t) => ({ text: t, options: { bold: true, color: C.text1, fontSize: 13, valign: "middle" } });
  const rows = [
    [{ text: "", options: { fill: { color: C.accent4 } } }, hdr("Naturtrade"), hdr("Szafi Products"), hdr("Hunorganic (It's Us)"), hdr("Import brands")],
    [rowL("Own flour milling"), yes, no, no, no],
    [rowL("Own production"), yes, part, part, yes],
    [rowL("Healthy instant (cups)"), yes, no, no, part],
    [rowL("Plant-based premixes"), yes, no, no, no],
    [rowL("Export & GCC relationships"), yes, part, no, yes],
    [rowL("Local speed and flexibility"), yes, yes, yes, no],
  ];
  s.addTable(rows, { x: MX, y: 1.8, w: 7.7, colW: [2.7, 1.25, 1.25, 1.25, 1.25], rowH: [0.6, 0.45, 0.45, 0.45, 0.45, 0.45, 0.45], border: { type: "solid", color: HEX.line, pt: 0.75 }, fontFace: "Calibri", objectName: "competition matrix" });
  s.addText("● yes   ◐ partly   ○ no   ·   Import brands: e.g. Schär, dmBio", { x: MX, y: 5.3, w: 7, h: 0.3, fontSize: 11, color: C.accent5, margin: 0, isTextBox: true, objectName: "matrix legend" });
  const comp = [
    ["Szafi Products Kft.", "2024 revenue HUF 1.7bn · EBIT HUF 194M", "Similar portfolio. No own milling; pasta is contract-made. Ownership changed 4 years ago and growth has slowed since."],
    ["Hunorganic Kft. (It's Us)", "Revenue HUF 910M · EBIT HUF 71M", "Similar portfolio. An early multinational-chain listing and pioneer brand, acquired by Foodnet (Alpro)."],
  ];
  for (let i = 0; i < 2; i++) {
    const y = 1.8 + i * 2.05;
    shadowCard(s, 8.6, y, 4.13, 1.85, "competitor card " + (i + 1));
    s.addText([T(comp[i][0], { bold: true, fontSize: 15, color: C.text1, breakLine: true }), T(comp[i][1], { bold: true, fontSize: 12, color: C.accent3, breakLine: true }), T(comp[i][2], { fontSize: 12, color: C.accent5 })],
      { x: 8.82, y: y + 0.15, w: 3.75, h: 1.6, margin: 0, valign: "top", isTextBox: true, objectName: "competitor text " + (i + 1) });
  }
  card(s, 8.6, 5.95, 4.13, 0.75, C.background2, "precedent card");
  s.addText([T("Exit precedent: ", { bold: true, color: C.accent3 }), T("Hunorganic was bought by a strategic buyer.", { color: C.text1 })], { x: 8.8, y: 5.95, w: 3.8, h: 0.75, fontSize: 13, margin: 0, valign: "middle", isTextBox: true, objectName: "precedent text" });
  s.addText("Our edge: own milling, a constant stream of new products, strong export potential and GCC market knowledge and relationships.", { x: MX, y: 5.85, w: 7.7, h: 0.8, fontSize: 14, bold: true, italic: true, color: C.text1, margin: 0, valign: "top", isTextBox: true, objectName: "edge statement" });

  // ======================= 26. SWOT =======================
  s = pres.addSlide({ masterName: "CONTENT_LIGHT", sectionTitle: "Competition" });
  heads(s, "05 · SWOT", "Clear strengths, known risks, a plan for both");
  const swot = [
    ["S", "Strengths", C.accent1, "FFFFFF", ["Our own premium milled flour", "Constant R&D-driven product innovation", "Contract manufacturing for competitors"]],
    ["W", "Weaknesses", HEX.line, "FFFFFF", ["Undercapitalised, like most Hungarian SMEs", "Smaller marketing budget than Western peers", "Export growth held back by capacity"]],
    ["O", "Opportunities", C.accent2, "FFFFFF", ["Contract manufacturing: a HUF 1.2bn opportunity", "Export: a HUF 2.1bn opportunity", "Western Europe, North America, Middle East"]],
    ["T", "Threats", C.accent6, "FFFFFF", ["Unpredictable regulation: special taxes, price and margin caps", "Competition", "Global shocks: tariffs, Covid, the war in Ukraine"]],
  ];
  for (let i = 0; i < 4; i++) {
    const col = i % 2, row = Math.floor(i / 2);
    const x = MX + col * 6.15, y = 1.8 + row * 2.45, w = 5.95, h = 2.25;
    shadowCard(s, x, y, w, h, "swot card " + swot[i][1]);
    numBadge(s, x + 0.3, y + 0.3, 0.8, swot[i][0], swot[i][2]);
    s.addText(swot[i][1], { x: x + 1.3, y: y + 0.3, w: 4, h: 0.45, fontFace: "Cambria", fontSize: 20, bold: true, color: C.text1, margin: 0, isTextBox: true, objectName: "swot head " + (i + 1) });
    s.addText(swot[i][4].map((t, k) => T(t, { bullet: true, breakLine: k < swot[i][4].length - 1 })),
      { x: x + 1.3, y: y + 0.82, w: w - 1.5, h: 1.3, fontSize: 14, color: C.text1, margin: 0, valign: "top", paraSpaceAfter: 4, isTextBox: true, objectName: "swot list " + (i + 1) });
  }
  s.addText([T("How we respond: ", { bold: true, color: C.accent3 }), T("a constant flow of new products and cost advantages (wages, energy, a faster and more flexible organisation)")],
    { x: MX, y: 6.62, w: 12.1, h: 0.3, fontSize: 12, color: C.text1, margin: 0, isTextBox: true, objectName: "swot response" });

  // ======================= 27. GROWTH LEVERS =======================
  s = pres.addSlide({ masterName: "CONTENT_DARK", sectionTitle: "Competition" });
  heads(s, "05 · Growth levers", "Two breakout opportunities worth HUF 3.3bn");
  const lev = [
    ["FaIndustry", "HUF 1.2bn", "Contract manufacturing", "We already produce for several Hungarian and international companies. The new plant lets us scale this into a stable second income stream that fills capacity."],
    ["FaGlobe", "HUF 2.1bn", "Export", "Western Europe, North America and the Middle East are the largest expected markets. We sell in EUR, which limits currency exposure."],
  ];
  for (let i = 0; i < 2; i++) {
    const x = MX + i * 6.15, y = 1.9, w = 5.95, h = 4.6;
    card(s, x, y, w, h, HEX.darkCard, "lever card " + (i + 1));
    await badge(s, x + 0.4, y + 0.4, 0.95, lev[i][0], i ? C.accent1 : C.accent2, HEX.icon, "lever badge " + (i + 1));
    s.addText(lev[i][1], { x: x + 0.4, y: y + 1.6, w: w - 0.8, h: 1.0, fontFace: "Cambria", fontSize: 50, bold: true, color: i ? C.accent3 : HEX.warmTxt, margin: 0, isTextBox: true, objectName: "lever stat " + (i + 1) });
    s.addText(lev[i][2], { x: x + 0.4, y: y + 2.6, w: w - 0.8, h: 0.45, fontSize: 20, bold: true, color: C.text1, margin: 0, isTextBox: true, objectName: "lever head " + (i + 1) });
    s.addText(lev[i][3], { x: x + 0.4, y: y + 3.1, w: w - 0.8, h: 1.3, fontSize: 14, color: HEX.muteOnDark, margin: 0, valign: "top", isTextBox: true, objectName: "lever text " + (i + 1) });
  }

  // ======================= 28. THE ASK =======================
  pres.addSection({ title: "The investment" });
  s = pres.addSlide({ masterName: "CONTENT_LIGHT", sectionTitle: "The investment" });
  heads(s, "06 · The ask", "HUF 800M–1bn to build our Röszke plant");
  const ix = 7.0, iw = W - MX - 7.0, ih = 3.25;
  s.addImage({ data: await photo(ai("plant.jpg", "image60.jpg"), iw, ih), x: ix, y: 1.75, w: iw, h: ih, objectName: "Röszke plant concept image", altText: "Concept visualisation of the planned Röszke plant" });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: ix + 0.15, y: 1.75 + ih - 0.43, w: 2.75, h: 0.28, rectRadius: 0.05, fill: { color: C.background1, transparency: 15 }, line: { type: "none" }, objectName: "concept label bg" });
  s.addText("Röszke plant · concept visualisation", { x: ix + 0.15, y: 1.75 + ih - 0.43, w: 2.75, h: 0.28, fontSize: 10, color: C.text1, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: "concept label" });
  s.addText("HUF 800M–1bn", { x: MX, y: 1.7, w: 6.2, h: 0.95, fontFace: "Cambria", fontSize: 48, bold: true, color: HEX.warmTxt, margin: 0, isTextBox: true, objectName: "ask amount" });
  s.addText("Investment sought. 100% goes into the new production plant.", { x: MX, y: 2.65, w: 6.2, h: 0.4, fontSize: 16, color: C.text1, margin: 0, isTextBox: true, objectName: "ask subtitle" });
  const askItems = [
    [HEX.sage, "Greenfield plant · ~HUF 700M", "Relocation to Röszke, 5 minutes from today's site: concrete base, road, fence, utilities and the production hall"],
    [HEX.gold, "Production technology · HUF 100–300M", "New machines and accessories, line installation and commissioning in the new hall"],
  ];
  for (let i = 0; i < 2; i++) {
    const y = 3.35 + i * 0.95;
    s.addShape(pres.shapes.OVAL, { x: MX, y: y + 0.06, w: 0.3, h: 0.3, fill: { color: askItems[i][0] }, line: { type: "none" }, objectName: "ask dot " + (i + 1) });
    s.addText([T(askItems[i][1], { bold: true, fontSize: 16, color: C.text1, breakLine: true }), T(askItems[i][2], { fontSize: 13, color: C.accent5 })],
      { x: MX + 0.5, y, w: 5.75, h: 0.85, margin: 0, valign: "top", isTextBox: true, objectName: "ask item " + (i + 1) });
  }
  // scale bar 0 → 1bn
  const bx = MX, bw = 12.13, by = 5.45, bh = 0.42, sc = bw / 1000;
  s.addShape(pres.shapes.RECTANGLE, { x: bx, y: by, w: 700 * sc, h: bh, fill: { color: HEX.sage }, line: { color: C.background1, width: 1 }, objectName: "bar plant" });
  s.addShape(pres.shapes.RECTANGLE, { x: bx + 700 * sc, y: by, w: 100 * sc, h: bh, fill: { color: HEX.gold }, line: { color: C.background1, width: 1 }, objectName: "bar technology min" });
  s.addShape(pres.shapes.RECTANGLE, { x: bx + 800 * sc, y: by, w: 200 * sc, h: bh, fill: { color: HEX.gold, transparency: 55 }, line: { color: HEX.gold, width: 1, dashType: "dash" }, objectName: "bar technology range" });
  s.addText("Greenfield plant", { x: bx + 0.15, y: by, w: 4, h: bh, fontSize: 13, bold: true, color: C.text1, valign: "middle", margin: 0, isTextBox: true, objectName: "bar label plant" });
  s.addText("Technology", { x: bx + 700 * sc, y: by, w: 300 * sc, h: bh, fontSize: 13, bold: true, color: C.text1, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: "bar label tech" });
  [[0, "0"], [700, "~700M"], [800, "800M"], [1000, "1bn"]].forEach(([v, t], k) => {
    const x = bx + v * sc;
    s.addText(t, { x: k === 3 ? x - 1.0 : (k === 0 ? x : x - 0.5), y: by + bh + 0.06, w: 1.0, h: 0.28, fontSize: 11, color: C.accent5, align: k === 3 ? "right" : (k === 0 ? "left" : "center"), margin: 0, isTextBox: true, objectName: "bar tick " + t });
  });
  s.addText("Separately, our winning ÉLIP grant co-finances further capacity expansion.", { x: MX, y: 6.55, w: 12, h: 0.3, fontSize: 11, italic: true, color: C.accent5, margin: 0, isTextBox: true, objectName: "grant note" });
  s.addNotes("The ask is HUF 800M to 1bn, used entirely for the plant: about HUF 700M for the greenfield site and hall, the balance for production technology.");

  // ======================= 29. CAPACITY =======================
  s = pres.addSlide({ masterName: "CONTENT_LIGHT", sectionTitle: "The investment" });
  heads(s, "06 · Capacity", "The capacity build-out multiplies output");
  s.addImage({ data: await photo("image60.jpg", 4.6, 2.35), x: W - MX - 4.6, y: 1.75, w: 4.6, h: 2.35, objectName: "production hall photo", altText: "Production hall" });
  s.addImage({ data: await photo("image6.jpg", 4.6, 2.35), x: W - MX - 4.6, y: 4.25, w: 4.6, h: 2.35, objectName: "milling equipment photo", altText: "Milling equipment" });
  const capa = [
    ["Flour milling", "270", "1,270", "kg / hour", "4.7×", "A third, 1,000 kg/h mill joins the two existing mills"],
    ["Flour blending", "750", "2,550", "kg / hour", "3.4×", "A 550 L high-capacity mixer for Bake-Free"],
    ["Instant cups", "1 line", "2 lines + filler", "automatic", "2×+", "A second automatic packer and an automatic cup filler"],
  ];
  for (let i = 0; i < 3; i++) {
    const y = 1.8 + i * 1.6, x = MX, w = 7.2;
    shadowCard(s, x, y, w, 1.42, "capacity card " + (i + 1));
    s.addText([T(capa[i][0], { bold: true, fontSize: 16, color: C.text1, breakLine: true }), T(capa[i][5], { fontSize: 12, color: C.accent5 })],
      { x: x + 0.25, y: y + 0.15, w: 2.6, h: 1.15, margin: 0, valign: "middle", isTextBox: true, objectName: "capacity label " + (i + 1) });
    s.addText([T(capa[i][1], { fontSize: 18, color: C.accent5, bold: true }), T("   →   ", { fontSize: 16, color: HEX.warmTxt, bold: true }), T(capa[i][2], { fontSize: 18, color: C.text1, bold: true, breakLine: true }), T(capa[i][3], { fontSize: 11, color: C.accent5 })],
      { x: x + 2.95, y: y + 0.15, w: 2.6, h: 1.15, margin: 0, valign: "middle", isTextBox: true, objectName: "capacity values " + (i + 1) });
    s.addText(capa[i][4], { x: x + 5.6, y: y + 0.15, w: 1.45, h: 1.15, fontFace: "Cambria", fontSize: 30, bold: true, color: C.accent3, align: "right", margin: 0, valign: "middle", isTextBox: true, objectName: "capacity multiple " + (i + 1) });
  }
  s.addText("The new plant is planned to triple production capacity within two years, towards ~10M units a year.", { x: MX, y: 6.62, w: 7.2, h: 0.3, fontSize: 11, italic: true, color: C.accent5, margin: 0, isTextBox: true, objectName: "capacity note" });

  // ======================= 30. ROADMAP =======================
  s = pres.addSlide({ masterName: "CONTENT_LIGHT", sectionTitle: "The investment" });
  heads(s, "06 · Roadmap", "From the Szeged plant to an international exit");
  const road = [
    ["2025", "Foundations", ["First KSA exports", "ÉLIP grant won", "NOPEX London"]],
    ["2026", "Build", ["Röszke greenfield works", "New Cupster pack (HUF 599)", "Oosty 'meat' mixes"]],
    ["2027", "Scale", ["Plant ramp-up", "Multinational & discount chains", "KSA up to €1.5M"]],
    ["2028–29", "Expand", ["Western Europe B2B + DTC", "USA, Canada, GCC", "Contract manufacturing at scale"]],
    ["2030", "Exit", ["Sale to a financial or strategic investor", "3–5× target"]],
  ];
  const rw = 2.3, rg = 0.15, ry = 2.55;
  s.addShape(pres.shapes.LINE, { x: MX + 0.3, y: ry + 0.3, w: 11.5, h: 0, line: { color: HEX.line, width: 3 }, objectName: "roadmap line" });
  for (let i = 0; i < 5; i++) {
    const x = MX + i * (rw + rg);
    numBadge(s, x + 0.05, ry, 0.6, i + 1, i === 4 ? C.accent2 : C.accent1, C.text2);
    s.addText(road[i][0], { x, y: 1.75, w: rw, h: 0.6, fontFace: "Cambria", fontSize: 26, bold: true, color: i === 4 ? HEX.warmTxt : C.accent3, margin: 0, isTextBox: true, objectName: "roadmap year " + (i + 1) });
    shadowCard(s, x, 3.45, rw, 2.75, "roadmap card " + (i + 1));
    s.addText(road[i][1], { x: x + 0.2, y: 3.6, w: rw - 0.4, h: 0.45, fontSize: 17, bold: true, color: C.text1, margin: 0, isTextBox: true, objectName: "roadmap phase " + (i + 1) });
    s.addText(road[i][2].map((t, k) => T(t, { bullet: true, breakLine: k < road[i][2].length - 1 })),
      { x: x + 0.2, y: 4.1, w: rw - 0.35, h: 2.3, fontSize: 13, color: C.accent5, margin: 0, valign: "top", paraSpaceAfter: 6, isTextBox: true, objectName: "roadmap items " + (i + 1) });
  }

  // ======================= 31. FINANCIAL PLAN =======================
  s = pres.addSlide({ masterName: "CONTENT_LIGHT", sectionTitle: "The investment" });
  heads(s, "06 · Financial plan", "Revenue to HUF 2.5bn by 2030, at a ~20% EBITDA margin");
  const yrs = ["2025A", "2026P", "2027P", "2028P", "2029P", "2030P"];
  s.addChart([
    { type: pres.charts.BAR, data: [
      { name: "Net revenue (HUF M)", labels: yrs, values: [732.0, 873.4, 1121.5, 1468.4, 1915.5, 2514.1] },
      { name: "EBITDA (HUF M)", labels: yrs, values: [181.5, 201.4, 208.3, 320.1, 377.2, 524.8] },
    ], options: { chartColors: [HEX.sage, HEX.gold], barGapWidthPct: 45 } },
    { type: pres.charts.LINE, data: [{ name: "EBITDA margin", labels: yrs, values: [0.248, 0.231, 0.186, 0.218, 0.197, 0.209] }],
      options: { chartColors: [HEX.green], lineSize: 3, lineDataSymbol: "circle", lineDataSymbolSize: 8, secondaryValAxis: true, secondaryCatAxis: true } },
  ], {
    x: MX, y: 1.7, w: 8.4, h: 4.95, showLegend: true, legendPos: "t", ...chartBase,
    valAxes: [{ showValAxisTitle: false, valAxisMinVal: 0, valAxisMaxVal: 3000, valAxisLabelFormatCode: "#,##0", valAxisLabelColor: HEX.grey, valGridLine: { color: "E3EAE8", size: 0.75 } },
      { showValAxisTitle: false, valAxisMinVal: 0, valAxisMaxVal: 0.5, valAxisLabelFormatCode: "0%", valAxisLabelColor: HEX.grey, valGridLine: { style: "none" } }],
    catAxes: [{ catAxisTitle: "" }, { catAxisHidden: true }],
    showValue: false, objectName: "financial plan chart",
  });
  const fin = [["3.4×", "revenue growth, 2025 → 2030"], ["~28%", "revenue CAGR, 2025–2030"], ["HUF 525M", "planned EBITDA in 2030"]];
  for (let i = 0; i < 3; i++) {
    const y = 1.8 + i * 1.6;
    card(s, 9.3, y, 3.43, 1.42, C.background2, "fin card " + (i + 1));
    s.addText([T(fin[i][0], { fontFace: "Cambria", fontSize: 30, bold: true, color: C.accent3, breakLine: true }), T(fin[i][1], { fontSize: 13, color: C.text1 })],
      { x: 9.5, y: y + 0.12, w: 3.1, h: 1.2, margin: 0, valign: "middle", isTextBox: true, objectName: "fin stat " + (i + 1) });
  }
  source(s, "Company financial plan. A = actual, P = plan. Revenue: 732 → 873 → 1,122 → 1,468 → 1,916 → 2,514. EBITDA: 182 → 201 → 208 → 320 → 377 → 525 (HUF M).");
  s.addNotes("The 2027 margin dip reflects the plant ramp-up. From 2028 operating leverage brings the margin back above 20%.");

  // ======================= 32. EXIT =======================
  s = pres.addSlide({ masterName: "CONTENT_DARK", sectionTitle: "The investment" });
  heads(s, "06 · Exit", "A clear path to a 2030 exit");
  s.addText("3–5×", { x: MX, y: 1.9, w: 5.4, h: 1.9, fontFace: "Cambria", fontSize: 110, bold: true, color: HEX.warmTxt, margin: 0, isTextBox: true, objectName: "exit multiple" });
  s.addText("Target return to investors in 5 years", { x: MX, y: 3.85, w: 5.4, h: 0.5, fontSize: 20, color: C.text1, margin: 0, isTextBox: true, objectName: "exit multiple label" });
  s.addText("The value drivers are scale from the Röszke plant, export share, high-margin DTC, Cupster and Oosty growth, and contract manufacturing.", { x: MX, y: 4.5, w: 5.2, h: 1.3, fontSize: 14, color: HEX.muteOnDark, margin: 0, valign: "top", isTextBox: true, objectName: "exit drivers" });
  const ex = [
    ["FaLandmark", "Financial investor", "A sale to an international financial investor such as a PE fund or food-platform roll-up."],
    ["FaHandshake", "Strategic investor", "Hungarian or international food groups, once we reach a larger company value. Precedent: Hunorganic was acquired by Foodnet (Alpro)."],
  ];
  for (let i = 0; i < 2; i++) {
    const y = 1.9 + i * 2.35;
    card(s, 6.6, y, 6.13, 2.1, HEX.darkCard, "exit card " + (i + 1));
    await badge(s, 6.9, y + 0.35, 0.8, ex[i][0], C.accent1, HEX.icon, "exit badge " + (i + 1));
    s.addText([T(ex[i][1], { bold: true, fontSize: 19, color: C.text1, breakLine: true }), T(ex[i][2], { fontSize: 14, color: HEX.muteOnDark })],
      { x: 7.95, y: y + 0.25, w: 4.6, h: 1.65, margin: 0, valign: "top", isTextBox: true, objectName: "exit text " + (i + 1) });
  }

  // ======================= 33. WHY INVEST =======================
  s = pres.addSlide({ masterName: "CONTENT_LIGHT", sectionTitle: "The investment" });
  heads(s, "Summary", "Why invest in Naturtrade");
  s.addImage({ data: await photo("image61.jpg", 4.3, 4.85), x: W - MX - 4.3, y: 1.75, w: 4.3, h: 4.85, objectName: "summary product photo", altText: "Bake-Free cake" });
  const wi = [
    ["FaWheatAwn", "Real moat", "Our own milling plus full in-house production, a chain competitors can't copy quickly"],
    ["FaChartLine", "Growing markets", "Gluten-free +7–10% a year, plant-based meat +15%, GCC demand rising"],
    ["FaLightbulb", "Category creators", "Cupster and Oosty open new shelf space, not just compete for it"],
    ["FaGlobe", "Proven export", "UK, Western Europe, Amazon and a state-backed KSA partnership"],
    ["FaRocket", "Capital-efficient plan", "HUF 800M–1bn builds the plant, with a 3–5× target by 2030"],
  ];
  for (let i = 0; i < 5; i++) {
    const y = 1.8 + i * 0.97;
    await badge(s, MX, y + 0.05, 0.65, wi[i][0], i % 2 ? C.accent2 : C.accent1, HEX.icon, "invest badge " + (i + 1));
    s.addText([T(wi[i][1], { bold: true, fontSize: 16, color: C.text1, breakLine: true }), T(wi[i][2], { fontSize: 13, color: C.accent5 })],
      { x: MX + 0.9, y, w: 6.6, h: 0.85, margin: 0, valign: "middle", isTextBox: true, objectName: "invest reason " + (i + 1) });
  }

  // ======================= 34. THANK YOU =======================
  s = pres.addSlide({ masterName: "COVER_DARK", sectionTitle: "The investment" });
  s.addImage({ data: await photo(ai("handshake.jpg", "image26.jpg"), 5.6, 7.5), x: W - 5.6, y: 0, w: 5.6, h: 7.5, objectName: "closing celebration image", altText: "Business partners shaking hands while colleagues celebrate" });
  s.addText("THANK YOU", { x: MX + 0.1, y: 1.4, w: 6.8, h: 0.35, fontSize: 13, bold: true, color: C.accent3, charSpacing: 4, margin: 0, isTextBox: true, objectName: "closing eyebrow" });
  s.addText("Let's grow it together.", { x: MX + 0.1, y: 1.85, w: 7, h: 1.6, fontFace: "Cambria", fontSize: 48, bold: true, color: C.text1, margin: 0, valign: "top", isTextBox: true, objectName: "closing title" });
  s.addText([
    T("Miklós Véha", { bold: true, color: C.text1, fontSize: 18, breakLine: true }),
    T("Founder & CEO, Naturtrade Hungary Kft.", { color: HEX.muteOnDark, fontSize: 14, breakLine: true }),
    T(" ", { fontSize: 8, breakLine: true }),
    T("miklos.veha@naturtrade.hu", { color: HEX.warmTxt, fontSize: 15, breakLine: true }),
    T("+36 30 637 4535", { color: HEX.warmTxt, fontSize: 15, breakLine: true }),
    T("millandfolks.com  ·  edenpremium.hu", { color: HEX.muteOnDark, fontSize: 14 }),
  ], { x: MX + 0.1, y: 3.7, w: 6.5, h: 2.2, margin: 0, valign: "top", isTextBox: true, objectName: "contact" });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: MX + 0.1, y: 6.0, w: 0.95, h: 0.95, rectRadius: 0.08, fill: { color: C.background1 }, line: { type: "none" }, objectName: "closing logo tile" });
  const lg2 = await logo("image1.png", 0.78, 0.78);
  s.addImage({ data: lg2.data, x: MX + 0.1 + (0.95 - lg2.w) / 2, y: 6.0 + (0.95 - lg2.h) / 2, w: lg2.w, h: lg2.h, objectName: "closing logo", altText: "Naturtrade logo" });

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
}
main().catch((e) => { console.error(e); process.exit(1); });
