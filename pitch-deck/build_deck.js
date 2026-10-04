// RepairOS — CodeSlayer 2K26 submission deck (6 slides, organiser template title bar).
// Build:  npm install && node build_deck.js   (writes RepairOS_CodeSlayer2K26.pptx next to this file)
// Edit TEAM below, rebuild, then export the PPTX to PDF for submission.

const path = require("path");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const Fi = require("react-icons/fi");
const Bs = require("react-icons/bs");
const Fa = require("react-icons/fa");

// ---------------------------------------------------------------- team details
const TEAM = {
  name: "Team Name",
  members: [
    { name: "Member 1", role: "Team lead · product, decision logic & pitch" },
    { name: "Member 2", role: "Frontend · Next.js PWA, self-tests, Passport UI" },
    { name: "Member 3", role: "Backend · FastAPI, Bayesian engine, data model" },
    { name: "Member 4", role: "AI/ML · LLM extraction, OCR/vision, RAG" },
  ],
};

const OUT = process.argv[2] || path.join(__dirname, "RepairOS_CodeSlayer2K26.pptx");
const ASSET = (f) => path.join(__dirname, "assets", f);

// ---------------------------------------------------------------- theme
const HEX = {
  ink: "141821", slate: "1B2430", light: "F4F5F7", crimson: "B3122E", green: "1E8E5A",
  amber: "C77C02", muted: "5B6472", border: "DDE1E6", tint: "FBEAED", white: "FFFFFF",
};
const THEME = {
  name: "RepairOS CodeSlayer",
  headFontFace: "Georgia",
  bodyFontFace: "Arial",
  colors: {
    dk1: HEX.ink, lt1: HEX.white, dk2: HEX.slate, lt2: HEX.light,
    accent1: HEX.crimson, accent2: HEX.green, accent3: HEX.amber, accent4: HEX.muted,
    accent5: HEX.border, accent6: HEX.tint, hlink: HEX.crimson, folHlink: "7A0C1F",
  },
};
const SERIF = "Georgia";

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5 in — same canvas as the organiser template
pres.title = "RepairOS — Diagnose Before You Replace";
pres.subject = "CodeSlayer 2K26 submission";
pres.author = TEAM.name;
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };

const S = pres.SchemeColor;
const C = {
  ink: S.text1, white: S.background1, slate: S.text2, light: S.background2,
  crimson: S.accent1, green: S.accent2, amber: S.accent3, muted: S.accent4, border: S.accent5, tint: S.accent6,
};

// ---------------------------------------------------------------- icon helper
async function icon(Comp, hex, size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Comp, { color: "#" + hex, size }));
  const buf = await sharp(Buffer.from(svg)).resize(size, size).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

// ---------------------------------------------------------------- layouts
const LOGOS_CONTENT = [
  { image: { x: 10.6, y: -0.12, w: 1.25, h: 1.21, path: ASSET("devsphere-logo.png") } },
  { text: { text: "DevSphereIndia", options: { x: 10.33, y: 0.8, w: 1.8, h: 0.26, fontSize: 11, bold: true, color: C.ink, align: "center", margin: 0 } } },
  { image: { x: 11.88, y: -0.06, w: 1.14, h: 1.14, path: ASSET("codeslayer-logo.png") } },
];
const OVAL = (x, y) => [
  { text: { text: TEAM.name, options: { shape: pres.shapes.OVAL, x, y, w: 2.45, h: 0.93, line: { color: "595959", width: 0.75 }, fontSize: 17, color: C.ink, align: "center", valign: "middle", margin: 4 } } },
];

pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: C.white },
  objects: [
    ...LOGOS_CONTENT,
    ...OVAL(0.24, 1.02),
    {
      placeholder: {
        options: { name: "title", type: "title", x: 2.75, y: 0.16, w: 7.75, h: 1.08, fontFace: SERIF, fontSize: 27, bold: true, color: C.ink, align: "center", valign: "middle", margin: 0 },
        text: "",
      },
    },
  ],
});

pres.defineSlideMaster({
  title: "COVER",
  background: { color: C.white },
  objects: [
    { image: { x: -0.08, y: -0.22, w: 13.49, h: 4.49, path: ASSET("banner.png") } },
    { image: { x: 10.46, y: 0.01, w: 1.25, h: 1.21, path: ASSET("devsphere-logo.png") } },
    { text: { text: "DevSphereIndia", options: { x: 10.19, y: 0.97, w: 1.8, h: 0.24, fontSize: 11, bold: true, color: C.ink, align: "center", margin: 0 } } },
    { image: { x: 11.86, y: 0.1, w: 1.16, h: 1.16, path: ASSET("codeslayer-logo.png") } },
    ...OVAL(0.24, 1.18),
    { image: { x: 4.533, y: 1.664, w: 4.26, h: 0.447, path: ASSET("codeslayer-wordmark.png") } },
  ],
});

// ---------------------------------------------------------------- drawing helpers
function panel(slide, x, y, w, h, name, fill = C.light) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h, rectRadius: 0.07, fill: { color: fill }, line: { color: C.border, width: 0.75 }, objectName: name,
  });
}

function sectionHead(slide, x, y, ico, text, w = 5) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 0.4, h: 0.4, rectRadius: 0.08, fill: { color: C.tint }, line: { color: C.tint } });
  slide.addImage({ data: ico, x: x + 0.09, y: y + 0.09, w: 0.22, h: 0.22 });
  slide.addText(text, { x: x + 0.52, y, w, h: 0.4, fontSize: 18, bold: true, color: C.ink, valign: "middle", margin: 0, isTextBox: true, objectName: "Section: " + text });
}

function lede(slide, runs) {
  slide.addText(runs, { x: 3.0, y: 1.1, w: 9.85, h: 0.77, fontSize: 15, color: C.ink, valign: "middle", margin: 0, isTextBox: true, objectName: "Key message" });
}

function footer(slide, runs) {
  slide.addText(runs, { x: 0.5, y: 6.99, w: 12.35, h: 0.42, fontSize: 9, color: C.muted, valign: "top", margin: 0, isTextBox: true, objectName: "Sources" });
}

function pill(slide, x, y, w, text, fill) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h: 0.26, rectRadius: 0.13, fill: { color: fill }, line: { color: fill } });
  slide.addText(text, { x, y, w, h: 0.26, fontSize: 8.5, bold: true, color: C.white, align: "center", valign: "middle", margin: 0, charSpacing: 1, isTextBox: true });
}

function arrow(slide, x1, y1, x2, y2, opts = {}) {
  slide.addShape(pres.shapes.LINE, {
    x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(x2 - x1), h: Math.abs(y2 - y1),
    flipH: x2 < x1, flipV: y2 < y1,
    line: { color: opts.color || C.muted, width: opts.width || 1.25, endArrowType: "triangle", beginArrowType: opts.both ? "triangle" : undefined },
  });
}

const sup = (n) => ({ text: String(n), options: { superscript: true } });

// pptxgenjs writes Office's stock palette into the theme; swap in ours so scheme colors resolve correctly.
async function applyThemeColors(file, theme) {
  const fs = require("fs");
  const JSZip = require(require.resolve("jszip", { paths: [require.resolve("pptxgenjs")] }));
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  const part = "ppt/theme/theme1.xml";
  const slots = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];
  const scheme = `<a:clrScheme name="${theme.name}">` + slots.map((k) => `<a:${k}><a:srgbClr val="${theme.colors[k]}"/></a:${k}>`).join("") + "</a:clrScheme>";
  const xml = (await zip.file(part).async("string"))
    .replace(/<a:clrScheme\b[\s\S]*?<\/a:clrScheme>/, () => scheme)
    .replace(/(<a:(?:theme|fontScheme)\b[^>]*?\bname=")[^"]*"/g, (_, head) => `${head}${theme.name}"`);
  zip.file(part, xml);
  fs.writeFileSync(file, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
}

// ================================================================= build
(async () => {
  const I = {
    problem: await icon(Fi.FiAlertTriangle, HEX.crimson),
    users: await icon(Fi.FiUsers, HEX.crimson),
    features: await icon(Fi.FiZap, HEX.crimson),
    different: await icon(Fi.FiAward, HEX.crimson),
    stack: await icon(Fi.FiLayers, HEX.crimson),
    arch: await icon(Fi.FiShare2, HEX.crimson),
    feasible: await icon(Fi.FiCheckCircle, HEX.crimson),
    shield: await icon(Fi.FiShield, HEX.crimson),
    usp: await icon(Fi.FiStar, HEX.crimson),
    biz: await icon(Fi.FiTrendingUp, HEX.crimson),
    identify: await icon(Fi.FiTag, HEX.white),
    evidence: await icon(Fi.FiActivity, HEX.white),
    diagnose: await icon(Fi.FiCpu, HEX.white),
    decide: await icon(Fa.FaBalanceScale, HEX.white),
    passport: await icon(Fi.FiBookOpen, HEX.white),
    full: await icon(Bs.BsCircleFill, HEX.ink),
    fullRed: await icon(Bs.BsCircleFill, HEX.crimson),
    half: await icon(Bs.BsCircleHalf, HEX.ink),
    none: await icon(Bs.BsCircle, "9AA1AC"),
    check: await icon(Fi.FiCheck, HEX.green),
    qr: await icon(Bs.BsQrCode, HEX.ink),
    loop: await icon(Fi.FiRefreshCw, HEX.white),
    dot: await icon(Bs.BsCircleFill, HEX.crimson),
  };

  // =============================================================== 1. COVER
  pres.addSection({ title: "Title" });
  {
    const s = pres.addSlide({ masterName: "COVER", sectionTitle: "Title" });
    const label = (text, y) => s.addText(text, { x: 0.75, y, w: 2.4, h: 0.42, fontSize: 20, bold: true, color: C.ink, valign: "middle", margin: 0, isTextBox: true });

    label("Project Title :", 2.62);
    s.addText(
      [
        { text: "RepairOS", options: { fontFace: SERIF, fontSize: 28, bold: true, color: C.crimson } },
        { text: "   Diagnose Before You Replace", options: { fontSize: 16, italic: true, color: C.muted } },
      ],
      { x: 2.95, y: 2.53, w: 6.85, h: 0.6, valign: "middle", margin: 0, isTextBox: true, objectName: "Project title" }
    );
    s.addText(
      "An evidence-first AI platform that tells phone and laptop owners what failed, what a fair fix should cost, and whether to repair, service or replace.",
      { x: 2.95, y: 3.1, w: 6.7, h: 0.58, fontSize: 12.5, color: C.ink, valign: "top", margin: 0, isTextBox: true, objectName: "One-line pitch" }
    );

    label("Team Name :", 3.86);
    s.addText(TEAM.name, { x: 2.95, y: 3.86, w: 6.5, h: 0.42, fontSize: 18, color: C.ink, valign: "middle", margin: 0, isTextBox: true, objectName: "Team name" });

    s.addText(
      [
        { text: "Team Members ", options: { fontSize: 20, bold: true } },
        { text: "(Names + Roles Briefly)", options: { fontSize: 15, italic: true } },
        { text: ":", options: { fontSize: 20, bold: true } },
      ],
      { x: 0.75, y: 4.46, w: 7.5, h: 0.42, color: C.ink, valign: "middle", margin: 0, isTextBox: true }
    );
    TEAM.members.forEach((m, i) => {
      const x = 0.98, y = 4.95 + i * 0.31;
      s.addImage({ data: I.dot, x: x - 0.2, y: y + 0.105, w: 0.09, h: 0.09 });
      s.addText(
        [
          { text: m.name, options: { bold: true, fontSize: 13, color: C.ink } },
          { text: "  —  " + m.role, options: { fontSize: 11.5, color: C.muted } },
        ],
        { x, y, w: 8.6, h: 0.3, valign: "middle", margin: 0, isTextBox: true, objectName: "Member " + (i + 1) }
      );
    });

    label("Tracks Chosen:", 6.3);
    s.addText(
      [
        { text: "Open Innovation", options: { fontSize: 18, bold: true, color: C.crimson } },
        { text: "   with an AI/ML core and a Sustainability outcome", options: { fontSize: 13, color: C.muted } },
      ],
      { x: 3.3, y: 6.3, w: 6.6, h: 0.42, valign: "middle", margin: 0, isTextBox: true, objectName: "Track" }
    );

    // --- illustrative product screen (phone mock-up)
    const PX = 10.05, PY = 2.38, PW = 2.6, PH = 4.72;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: PX, y: PY, w: PW, h: PH, rectRadius: 0.32, fill: { color: C.slate }, line: { color: C.slate },
      shadow: { type: "outer", color: "000000", opacity: 0.22, blur: 14, offset: 5, angle: 90 }, objectName: "Phone frame",
    });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: PX + 0.1, y: PY + 0.1, w: PW - 0.2, h: PH - 0.2, rectRadius: 0.24, fill: { color: C.white }, line: { color: C.white } });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: PX + PW / 2 - 0.34, y: PY + 0.17, w: 0.68, h: 0.1, rectRadius: 0.05, fill: { color: C.slate }, line: { color: C.slate } });

    const cx = PX + 0.24, cw = PW - 0.48;
    const t = (text, x, y, w, h, o = {}) => s.addText(text, { x, y, w, h, margin: 0, valign: "middle", isTextBox: true, color: C.ink, fontSize: 8, ...o });
    t("RepairOS", cx, PY + 0.36, 1.2, 0.24, { fontFace: SERIF, bold: true, fontSize: 11.5, color: C.crimson });
    t("Report R-2417", cx + 1.0, PY + 0.36, cw - 1.0, 0.24, { fontSize: 7, color: C.muted, align: "right" });
    t("Dell Inspiron 15 3511", cx, PY + 0.66, cw, 0.2, { bold: true, fontSize: 9 });
    t("Laptop · 3.4 yrs old · out of warranty", cx, PY + 0.85, cw, 0.17, { fontSize: 7, color: C.muted });

    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: cx, y: PY + 1.1, w: cw, h: 0.4, rectRadius: 0.08, fill: { color: C.light }, line: { color: C.light } });
    t("“Shuts down the moment I unplug the charger.”", cx + 0.08, PY + 1.1, cw - 0.16, 0.4, { italic: true, fontSize: 7.5 });

    t("EVIDENCE", cx, PY + 1.6, cw, 0.16, { fontSize: 6.5, bold: true, color: C.muted, charSpacing: 1 });
    ["Battery report: 31% of design capacity", "Runs normally on AC power", "Photo check: no battery swelling"].forEach((e, i) => {
      s.addImage({ data: I.check, x: cx, y: PY + 1.8 + i * 0.18, w: 0.12, h: 0.12 });
      t(e, cx + 0.17, PY + 1.77 + i * 0.18, cw - 0.17, 0.18, { fontSize: 7 });
    });

    t("LIKELY CAUSE", cx, PY + 2.4, cw, 0.16, { fontSize: 6.5, bold: true, color: C.muted, charSpacing: 1 });
    [["Battery wear", 0.82], ["Charging circuit", 0.12], ["Adapter / cable", 0.06]].forEach(([n, v], i) => {
      const y = PY + 2.6 + i * 0.2;
      t(n, cx, y, 0.92, 0.17, { fontSize: 7 });
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: cx + 0.95, y: y + 0.055, w: 0.82, h: 0.07, rectRadius: 0.035, fill: { color: C.light }, line: { color: C.light } });
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: cx + 0.95, y: y + 0.055, w: Math.max(0.07, 0.82 * v), h: 0.07, rectRadius: 0.035, fill: { color: i === 0 ? C.crimson : C.muted }, line: { color: i === 0 ? C.crimson : C.muted } });
      t(Math.round(v * 100) + "%", cx + 1.8, y, cw - 1.8, 0.17, { fontSize: 7, bold: true, align: "right" });
    });

    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: cx, y: PY + 3.27, w: cw, h: 0.92, rectRadius: 0.08, fill: { color: "E6F3EC" }, line: { color: "BFE0CC" } });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: cx + 0.09, y: PY + 3.36, w: 0.6, h: 0.2, rectRadius: 0.1, fill: { color: C.green }, line: { color: C.green } });
    t("REPAIR", cx + 0.09, PY + 3.36, 0.6, 0.2, { fontSize: 6.5, bold: true, color: C.white, align: "center", charSpacing: 1 });
    t("Battery replacement", cx + 0.76, PY + 3.36, cw - 0.82, 0.2, { fontSize: 8, bold: true });
    t([{ text: "₹3,500–5,000", options: { bold: true } }, { text: "  vs  ₹45,000+ new laptop" }], cx + 0.09, PY + 3.62, cw - 0.18, 0.2, { fontSize: 7.5 });
    t("≈ ₹180 per extra month · repairability 8/10", cx + 0.09, PY + 3.85, cw - 0.18, 0.26, { fontSize: 6.5, color: C.muted });

    s.addImage({ data: I.qr, x: cx, y: PY + 4.27, w: 0.24, h: 0.24 });
    t("Saved to Repair Passport", cx + 0.32, PY + 4.25, cw - 0.32, 0.15, { fontSize: 7.5, bold: true });
    t("shareable · tamper-evident history", cx + 0.32, PY + 4.39, cw - 0.32, 0.14, { fontSize: 6.5, color: C.muted });

    s.addText("Illustrative prototype screen", { x: PX, y: PY + PH + 0.04, w: PW, h: 0.2, fontSize: 8, italic: true, color: C.muted, align: "center", margin: 0, isTextBox: true });

    s.addNotes(
      "RepairOS is the decision layer between a broken device and a replacement purchase. In one guided session it identifies the device, collects real evidence (including the device's own health data), ranks likely faults with confidence, and turns that into a costed repair / service / replace decision that is saved to a Repair Passport. The screen on the right is an illustrative prototype result for a laptop that shuts down when unplugged."
    );
  }

  // =============================================================== 2. PROBLEM
  pres.addSection({ title: "Problem" });
  {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Problem" });
    s.addText("PROBLEM STATEMENT &\nTARGET AUDIENCE", { placeholder: "title" });
    lede(s, [
      { text: "When a phone or laptop fails, owners " },
      { text: "decide blind", options: { bold: true, color: C.crimson } },
      { text: " —", options: { breakLine: true } },
      { text: "so fixable devices get replaced, and repairs get overpaid." },
    ]);

    // Problem panel
    panel(s, 0.5, 2.25, 8.3, 4.66, "Problem panel");
    sectionHead(s, 0.7, 2.4, I.problem, "Problem We're Solving");
    s.addText(
      "A device fault is an information problem before it is a technical one. Owners can't tell what failed, what a fair repair should cost, or whether the device is still worth fixing — and the people who can tell them usually profit from the answer.",
      { x: 0.7, y: 2.9, w: 7.9, h: 0.78, fontSize: 12.5, color: C.ink, valign: "top", margin: 0, isTextBox: true, objectName: "Problem definition" }
    );

    const stats = [
      ["43%", [{ text: "of Indian smartphone users who used a service centre found repair cost high to very high" }, sup(1)]],
      ["4 in 10", [{ text: "needed more than one service-centre visit to get a single issue resolved" }, sup(1)]],
      ["1.75 Mt", [{ text: "e-waste generated in India in FY24 — up 72.5% in five years" }, sup(2)]],
      ["72%", [{ text: "of a smartphone's climate impact comes from making, shipping and disposing of it" }, sup(3)]],
    ];
    stats.forEach(([num, label], i) => {
      const x = 0.7 + i * (1.86 + 0.153), y = 3.78;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 1.86, h: 1.5, rectRadius: 0.06, fill: { color: C.white }, line: { color: C.border, width: 0.75 }, objectName: "Stat " + num });
      s.addText(num, { x: x + 0.14, y: y + 0.1, w: 1.6, h: 0.5, fontSize: 26, bold: true, color: C.crimson, valign: "middle", margin: 0, isTextBox: true });
      s.addText(label, { x: x + 0.14, y: y + 0.62, w: 1.62, h: 0.82, fontSize: 10, color: C.ink, valign: "top", margin: 0, isTextBox: true });
    });

    s.addText("WHERE THE REPAIR DECISION BREAKS TODAY", { x: 0.7, y: 5.43, w: 7.9, h: 0.24, fontSize: 9, bold: true, color: C.muted, charSpacing: 1, margin: 0, isTextBox: true });
    const chain = ["Device fails", "Owner can't pin down the fault", "Opaque quote, often a full-module swap", "Replaced “to be safe”", "Avoidable cost and e-waste"];
    chain.forEach((c, i) => {
      const w = 1.42, x = 0.7 + i * (w + 0.2), y = 5.74, last = i === chain.length - 1;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h: 0.98, rectRadius: 0.06, fill: { color: last ? C.crimson : C.white }, line: { color: last ? C.crimson : C.border, width: 0.75 } });
      s.addText(
        [{ text: "0" + (i + 1), options: { fontSize: 8.5, bold: true, color: last ? C.white : C.crimson, breakLine: true } }, { text: c, options: { fontSize: 10.5, bold: true, color: last ? C.white : C.ink } }],
        { x: x + 0.1, y: y + 0.08, w: w - 0.2, h: 0.82, valign: "top", margin: 0, isTextBox: true, paraSpaceAfter: 2 }
      );
      if (!last) arrow(s, x + w + 0.03, y + 0.49, x + w + 0.17, y + 0.49);
    });

    // Target users panel
    panel(s, 9.05, 2.25, 3.8, 4.66, "Target users panel");
    sectionHead(s, 9.25, 2.4, I.users, "Target Users", 2.8);
    pill(s, 9.25, 2.93, 1.45, "B2C · PRIMARY", C.crimson);
    const b2c = [
      "Phone & laptop owners aged 18–40: students, young professionals, gig workers",
      "Out-of-warranty devices 2–5 years old — where repair-or-replace is decided",
      "Used-device buyers and sellers who need proof of condition",
    ];
    s.addText(b2c.map((t, i) => ({ text: t, options: { bullet: { indent: 12 }, breakLine: i < b2c.length - 1 } })), {
      x: 9.25, y: 3.25, w: 3.45, h: 1.35, fontSize: 10.5, color: C.ink, valign: "top", margin: 0, paraSpaceAfter: 4, isTextBox: true, objectName: "B2C segments",
    });
    pill(s, 9.25, 4.66, 1.45, "B2B · REVENUE", C.slate);
    const b2b = [
      "Independent repair shops: pre-triaged leads, digital job cards",
      "Refurbishers & device insurers: grading and claim triage",
      "Campus and SME IT teams: fleet repair-vs-replace calls",
    ];
    s.addText(b2b.map((t, i) => ({ text: t, options: { bullet: { indent: 12 }, breakLine: i < b2b.length - 1 } })), {
      x: 9.25, y: 4.98, w: 3.45, h: 1.12, fontSize: 10.5, color: C.ink, valign: "top", margin: 0, paraSpaceAfter: 4, isTextBox: true, objectName: "B2B segments",
    });
    s.addText("740M+", { x: 9.25, y: 6.2, w: 1.3, h: 0.55, fontSize: 22, bold: true, color: C.crimson, valign: "middle", margin: 0, isTextBox: true });
    s.addText([{ text: "active smartphones in India — the B2C base, before laptops" }, sup(4)], { x: 10.6, y: 6.2, w: 2.1, h: 0.55, fontSize: 9, color: C.muted, valign: "middle", margin: 0, isTextBox: true });

    footer(s, [
      sup(1), { text: " Counterpoint Research, India Smartphone After-Sales Service Consumer Study (Sep 2025)    " },
      sup(2), { text: " CPCB data, MoHUA written reply in Rajya Sabha (16 Dec 2024)", options: { breakLine: true } },
      sup(3), { text: " European Environmental Bureau / Coolproducts, “Coolproducts don't cost the Earth” (2019)    " },
      sup(4), { text: " Counterpoint Research, active smartphone installed base (2025)" },
    ]);
    s.addNotes(
      "Talk track: the failure is not the hard part — the decision is. Owners lack three answers: what failed, what a fair fix costs, and whether the device is worth fixing. Today those answers come from someone who profits from the outcome.\n\nSources:\n1. Counterpoint Research — Survey: 4 in 10 India smartphone consumers require repeat visits for repairs (15 Sep 2025): https://www.counterpointresearch.com/insights/survey-4-in-10-india-smartphone-consumers-require-repeat-visits%20for-repairs\n2. CPCB data via MoHUA written reply, Rajya Sabha, 16 Dec 2024 (1.751 MT in FY2023-24 vs 1.01 MT in FY2019-20): https://www.downtoearth.org.in/waste/indias-e-waste-surges-by-73-in-5-years\n3. European Environmental Bureau, Coolproducts don't cost the Earth (2019): https://eeb.org/wp-content/uploads/2019/09/Coolproducts-briefing.pdf\n4. Counterpoint Research, global active smartphone installed base 2025 (India >740M)."
    );
  }

  // =============================================================== 3. SOLUTION
  pres.addSection({ title: "Solution" });
  {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Solution" });
    s.addText("OUR UNIQUE SOLUTION", { placeholder: "title" });
    lede(s, [
      { text: "RepairOS turns a vague complaint into " },
      { text: "device evidence, a ranked diagnosis and a costed decision", options: { bold: true, color: C.crimson } },
      { text: " — in under five minutes." },
    ]);

    panel(s, 0.5, 2.25, 6.4, 4.66, "Key features panel");
    sectionHead(s, 0.7, 2.4, I.features, "Key Features");
    const feats = [
      [I.identify, "Identify the device", "Snap the model label or service tag; OCR pulls specs, device age and known failure patterns for that model."],
      [I.evidence, "Collect real evidence", "Adaptive English/Hinglish questions, photos, in-browser self-tests (touch, audio, camera, sensors) and OS health reports such as battery wear."],
      [I.diagnose, "Diagnose with confidence", "A Bayesian engine ranks likely faults with a confidence score and asks the single most informative next question."],
      [I.decide, "Decide in rupees", "Repairability score on the six factors used by India's Repairability Index, plus ₹ per extra month of life, fair-price band and quote check."],
      [I.passport, "Track in a Repair Passport", "A shareable, tamper-evident device history (QR) for resale, warranty claims and the next diagnosis."],
    ];
    s.addShape(pres.shapes.LINE, { x: 0.93, y: 3.12, w: 0, h: 3.1, line: { color: C.border, width: 1.5 } });
    feats.forEach(([ic, title, desc], i) => {
      const y = 2.93 + i * 0.78;
      s.addShape(pres.shapes.OVAL, { x: 0.72, y, w: 0.42, h: 0.42, fill: { color: C.crimson }, line: { color: C.white, width: 1.5 } });
      s.addImage({ data: ic, x: 0.83, y: y + 0.11, w: 0.2, h: 0.2 });
      s.addText(title, { x: 1.32, y: y - 0.02, w: 5.4, h: 0.26, fontSize: 12.5, bold: true, color: C.ink, valign: "middle", margin: 0, isTextBox: true });
      s.addText(desc, { x: 1.32, y: y + 0.25, w: 5.45, h: 0.46, fontSize: 10, color: C.muted, valign: "top", margin: 0, isTextBox: true });
    });

    panel(s, 7.15, 2.25, 5.7, 4.66, "Why different panel");
    sectionHead(s, 7.35, 2.4, I.different, "Why It's Different");
    const cols = ["AI\nchatbots", "DIY repair\nguides", "Repair /\ntrade-in apps", "OEM\napps", "RepairOS"];
    const rows = [
      ["Evidence from the device itself", [0, 0, 0, 2, 2]],
      ["Ranked diagnosis + confidence", [1, 1, 1, 1, 2]],
      ["Repair-vs-replace economics (₹)", [1, 0, 1, 0, 2]],
      ["Outcome-neutral advice", [2, 1, 0, 0, 2]],
      ["Persistent device history", [0, 0, 1, 1, 2]],
    ];
    const LX = 7.35, CX0 = 9.3, CW = 0.69, TY = 2.95, RH = 0.37;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: CX0 + 4 * CW + 0.02, y: TY - 0.02, w: CW - 0.04, h: 0.5 + rows.length * RH + 0.04, rectRadius: 0.06, fill: { color: C.tint }, line: { color: C.tint } });
    cols.forEach((c, j) => {
      s.addText(c, { x: CX0 + j * CW, y: TY, w: CW, h: 0.46, fontSize: 8, bold: true, color: j === 4 ? C.crimson : C.ink, align: "center", valign: "bottom", margin: 0, isTextBox: true });
    });
    rows.forEach(([label, v], r) => {
      const y = TY + 0.5 + r * RH;
      s.addShape(pres.shapes.LINE, { x: LX, y, w: 5.35, h: 0, line: { color: C.border, width: 0.75 } });
      s.addText(label, { x: LX, y, w: 1.95, h: RH, fontSize: 9.5, color: C.ink, valign: "middle", margin: 0, isTextBox: true });
      v.forEach((k, j) => {
        const ic = k === 2 ? (j === 4 ? I.fullRed : I.full) : k === 1 ? I.half : I.none;
        s.addImage({ data: ic, x: CX0 + j * CW + CW / 2 - 0.09, y: y + RH / 2 - 0.09, w: 0.18, h: 0.18 });
      });
    });
    const ly = TY + 0.56 + rows.length * RH;
    [[I.full, "Full"], [I.half, "Partial"], [I.none, "None"]].forEach(([ic, t], i) => {
      s.addImage({ data: ic, x: LX + i * 0.85, y: ly + 0.05, w: 0.13, h: 0.13 });
      s.addText(t, { x: LX + i * 0.85 + 0.18, y: ly, w: 0.65, h: 0.23, fontSize: 8.5, color: C.muted, valign: "middle", margin: 0, isTextBox: true });
    });

    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 7.35, y: 5.62, w: 5.3, h: 1.17, rectRadius: 0.06, fill: { color: C.slate }, line: { color: C.slate }, objectName: "Novelty card" });
    s.addText(
      [
        { text: "NOVELTY FACTOR", options: { fontSize: 8.5, bold: true, color: "F29AAA", charSpacing: 1, breakLine: true } },
        { text: "Evidence-first: ", options: { bold: true } },
        { text: "reads the device's own health data, not just the user's words.", options: { breakLine: true } },
        { text: "Decision-first: ", options: { bold: true } },
        { text: "answers “what should I do?” in rupees and months — then remembers it." },
      ],
      { x: 7.52, y: 5.7, w: 5.0, h: 1.02, fontSize: 10.5, color: C.white, valign: "middle", margin: 0, paraSpaceAfter: 3, isTextBox: true }
    );

    footer(s, [{ text: "Comparison uses representative products and their publicly documented features: AI chatbots (ChatGPT, Gemini), DIY guides (iFixit), repair/trade-in apps (Cashify), OEM diagnostics (Samsung Members)." }]);
    s.addNotes(
      "Five steps, one session: identify → collect evidence → diagnose → decide → track. Two things are new: (1) evidence comes from the device itself — Windows battery reports, macOS system_profiler, browser self-tests — not only from the user's description; (2) the output is a decision with economics (₹ per extra month of use, fair-price band, quote check), not a list of possible faults. The repairability score uses the same six factors as India's Repairability Index framework and the EU A–E label: disassembly depth, fasteners, tools, spare-part availability, software updates and repair information."
    );
  }

  // =============================================================== 4. TECH
  pres.addSection({ title: "Technology" });
  {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Technology" });
    s.addText("TECH STACK + ARCHITECTURE", { placeholder: "title" });
    lede(s, [
      { text: "Language models read the messy input; " },
      { text: "a deterministic, auditable engine makes the call.", options: { bold: true, color: C.crimson } },
    ]);

    panel(s, 0.5, 2.25, 3.85, 4.66, "Tech stack panel");
    sectionHead(s, 0.7, 2.4, I.stack, "Tech Stack", 3);
    const stack = [
      ["FRONTEND", "Next.js 14 · React · TypeScript · Tailwind · PWA with Web APIs for self-tests"],
      ["BACKEND", "Python · FastAPI · Pydantic JSON contracts"],
      ["DIAGNOSTIC ENGINE", "Bayesian fault model (NumPy) · YAML rule graph · info-gain question picker"],
      ["AI SERVICES", "LLM API in JSON-schema mode · vision model · Tesseract OCR"],
      ["DATA", "Supabase: PostgreSQL, Auth, Storage · pgvector for RAG"],
      ["OUTPUT", "WeasyPrint PDF · QR share links · SHA-256 hash-chained Passport"],
      ["DEPLOY", "Vercel · Render · GitHub Actions CI"],
    ];
    const stackRuns = [];
    stack.forEach(([k, v], i) => {
      stackRuns.push({ text: k, options: { fontSize: 8.5, bold: true, color: C.crimson, charSpacing: 1, breakLine: true, paraSpaceBefore: i ? 7 : 0 } });
      stackRuns.push({ text: v, options: { fontSize: 10, color: C.ink, breakLine: i < stack.length - 1 } });
    });
    s.addText(stackRuns, { x: 0.7, y: 2.95, w: 3.5, h: 3.85, valign: "top", margin: 0, isTextBox: true, objectName: "Tech stack list" });

    panel(s, 4.6, 2.25, 8.25, 4.66, "Architecture panel");
    sectionHead(s, 4.8, 2.4, I.arch, "Architecture", 3);

    const box = (x, y, w, h, title, sub, o = {}) => {
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.06, fill: { color: o.fill || C.white }, line: { color: o.line || C.border, width: 0.9 } });
      s.addText(
        [{ text: title, options: { bold: true, fontSize: o.ts || 10, color: o.tc || C.ink, breakLine: true } }, { text: sub, options: { fontSize: o.ss || 8.5, color: o.sc || C.muted } }],
        { x: x + 0.1, y: y + 0.04, w: w - 0.2, h: h - 0.08, valign: "middle", margin: 0, isTextBox: true }
      );
    };
    box(4.8, 2.9, 3.55, 0.6, "Client · Next.js PWA", "Guided intake · self-tests · uploads · Passport viewer");
    box(8.95, 2.9, 3.7, 0.6, "API · FastAPI orchestrator", "Sessions · auth · schema-validated JSON contracts");
    arrow(s, 8.38, 3.2, 8.92, 3.2, { both: true });
    arrow(s, 10.8, 3.5, 10.8, 3.7);

    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 4.8, y: 3.72, w: 7.85, h: 1.42, rectRadius: 0.06, fill: { color: C.tint }, line: { color: "F1C4CC", width: 0.9 }, objectName: "Core" });
    s.addText("REPAIROS CORE — every verdict is traceable to its evidence", { x: 4.95, y: 3.76, w: 7.5, h: 0.22, fontSize: 8.5, bold: true, color: C.crimson, charSpacing: 1, margin: 0, isTextBox: true });
    const stages = [
      ["Evidence extractor", "LLM → JSON symptoms, OCR, vision cues, log parsers"],
      ["Diagnostic engine", "Bayesian ranking + next-best question"],
      ["Safety gate", "Swelling, liquid or heat → stop & route to a pro"],
      ["Decision engine", "Repairability, ₹\u00A0economics, fair-price check"],
      ["Report & Passport", "PDF report, QR link, hash-chained history"],
    ];
    const SW = 1.41, SG = 0.15, SX = 4.9, SY = 4.02, SH = 1.0;
    stages.forEach(([t, d], i) => {
      const x = SX + i * (SW + SG);
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: SY, w: SW, h: SH, rectRadius: 0.05, fill: { color: C.white }, line: { color: C.border, width: 0.75 } });
      s.addShape(pres.shapes.OVAL, { x: x + 0.08, y: SY + 0.09, w: 0.24, h: 0.24, fill: { color: C.crimson }, line: { color: C.crimson } });
      s.addText(String(i + 1), { x: x + 0.08, y: SY + 0.09, w: 0.24, h: 0.24, fontSize: 9, bold: true, color: C.white, align: "center", valign: "middle", margin: 0, isTextBox: true });
      s.addText(t, { x: x + 0.37, y: SY + 0.05, w: SW - 0.42, h: 0.34, fontSize: 9, bold: true, color: C.ink, valign: "middle", margin: 0, isTextBox: true });
      s.addText(d, { x: x + 0.09, y: SY + 0.42, w: SW - 0.16, h: 0.55, fontSize: 8.5, color: C.muted, valign: "top", margin: 0, isTextBox: true });
      if (i < stages.length - 1) arrow(s, x + SW + 0.01, SY + SH / 2, x + SW + SG - 0.01, SY + SH / 2, { width: 1 });
    });

    // Data and AI services: arrows point the way data flows (into the core, or out to storage)
    box(4.8, 5.44, 2.45, 0.62, "AI services", "LLM · vision · OCR APIs");
    box(7.4, 5.44, 2.75, 0.62, "Device knowledge base", "Specs · failure priors · parts · ₹ bands (pgvector)");
    box(10.3, 5.44, 2.35, 0.62, "Supabase", "Postgres · Auth · Storage");
    arrow(s, 5.6, 5.42, 5.6, 5.16);
    arrow(s, 8.18, 5.42, 8.18, 5.16);
    arrow(s, 9.4, 5.42, 9.4, 5.16);
    arrow(s, 11.9, 5.16, 11.9, 5.42);

    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 4.8, y: 6.24, w: 7.85, h: 0.52, rectRadius: 0.06, fill: { color: C.slate }, line: { color: C.slate }, objectName: "Learning loop" });
    s.addImage({ data: I.loop, x: 4.97, y: 6.38, w: 0.24, h: 0.24 });
    s.addText(
      [{ text: "Learning loop: ", options: { bold: true } }, { text: "shop-verified repair outcomes update each model's failure priors — every repair sharpens the next diagnosis." }],
      { x: 5.32, y: 6.24, w: 7.2, h: 0.52, fontSize: 10, color: C.white, valign: "middle", margin: 0, isTextBox: true }
    );
    s.addNotes(
      "Request flow: the PWA collects answers, photos, self-test results and uploaded OS reports; FastAPI validates every payload against JSON schemas; the LLM is used only to turn messy text into structured symptoms (and to phrase the explanation) — it never makes the decision. The Bayesian engine combines model-specific priors from the knowledge base with the evidence to rank faults and pick the next question by expected information gain. The safety gate can override everything. The decision engine computes repairability and economics; the report and Passport are generated last and stored in Supabase."
    );
  }

  // =============================================================== 5. FEASIBILITY
  pres.addSection({ title: "Feasibility" });
  {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Feasibility" });
    s.addText("FEASIBILITY AND SHOWSTOPPERS", { placeholder: "title" });
    lede(s, [
      { text: "Scoped to " },
      { text: "one believable end-to-end flow", options: { bold: true, color: C.crimson } },
      { text: " —", options: { breakLine: true } },
      { text: "built on APIs, rules and public data, with no model training required." },
    ]);

    panel(s, 0.5, 2.25, 5.45, 3.62, "Feasibility panel");
    sectionHead(s, 0.7, 2.4, I.feasible, "Feasibility", 3);
    [["2", "device categories: phones + laptops"], ["10", "high-frequency fault families"], ["25", "top-selling models seeded in India"]].forEach(([n, l], i) => {
      const x = 0.7 + i * 1.7;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 2.92, w: 1.58, h: 0.86, rectRadius: 0.06, fill: { color: C.white }, line: { color: C.border, width: 0.75 } });
      s.addText(n, { x: x + 0.12, y: 2.97, w: 1.4, h: 0.38, fontSize: 21, bold: true, color: C.crimson, valign: "middle", margin: 0, isTextBox: true });
      s.addText(l, { x: x + 0.12, y: 3.35, w: 1.4, h: 0.4, fontSize: 9, color: C.ink, valign: "top", margin: 0, isTextBox: true });
    });
    const why = [
      "No model training: pretrained LLM/vision APIs; reasoning is rules + Bayes in plain Python",
      "Real device evidence is free: Windows battery report, macOS system_profiler, browser sensor APIs",
      "Knowledge base seeded from public spec sheets, service manuals and parts listings",
      "Skills map 1:1 to modules (frontend, engine, AI, product); free-tier hosting",
    ];
    s.addText(why.map((t, i) => ({ text: t, options: { bullet: { indent: 12 }, breakLine: i < why.length - 1 } })), {
      x: 0.7, y: 3.95, w: 5.1, h: 1.85, fontSize: 11, color: C.ink, valign: "top", margin: 0, paraSpaceAfter: 7, isTextBox: true, objectName: "Why buildable",
    });

    panel(s, 6.2, 2.25, 6.65, 3.62, "Showstoppers panel");
    sectionHead(s, 6.4, 2.4, I.shield, "Showstoppers", 3);
    s.addText("RISK", { x: 6.4, y: 2.9, w: 2.0, h: 0.22, fontSize: 8.5, bold: true, color: C.muted, charSpacing: 1, margin: 0, isTextBox: true });
    s.addText("MITIGATION", { x: 8.55, y: 2.9, w: 4.1, h: 0.22, fontSize: 8.5, bold: true, color: C.muted, charSpacing: 1, margin: 0, isTextBox: true });
    const risks = [
      ["Wrong or over-confident diagnosis", "LLM only extracts; the Bayesian engine decides and shows confidence. Low confidence ⇒ “get it inspected”, never “replace”."],
      ["Unsafe DIY advice", "Hard safety gate: swollen battery, liquid or burning smell stops self-help and routes to a professional."],
      ["Thin or volatile price data", "₹ bands labelled indicative, with source and date; partner-shop quotes refine them."],
      ["Long tail of device models", "Category-level priors as fallback; manual model entry when OCR fails."],
      ["LLM latency, cost or outage", "Schema-bound, cached calls; rule-only fallback keeps the core flow running."],
      ["Privacy of IMEI and photos", "Hashed device IDs, consent-first uploads; photos deleted unless saved to the Passport."],
    ];
    risks.forEach(([r, m], i) => {
      const y = 3.14 + i * 0.44;
      s.addShape(pres.shapes.LINE, { x: 6.4, y, w: 6.25, h: 0, line: { color: C.border, width: 0.75 } });
      s.addImage({ data: I.dot, x: 6.4, y: y + 0.18, w: 0.08, h: 0.08 });
      s.addText(r, { x: 6.56, y, w: 1.9, h: 0.44, fontSize: 9.5, bold: true, color: C.ink, valign: "middle", margin: 0, isTextBox: true });
      s.addText(m, { x: 8.55, y, w: 4.15, h: 0.44, fontSize: 9.5, color: C.ink, valign: "middle", margin: 0, isTextBox: true });
    });

    s.addText(
      [{ text: "24-HOUR BUILD PLAN", options: { bold: true, color: C.ink } }, { text: "  ·  demo-first: one complete flow (“laptop won't hold charge”) working by hour 15", options: { color: C.muted } }],
      { x: 0.5, y: 6.02, w: 12.35, h: 0.24, fontSize: 9, charSpacing: 0.5, margin: 0, isTextBox: true }
    );
    const plan = [
      [4, "0–4h", "Scaffold: UI, DB, API contracts"],
      [6, "4–10h", "Knowledge base + Bayesian engine"],
      [5, "10–15h", "AI evidence: LLM, OCR, battery report"],
      [4, "15–19h", "Decision, economics, safety gate"],
      [3, "19–22h", "Report + Repair Passport"],
      [2, "22–24h", "Test & demo polish"],
    ];
    let px = 0.5;
    const unit = (12.35 - 5 * 0.05) / 24;
    plan.forEach(([hrs, h, t], i) => {
      const w = hrs * unit;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: px, y: 6.3, w, h: 0.62, rectRadius: 0.05, fill: { color: i % 2 ? "2C3A4B" : C.slate }, line: { color: i % 2 ? "2C3A4B" : C.slate } });
      s.addText(
        [{ text: h, options: { bold: true, color: "F29AAA", fontSize: 8.5, breakLine: true } }, { text: t, options: { color: C.white, fontSize: 9 } }],
        { x: px + 0.08, y: 6.32, w: w - 0.14, h: 0.58, valign: "middle", margin: 0, isTextBox: true }
      );
      px += w + 0.05;
    });
    s.addNotes(
      "Feasibility rests on scope discipline: two device categories, ten high-frequency fault families, ~25 popular models. Nothing needs training — the LLM and vision model are called through APIs with JSON-schema output, and the diagnostic engine is a small Bayesian model plus a rule graph in Python. Real evidence is available for free: `powercfg /batteryreport` on Windows gives design vs full-charge capacity, macOS `system_profiler SPPowerDataType` gives cycle count and condition, and browser APIs cover touch, audio, camera and motion sensors. All prices, scores and confidence values in the prototype are labelled indicative/model-generated; RepairOS does not issue official repairability ratings or safety certification."
    );
  }

  // =============================================================== 6. USP + BUSINESS
  pres.addSection({ title: "USP and Business Model" });
  {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "USP and Business Model" });
    s.addText("USP(UNIQUE SELLING PRICE) &\nBUSINESS MODEL", { placeholder: "title" });
    lede(s, [
      { text: "We earn whether a device is " },
      { text: "repaired, resold or recycled", options: { bold: true, color: C.crimson } },
      { text: " — so the advice stays honest." },
    ]);

    panel(s, 0.5, 2.25, 5.75, 4.66, "USP panel");
    sectionHead(s, 0.7, 2.4, I.usp, "USP (Unique Selling Proposition)", 4.6);
    const usps = [
      ["Evidence, not opinions", "Reads the device's own health data — battery wear, self-tests, photos — and shows the confidence behind every verdict."],
      ["The decision, not just a diagnosis", "Repair vs replace in ₹ per month of extra life, with a fair-price band and a check on any quote a shop gives you."],
      ["Outcome-neutral by design", "Revenue from every path — repair, resale, recycling — so there is no reason to push a new device."],
      ["Repair Passport", "A portable, tamper-evident service history that makes used devices easier to trust, insure and resell."],
    ];
    usps.forEach(([t, d], i) => {
      const y = 2.95 + i * 0.75;
      s.addText("0" + (i + 1), { x: 0.7, y, w: 0.55, h: 0.32, fontSize: 18, bold: true, color: C.crimson, valign: "middle", margin: 0, isTextBox: true });
      s.addText(t, { x: 1.28, y, w: 4.8, h: 0.28, fontSize: 12, bold: true, color: C.ink, valign: "middle", margin: 0, isTextBox: true });
      s.addText(d, { x: 1.28, y: y + 0.28, w: 4.8, h: 0.42, fontSize: 10, color: C.muted, valign: "top", margin: 0, isTextBox: true });
    });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.7, y: 5.98, w: 5.35, h: 0.82, rectRadius: 0.06, fill: { color: C.slate }, line: { color: C.slate }, objectName: "Why now" });
    s.addText(
      [
        { text: "WHY NOW  ", options: { bold: true, color: "F29AAA", charSpacing: 1 } },
        { text: "India's Repairability Index framework (DoCA, 2025) and the EU's A–E repairability label (June 2025) score the same six factors RepairOS uses. We are the consumer-facing layer for them." },
      ],
      { x: 0.85, y: 6.0, w: 5.08, h: 0.78, fontSize: 9.5, color: C.white, valign: "middle", margin: 0, isTextBox: true }
    );

    panel(s, 6.5, 2.25, 6.35, 4.66, "Business model panel");
    sectionHead(s, 6.7, 2.4, I.biz, "Business Model", 3);
    const TX = [6.7, 9.62, 11.32], TW = [2.82, 1.62, 1.38];
    ["REVENUE STREAM", "WHO PAYS", "PRICING (PLANNED)"].forEach((h, j) => s.addText(h, { x: TX[j], y: 2.9, w: TW[j], h: 0.22, fontSize: 8, bold: true, color: C.muted, charSpacing: 1, margin: 0, isTextBox: true }));
    const streams = [
      ["Free diagnosis + basic Passport", "Consumers", "₹0 · acquisition"],
      ["Verified Repair Report for resale or insurance", "Consumers", "₹99 / report"],
      ["Booked-repair commission", "Partner shops", "8–10% of ticket"],
      ["Shop OS: pre-triage, job cards, parts lookup", "Repair shops", "₹499 / month"],
      ["Diagnostics API: claim & grading triage", "Insurers, refurbishers, IT fleets", "Per device"],
      ["Trade-in & recycling referrals", "Refurbishers, EPR recyclers", "Per referral"],
    ];
    streams.forEach((row, i) => {
      const y = 3.14 + i * 0.4;
      s.addShape(pres.shapes.LINE, { x: 6.7, y, w: 5.98, h: 0, line: { color: C.border, width: 0.75 } });
      s.addText(row[0], { x: TX[0], y, w: TW[0], h: 0.4, fontSize: 9.5, color: C.ink, valign: "middle", margin: 0, isTextBox: true });
      s.addText(row[1], { x: TX[1], y, w: TW[1], h: 0.4, fontSize: 9, color: C.muted, valign: "middle", margin: 0, isTextBox: true });
      s.addText(row[2], { x: TX[2], y, w: TW[2], h: 0.4, fontSize: 9.5, bold: true, color: C.crimson, valign: "middle", margin: 0, isTextBox: true });
    });

    s.addText("GO-TO-MARKET", { x: 6.7, y: 5.6, w: 3, h: 0.22, fontSize: 8, bold: true, color: C.muted, charSpacing: 1, margin: 0, isTextBox: true });
    const gtm = [
      ["Pilot · 0–6 months", "One metro repair hub + 2 campuses; 30 partner shops"],
      ["Scale · 6–18 months", "Multi-city shop network; verified outcomes sharpen priors"],
      ["Platform · 18 months+", "B2B API for insurers, refurbishers and OEM after-sales"],
    ];
    gtm.forEach(([t, d], i) => {
      const w = 1.89, x = 6.7 + i * (w + 0.155), y = 5.86;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h: 0.94, rectRadius: 0.06, fill: { color: C.white }, line: { color: C.border, width: 0.75 } });
      s.addText(
        [{ text: t, options: { bold: true, color: C.crimson, fontSize: 9, breakLine: true } }, { text: d, options: { color: C.ink, fontSize: 9 } }],
        { x: x + 0.1, y: y + 0.06, w: w - 0.2, h: 0.82, valign: "top", margin: 0, paraSpaceAfter: 2, isTextBox: true }
      );
      if (i < gtm.length - 1) arrow(s, x + w + 0.02, y + 0.47, x + w + 0.135, y + 0.47, { width: 1 });
    });

    footer(s, [
      { text: "Sources: Dept. of Consumer Affairs, Govt. of India — Repairability Index framework for the mobile & electronics sector (committee report, 2025) · EU Delegated Regulation 2023/1669 (repairability class on smartphone and tablet labels from 20 Jun 2025) · MoEFCC, E-Waste (Management) Rules, 2022. Pricing is planned and indicative." },
    ]);
    s.addNotes(
      "Monetisation is deliberately outcome-neutral: consumers get diagnosis free; we earn from verified reports, booked-repair commissions, shop SaaS, B2B diagnostics API, and trade-in/recycling referrals when replacement genuinely is the right call. That alignment is the core of the USP — the recommendation never depends on which path makes us money.\n\nSources:\n- Department of Consumer Affairs: committee on Repairability Index for mobile and electronics (constituted Sep 2024, report 2025; parameters: disassembly depth, repair information, spare-part availability, software updates, tools, fasteners).\n- EU Delegated Regulation 2023/1669 — energy label for smartphones and tablets incl. repairability class A–E, applicable from 20 June 2025: https://energy-efficient-products.ec.europa.eu/product-list/smartphones-and-tablets_en\n- E-Waste (Management) Rules, 2022 (MoEFCC), in force 1 Apr 2023."
    );
  }

  await pres.writeFile({ fileName: OUT });
  await applyThemeColors(OUT, THEME);
  console.log("wrote " + OUT);
})().catch((e) => {
  console.error(e);
  process.exit(1);
});
