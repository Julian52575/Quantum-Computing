// Build the defense deck. Usage: node build_deck.js <out.pptx>
const path = require("path");
const fs = require("fs");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const { applyTheme } = require(process.env.PPTX_SKILL + "/scripts/apply_theme.js");

const OUT = process.argv[2] || "deck.pptx";
const D = JSON.parse(fs.readFileSync(path.join(__dirname, "data.json"), "utf8"));
const IMG = (f) => path.join(__dirname, f);

const THEME = {
  name: "Order Selection QAOA",
  headFontFace: "Cambria",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "1B2A33", lt1: "FFFFFF", dk2: "264653", lt2: "EEF5F4",
    accent1: "2A9D8F", accent2: "E76F51", accent3: "E9C46A",
    accent4: "264653", accent5: "8D99AE", accent6: "F4A261",
    hlink: "2A9D8F", folHlink: "264653",
  },
};
const HEX = THEME.colors;
const MUTED = "5B6770";

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.title = "Conflict-free order selection with QAOA";
pres.subject = "Tech 5 Quantum Computing final defense";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
const C = pres.SchemeColor;

const W = 13.333;
const MX = 0.6; // side margin
const CW = W - 2 * MX; // content width

// ------------------------------------------------------------- layouts
pres.defineSlideMaster({
  title: "TITLE_DARK",
  background: { color: C.text2 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: MX, y: 2.0, w: 7.2, h: 2.2, fontSize: 42, bold: true,
      color: C.background1, valign: "bottom", align: "left", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: MX, y: 4.45, w: 7.0, h: 1.6, fontSize: 18,
      color: "CFE3E0", valign: "top", margin: 0 }, text: "" } },
  ],
});

pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: C.background1 },
  margin: [0.5, 0.6, 0.6, 0.6],
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: MX, y: 0.78, w: CW, h: 0.85, fontSize: 28, bold: true,
      color: C.text2, valign: "top", align: "left", margin: 0 }, text: "" } },
    { text: { text: "Conflict-free order selection with QAOA  •  Tech 5 Quantum Computing",
      options: { x: MX, y: 7.05, w: 8, h: 0.3, fontSize: 10, color: MUTED, margin: 0 } } },
  ],
  slideNumber: { x: W - MX - 0.6, y: 7.05, w: 0.6, h: 0.3, fontSize: 10, color: MUTED, align: "right" },
});

pres.defineSlideMaster({
  title: "CLOSING_DARK",
  background: { color: C.text2 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: MX, y: 0.55, w: CW, h: 0.9, fontSize: 32, bold: true,
      color: C.background1, valign: "top", align: "left", margin: 0 }, text: "" } },
  ],
  slideNumber: { x: W - MX - 0.6, y: 7.05, w: 0.6, h: 0.3, fontSize: 10, color: "CFE3E0", align: "right" },
});

// ------------------------------------------------------------- helpers
async function icon(Comp, color) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Comp, { color: "#" + color, size: 256 }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

// Defense agenda tracker: the repeated motif, one pill per defense part.
const PARTS = ["Problem & model", "Classical baseline", "Quantum workflow", "QAOA results", "Critical analysis"];
function tracker(slide, active) {
  const pw = 2.1, gap = 0.12, h = 0.32;
  PARTS.forEach((label, i) => {
    const on = i === active;
    slide.addText(`${i + 1}  ${label}`, {
      shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.16, x: MX + i * (pw + gap), y: 0.3, w: pw, h,
      fill: { color: on ? C.accent1 : C.background2 }, color: on ? C.background1 : MUTED,
      fontSize: 11, bold: on, align: "center", valign: "middle", margin: 0, objectName: `tracker-${i + 1}`,
    });
  });
}

function contentSlide(part, title, section) {
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: section });
  s.addText(title, { placeholder: "title" });
  if (part !== null) tracker(s, part);
  return s;
}

function card(slide, x, y, w, h, name, fill) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h, rectRadius: 0.12, fill: { color: fill || C.background2 }, line: { type: "none" }, objectName: name,
  });
}

// Formula text: "_{..}" subscript, "^{..}" superscript.
function formula(src, base) {
  const runs = [];
  const re = /([_^])\{([^}]*)\}/g;
  let last = 0, m;
  while ((m = re.exec(src))) {
    if (m.index > last) runs.push({ text: src.slice(last, m.index), options: { ...base } });
    runs.push({ text: m[2], options: { ...base, [m[1] === "_" ? "subscript" : "superscript"]: true } });
    last = re.lastIndex;
  }
  if (last < src.length) runs.push({ text: src.slice(last), options: { ...base } });
  return runs;
}

function stat(slide, x, y, w, big, label, name, color, size) {
  slide.addText(big, { x, y, w, h: 0.85, fontSize: size || 40, bold: true, color: color || C.accent1, fontFace: THEME.headFontFace,
    margin: 0, valign: "bottom", isTextBox: true, objectName: name + "-value" });
  slide.addText(label, { x, y: y + 0.9, w, h: 0.6, fontSize: 14, color: MUTED, margin: 0, valign: "top",
    isTextBox: true, objectName: name + "-label" });
}

function bullets(items, size = 16) {
  return items.map((t, i) => {
    const runs = Array.isArray(t) ? t : [{ text: t }];
    return runs.map((r, k) => ({
      text: r.text,
      options: { ...(r.options || {}), ...(k === 0 ? { bullet: true, paraSpaceAfter: 8 } : {}),
        ...(k === runs.length - 1 && i < items.length - 1 ? { breakLine: true } : {}), fontSize: size },
    }));
  }).flat();
}

const pct = (v, d = 0) => `${(100 * v).toFixed(d)} %`;
const CHART_FONT = { catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt",
  legendFontFace: "+mn-lt", titleFontFace: "+mn-lt", catAxisTitleFontFace: "+mn-lt", valAxisTitleFontFace: "+mn-lt" };
const chartBase = () => ({
  ...CHART_FONT, catAxisLabelColor: MUTED, valAxisLabelColor: MUTED, catAxisLabelFontSize: 12, valAxisLabelFontSize: 12,
  dataLabelFontSize: 11, dataLabelColor: "1B2A33", legendFontSize: 12, legendColor: "1B2A33",
  valGridLine: { color: "DDE3E6", size: 0.75 }, catGridLine: { style: "none" }, titleFontSize: 14, titleColor: "1B2A33",
  catAxisTitleColor: MUTED, valAxisTitleColor: MUTED, catAxisTitleFontSize: 12, valAxisTitleFontSize: 12,
});

async function main() {
  const icons = {
    tower: await icon(fa.FaBroadcastTower, HEX.accent1), ad: await icon(fa.FaBullhorn, HEX.accent1),
    truck: await icon(fa.FaTruck, HEX.accent1), chart: await icon(fa.FaChartLine, HEX.accent1),
    check: await icon(fa.FaCheck, HEX.lt1), cross: await icon(fa.FaTimes, HEX.lt1), arrow: await icon(fa.FaArrowRight, HEX.accent5),
    redo: await icon(fa.FaRedo, HEX.accent2),
  };

  const th = (t, o = {}) => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.text2 }, ...o } });
  const td = (t, o = {}) => ({ text: t, options: { color: "1B2A33", ...o } });

  // ---------------------------------------------------- 1. title
  pres.addSection({ title: "Introduction" });
  let s = pres.addSlide({ masterName: "TITLE_DARK", sectionTitle: "Introduction" });
  s.addText("Conflict-free order selection with QAOA", { placeholder: "title" });
  s.addText("From an industrial scheduling problem to a quantum circuit\nTech 5 Quantum Computing  •  Final defense  •  EPITECH 2026/2027",
    { placeholder: "body" });
  card(s, 8.3, 1.9, 4.45, 3.7, "title-graph-card", C.background1);
  s.addImage({ path: IMG("graph_opt.png"), x: 8.45, y: 2.55, w: 4.15, h: 4.15 * 602 / 1290, altText: "Conflict graph of 8 orders with the optimal selection highlighted" });
  s.addText("8 orders, 10 conflicts: accept 1, 4, 5 and 8 for a profit of 17", { x: 8.5, y: 4.65, w: 4.05, h: 0.7, fontSize: 13,
    color: MUTED, align: "center", isTextBox: true, objectName: "title-graph-caption" });
  s.addNotes("Introduce the team. Our project takes a real decision problem, which orders a workshop should accept when some of them conflict, and carries it through the full chain of the module: industry problem, math model, QUBO, Hamiltonian, QAOA, and critical analysis. The picture on the right is our instance and its optimal answer, which we will come back to.");

  // ---------------------------------------------------- 2. chain + agenda
  s = contentSlide(null, "One modeling chain, followed step by step", "Introduction");
  const steps = [
    ["Industry problem", "Workshop orders with shared resources"],
    ["Math model", "Binary choices, profit, conflict rule"],
    ["QUBO", "f(x) with a conflict penalty"],
    ["Hamiltonian", "Z fields + ZZ couplings"],
    ["QAOA", "Circuit + classical optimizer"],
    ["Analysis", "Layers, penalty, size, noise"],
  ];
  const sw = 1.83, sg = 0.25, sy = 2.0;
  steps.forEach(([t, d], i) => {
    const x = MX + i * (sw + sg);
    card(s, x, sy, sw, 2.35, `chain-card-${i + 1}`);
    s.addText(String(i + 1), { shape: pres.shapes.OVAL, x: x + 0.2, y: sy + 0.2, w: 0.5, h: 0.5, fill: { color: C.accent1 },
      color: C.background1, fontSize: 16, bold: true, align: "center", valign: "middle", margin: 0, objectName: `chain-num-${i + 1}` });
    s.addText(t, { x: x + 0.2, y: sy + 0.8, w: sw - 0.4, h: 0.6, fontSize: 16, bold: true, color: C.text2, margin: 0, valign: "top",
      isTextBox: true, objectName: `chain-title-${i + 1}` });
    s.addText(d, { x: x + 0.2, y: sy + 1.42, w: sw - 0.4, h: 0.85, fontSize: 13, color: MUTED, margin: 0, valign: "top",
      isTextBox: true, objectName: `chain-desc-${i + 1}` });
    if (i < steps.length - 1) s.addImage({ data: icons.arrow, x: x + sw + 0.04, y: sy + 0.95, w: 0.17, h: 0.17, altText: "" });
  });
  s.addText("Defense agenda", { x: MX, y: 4.65, w: 4, h: 0.4, fontSize: 18, bold: true, color: C.text2, margin: 0,
    isTextBox: true, objectName: "agenda-heading" });
  const agenda = ["Chain steps 1 to 4", "Brute-force ground truth", "Chain step 5", "Same instance as part 2", "Chain step 6"];
  const aw = (CW - 4 * 0.25) / 5;
  agenda.forEach((sub, i) => {
    s.addText([{ text: `${i + 1}  ${PARTS[i]}`, options: { bold: true, color: C.background1, fontSize: 14, breakLine: true } },
      { text: sub, options: { color: "D7ECE9", fontSize: 12 } }],
      { shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.12, x: MX + i * (aw + 0.25), y: 5.2, w: aw, h: 0.95,
        fill: { color: i % 2 ? C.text2 : C.accent1 }, margin: [12, 12, 0, 0], valign: "middle", objectName: `agenda-${i + 1}` });
  });
  s.addNotes("This is the storyline from the kickoff deck, page 4. Each step produces something concrete that feeds the next one, and each step of our notebook ends with a check against the previous one. The defense follows the five parts shown at the top of every slide.");

  // ---------------------------------------------------- 3. industry problem
  pres.addSection({ title: "Problem & model" });
  s = contentSlide(0, "A workshop gets more orders than it can process", "Problem & model");
  s.addText("Each order brings a profit and needs scarce resources. A resource serves one order per week.",
    { x: MX, y: 1.95, w: 6.0, h: 0.85, fontSize: 18, color: "1B2A33", valign: "top", margin: 0, isTextBox: true, objectName: "industry-context" });
  [["Decision", "which orders to accept"], ["Goal", "maximize the total profit"], ["Rule", "never accept two orders that share a resource"]].forEach(([k, v], i) => {
    const y = 3.0 + i * 0.75;
    s.addText(k, { shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.1, x: MX, y, w: 1.4, h: 0.5, fill: { color: C.accent1 },
      color: C.background1, fontSize: 15, bold: true, align: "center", valign: "middle", margin: 0, objectName: `industry-key-${i + 1}` });
    s.addText(v, { x: MX + 1.6, y, w: 4.6, h: 0.5, fontSize: 18, color: "1B2A33", valign: "middle", margin: 0, isTextBox: true, objectName: `industry-val-${i + 1}` });
  });
  // Mirrors RESOURCES in the notebook: each resource shared by two orders is one conflict edge.
  const resources = [["Machine", [["CNC mill", "1, 2"], ["Laser cutter", "3, 4"], ["Paint booth", "5, 6"], ["Lathe", "7, 8"]]],
    ["Worker", [["Certified welder", "1, 3"], ["Quality inspector", "4, 6"], ["Electrician", "5, 7"]]],
    ["Time slot", [["Monday dock slot", "2, 4"], ["Wednesday dock slot", "3, 5"], ["Friday dock slot", "6, 8"]]]];
  const resRows = [[th("Type"), th("Resource (capacity 1)"), th("Orders", { align: "center" })]];
  resources.forEach(([type, items], gi) => items.forEach(([name, orders], k) => resRows.push([
    ...(k === 0 ? [td(type, { rowspan: items.length, bold: true, color: C.text2, fill: { color: gi % 2 ? "FFFFFF" : HEX.lt2 } })] : []),
    td(name), td(orders, { align: "center", bold: true, color: C.accent2 })])));
  s.addTable(resRows, { x: 7.1, y: 1.95, w: 5.63, colW: [1.4, 2.73, 1.5], fontSize: 13, rowH: 0.3, valign: "middle",
    border: { type: "solid", pt: 0.75, color: "DDE3E6" }, margin: [0.02, 0.1, 0.02, 0.1], objectName: "resource-table" });
  card(s, MX, 5.55, CW, 1.15, "combinatorics-card", C.text2);
  s.addText([{ text: "2", options: { fontSize: 36, bold: true, color: C.accent3 } }, { text: "n", options: { fontSize: 36, bold: true, color: C.accent3, superscript: true } }],
    { x: MX + 0.3, y: 5.6, w: 0.9, h: 1.05, margin: 0, valign: "middle", fontFace: THEME.headFontFace, isTextBox: true, objectName: "combinatorics-value" });
  s.addText("possible selections: 256 for our 8 orders, about 10¹⁵ for 50 orders. Checking them all does not scale.",
    { x: MX + 1.25, y: 5.6, w: CW - 1.55, h: 1.05, fontSize: 17, color: C.background1, margin: 0, valign: "middle", isTextBox: true, objectName: "combinatorics-text" });
  s.addNotes("The business story. Each order needs some resources, and each resource can serve only one order in the planning week. Two orders that share a resource conflict: the table gives the ten resources behind our ten conflicts, a mix of machines, certified workers and dock slots. That mix matters, as we will see on the next slide. If asked about real capacities such as machine-hours: those need knapsack-type constraints with slack qubits, which is a different problem. The number of possible selections doubles with every order: 256 for our instance, but around 10^15 for 50 orders.");

  // ---------------------------------------------------- 4. MWIS
  s = contentSlide(0, "Accepting orders = picking an independent set", "Problem & model");
  card(s, MX, 1.95, 6.4, 3.45, "graph-card");
  s.addImage({ path: IMG("graph.png"), x: MX + 0.25, y: 2.1, w: 5.9, h: 5.9 * 602 / 1290, altText: "Conflict graph: 8 orders as vertices, 10 conflicts as edges, shaped like a ladder" });
  s.addText("Vertex = order (with its profit p)  •  edge = conflict", { x: MX + 0.25, y: 4.9, w: 5.9, h: 0.4, fontSize: 13, color: MUTED,
    align: "center", margin: 0, isTextBox: true, objectName: "graph-caption" });
  s.addText([{ text: "Maximum Weight Independent Set (MWIS)", options: { bold: true, fontSize: 18, color: C.text2, breakLine: true } },
    { text: "A valid selection has no edge inside it. Best profit = maximum weight independent set, an NP-hard problem.", options: { fontSize: 14, color: "1B2A33" } }],
    { x: 7.35, y: 1.95, w: 5.38, h: 1.2, margin: 0, valign: "top", isTextBox: true, objectName: "mwis-definition" });
  const uses = [["tower", "Telecoms", "Interfering transmitters never share a frequency"], ["ad", "Advertising", "Sell mutually exclusive ad placements"],
    ["truck", "Logistics", "Deliveries that compete for a vehicle or dock"], ["chart", "Finance", "Pick non-overlapping projects or trades"]];
  uses.forEach(([ic, t, d], i) => {
    const y = 3.3 + i * 0.55;
    s.addImage({ data: icons[ic], x: 7.35, y: y + 0.07, w: 0.3, h: 0.3, altText: t });
    s.addText([{ text: t + ": ", options: { bold: true, color: C.text2 } }, { text: d, options: { color: "1B2A33" } }],
      { x: 7.8, y, w: 4.93, h: 0.45, fontSize: 14, margin: 0, valign: "middle", isTextBox: true, objectName: `use-${i + 1}` });
  });
  card(s, MX, 5.65, CW, 1.05, "scope-card", "FBE3DC");
  s.addText([{ text: "Scope. ", options: { bold: true, color: C.accent2 } },
    { text: "If every conflict came from overlapping time intervals on one machine, the graph would be an interval graph and dynamic programming would solve it in polynomial time. Our conflicts mix machines, workers and slots, which gives general graphs: that is where MWIS is hard." }],
    { x: MX + 0.3, y: 5.7, w: CW - 0.6, h: 0.95, fontSize: 14, color: "1B2A33", margin: 0, valign: "middle", isTextBox: true, objectName: "scope-text" });
  s.addNotes("Orders become vertices, conflicts become edges. A valid selection is an independent set, and the most profitable one is the maximum weight independent set, which is NP-hard. The same model appears in telecoms, advertising, logistics and finance. Be ready for the question: if conflicts were only time overlaps on a single machine, this would be weighted interval scheduling, solvable exactly by dynamic programming. Mixed resources give general graphs.");

  // ---------------------------------------------------- 5. model
  s = contentSlide(0, "The model: binary choices, a profit, a conflict rule", "Problem & model");
  const F = { fontFace: "Cambria Math", fontSize: 22, color: C.text2 };
  const cols = [
    ["Variables", formula("x_{i} ∈ {0, 1}", F), "x_{i} = 1 if order i is accepted"],
    ["Objective", formula("max  Σ_{i} p_{i} x_{i}", F), "total profit of the accepted orders"],
    ["Constraints", formula("x_{i} + x_{j} ≤ 1", F), "for every conflict (i, j) ∈ E"],
  ];
  const cwid = (CW - 2 * 0.3) / 3;
  cols.forEach(([t, f, d], i) => {
    const x = MX + i * (cwid + 0.3);
    card(s, x, 1.95, cwid, 2.0, `model-card-${i + 1}`);
    s.addText(t, { x: x + 0.3, y: 2.1, w: cwid - 0.6, h: 0.4, fontSize: 15, bold: true, color: C.accent1, margin: 0, isTextBox: true, objectName: `model-title-${i + 1}` });
    s.addText(f, { x: x + 0.3, y: 2.55, w: cwid - 0.6, h: 0.7, margin: 0, valign: "middle", isTextBox: true, objectName: `model-formula-${i + 1}` });
    s.addText(formula(d, { fontSize: 14, color: MUTED }), { x: x + 0.3, y: 3.3, w: cwid - 0.6, h: 0.5, margin: 0, valign: "top", isTextBox: true, objectName: `model-desc-${i + 1}` });
  });
  card(s, MX, 4.2, CW, 1.15, "key-card", C.text2);
  s.addText([{ text: "Key observation  ", options: { bold: true, color: C.accent3, fontSize: 16 } },
    ...formula("For binary variables, x_{i} + x_{j} ≤ 1  ⇔  x_{i} · x_{j} = 0. The constraint becomes a product, which a QUBO can absorb without extra variables.", { color: C.background1, fontSize: 16 })],
    { x: MX + 0.3, y: 4.25, w: CW - 0.6, h: 1.05, margin: 0, valign: "middle", isTextBox: true, objectName: "key-text" });
  // Instance table
  const head = [{ text: "Order", options: { bold: true, color: C.background1, fill: { color: C.accent1 } } },
    ...PROFITS().map((_, i) => ({ text: String(i + 1), options: { bold: true, color: C.background1, fill: { color: C.accent1 }, align: "center" } }))];
  const row = [{ text: "Profit p", options: { bold: true, color: C.text2 } }, ...PROFITS().map((p) => ({ text: String(p), options: { align: "center", color: "1B2A33" } }))];
  s.addTable([head, row], { x: MX, y: 5.6, w: 7.4, colW: [1.4, ...Array(8).fill(0.75)], fontSize: 14, border: { type: "solid", pt: 0.75, color: "DDE3E6" },
    rowH: 0.4, margin: 0.05, objectName: "instance-table" });
  s.addText("Conflicts: (1,2) (1,3) (2,4) (3,4) (3,5) (4,6) (5,6) (5,7) (6,8) (7,8)", { x: 8.3, y: 5.6, w: 4.43, h: 0.8, fontSize: 14,
    color: "1B2A33", margin: 0, valign: "middle", isTextBox: true, objectName: "instance-conflicts" });
  s.addNotes("Three ingredients. One binary variable per order, a linear objective, and one inequality per conflict. The key observation: for binary variables, 'at most one of the two' is the same as 'their product is zero'. This is what lets us build a QUBO with exactly one qubit per order. Bottom: our 8-order instance.");

  // ---------------------------------------------------- 6. QUBO
  s = contentSlide(0, "Conflicts become a penalty, and A > max pᵢ is enough", "Problem & model");
  card(s, MX, 1.95, CW, 1.2, "qubo-formula-card", C.text2);
  s.addText(formula("min  f(x) = − Σ_{i} p_{i} x_{i}  +  A · Σ_{(i,j) ∈ E} x_{i} x_{j}", { fontFace: "Cambria Math", fontSize: 28, color: C.background1 }),
    { x: MX, y: 1.95, w: CW, h: 1.2, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: "qubo-formula" });
  s.addText("Why A > max pᵢ guarantees a valid optimum", { x: MX, y: 3.45, w: 7.5, h: 0.4, fontSize: 18, bold: true, color: C.text2, margin: 0, isTextBox: true, objectName: "penalty-heading" });
  const why = ["Take any assignment that breaks a conflict (i, j)", "Drop order j: profit falls by pⱼ ≤ max pᵢ", "Penalty falls by at least A",
    "If A > max pᵢ, f strictly decreases, so no conflict-breaking assignment can be optimal"];
  why.forEach((t, i) => {
    const y = 4.0 + i * 0.62;
    s.addText(String(i + 1), { shape: pres.shapes.OVAL, x: MX, y, w: 0.42, h: 0.42, fill: { color: C.accent1 }, color: C.background1, fontSize: 14, bold: true,
      align: "center", valign: "middle", margin: 0, objectName: `why-num-${i + 1}` });
    s.addText(t, { x: MX + 0.6, y: y - 0.04, w: 6.9, h: 0.5, fontSize: 16, color: "1B2A33", margin: 0, valign: "middle", isTextBox: true, objectName: `why-text-${i + 1}` });
  });
  card(s, 8.6, 3.45, 4.13, 3.25, "penalty-stat-card");
  stat(s, 8.9, 3.6, 3.6, "A = 7", "smallest integer above the largest profit (6)", "penalty-stat");
  stat(s, 8.9, 5.05, 3.6, "−12 vs −17", "best conflict-breaking assignment vs the optimum: the penalty works", "penalty-check", C.accent2);
  s.addNotes("The constraint moves into the objective as a penalty that is zero for valid selections. The argument for A: dropping one order from a conflicting pair loses at most the largest profit but removes at least A of penalty, so as long as A is larger than the largest profit, breaking a conflict never pays. A = 7 is the smallest integer that works. The brute force confirms it: the best conflict-breaking assignment scores -12, well above the optimum at -17.");

  // ---------------------------------------------------- 7. converters
  s = contentSlide(0, "How the QUBO is built changes qubits and penalty", "Problem & model");
  s.addTable([
    [th("Construction"), th("Qubits"), th("Penalty"), th("Optimum"), th("Comment")],
    [td("Step-by-step converters (slack variables, notebook 2)"), td("18", { bold: true, color: C.accent2 }), td("33"), td("{1, 4, 5, 8}"), td("One slack bit per conflict: more than twice the qubits")],
    [td("QuadraticProgramToQubo"), td("8"), td("363", { bold: true, color: C.accent2 }), td("{1, 4, 5, 8}"), td("Spots xᵢ + xⱼ ≤ 1, but picks a very safe, very large penalty")],
    [td("Hand-built (our choice)", { bold: true }), td("8", { bold: true, color: C.accent1 }), td("7", { bold: true, color: C.accent1 }), td("{1, 4, 5, 8}"), td("Product trick + smallest valid penalty")],
  ], { x: MX, y: 2.0, w: CW, colW: [3.4, 1.0, 1.1, 1.6, 5.033], fontSize: 15, rowH: [0.5, 0.75, 0.75, 0.75], valign: "middle",
    border: { type: "solid", pt: 0.75, color: "DDE3E6" }, margin: [0.04, 0.12, 0.04, 0.12], objectName: "converter-table" });
  stat(s, MX, 5.15, 3.6, "Same optimum", "all three QUBOs agree with brute force", "conv-same", undefined, 32);
  stat(s, MX + 4.1, 5.15, 3.6, "18 → 8 qubits", "fewer qubits = shallower circuit, less noise", "conv-qubits", C.accent2, 32);
  stat(s, MX + 8.2, 5.15, 3.9, "363 → 7", "the penalty scale matters for QAOA (slide 16)", "conv-penalty", C.accent2, 32);
  s.addNotes("We compared three ways to build the QUBO. The step-by-step pipeline from notebook 2 adds one slack bit per conflict: 18 qubits. Qiskit's QuadraticProgramToQubo recognizes the at-most-one pattern and avoids slack, but chooses a penalty of 363. Our hand-built QUBO uses 8 qubits and A = 7. All three have the same optimum, so for brute force they are equivalent, but not for QAOA, as we show in the analysis part.");

  // ---------------------------------------------------- 8. Hamiltonian
  s = contentSlide(0, "xᵢ = (1 − zᵢ)/2 turns the QUBO into Z and ZZ terms", "Problem & model");
  const H = { fontFace: "Cambria Math", fontSize: 17, color: "1B2A33" };
  const lines = [
    [{ text: "Profit term", options: { bold: true, color: C.accent1, fontSize: 15, breakLine: true } }],
    [...formula("− p_{i} x_{i}  =  − p_{i}/2  +  (p_{i}/2) Z_{i}", H), { text: "", options: { breakLine: true } }],
    [{ text: "Penalty term", options: { bold: true, color: C.accent2, fontSize: 15, breakLine: true } }],
    [...formula("A x_{i} x_{j}  =  (A/4)(1 − Z_{i} − Z_{j} + Z_{i} Z_{j})", H), { text: "", options: { breakLine: true } }],
    [{ text: "Result (dᵢ = number of conflicts of order i)", options: { bold: true, color: C.text2, fontSize: 15, breakLine: true } }],
    [...formula("H_{C} = Σ_{i} (p_{i}/2 − A d_{i}/4) Z_{i}  +  Σ_{E} (A/4) Z_{i} Z_{j}", H)],
  ].flat();
  card(s, MX, 1.95, 6.5, 3.55, "derivation-card");
  s.addText(lines, { x: MX + 0.3, y: 2.05, w: 6.0, h: 3.35, margin: 0, valign: "middle", paraSpaceAfter: 6, isTextBox: true, objectName: "derivation" });
  const hz = [-1.0, -2.0, -3.25, -2.25, -4.25, -2.75, -2.0, -1.5];
  s.addChart(pres.charts.BAR, [{ name: "Z coefficient hᵢ", labels: hz.map((_, i) => `Order ${i + 1}`), values: hz }], {
    x: 7.4, y: 1.85, w: 5.33, h: 3.75, barDir: "col", ...chartBase(), chartColors: [HEX.accent2], showValue: true, dataLabelPosition: "outEnd",
    dataLabelFormatCode: "0.00", showLegend: false, showTitle: true, title: "Z field per order (A = 7): conflicts outweigh profit",
    valAxisMinVal: -5, valAxisMaxVal: 0, valAxisLabelFormatCode: "0", objectName: "z-chart",
  });
  stat(s, MX, 5.75, 3.4, "+1.75", "ZZ coupling A/4 on each of the 10 conflicts", "zz-stat");
  stat(s, MX + 3.9, 5.75, 3.0, "1.5", "constant offset kept to read f(x)", "offset-stat", C.text2);
  stat(s, MX + 7.4, 5.75, 4.6, "0.0", "max error of ⟨x|H|x⟩ + offset vs f(x), over all 256 states", "check-stat", C.accent1);
  s.addNotes("Each binary variable becomes a spin with x = (1 - z)/2, then z becomes the Pauli Z operator. Profit gives a Z field; each penalty term gives a ZZ coupling, but also Z fields. That is why every order's Z coefficient is negative here: the penalty contributions outweigh the profit, which pulls orders towards rejection, while the ZZ couplings punish accepting both ends of a conflict. Our hand derivation matches Qiskit's to_ising exactly, and the diagonal of H plus the offset equals f(x) on all 256 basis states.");

  // ---------------------------------------------------- 9. exhaustive search
  pres.addSection({ title: "Classical baseline" });
  s = contentSlide(1, "Brute force: unique optimum {1, 4, 5, 8}, profit 17", "Classical baseline");
  const rk = [["1", "−17", "{1, 4, 5, 8}", "yes"], ["2", "−15", "{2, 3, 6, 7}", "yes"], ["3", "−15", "{1, 4, 8}", "yes"], ["4", "−14", "{1, 4, 7}", "yes"],
    ["5", "−13", "{1, 6, 7}", "yes"], ["…", "", "", ""], ["first invalid", "−12", "{1, 4, 6, 7}", "no"]];
  s.addTable([[th("Rank"), th("f(x)"), th("Orders accepted"), th("Valid")],
    ...rk.map((r, i) => r.map((c, k) => td(c, { bold: i === 0, color: i === 6 ? C.accent2 : i === 0 ? C.accent1 : "1B2A33", align: k === 2 ? "left" : "center" })))],
    { x: MX, y: 1.95, w: 6.0, colW: [1.5, 1.0, 2.4, 1.1], fontSize: 15, rowH: 0.45, border: { type: "solid", pt: 0.75, color: "DDE3E6" }, margin: [0.04, 0.1, 0.04, 0.1],
      valign: "middle", objectName: "ranking-table" });
  card(s, 7.0, 1.95, 5.73, 3.15, "opt-graph-card");
  s.addImage({ path: IMG("graph_opt.png"), x: 7.2, y: 2.15, w: 5.33, h: 5.33 * 602 / 1290, altText: "Optimal selection: orders 1, 4, 5 and 8 highlighted in the conflict graph" });
  s.addText("Optimal selection (green): no two accepted orders share an edge", { x: 7.2, y: 4.62, w: 5.33, h: 0.4, fontSize: 13, color: MUTED, align: "center",
    margin: 0, isTextBox: true, objectName: "opt-caption" });
  stat(s, 7.0, 5.3, 2.6, "1.1 ms", "to check all 256 cases", "bf-time", C.text2);
  stat(s, 9.9, 5.3, 2.8, "2 ties", "for second place (profit 15)", "bf-ties", C.text2);
  s.addNotes("We enumerate all 256 assignments. The optimum is unique: orders 1, 4, 5 and 8, profit 17. Two different plans tie for second at profit 15, so a near miss can still be a good plan. The first assignment that breaks a conflict only appears at -12. Brute force takes about a millisecond here: this is our ground truth, and Qiskit's exact NumPyMinimumEigensolver agrees.");

  // ---------------------------------------------------- 10. landscape
  s = contentSlide(1, "Only 16 % of all choices are valid", "Classical baseline");
  const buckets = [];
  for (let lo = -17; lo <= 38; lo += 5) buckets.push([lo, Math.min(lo + 4, 38)]);
  const sumIn = (arr, lo, hi) => D.hist.bins.reduce((acc, b, i) => acc + (b >= lo && b <= hi ? arr[i] : 0), 0);
  const blabels = buckets.map(([lo, hi]) => `${lo} to ${hi}`.replace(/-/g, "−"));
  s.addChart(pres.charts.BAR, [
    { name: "valid", labels: blabels, values: buckets.map(([lo, hi]) => sumIn(D.hist.valid, lo, hi)) },
    { name: "breaks a conflict", labels: blabels, values: buckets.map(([lo, hi]) => sumIn(D.hist.invalid, lo, hi)) },
  ], { x: MX, y: 1.85, w: 7.6, h: 4.9, barDir: "col", barGrouping: "stacked", ...chartBase(), chartColors: [HEX.accent1, HEX.accent2],
    showLegend: true, legendPos: "t", showTitle: true, title: "Number of assignments per value of f(x)", showCatAxisTitle: true, catAxisTitle: "f(x)  (lower is better)",
    catAxisLabelFontSize: 10, catAxisLabelRotate: -45, objectName: "landscape-chart" });
  s.addText("Random guessing, the bar QAOA must clear", { x: 8.65, y: 1.95, w: 4.08, h: 0.45, fontSize: 17, bold: true, color: C.text2, margin: 0, isTextBox: true, objectName: "random-heading" });
  stat(s, 8.65, 2.45, 4.0, pct(D.random["valid rate"]), "valid rate (41 of 256)", "rand-valid");
  stat(s, 8.65, 3.95, 4.0, pct(D.random["P(optimum)"], 1), "chance of the optimum (1 of 256)", "rand-popt");
  stat(s, 8.65, 5.45, 4.0, D.random["approx. ratio"].toFixed(2), "expected profit / 17 (invalid = 0)", "rand-ratio");
  s.addNotes("The energy landscape of all 256 assignments. Valid ones, in green, are all on the left: the penalty pushes every conflict-breaking assignment to higher f. Only 16 percent are valid. These three numbers are what uniform random guessing gives, and they are the baseline every QAOA result is compared to: valid rate, probability of the optimum, and approximation ratio.");

  // ---------------------------------------------------- 11. QAOA workflow
  pres.addSection({ title: "Quantum workflow" });
  s = contentSlide(2, "QAOA: the circuit proposes, the optimizer adjusts", "Quantum workflow");
  const box = (x, y, w, h, title, body, name, fill, tcol) => {
    s.addText([{ text: title, options: { bold: true, fontSize: 16, color: tcol || C.text2, breakLine: true } }, ...formula(body, { fontSize: 14, color: tcol || "1B2A33" })],
      { shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.12, x, y, w, h, fill: { color: fill || C.background2 }, margin: [12, 12, 0, 0], valign: "middle", objectName: name });
  };
  const by = 2.35, bh = 1.35;
  box(MX, by, 2.5, bh, "1. Prepare", "|+⟩^{⊗8}: every selection equally likely (= random guess)", "wf-prepare");
  card(s, 3.45, 1.95, 5.6, 2.15, "wf-layer-frame", "D7ECE9");
  s.addText("repeated p times (\"layers\")", { x: 3.6, y: 2.0, w: 5.3, h: 0.3, fontSize: 12, italic: true, color: C.text2, margin: 0, isTextBox: true, objectName: "wf-layer-label" });
  box(3.6, by, 2.55, bh, "2. Cost layer", "exp(−iγ_{k} H_{C}): phases favor low f(x)", "wf-cost", C.background1);
  box(6.35, by, 2.55, bh, "3. Mixer layer", "exp(−iβ_{k} Σ X_{i}): moves probability between selections", "wf-mixer", C.background1);
  box(9.35, by, 3.38, bh, "4. Measure", "samples x, estimate the energy ⟨H_{C}⟩", "wf-measure");
  s.addImage({ data: icons.arrow, x: 3.13, y: by + 0.55, w: 0.25, h: 0.25, altText: "" });
  s.addImage({ data: icons.arrow, x: 9.07, y: by + 0.55, w: 0.25, h: 0.25, altText: "" });
  box(3.45, 4.6, 9.28, 1.2, "5. Classical optimizer (COBYLA)", "updates the 2p angles (γ, β) to lower ⟨H_{C}⟩, then steps 2 to 4 run again", "wf-optimizer", C.text2, C.background1);
  s.addImage({ data: icons.redo, x: 2.55, y: 4.95, w: 0.5, h: 0.5, altText: "Loop back to the cost layer" });
  s.addText("Hybrid: the quantum computer only evaluates the energy, the classical computer steers the angles.", { x: MX, y: 6.15, w: CW, h: 0.5, fontSize: 16,
    italic: true, color: C.text2, margin: 0, isTextBox: true, objectName: "wf-hybrid" });
  s.addNotes("QAOA in four steps. Start from the uniform superposition: every selection equally likely, which is exactly random guessing. Then alternate p times a cost layer, which adds phases that depend on f(x), and a mixer layer, which moves probability between selections. Measure and estimate the average energy. A classical optimizer, COBYLA, updates the angles and the loop repeats. This is why QAOA is called hybrid.");

  // ---------------------------------------------------- 12. gates
  s = contentSlide(2, "Gates: RZ per order, RZZ per conflict, RX per qubit", "Quantum workflow");
  card(s, MX, 1.95, CW, 3.2, "circuit-card", C.background1);
  s.addImage({ path: IMG("circuit.png"), x: MX + 0.1, y: 2.0, w: CW - 0.2, h: (CW - 0.2) * 698 / 2882, altText: "One QAOA layer on 8 qubits: H gates, RZ per order, RZZ per conflict, RX mixer" });
  const gm = [["H (shown as U)", "prepares |+⟩ on each of the 8 qubits"], ["RZ(2γhᵢ)", "one per order, from the Z fields"],
    ["RZZ(2γ·A/4)", "one per conflict: 10 per layer = 20 CNOTs"], ["RX(2β)", "the mixer, one per qubit"]];
  const gw = (CW - 3 * 0.25) / 4;
  gm.forEach(([g, d], i) => {
    const x = MX + i * (gw + 0.25);
    card(s, x, 5.3, gw, 1.4, `gate-card-${i + 1}`);
    s.addText([{ text: g, options: { bold: true, fontSize: 16, color: C.accent1, breakLine: true } }, { text: d, options: { fontSize: 13, color: "1B2A33" } }],
      { x: x + 0.2, y: 5.35, w: gw - 0.4, h: 1.3, margin: 0, valign: "middle", isTextBox: true, objectName: `gate-text-${i + 1}` });
  });
  s.addNotes("One layer of the circuit on 8 qubits, one per order. Hadamards prepare the superposition. The Z fields become RZ rotations, one per order. Each conflict becomes an RZZ rotation, which on hardware is two CNOTs around an RZ, so 20 CNOTs per layer. The mixer is one RX per qubit. Two practical choices in the notebook: we divide the QUBO by A so energies are of order 1 and the angles live in a sensible range, and we run COBYLA from 5 random starting points and keep the best.");

  // ---------------------------------------------------- 13. headline results
  pres.addSection({ title: "QAOA results" });
  s = contentSlide(3, `2-layer QAOA: ${pct(D.qaoa["valid rate"])} of samples are valid plans`, "QAOA results");
  const keys = ["valid rate", "P(optimum)", "approx. ratio"];
  const klab = ["Valid rate", "P(optimum)", "Approx. ratio"];
  const fmt = [(v) => pct(v), (v) => pct(v, 1), (v) => v.toFixed(2)];
  keys.forEach((k, i) => {
    const y = 1.95 + i * 1.6;
    card(s, MX, y, 5.4, 1.4, `headline-card-${i + 1}`);
    s.addText(klab[i], { x: MX + 0.3, y: y + 0.12, w: 4.8, h: 0.35, fontSize: 14, bold: true, color: MUTED, margin: 0, isTextBox: true, objectName: `headline-label-${i + 1}` });
    s.addText([{ text: fmt[i](D.random[k]), options: { color: MUTED, fontSize: 24 } }, { text: "  →  ", options: { color: MUTED, fontSize: 24 } },
      { text: fmt[i](D.qaoa[k]), options: { color: C.accent1, fontSize: 36, bold: true } }],
      { x: MX + 0.3, y: y + 0.45, w: 4.8, h: 0.8, margin: 0, valign: "middle", fontFace: THEME.headFontFace, isTextBox: true, objectName: `headline-value-${i + 1}` });
  });
  s.addChart(pres.charts.BAR, [
    { name: "Random guessing", labels: klab, values: keys.map((k) => D.random[k]) },
    { name: "QAOA, 2 layers (emulator)", labels: klab, values: keys.map((k) => D.qaoa[k]) },
  ], { x: 6.4, y: 1.85, w: 6.33, h: 4.9, barDir: "col", ...chartBase(), chartColors: [HEX.accent5, HEX.accent1], showValue: true, dataLabelPosition: "outEnd",
    dataLabelFormatCode: "0.00", valAxisMinVal: 0, valAxisMaxVal: 1, valAxisLabelFormatCode: "0.0", showLegend: true, legendPos: "t",
    showTitle: true, title: "Same instance, ideal emulator", objectName: "headline-chart" });
  s.addNotes(`On the same 8-order instance, with 2 layers on the ideal emulator. The valid rate goes from 16 to ${pct(D.qaoa["valid rate"])}. The probability of measuring the optimum goes from 0.4 to ${pct(D.qaoa["P(optimum)"], 1)}, about 10 times better than random. The expected profit reaches half of the optimum, counting invalid plans as zero. These are exact probabilities from the final state, so no shot noise.`);

  // ---------------------------------------------------- 14. distribution
  s = contentSlide(3, "But the optimum is only the 6th most likely outcome", "QAOA results");
  const tl = D.top.map((t) => `{${t.orders.join(",")}}`);
  const tcol = D.top.map((t) => (t.opt ? HEX.dk2 : t.valid ? HEX.accent1 : HEX.accent2));
  s.addChart(pres.charts.BAR, [{ name: "probability", labels: tl, values: D.top.map((t) => t.p) }], {
    x: MX, y: 1.85, w: 8.4, h: 4.9, barDir: "col", ...chartBase(), chartColors: tcol, showValue: false, showLegend: false,
    valAxisLabelFormatCode: "0%", showTitle: true, title: "15 most likely selections (dark = optimum, all 15 are valid)",
    catAxisLabelRotate: -45, catAxisLabelFontSize: 11, objectName: "top15-chart",
  });
  card(s, 9.35, 1.95, 3.38, 4.75, "dist-notes-card");
  s.addText(bullets([
    "All 15 most likely outcomes are valid plans",
    "The top two, {1,6,7} and {1,4,7}, earn 13 and 14, not 17",
    "Qiskit reports the best sample, so it does return {1,4,5,8}: P(optimum) is the fairer measure",
  ], 15), { x: 9.6, y: 2.15, w: 2.95, h: 4.4, color: "1B2A33", margin: 0, valign: "top", objectName: "dist-notes" });
  s.addNotes("Looking inside the distribution. All 15 most likely outcomes are valid, so QAOA has learned the conflict rule. But it spreads probability over many good plans: the optimum is only 6th. MinimumEigenOptimizer reports the best sample it saw, which is the optimum, but that hides how often you would actually get it. That is why we report P(optimum) rather than just the returned answer.");

  // ---------------------------------------------------- 15. layers
  pres.addSection({ title: "Critical analysis" });
  s = contentSlide(4, "More layers did not help: 2 layers beat 3 and 4", "Critical analysis");
  const reps = D.layers.map((r) => `p = ${r.reps}`);
  s.addChart(pres.charts.LINE, ["valid rate", "approx. ratio", "P(optimum)"].map((k) => ({ name: k, labels: reps, values: D.layers.map((r) => r[k]) })), {
    x: MX, y: 1.85, w: 7.4, h: 4.9, ...chartBase(), chartColors: [HEX.accent1, HEX.dk2, HEX.accent2], lineSize: 2.5, lineDataSymbolSize: 8,
    showValue: true, dataLabelPosition: "t", dataLabelFormatCode: "0.00", valAxisMinVal: 0, valAxisMaxVal: 1, valAxisLabelFormatCode: "0.0",
    showLegend: true, legendPos: "t", showTitle: true, title: "Solution quality vs number of layers (emulator)", objectName: "layers-chart",
  });
  s.addTable([[th("Layers"), th("Angles"), th("CNOTs"), th("Time")],
    ...D.layers.map((r) => [td(String(r.reps), { align: "center" }), td(String(r.parameters), { align: "center" }), td(String(20 * r.reps), { align: "center" }), td(`${r["time (s)"].toFixed(1)} s`, { align: "center" })])],
    { x: 8.4, y: 1.95, w: 4.33, colW: [1.0, 1.0, 1.1, 1.233], fontSize: 14, rowH: 0.42, border: { type: "solid", pt: 0.75, color: "DDE3E6" }, valign: "middle", objectName: "layers-table" });
  s.addText(bullets([
    "Theory: with the best angles, more layers can only help",
    "Practice: COBYLA must tune 2p angles and gets stuck in poorer local minima",
    "Each layer adds 20 CNOTs, so more noise on hardware",
  ], 14), { x: 8.4, y: 4.25, w: 4.33, h: 2.5, color: "1B2A33", margin: 0, valign: "top", objectName: "layers-notes" });
  s.addNotes("We swept the number of layers from 1 to 4. In theory, with optimal angles, adding layers can never hurt, and as p goes to infinity QAOA approaches adiabatic evolution. In practice 2 layers gave the best result: with more layers there are more angles to tune and COBYLA lands in worse local minima. Time and circuit depth also grow linearly, which costs even more on real hardware.");

  // ---------------------------------------------------- 16. penalty
  s = contentSlide(4, "A bigger penalty keeps plans valid, hides the best", "Critical analysis");
  const plab = D.penalty.map((r) => `A = ${Math.round(r.A)}`);
  s.addChart(pres.charts.BAR, [
    { name: "valid rate", labels: plab, values: D.penalty.map((r) => r["valid rate"]) },
    { name: "P(optimum)", labels: plab, values: D.penalty.map((r) => r["P(optimum)"]) },
  ], { x: MX, y: 1.85, w: 7.6, h: 4.9, barDir: "col", ...chartBase(), chartColors: [HEX.accent1, HEX.accent2], showValue: true, dataLabelPosition: "outEnd",
    dataLabelFormatCode: "0.00", valAxisMinVal: 0, valAxisMaxVal: 1, valAxisLabelFormatCode: "0.0", showLegend: true, legendPos: "t",
    showTitle: true, title: "Effect of the penalty weight (2 layers, emulator)", objectName: "penalty-chart" });
  const pFirst = D.penalty[0]["P(optimum)"], pLast = D.penalty[D.penalty.length - 1]["P(optimum)"];
  stat(s, 8.65, 1.95, 4.0, `${pct(pFirst, 1)} → ${pct(pLast, 1)}`, "P(optimum) from A = 7 to A = 363", "pen-popt", C.accent2);
  s.addText(bullets([
    "Every A > 6 gives the same optimum: brute force cannot tell them apart",
    "After rescaling, a large A shrinks the profit terms next to the penalty",
    "QAOA still avoids conflicts, but stops separating good plans from the best one",
    "Use the smallest valid penalty, not the converter's default",
  ], 14), { x: 8.65, y: 3.6, w: 4.08, h: 3.1, color: "1B2A33", margin: 0, valign: "top", objectName: "penalty-notes" });
  s.addNotes("This is where the converter comparison matters. Every penalty above 6 gives the same optimum, so brute force cannot tell them apart. For QAOA they are very different. With a large A, the profit differences become tiny next to the penalty: QAOA still learns to avoid conflicts, the valid rate stays high, but it can no longer distinguish good plans from the best one. P(optimum) drops from about 4 percent to 0.6 percent with the converter's penalty of 363.");

  // ---------------------------------------------------- 17. size
  s = contentSlide(4, "More orders: fixed-depth QAOA loses its edge", "Critical analysis");
  const nl = D.size.map((r) => `n = ${r.n}`);
  s.addChart(pres.charts.LINE, [
    { name: "QAOA, 2 layers", labels: nl, values: D.size.map((r) => r["P(optimum)"]) },
    { name: "random guessing", labels: nl, values: D.size.map((r) => r["random P(opt)"]) },
  ], { x: MX, y: 1.85, w: 7.4, h: 4.9, ...chartBase(), chartColors: [HEX.accent1, HEX.accent5], lineSize: 2.5, lineDataSymbolSize: 8,
    valAxisLogScaleBase: 10, valAxisMinVal: 0.0001, valAxisMaxVal: 1, valAxisLabelFormatCode: "0.00%", showLegend: true, legendPos: "t",
    showTitle: true, title: "P(optimum) on random conflict graphs (log scale)", objectName: "size-chart" });
  s.addTable([[th("n"), th("Conflicts"), th("Brute force"), th("QAOA sim.")],
    ...D.size.map((r) => [td(String(r.n), { align: "center" }), td(String(r.conflicts), { align: "center" }),
      td(`${(r["brute force (s)"] * 1000).toFixed(1)} ms`, { align: "center" }), td(`${r["QAOA (s)"].toFixed(1)} s`, { align: "center" })])],
    { x: 8.4, y: 1.95, w: 4.33, colW: [0.7, 1.1, 1.3, 1.233], fontSize: 13, rowH: 0.36, border: { type: "solid", pt: 0.75, color: "DDE3E6" }, valign: "middle", objectName: "size-table" });
  s.addText(bullets([
    "Brute force doubles per order: about 10¹⁵ cases at n = 50",
    "Simulating QAOA is exponential too (2ⁿ amplitudes): a cost of the emulator, not of the hardware",
    "At n = 10, QAOA barely beats guessing: instances differ, not only by size",
  ], 13), { x: 8.4, y: 4.3, w: 4.33, h: 2.45, color: "1B2A33", margin: 0, valign: "top", objectName: "size-notes" });
  s.addNotes("We generated random conflict graphs from 4 to 12 orders. Brute force doubles with each order: it is still instant here but reaches 10^15 cases at 50 orders. Simulating QAOA is also exponential, because the emulator stores 2^n amplitudes; on hardware one circuit costs about p times the number of conflicts in gates. With a fixed number of layers, P(optimum) falls with size and on the 10-order instance QAOA barely beats random guessing. Keeping quality would require more layers, so deeper circuits.");

  // ---------------------------------------------------- 18. hardware
  s = contentSlide(4, "On a real chip: bigger circuit, noisier results", "Critical analysis");
  const T = D.transpile;
  card(s, MX, 1.95, 5.2, 2.35, "transpile-card");
  s.addText("Transpiled for IBM Torino (heavy-hex layout)", { x: MX + 0.3, y: 2.05, w: 4.6, h: 0.4, fontSize: 14, bold: true, color: MUTED, margin: 0, isTextBox: true, objectName: "transpile-heading" });
  stat(s, MX + 0.3, 2.45, 2.2, `${T.logical_2q} → ${T.isa_2q}`, "two-qubit gates", "tq-gates", C.accent2, 32);
  stat(s, MX + 2.7, 2.45, 2.3, `${T.logical_depth} → ${T.isa_depth}`, "circuit depth", "tq-depth", C.accent2, 32);
  s.addText("The conflict graph contains squares; the chip has none, so the transpiler inserts SWAP gates to bring conflicting orders next to each other.",
    { x: MX, y: 4.45, w: 5.2, h: 1.0, fontSize: 14, color: "1B2A33", margin: 0, valign: "top", isTextBox: true, objectName: "swap-note" });
  card(s, MX, 5.6, 5.2, 1.1, "hw-pending-card", "FBE3DC");
  s.addText([{ text: "Real IBM run: to complete. ", options: { bold: true, color: C.accent2 } },
    { text: "Same angles, 4096 shots; set RUN_ON_HARDWARE = True in section 8 of the notebook." }],
    { x: MX + 0.25, y: 5.65, w: 4.7, h: 1.0, fontSize: 13, color: "1B2A33", margin: 0, valign: "middle", isTextBox: true, objectName: "hw-pending" });
  s.addChart(pres.charts.BAR, [
    { name: "Random guessing", labels: klab, values: keys.map((k) => D.random[k]) },
    { name: "Ideal emulator", labels: klab, values: keys.map((k) => D.qaoa[k]) },
    { name: "Noisy emulator (Torino noise model)", labels: klab, values: keys.map((k) => D.noisy[k]) },
  ], { x: 6.2, y: 1.85, w: 6.53, h: 4.9, barDir: "col", ...chartBase(), chartColors: [HEX.accent5, HEX.accent1, HEX.accent2], showValue: true,
    dataLabelPosition: "outEnd", dataLabelFormatCode: "0.00", valAxisMinVal: 0, valAxisMaxVal: 1, valAxisLabelFormatCode: "0.0", showLegend: true, legendPos: "t",
    showTitle: true, title: "Same circuit, same angles: effect of noise", objectName: "noise-chart" });
  s.addNotes(`We reuse the angles from the emulator and only sample, so the comparison isolates the hardware. Transpiled for IBM Torino, the circuit goes from ${T.logical_2q} to ${T.isa_2q} two-qubit gates and from depth ${T.logical_depth} to ${T.isa_depth}: the squares in our conflict graph do not exist on the heavy-hex chip, so SWAPs are needed. With Torino's noise model, the valid rate falls from ${pct(D.qaoa["valid rate"])} to ${pct(D.noisy["valid rate"])}: noise pushes the distribution back towards uniform. The valid rate is the most fragile metric, since one bit flip on a conflicting pair breaks a plan. TO UPDATE once the real IBM run is done.`);

  // ---------------------------------------------------- 19. conclusion
  pres.addSection({ title: "Conclusion" });
  s = pres.addSlide({ masterName: "CLOSING_DARK", sectionTitle: "Conclusion" });
  s.addText("What this project shows, and what it does not", { placeholder: "title" });
  const colw = (CW - 0.4) / 2;
  const col = (x, ic, color, head, items, name) => {
    s.addShape(pres.shapes.OVAL, { x, y: 1.75, w: 0.55, h: 0.55, fill: { color }, line: { type: "none" }, objectName: name + "-icon-bg" });
    s.addImage({ data: ic, x: x + 0.14, y: 1.89, w: 0.27, h: 0.27, altText: head });
    s.addText(head, { x: x + 0.75, y: 1.75, w: colw - 0.75, h: 0.55, fontSize: 20, bold: true, color: C.background1, margin: 0, valign: "middle", isTextBox: true, objectName: name + "-head" });
    s.addText(bullets(items, 18), { x, y: 2.55, w: colw, h: 2.9, color: "E6F0EE", margin: 0, valign: "top", objectName: name + "-list" });
  };
  col(MX, icons.check, C.accent1, "Shows", [
    "The full chain works: workshop orders → MWIS → QUBO with one qubit per order → Ising → QAOA",
    "Modeling choices matter before any quantum code runs: 18 vs 8 qubits, penalty 363 vs 7",
    `QAOA beats random guessing clearly: ${pct(D.qaoa["valid rate"])} valid, about 10× the chance of the optimum`,
  ], "shows");
  col(MX + colw + 0.4, icons.cross, C.accent2, "Does not show", [
    "Any speedup: brute force solves this instance in about 1 ms",
    `That QAOA reliably returns the optimum: ${pct(D.qaoa["P(optimum)"], 1)}, and falling with size`,
    "That today's noisy hardware can handle industrial sizes (thousands of orders)",
  ], "doesnot");
  card(s, MX, 5.6, CW, 1.1, "next-card", "1F3B45");
  s.addText([{ text: "Next steps  ", options: { bold: true, color: C.accent3 } },
    { text: "CVaR-QAOA (optimize the best samples, not the mean)  •  warm start from a greedy plan  •  mixers that only explore valid plans, removing the penalty" }],
    { x: MX + 0.3, y: 5.65, w: CW - 0.6, h: 1.0, fontSize: 15, color: C.background1, margin: 0, valign: "middle", isTextBox: true, objectName: "next-text" });
  s.addNotes("To conclude. We carried a real decision problem through the whole chain, and learned that the modeling choices, qubit count and penalty, matter as much as the algorithm. QAOA clearly beats random guessing on this instance. But we do not claim a speedup: brute force takes a millisecond, QAOA returns the optimum only about 4 percent of the time, and that drops with size. Next steps we would explore: CVaR-QAOA, warm starts, and constraint-preserving mixers. Thank you, questions welcome.");

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
}

function PROFITS() { return [5, 3, 4, 6, 2, 5, 3, 4]; }

main().catch((e) => { console.error(e); process.exit(1); });
