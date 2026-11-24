// Deck built from WorkshopOrders.ipynb. Usage: node build_deck.js <out.pptx>
// Figures come from make_assets.py; every number is copied from the notebook outputs.
const path = require("path");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const { applyTheme } = require(process.env.PPTX_SKILL + "/scripts/apply_theme.js");

const OUT = process.argv[2] || "deck.pptx";
const IMG = (f) => path.join(__dirname, f);

// Same theme as ../build_deck.js so both decks read as one set.
const THEME = {
  name: "Workshop Orders QAOA",
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
const INK = "1B2A33";

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.title = "Optimisation du profit d'un atelier avec QAOA";
pres.subject = "Tech 5 Quantum Computing - soutenance";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
const C = pres.SchemeColor;

const W = 13.333;
const MX = 0.6;
const CW = W - 2 * MX;

// ------------------------------------------------------------- layouts
pres.defineSlideMaster({
  title: "TITLE_DARK",
  background: { color: C.text2 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: MX, y: 1.9, w: 7.2, h: 2.3, fontSize: 40, bold: true,
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
    { text: { text: "Optimisation du profit d'un atelier avec QAOA  •  Tech 5 Quantum Computing",
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

// Agenda tracker: the repeated motif, one pill per part of the talk.
const PARTS = ["Problème", "Recherche exhaustive", "QAOA", "Résultats", "Analyse"];
function tracker(slide, active) {
  const pw = 2.3, gap = 0.12, h = 0.32;
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

// Multi-line formula: one paragraph per "\n" so subscripts never split a line.
function lines(src, base) {
  const parts = src.split("\n");
  return parts.flatMap((l, i) => {
    const runs = formula(l, base);
    if (i < parts.length - 1) runs[runs.length - 1].options.breakLine = true;
    return runs;
  });
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
const B = (t) => ({ text: t, options: { bold: true, color: C.text2 } });
const T = (t) => ({ text: t });

const CHART_FONT = { catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt",
  legendFontFace: "+mn-lt", titleFontFace: "+mn-lt", catAxisTitleFontFace: "+mn-lt", valAxisTitleFontFace: "+mn-lt" };
const chartBase = () => ({
  ...CHART_FONT, catAxisLabelColor: MUTED, valAxisLabelColor: MUTED, catAxisLabelFontSize: 12, valAxisLabelFontSize: 12,
  dataLabelFontSize: 11, dataLabelColor: INK, legendFontSize: 12, legendColor: INK,
  valGridLine: { color: "DDE3E6", size: 0.75 }, catGridLine: { style: "none" }, titleFontSize: 14, titleColor: INK,
  catAxisTitleColor: MUTED, valAxisTitleColor: MUTED, catAxisTitleFontSize: 12, valAxisTitleFontSize: 12,
});

const th = (t, o = {}) => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.text2 }, ...o } });
const td = (t, o = {}) => ({ text: t, options: { color: INK, ...o } });
const tableOpts = (o) => ({ border: { type: "solid", pt: 0.75, color: "DDE3E6" }, margin: [0.03, 0.1, 0.03, 0.1], valign: "middle", ...o });

async function main() {
  const icons = {
    tower: await icon(fa.FaBroadcastTower, HEX.accent1), truck: await icon(fa.FaTruck, HEX.accent1),
    calendar: await icon(fa.FaCalendarAlt, HEX.accent1), chart: await icon(fa.FaChartLine, HEX.accent1),
    arrow: await icon(fa.FaArrowRight, HEX.accent5), redo: await icon(fa.FaRedo, HEX.accent2),
    clock: await icon(fa.FaHourglassHalf, HEX.lt1), memory: await icon(fa.FaMemory, HEX.lt1), noise: await icon(fa.FaWaveSquare, HEX.lt1),
    warn: await icon(fa.FaExclamationTriangle, HEX.accent2),
  };

  // ---------------------------------------------------- 1. titre
  pres.addSection({ title: "Introduction" });
  let s = pres.addSlide({ masterName: "TITLE_DARK", sectionTitle: "Introduction" });
  s.addText("Optimisation du profit d'un atelier avec QAOA", { placeholder: "title" });
  s.addText("Du choix des commandes au circuit quantique : recherche exhaustive, QAOA émulé et QAOA sur IBM ibm_fez\nTech 5 Quantum Computing  •  Soutenance",
    { placeholder: "body" });
  card(s, 8.3, 1.9, 4.45, 3.7, "title-graph-card", C.background1);
  s.addImage({ path: IMG("graph_opt.png"), x: 8.45, y: 2.25, w: 4.15, h: 4.15 * 608 / 1290, altText: "Graphe de conflits des 8 commandes, sélection optimale en orange" });
  s.addText("8 commandes, 10 conflits : accepter 1, 4, 5 et 8 rapporte un profit de 17", { x: 8.5, y: 4.4, w: 4.05, h: 0.9, fontSize: 13,
    color: MUTED, align: "center", valign: "middle", isTextBox: true, objectName: "title-graph-caption" });
  s.addNotes("Présentation du sujet. Un atelier reçoit plus de commandes qu'il ne peut en traiter, et certaines se disputent les mêmes ressources. On suit toute la chaîne : problème industriel, modèle mathématique, QUBO, Hamiltonien, circuit QAOA, puis comparaison entre recherche exhaustive, QAOA émulé et QAOA exécuté sur une vraie machine IBM. L'image de droite est notre instance et sa réponse optimale.");

  // ---------------------------------------------------- 2. le problème
  pres.addSection({ title: "Problème" });
  s = contentSlide(0, "Un atelier reçoit plus de commandes qu'il ne peut en traiter", "Problème");
  s.addText("Chaque commande rapporte un profit et mobilise des ressources. Une ressource ne sert qu'une commande à la fois.",
    { x: MX, y: 1.95, w: 6.0, h: 0.85, fontSize: 18, color: INK, valign: "top", margin: 0, isTextBox: true, objectName: "context" });
  [["Décision", "à quelles commandes répondre ?"], ["Objectif", "maximiser le profit total"], ["Règle", "jamais deux commandes qui partagent une ressource"]].forEach(([k, v], i) => {
    const y = 3.0 + i * 0.75;
    s.addText(k, { shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.1, x: MX, y, w: 1.4, h: 0.5, fill: { color: C.accent1 },
      color: C.background1, fontSize: 15, bold: true, align: "center", valign: "middle", margin: 0, objectName: `key-${i + 1}` });
    s.addText(v, { x: MX + 1.6, y, w: 4.6, h: 0.5, fontSize: 18, color: INK, valign: "middle", margin: 0, isTextBox: true, objectName: `val-${i + 1}` });
  });
  const resources = [["Machine", [["Fraiseuse CNC", "1, 2"], ["Découpeuse laser", "3, 4"], ["Cabine de peinture", "5, 6"], ["Tour", "7, 8"]]],
    ["Ouvrier", [["Soudeur certifié", "1, 3"], ["Inspecteur qualité", "4, 6"], ["Électricien", "5, 7"]]],
    ["Créneau", [["Lundi", "2, 4"], ["Mercredi", "3, 5"], ["Vendredi", "6, 8"]]]];
  const resRows = [[th("Type"), th("Ressource (capacité 1)"), th("Commandes", { align: "center" })]];
  resources.forEach(([type, items], gi) => items.forEach(([name, orders], k) => resRows.push([
    ...(k === 0 ? [td(type, { rowspan: items.length, bold: true, color: C.text2, fill: { color: gi % 2 ? "FFFFFF" : HEX.lt2 } })] : []),
    td(name), td(orders, { align: "center", bold: true, color: C.accent2 })])));
  s.addTable(resRows, tableOpts({ x: 7.1, y: 1.95, w: 5.63, colW: [1.3, 2.73, 1.6], fontSize: 13, rowH: 0.3, objectName: "resource-table" }));
  card(s, MX, 5.55, CW, 1.15, "combinatorics-card", C.text2);
  s.addText([{ text: "2", options: { fontSize: 36, bold: true, color: C.accent3 } }, { text: "n", options: { fontSize: 36, bold: true, color: C.accent3, superscript: true } }],
    { x: MX + 0.3, y: 5.6, w: 0.9, h: 1.05, margin: 0, valign: "middle", fontFace: THEME.headFontFace, isTextBox: true, objectName: "combinatorics-value" });
  s.addText("sélections possibles (acceptée ou refusée par commande) : 256 pour nos 8 commandes, 65 536 pour 16, environ 10¹⁵ pour 50.",
    { x: MX + 1.25, y: 5.6, w: CW - 1.55, h: 1.05, fontSize: 17, color: C.background1, margin: 0, valign: "middle", isTextBox: true, objectName: "combinatorics-text" });
  s.addNotes("L'atelier a 4 machines, 3 ouvriers et 3 créneaux, chacun utilisable par une seule commande. Il doit choisir parmi 8 commandes. Le tableau donne, pour chaque ressource, les deux commandes qui la demandent : ce sont nos 10 conflits. Chaque commande est acceptée ou refusée, donc il y a 2^n sélections possibles, ce qui double à chaque commande ajoutée.");

  // ---------------------------------------------------- 3. MWIS
  s = contentSlide(0, "Choisir les commandes = trouver un stable de poids maximum", "Problème");
  card(s, MX, 1.95, 6.4, 3.5, "graph-card");
  s.addImage({ path: IMG("graph.png"), x: MX + 0.25, y: 2.05, w: 5.9, h: 5.9 * 608 / 1290, altText: "Graphe de conflits : 8 commandes en sommets, 10 conflits en arêtes, en forme d'échelle" });
  s.addText("Sommet = commande (profit p)  •  arête = conflit de ressource", { x: MX + 0.25, y: 4.9, w: 5.9, h: 0.4, fontSize: 13, color: MUTED,
    align: "center", margin: 0, isTextBox: true, objectName: "graph-caption" });
  card(s, MX, 5.65, 6.4, 1.05, "formula-card", C.text2);
  s.addText(formula("max Σ_{i} p_{i} x_{i}   sous   x_{i} x_{j} = 0  ∀ (i, j) ∈ E,   x_{i} ∈ {0, 1}", { color: C.background1, fontSize: 20 }),
    { x: MX + 0.3, y: 5.7, w: 5.8, h: 0.95, margin: 0, valign: "middle", fontFace: THEME.headFontFace, isTextBox: true, objectName: "mwis-formula" });
  s.addText([{ text: "Maximum Weight Independent Set (MWIS)", options: { bold: true, fontSize: 18, color: C.text2, breakLine: true } },
    { text: "Une sélection valide ne contient aucune arête. Le meilleur profit est un stable de poids maximum : un problème NP-difficile.", options: { fontSize: 15, color: INK } }],
    { x: 7.35, y: 1.95, w: 5.38, h: 1.35, margin: 0, valign: "top", isTextBox: true, objectName: "mwis-definition" });
  s.addText("Le même modèle en industrie", { x: 7.35, y: 3.5, w: 5.38, h: 0.4, fontSize: 16, bold: true, color: C.text2, margin: 0, isTextBox: true, objectName: "uses-heading" });
  const uses = [["calendar", "Production", "commandes ou tâches qui se disputent machines et équipes"], ["tower", "Télécoms", "émetteurs voisins qui ne partagent pas une fréquence"],
    ["truck", "Logistique", "livraisons en concurrence pour un quai ou un véhicule"], ["chart", "Finance", "projets ou ordres mutuellement exclusifs"]];
  uses.forEach(([ic, t, d], i) => {
    const y = 4.0 + i * 0.68;
    s.addImage({ data: icons[ic], x: 7.35, y: y + 0.1, w: 0.3, h: 0.3, altText: t });
    s.addText([{ text: t + " : ", options: { bold: true, color: C.text2 } }, { text: d, options: { color: INK } }],
      { x: 7.8, y, w: 4.93, h: 0.55, fontSize: 14, margin: 0, valign: "middle", isTextBox: true, objectName: `use-${i + 1}` });
  });
  s.addNotes("Les commandes deviennent des sommets, les conflits des arêtes. Une sélection valide est un ensemble de sommets sans arête entre eux : un stable. Le profit maximal correspond au stable de poids maximum, MWIS, NP-difficile. x_i vaut 1 si la commande i est acceptée, et x_i x_j = 0 exprime exactement l'interdiction d'accepter deux commandes en conflit. Le même modèle sert en production, télécoms, logistique et finance.");

  // ---------------------------------------------------- 4. recherche exhaustive
  pres.addSection({ title: "Recherche exhaustive" });
  s = contentSlide(1, "La recherche exhaustive donne la vérité terrain en millisecondes", "Recherche exhaustive");
  const exRows = [
    ["Atelier de base", "8 commandes", "256", "41\u00A0(16,0\u00A0%)", "17", "0,123\u00A0ms"],
    ["Atelier x2", "16 commandes", "65\u00A0536", "1\u00A0681\u00A0(2,6\u00A0%)", "34", "27,9\u00A0ms"],
  ];
  const heads = ["Sélections testées (2ⁿ)", "Sélections valides", "Profit optimal", "Durée moyenne"];
  heads.forEach((h, k) => s.addText(h, { x: MX + 2.5 + k * 2.4, y: 1.95, w: 2.3, h: 0.4, fontSize: 13, color: MUTED, bold: true,
    margin: 0, valign: "bottom", isTextBox: true, objectName: `ex-head-${k + 1}` }));
  exRows.forEach((r, i) => {
    const y = 2.5 + i * 1.3;
    card(s, MX, y, CW, 1.1, `ex-row-${i + 1}`, HEX.lt2);
    s.addText([{ text: r[0], options: { bold: true, color: C.text2, fontSize: 16, breakLine: true } }, { text: r[1], options: { color: MUTED, fontSize: 13 } }],
      { x: MX + 0.25, y, w: 2.0, h: 1.1, margin: 0, valign: "middle", isTextBox: true, objectName: `ex-name-${i + 1}` });
    r.slice(2).forEach((v, k) => s.addText(v, { x: MX + 2.5 + k * 2.4, y, w: 2.3, h: 1.1, fontSize: k === 1 ? 22 : 26, bold: true,
      color: k === 2 ? C.accent2 : C.accent1, fontFace: THEME.headFontFace, margin: 0, valign: "middle", isTextBox: true, objectName: `ex-val-${i + 1}-${k + 1}` }));
  });
  s.addText([B("Solution optimale : "), T("commandes [1, 4, 5, 8] pour l'atelier de base, [1, 4, 5, 8, 9, 12, 13, 16] pour l'atelier x2.")],
    { x: MX, y: 5.2, w: CW, h: 0.45, fontSize: 15, color: INK, margin: 0, valign: "middle", isTextBox: true, objectName: "ex-solution" });
  card(s, MX, 5.8, CW, 0.9, "ex-growth-card", C.text2);
  s.addText([{ text: "×2 commandes → ×256 candidats → ×227 en durée. ", options: { bold: true, color: C.accent3 } },
    { text: "Environ 0,43 µs par candidat, sur un cœur (Intel Core Ultra 7 155H, moyenne sur 10 000 exécutions).", options: { color: C.background1 } }],
    { x: MX + 0.3, y: 5.8, w: CW - 0.6, h: 0.9, fontSize: 16, margin: 0, valign: "middle", isTextBox: true, objectName: "ex-growth" });
  s.addNotes("L'algorithme exhaustif teste les 2^n sélections, écarte celles qui violent un conflit et garde la plus rentable. Sur 8 commandes : 256 candidats, 41 valides, profit 17 avec les commandes 1, 4, 5, 8, en 0,12 ms. On double ensuite l'atelier en juxtaposant deux copies : 65 536 candidats, profit 34, 28 ms. La durée est multipliée par 227, proche des 256 attendus : c'est la croissance exponentielle. Ce résultat sert de référence pour juger QAOA.");

  // ---------------------------------------------------- 5. QUBO -> Hamiltonien
  pres.addSection({ title: "QAOA" });
  s = contentSlide(2, "Du modèle au Hamiltonien : la pénalité remplace la contrainte", "QAOA");
  const steps5 = [
    ["1  QUBO", "f(x) = −Σ_{i} p_{i} x_{i} + A Σ_{(i,j)∈E} x_{i} x_{j}", "A = 7 > max p_{i} = 6 : une sélection en conflit n'est jamais optimale."],
    ["2  Substitution", "x_{i} = (1 − Z_{i}) / 2", "Z_{i} vaut +1 si la commande est refusée, −1 si elle est acceptée."],
    ["3  Hamiltonien d'Ising", "H = Σ h_{i} Z_{i} + Σ J_{ij} Z_{i} Z_{j} + C", "J_{ij} = A/4 = 1,75\nh_{i} = p_{i}/2 − A d_{i}/4\nC = 1,5"],
  ];
  const w5 = (CW - 2 * 0.4) / 3;
  steps5.forEach(([t, f, d], i) => {
    const x = MX + i * (w5 + 0.4);
    card(s, x, 1.95, w5, 3.0, `step5-${i + 1}`);
    s.addText(t, { x: x + 0.25, y: 2.1, w: w5 - 0.5, h: 0.45, fontSize: 18, bold: true, color: C.text2, margin: 0, isTextBox: true, objectName: `step5-title-${i + 1}` });
    s.addText(formula(f, { fontSize: i === 0 ? 15 : 17, color: INK }), { x: x + 0.25, y: 2.65, w: w5 - 0.5, h: 1.0, margin: 0, valign: "middle",
      fontFace: THEME.headFontFace, isTextBox: true, objectName: `step5-formula-${i + 1}` });
    s.addText(lines(d, { fontSize: 14, color: MUTED }), { x: x + 0.25, y: 3.75, w: w5 - 0.5, h: 1.05, margin: 0, valign: "top", isTextBox: true, objectName: `step5-desc-${i + 1}` });
    if (i < 2) s.addImage({ data: icons.arrow, x: x + w5 + 0.11, y: 3.3, w: 0.18, h: 0.18, altText: "" });
  });
  [["8", "qubits, un par commande"], ["18", "termes de Pauli (8 Z, 10 ZZ)"], ["H|x⟩ = f(x)|x⟩", "l'état fondamental est la meilleure sélection"]].forEach(([v, l], i) =>
    stat(s, MX + i * 4.1, 5.15, i === 2 ? 4.0 : 3.6, v, l, `h-stat-${i + 1}`, i === 2 ? C.accent2 : C.accent1, i === 2 ? 30 : 40));
  s.addNotes("QAOA ne sait pas gérer de contrainte : on la déplace dans l'objectif sous forme de pénalité A par paire en conflit. Avec A plus grand que le plus gros profit, retirer une commande en conflit fait toujours baisser f, donc le minimum est une sélection valide. Qiskit applique ensuite la substitution x = (1 − Z)/2 avec to_ising(), ce qui donne un Hamiltonien d'Ising de 8 qubits et 18 termes. Comme la substitution est exacte, l'énergie d'un état de base est le coût f(x) : trouver l'état fondamental, c'est trouver la meilleure sélection.");

  // ---------------------------------------------------- 6. circuit
  s = contentSlide(2, "Le circuit QAOA alterne coût et mélange sur p couches", "QAOA");
  const ch = CW * 698 / 2804;
  s.addImage({ path: IMG("circuit.png"), x: MX, y: 1.85, w: CW, h: ch, altText: "Une couche QAOA sur 8 qubits : préparation, rotations RZ, une porte ZZ par conflit, rotations RX" });
  s.addText("Une couche (p = 1) : préparation |+⟩, R_Z (champs h_i), une porte ZZ par conflit (J_ij), puis le mélangeur R_X",
    { x: MX, y: 1.9 + ch, w: CW, h: 0.35, fontSize: 12, color: MUTED, margin: 0, align: "center", isTextBox: true, objectName: "circuit-caption" });
  card(s, MX, 5.35, 6.6, 1.35, "state-card", C.text2);
  s.addText(formula("|ψ(γ, β)⟩ = U_{B}(β_{p}) U_{C}(γ_{p}) … U_{B}(β_{1}) U_{C}(γ_{1}) |+⟩^{⊗n}", { color: C.background1, fontSize: 19 }).concat(
    [{ text: "", options: { breakLine: true } }], formula("U_{C}(γ) = e^{−iγH}  marque les bonnes sélections,  U_{B}(β) = e^{−iβΣX}  les mélange", { color: "CFE3E0", fontSize: 14 })),
    { x: MX + 0.3, y: 5.35, w: 6.0, h: 1.35, margin: 0, valign: "middle", fontFace: THEME.headFontFace, isTextBox: true, objectName: "state-formula" });
  [["3", "couches (p)"], ["6", "paramètres (γ, β)"], ["46", "de profondeur"], ["60", "portes à 2 qubits"]].forEach(([v, l], i) =>
    stat(s, 7.55 + i * 1.32, 5.2, 1.3, v, l, `circ-stat-${i + 1}`, C.accent1, 28));
  s.addNotes("QAOAAnsatz construit le circuit à partir de H. On part de la superposition uniforme : toutes les sélections ont la même amplitude. U_C ajoute à chaque sélection une phase proportionnelle à son coût, U_B mélange les sélections en inversant des bits. On répète p fois. Avec p = 3 : 6 paramètres, profondeur 46 et 60 portes à 2 qubits, soit 2 portes CX par conflit et par couche.");

  // ---------------------------------------------------- 7. boucle hybride
  s = contentSlide(2, "Une boucle hybride : le quantique mesure, le classique ajuste", "QAOA");
  const loop = [
    ["Préparer", "le circuit |ψ(γ, β)⟩ avec les angles courants"],
    ["Estimer", "l'énergie ⟨H⟩ (StatevectorEstimator, sans bruit)"],
    ["Ajuster", "COBYLA met à jour (γ, β) pour faire baisser ⟨H⟩"],
    ["Échantillonner", "4096 tirages, garder la meilleure sélection valide"],
  ];
  const lw = (CW - 3 * 0.45) / 4, ly = 2.0;
  loop.forEach(([t, d], i) => {
    const x = MX + i * (lw + 0.45);
    card(s, x, ly, lw, 2.2, `loop-card-${i + 1}`, i === 3 ? "FBE3DC" : HEX.lt2);
    s.addText(String(i + 1), { shape: pres.shapes.OVAL, x: x + 0.2, y: ly + 0.2, w: 0.5, h: 0.5, fill: { color: i === 3 ? C.accent2 : C.accent1 },
      color: C.background1, fontSize: 16, bold: true, align: "center", valign: "middle", margin: 0, objectName: `loop-num-${i + 1}` });
    s.addText(t, { x: x + 0.2, y: ly + 0.85, w: lw - 0.4, h: 0.45, fontSize: 18, bold: true, color: C.text2, margin: 0, isTextBox: true, objectName: `loop-title-${i + 1}` });
    s.addText(d, { x: x + 0.2, y: ly + 1.3, w: lw - 0.4, h: 0.8, fontSize: 14, color: MUTED, margin: 0, valign: "top", isTextBox: true, objectName: `loop-desc-${i + 1}` });
    if (i < 3) s.addImage({ data: icons.arrow, x: x + lw + 0.13, y: ly + 1.0, w: 0.19, h: 0.19, altText: "" });
  });
  s.addImage({ data: icons.redo, x: MX + 0.2, y: 4.45, w: 0.3, h: 0.3, altText: "Boucle" });
  s.addText("Étapes 1 à 3 répétées jusqu'à convergence (300 itérations max), depuis 3 points de départ aléatoires : le paysage d'énergie a des minima locaux.",
    { x: MX + 0.65, y: 4.35, w: CW - 0.65, h: 0.5, fontSize: 15, color: INK, margin: 0, valign: "middle", isTextBox: true, objectName: "loop-note" });
  [["871", "évaluations de l'énergie par exécution"], ["5,0 s", "d'optimisation COBYLA (8 qubits)"], ["−8,09", "énergie minimale ⟨H⟩ + C"], ["65,7 %", "des tirages finaux sont valides"]].forEach(([v, l], i) =>
    stat(s, MX + i * (lw + 0.45), 5.1, lw, v, l, `loop-stat-${i + 1}`, i === 3 ? C.accent2 : C.accent1, 34));
  s.addNotes("Le circuit prépare un état, on estime son énergie moyenne, et l'optimiseur classique COBYLA ajuste les angles pour la faire baisser. COBYLA ne voit que la valeur de l'énergie, pas le circuit. On relance depuis trois points de départ, car il y a des minima locaux. À la fin, on mesure le circuit 4096 fois, on écarte les sélections en conflit, car la pénalité ne garantit la validité qu'à l'optimum, et on garde la meilleure.");

  // ---------------------------------------------------- 8. résultats QAOA émulé
  pres.addSection({ title: "Résultats" });
  s = contentSlide(3, "QAOA émulé retrouve l'optimum, sans le garantir", "Résultats");
  [["17 / 34", "même profit optimal que la recherche exhaustive (8 puis 16 commandes)"], ["8/10  •  3/3", "exécutions qui trouvent l'optimum (base, puis x2)"], ["4,9 s / 113 s", "durée moyenne par exécution (base, puis x2)"]].forEach(([v, l], i) =>
    stat(s, MX, 1.95 + i * 1.6, 4.6, v, l, `qaoa-stat-${i + 1}`, i === 1 ? C.accent2 : C.accent1, 36));
  s.addChart(pres.charts.BAR, [
    { name: "Hasard (part valide de l'espace)", labels: ["Atelier de base (8)", "Atelier x2 (16)"], values: [16.0, 2.6] },
    { name: "QAOA émulé (tirages valides)", labels: ["Atelier de base (8)", "Atelier x2 (16)"], values: [62.9, 47.0] },
  ], {
    ...chartBase(), x: 5.7, y: 1.9, w: 7.03, h: 3.9, barDir: "col", barGrouping: "clustered", barGapWidthPct: 60,
    chartColors: [HEX.accent5, HEX.accent1], showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: '0.0" %"',
    valAxisMinVal: 0, valAxisMaxVal: 80, valAxisMajorUnit: 20, valAxisLabelFormatCode: '0" %"',
    showLegend: true, legendPos: "b", showTitle: true, title: "Part des sélections valides : QAOA concentre la mesure", objectName: "valid-chart",
  });
  card(s, 5.7, 5.95, 7.03, 0.8, "signif-card", "FBE3DC");
  s.addText("À 16 commandes, 4096 tirages au hasard n'ont que 6,1 % de chances de contenir l'optimum. QAOA l'a trouvé 5 fois sur 5.",
    { x: 5.9, y: 5.95, w: 6.63, h: 0.8, fontSize: 14, color: INK, margin: 0, valign: "middle", isTextBox: true, objectName: "signif-text" });
  s.addNotes("Résultats du benchmark : 10 exécutions sur l'atelier de base, 3 sur l'atelier x2, chacune avec une graine différente. QAOA retrouve le profit optimal, mais l'a manqué 2 fois sur 10 sur l'atelier de base, probablement à cause d'un minimum local de COBYLA. Le graphe montre que QAOA concentre la mesure sur les sélections valides : ×3,9 par rapport au hasard sur 8 commandes, ×18,3 sur 16. Sur 8 commandes, trouver l'optimum ne prouve rien, car 4096 tirages couvrent largement 256 sélections ; sur 16, c'est significatif.");

  // ---------------------------------------------------- 9. émulateur vs IBM
  s = contentSlide(3, "Sur IBM ibm_fez, le bruit coûte des tirages valides, pas l'optimum", "Résultats");
  const cmpRows = [
    [th("Indicateur"), th("Émulé x1", { align: "center" }), th("IBM x1", { align: "center" }), th("Émulé x2", { align: "center" }), th("IBM x2", { align: "center" })],
    ["Profondeur du circuit", "47", "278", "47", "250"],
    ["Portes à 2 qubits", "60", "108", "120", "222"],
    ["Fidélité estimée", "100 %", "69,8 %", "100 %", "46,1 %"],
    ["Tirages valides", "62,9 %", "50,6 %", "47,0 %", "26,6 %"],
    ["Profit trouvé", "17", "17", "34", "34"],
  ].map((r, i) => i === 0 ? r : r.map((v, k) => td(v, { align: k ? "center" : "left", bold: k === 0 || k % 2 === 0, color: k === 0 ? C.text2 : (k % 2 === 0 ? C.accent2 : INK),
    fill: { color: i % 2 ? "FFFFFF" : HEX.lt2 } })));
  s.addTable(cmpRows, tableOpts({ x: MX, y: 1.95, w: 7.4, colW: [2.6, 1.2, 1.2, 1.2, 1.2], fontSize: 14, rowH: 0.45, objectName: "ibm-table" }));
  card(s, 8.4, 1.95, 4.33, 2.7, "machine-card");
  s.addText([{ text: "ibm_fez", options: { bold: true, fontSize: 20, color: C.text2, breakLine: true } },
    { text: "IBM Heron r2, 156 qubits", options: { fontSize: 14, color: MUTED, breakLine: true, paraSpaceAfter: 8 } },
    ...bullets([[B("T1 / T2 : "), T("120 / 90 µs")], [B("Erreur cz : "), T("0,26 %")], [B("Erreur de lecture : "), T("1,03 %")],
      [B("Angles : "), T("ceux de l'émulateur, 4096 tirages")]], 14)],
    { x: 8.65, y: 2.05, w: 3.85, h: 2.5, margin: 0, valign: "top", color: INK, isTextBox: true, objectName: "machine-text" });
  s.addText([B("La transpilation multiplie la profondeur par 5,9 "), T("(portes natives cz, rz, sx et topologie de la puce). La fidélité suit le nombre de portes à 2 qubits : ×2,06 portes → ×0,66 fidélité, et la part valide passe de 80 % à 57 % de celle de l'émulateur.")],
    { x: MX, y: 4.8, w: 7.4, h: 1.1, fontSize: 14, color: INK, margin: 0, valign: "top", isTextBox: true, objectName: "ibm-takeaway" });
  card(s, 8.4, 4.8, 4.33, 1.9, "incident-card", "FBE3DC");
  s.addImage({ data: icons.warn, x: 8.6, y: 4.97, w: 0.3, h: 0.3, altText: "Incident" });
  s.addText([{ text: "Incident : job x2 bloqué en RUNNING", options: { bold: true, color: C.accent2, breakLine: true } },
    { text: "Sur 3 machines, sans jamais s'exécuter. Correctif : un PUB par groupe de 8 qubits, puis on recolle les tirages (terminé en 7,5 s).", options: { color: INK } }],
    { x: 9.0, y: 4.9, w: 3.55, h: 1.7, fontSize: 13, margin: 0, valign: "top", isTextBox: true, objectName: "incident-text" });
  s.addNotes("Pour juger le vrai matériel, on reprend les angles trouvés par l'émulateur et on ne fait qu'échantillonner sur ibm_fez. La transpilation vers les portes natives et la topologie fait passer la profondeur de 47 à 278. Le bruit fait baisser la part de tirages valides : 50,6 % au lieu de 62,9 % sur 8 qubits, 26,6 % au lieu de 47 % sur 16. Mais la machine retrouve quand même le profit optimal dans les deux cas. Anecdote utile : le job x2 envoyé d'un bloc restait bloqué en RUNNING côté IBM sur trois machines ; le découper en un PUB par groupe de qubits a réglé le problème.");

  // ---------------------------------------------------- 10. temps
  s = contentSlide(3, "À cette taille, la recherche exhaustive gagne sans discussion", "Résultats");
  s.addChart(pres.charts.BAR, [
    { name: "Exhaustive", labels: ["Atelier de base (8)", "Atelier x2 (16)"], values: [0.123, 27.9] },
    { name: "QAOA émulé", labels: ["Atelier de base (8)", "Atelier x2 (16)"], values: [4881, 112591] },
    { name: "QAOA IBM (échantillonnage)", labels: ["Atelier de base (8)", "Atelier x2 (16)"], values: [4512, 5982] },
    { name: "QAOA IBM (de bout en bout)", labels: ["Atelier de base (8)", "Atelier x2 (16)"], values: [16200, 36500] },
  ], {
    ...chartBase(), x: MX, y: 1.9, w: 7.6, h: 4.85, barDir: "col", barGrouping: "clustered", barGapWidthPct: 50,
    chartColors: [HEX.accent1, HEX.accent6, HEX.dk2, HEX.accent5], showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "#,##0.###",
    valAxisLogScaleBase: 10, valAxisMinVal: 0.001, valAxisMaxVal: 1000000, valAxisLabelFormatCode: "#,##0.##", valAxisTitle: "ms (échelle log)", showValAxisTitle: true,
    showLegend: true, legendPos: "b", showTitle: true, title: "Durée par exécution (ms)", objectName: "time-chart",
  });
  s.addText(bullets([
    [B("Exhaustive : "), T("0,12 ms et 28 ms, de 215 à 132 000 fois plus rapide que QAOA.")],
    [B("Échantillonnage : "), T("IBM rejoint l'émulateur à 8 qubits (×1,08) et le dépasse à 16 (×18,8).")],
    [B("De bout en bout : "), T("l'émulateur gagne à 8 commandes (4,9 s contre 16,2 s), IBM à 16 (36,5 s contre 113 s).")],
    [B("Surcoût hors calcul : "), T("~5 s de transpilation et d'envoi, 1 à 2 s de file d'attente.")],
  ], 15), { x: 8.55, y: 1.95, w: 4.18, h: 4.8, margin: 0, valign: "top", color: INK, isTextBox: true, objectName: "time-bullets" });
  s.addNotes("Échelle logarithmique. La recherche exhaustive est de loin la plus rapide : 65 536 candidats se testent en 28 ms. Pour QAOA, il faut distinguer deux durées sur IBM : l'échantillonnage seul, et le bout en bout qui ajoute l'optimisation locale des angles, la transpilation, l'envoi et la file d'attente. De bout en bout, l'émulateur est plus rapide à 8 qubits, car l'aller-retour vers IBM coûte plus que simuler 8 qubits ; à 16 qubits, IBM devient 3,1 fois plus rapide.");

  // ---------------------------------------------------- 11. taille
  pres.addSection({ title: "Analyse" });
  s = contentSlide(4, "Taille du problème : chaque méthode bute sur une limite différente", "Analyse");
  const growRows = [
    [th("Base → x2"), th("Exhaustive", { align: "center" }), th("QAOA émulé", { align: "center" }), th("QAOA IBM", { align: "center" })],
    ["Durée", "×227", "×23", "×1,33 (échantillonnage)"],
    ["Temps CPU local", "×226", "×276", "n/a (QPU)"],
    ["Ce qui grandit", "candidats : ×256 (2⁸)", "évaluations : ×1,03", "portes 2 qubits : ×2,06"],
  ].map((r, i) => i === 0 ? r : r.map((v, k) => td(v, { align: k ? "center" : "left", bold: k === 0, color: k === 0 ? C.text2 : INK, fill: { color: i % 2 ? "FFFFFF" : HEX.lt2 } })));
  s.addTable(growRows, tableOpts({ x: MX, y: 1.95, w: CW, colW: [2.8, 3.11, 3.11, 3.11], fontSize: 15, rowH: 0.48, objectName: "growth-table" }));
  const limits = [
    ["clock", "Exhaustive : le temps", "La durée suit 2ⁿ. À 0,43 µs par candidat : ~30 min pour 32 commandes, ~5 jours pour 40, ~15 ans pour 50."],
    ["memory", "QAOA émulé : la mémoire", "Un vecteur d'état coûte 16 × 2ⁿ octets. Sur nos 14,6 Go, l'émulation s'arrête vers 29 commandes."],
    ["noise", "QAOA IBM : le bruit", "~0,33 % de fidélité perdue par porte à 2 qubits : ~23 % à 32 commandes, ~5 % à 64, proche du hasard."],
  ];
  const limW = (CW - 2 * 0.35) / 3;
  limits.forEach(([ic, t, d], i) => {
    const x = MX + i * (limW + 0.35);
    card(s, x, 4.2, limW, 2.5, `limit-${i + 1}`, C.text2);
    s.addText("", { shape: pres.shapes.OVAL, x: x + 0.25, y: 4.4, w: 0.55, h: 0.55, fill: { color: i === 2 ? C.accent2 : C.accent1 }, line: { type: "none" }, objectName: `limit-icon-bg-${i + 1}` });
    s.addImage({ data: icons[ic], x: x + 0.37, y: 4.52, w: 0.31, h: 0.31, altText: t });
    s.addText(t, { x: x + 0.95, y: 4.4, w: limW - 1.15, h: 0.55, fontSize: 16, bold: true, color: C.background1, margin: 0, valign: "middle", isTextBox: true, objectName: `limit-title-${i + 1}` });
    s.addText(d, { x: x + 0.25, y: 5.1, w: limW - 0.5, h: 1.5, fontSize: 14, color: "D7ECE9", margin: 0, valign: "top", isTextBox: true, objectName: `limit-desc-${i + 1}` });
  });
  s.addNotes("Passer de 8 à 16 commandes. La recherche exhaustive suit 2^n : ×227 en durée. QAOA émulé n'a presque pas besoin de plus d'évaluations (871 puis 894), mais chaque évaluation simule 2^n amplitudes, donc son temps CPU explose aussi (×276), masqué par le parallélisme. Sur IBM, l'échantillonnage n'augmente que de ×1,33, mais le nombre de portes double et la fidélité baisse. Chaque méthode a donc sa limite : le temps, la mémoire, ou le bruit. Les chiffres à 32, 40, 50 ou 64 commandes sont des extrapolations du notebook, pas des mesures.");

  // ---------------------------------------------------- 12. couches
  s = contentSlide(4, "Nombre de couches p : plus d'expressivité, plus de bruit", "Analyse");
  s.addText(bullets([
    [B("Paramètres : "), T("2p angles à optimiser (6 pour p = 3). Plus de dimensions pour COBYLA, plus de minima locaux : d'où les 3 départs.")],
    [B("Circuit : "), T("2 portes CX par conflit et par couche, soit 2 × |E| × p = 60 portes à p = 3. La profondeur croît linéairement avec p.")],
    [B("Qualité : "), T("un p petit ne garantit pas l'optimum (2 échecs sur 10 à p = 3) ; quand p augmente, QAOA tend vers l'évolution adiabatique et l'optimum exact.")],
    [B("Sur matériel : "), T("chaque couche ajoute ~36 portes cz après transpilation, donc ~11 % de fidélité en moins.")],
  ], 15), { x: MX, y: 1.95, w: 6.2, h: 4.0, margin: 0, valign: "top", color: INK, isTextBox: true, objectName: "layers-bullets" });
  card(s, MX, 6.0, 6.2, 0.75, "layers-note", "FBE3DC");
  s.addText("Le notebook mesure uniquement p = 3 : un balayage de p reste à faire.", { x: MX + 0.25, y: 6.0, w: 5.7, h: 0.75, fontSize: 14,
    color: INK, margin: 0, valign: "middle", isTextBox: true, objectName: "layers-note-text" });
  // Estimate: fidelity measured at p=3 (69.8 %, 108 cz) scaled per layer; only p=3 is a measurement.
  const fid = [1, 2, 3, 4, 5].map((p) => Math.round(1000 * Math.pow(0.698, p / 3)) / 10);
  s.addChart(pres.charts.BAR, [{ name: "Fidélité estimée", labels: ["p = 1", "p = 2", "p = 3 (mesuré)", "p = 4", "p = 5"], values: fid }], {
    ...chartBase(), x: 7.15, y: 1.9, w: 5.58, h: 4.85, barDir: "col", barGapWidthPct: 60, chartColors: [HEX.accent1],
    showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: '0.0" %"', valAxisMinVal: 0, valAxisMaxVal: 100, valAxisMajorUnit: 25,
    valAxisLabelFormatCode: '0" %"', showLegend: false, showTitle: true, title: "Fidélité estimée sur ibm_fez selon p (atelier de base)", objectName: "layers-chart",
  });
  s.addNotes("Le notebook fixe p = 3 sans faire varier p, donc cette diapositive raisonne à partir de ce qui est mesuré. Plus de couches donnent plus d'expressivité : en théorie, quand p tend vers l'infini, QAOA retrouve l'optimum exact. Mais chaque couche ajoute deux paramètres à optimiser et environ 36 portes cz sur ibm_fez. Le graphe extrapole la fidélité mesurée à p = 3 (69,8 % pour 108 portes cz) aux autres valeurs de p : seule la barre p = 3 est une mesure. Il y a donc un compromis : plus de couches améliorent la solution idéale, mais le bruit grandit. La prochaine étape est un balayage de p, sur émulateur puis sur IBM.");

  // ---------------------------------------------------- 13. bilan
  s = contentSlide(4, "Bilan : exacte et rapide à petite taille, QAOA est prometteur mais bruité", "Analyse");
  const bil = [
    [th(""), th("Recherche exhaustive"), th("QAOA émulé"), th("QAOA natif (IBM)")],
    ["Optimum", "Garanti", "Probable (85 % des exécutions)", "Trouvé 4/4, angles fournis par l'émulateur"],
    ["Vitesse à 8 / 16 commandes", "Meilleure : 0,12 ms / 28 ms", "4,9 s / 113 s", "16,2 s / 36,5 s de bout en bout"],
    ["Croissance mesurée", "2ⁿ (×227)", "2ⁿ en CPU (×276)", "Quasi linéaire (×1,33 en échantillonnage)"],
    ["Limite", "Temps : ~15 ans à 50 commandes", "Mémoire : ~29 commandes sur 14,6 Go", "Bruit : fidélité ~5 % vers 64 commandes"],
  ].map((r, i) => i === 0 ? r : r.map((v, k) => td(v, { bold: k === 0 || (i === 2 && k === 1) || (i === 1 && k === 1), color: k === 0 ? C.text2 : INK, fill: { color: i % 2 ? "FFFFFF" : HEX.lt2 } })));
  s.addTable(bil, tableOpts({ x: MX, y: 1.95, w: CW, colW: [2.6, 3.1, 3.1, 3.33], fontSize: 15, rowH: 0.62, objectName: "summary-table" }));
  s.addText([B("Pour cet atelier, la recherche exhaustive est le bon choix. "), T("QAOA émulé sert à concevoir le circuit et trouver les angles ; QAOA natif est la seule variante dont le coût n'explose pas avec n, mais son bruit grandit avec le nombre de portes.")],
    { x: MX, y: 5.35, w: CW, h: 1.3, fontSize: 16, color: INK, margin: 0, valign: "top", isTextBox: true, objectName: "summary-text" });
  s.addNotes("Le tableau résume la comparaison. Seule la recherche exhaustive garantit l'optimum, et c'est la plus rapide à 8 et 16 commandes. QAOA émulé ne peut pas apporter de gain de performance : simuler un ordinateur quantique coûte lui-même 2^n. QAOA natif est le seul dont le coût ne croît presque pas, mais il dépend de l'émulateur pour ses angles, et sa fidélité baisse avec la taille du circuit.");

  // ---------------------------------------------------- 14. conclusion
  pres.addSection({ title: "Conclusion" });
  s = pres.addSlide({ masterName: "CLOSING_DARK", sectionTitle: "Conclusion" });
  s.addText("Aucun avantage quantique ici, mais une chaîne complète et validée", { placeholder: "title" });
  const take = [
    ["Modéliser", "Un problème d'atelier réel devient un MWIS, puis un QUBO, puis un Hamiltonien d'Ising de 8 à 16 qubits."],
    ["Valider", "Recherche exhaustive, QAOA émulé et QAOA sur ibm_fez trouvent le même optimum : profit 17, puis 34."],
    ["Comparer", "La force brute gagne à cette taille ; QAOA concentre la mesure (×18 sur le hasard) et IBM passe à l'échelle en temps, pas en bruit."],
  ];
  const tw = (CW - 2 * 0.35) / 3;
  take.forEach(([t, d], i) => {
    const x = MX + i * (tw + 0.35);
    card(s, x, 1.75, tw, 2.5, `take-${i + 1}`, "2F5A68");
    s.addText(t, { x: x + 0.3, y: 1.95, w: tw - 0.6, h: 0.5, fontSize: 20, bold: true, color: C.accent3, margin: 0, isTextBox: true, objectName: `take-title-${i + 1}` });
    s.addText(d, { x: x + 0.3, y: 2.55, w: tw - 0.6, h: 1.6, fontSize: 15, color: C.background1, margin: 0, valign: "top", isTextBox: true, objectName: `take-desc-${i + 1}` });
  });
  s.addText("Pour aller plus loin", { x: MX, y: 4.55, w: 6, h: 0.45, fontSize: 18, bold: true, color: C.background1, margin: 0, isTextBox: true, objectName: "next-heading" });
  s.addText(bullets([
    "Un atelier de plusieurs dizaines de commandes, où la force brute n'est plus praticable",
    "Un balayage du nombre de couches p, et l'optimisation des angles directement sur le matériel",
    "Des machines moins bruitées ou la correction d'erreurs, et une comparaison avec des solveurs classiques spécialisés (PLNE, heuristiques MWIS)",
  ], 15), { x: MX, y: 5.05, w: CW, h: 1.75, margin: 0, valign: "top", color: "D7ECE9", isTextBox: true, objectName: "next-bullets" });
  s.addNotes("Conclusion. On a parcouru toute la chaîne, du problème d'atelier au circuit exécuté sur un vrai processeur quantique, et les trois méthodes sont d'accord sur l'optimum. À cette taille, aucun avantage quantique n'est démontré : la force brute est exacte, simple et beaucoup plus rapide. Pour que QAOA devienne intéressant, il faudrait un problème bien plus grand, du matériel moins bruité, et une comparaison honnête avec les meilleurs solveurs classiques. Merci, nous prenons vos questions.");

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
}

main().catch((e) => { console.error(e); process.exit(1); });
