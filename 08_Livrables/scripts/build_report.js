// Le rapport Word : une section par rôle, la prose du groupe avec chaque jeton [F12] résolu
// contre le registre, les annexes, et (hors version finale) le contrôle des citations.
// Lit rapport-sujet-2.data.json (dataroom report) ou, à défaut, report-data.json
// (build_dossiers.py). N'écrit aucun chiffre lui-même : une valeur vient du registre ou
// le jeton est marqué inconnu.
"use strict";
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType,
  AlignmentType, PageBreak, TableOfContents, Footer, PageNumber, ShadingType, LevelFormat, BorderStyle,
  ExternalHyperlink,
} = require("docx");

const DATA = JSON.parse(fs.readFileSync(process.argv[2] || path.join(__dirname, "report-data.json"), "utf8"));
const OUT = process.argv[3] || "/home/ubuntu/capstone-livrables/rapport-sujet-2.docx";
const FINAL = process.argv.includes("--final") || DATA.final === true;
const TEXTS = DATA.texts || {};
const CONTROLS = DATA.controls || null;
const TEXT_WIDTH = 9026; // A4, marges de 2,54 cm

const bodyFont = { font: "Arial", size: 22 };
const smallFont = { font: "Arial", size: 18 };
const HEADINGS = [HeadingLevel.HEADING_1, HeadingLevel.HEADING_2, HeadingLevel.HEADING_3];
const numberingRefs = new Set(["numbers", "numbers2", "role1", "role2", "role3", "role4", "role5", "role6"]);
let dynamicLists = 0;

function p(text, opts = {}) {
  return new Paragraph({
    spacing: { after: 120, line: 300 },
    ...opts.para,
    children: [new TextRun({ text, ...bodyFont, ...opts.run })],
  });
}
function runs(parts, para = {}) {
  return new Paragraph({ spacing: { after: 120, line: 300 }, ...para, children: parts.map((x) => new TextRun({ ...bodyFont, ...x })) });
}
function h(text, level) {
  return new Paragraph({ heading: level, spacing: { before: 240, after: 120 }, children: [new TextRun({ text, font: "Arial" })] });
}
function bullets(items) {
  return items.map((t) => new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 80 },
    children: [new TextRun({ text: t, ...bodyFont })] }));
}
function numbered(items, ref) {
  numberingRefs.add(ref);
  return items.map((t) => new Paragraph({ numbering: { reference: ref, level: 0 }, spacing: { after: 80 },
    children: [new TextRun({ text: t, ...bodyFont })] }));
}
function placeholder(text) {
  return new Paragraph({
    spacing: { before: 160, after: 160 },
    shading: { type: ShadingType.CLEAR, color: "auto", fill: "FFF2A8" },
    children: [new TextRun({ text, ...bodyFont, italics: true })],
  });
}
function cell(content, width, opts = {}) {
  const children = Array.isArray(content)
    ? inlineRuns(content, smallFont)
    : [new TextRun({ text: String(content == null ? "" : content), ...smallFont, bold: !!opts.head })];
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    shading: opts.head ? { type: ShadingType.CLEAR, color: "auto", fill: "E7E6E6" } : undefined,
    margins: { top: 40, bottom: 40, left: 80, right: 80 },
    children: [new Paragraph({ spacing: { after: 0 }, alignment: opts.right ? AlignmentType.RIGHT : AlignmentType.LEFT, children })],
  });
}
function table(headers, rows, widths, rightCols = []) {
  const total = widths.reduce((a, b) => a + b, 0);
  if (total !== TEXT_WIDTH) throw new Error("largeurs de colonnes : " + total + " ≠ " + TEXT_WIDTH);
  return new Table({
    width: { size: TEXT_WIDTH, type: WidthType.DXA },
    columnWidths: widths,
    rows: [
      new TableRow({ tableHeader: true, children: headers.map((t, i) => cell(t, widths[i], { head: true })) }),
      ...rows.map((r) => new TableRow({ children: r.map((t, i) => cell(t, widths[i], { right: rightCols.includes(i) })) })),
    ],
  });
}
function evenWidths(n) {
  const base = Math.floor(TEXT_WIDTH / n);
  const widths = Array(n).fill(base);
  widths[n - 1] += TEXT_WIDTH - base * n;
  return widths;
}
const pageBreak = () => new Paragraph({ children: [new PageBreak()] });

// ---- la prose du groupe : runs et blocs
// Un jeton résolu s'imprime « valeur unité (F12) » ; « † » après l'identifiant quand personne
// n'a relu la ligne ; un jeton inconnu reste visible, en rouge sur jaune.
function inlineRuns(inlines, base = bodyFont) {
  const out = [];
  for (const r of inlines || []) {
    const fmt = { ...base, bold: !!r.bold, italics: !!r.italic };
    if (r.kind === "text") out.push(new TextRun({ text: r.text, ...fmt }));
    else if (r.kind === "code") out.push(new TextRun({ text: r.text, ...fmt, font: "Consolas" }));
    else if (r.kind === "link") out.push(new ExternalHyperlink({ link: r.url, children: [new TextRun({ text: r.text, ...fmt, style: "Hyperlink" })] }));
    else if (r.kind === "wikilink") out.push(new TextRun({ text: r.text, ...fmt, italics: true }));
    else if (r.kind === "token") {
      if (r.status !== "ok") out.push(new TextRun({ text: `[${r.id} ?]`, ...fmt, bold: true, color: "C00000", highlight: "yellow" }));
      else {
        out.push(new TextRun({ text: `${r.value}${r.unit ? " " + r.unit : ""}`, ...fmt }));
        out.push(new TextRun({ text: ` (${r.id}${r.checked ? "" : "†"})`, ...smallFont, color: "777777", bold: false, italics: !!r.italic }));
      }
    }
  }
  return out;
}
function renderBlocks(blocks, opts = {}) {
  const baseLevel = opts.baseLevel || 1;
  const out = [];
  for (const b of blocks || []) {
    if (b.kind === "heading") {
      const level = Math.min(3, Math.max(baseLevel + 1, baseLevel + b.level - 1));
      out.push(new Paragraph({ heading: HEADINGS[level - 1], spacing: { before: 240, after: 120 }, children: inlineRuns(b.inlines, { font: "Arial" }) }));
    } else if (b.kind === "paragraph") {
      out.push(new Paragraph({ spacing: { after: 120, line: 300 }, children: inlineRuns(b.inlines) }));
    } else if (b.kind === "quote") {
      out.push(new Paragraph({ spacing: { after: 120, line: 300 }, indent: { left: 540 },
        border: { left: { style: BorderStyle.SINGLE, size: 6, color: "BBBBBB", space: 8 } },
        children: inlineRuns(b.inlines, { ...bodyFont, italics: true }) }));
    } else if (b.kind === "list") {
      let ref = "bullets";
      if (b.ordered) { ref = `list${dynamicLists++}`; numberingRefs.add(ref); }
      for (const item of b.items) {
        out.push(new Paragraph({ numbering: { reference: ref, level: 0 }, spacing: { after: 80 }, children: inlineRuns(item) }));
      }
    } else if (b.kind === "todo") {
      if (!FINAL) out.push(placeholder(b.text));
    } else if (b.kind === "table" && b.rows.length) {
      const n = Math.max(...b.rows.map((r) => r.length));
      const widths = evenWidths(n);
      const pad = (row) => { const r = row.slice(); while (r.length < n) r.push([]); return r; };
      out.push(new Table({
        width: { size: TEXT_WIDTH, type: WidthType.DXA }, columnWidths: widths,
        rows: b.rows.map((row, i) => new TableRow({ tableHeader: i === 0,
          children: pad(row).map((c, j) => cell(c, widths[j], { head: i === 0 })) })),
      }));
      out.push(new Paragraph({ spacing: { after: 120 }, children: [] }));
    }
  }
  return out;
}
function prose(key) {
  const t = TEXTS[key];
  return t && t.has_prose ? t : null;
}

const children = [];

// ---- page de titre
children.push(
  new Paragraph({ spacing: { before: 2400, after: 240 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: DATA.title, font: "Arial", size: 40, bold: true })] }),
  new Paragraph({ spacing: { after: 600 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: DATA.subtitle, font: "Arial", size: 26, color: "555555" })] }),
  p("Rapport écrit — remise le 25 novembre 2026", { para: { alignment: AlignmentType.CENTER } }),
  p("Groupe : [[prénoms et noms, un par rôle]]", { para: { alignment: AlignmentType.CENTER }, run: { italics: true } }),
  p(`${FINAL ? "Rapport" : "Squelette"} généré le ${DATA.generated} depuis la dataroom du groupe (${DATA.n_figures} chiffres sourcés ; contrôles du classeur : ${DATA.checks}). ` +
    "Chaque chiffre cité renvoie à l'annexe A : fichier, page, URL, date de consultation." +
    (FINAL ? "" : " Les passages surlignés sont à rédiger par l'auteur de la section ; « † » après un identifiant : ligne que personne n'a encore relue."),
    { para: { alignment: AlignmentType.CENTER, spacing: { before: 1200 } }, run: { ...smallFont, color: "555555" } }),
  pageBreak(),
  h("Sommaire", HeadingLevel.HEADING_1),
  new TableOfContents("Sommaire", { hyperlink: true, headingStyleRange: "1-2" }),
  pageBreak(),
);

// ---- réponse en une page
children.push(h("La réponse en une page", HeadingLevel.HEADING_1), p(DATA.thread));
children.push(table(["Indicateur", "Valeur", "Ce qu'il faut en lire"],
  DATA.tiles.map((t) => [t[0], t[1], t[2]]), [3200, 1600, 4226], [1]));
const reco = prose("recommandation");
if (reco) children.push(h("La recommandation", HeadingLevel.HEADING_2), ...renderBlocks(reco.blocks, { baseLevel: 2 }));
else if (!FINAL) children.push(placeholder("[À rédiger] 08_Livrables/texte/recommandation.md : ce que nous recommandons, ce que cela coûte et à qui, ce qui nous ferait changer d'avis."));
children.push(pageBreak());

// ---- les questions du cours
children.push(h("Les quatre questions de recherche de Veolia, et où le rapport y répond", HeadingLevel.HEADING_1));
children.push(...numbered(DATA.research_questions, "numbers"));
children.push(h("Le chantier « capacité » en cinq blocs", HeadingLevel.HEADING_2));
children.push(...numbered(DATA.capacity_blocks, "numbers2"));
children.push(pageBreak());

// ---- une section par rôle
for (const role of DATA.roles) {
  const text = prose(`role-${role.n}`);
  const signed = TEXTS[`role-${role.n}`] && TEXTS[`role-${role.n}`].meta ? TEXTS[`role-${role.n}`].meta.signe_par : "";
  children.push(h(`Rôle ${role.n} — ${role.name}`, HeadingLevel.HEADING_1));
  children.push(p(role.brief, { run: { italics: true, color: "555555" } }));
  if (signed) children.push(p(`Section signée par : ${signed}`, { run: { italics: true, color: "555555" } }));
  else if (!FINAL) children.push(placeholder(`[Section signée par : ______ ] — ${role.tabs ? "dans le classeur : " + role.tabs : ""}`));
  if (text) {
    children.push(...renderBlocks(text.blocks, { baseLevel: 1 }));
  } else if (!FINAL) {
    children.push(h("Réponse (générée depuis le classeur, à réécrire)", HeadingLevel.HEADING_2), p(role.answer));
    children.push(h("Raisonnement (généré)", HeadingLevel.HEADING_2), ...numbered(role.reasoning, `role${role.n}`));
    children.push(placeholder(`[À rédiger] 08_Livrables/texte/role-${role.n}.md : 6 à 9 pages, chaque chiffre cité par son jeton [F…] du registre.`));
  }
  children.push(h("Chiffres cités", HeadingLevel.HEADING_2));
  children.push(table(["ID", "Chiffre", "Période", "Valeur", "Source (fichier, page)", "Relu par"],
    role.figures.map((f) => [f[0], f[1], f[2], f[3], f[4], f[5] || "—"]), [900, 3000, 1000, 1300, 2126, 700], [3]));
  if (!FINAL) {
    children.push(h("Ce que nous n'avons pas pu établir", HeadingLevel.HEADING_2), ...bullets(role.open));
    children.push(h("Questions probables du jury", HeadingLevel.HEADING_2));
    for (const [q, a] of role.qa) {
      children.push(runs([{ text: q, bold: true }]));
      children.push(p(a));
    }
  }
  children.push(pageBreak());
}

// ---- la fourchette
children.push(h("La fourchette de capacité, pour les autres périmètres", HeadingLevel.HEADING_1));
children.push(p("Le nombre que les groupes « technologies de l'eau », « déchets dangereux » et « bioénergie » doivent respecter, dans ses trois scénarios et sous ses deux plafonds."));
children.push(table(["Fin 2027", "Défavorable", "Central", "Favorable"], DATA.range_rows, [4226, 1600, 1600, 1600], [1, 2, 3]));
if (!FINAL) children.push(h("Ce qu'il reste à faire avant la remise", HeadingLevel.HEADING_2), ...bullets(DATA.todo));
children.push(pageBreak());

// ---- annexe A : le registre
children.push(h("Annexe A — Registre des chiffres cités", HeadingLevel.HEADING_1));
children.push(p(`${DATA.register.length} chiffres, chacun avec le fichier joint, la page, l'URL d'origine et la date de consultation. « Relu par » vide : un seul membre a lu la page.`, { run: smallFont }));
children.push(table(["ID", "Chiffre", "Valeur", "Période", "Citation", "Rôle", "Relu"],
  DATA.register.map((r) => [r[0], r[1], r[2], r[3], r[4], r[5], r[6] || "—"]), [600, 2400, 1000, 900, 3326, 400, 400], [2]));
if (DATA.cited_deals && DATA.cited_deals.length) {
  children.push(h("Opérations citées (deals.csv)", HeadingLevel.HEADING_2));
  children.push(table(["ID", "Opération", "Montant", "Devise", "Source", "Relecture"],
    DATA.cited_deals.map((d) => [d[0], d[1], d[2], d[3], d[4], d[5]]), [600, 2600, 1000, 700, 3126, 1000], [2]));
}
if (DATA.cited_targets && DATA.cited_targets.length) {
  children.push(h("Cibles citées (targets.csv)", HeadingLevel.HEADING_2));
  children.push(table(["ID", "Cible", "Valeur estimée", "Devise", "Source", "Relecture"],
    DATA.cited_targets.map((t) => [t[0], t[1], t[2], t[3], t[4], t[5]]), [600, 2600, 1000, 700, 3126, 1000], [2]));
}
children.push(pageBreak());

// ---- annexe B : outils d'IA
children.push(h("Annexe B — Outils d'intelligence artificielle utilisés", HeadingLevel.HEADING_1));
children.push(...bullets([
  "Dataroom : un serveur MCP privé au groupe, qui indexe les documents publics versés dans le vault et ne sait que chercher, lire une page et citer (fichier, page, URL, date de consultation). Il ne calcule rien, ne résume rien et ne produit aucun chiffre : chaque chiffre du registre a été lu par un membre sur la page citée, puis relu par un autre.",
  "Claude (Anthropic), via Claude Code : a construit le serveur, le registre, les contrôles automatiques, le générateur du classeur Excel et celui de ce rapport. Le classeur ne contient que des formules ; ses entrées sont les lignes du registre.",
  "LibreOffice (recalcul sans interface) : recalcule le classeur après chaque génération et signale toute erreur de formule ; aucune valeur calculée n'est saisie à la main.",
  "Règle tenue : aucun nombre, aucune source et aucune citation n'ont été produits par un modèle de langage. Le modèle construit l'outil qui calcule ; il ne calcule pas.",
]));
const ia = prose("outils-ia");
if (ia) children.push(...renderBlocks(ia.blocks, { baseLevel: 1 }));
else if (!FINAL) children.push(placeholder("[À rédiger] 08_Livrables/texte/outils-ia.md : ce que le groupe a décidé sans l'outil, les cas où l'outil a été contredit après lecture, puis un tableau membre / outil / usage / période."));

// ---- contrôle des citations (jamais dans la version finale)
if (CONTROLS && !FINAL) {
  children.push(pageBreak(), h("Contrôle des citations (retiré de la version finale)", HeadingLevel.HEADING_1));
  const files = Object.keys(CONTROLS.words || {});
  if (files.length) {
    const [lo, hi] = CONTROLS.pages_per_role || [6, 9];
    children.push(table(["Section", "Pages (PDF)", "Mots", "Cible", "À rédiger"],
      files.map((f) => [f, CONTROLS.pages && CONTROLS.pages[f] != null ? CONTROLS.pages[f] : "—", CONTROLS.words[f],
        f.startsWith("role-") ? `${lo}-${hi}` : "—", CONTROLS.todos[f] || 0]), [3026, 1500, 1500, 1500, 1500], [1, 2, 4]));
    if (CONTROLS.pages_total) children.push(p(`Total : ${CONTROLS.pages_total} pages (cible ${(CONTROLS.pages_total_target || [50, 60]).join("-")}).`, { run: smallFont }));
  }
  const list = (title, items, fmt) => {
    children.push(h(title, HeadingLevel.HEADING_2));
    children.push(...(items.length ? bullets(items.map(fmt)) : [p("aucun", { run: { italics: true } })]));
  };
  list(`Jetons inconnus (${CONTROLS.unknown.length})`, CONTROLS.unknown, (x) => `${x[0]} : [${x[1]}] ne correspond à aucune ligne des registres`);
  list(`Lignes citées que personne n'a relues (${CONTROLS.unread.length})`, CONTROLS.unread, (x) => `${x[0]} : ${x[1]} — ${x[2]}`);
  list(`Lignes citées contestées (${CONTROLS.disputed.length})`, CONTROLS.disputed, (x) => `${x[0]} : ${x[1]} — ${x[2]}`);
  list("Fichiers absents ou vides", [...(CONTROLS.missing || []), ...(CONTROLS.empty || [])], (x) => x);
  const uncited = Object.entries(CONTROLS.uncited || {});
  list("Chiffres du registre jamais cités par leur rôle", uncited, ([role, ids]) => `rôle ${role} : ${ids.length}${ids.length ? " — " + ids.slice(0, 20).join(", ") + (ids.length > 20 ? " …" : "") : ""}`);
}

const doc = new Document({
  creator: "Groupe sujet 2",
  title: DATA.title,
  styles: { default: { document: { run: bodyFont } } },
  numbering: {
    config: [
      { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 540, hanging: 270 } } } }] },
      ...[...numberingRefs].map((ref) => ({
        reference: ref, levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 540, hanging: 360 } } } }],
      })),
    ],
  },
  sections: [{
    properties: { page: { margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
      children: [new TextRun({ children: ["Page ", PageNumber.CURRENT], ...smallFont, color: "777777" })] })] }) },
    children,
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(OUT, buf);
  console.log("écrit", OUT, Math.round(buf.length / 1024), "Ko" + (FINAL ? " (version finale)" : ""));
});
