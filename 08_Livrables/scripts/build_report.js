// Squelette Word du rapport : une section par rôle, chiffres cités avec leur source, annexes.
// Lit report-data.json (produit par build_dossiers.py) ; n'écrit aucun chiffre lui-même.
"use strict";
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType,
  AlignmentType, PageBreak, TableOfContents, Footer, PageNumber, ShadingType, LevelFormat, BorderStyle,
} = require("docx");

const DATA = JSON.parse(fs.readFileSync(process.argv[2] || path.join(__dirname, "report-data.json"), "utf8"));
const OUT = process.argv[3] || "/home/ubuntu/capstone-livrables/rapport-sujet-2.docx";
const TEXT_WIDTH = 9026; // A4, marges de 2,54 cm

const bodyFont = { font: "Arial", size: 22 };
const smallFont = { font: "Arial", size: 18 };

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
function cell(text, width, opts = {}) {
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    shading: opts.head ? { type: ShadingType.CLEAR, color: "auto", fill: "E7E6E6" } : undefined,
    margins: { top: 40, bottom: 40, left: 80, right: 80 },
    children: [new Paragraph({ spacing: { after: 0 }, alignment: opts.right ? AlignmentType.RIGHT : AlignmentType.LEFT,
      children: [new TextRun({ text: String(text == null ? "" : text), ...smallFont, bold: !!opts.head })] })],
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
const pageBreak = () => new Paragraph({ children: [new PageBreak()] });

const children = [];

// ---- page de titre
children.push(
  new Paragraph({ spacing: { before: 2400, after: 240 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: DATA.title, font: "Arial", size: 40, bold: true })] }),
  new Paragraph({ spacing: { after: 600 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: DATA.subtitle, font: "Arial", size: 26, color: "555555" })] }),
  p("Rapport écrit — remise le 25 novembre 2026", { para: { alignment: AlignmentType.CENTER } }),
  p("Groupe : [[prénoms et noms, un par rôle]]", { para: { alignment: AlignmentType.CENTER }, run: { italics: true } }),
  p(`Squelette généré le ${DATA.generated} depuis la dataroom du groupe (${DATA.n_figures} chiffres sourcés ; contrôles du classeur : ${DATA.checks}). ` +
    "Chaque chiffre cité renvoie à l'annexe A : fichier, page, URL, date de consultation. Les passages surlignés sont à rédiger par l'auteur de la section.",
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
children.push(placeholder("[À rédiger] La recommandation en trois phrases : ce que nous recommandons, ce que cela coûte et à qui, ce qui nous ferait changer d'avis."));
children.push(pageBreak());

// ---- les questions du cours
children.push(h("Les quatre questions de recherche de Veolia, et où le rapport y répond", HeadingLevel.HEADING_1));
children.push(...numbered(DATA.research_questions, "numbers"));
children.push(h("Le chantier « capacité » en cinq blocs", HeadingLevel.HEADING_2));
children.push(...numbered(DATA.capacity_blocks, "numbers2"));
children.push(pageBreak());

// ---- une section par rôle
for (const role of DATA.roles) {
  children.push(h(`Rôle ${role.n} — ${role.name}`, HeadingLevel.HEADING_1));
  children.push(p(role.brief, { run: { italics: true, color: "555555" } }));
  children.push(placeholder(`[Section signée par : ______ ] — ${role.tabs ? "dans le classeur : " + role.tabs : ""}`));
  children.push(h("Réponse", HeadingLevel.HEADING_2), p(role.answer));
  children.push(h("Raisonnement", HeadingLevel.HEADING_2), ...numbered(role.reasoning, `role${role.n}`));
  children.push(placeholder("[À rédiger] Le développement de la section : 6 à 9 pages, chaque chiffre avec son identifiant F… du registre."));
  children.push(h("Chiffres cités", HeadingLevel.HEADING_2));
  children.push(table(["ID", "Chiffre", "Période", "Valeur", "Source (fichier, page)", "Relu par"],
    role.figures.map((f) => [f[0], f[1], f[2], f[3], f[4], f[5] || "—"]), [900, 3000, 1000, 1300, 2126, 700], [3]));
  children.push(h("Ce que nous n'avons pas pu établir", HeadingLevel.HEADING_2), ...bullets(role.open));
  children.push(h("Questions probables du jury", HeadingLevel.HEADING_2));
  for (const [q, a] of role.qa) {
    children.push(runs([{ text: q, bold: true }]));
    children.push(p(a));
  }
  children.push(pageBreak());
}

// ---- la fourchette
children.push(h("La fourchette de capacité, pour les autres périmètres", HeadingLevel.HEADING_1));
children.push(p("Le nombre que les groupes « technologies de l'eau », « déchets dangereux » et « bioénergie » doivent respecter, dans ses trois scénarios et sous ses deux plafonds."));
children.push(table(["Fin 2027", "Défavorable", "Central", "Favorable"], DATA.range_rows, [4226, 1600, 1600, 1600], [1, 2, 3]));
children.push(h("Ce qu'il reste à faire avant la remise", HeadingLevel.HEADING_2), ...bullets(DATA.todo));
children.push(pageBreak());

// ---- annexe A : le registre
children.push(h("Annexe A — Registre des chiffres cités", HeadingLevel.HEADING_1));
children.push(p(`${DATA.register.length} chiffres, chacun avec le fichier joint, la page, l'URL d'origine et la date de consultation. « Relu par » vide : un seul membre a lu la page.`, { run: smallFont }));
children.push(table(["ID", "Chiffre", "Valeur", "Période", "Citation", "Rôle", "Relu"],
  DATA.register.map((r) => [r[0], r[1], r[2], r[3], r[4], r[5], r[6] || "—"]), [600, 2400, 1000, 900, 3326, 400, 400], [2]));
children.push(pageBreak());

// ---- annexe B : outils d'IA
children.push(h("Annexe B — Outils d'intelligence artificielle utilisés", HeadingLevel.HEADING_1));
children.push(...bullets([
  "Dataroom : un serveur MCP privé au groupe, qui indexe les documents publics versés dans le vault et ne sait que chercher, lire une page et citer (fichier, page, URL, date de consultation). Il ne calcule rien, ne résume rien et ne produit aucun chiffre : chaque chiffre du registre a été lu par un membre sur la page citée, puis relu par un autre.",
  "Claude (Anthropic), via Claude Code : a construit le serveur, le registre, les contrôles automatiques, le générateur du classeur Excel et celui de ce squelette. Le classeur ne contient que des formules ; ses entrées sont les lignes du registre.",
  "LibreOffice (recalcul sans interface) : recalcule le classeur après chaque génération et signale toute erreur de formule ; aucune valeur calculée n'est saisie à la main.",
  "Règle tenue : aucun nombre, aucune source et aucune citation n'ont été produits par un modèle de langage. Le modèle construit l'outil qui calcule ; il ne calcule pas.",
]));
children.push(placeholder("[À rédiger] Ce que le groupe a décidé sans l'outil : le choix des hypothèses, la recommandation, les cas où l'outil a été contredit après lecture. Puis un tableau : membre, outil, usage, période."));

const doc = new Document({
  creator: "Groupe sujet 2",
  title: DATA.title,
  styles: { default: { document: { run: bodyFont } } },
  numbering: {
    config: [
      { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 540, hanging: 270 } } } }] },
      ...["numbers", "numbers2", "role1", "role2", "role3", "role4", "role5", "role6"].map((ref) => ({
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
  console.log("écrit", OUT, Math.round(buf.length / 1024), "Ko");
});
