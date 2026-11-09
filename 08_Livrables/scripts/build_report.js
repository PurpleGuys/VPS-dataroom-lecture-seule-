// Le rapport Word du sujet 2, mis en page selon la « Méthodologie du rapport » de l'EDHEC
// (BBA3, 2025-2026), appliquée à la lettre :
//   page de couverture (logo, formation, entreprise, titre, prénoms et NOMS, année, responsable),
//   page blanche, remerciements, sommaire paginé qui est la page 1, introduction, parties numérotées
//   (1, 1.1, 1.1.1) aux titres « habillés », conclusion, bibliographie raisonnée, glossaire des termes
//   en italique dans le texte, table des annexes en dernière page numérotée, annexes paginées à part
//   en chiffres romains. Times New Roman 12, titres en gras 12 avec un seul retour avant et après,
//   interligne 1,5, marges de 2,5 cm, texte justifié ; les notes de bas de page remplacent les
//   parenthèses : chaque chiffre cité par son jeton [F12] porte sa source en note.
// Lit rapport-sujet-2.data.json (dataroom report) ou, à défaut, report-data.json (build_dossiers.py).
// N'écrit aucun chiffre lui-même : une valeur vient du registre ou le jeton est marqué inconnu.
// Écrit, à côté du .docx, la liste ordonnée des titres (…headings.json) : dataroom report la
// confronte au PDF pour numéroter le sommaire et la table des annexes au second passage.
"use strict";
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType,
  AlignmentType, PageBreak, Footer, PageNumber, ShadingType, LevelFormat, BorderStyle,
  ExternalHyperlink, FootnoteReferenceRun, ImageRun, NumberFormat, Tab, TabStopType, LeaderType,
} = require("docx");

const DATA = JSON.parse(fs.readFileSync(process.argv[2] || path.join(__dirname, "report-data.json"), "utf8"));
const OUT = process.argv[3] || "rapport-sujet-2.docx";
const HEADINGS_OUT = OUT.replace(/\.docx$/i, "") + ".headings.json";
const FINAL = process.argv.includes("--final") || DATA.final === true;
const TEXTS = DATA.texts || {};
const CONTROLS = DATA.controls || null;
const COVER = DATA.cover || {};
const PAGINATION = DATA.pagination || {};
const TOC_PAGES = new Map((PAGINATION.toc || []).map((h) => [h.text, h.page]));
const ANNEX_PAGES = new Map((PAGINATION.annexes || []).map((h) => [h.text, h.page]));
const GLOSSARY = DATA.glossary || [];
const NOTES = DATA.notes || {};   // la note de bas de page de chaque jeton, écrite par report.py
// L'édition web du même rapport (dataroom rapport-web) : chaque chiffre y ouvre sa page source, surlignée.
const WEB = DATA.web_edition || "https://dataroom.is42.fr/visualisation/rapport.html";

// ---- la méthode, en chiffres (twips : 1 cm = 567 ; tailles en demi-points : 12 pt = 24)
const FONT = "Times New Roman";
const BODY = 24;                       // « Times New Roman 12 »
const SMALL = 20;                      // 10 pt : tableaux, légendes
const NOTE = 18;                       // 9 pt : notes de bas de page
const LINE = 360;                      // « interligne 1,5 »
const MARGIN = 1417;                   // « 2,5 cm pour toutes les marges »
const PAGE_W = 11906, PAGE_H = 16838;  // A4
const TEXT_WIDTH = PAGE_W - 2 * MARGIN; // 9072
const RETURN = 360;                    // « un seul retour avant et après » un titre : une ligne à 1,5

const bodyFont = { font: FONT, size: BODY };
const smallFont = { font: FONT, size: SMALL };
const HEADING_LEVELS = [HeadingLevel.HEADING_1, HeadingLevel.HEADING_2, HeadingLevel.HEADING_3];
const numberingRefs = new Set(["numbers", "numbers2", "role1", "role2", "role3", "role4", "role5", "role6"]);
let dynamicLists = 0;

// ---- notes de bas de page : une par source et par partie, à sa première citation dans la partie
const footnotes = {};
let footnoteId = 0;
let noted = new Set();   // les groupes de sources déjà annotés dans la partie en cours
let groups = new Map();
let modelVersionShown = false;  // la version de la fourchette, dite dans la première note du modèle  // groupe -> { ids, unread, rest } : une note par page de document (ou par onglet du classeur) et par partie
const PARA_AFTER = 120;  // 6 pt entre deux paragraphes : pas de ligne vide, la méthode ne veut qu'un retour autour des titres
function refOf(id) { const R = DATA.refs || {}; return R[id] || R[String(id).split("-")[0]] || ""; }
function noteKey(r) {
  if (r.rowkind === "model") return "M|" + String(NOTES[r.id] || r.id).replace(/^\S+ — /, "");
  return "R|" + (r.file || r.id) + "#" + (r.page || "");
}
function collectGroups(blocks) {
  groups = new Map();
  const visit = (x) => {
    if (Array.isArray(x)) { x.forEach(visit); return; }
    if (!x || typeof x !== "object") return;
    if (x.kind === "token" && x.status === "ok") {
      const key = noteKey(x);
      const g = groups.get(key) || { ids: [], unread: false, model: x.rowkind === "model" };
      if (!g.ids.includes(x.id)) g.ids.push(x.id);
      if (!x.checked && x.rowkind !== "model") g.unread = true;
      groups.set(key, g);
    }
    for (const k of ["inlines", "items", "rows"]) if (x[k]) visit(x[k]);
  };
  visit(blocks);
}
function groupNote(key, r) {
  const g = groups.get(key) || { ids: [r.id], unread: !r.checked, model: r.rowkind === "model" };
  if (g.model) {
    // la version de la fourchette n'est dite qu'une fois : dans la première note du modèle
    let text = `${g.ids.join(", ")} — ${key.slice(2)}`;
    if (!modelVersionShown && DATA.model_version) { text = text.replace(/\.$/, "") + ` ; ${DATA.model_version}, valable pour toutes les notes du modèle.`; modelVersionShown = true; }
    return text;
  }
  const ref = refOf(g.ids[0]) || String(NOTES[g.ids[0]] || r.cite || r.file || "source au registre").replace(/^\S+ — /, "").replace(/ ; non relue\.$|\.$/, "");
  return `${g.ids.join(", ")} — ${ref}${g.unread ? (g.ids.length > 1 ? " ; non relues" : " ; non relue") : ""}.`;
}
function footnote(text) {
  footnotes[++footnoteId] = { children: [new Paragraph({ spacing: { after: 0, line: 240 }, alignment: AlignmentType.LEFT,
    children: [new TextRun({ text, font: FONT, size: NOTE })] })] };
  return new FootnoteReferenceRun(footnoteId);
}

// ---- le glossaire : ses termes apparaissent en italique dans le texte (§ 4.7)
const glossaryRegex = (() => {
  if (!GLOSSARY.length) return null;
  const esc = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const variants = [];
  for (const { term } of GLOSSARY) {
    variants.push(esc(term));
    if (term[0] === term[0].toLowerCase()) variants.push(esc(term[0].toUpperCase() + term.slice(1)));
  }
  variants.sort((a, b) => b.length - a.length);
  return new RegExp("(?<![\\p{L}\\p{N}-])(" + variants.join("|") + ")(?![\\p{L}\\p{N}-])", "gu");
})();
function textRuns(text, fmt, italicizeGlossary) {
  if (!italicizeGlossary || !glossaryRegex || fmt.italics) return [new TextRun({ text, ...fmt })];
  const out = [];
  let last = 0;
  for (const m of text.matchAll(glossaryRegex)) {
    if (m.index > last) out.push(new TextRun({ text: text.slice(last, m.index), ...fmt }));
    out.push(new TextRun({ text: m[0], ...fmt, italics: true }));
    last = m.index + m[0].length;
  }
  if (last < text.length) out.push(new TextRun({ text: text.slice(last), ...fmt }));
  return out;
}

// ---- paragraphes
function p(text, opts = {}) {
  return new Paragraph({
    spacing: { after: 0, line: LINE },
    alignment: AlignmentType.JUSTIFIED,
    ...opts.para,
    children: [new TextRun({ text, ...bodyFont, ...opts.run })],
  });
}
function runs(parts, para = {}) {
  return new Paragraph({ spacing: { after: 0, line: LINE }, alignment: AlignmentType.JUSTIFIED, ...para,
    children: parts.map((x) => new TextRun({ ...bodyFont, ...x })) });
}
function blank() {
  return new Paragraph({ spacing: { after: 0, line: LINE }, children: [] });
}
function bullets(items) {
  return items.map((t) => new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 0, line: LINE },
    alignment: AlignmentType.JUSTIFIED, children: [new TextRun({ text: t, ...bodyFont })] }));
}
function numbered(items, ref) {
  numberingRefs.add(ref);
  return items.map((t) => new Paragraph({ numbering: { reference: ref, level: 0 }, spacing: { after: 0, line: LINE },
    alignment: AlignmentType.JUSTIFIED, children: [new TextRun({ text: t, ...bodyFont })] }));
}
// Un passage à écrire, ou une aide générée : visible, sur fond jaune, jamais dans la version finale.
function placeholder(text) {
  return new Paragraph({
    spacing: { before: 120, after: 120, line: LINE },
    shading: { type: ShadingType.CLEAR, color: "auto", fill: "FFF2A8" },
    children: [new TextRun({ text, ...bodyFont, italics: true })],
  });
}
function caption(text) {
  return new Paragraph({ spacing: { before: 60, after: RETURN, line: 240 }, alignment: AlignmentType.LEFT,
    children: [new TextRun({ text, ...smallFont, italics: true })] });
}
const pageBreak = () => new Paragraph({ children: [new PageBreak()] });

// ---- titres : en gras, taille 12, un seul retour avant et après ; chaque titre est noté pour le sommaire
const HEADINGS = [];
function heading(text, level, opts = {}) {
  HEADINGS.push({ text, level, annex: !!opts.annex, label: opts.label || "", role: opts.role || 0, toc: opts.toc !== false });
  return new Paragraph({
    heading: HEADING_LEVELS[level - 1], keepNext: true, keepLines: true, pageBreakBefore: !!opts.newPage,
    spacing: { before: RETURN, after: RETURN, line: LINE }, alignment: AlignmentType.LEFT,
    children: [new TextRun({ text, ...bodyFont, bold: true, color: "000000" })],
  });
}

// ---- tableaux
function cell(content, width, opts = {}) {
  const children = Array.isArray(content)
    ? inlineRuns(content, smallFont, { glossary: false })
    : [new TextRun({ text: String(content == null ? "" : content), ...smallFont, bold: !!opts.head })];
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    shading: opts.head ? { type: ShadingType.CLEAR, color: "auto", fill: "E7E6E6" } : undefined,
    margins: { top: 40, bottom: 40, left: 80, right: 80 },
    children: [new Paragraph({ spacing: { after: 0, line: 240 }, alignment: opts.right ? AlignmentType.RIGHT : AlignmentType.LEFT, children })],
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

// ---- la prose du groupe : runs et blocs
// Un jeton résolu s'imprime « valeur unité » en notation française (25 423 M€, 5,1 %), suivi d'un appel de note.
// [M:…] est un résultat du classeur du groupe : sa note dit quel onglet et quelle version. Pour les autres, la note donne l'identifiant
// du registre, la source (fichier, page, URL, date de consultation) et dit si la ligne a été relue.
// Un jeton inconnu reste visible, en rouge sur jaune. Les termes du glossaire passent en italique.
function inlineRuns(inlines, base = bodyFont, opts = {}) {
  const glossary = opts.glossary !== false;
  const out = [];
  for (const r of inlines || []) {
    const fmt = { ...base, bold: !!r.bold, italics: !!r.italic };
    if (r.kind === "text") out.push(...textRuns(r.text, fmt, glossary));
    else if (r.kind === "code") out.push(new TextRun({ text: r.text, ...fmt, font: "Courier New" }));
    else if (r.kind === "link") out.push(new ExternalHyperlink({ link: r.url, children: [new TextRun({ text: r.text, ...fmt, style: "Hyperlink" })] }));
    else if (r.kind === "wikilink") out.push(new TextRun({ text: r.text, ...fmt, italics: true }));
    else if (r.kind === "token") {
      if (r.status !== "ok") out.push(new TextRun({ text: `[${r.id} ?]`, ...fmt, bold: true, color: "C00000", highlight: "yellow" }));
      else {
        out.push(new TextRun({ text: r.display || `${r.value}${r.unit ? " " + r.unit : ""}`, ...fmt }));
        const key = noteKey(r);
        if (!noted.has(key)) {
          noted.add(key);
          out.push(footnote(groupNote(key, r)));
        }
      }
    }
  }
  return out;
}
function plain(inlines) {
  return (inlines || []).map((r) => (r.kind === "token" ? (r.status === "ok" ? (r.display || `${r.value}${r.unit ? " " + r.unit : ""}`) : `[${r.id} ?]`) : r.text)).join("").trim();
}
// Les blocs d'un fichier texte. `part` numérote les sous-parties (1.1, 1.1.1) ; sans lui (introduction,
// conclusion, remerciements) les sous-titres restent sans numéro, comme la méthode le permet.
function renderBlocks(blocks, opts = {}) {
  const part = opts.part || 0;
  noted = new Set();   // chaque partie redonne la source de ses chiffres
  collectGroups(blocks);
  const counters = [0, 0];
  const out = [];
  for (const b of blocks || []) {
    if (b.kind === "heading") {
      const level = Math.min(3, Math.max(2, b.level));
      let number = "";
      if (part) {
        if (level === 2) { counters[0] += 1; counters[1] = 0; number = `${part}.${counters[0]} `; }
        else { counters[1] += 1; number = `${part}.${counters[0] || 1}.${counters[1]} `; }
      }
      // dans une annexe, les sous-titres restent hors du sommaire : il ne liste que le corps (§ 5.4)
      out.push(heading(number + plain(b.inlines), level, opts.inAnnex ? { toc: false } : {}));
    } else if (b.kind === "paragraph") {
      out.push(new Paragraph({ spacing: { after: PARA_AFTER, line: LINE }, alignment: AlignmentType.JUSTIFIED, children: inlineRuns(b.inlines) }));
    } else if (b.kind === "quote") {
      // Une citation recopiée exactement, entre guillemets français (§ 5.6) ; le retrait la signale.
      const text = plain(b.inlines);
      const open = /^[«"“]/.test(text), close = /[»"”]$/.test(text);
      out.push(new Paragraph({ spacing: { after: PARA_AFTER, line: LINE }, alignment: AlignmentType.JUSTIFIED, indent: { left: 567, right: 567 },
        children: [...(open ? [] : [new TextRun({ text: "« ", ...bodyFont })]), ...inlineRuns(b.inlines, bodyFont, { glossary: false }), ...(close ? [] : [new TextRun({ text: " »", ...bodyFont })])] }));
    } else if (b.kind === "list") {
      let ref = "bullets";
      if (b.ordered) { ref = `list${dynamicLists++}`; numberingRefs.add(ref); }
      b.items.forEach((item, i) => {
        out.push(new Paragraph({ numbering: { reference: ref, level: 0 }, spacing: { after: i === b.items.length - 1 ? PARA_AFTER : 0, line: LINE },
          alignment: AlignmentType.JUSTIFIED, children: inlineRuns(item) }));
      });
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
      out.push(new Paragraph({ spacing: { after: PARA_AFTER, line: 240 }, children: [] }));
    }
  }
  return out;
}
function prose(key) {
  const t = TEXTS[key];
  return t && t.has_prose ? t : null;
}
function signature(key) {
  const meta = TEXTS[key] && TEXTS[key].meta ? TEXTS[key].meta : {};
  const signed = (meta.signe_par || "").trim(), read = (meta.relu_par || "").trim();
  if (!signed) return null;
  return `Partie rédigée par ${signed}${read ? `, relue par ${read}` : ""}.`;
}
// Une entrée du sommaire ou de la table des annexes : le titre, des points de conduite, la page.
function tocLine(text, page, level) {
  return new Paragraph({
    spacing: { after: level === 1 ? 60 : 0, line: 276 }, alignment: AlignmentType.LEFT,
    indent: level > 1 ? { left: 567 * (level - 1) } : undefined,
    tabStops: [{ type: TabStopType.RIGHT, position: TEXT_WIDTH, leader: LeaderType.DOT }],
    children: [new TextRun({ children: [text, new Tab(), page || ""], ...bodyFont, bold: level === 1 })],
  });
}

// =================================================================== I. Pages liminaires, sans numéro
const front = [];

// ---- page de couverture (§ 5.2) : logo, formation, entreprise, titre, prénoms et NOMS, année, responsable
const center = (children, spacing) => new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 0, line: LINE, ...spacing }, children });
if (COVER.logo && fs.existsSync(COVER.logo)) {
  const type = /\.jpe?g$/i.test(COVER.logo) ? "jpg" : "png";
  front.push(center([new ImageRun({ type, data: fs.readFileSync(COVER.logo), transformation: { width: 220, height: 80 } })], { before: 0, after: 240 }));
} else if (!FINAL) {
  front.push(placeholder("[Logo de l'EDHEC : déposer 08_Livrables/scripts/logo-edhec.png]"));
}
front.push(center([new TextRun({ text: COVER.formation || "EDHEC Business School", ...bodyFont })], { after: 0 }));
front.push(center([new TextRun({ text: `Rapport réalisé pour ${COVER.entreprise || "Veolia Environnement"}`, ...bodyFont })], { after: 1800 }));
front.push(center([new TextRun({ text: DATA.title, font: FONT, size: 40, bold: true })], { after: 240 }));
front.push(center([new TextRun({ text: DATA.subtitle, font: FONT, size: 26 })], { after: DATA.edition ? 240 : 1800 }));
// Deux éditions du même rapport : la courte, conforme au cahier des charges, et la longue, remise en complément.
if (DATA.edition) front.push(center([new TextRun({ text: DATA.edition === "longue" ? "Version longue" : "Version courte", ...bodyFont, italics: true })], { after: 1560 }));
if (COVER.auteurs && COVER.auteurs.length) {
  for (const name of COVER.auteurs) front.push(center([new TextRun({ text: name, ...bodyFont })]));
} else if (!FINAL) {
  front.push(placeholder("[Prénom et NOM de chaque auteur : roles.yaml, clé « auteurs »]"));
}
front.push(center([new TextRun({ text: COVER.annee || "Année 2026-2027", ...bodyFont })], { before: 600 }));
front.push(center([new TextRun({ text: COVER.encadrant ? `Responsable : ${COVER.encadrant}` : "", ...bodyFont })], { before: 240 }));
if (!COVER.encadrant && !FINAL) front.push(placeholder("[Nom du responsable : roles.yaml, clé « encadrant »]"));
if (!FINAL) {
  // brouillon ou squelette selon l'état réel du texte ; date à la française (« 9 novembre 2026 »)
  const todoCount = Object.values((CONTROLS && CONTROLS.todos) || {}).reduce((a, b) => a + (b || 0), 0);
  const MONTHS = ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre"];
  const iso = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(DATA.generated || ""));
  const when = iso ? `${Number(iso[3])} ${MONTHS[Number(iso[2]) - 1]} ${iso[1]}` : String(DATA.generated || "");
  front.push(center([new TextRun({ text: `${todoCount ? "Squelette" : "Brouillon à relire,"} généré le ${when} depuis la dataroom du groupe : ${DATA.n_figures} chiffres sourcés, contrôles du classeur ${DATA.checks}. `
    + (todoCount ? "Les passages surlignés sont à rédiger par l'auteur de la partie ; les notes" : "Les notes")
    + " de bas de page disent quelles lignes du registre n'ont pas encore été relues.", ...smallFont, color: "555555" })], { before: 600, line: 240 }));
  front.push(center([new TextRun({ text: "Édition vérifiable, où chaque chiffre ouvre sa page source : ", ...smallFont, color: "555555" }),
    new ExternalHyperlink({ link: WEB, children: [new TextRun({ text: WEB, ...smallFont, style: "Hyperlink" })] })], { line: 240 }));
}
// ---- la page blanche qui suit la couverture (§ 5.3)
front.push(pageBreak(), blank(), pageBreak());
// ---- remerciements (§ 5.3) : sobres, écrits par le groupe
front.push(heading("Remerciements", 1, { toc: false }));
const thanks = prose("remerciements");
if (thanks) front.push(...renderBlocks(thanks.blocks));
else if (!FINAL) front.push(placeholder("[À rédiger] 08_Livrables/texte/remerciements.md : des remerciements sobres, à Veolia et aux personnes qui ont éclairé le travail."));

// =================================================================== II. Le corps, numéroté à partir du sommaire
const body = [];
// ---- le sommaire (§ 5.4), page 1 : rempli au second passage, quand le PDF a donné ses pages
const tocMarker = body.length;
body.push(heading("Sommaire", 1, { toc: false }));
// (les entrées sont insérées à la fin, quand tous les titres sont connus)

// ---- introduction (§ 4.3) : rédigée en dernier par le groupe ; situe le cadre, présente la mission, annonce les résultats et le plan
body.push(heading("Introduction", 1, { newPage: true }));
const intro = prose("introduction");
if (intro) body.push(...renderBlocks(intro.blocks));
else if (!FINAL) {
  body.push(placeholder("[À rédiger en dernier] 08_Livrables/texte/introduction.md : le cadre (Veolia, GreenUp 2027, le sujet 2), la mission du groupe, les résultats annoncés, le plan. Ci-dessous, la matière générée depuis le classeur, à réécrire :"));
  body.push(p(DATA.thread), blank());
  body.push(runs([{ text: "Les quatre questions de recherche de Veolia", bold: true }]));
  body.push(...numbered(DATA.research_questions, "numbers"), blank());
  body.push(runs([{ text: "Le chantier « capacité » en cinq blocs", bold: true }]));
  body.push(...numbered(DATA.capacity_blocks, "numbers2"));
}

// ---- le développement : une partie numérotée par rôle, puis la recommandation (§ 4.2, § 4.4)
let part = 0;
for (const role of DATA.roles) {
  part += 1;
  const key = `role-${role.n}`;
  const text = prose(key);
  const title = (TEXTS[key] && TEXTS[key].title) || role.name;
  body.push(heading(`${part}. ${title}`, 1, { role: role.n , newPage: true }));
  const signed = signature(key);
  if (signed) body.push(p(signed, { run: { italics: true } }), blank());
  else if (!FINAL) body.push(placeholder(`[Partie rédigée par : ______ — frontmatter signe_par / relu_par de texte/${key}.md] ${role.tabs ? "Dans le classeur : " + role.tabs : ""}`));
  // Une partie rédigée se lit seule : les aides générées (consigne, chiffres du rôle, questions du jury)
  // ne s'impriment qu'à côté d'une partie encore vide.
  if (!FINAL && !text) body.push(p(role.brief, { run: { italics: true, color: "555555" } }), blank());
  if (text) {
    body.push(...renderBlocks(text.blocks, { part }));
  } else if (!FINAL) {
    body.push(placeholder(`[À rédiger] 08_Livrables/texte/${key}.md : 6 à 9 pages, chaque chiffre cité par son jeton [F…] du registre ; des titres habillés (ce que la partie démontre), jamais « analyse » ou « résultats ».`));
    body.push(runs([{ text: "Réponse générée depuis le classeur, à réécrire", bold: true }]), p(role.answer), blank());
    body.push(runs([{ text: "Raisonnement généré", bold: true }]), ...numbered(role.reasoning, `role${role.n}`), blank());
  }
  if (!FINAL && !text) {
    body.push(runs([{ text: "Chiffres du registre pour cette partie", bold: true }]));
    body.push(table(["ID", "Chiffre", "Période", "Valeur", "Source (fichier, page)", "Relu par"],
      role.figures.map((f) => [f[0], f[1], f[2], f[3], f[4], f[5] || "—"]), [900, 3000, 1000, 1300, 2172, 700], [3]));
    body.push(caption(`Aide à la rédaction, retirée de la version finale : les ${role.figures.length} lignes du registre attribuées au rôle ${role.n}.`));
    body.push(runs([{ text: "Ce que nous n'avons pas pu établir", bold: true }]), ...bullets(role.open), blank());
    body.push(runs([{ text: "Questions probables du jury", bold: true }]));
    for (const [q, a] of role.qa) { body.push(runs([{ text: q, bold: true }])); body.push(p(a)); }
  }
}
// ---- la recommandation, dernière partie du développement
part += 1;
const reco = prose("recommandation");
const recoTitle = (TEXTS.recommandation && TEXTS.recommandation.title) || "Ce que nous recommandons";
body.push(heading(`${part}. ${recoTitle}`, 1, { newPage: true }));
const recoSigned = signature("recommandation");
if (recoSigned) body.push(p(recoSigned, { run: { italics: true } }), blank());
if (reco) body.push(...renderBlocks(reco.blocks, { part }));
else if (!FINAL) body.push(placeholder("[À rédiger] 08_Livrables/texte/recommandation.md : ce que nous recommandons, ce que cela coûte et à qui, ce qui nous ferait changer d'avis. Chiffres par jeton."));
body.push(table(["Indicateur", "Valeur", "Ce qu'il faut en lire"], DATA.tiles.map((t) => [t[0], t[1], t[2]]), [3200, 1600, 4272], [1]));
body.push(caption(`Tableau ${part}.1 — Les indicateurs de la réponse, calculés dans le classeur du groupe à partir des chiffres du registre (${DATA.n_figures} lignes sourcées).`));

// ---- conclusion (§ 4.5) : les points forts à retenir
body.push(heading("Conclusion", 1, { newPage: true }));
const concl = prose("conclusion");
if (concl) body.push(...renderBlocks(concl.blocks));
else if (!FINAL) {
  body.push(placeholder("[À rédiger] 08_Livrables/texte/conclusion.md : les points forts à retenir, la réponse en trois phrases, ce que le groupe retient de la méthode."));
  body.push(runs([{ text: "Ce qu'il reste à faire avant la remise", bold: true }]), ...bullets(DATA.todo));
}

// ---- bibliographie raisonnée (§ 4.6, § 5.7.3) : les documents cités, par type de source, puis par auteur
body.push(heading("Bibliographie", 1, { newPage: true }));
const biblio = DATA.bibliography || [];
if (biblio.length) {
  body.push(p("Les documents cités dans ce rapport, classés par type de source puis par ordre alphabétique d'auteur. Chaque référence donne la date de consultation ; les pages citées figurent dans les notes de bas de page et à l'annexe I."), blank());
  for (const group of biblio) {
    body.push(runs([{ text: group.category, bold: true }]));
    for (const entry of group.entries) {
      body.push(new Paragraph({ spacing: { after: 0, line: LINE }, alignment: AlignmentType.LEFT, indent: { left: 567, hanging: 567 },
        children: [new TextRun({ text: entry, ...bodyFont })] }));
    }
    body.push(blank());
  }
} else if (!FINAL) {
  body.push(placeholder("La bibliographie se remplit d'elle-même : chaque document dont une ligne du registre est citée par jeton y prend sa place, par type de source puis par auteur."));
}

// ---- glossaire (§ 4.7) : les termes techniques et sigles, en italique dans le texte, définis ici par ordre alphabétique
body.push(heading("Glossaire", 1, { newPage: true }));
if (GLOSSARY.length) {
  body.push(p("Les termes techniques, expressions étrangères et sigles employés dans le rapport ; ils apparaissent en italique dans le texte."), blank());
  for (const { term, definition } of GLOSSARY) {
    body.push(new Paragraph({ spacing: { after: 0, line: LINE }, alignment: AlignmentType.JUSTIFIED, indent: { left: 567, hanging: 567 },
      children: [new TextRun({ text: term, ...bodyFont, bold: true }), new TextRun({ text: ` : ${definition}`, ...bodyFont })] }));
  }
} else if (!FINAL) {
  body.push(placeholder("Le glossaire se remplit d'elle-même des termes définis dans report.py (GLOSSARY) que le texte emploie."));
}

// =================================================================== III. Les annexes, paginées à part en chiffres romains (§ 5.8)
// Chaque annexe est numérotée, titrée, légendée ; le texte doit la commenter (le contrôle le vérifie).
const annexes = [];
const ROMAN = ["I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X"];
let annexNo = 0;
function annex(title, children, legend) {
  const label = `Annexe ${ROMAN[annexNo++] || annexNo}`;
  const text = `${label} — ${title}`;
  if (annexes.length) annexes.push(pageBreak());
  annexes.push(heading(text, 1, { annex: true, label }));
  annexes.push(caption(legend));
  annexes.push(...children);
}
// Les identifiants cités dans la prose, dans l'ordre de lecture : l'annexe I ne reprend qu'eux.
const citedIds = new Set();
(function walk(node) {
  if (Array.isArray(node)) node.forEach(walk);
  else if (node && typeof node === "object") {
    if (node.kind === "token" && node.status === "ok") citedIds.add(String(node.id).split("-")[0]);
    for (const k of ["blocks", "inlines", "items", "rows"]) if (node[k]) walk(node[k]);
  }
})(Object.values(TEXTS));
const citedRows = DATA.register.filter((r) => citedIds.has(String(r[0])));
annex("Registre des chiffres cités", [
  ...(citedRows.length
    ? [table(["ID", "Chiffre", "Valeur", "Période", "Source (voir la bibliographie)", "Rôle", "Relu"],
        citedRows.map((r) => [r[0], r[1], (DATA.displays || {})[r[0]] || r[2], r[3], (DATA.refs || {})[r[0]] || r[4], r[5], r[6] || "—"]),
        [600, 2900, 1100, 1000, 2272, 600, 600], [2])]
    : [FINAL ? p("Aucun chiffre cité.") : placeholder(`Aucun chiffre encore cité par jeton : l'annexe se remplira des lignes citées parmi les ${DATA.register.length} du registre.`)]),
], `${citedRows.length} chiffres cités dans le texte, sur ${DATA.register.length} au registre du groupe. Chaque source renvoie à la bibliographie, qui donne le titre, l'adresse et la date de consultation ; le registre garde le fichier joint. « Relu » vide : un seul membre a lu la page.`);
if (DATA.cited_deals && DATA.cited_deals.length) {
  annex("Opérations citées", [
    table(["ID", "Opération", "Montant", "Devise", "Source", "Relecture"],
      DATA.cited_deals.map((d) => [d[0], d[1], d[2], d[3], d[4], d[5]]), [600, 2600, 1000, 700, 3172, 1000], [2]),
  ], "Les opérations du registre deals.csv citées dans le texte par leur jeton [D…].");
}
if (DATA.cited_targets && DATA.cited_targets.length) {
  annex("Cibles citées", [
    table(["ID", "Cible", "Valeur estimée", "Devise", "Source", "Relecture"],
      DATA.cited_targets.map((t) => [t[0], t[1], t[2], t[3], t[4], t[5]]), [600, 2600, 1000, 700, 3172, 1000], [2]),
  ], "Les cibles du registre targets.csv citées dans le texte par leur jeton [T…].");
}
annex("La fourchette de capacité pour les autres périmètres", [
  p("Le nombre que les groupes « technologies de l'eau », « déchets dangereux » et « bioénergie » doivent respecter, dans ses trois scénarios et sous ses deux plafonds."), blank(),
  table(["Fin 2027", "Défavorable", "Central", "Favorable"], DATA.range_rows, [4272, 1600, 1600, 1600], [1, 2, 3]),
], `Source : classeur du groupe, modele-greenup-2027.xlsx${DATA.model_version ? ", " + DATA.model_version : ""}, calculé à partir des chiffres du registre.`);
const AI_FACTS = (DATA.ai_facts && DATA.ai_facts.length) ? DATA.ai_facts.filter((f) => !f.startsWith("[[")) : [
  "Dataroom : un serveur MCP construit par le groupe pour le sujet, qui indexe les documents publics versés dans le vault et ne sait que chercher, lire une page et citer (fichier, page, URL, date de consultation). Il ne calcule rien et ne résume rien.",
  "Claude (Anthropic), via Claude Code : a construit le serveur, le registre, les contrôles automatiques, le générateur du classeur Excel et celui de ce rapport.",
  "[Provenance et relecture des lignes du registre : générer le rapport par « dataroom report » pour les compter.]",
];
const ia = prose("outils-ia");
annex("Outils d'intelligence artificielle utilisés", [
  ...bullets(AI_FACTS), blank(),
  ...(ia ? renderBlocks(ia.blocks, { inAnnex: true }) : (!FINAL ? [placeholder("[À rédiger] 08_Livrables/texte/outils-ia.md : ce que le groupe a décidé sans l'outil, les cas où l'outil a été contredit après lecture, puis un tableau membre / outil / usage / période.")] : [])),
], "Annexe demandée par le cours : les outils, ce qu'ils ont fait, ce que le groupe a vérifié lui-même. Les phrases comptées le sont depuis les registres, le jour de la génération.");

// ---- la table des annexes (§ 5.8), dernière page numérotée du corps
body.push(heading("Table des annexes", 1, { newPage: true }));
for (const h of HEADINGS.filter((x) => x.annex)) body.push(tocLine(h.text, ANNEX_PAGES.get(h.text) || "", 1));

// ---- le sommaire, maintenant que tous les titres sont connus : parties et sous-parties, puis les annexes en un mot
const tocEntries = HEADINGS.filter((h) => h.toc && !h.annex && h.level <= 2).map((h) => tocLine(h.text, TOC_PAGES.get(h.text) || "", h.level));
const firstAnnex = HEADINGS.find((h) => h.annex);
if (firstAnnex) tocEntries.push(tocLine("Annexes (paginées à part, en chiffres romains)", ANNEX_PAGES.get(firstAnnex.text) || "", 1));
body.splice(tocMarker + 1, 0, ...tocEntries);

// =================================================================== IV. Contrôle des citations (jamais dans la version finale)
const controls = [];
if (CONTROLS && !FINAL) {
  controls.push(heading("Contrôle des citations et de la méthode (retiré de la version finale)", 1, { toc: false }));
  const files = Object.keys(CONTROLS.words || {});
  if (files.length) {
    // l'édition longue est la version de référence : les cibles de pages ne valent que pour la courte
    const longEdition = DATA.edition === "longue";
    const [lo, hi] = CONTROLS.pages_per_role || [6, 9];
    controls.push(table(["Fichier", "Pages (PDF)", "Mots", "Caractères", "Cible pages", "À rédiger"],
      files.map((f) => [f, CONTROLS.pages && CONTROLS.pages[f] != null ? CONTROLS.pages[f] : "—", CONTROLS.words[f],
        CONTROLS.chars && CONTROLS.chars[f] != null ? CONTROLS.chars[f] : "—",
        f.startsWith("role-") && !longEdition ? `${lo}-${hi}` : "—", CONTROLS.todos[f] || 0]), [2772, 1200, 1100, 1300, 1300, 1400], [1, 2, 3, 5]));
    const target = (CONTROLS.pages_total_target || [50, 60]).join("-");
    controls.push(caption(`Corps du rapport, du sommaire à la première annexe : ${CONTROLS.pages_body || "—"} pages ${longEdition ? "(édition de référence, sans cible de pages)" : `(cible ${target})`} ; document entier : ${CONTROLS.pages_total || "—"} pages ; ${CONTROLS.chars_total || 0} caractères espaces compris dans le corps (la méthode en demande ${(DATA.method && DATA.method.chars_min) || 50000} au minimum).`));
  }
  const list = (title, items, fmt) => {
    controls.push(runs([{ text: title, bold: true }]));
    controls.push(...(items.length ? bullets(items.map(fmt)) : [p("aucun", { run: { italics: true } })]), blank());
  };
  list(`Ce que la méthode du rapport ferait souligner (${(CONTROLS.warnings || []).length})`, CONTROLS.warnings || [], (x) => x);
  list(`Jetons inconnus (${CONTROLS.unknown.length})`, CONTROLS.unknown, (x) => `${x[0]} : [${x[1]}] ne correspond à aucune ligne des registres`);
  const unreadBy = {};
  for (const [f, id] of CONTROLS.unread) (unreadBy[f] = unreadBy[f] || []).push(id);
  list(`Lignes citées que personne n'a relues (${CONTROLS.unread.length})`, Object.entries(unreadBy),
    ([f, ids]) => `${f} : ${ids.length} ligne(s) — ${ids.slice(0, 40).join(", ")}${ids.length > 40 ? " …" : ""}`);
  list(`Lignes citées contestées (${CONTROLS.disputed.length})`, CONTROLS.disputed, (x) => `${x[0]} : ${x[1]} — ${x[2]}`);
  list("Fichiers absents ou vides", [...(CONTROLS.missing || []), ...(CONTROLS.empty || [])], (x) => x);
  list("Pages liminaires à écrire", [...(CONTROLS.cover_missing || []).map((x) => `couverture : ${x}`), ...(CONTROLS.front_missing || [])], (x) => x);
  const uncited = Object.entries(CONTROLS.uncited || {});
  list("Chiffres du registre jamais cités par leur rôle", uncited, ([role, ids]) => `rôle ${role} : ${ids.length}${ids.length ? " — " + ids.slice(0, 20).join(", ") + (ids.length > 20 ? " …" : "") : ""}`);
}

// =================================================================== le document
const pageProps = (pageNumbers) => ({ page: { size: { width: PAGE_W, height: PAGE_H }, margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN }, ...(pageNumbers ? { pageNumbers } : {}) } });
const numberFooter = () => new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 0 },
  children: [new TextRun({ children: [PageNumber.CURRENT], ...smallFont })] })] });
const textFooter = (text) => new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 0 },
  children: [new TextRun({ text, ...smallFont, color: "777777" })] })] });
const headingStyle = (id, name, level) => ({
  id, name, basedOn: "Normal", next: "Normal", quickFormat: true,
  run: { font: FONT, size: BODY, bold: true, color: "000000", italics: false },
  paragraph: { spacing: { before: RETURN, after: RETURN, line: LINE }, outlineLevel: level, keepNext: true },
});

const sections = [
  { properties: pageProps(null), footers: { default: textFooter("") }, children: front },
  { properties: pageProps({ start: 1, formatType: NumberFormat.DECIMAL }), footers: { default: numberFooter() }, children: body },
  { properties: pageProps({ start: 1, formatType: NumberFormat.UPPER_ROMAN }), footers: { default: numberFooter() }, children: annexes },
];
if (controls.length) sections.push({ properties: pageProps(null), footers: { default: textFooter("Contrôle — hors version finale") }, children: controls });

const doc = new Document({
  creator: "Groupe sujet 2",
  title: DATA.title,
  styles: {
    default: { document: { run: bodyFont, paragraph: { spacing: { line: LINE, after: 0 }, alignment: AlignmentType.JUSTIFIED } } },
    paragraphStyles: [
      headingStyle("Heading1", "Heading 1", 0),
      headingStyle("Heading2", "Heading 2", 1),
      headingStyle("Heading3", "Heading 3", 2),
      { id: "FootnoteText", name: "Footnote Text", basedOn: "Normal", run: { font: FONT, size: NOTE }, paragraph: { spacing: { line: 240, after: 0 }, alignment: AlignmentType.LEFT } },
    ],
  },
  footnotes,
  numbering: {
    config: [
      { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "–", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 567, hanging: 283 } } } }] },
      ...[...numberingRefs].map((ref) => ({
        reference: ref, levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 567, hanging: 360 } } } }],
      })),
    ],
  },
  sections,
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(OUT, buf);
  fs.writeFileSync(HEADINGS_OUT, JSON.stringify(HEADINGS.map(({ text, level, annex, label, role }) => ({ text, level, annex, label, role })), null, 1));
  console.log("écrit", OUT, Math.round(buf.length / 1024), "Ko", `(${HEADINGS.length} titres, ${footnoteId} notes)` + (FINAL ? " (version finale)" : ""));
});
