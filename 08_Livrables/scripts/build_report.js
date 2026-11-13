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
// La mise en page se tient seule, dans Word comme dans LibreOffice : contrôle des veuves et orphelines, titres liés à
// ce qui les suit et coupés en lignes équilibrées, lignes de tableau insécables, colonnes assez larges pour leur mot
// le plus long, petits tableaux d'un seul tenant avec la phrase qui les annonce, sauts de page portés par le titre
// qui ouvre la page. dataroom report relit ensuite le PDF (layout_findings) et dit ce qui a échappé.
// Écrit, à côté du .docx, la liste ordonnée des titres (…headings.json) : dataroom report la
// confronte au PDF pour numéroter le sommaire et la table des annexes au second passage.
"use strict";
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun: RawTextRun, HeadingLevel, Table, TableRow, TableCell, WidthType,
  AlignmentType, Footer, PageNumber, ShadingType, LevelFormat, BorderStyle,
  ExternalHyperlink, FootnoteReferenceRun, ImageRun, NumberFormat, Tab, TabStopType, LeaderType, UnderlineType,
  HorizontalPositionRelativeFrom, VerticalPositionRelativeFrom, TextWrappingType, HeightRule, VerticalAlign,
} = require("docx");

// Aucun tiret long dans le document : dans toutes les chaînes reçues (notes, titres, légendes), « A — B » devient « A : B ».
const noDash = (v) => (typeof v === "string" ? v.replace(/\s—\s/g, " : ")
  : Array.isArray(v) ? v.map(noDash) : v && typeof v === "object" ? Object.fromEntries(Object.entries(v).map(([k, x]) => [k, noDash(x)])) : v);
const DATA = noDash(JSON.parse(fs.readFileSync(process.argv[2] || path.join(__dirname, "report-data.json"), "utf8")));
const OUT = process.argv[3] || "rapport-sujet-2.docx";
const HEADINGS_OUT = OUT.replace(/\.docx$/i, "") + ".headings.json";
const FINAL = process.argv.includes("--final") || DATA.final === true;
// L'édition mise en forme : même texte, mêmes chiffres, mêmes notes ; l'habillage suit les « Essentiels de la charte
// graphique » de Veolia (mars 2024) : Arial (police des documents bureautiques), deux couleurs d'une même teinte (marine
// et bleu ciel), texte en noir à 80 %, un marqueur circulaire sur la couverture et les ouvertures de partie. Pas de rouge :
// la charte le réserve aux usages « corporate » ; seul le logotype l'est.
const THEME = (process.argv.find((a) => a.startsWith("--theme=")) || "").slice(8) || String(DATA.theme || "");
const VEOLIA = THEME === "veolia";
const MARINE = "002D62", SKY = "8DACCD", SKY_PALE = "E3EAF2", INK = "55555A", RULE = "D1D3D4";
const ASSETS = path.join(__dirname, "veolia");
// L'édition anglaise : --lang=en. Les textes arrivent traduits (dataroom report, par DeepL) ; les chaînes écrites ici
// passent par T() : les nombres sortent du gabarit ({0}, {1}…), la table DATA.i18n donne l'anglais, et ce qu'elle ne
// connaît pas est noté dans <rapport>.i18n.json pour être traduit avant le passage suivant.
const LANG = (process.argv.find((a) => a.startsWith("--lang=")) || "").slice(7) || String(DATA.lang || "");
const EN = LANG === "en";
const I18N = DATA.i18n || {};
const I18N_MISSING = new Set();
const enNum = (n) => String(n).replace(/[\u202f\u00a0 ](?=\d{3})/g, ",").replace(/(\d),(\d{1,2})$/, "$1.$2");
function T(text) {
  if (!EN) return text;
  const args = [];
  const tpl = String(text).replace(/\d{1,3}(?:[\u202f\u00a0 ]\d{3})+|\d+(?:[.,]\d+)?/g, (m) => { args.push(m); return `{${args.length - 1}}`; });
  const got = I18N[tpl];
  if (got === undefined) { I18N_MISSING.add(tpl); return String(text); }
  return got.replace(/\{(\d+)\}/g, (_, k) => (args[Number(k)] === undefined ? "" : enNum(args[Number(k)])));
}
const SEP = EN ? ": " : " : ";   // la ponctuation haute : sans espace devant, en anglais
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
const FONT = VEOLIA ? "Arial" : "Times New Roman";
const BODY = VEOLIA ? 21 : 24;         // « Times New Roman 12 » ; Arial 10,5 dans l'édition mise en forme
const SMALL = VEOLIA ? 18 : 20;        // 10 pt : tableaux, légendes
const NOTE = VEOLIA ? 16 : 18;         // 9 pt : notes de bas de page
const LINE = VEOLIA ? 300 : 360;       // « interligne 1,5 » ; 1,25 dans l'édition mise en forme
const ALIGN = VEOLIA ? AlignmentType.LEFT : AlignmentType.JUSTIFIED;   // la charte compose au fer à gauche
const MARGIN = 1417;                   // « 2,5 cm pour toutes les marges » : en haut et en bas
// Le rapport est relié : 3 cm à gauche, 2 cm à droite. La largeur du texte ne change pas (16 cm), la mise en page non plus.
const MARGIN_LEFT = 1701, MARGIN_RIGHT = 1133;
const PAGE_W = 11906, PAGE_H = 16838;  // A4
const TEXT_WIDTH = PAGE_W - MARGIN_LEFT - MARGIN_RIGHT; // 9072
const RETURN = 360;                    // « un seul retour avant et après » un titre : une ligne à 1,5
const TOC_RIGHT = 567;                 // sommaire : 1 cm de retrait à droite, où viennent les points de conduite et la page
const CELL_PAD = 160;                  // marges gauche et droite d'une cellule (80 + 80)
const WJ = "\u2060", ZWSP = "\u200b";  // WORD JOINER : rien ne se coupe autour ; ZERO WIDTH SPACE : une coupure permise

const bodyFont = { font: FONT, size: BODY };
const smallFont = { font: FONT, size: SMALL };
const HEADING_LEVELS = [HeadingLevel.HEADING_1, HeadingLevel.HEADING_2, HeadingLevel.HEADING_3];

// ---- la typographie française : les espaces qui ne se coupent pas. Les textes sont tapés avec des espaces simples ;
// ce qui s'imprime passe par typo(), qui lie ce que la règle veut lié. Une espace insécable (U+00A0) après « et avant
// » ; : ? !, entre « p. », « n° », « § », « art. » et leur numéro, entre un jour et son mois, entre IFRS ou IAS et son
// numéro, à l'intérieur d'une unité (« Md AUD ») ; une espace fine insécable (U+202F), celle des valeurs du registre,
// entre les milliers et entre un nombre et son unité. Rien d'autre ne change : ni tiret, ni apostrophe, ni mot.
// Sans effet sur un texte déjà traité. dataroom report relit le PDF et signale ce qui lui aurait échappé (« ponctuation »).
const NBSP = "\u00a0", NNBSP = "\u202f";
const MONTHS = "janvier|février|mars|avril|mai|juin|juillet|août|septembre|octobre|novembre|décembre";
const CURRENCIES = "USD|AUD|GBP|CAD|EUR";
const TYPO = [
  [/« /g, "«" + NBSP],
  [/ »/g, NBSP + "»"],
  [/ (?=[;:?!])/g, NBSP],
  // les milliers tapés avec une espace : « 5 747 », jamais une année suivie d'un nombre (« 2026 150 »), ni un identifiant
  [/(?<![\p{L}\p{N},.])\d{1,3}(?: \d{3})+(?!\d)/gu, (m) => m.replace(/ /g, NNBSP)],
  // l'unité qui contient une espace : « Md AUD », « k USD »
  [new RegExp(`(?<![\\p{L}\\p{N}])(Mds?|M|k) (?=(?:${CURRENCIES})(?![\\p{L}\\p{N}]))`, "gu"), "$1" + NBSP],
  // un nombre et son unité : « 3 % », « 229 M€ », « 470 M$ », « 50 992 k USD », « 12 kt »
  [new RegExp(`(\\d) (?=(?:(?:Mds?|M|k)?[€$£]|[%‰]|kt|Mt|(?:(?:Mds?|M|k)${NBSP})?(?:${CURRENCIES}))(?![\\p{L}\\p{N}]))`, "gu"), "$1" + NNBSP],
  // « p. 12 », « pp. 12 », « n° 3 », « § 4 », « art. L » : l'abréviation garde son numéro
  [/(^|[\s(«])(pp?\.|[nN]°|§|[aA]rt\.) (?=\d|[IVXLC]+(?:[\s.,;:)]|$))/g, "$1$2" + NBSP],
  // le jour et son mois : « 13 novembre », « 1er janvier »
  [new RegExp(`(?<![\\p{L}\\p{N},.${NNBSP}])(\\d{1,2}|1er) (?=(?:${MONTHS})(?![\\p{L}]))`, "gu"), "$1" + NBSP],
  // la norme et son numéro : « IFRS 16 », « IAS 37 »
  [/(?<![\p{L}\p{N}])(IFRS|IAS) (?=\d)/gu, "$1" + NBSP],
];
function typo(text) {
  let s = String(text == null ? "" : text);
  if (EN) return s.replace(/(^|[\s(])(pp?\.) (?=\d)/g, "$1$2" + NBSP);   // en anglais : seul « p. 12 » reste lié
  if (!s.includes(" ")) return s;
  for (const [pattern, by] of TYPO) s = s.replace(pattern, by);
  return s;
}
// Tout ce qui s'imprime passe par ici : un TextRun de la bibliothèque, dont le texte a reçu ses espaces insécables.
// `verbatim: true` le laisse tel quel (du code, une adresse).
function TextRun(options) {
  if (typeof options === "string") return new RawTextRun(typo(options));
  const { verbatim, ...rest } = options || {};
  if (VEOLIA && rest.color === undefined && rest.style !== "Hyperlink") rest.color = INK;
  if (!verbatim) {
    if (typeof rest.text === "string") rest.text = typo(rest.text);
    if (Array.isArray(rest.children)) rest.children = rest.children.map((c) => (typeof c === "string" ? typo(c) : c));
  }
  return new RawTextRun(rest);
}

// ---- les chasses de Times New Roman, en 1/2048 d'em (Liberation Serif, qui la remplace dans LibreOffice, a les mêmes :
// c'est sa raison d'être). De quoi couper un titre en lignes équilibrées et donner à chaque colonne la largeur de son
// mot le plus long, sans attendre le PDF. Les 224 premiers caractères (U+0020 à U+00FF), puis ADVANCE_CHARS.
const ADVANCE = {
  r: [512,682,836,1024,1024,1706,1593,369,682,682,1024,1155,512,682,512,569,1024,1024,1024,1024,1024,1024,1024,1024,1024,1024,569,569,1155,1155,1155,909,1886,1479,1366,1366,1479,1251,1139,1479,1479,682,797,1479,1251,1821,1479,1479,1139,1479,1366,1139,1251,1479,1479,1933,1479,1479,1251,682,569,682,961,1024,682,909,1024,909,1024,909,682,1024,1024,569,569,1024,569,1593,1024,1024,1024,1024,682,797,569,1024,1024,1479,1024,1024,909,983,410,983,1108,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,512,682,1024,1024,1024,1024,410,1024,682,1556,565,1024,1155,682,1556,1024,819,1124,614,614,682,1180,928,682,682,614,635,1024,1536,1536,1536,909,1479,1479,1479,1479,1479,1479,1821,1366,1251,1251,1251,1251,682,682,682,682,1479,1479,1479,1479,1479,1479,1479,1155,1479,1479,1479,1479,1479,1479,1139,1024,909,909,909,909,909,909,1366,909,909,909,909,909,569,569,569,569,1024,1024,1024,1024,1024,1024,1024,1124,1024,1024,1024,1024,1024,1024,1024,1024],
  b: [512,682,1137,1024,1024,2048,1706,569,682,682,1024,1167,512,682,512,569,1024,1024,1024,1024,1024,1024,1024,1024,1024,1024,682,682,1167,1167,1167,1024,1905,1479,1366,1479,1479,1366,1251,1593,1593,797,1024,1593,1366,1933,1479,1593,1251,1593,1479,1139,1366,1479,1479,2048,1479,1479,1366,682,569,682,1190,1024,682,1024,1139,909,1139,909,682,1024,1139,569,682,1139,569,1706,1139,1024,1139,1139,909,797,682,1139,1024,1479,1024,1024,909,807,451,807,1065,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,512,682,1024,1024,1024,1024,451,1024,682,1530,614,1024,1167,682,1530,1024,819,1124,614,614,682,1180,1106,683,682,614,676,1024,1536,1536,1536,1024,1479,1479,1479,1479,1479,1479,2048,1479,1366,1366,1366,1366,797,797,797,797,1479,1479,1593,1593,1593,1593,1593,1167,1593,1479,1479,1479,1479,1479,1251,1139,1024,1024,1024,1024,1024,1024,1479,909,909,909,909,909,569,569,569,569,1024,1139,1024,1024,1024,1024,1024,1124,1024,1139,1139,1139,1139,1024,1139,1024],
  i: [512,682,860,1024,1024,1706,1593,438,682,682,1024,1382,512,682,512,569,1024,1024,1024,1024,1024,1024,1024,1024,1024,1024,682,682,1382,1382,1382,1024,1884,1251,1251,1366,1479,1251,1251,1479,1479,682,909,1366,1139,1706,1366,1479,1251,1479,1251,1024,1139,1479,1251,1706,1251,1139,1139,797,569,797,864,1024,682,1024,1024,909,1024,909,569,1024,1024,569,569,909,569,1479,1024,1024,1024,1024,797,797,569,1024,909,1366,909,909,797,819,563,819,1108,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,512,797,1024,1024,1024,1024,563,1024,682,1556,565,1024,1382,682,1556,1024,819,1124,614,614,682,1180,1071,512,682,614,635,1024,1536,1536,1536,1024,1251,1251,1251,1251,1251,1251,1821,1366,1251,1251,1251,1251,682,682,682,682,1479,1366,1479,1479,1479,1479,1479,1382,1479,1479,1479,1479,1479,1139,1251,1024,1024,1024,1024,1024,1024,1024,1366,909,909,909,909,909,569,569,569,569,1024,1024,1024,1024,1024,1024,1024,1124,1024,1024,1024,1024,1024,909,1024,909],
  bi: [512,797,1137,1024,1024,1706,1593,569,682,682,1024,1167,512,682,512,569,1024,1024,1024,1024,1024,1024,1024,1024,1024,1024,682,682,1167,1167,1167,1024,1704,1366,1366,1366,1479,1366,1366,1479,1593,797,1024,1366,1251,1821,1479,1479,1251,1479,1366,1139,1251,1479,1366,1821,1366,1251,1251,682,569,682,1167,1024,682,1024,1024,909,1024,909,682,1024,1139,569,569,1024,569,1593,1139,1024,1024,1024,797,797,569,1139,909,1366,1024,909,797,713,451,713,1167,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,512,797,1024,1024,1024,1024,451,1024,682,1530,545,1024,1241,682,1530,1024,819,1124,614,614,682,1180,1024,512,682,614,614,1024,1536,1536,1536,1024,1366,1366,1366,1366,1366,1366,1933,1366,1366,1366,1366,1366,797,797,797,797,1479,1479,1479,1479,1479,1479,1479,1167,1479,1479,1479,1479,1479,1251,1251,1024,1024,1024,1024,1024,1024,1024,1479,909,909,909,909,909,569,569,569,569,1024,1139,1024,1024,1024,1024,1024,1124,1024,1139,1139,1139,1139,909,1024,909],
};
const ADVANCE_CHARS = "\u0152\u0153\u0178\u2009\u200b\u2010\u2012\u2013\u2014\u2015\u2016\u2017\u2018\u2019\u201a\u201b\u201c\u201d\u201e\u201f\u2020\u2021\u2022\u2026\u202f\u2030\u2032\u2033\u2039\u203a\u20ac\u2122\u2190\u2191\u2192\u2193\u2212\u2248\u2260\u2264\u2265";
const ADVANCE_MORE = {
  r: [1821,1479,1479,410,0,682,1024,1024,2048,2048,575,1024,682,682,682,682,909,909,909,909,1024,1024,717,2048,410,2048,448,853,682,682,1024,2007,2048,1024,2048,1024,1155,1124,1124,1124,1124],
  b: [2048,1479,1479,410,0,682,1024,1024,2048,2048,664,1024,682,682,682,682,1024,1024,1024,1024,1024,1024,717,2048,410,2048,576,1131,682,682,1024,2048,2048,1024,2048,1024,1167,1124,1124,1124,1124],
  i: [1933,1366,1139,410,0,682,1024,1024,1821,2048,575,1024,682,682,682,682,1139,1139,1139,1139,1024,1024,717,1821,410,2048,448,853,682,682,1024,2007,2048,1024,2048,1024,1382,1124,1124,1124,1124],
  bi: [1933,1479,1251,410,0,682,1024,1024,2048,2048,708,1024,682,682,682,682,1024,1024,1024,1024,1024,1024,717,2048,410,2048,576,1131,682,682,1024,2048,2048,1024,2048,1024,1241,1124,1124,1124,1124],
};
function advance(ch, style) {
  const cp = ch.codePointAt(0);
  if (cp === 0x200b || cp === 0x2060 || cp === 0xad) return 0;
  const more = ADVANCE_CHARS.indexOf(ch);
  const v = cp >= 32 && cp < 256 ? ADVANCE[style][cp - 32] : (more >= 0 ? ADVANCE_MORE[style][more] : 0);
  return Math.round((v || 1024) * (VEOLIA ? 1.14 : 1));   // un caractère hors table compte pour un demi-em ; Arial : environ 14 % de plus
}
// Largeur d'un texte en twips ; `size` en demi-points, `style` r, b, i ou bi.
function textWidth(text, size, style = "r") {
  let units = 0;
  for (const ch of String(text)) units += advance(ch, style);
  return (units / 2048) * size * 10;
}
// Où Word et LibreOffice peuvent couper une ligne : à une espace simple (qui disparaît en fin de ligne), après un trait
// d'union ou un tiret suivi d'une lettre (jamais devant un chiffre), à une espace sans chasse (U+200B) ; jamais autour
// d'un WORD JOINER. Chaque segment dit s'il est suivi d'une espace.
function segments(text) {
  const out = [];
  const chars = Array.from(typo(text));   // les espaces insécables d'abord : on mesure ce qui s'imprime
  let cur = "";
  for (let i = 0; i < chars.length; i++) {
    const ch = chars[i];
    if (ch === " ") {
      if (cur) out.push({ text: cur, space: true });
      else if (out.length) out[out.length - 1].space = true;
      cur = "";
    } else if (ch === ZWSP) {
      if (cur) out.push({ text: cur, space: false });
      cur = "";
    } else {
      cur += ch;
      const next = chars[i + 1];
      if ("-‐–—".includes(ch) && cur.length > 1 && next && /\p{L}/u.test(next) && chars[i - 1] !== WJ) {
        out.push({ text: cur, space: false });
        cur = "";
      }
    }
  }
  if (cur) out.push({ text: cur, space: false });
  return out;
}
// La coupure naturelle : chaque ligne prend autant de segments qu'elle peut.
function layoutLines(segs, maxWidth, size, style) {
  const lines = [];
  const space = textWidth(" ", size, style);
  let line = [], width = 0;
  for (const s of segs) {
    const w = textWidth(s.text, size, style);
    const sep = line.length && line[line.length - 1].space ? space : 0;
    if (line.length && width + sep + w > maxWidth) {
      lines.push(line);
      line = [s];
      width = w;
    } else {
      line.push(s);
      width += sep + w;
    }
  }
  if (line.length) lines.push(line);
  return lines;
}
const lineText = (line) => line.map((s, i) => s.text + (i < line.length - 1 && s.space ? " " : "")).join("");
// Une coupure équilibrée : autant de lignes que la coupure naturelle, mais la plus étroite qui les tienne, si bien que
// la dernière ligne ne porte pas un ou deux mots seuls. 2 pt de marge sur la largeur mesurée ; null quand le texte
// tient sur une ligne, ou n'a pas besoin de cette marge pour tenir en autant de lignes.
function balancedLines(text, maxWidth, size, style) {
  const segs = segments(text);
  const natural = layoutLines(segs, maxWidth, size, style).length;
  const room = maxWidth - 40;
  if (natural < 2 || layoutLines(segs, room, size, style).length > natural) return null;
  let lo = 0, hi = room;
  for (let k = 0; k < 30; k++) {
    const mid = (lo + hi) / 2;
    if (layoutLines(segs, mid, size, style).length <= natural) hi = mid; else lo = mid;
  }
  const lines = layoutLines(segs, hi, size, style);
  // Garde-fou : si la dernière ligne n'a encore qu'un ou deux mots, elle prend le dernier mot de la précédente.
  const words = (line) => lineText(line).split(" ").length;
  const n = lines.length;
  while (n > 1 && words(lines[n - 1]) <= 2 && lines[n - 2].length > 1 && words(lines[n - 2]) > 3) {
    const moved = lines[n - 2].pop();
    lines[n - 1].unshift(moved);
    if (textWidth(lineText(lines[n - 1]), size, style) > room) { lines[n - 1].shift(); lines[n - 2].push(moved); break; }
  }
  // chaque ligne garde l'espace qui la suivait (rien après un trait d'union) : le texte du titre ne change pas
  return lines.map((line, i) => lineText(line) + (i < n - 1 && line[line.length - 1].space ? " " : ""));
}
// Les runs d'un texte coupé en lignes : un retour à la ligne explicite entre elles, l'espace gardée en fin de ligne
// (le texte du titre, et donc le signet du PDF, reste le même).
function linedRuns(text, lines, fmt) {
  if (!lines) return [new TextRun({ text, ...fmt })];
  return lines.map((line, i) => new TextRun({ text: line, ...(i ? { break: 1 } : {}), ...fmt }));
}
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
  if (r.rowkind === "model") return "M|" + String(NOTES[r.id] || r.id).replace(/^\S+ (?:—|:) /, "");
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
// Une note du classeur se lit comme une référence : « Classeur du groupe, onglet Trajectoire, scénario défavorable
// combiné », la clé du jeton (M:HEAD27@def) à la fin, pour qui veut retrouver la cellule.
function classeurNote(where) {
  let s = String(where || "").trim().replace(/\.$/, "");
  let tail = "";
  const paren = /\s*(\([^()]*\))$/.exec(s);
  if (paren) { tail = " " + paren[1]; s = s.slice(0, paren.index); }
  const m = /(?:^|,\s*)onglet (.+?) du classeur(?: du groupe)?$/.exec(s);
  if (!m) return `Classeur du groupe : ${s}${tail}`;
  const rest = s.slice(0, m.index).replace(/[,\s]+$/, "");
  return `Classeur du groupe, onglet ${m[1]}${rest ? ", " + rest : ""}${tail}`;
}
function groupNote(key, r) {
  const g = groups.get(key) || { ids: [r.id], unread: !r.checked, model: r.rowkind === "model" };
  if (g.model) {
    // la version de la fourchette n'est dite qu'une fois : dans la première note du modèle
    let text = T(classeurNote(key.slice(2)));
    if (!modelVersionShown && DATA.model_version) { text += EN ? `; ${DATA.model_version}, valid for every workbook note` : ` ; ${DATA.model_version}, valable pour toutes les notes du classeur`; modelVersionShown = true; }
    return `${text}${EN ? ";" : " ;"} ${T(g.ids.length > 1 ? "clés" : "clé")} ${g.ids.join(", ")}.`;
  }
  const ref = refOf(g.ids[0]) || String(NOTES[g.ids[0]] || r.cite || r.file || "source au registre").replace(/^\S+ (?:—|:) /, "").replace(/ ; non relue\.$|\.$/, "");
  return `${g.ids.join(", ")}${SEP}${ref}.`;
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
  const spaced = (s) => esc(s).replace(/ /g, `[ ${NBSP}${NNBSP}]`);
  for (const { term } of GLOSSARY) {
    variants.push(spaced(term));
    if (term[0] === term[0].toLowerCase()) variants.push(spaced(term[0].toUpperCase() + term.slice(1)));
  }
  variants.sort((a, b) => b.length - a.length);
  return new RegExp("(?<![\\p{L}\\p{N}-])(" + variants.join("|") + ")(?![\\p{L}\\p{N}-])", "gu");
})();
function textRuns(text, fmt, italicizeGlossary) {
  if (!italicizeGlossary || !glossaryRegex || fmt.italics) return [new TextRun({ text, ...fmt })];
  text = typo(text);   // avant de couper : « IFRS 16 » ou « 3 % » ne doivent pas se trouver de part et d'autre d'un terme
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
    alignment: ALIGN,
    ...opts.para,
    children: [new TextRun({ text, ...bodyFont, ...opts.run })],
  });
}
// Un paragraphe tout en gras est un intertitre : il reste avec ce qu'il annonce.
function runs(parts, para = {}) {
  const intertitle = parts.length > 0 && parts.every((x) => x.bold);
  return new Paragraph({ spacing: { after: 0, line: LINE }, alignment: ALIGN,
    ...(intertitle ? { keepNext: true, keepLines: true } : {}), ...para,
    children: parts.map((x) => new TextRun({ ...bodyFont, ...x })) });
}
function blank() {
  return new Paragraph({ spacing: { after: 0, line: LINE }, children: [] });
}
function bullets(items) {
  return items.map((t) => new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 0, line: LINE },
    alignment: ALIGN, children: [new TextRun({ text: t, ...bodyFont })] }));
}
function numbered(items, ref) {
  numberingRefs.add(ref);
  return items.map((t) => new Paragraph({ numbering: { reference: ref, level: 0 }, spacing: { after: 0, line: LINE },
    alignment: ALIGN, children: [new TextRun({ text: t, ...bodyFont })] }));
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
// La légende posée au-dessus de son tableau, et qui ne le quitte pas.
function captionAbove(text) {
  return new Paragraph({ keepNext: true, keepLines: true, spacing: { before: 60, after: 60, line: 240 }, alignment: AlignmentType.LEFT,
    children: [new TextRun({ text, ...smallFont, italics: true })] });
}

// ---- titres : en gras, taille 12, un seul retour avant et après ; chaque titre est noté pour le sommaire.
// heading() rend une marque ; settle() en fait un paragraphe quand sa place est connue : sans espace avant quand il
// ouvre une page (saut porté par le titre lui-même, ou début de section) ou suit directement un autre titre, dont le
// retour suffit ; coupé en lignes équilibrées quand il en prend plusieurs.
const HEADINGS = [];
function heading(text, level, opts = {}) {
  HEADINGS.push({ text, level, annex: !!opts.annex, label: opts.label || "", role: opts.role || 0, toc: opts.toc !== false });
  return { isHeading: true, text, level, newPage: !!opts.newPage };
}
const isHeading = (x) => !!x && x.isHeading === true;
// Les titres que LibreOffice a laissés seuls en bas d'une page du PDF précédent, quand le paragraphe auquel ils sont
// liés part à la page suivante avec ses notes : dataroom report les relève (layout_findings) et le titre ouvre alors
// la page, ce que « lié au suivant » aurait dû faire.
// La mise à plat est celle de dataroom report (_one_line), qui nomme ces titres d'après les signets du PDF : sans
// caractère invisible, un mot composé coupé après son trait d'union recollé, les espaces et les retours réduits à une
// espace.
const oneLine = (s) => String(s || "").replace(/[\u200b\u2060\u00ad]/g, "")
  .replace(/(?<=[-\u2010\u2013\u2014])\r?\n/g, "").replace(/\s+/g, " ").trim();
const PAGE_BREAKS = new Set(((CONTROLS && CONTROLS.layout && CONTROLS.layout.page_breaks) || []).map(oneLine));
const TIGHTEN = new Set(((CONTROLS && CONTROLS.layout && CONTROLS.layout.tighten) || []).map(oneLine));
const tightFor = (title) => (!VEOLIA ? 0 : TIGHTEN.has(oneLine(title) + "!!") ? 2 : TIGHTEN.has(oneLine(title)) ? 1 : 0);
// Édition mise en forme. Niveau 1 : le titre, en blanc, sur un bandeau marine posé en haut de la page (il ouvre toujours
// une page) ; l'anneau du bandeau est le marqueur circulaire de la charte. Niveau 2 : marine, souligné d'un filet bleu ciel.
const BAND_PX = [794, 192], BAND_TWIPS = 2880, H1_SIZE = 32, H1_LINE = 420, H1_INDENT = 2100;
const floating = (file, width, height, x = 0, y = 0) => new ImageRun({ type: "png", data: fs.readFileSync(path.join(ASSETS, file)),
  transformation: { width, height },
  floating: { horizontalPosition: { relative: HorizontalPositionRelativeFrom.PAGE, offset: x },
    verticalPosition: { relative: VerticalPositionRelativeFrom.PAGE, offset: y },
    behindDocument: true, allowOverlap: true, wrap: { type: TextWrappingType.NONE } } });
function veoliaHeading(h, before, first) {
  if (h.level === 1) {
    const lines = balancedLines(h.text, TEXT_WIDTH - H1_INDENT, H1_SIZE, "b");
    const n = lines ? lines.length : 1;
    return new Paragraph({ heading: HEADING_LEVELS[0], keepNext: true, keepLines: true, pageBreakBefore: !first,
      spacing: { before: 0, after: Math.max(240, BAND_TWIPS - MARGIN - n * H1_LINE + 320), line: H1_LINE, lineRule: "exact" },
      alignment: AlignmentType.LEFT, indent: { right: H1_INDENT },
      children: [floating("bandeau.png", BAND_PX[0], BAND_PX[1]), ...linedRuns(h.text, lines, { font: FONT, size: H1_SIZE, bold: true, color: "FFFFFF" })] });
  }
  const size = h.level === 2 ? 26 : 22;
  return new Paragraph({ heading: HEADING_LEVELS[h.level - 1], keepNext: true, keepLines: true, pageBreakBefore: h.newPage,
    spacing: { before: before ? (h.level === 2 ? (h.tight === 2 ? 200 : 320) : (h.tight === 2 ? 140 : 220)) : 0,
      after: h.level === 2 ? (h.tight === 2 ? 100 : 160) : (h.tight === 2 ? 50 : 90), line: h.level === 2 ? 320 : 280 }, alignment: AlignmentType.LEFT,
    ...(h.level === 2 ? { border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: SKY, space: 5 } } } : {}),
    children: linedRuns(h.text, balancedLines(h.text, TEXT_WIDTH, size, "b"), { font: FONT, size, bold: true, color: h.level === 2 ? MARINE : INK }) });
}
function headingParagraph(h, before, first) {
  if (VEOLIA) return veoliaHeading(h, before, first);
  return new Paragraph({
    heading: HEADING_LEVELS[h.level - 1], keepNext: true, keepLines: true, pageBreakBefore: h.newPage,
    spacing: { before, after: RETURN, line: LINE }, alignment: AlignmentType.LEFT,
    children: linedRuns(h.text, balancedLines(h.text, TEXT_WIDTH, BODY, "b"), { ...bodyFont, bold: true, color: "000000" }),
  });
}
function settle(children) {
  // un titre laissé seul ouvre la page ; s'il suit d'autres titres, c'est le premier de la pile qui l'ouvre
  const breakAt = new Set();
  children.forEach((x, i) => {
    if (!isHeading(x) || !PAGE_BREAKS.has(oneLine(x.text))) return;
    let j = i;
    while (j > 0 && isHeading(children[j - 1])) j -= 1;
    breakAt.add(j);
  });
  return children.map((x, i) => {
    if (!isHeading(x)) return x;
    const h = breakAt.has(i) && i > 0 ? { ...x, newPage: true } : x;
    return headingParagraph(h, i === 0 || h.newPage || isHeading(children[i - 1]) ? 0 : RETURN, i === 0);
  });
}

// ---- tableaux
// Le texte d'une cellule, protégé : une date (31/12/2025) ou une période (2024-2027) ne se coupe pas ; une adresse ou
// un nom de fichier se coupe après « / », « _ » ou un trait d'union plutôt qu'au milieu d'un mot.
const URLISH = /:\/\/|^www\.|\.(?:pdf|html?|xlsx?|csv|docx?)\b/i;
function protect(text) {
  return String(text == null ? "" : text).split(/( +)/).map((tok) => {
    if (!tok.trim()) return tok;
    if (URLISH.test(tok)) return tok.replace(/([\/_?&=#])(?=[^\/\s])|(-)(?=\d)/g, (m, a, b) => (a || b) + ZWSP);
    return tok.replace(/(\d)([\/-])(?=\d)/g, `$1${WJ}$2${WJ}`);
  }).join("");
}
const protectInlines = (inlines) => (inlines || []).map((r) => (r.kind === "text" ? { ...r, text: protect(r.text) } : r));
// Ce que la cellule imprime, pour mesurer ses mots.
const cellText = (content) => (Array.isArray(content) ? plain(protectInlines(content)) : protect(content));
const cellBold = (content, head) => head || (Array.isArray(content) && content.some((r) => r.bold));
function cell(content, width, opts = {}) {
  const children = Array.isArray(content)
    ? inlineRuns(protectInlines(content), { ...smallFont, bold: !!opts.head, ...(VEOLIA && opts.head ? { color: "FFFFFF" } : {}) }, { glossary: false })
    : [new TextRun({ text: protect(content), ...smallFont, bold: !!opts.head, ...(VEOLIA && opts.head ? { color: "FFFFFF" } : {}) })];
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    shading: opts.head ? { type: ShadingType.CLEAR, color: "auto", fill: VEOLIA ? MARINE : "E7E6E6" } : undefined,
    margins: { top: VEOLIA ? 70 : 40, bottom: VEOLIA ? 70 : 40, left: 80, right: 80 },
    children: [new Paragraph({ keepNext: !!opts.keepNext, spacing: { after: 0, line: 240 },
      alignment: opts.right ? AlignmentType.RIGHT : AlignmentType.LEFT, children })],
  });
}
// Chaque colonne reçoit au moins la largeur de son mot insécable le plus long, marges de la cellule comprises,
// prise sur les colonnes qui en ont de reste : aucun mot ne se coupe en deux dans une cellule.
function fitWidths(rows, base) {
  const need = base.map(() => 0);
  rows.forEach((row, i) => row.forEach((c, j) => {
    const style = cellBold(c, i === 0) ? "b" : "r";
    for (const s of segments(cellText(c))) need[j] = Math.max(need[j], textWidth(s.text, SMALL, style));
  }));
  const min = need.map((w) => Math.ceil(w + CELL_PAD + 40));
  const short = base.map((w, j) => Math.max(0, min[j] - w));
  const deficit = short.reduce((a, b) => a + b, 0);
  if (!deficit) return base.slice();
  const slack = base.map((w, j) => (short[j] ? 0 : Math.max(0, w - min[j])));
  const totalSlack = slack.reduce((a, b) => a + b, 0);
  if (totalSlack < deficit) {
    console.error(`tableau : ${Math.round(deficit - totalSlack)} twips manquent pour que chaque mot tienne dans sa colonne`);
    const scale = TEXT_WIDTH / min.reduce((a, b) => a + b, 0);
    const out = min.map((w) => Math.floor(w * scale));
    out[out.length - 1] += TEXT_WIDTH - out.reduce((a, b) => a + b, 0);
    return out;
  }
  const out = base.map((w, j) => (short[j] ? min[j] : Math.floor(w - (slack[j] * deficit) / totalSlack)));
  const widest = out.indexOf(Math.max(...out.filter((w, j) => !short[j])));
  out[widest] += TEXT_WIDTH - out.reduce((a, b) => a + b, 0);
  return out;
}
// Un petit tableau (au plus 12 lignes et environ un tiers de page, notes de bas de page qu'il appelle comprises)
// reste d'un seul tenant : chaque ligne, sauf la dernière, est liée à la suivante. Les plus grands se coupent entre
// deux lignes, jamais au milieu d'une, et leurs dernières lignes restent ensemble : au moins deux, et assez pour
// faire TABLE_TAIL_LINES lignes de texte (quatre lignes de tableau au plus). La dernière page d'un long tableau ne
// porte donc jamais une ligne seule sous l'en-tête répété, et un tableau aux lignes hautes ne laisse pas un grand
// blanc au bas de la page qui précède.
const SMALL_TABLE_ROWS = 12, SMALL_TABLE_HEIGHT = 4700;   // twips : un peu plus du tiers de la hauteur utile
const TABLE_TAIL_ROWS = 4, TABLE_TAIL_LINES = 6;
const rowLines = (row, widths, head) => Math.max(1, ...row.map((c, j) =>
  layoutLines(segments(cellText(c)), widths[j] - CELL_PAD, SMALL, cellBold(c, head) ? "b" : "r").length));
function tableHeight(rows, widths) {
  const notes = new Set();
  for (const row of rows) for (const c of row) for (const r of (Array.isArray(c) ? c : [])) if (r.kind === "token" && r.status === "ok") notes.add(noteKey(r));
  return notes.size * 320 + rows.reduce((total, row, i) => total + 80 + 230 * rowLines(row, widths, i === 0), 0);
}
const isSmallTable = (rows, widths) => rows.length <= SMALL_TABLE_ROWS && tableHeight(rows, widths) <= SMALL_TABLE_HEIGHT;
// Combien de lignes de la fin restent ensemble (l'en-tête, rows[0], n'en est pas une).
function tableTail(rows, widths) {
  let kept = 0, lines = 0;
  while (kept < Math.min(TABLE_TAIL_ROWS, rows.length - 1) && (kept < 2 || lines < TABLE_TAIL_LINES)) {
    kept += 1;
    lines += rowLines(rows[rows.length - kept], widths, false);
  }
  return kept;
}
function tableRows(rows, widths, rightCols = []) {
  const small = isSmallTable(rows, widths);
  const tail = small ? 0 : tableTail(rows, widths);
  return rows.map((row, i) => new TableRow({ ...(i === 0 ? { tableHeader: true } : {}), cantSplit: true,
    children: row.map((c, j) => cell(c, widths[j], { head: i === 0, right: i > 0 && rightCols.includes(j),
      keepNext: i === 0 || (i < rows.length - 1 && (small || i >= rows.length - tail)) })) }));
}
// Édition mise en forme : des filets horizontaux seulement, marine en haut et en bas, gris entre les lignes.
const NO_LINE = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const VEOLIA_BORDERS = VEOLIA ? { borders: { top: { style: BorderStyle.SINGLE, size: 8, color: MARINE }, bottom: { style: BorderStyle.SINGLE, size: 8, color: MARINE },
  left: NO_LINE, right: NO_LINE, insideVertical: NO_LINE, insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: RULE } } } : {};
function table(headers, rows, widths, rightCols = []) {
  const total = widths.reduce((a, b) => a + b, 0);
  if (total !== TEXT_WIDTH) throw new Error("largeurs de colonnes : " + total + " ≠ " + TEXT_WIDTH);
  const all = [headers, ...rows];
  const fitted = fitWidths(all, widths);
  return new Table({ width: { size: TEXT_WIDTH, type: WidthType.DXA }, columnWidths: fitted, ...VEOLIA_BORDERS, rows: tableRows(all, fitted, rightCols) });
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
// Deux jetons voisins (séparés par rien ou par des espaces) n'appellent qu'une note, qui les source tous deux ; un
// jeton qui n'affiche rien (une cible sans prix publié) colle son appel au mot qui le précède.
const shown = (r) => (r.status === "ok" ? (r.display || `${r.value}${r.unit ? " " + r.unit : ""}`) : `[${r.id} ?]`);
function joinTokens(inlines) {
  const out = [];
  const list = inlines || [];
  for (let i = 0; i < list.length; i++) {
    const r = list[i];
    if (r.kind !== "token") { out.push(r); continue; }
    const group = { kind: "tokens", items: [r], seps: [] };
    let j = i + 1;
    for (;;) {
      if (j < list.length && list[j].kind === "token") { group.seps.push(""); group.items.push(list[j]); j += 1; continue; }
      if (j + 1 < list.length && list[j].kind === "text" && !list[j].text.trim() && list[j + 1].kind === "token") {
        group.seps.push(list[j].text); group.items.push(list[j + 1]); j += 2; continue;
      }
      break;
    }
    i = j - 1;
    const prev = out[out.length - 1];
    if (!shown(r).trim() && prev && prev.kind === "text") out[out.length - 1] = { ...prev, text: prev.text.replace(/\s+$/, "") };
    out.push(group);
  }
  return out;
}
function inlineRuns(inlines, base = bodyFont, opts = {}) {
  const glossary = opts.glossary !== false;
  const out = [];
  for (const r of joinTokens(inlines)) {
    const fmt = { ...base, bold: !!r.bold || !!base.bold, italics: !!r.italic };
    if (r.kind === "text") out.push(...textRuns(r.text, fmt, glossary));
    else if (r.kind === "code") out.push(new TextRun({ text: r.text, ...fmt, font: "Courier New", verbatim: true }));
    else if (r.kind === "link") out.push(new ExternalHyperlink({ link: r.url, children: [new TextRun({ text: r.text, ...fmt, style: "Hyperlink" })] }));
    // l'adresse d'une source, imprimée comme le texte qui l'entoure : `text` porte déjà ses coupures permises
    else if (r.kind === "address") out.push(new ExternalHyperlink({ link: r.url, children: [new TextRun({ text: r.text, ...fmt, color: "000000", underline: { type: UnderlineType.NONE }, verbatim: true })] }));
    else if (r.kind === "wikilink") out.push(new TextRun({ text: r.text, ...fmt, italics: true }));
    else if (r.kind === "tokens") {
      const notes = [];
      let printed = false;
      r.items.forEach((t, k) => {
        const tfmt = { ...base, bold: !!t.bold || !!base.bold, italics: !!t.italic };
        const text = shown(t);
        if (text.trim() && printed) out.push(new TextRun({ text: r.seps[k - 1] || " ", ...tfmt }));
        if (t.status !== "ok") out.push(new TextRun({ text, ...tfmt, bold: true, color: "C00000", highlight: "yellow" }));
        else {
          if (text.trim()) out.push(new TextRun({ text, ...tfmt }));
          const key = noteKey(t);
          if (!noted.has(key)) {
            noted.add(key);
            notes.push(groupNote(key, t));
          }
        }
        printed = printed || !!text.trim();
      });
      if (notes.length) out.push(footnote(notes.join(" ")));
    }
  }
  return out;
}
function plain(inlines) {
  return (inlines || []).map((r) => (r.kind === "token" ? shown(r) : r.text)).join("").trim();
}
// Les blocs d'un fichier texte. `part` numérote les sous-parties (1.1, 1.1.1) ; sans lui (introduction,
// conclusion, remerciements) les sous-titres restent sans numéro, comme la méthode le permet.
function renderBlocks(blocks, opts = {}) {
  const AFTER = opts.tight === 2 ? 20 : opts.tight ? 50 : PARA_AFTER, LN = opts.tight === 2 ? LINE - 18 : opts.tight ? LINE - 10 : LINE;   // une partie resserrée, d'un cran ou de deux
  const part = opts.part || 0;
  noted = new Set();   // chaque partie redonne la source de ses chiffres
  collectGroups(blocks);
  const counters = [0, 0];
  const out = [];
  // Les tableaux du texte : largeurs ajustées à leurs mots, et s'ils sont petits, d'un seul tenant avec la phrase
  // qui les annonce.
  const layouts = new Map();
  for (const b of blocks || []) {
    if (b.kind !== "table" || !b.rows.length) continue;
    const n = Math.max(...b.rows.map((r) => r.length));
    const rows = b.rows.map((row) => { const r = row.slice(); while (r.length < n) r.push([]); return r; });
    const widths = fitWidths(rows, evenWidths(n));
    layouts.set(b, { rows, widths, small: isSmallTable(rows, widths) });
  }
  let boxed = false;   // édition mise en forme : la liste de ce qui n'a pas pu être établi est encadrée
  (blocks || []).forEach((b, k) => {
    const next = (blocks || [])[k + 1];
    if (b.kind === "heading") boxed = VEOLIA && (b.box === true || /^Ce que nous n.avons pas pu établir/.test(plain(b.inlines)));
    const announces = !!next && layouts.has(next) && layouts.get(next).small;
    if (b.kind === "heading") {
      const level = Math.min(3, Math.max(2, b.level));
      let number = "";
      if (part) {
        if (level === 2) { counters[0] += 1; counters[1] = 0; number = `${part}.${counters[0]} `; }
        else { counters[1] += 1; number = `${part}.${counters[0] || 1}.${counters[1]} `; }
      }
      // dans une annexe, les sous-titres restent hors du sommaire : il ne liste que le corps (§ 5.4)
      out.push({ ...heading(number + plain(b.inlines), level, opts.inAnnex ? { toc: false } : {}), tight: opts.tight || 0 });
    } else if (b.kind === "paragraph") {
      // un paragraphe tout en gras est un intertitre ; la phrase qui annonce un petit tableau reste avec lui
      const intertitle = b.inlines.length > 0 && b.inlines.every((r) => r.bold || (r.kind === "text" && !r.text.trim()));
      out.push(new Paragraph({ spacing: { after: AFTER, line: LN }, alignment: ALIGN,
        ...(intertitle || announces ? { keepNext: true } : {}), ...(intertitle ? { keepLines: true } : {}), children: inlineRuns(b.inlines) }));
    } else if (b.kind === "quote") {
      // Une citation recopiée exactement, entre guillemets français (§ 5.6) ; le retrait la signale.
      const text = plain(b.inlines);
      const open = /^[«"“]/.test(text), close = /[»"”]$/.test(text);
      out.push(new Paragraph({ spacing: { after: AFTER, line: LN }, alignment: ALIGN, indent: { left: 567, right: 567 },
        children: [...(open ? [] : [new TextRun({ text: EN ? "“" : "« ", ...bodyFont })]), ...inlineRuns(b.inlines, bodyFont, { glossary: false }), ...(close ? [] : [new TextRun({ text: EN ? "”" : " »", ...bodyFont })])] }));
    } else if (b.kind === "list") {
      let ref = "bullets";
      if (b.ordered) { ref = `list${dynamicLists++}`; numberingRefs.add(ref); }
      b.items.forEach((item, i) => {
        out.push(new Paragraph({ numbering: { reference: ref, level: 0 }, spacing: { after: i === b.items.length - 1 ? AFTER : 0, line: LN },
          ...(boxed ? { shading: { type: ShadingType.CLEAR, color: "auto", fill: SKY_PALE },
            border: { left: { style: BorderStyle.SINGLE, size: 24, color: MARINE, space: 6 } } } : {}),
          alignment: ALIGN, children: inlineRuns(item) }));
      });
    } else if (b.kind === "todo") {
      if (!FINAL) out.push(placeholder(b.text));
    } else if (b.kind === "table" && b.rows.length) {
      const { rows, widths } = layouts.get(b);
      out.push(new Table({ width: { size: TEXT_WIDTH, type: WidthType.DXA }, columnWidths: widths, ...VEOLIA_BORDERS, rows: tableRows(rows, widths) }));
      out.push(new Paragraph({ spacing: { after: AFTER, line: 240 }, children: [] }));
    }
  });
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
  return `Partie rédigée par ${signed}.`;
}
// Une entrée du sommaire ou de la table des annexes : le titre, des points de conduite, la page. Le titre s'arrête
// à 1 cm de la marge (retrait à droite) et la tabulation va jusqu'à elle : la dernière ligne garde toujours des
// points et son numéro, aucune ligne n'est faite que de points. Une partie (niveau 1) ne finit jamais une page du
// sommaire : elle reste avec sa première sous-partie ; aucune entrée ne se coupe entre deux pages.
function tocLine(text, page, level) {
  const left = level > 1 ? 567 * (level - 1) : 0;
  const fmt = { ...bodyFont, bold: level === 1 };
  return new Paragraph({
    spacing: { after: level === 1 ? 60 : 0, line: 276 }, alignment: AlignmentType.LEFT,
    indent: { left, right: TOC_RIGHT }, keepNext: level === 1, keepLines: true,
    tabStops: [{ type: TabStopType.RIGHT, position: TEXT_WIDTH, leader: LeaderType.DOT }],
    children: [...linedRuns(text, balancedLines(text, TEXT_WIDTH - left - TOC_RIGHT, BODY, level === 1 ? "b" : "r"), fmt),
      new TextRun({ children: [new Tab(), page || ""], ...fmt })],
  });
}

// =================================================================== I. Pages liminaires, sans numéro
const front = [];

// ---- page de couverture (§ 5.2) : logo, formation, entreprise, titre, prénoms et NOMS, année, responsable
const center = (children, spacing) => new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 0, line: LINE, ...spacing }, children });
// Le logo de l'école et, s'il est déposé à côté de ce script (logo-veolia.png), celui de l'entreprise : sur une même ligne,
// à la même hauteur, chacun dans ses proportions (lues dans l'en-tête du PNG ; un JPEG garde l'ancien cadre 220 × 80).
const LOGO_HEIGHT = 46;
const logoRun = (file) => {
  const data = fs.readFileSync(file);
  const png = data.length > 24 && data.toString("latin1", 1, 4) === "PNG";
  const transformation = png ? { width: Math.round(LOGO_HEIGHT * data.readUInt32BE(16) / data.readUInt32BE(20)), height: LOGO_HEIGHT } : { width: 220, height: 80 };
  return new ImageRun({ type: /\.jpe?g$/i.test(file) ? "jpg" : "png", data, transformation });
};
const COMPANY_LOGO = path.join(__dirname, "logo-veolia.png");
if (VEOLIA) {
  // Le fond (marine, liseré, anneau, logotypes) est une image ; le texte se pose dessus dans un tableau sans filets, à
  // hauteurs fixes : le haut sur le marine, le bas sur le blanc.
  const W = (text, o = {}) => new TextRun({ text, font: FONT, color: "FFFFFF", ...o });
  const G = (text, o = {}) => new TextRun({ text, font: FONT, color: INK, size: 21, ...o });
  const line = (children, o = {}) => new Paragraph({ alignment: AlignmentType.LEFT, spacing: { after: 0, line: 276, ...(o.spacing || {}) }, ...(o.indent ? { indent: o.indent } : {}), children });
  const blankCell = { borders: { top: NO_LINE, bottom: NO_LINE, left: NO_LINE, right: NO_LINE }, margins: { top: 0, bottom: 0, left: 0, right: 0 } };
  const MONTHS = ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre"];
  const iso = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(DATA.generated || ""));
  const when = iso ? `${Number(iso[3])} ${MONTHS[Number(iso[2]) - 1]} ${iso[1]}` : String(DATA.generated || "");
  const top = [
    line([W(COVER.formation || "EDHEC Business School", { size: 19 })]),
    line([W(DATA.title, { size: 48, bold: true })], { spacing: { before: 760, line: 600, lineRule: "exact" }, indent: { right: 2700 } }),
    line([W(DATA.subtitle, { size: 22 })], { spacing: { before: 360, line: 300 }, indent: { right: 3300 } }),
  ];
  const bottom = [
    line([G(`${T("Rapport réalisé pour")} ${COVER.entreprise || "Veolia Environnement"}`, { bold: true, color: MARINE, size: 24 })]),
    line([G((COVER.auteurs || []).join(", "))], { spacing: { before: 200 } }),
    line([G([COVER.encadrant ? `${T("Responsable")}${SEP}${COVER.encadrant}` : "", T(COVER.annee || "")].filter(Boolean).join(" | "))], { spacing: { before: 80 } }),
    line([G(T("Travail académique d'étudiants de l'EDHEC. Ce document n'est pas une publication de Veolia."), { size: 17, italics: true })], { spacing: { before: 280 } }),
    ...(FINAL ? [] : [line([G(T(`Rapport généré le ${when} depuis la dataroom du groupe : ${DATA.n_figures} chiffres sourcés, contrôles du classeur ${DATA.checks}. Édition vérifiable :`) + " ", { size: 15 }),
      new ExternalHyperlink({ link: WEB, children: [new TextRun({ text: WEB, font: FONT, size: 15, style: "Hyperlink" })] })], { spacing: { before: 100, line: 220 } })]),
  ];
  // le fond s'ancre dans un paragraphe hors du tableau : dans une cellule, LibreOffice le placerait par rapport à elle
  front.push(new Paragraph({ spacing: { after: 0, line: 20, lineRule: "exact" }, children: [floating("couverture.png", 794, 1123)] }));
  front.push(new Table({ width: { size: TEXT_WIDTH, type: WidthType.DXA }, columnWidths: [TEXT_WIDTH],
    borders: { top: NO_LINE, bottom: NO_LINE, left: NO_LINE, right: NO_LINE, insideHorizontal: NO_LINE, insideVertical: NO_LINE },
    rows: [new TableRow({ height: { value: 8830, rule: HeightRule.EXACT }, children: [new TableCell({ ...blankCell, width: { size: TEXT_WIDTH, type: WidthType.DXA }, children: top })] }),
      new TableRow({ height: { value: 4000, rule: HeightRule.EXACT }, children: [new TableCell({ ...blankCell, width: { size: TEXT_WIDTH, type: WidthType.DXA }, verticalAlign: VerticalAlign.TOP, children: bottom })] })] }));
} else {
if (COVER.logo && fs.existsSync(COVER.logo)) {
  const logos = [logoRun(COVER.logo)];
  if (fs.existsSync(COMPANY_LOGO)) logos.push(new TextRun({ text: "\u00A0".repeat(16), ...bodyFont }), logoRun(COMPANY_LOGO));
  front.push(center(logos, { before: 0, after: 240 }));
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
  front.push(center([new TextRun({ text: `${todoCount ? "Squelette" : "Rapport"} généré le ${when} depuis la dataroom du groupe : ${DATA.n_figures} chiffres sourcés, contrôles du classeur ${DATA.checks}.`
    + (todoCount ? " Les passages surlignés sont à rédiger par l'auteur de la partie." : ""), ...smallFont, color: "555555" })], { before: 600, line: 240 }));
  front.push(center([new TextRun({ text: "Édition vérifiable, où chaque chiffre ouvre sa page source : ", ...smallFont, color: "555555" }),
    new ExternalHyperlink({ link: WEB, children: [new TextRun({ text: WEB, ...smallFont, style: "Hyperlink" })] })], { line: 240 }));
}
}
// ---- la page blanche qui suit la couverture (§ 5.3) : un paragraphe vide qui ouvre sa page ; les remerciements,
// la suivante. Chaque saut de page est porté par ce qui ouvre la page, jamais par un paragraphe de saut.
front.push(new Paragraph({ pageBreakBefore: true, spacing: { after: 0, line: LINE }, children: [] }));
// ---- remerciements (§ 5.3) : sobres, écrits par le groupe
front.push(heading(T("Remerciements"), 1, { toc: false, newPage: true }));
const thanks = prose("remerciements");
if (thanks) front.push(...renderBlocks(thanks.blocks));
else if (!FINAL) front.push(placeholder("[À rédiger] 08_Livrables/texte/remerciements.md : des remerciements sobres, à Veolia et aux personnes qui ont éclairé le travail."));

// =================================================================== II. Le corps, numéroté à partir du sommaire
const body = [];
// ---- le sommaire (§ 5.4), page 1 : rempli au second passage, quand le PDF a donné ses pages
const tocMarker = body.length;
body.push(heading(T("Sommaire"), 1, { toc: false }));
// (les entrées sont insérées à la fin, quand tous les titres sont connus)

// ---- introduction (§ 4.3) : rédigée en dernier par le groupe ; situe le cadre, présente la mission, annonce les résultats et le plan
if (VEOLIA && (DATA.tiles || []).length) {
  // Les indicateurs de la réponse (ceux du tableau de la recommandation), en tuiles : un nombre, son nom, sa lecture.
  body.push(heading(T("Chiffres clés"), 1, { toc: false, newPage: true }));
  const COLS = 2, wide = Math.floor(TEXT_WIDTH / COLS), gap = { style: BorderStyle.SINGLE, size: 36, color: "FFFFFF" };
  const tile = (t) => new TableCell({ width: { size: wide, type: WidthType.DXA }, shading: { type: ShadingType.CLEAR, color: "auto", fill: SKY_PALE },
    borders: { top: gap, bottom: gap, left: gap, right: gap }, margins: { top: 220, bottom: 220, left: 260, right: 200 },
    children: t ? [new Paragraph({ alignment: AlignmentType.LEFT, spacing: { after: 60, line: 560, lineRule: "exact" }, children: [new TextRun({ text: t[1], font: FONT, size: 48, bold: true, color: MARINE })] }),
      new Paragraph({ alignment: AlignmentType.LEFT, spacing: { after: 60, line: 260 }, children: [new TextRun({ text: t[0], font: FONT, size: 20, bold: true, color: MARINE })] }),
      new Paragraph({ alignment: AlignmentType.LEFT, spacing: { after: 0, line: 240 }, children: [new TextRun({ text: t[2], font: FONT, size: 17 })] })] : [new Paragraph({ children: [] })] });
  const rows = [];
  for (let i = 0; i < DATA.tiles.length; i += COLS) rows.push(new TableRow({ cantSplit: true, children: Array.from({ length: COLS }, (_, j) => tile(DATA.tiles[i + j])) }));
  body.push(new Table({ width: { size: wide * COLS, type: WidthType.DXA }, columnWidths: Array(COLS).fill(wide),
    borders: { top: NO_LINE, bottom: NO_LINE, left: NO_LINE, right: NO_LINE, insideHorizontal: NO_LINE, insideVertical: NO_LINE }, rows }));
  body.push(caption(T(`Source : classeur du groupe, calculé à partir des chiffres du registre (${DATA.n_figures} lignes sourcées). Le détail figure dans la dernière partie.`)));
}
body.push(heading(T("Introduction"), 1, { newPage: true }));
const intro = prose("introduction");
if (intro) body.push(...renderBlocks(intro.blocks, { tight: tightFor("Introduction") }));
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
  else if (!FINAL && !VEOLIA) body.push(placeholder(`[Partie rédigée par : ______ ; frontmatter signe_par de texte/${key}.md] ${role.tabs ? "Dans le classeur : " + role.tabs : ""}`));
  // Une partie rédigée se lit seule : les aides générées (consigne, chiffres du rôle, questions du jury)
  // ne s'impriment qu'à côté d'une partie encore vide.
  if (!FINAL && !text) body.push(p(role.brief, { run: { italics: true, color: "555555" } }), blank());
  if (text) {
    body.push(...renderBlocks(text.blocks, { part, tight: tightFor(`${part}. ${title}`) }));
  } else if (!FINAL) {
    body.push(placeholder(`[À rédiger] 08_Livrables/texte/${key}.md : 6 à 9 pages, chaque chiffre cité par son jeton [F…] du registre ; des titres habillés (ce que la partie démontre), jamais « analyse » ou « résultats ».`));
    body.push(runs([{ text: "Réponse générée depuis le classeur, à réécrire", bold: true }]), p(role.answer), blank());
    body.push(runs([{ text: "Raisonnement généré", bold: true }]), ...numbered(role.reasoning, `role${role.n}`), blank());
  }
  if (!FINAL && !text) {
    body.push(runs([{ text: "Chiffres du registre pour cette partie", bold: true }]));
    body.push(table(["ID", "Chiffre", "Période", "Valeur", "Source (fichier, page)"],
      role.figures.map((f) => [f[0], f[1], f[2], f[3], f[4]]), [900, 3200, 1100, 1400, 2472], [3]));
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
if (reco) body.push(...renderBlocks(reco.blocks, { part, tight: tightFor(`${part}. ${recoTitle}`) }));
else if (!FINAL) body.push(placeholder("[À rédiger] 08_Livrables/texte/recommandation.md : ce que nous recommandons, ce que cela coûte et à qui, ce qui nous ferait changer d'avis. Chiffres par jeton."));
body.push(captionAbove(T(`Tableau ${part}.1 : Les indicateurs de la réponse, calculés dans le classeur du groupe à partir des chiffres du registre (${DATA.n_figures} lignes sourcées).`)));
body.push(table(["Indicateur", "Valeur", "Ce qu'il faut en lire"].map(T), DATA.tiles.map((t) => [t[0], t[1], t[2]]), [3200, 1600, 4272], [1]));

// ---- conclusion (§ 4.5) : les points forts à retenir
body.push(heading(T("Conclusion"), 1, { newPage: true }));
const concl = prose("conclusion");
if (concl) body.push(...renderBlocks(concl.blocks, { tight: tightFor("Conclusion") }));
else if (!FINAL) {
  body.push(placeholder("[À rédiger] 08_Livrables/texte/conclusion.md : les points forts à retenir, la réponse en trois phrases, ce que le groupe retient de la méthode."));
  body.push(runs([{ text: "Ce qu'il reste à faire avant la remise", bold: true }]), ...bullets(DATA.todo));
}

// ---- bibliographie raisonnée (§ 4.6, § 5.7.3) : les documents cités, par type de source, puis par auteur
// Une référence : « Auteur (année), « Titre », site, adresse, consulté le J mois AAAA. » L'adresse est un lien, imprimé
// comme le texte ; elle peut se couper après / _ ? & = # et, après un trait d'union, devant un chiffre (une espace sans
// chasse, que LibreOffice retire du PDF). « consulté le J mois AAAA. » reste d'un seul tenant : plus de « 2026. » seul
// sur une ligne. Une référence d'une autre forme s'imprime telle quelle.
// « consulté le 6 novembre 2026 » → « accessed on 6 November 2026 »
const EN_MONTHS = { janvier: "January", "février": "February", mars: "March", avril: "April", mai: "May", juin: "June", juillet: "July", "août": "August",
  septembre: "September", octobre: "October", novembre: "November", "décembre": "December" };
const enAccessed = (s) => String(s).replace(/consulté le (\d{1,2})(?:er)? (\p{L}+) (\d{4})/gu, (_, d, mo, y) => `accessed on ${d} ${EN_MONTHS[mo] || mo} ${y}`)
  .replace(/consulté le /g, "accessed on ");
const enQuotes = (s) => String(s).replace(/«[\u00a0\u202f ]?/g, "“").replace(/[\u00a0\u202f ]?»/g, "”");
function biblioRuns(entry) {
  const m = /^(.*?, )(https?:\/\/\S+?)(, )(consulté le .*)$/.exec(entry);
  if (!m) return [new TextRun({ text: EN ? enQuotes(enAccessed(entry)) : entry, ...bodyFont })];
  const [, headFr, url, comma, consultedFr] = m;
  const consulted = EN ? enAccessed(consultedFr) : consultedFr, head = EN ? enQuotes(headFr) : headFr;
  const shown = url.replace(/([\/_?&=#])(?=[^\/\s])|(-)(?=\d)/g, (x, a, b) => (a || b) + ZWSP);
  return [
    new TextRun({ text: head, ...bodyFont }),
    new ExternalHyperlink({ link: url, children: [new TextRun({ text: shown, ...bodyFont, color: "000000", underline: { type: UnderlineType.NONE }, verbatim: true })] }),
    new TextRun({ text: comma + consulted.replace(/ /g, NBSP), ...bodyFont }),
  ];
}
body.push(heading(T("Bibliographie"), 1, { newPage: true }));
const biblio = DATA.bibliography || [];
if (biblio.length) {
  body.push(p(T("Les documents cités dans ce rapport, classés par type de source puis par ordre alphabétique d'auteur. Chaque référence donne la date de consultation ; les pages citées figurent dans les notes de bas de page et à l'annexe I.")));
  for (const group of biblio) {
    // la rubrique, liée à sa première référence ; le retour qui la précède est un espace avant, qui tombe en haut de page
    body.push(runs([{ text: group.category, bold: true }], { spacing: { before: RETURN, after: 0, line: LINE } }));
    for (const entry of group.entries) {
      body.push(new Paragraph({ keepLines: true, spacing: { after: 0, line: LINE }, alignment: AlignmentType.LEFT, indent: { left: 567, hanging: 567 },
        children: biblioRuns(entry) }));
    }
  }
} else if (!FINAL) {
  body.push(placeholder("La bibliographie se remplit d'elle-même : chaque document dont une ligne du registre est citée par jeton y prend sa place, par type de source puis par auteur."));
}

// ---- glossaire (§ 4.7) : les termes techniques et sigles, en italique dans le texte, définis ici par ordre alphabétique
body.push(heading(T("Glossaire"), 1, { newPage: true }));
if (GLOSSARY.length) {
  body.push(p(T("Les termes techniques, expressions étrangères et sigles employés dans le rapport ; ils apparaissent en italique dans le texte.")), blank());
  for (const { term, definition } of GLOSSARY) {
    body.push(new Paragraph({ spacing: { after: 0, line: LINE }, alignment: ALIGN, indent: { left: 567, hanging: 567 },
      children: [new TextRun({ text: term, ...bodyFont, bold: true }), new TextRun({ text: `${SEP}${definition}`, ...bodyFont })] }));
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
  const label = `${T("Annexe")} ${ROMAN[annexNo++] || annexNo}`;
  const text = `${label}${SEP}${T(title)}`;
  annexes.push(heading(text, 1, { annex: true, label, newPage: annexes.length > 0 }));
  annexes.push(caption(T(legend)));
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
// Les lignes sur lesquelles repose une conversion imprimée (le taux qui convertit une cible citée) : le lecteur les
// tient pour citées, l'annexe I les reprend avec les autres.
(DATA.cited_extra || []).forEach((id) => citedIds.add(String(id).split("-")[0]));
const citedRows = DATA.register.filter((r) => citedIds.has(String(r[0])));
// Une valeur trop longue pour sa colonne se coupe avant « par an », jamais entre « par » et « an ».
const valueCell = (v) => String(v == null ? "" : v).replace(/ par (an|action)$/, ` par${NBSP}$1`);
// La source d'une ligne d'annexe : la référence courte de la bibliographie (ou l'adresse d'une ligne sans fichier
// joint), à défaut la citation entière du registre. Une adresse est un lien, comme en bibliographie : sa cible est
// l'adresse telle quelle, celle que l'on suit ou que l'on copie, sans les espaces sans chasse qui la laissent se
// couper. Ce qui s'imprime ne change pas d'un caractère : le texte protégé de la cellule (protect), dont l'adresse
// est la part cliquable.
function sourceCell(id, citation) {
  const ref = (DATA.refs || {})[id] || citation || "-";
  const m = /^(https?:\/\/\S+?)(, consulté le .*)?$/.exec(ref);
  if (!m) return EN ? enAccessed(ref) : ref;
  const tail = m[2] || "", printed = protect(ref);
  return [{ kind: "address", url: m[1], text: printed.slice(0, printed.length - protect(tail).length) },
    { kind: "text", text: tail }].filter((r) => r.text);
}
// Le montant d'une opération ou d'une cible tient en une cellule, devise comprise ; un générateur plus ancien
// ne la fournit pas : la valeur et la devise du registre, telles quelles.
// Une mention sans chiffre (« non communiqué ») tient sur une ligne : la colonne s'élargit pour elle.
const amountCell = (shown, value, currency) => {
  const text = String((shown !== undefined && shown !== null ? shown : [value, currency].filter(Boolean).join(" ")) || "-");
  return /\d/.test(text) ? text : text.replace(/ /g, NBSP);
};
annex("Registre des chiffres cités", [
  ...(citedRows.length
    ? [table(["ID", "Chiffre", "Valeur", "Période", "Source", "Rôle"].map(T),
        citedRows.map((r) => [r[0], r[1], valueCell((DATA.displays || {})[r[0]] || r[2]), r[3], sourceCell(r[0], r[4]), r[5]]),
        [600, 3100, 1200, 1000, 2572, 600], [2])]
    : [FINAL ? p("Aucun chiffre cité.") : placeholder(`Aucun chiffre encore cité par jeton : l'annexe se remplira des lignes citées parmi les ${DATA.register.length} du registre.`)]),
], `${citedRows.length} chiffres cités dans le texte, sur ${DATA.register.length} au registre du groupe. Chaque source renvoie à la bibliographie, qui donne le titre, l'adresse et la date de consultation ; le registre garde le fichier joint.`);
if (DATA.cited_deals && DATA.cited_deals.length) {
  // [id, opération, montant, devise, citation, relecture, montant affiché]
  annex("Opérations citées", [
    table(["ID", "Opération", "Montant", "Source"].map(T),
      DATA.cited_deals.map((d) => [d[0], d[1], amountCell(d[6], d[2], d[3]), sourceCell(d[0], d[4])]), [600, 3700, 1500, 3272], [2]),
  ], "Les opérations du registre deals.csv citées dans le texte par leur jeton [D…]. Le montant est donné dans la devise de sa source ; un tiret : le registre n'en porte pas. Chaque source renvoie à la bibliographie.");
}
if (DATA.cited_targets && DATA.cited_targets.length) {
  // [id, nom, valeur estimée, devise, citation, relecture, en euros, taux, valeur affichée] : « en euros » et « taux »
  // quand le classeur convertit
  const euros = (t) => (t[6] ? `${t[6]}${t[7] ? ` (${t[7]})` : ""}` : "-");
  annex("Cibles citées", [
    table(["ID", "Cible", "Valeur estimée", "En euros (taux, date)", "Source"].map(T),
      DATA.cited_targets.map((t) => [t[0], t[1], amountCell(t[8], t[2], t[3]), euros(t), sourceCell(t[0], t[4])]), [600, 2000, 1400, 2100, 2972], [2]),
  ], "Les cibles du registre targets.csv citées dans le texte par leur jeton [T…]. « Valeur estimée » : dans la devise de sa source ; un tiret quand le registre n'en porte pas. « En euros » : la valeur convertie par le classeur du groupe (onglet Cibles), avec le taux employé, qui est une ligne datée du registre ou l'hypothèse de change du groupe ; un tiret quand le classeur ne convertit pas (prix en euros, non publié ou sans taux). Chaque source renvoie à la bibliographie ; une adresse seule : la ligne n'a pas de fichier joint au registre.");
}
annex("La fourchette de capacité pour les autres périmètres", [
  p(T("Le nombre que les groupes « technologies de l'eau », « déchets dangereux » et « bioénergie » doivent respecter, dans ses trois scénarios et sous ses deux plafonds.")), blank(),
  table(["Fin 2027", "Défavorable", "Central", "Favorable"].map(T), DATA.range_rows, [4272, 1600, 1600, 1600], [1, 2, 3]),
], EN ? `Source: group workbook, modele-greenup-2027.xlsx${DATA.model_version ? ", " + DATA.model_version : ""}, calculated from the figures of the register.`
  : `Source : classeur du groupe, modele-greenup-2027.xlsx${DATA.model_version ? ", " + DATA.model_version : ""}, calculé à partir des chiffres du registre.`);
// ---- la table des annexes (§ 5.8), dernière page numérotée du corps
body.push(heading(T("Table des annexes"), 1, { newPage: true }));
for (const h of HEADINGS.filter((x) => x.annex)) body.push(tocLine(h.text, ANNEX_PAGES.get(h.text) || "", 1));

// ---- le sommaire, maintenant que tous les titres sont connus : parties et sous-parties, puis les annexes en un mot
const tocEntries = HEADINGS.filter((h) => h.toc && !h.annex && h.level <= 2).map((h) => tocLine(h.text, TOC_PAGES.get(h.text) || "", h.level));
const firstAnnex = HEADINGS.find((h) => h.annex);
if (firstAnnex) tocEntries.push(tocLine(T("Annexes (paginées à part, en chiffres romains)"), ANNEX_PAGES.get(firstAnnex.text) || "", 1));
body.splice(tocMarker + 1, 0, ...tocEntries);

// =================================================================== IV. Contrôle des citations (jamais dans la version finale)
const controls = [];
if (CONTROLS && !FINAL) {
  if (!VEOLIA) controls.push(heading("Contrôle des citations et de la méthode (retiré de la version finale)", 1, { toc: false }));
  const files = Object.keys(CONTROLS.words || {}).filter((f) => f !== "outils-ia.md");
  if (files.length) {
    // l'édition longue est la version de référence : les cibles de pages ne valent que pour la courte
    const longEdition = DATA.edition === "longue";
    const [lo, hi] = CONTROLS.pages_per_role || [6, 9];
    controls.push(table(["Fichier", "Pages (PDF)", "Mots", "Caractères", "Cible pages", "À rédiger"],
      files.map((f) => [f, CONTROLS.pages && CONTROLS.pages[f] != null ? CONTROLS.pages[f] : "-", CONTROLS.words[f],
        CONTROLS.chars && CONTROLS.chars[f] != null ? CONTROLS.chars[f] : "-",
        f.startsWith("role-") && !longEdition ? `${lo}-${hi}` : "-", CONTROLS.todos[f] || 0]), [2772, 1200, 1100, 1300, 1300, 1400], [1, 2, 3, 5]));
    const target = (CONTROLS.pages_total_target || [50, 60]).join("-");
    controls.push(caption(`Corps du rapport, du sommaire à la première annexe : ${CONTROLS.pages_body || "-"} pages ${longEdition ? "(édition de référence, sans cible de pages)" : `(cible ${target})`} ; document entier : ${CONTROLS.pages_total || "-"} pages ; ${CONTROLS.chars_total || 0} caractères espaces compris dans le corps (la méthode en demande ${(DATA.method && DATA.method.chars_min) || 50000} au minimum).`));
  }
  const list = (title, items, fmt) => {
    controls.push(runs([{ text: title, bold: true }]));
    controls.push(...(items.length ? bullets(items.map(fmt)) : [p("aucun", { run: { italics: true } })]), blank());
  };
  list(`Ce que la méthode du rapport ferait souligner (${(CONTROLS.warnings || []).length})`, CONTROLS.warnings || [], (x) => x);
  list(`Jetons inconnus (${CONTROLS.unknown.length})`, CONTROLS.unknown, (x) => `${x[0]} : [${x[1]}] ne correspond à aucune ligne des registres`);
  list(`Lignes citées contestées (${CONTROLS.disputed.length})`, CONTROLS.disputed, (x) => `${x[0]} : ${x[1]} ; ${x[2]}`);
  list("Fichiers absents ou vides", [...(CONTROLS.missing || []), ...(CONTROLS.empty || [])].filter((x) => x !== "outils-ia.md"), (x) => x);
  list("Pages liminaires à écrire", [...(CONTROLS.cover_missing || []).map((x) => `couverture : ${x}`), ...(CONTROLS.front_missing || [])], (x) => x);
  const uncited = Object.entries(CONTROLS.uncited || {});
  list("Chiffres du registre jamais cités par leur rôle", uncited, ([role, ids]) => `rôle ${role} : ${ids.length}${ids.length ? " : " + ids.slice(0, 20).join(", ") + (ids.length > 20 ? " …" : "") : ""}`);
  // ce que dataroom report a relevé sur le PDF (layout_findings) : fins de partie, titres, tableaux, sommaire, notes
  const layout = CONTROLS.layout || {};
  if (layout.findings) {
    list(`Mise en page, relevée sur le PDF (${layout.findings.length})`, layout.findings, (f) => `${f.label} : ${f.text}`);
    if ((layout.page_breaks || []).length) list("Titres qui ouvrent une page (LibreOffice les laissait seuls en bas de la précédente)", layout.page_breaks, (t) => `« ${t} »`);
    const sizes = (layout.page_number_pt || []).map((x) => `${String(x).replace(".", ",")} pt`).join(", ");
    controls.push(p(`Polices du PDF : ${(layout.fonts || []).join(", ") || "-"} ; numéro de page : ${sizes || "-"}.`, { run: { italics: true } }));
  }
}

// =================================================================== le document
const pageProps = (pageNumbers) => ({ page: { size: { width: PAGE_W, height: PAGE_H }, margin: { top: MARGIN, right: MARGIN_RIGHT, bottom: MARGIN, left: MARGIN_LEFT }, ...(pageNumbers ? { pageNumbers } : {}) } });
// Le pied de page : 10 pt posés sur le paragraphe (style Footer et marque de paragraphe), que le champ PAGE reprend
// (LibreOffice ignore la taille portée par le run du champ).
const footerParagraph = (children) => new Paragraph({ style: "Footer", alignment: AlignmentType.CENTER, spacing: { after: 0 }, run: smallFont, children });
const numberFooter = () => (VEOLIA
  ? new Footer({ children: [new Paragraph({ style: "Footer", alignment: AlignmentType.LEFT, spacing: { after: 0 }, run: smallFont,
      tabStops: [{ type: TabStopType.RIGHT, position: TEXT_WIDTH }], border: { top: { style: BorderStyle.SINGLE, size: 4, color: RULE, space: 6 } },
      children: [new TextRun({ text: T("Rapport d'étudiants de l'EDHEC réalisé pour Veolia. Ce document n'est pas une publication de Veolia."), font: FONT, size: 14, color: "8A8A8E" }),
        new TextRun({ children: [new Tab(), PageNumber.CURRENT], font: FONT, size: SMALL, bold: true, color: MARINE })] })] })
  : new Footer({ children: [footerParagraph([new TextRun({ children: [PageNumber.CURRENT], ...smallFont })])] }));
const textFooter = (text) => new Footer({ children: [footerParagraph([new TextRun({ text, ...smallFont, color: "777777" })])] });
const headingStyle = (id, name, level) => ({
  id, name, basedOn: "Normal", next: "Normal", quickFormat: true,
  run: { font: FONT, size: BODY, bold: true, color: VEOLIA ? MARINE : "000000", italics: false },
  paragraph: { spacing: { before: RETURN, after: RETURN, line: LINE }, outlineLevel: level, keepNext: true },
});

const sections = [
  { properties: pageProps(null), footers: { default: textFooter("") }, children: settle(front) },
  { properties: pageProps({ start: 1, formatType: NumberFormat.DECIMAL }), footers: { default: numberFooter() }, children: settle(body) },
  { properties: pageProps({ start: 1, formatType: NumberFormat.UPPER_ROMAN }), footers: { default: numberFooter() }, children: settle(annexes) },
];
if (controls.length && !VEOLIA) sections.push({ properties: pageProps(null), footers: { default: textFooter("Contrôle, hors version finale") }, children: settle(controls) });

const doc = new Document({
  creator: "Groupe sujet 2",
  title: DATA.title,
  styles: {
    // le contrôle des veuves et orphelines, explicite : sans lui, ni Word ni LibreOffice ne l'appliquent à un .docx
    default: { document: { run: bodyFont, paragraph: { spacing: { line: LINE, after: 0 }, alignment: ALIGN, widowControl: true } } },
    paragraphStyles: [
      { id: "Footer", name: "footer", basedOn: "Normal", run: { font: FONT, size: SMALL }, paragraph: { spacing: { line: 240, after: 0 }, alignment: AlignmentType.CENTER } },
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
  if (EN) fs.writeFileSync(OUT.replace(/\.docx$/i, "") + ".i18n.json", JSON.stringify([...I18N_MISSING], null, 1));
  fs.writeFileSync(HEADINGS_OUT, JSON.stringify(HEADINGS.map(({ text, level, annex, label, role }) => ({ text, level, annex, label, role })), null, 1));
  console.log("écrit", OUT, Math.round(buf.length / 1024), "Ko", `(${HEADINGS.length} titres, ${footnoteId} notes)` + (FINAL ? " (version finale)" : ""));
});
