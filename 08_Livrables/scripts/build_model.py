"""Construit modele-greenup-2027.xlsx à partir du registre de la dataroom.

Chaque nombre bleu de l'onglet Entrées est une ligne du registre (fichier, page, URL).
Les cellules jaunes sont les hypothèses du groupe. Tout le reste est formule.
"""
import csv
import os
import re
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L

sys.path.insert(0, str(Path(__file__).parent))
import model_def as MD  # noqa: E402 - la définition du modèle, à côté du script

HERE = Path(__file__).parent
# Où sont les registres et où écrire : la chaîne (dataroom livrables) passe DATAROOM_ADMIN et DATAROOM_LIVRABLES ;
# à défaut, le vault autour du script (08_Livrables/scripts → ../../00_Admin), puis une copie à côté du script.
ADMIN = Path(os.environ.get("DATAROOM_ADMIN") or (HERE.parent.parent / "00_Admin"))
LIV = Path(os.environ.get("DATAROOM_LIVRABLES") or HERE.parent)
def admin_file(name):
    return ADMIN / name if (ADMIN / name).exists() else HERE / name
REG = admin_file("register.csv")
OUT = Path(sys.argv[1]) if len(sys.argv) > 1 else LIV / "modele-greenup-2027.xlsx"

# ---------------------------------------------------------------- styles
F = "Arial"
BLUE, BLACK, GREEN, GREY = "0000FF", "000000", "008000", "808080"
YELLOW = PatternFill("solid", fgColor="FFFF00")
HEAD = PatternFill("solid", fgColor="D9D9D9")
BAND = PatternFill("solid", fgColor="F2F2F2")
CHANGED = PatternFill("solid", fgColor="FCE4D6")
OKF = PatternFill("solid", fgColor="E2EFDA")
KOF = PatternFill("solid", fgColor="F8CBAD")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
TOPLINE = Border(top=Side(style="thin", color="000000"))

NF_M = '#,##0;(#,##0);"-"'
NF_M1 = '#,##0.0;(#,##0.0);"-"'
NF_D2 = '#,##0.00;(#,##0.00);"-"'
NF_X = '0.00"x";(0.00"x");"-"'
NF_P = '0.0%;(0.0%);"-"'
NF_I = '#,##0'


def font(color=BLACK, bold=False, size=10, italic=False):
    return Font(name=F, color=color, bold=bold, size=size, italic=italic)


def put(ws, ref, value, color=BLACK, bold=False, nf=None, fill=None, wrap=False, italic=False,
        size=10, align=None, border=None):
    c = ws[ref]
    c.value = value
    c.font = font(color, bold, size, italic)
    if nf:
        c.number_format = nf
    if fill:
        c.fill = fill
    if wrap or align:
        c.alignment = Alignment(wrap_text=wrap, vertical="top", horizontal=align)
    if border:
        c.border = border
    return c


def title(ws, text, sub=None):
    put(ws, "A1", text, bold=True, size=14)
    if sub:
        put(ws, "A2", sub, color=GREY, italic=True)


def header(ws, row, labels, start=1):
    for i, h in enumerate(labels):
        c = put(ws, f"{L(start + i)}{row}", h, bold=True, fill=HEAD, wrap=True)
        c.border = BOX


def widths(ws, spec):
    for col, w in spec.items():
        ws.column_dimensions[col].width = w


def q(sheet):
    return f"'{sheet}'"


# ---------------------------------------------------------------- registre
def parse(value: str) -> float:
    v = value.replace(" ", "").replace(" ", "").replace(" ", "").strip()
    if re.fullmatch(r"[-+]?\d{1,3}(,\d{3})+(\.\d+)?", v):
        return float(v.replace(",", ""))
    if re.fullmatch(r"-?\d+,\d+", v):
        return float(v.replace(",", "."))
    try:
        return float(v)
    except ValueError:
        return value.strip()  # libellé publié sans nombre (« high teens »)


rows = list(csv.DictReader(REG.open(encoding="utf-8")))
# Les opérations (00_Admin/deals.csv) : même règle que le registre, chaque ligne a sa source et sa page.
DEALS_CSV = admin_file("deals.csv")
DEALS = list(csv.DictReader(DEALS_CSV.open(encoding="utf-8"))) if DEALS_CSV.exists() else []
ENTREE_ROW = {}          # id -> ligne dans Entrées
FIRST = 6


def find(prefix, period, unit=None):
    hits = [r for r in rows if r["label"].startswith(prefix) and r["period"] == period
            and (unit is None or r["unit"] == unit)]
    if len(hits) != 1:
        raise SystemExit(f"{len(hits)} lignes pour {prefix!r} / {period!r} / {unit!r}: "
                         + ", ".join(h["id"] for h in hits))
    return hits[0]["id"]


USED: list[dict] = []   # chaque chiffre que le classeur lit, pour prioriser leur relecture


def E(prefix, period, unit=None):
    """Adresse absolue de la valeur d'une ligne du registre dans Entrées."""
    fid = find(prefix, period, unit)
    USED.append({"id": fid, "prefix": prefix, "period": period, "unit": unit or ""})
    return f"{q('Entrées')}!$C${ENTREE_ROW[fid]}", fid


wb = Workbook()
ws0 = wb.active
ws0.title = "Lisez-moi"
sheets = {}
for name in ["Entrées", "Hypothèses", "Levier", "Pont de dette", "Échéancier", "Cessions", "Trajectoire",
             "Sensibilité", "Distribution", "Segments", "Pont EBITDA", "Booster", "ESG", "Cibles", "Financement", "Vérifications", "Simulation"]:
    sheets[name] = wb.create_sheet(name)

# ================================================================ Entrées
ws = sheets["Entrées"]
title(ws, "Entrées — le registre de la dataroom, tel quel",
      "Une ligne par chiffre publié : valeur, unité, période, fichier, page, URL. "
      "Aucun nombre n'entre dans le modèle par un autre chemin.")
put(ws, "A3", "Valeur = le nombre imprimé, dans l'unité imprimée (les conversions se font dans les formules). "
    "« Tel qu'imprimé » garde la graphie d'origine. Source : 00_Admin/register.csv.", color=GREY, italic=True)
header(ws, 5, ["ID", "Libellé", "Valeur", "Unité", "Période", "Tel qu'imprimé", "Fichier", "Page",
               "Rôle", "URL source", "Consulté le", "Relu par"])
ROLE_NAME = {"1": "1 Périmètre", "2": "2 Contrainte", "3": "3 Capacité", "4": "4 Écart",
             "5": "5 ESG", "6": "6 Cibles"}
for i, r in enumerate(rows):
    n = FIRST + i
    ENTREE_ROW[r["id"]] = n
    val = parse(r["value"])
    if isinstance(val, str):
        put(ws, f"A{n}", r["id"], bold=True); put(ws, f"B{n}", r["label"])
        put(ws, f"C{n}", val, color=BLUE, align="right")
    else:
        nf = NF_I if float(val).is_integer() else '#,##0.00;-#,##0.00'
        put(ws, f"A{n}", r["id"], bold=True)
        put(ws, f"B{n}", r["label"])
        put(ws, f"C{n}", val, color=BLUE, nf=nf)
    put(ws, f"D{n}", r["unit"])
    put(ws, f"E{n}", r["period"])
    put(ws, f"F{n}", r["value"], color=GREY)
    put(ws, f"G{n}", r["file"])
    put(ws, f"H{n}", int(r["page"]), align="center")
    put(ws, f"I{n}", ROLE_NAME.get(r["workstream"], ""))
    put(ws, f"J{n}", r["source_url"], color=GREY)
    put(ws, f"K{n}", r["date_consulted"], color=GREY)
    put(ws, f"L{n}", r["checked_by"] or "à relire", color=GREY if r["checked_by"] else "C00000")
    if i % 2:
        for col in "ABCDEFGHIJKL":
            ws[f"{col}{n}"].fill = BAND
LAST_ENTREE = FIRST + len(rows) - 1
widths(ws, {"A": 6, "B": 62, "C": 13, "D": 13, "E": 13, "F": 13, "G": 52, "H": 6, "I": 12,
            "J": 40, "K": 12, "L": 11})
ws.freeze_panes = "C6"
ws.auto_filter.ref = f"A5:L{LAST_ENTREE}"

# ================================================================ Hypothèses
ws = sheets["Hypothèses"]
title(ws, "Hypothèses — ce que le groupe suppose, et pourquoi",
      "Seules les cellules jaunes se modifient. Base = scénario central ; Bas / Haut = bornes de la sensibilité.")
header(ws, 4, ["Code", "Hypothèse", "Base", "Bas", "Haut", "Unité", "Justification et ancrage", "Réf.", "Actif", "Valeur active"])
H = {}   # code -> row


def hyp(row, code, label, base, low, high, unit, why, refs, nf):
    H[code] = row
    put(ws, f"A{row}", code, bold=True)
    put(ws, f"B{row}", label, wrap=True)
    for col, v in (("C", base), ("D", low), ("E", high)):
        if v is None:
            put(ws, f"{col}{row}", "—", color=GREY, align="center")
            continue
        is_formula = isinstance(v, str) and v.startswith("=")
        color = (GREEN if "Entrées" in v else BLACK) if is_formula else BLUE
        put(ws, f"{col}{row}", v, color=color, nf=nf, fill=YELLOW)
    put(ws, f"F{row}", unit)
    put(ws, f"G{row}", why, wrap=True)
    put(ws, f"H{row}", refs, color=GREY)
    put(ws, f"I{row}", "Base", color=BLUE, fill=YELLOW, align="center")
    put(ws, f"J{row}", f'=IF(AND(I{row}="Bas",ISNUMBER(D{row})),D{row},IF(AND(I{row}="Haut",ISNUMBER(E{row})),E{row},C{row}))',
        nf=nf, bold=True)


g_lo, id_glo = E("Guidance 2026 : croissance organique de l'EBITDA, bas", "2026")
g_hi, id_ghi = E("Guidance 2026 : croissance organique de l'EBITDA, haut", "2026")
cni, id_cni = E("Guidance 2026 : croissance du résultat net courant, au moins", "2026")
ce_eb, id_ceeb = E("Clean Earth : EBITDA", "2026E")
ce_usd, id_ceusd = E("Clean Earth : prix d'acquisition", "01/06/2026", "M USD")
ce_eur, id_ceeur = E("Clean Earth : prix d'acquisition", "01/06/2026", "M EUR")
nfcf25, id_nfcf25 = E("Cash-flow libre net (net free cash flow)", "FY2025")
nfcf24, id_nfcf24 = E("Cash-flow libre net (net free cash flow)", "FY2024")
nfcfh125, id_nfcfh125 = E("Cash-flow libre net avant", "S1 2025")
eb25, id_eb25 = E("EBITDA (groupe)", "FY2025")
eb24, id_eb24 = E("EBITDA (groupe)", "FY2024")
div_sh26, id_divsh26 = E("Dividende versé aux actionnaires, approuvé", "2026")
div_t26, id_divt26 = E("Pont de dette S1 2026 : dividendes versés", "S1 2026")
sign26, id_sign26 = E("Signatures de cessions attendues", "2026")
ce_mult, id_cemult = E("Clean Earth : multiple", "2026e")
tuck_lo, id_tlo = E("GreenUp : tuck-ins, bas", "2024-2027")
tuck_hi, id_thi = E("GreenUp : tuck-ins, haut", "2024-2027")
syn_cost, id_sync = E("Clean Earth : coûts de mise en œuvre", "année 1-4")
cap27, id_cap27 = E("Engagement de levier du groupe : ≤", "2027")
hyb, id_hyb = E("Dettes hybrides", "30/06/2026")

r = 5
hyp(r, "g26", "Croissance organique de l'EBITDA en 2026", f"=AVERAGE({g_lo},{g_hi})/100",
    f"={g_lo}/100", f"={g_hi}/100", "%",
    "Milieu de la guidance « +5 % à +6 % » à périmètre et change constants. Bas et haut = les bornes publiées.",
    f"{id_glo}, {id_ghi}", NF_P); r += 1
hyp(r, "g27", "Croissance organique de l'EBITDA en 2027", "=C5", 0.04, 0.07, "%",
    "Aucune guidance 2027 publiée hors l'objectif ≥ 8 Md€ : on prolonge le rythme 2026. Bornes : hypothèse du groupe.",
    "—", NF_P); r += 1
hyp(r, "ceEb", "EBITDA de Clean Earth en année pleine", f"={ce_eb}", f"={ce_eb}*0.9", f"={ce_eb}*1.1", "M USD",
    "EBITDA 2026E ajusté post IFRS 16 de la présentation d'acquisition ; ±10 % pour le risque d'intégration.",
    id_ceeb, NF_M); r += 1
hyp(r, "syn", "Part des synergies réalisée en 2027", 0.25, 0, 0.5, "%",
    "Synergies attendues « à partir de 2027 » (CP S1 2026 p.3), 120 M$ en régime de croisière l'année 4 : montée linéaire supposée.",
    "F49", NF_P); r += 1
hyp(r, "fcfH2", "Cash-flow libre net du S2 2026", f"={nfcf25}-{nfcfh125}", f"=C{r}*0.8", f"=C{r}*1.1", "M EUR",
    "On reproduit le S2 2025 (année 1 178 moins S1 −451). Le BFR se reconstitue au S2 : c'est la saisonnalité qui fait baisser la dette en fin d'année.",
    f"{id_nfcf25}, {id_nfcfh125}", NF_M); r += 1
hyp(r, "conv", "Conversion cash-flow libre net / EBITDA en 2027",
    f"=AVERAGE({nfcf24}/{eb24},{nfcf25}/{eb25})", 0.14, 0.19, "%",
    "Moyenne 2024-2025 du ratio publié. Le bas couvre le surcoût d'intérêts de la dette Clean Earth.",
    f"{id_nfcf24}, {id_eb24}, {id_nfcf25}, {id_eb25}", NF_P); r += 1
hyp(r, "gDiv", "Croissance du dividende versé en 2027", f"={cni}/100", 0.04, 0.12, "%",
    "GreenUp : dividende « aligné sur le BPA » (p.66) ; guidance 2026 du résultat net courant : au moins +8 %.",
    id_cni, NF_P); r += 1
hyp(r, "s26", "Part des signatures 2026 encaissée avant le 31/12/2026", 0.5, 0, 1, "%",
    "Veolia attend ~0,5 Md€ de signatures en 2026 ; signer n'est pas encaisser. Le levier 2026 dépend « du calendrier des cessions » (présentation CE p.16).",
    id_sign26, NF_P); r += 1
hyp(r, "s27", "Part du reste du programme encaissée en 2027", 0.5, 0, 1, "%",
    "Programme > 2 Md€ à réaliser d'ici mi-2028 : la moitié du reste supposée encaissée en 2027, le solde au S1 2028.",
    "F25", NF_P); r += 1
hyp(r, "mDisp", "Multiple VE / EBITDA des actifs cédés", 10, 8, 12, "x",
    "Aucun multiple publié pour le programme. Plus le multiple est bas, plus on perd d'EBITDA par euro encaissé.",
    "—", '0.0"x"'); r += 1
hyp(r, "tuck", "Tuck-ins payés en 2027", f"={tuck_lo}", 0, f"={tuck_hi}", "Md EUR",
    "Bas de la fourchette GreenUp (0,5 à 1,0 Md€) : Clean Earth a déjà consommé l'enveloppe.",
    f"{id_tlo}, {id_thi}", NF_D2); r += 1
moffo25, id_moffo25 = E("Moody's : FFO (funds from operations)", "FY2025")
hyp(r, "ffo", "FFO (mesure des agences) en % de l'EBITDA", f"={moffo25}/{eb25}", f"=C{r}*0.95", f"=C{r}*1.03", "%",
    "FFO 2025 publié par Moody's (Exhibit 15, 5 160 M€) rapporté à l'EBITDA 2025 publié par Veolia. Le bas couvre le surcoût d'intérêts de la dette Clean Earth.",
    f"{id_moffo25}, {id_eb25}", NF_P); r += 1
VARIED = ["g26", "g27", "ceEb", "syn", "fcfH2", "conv", "gDiv", "s26", "s27", "mDisp", "tuck", "ffo", "mTuck"]
r += 1
put(ws, f"A{r}", "Hypothèses fixes (non soumises à la sensibilité)", bold=True); r += 1
hyp(r, "fx", "Change USD par EUR", f"={ce_usd}/{ce_eur}", None, None, "USD/EUR",
    "Taux implicite du prix de Clean Earth à la clôture (2 989 M$ = 2 542 M€). Le change sur la dette en dollars n'est pas modélisé.",
    f"{id_ceusd}, {id_ceeur}", '0.0000'); r += 1
hyp(r, "minDiv", "Dividendes versés aux minoritaires en 2027", f"=-{div_t26}-{div_sh26}", None, None, "M EUR",
    "Écart du S1 2026 entre dividendes totaux (1 394) et dividende Veolia (1 099), supposé constant.",
    f"{id_divt26}, {id_divsh26}", NF_M); r += 1
m_es, id_mes = E("Espagne : multiple moyen des tuck-ins", "2024-2025")
m_wts, id_mwts = E("WTS, rachat des 30 % : multiple", "2025e")
hyp(r, "mTuck", "Multiple VE / EBITDA des tuck-ins", f"={ce_mult}", f"={m_es}", f"={m_wts}", "x",
    "Celui de Clean Earth après synergies, faute de mieux.", id_cemult, '0.0"x"'); r += 1
hyp(r, "tuckH2", "Tuck-ins payés au S2 2026", 0, None, None, "Md EUR",
    "Enviropacific (137 M€) est déjà au S1 ; rien d'annoncé pour le S2.", "—", NF_D2); r += 1
hyp(r, "other", "Autres flux de dette par période (change, divers)", 0, None, None, "M EUR",
    "Change −260 M€ au S1 2026 : non prolongé, faute de base.", "—", NF_M); r += 1
hyp(r, "months", "Mois de Clean Earth consolidés en 2026", 7, None, None, "mois",
    "Clôture le 1er juin 2026 (amendement DEU p.46) : juin à décembre.", "F40", NF_I); r += 1
hyp(r, "synC", "Coûts de mise en œuvre des synergies décaissés en 2027", f"={syn_cost}/4", None, None, "M USD",
    "90 M$ attendus sur quatre ans, étalés à parts égales.", id_sync, NF_M1); r += 1
hyp(r, "cap", "Plafond de levier", f"={cap27}", None, None, "x",
    "Engagement ≤ 3x en 2027 (présentation CE p.16, GreenUp p.64).", f"{id_cap27}, F6", '0.0"x"'); r += 1
hyp(r, "hybPct", "Part des hybrides comptée en dette (définition élargie)", 0.5, None, None, "%",
    "Lecture d'analyste, pas celle de Veolia (qui les compte en capitaux propres, IAS 32.11). Sert uniquement dans Levier.",
    id_hyb, NF_P); r += 1
put(ws, f"A{r+1}", "Les formules de base vertes renvoient à l'onglet Entrées ; les nombres bleus sont des hypothèses saisies.",
    color=GREY, italic=True)
H_LAST = r - 1
widths(ws, {"A": 8, "B": 44, "C": 11, "D": 11, "E": 11, "F": 9, "G": 78, "H": 22, "I": 9, "J": 13})
ws.freeze_panes = "C5"
for rr in range(5, r + 1):
    ws.row_dimensions[rr].height = 30
# Sélecteur de scénario : une liste déroulante par hypothèse ; tout le classeur lit la colonne J.
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName
dv = DataValidation(type="list", formula1='"Base,Bas,Haut"', allow_blank=False)
dv.errorTitle, dv.error = "Scénario", "Base, Bas ou Haut"
ws.add_data_validation(dv); dv.add(f"I5:I{H_LAST}")
dvn = DataValidation(type="decimal", operator="between", formula1="-1000000", formula2="1000000", allow_blank=True)
dvn.errorTitle, dvn.error = "Hypothèse", "Une hypothèse est un nombre (ou une formule vers Entrées)"
ws.add_data_validation(dvn); dvn.add(f"C5:E{H_LAST}")
put(ws, "A3", f'="Scénario actif : "&IF(COUNTIF($I$5:$I${H_LAST},"Bas")+COUNTIF($I$5:$I${H_LAST},"Haut")=0,"base partout",'
    f'COUNTIF($I$5:$I${H_LAST},"Bas")+COUNTIF($I$5:$I${H_LAST},"Haut")&" hypothèse(s) hors base")', bold=True, color="C00000")
HSTAT = f"{q('Hypothèses')}!$A$3"
put(ws, f"I{H_LAST + 2}", "Actif : Base, Bas ou Haut. Une hypothèse sans bornes (—) reste à sa base. Les cellules « Valeur active » (J) "
    "sont lues par Trajectoire et par tous les onglets : c'est ici qu'on joue un scénario sans toucher un nombre.", color=GREY, italic=True)
for _code, _hrow in H.items():
    _dn = DefinedName(f"h_{_code}", attr_text=f"'Hypothèses'!$J${_hrow}")
    try:
        wb.defined_names[f"h_{_code}"] = _dn
    except TypeError:
        wb.defined_names.append(_dn)


def HY(code, col="J"):
    return f"{q('Hypothèses')}!${col}${H[code]}"


# ================================================================ Levier
ws = sheets["Levier"]
title(ws, "Rôle 2 — La contrainte : la définition de Veolia, la position de départ, le plafond",
      "Dette financière nette (IFRS 16 incluse) de clôture / EBITDA (IFRS 16 inclus) — URD 2025 p.355.")
header(ws, 4, ["", "Libellé", "Valeur", "Unité", "Réf. / calcul"])
A = {}


def line(ws, row, key, label, formula, unit, ref, nf=NF_M, bold=False, store=None, color=None):
    put(ws, f"B{row}", label, bold=bold)
    col = color or (GREEN if ("Entrées" in str(formula) or "Hypothèses" in str(formula)
                              or "Trajectoire" in str(formula)) and str(formula).count("!") == 1
                    and re.fullmatch(r"=[^+\-*/()]+", str(formula)) else BLACK)
    put(ws, f"C{row}", formula, color=col, nf=nf, bold=bold)
    put(ws, f"D{row}", unit, color=GREY)
    put(ws, f"E{row}", ref, color=GREY)
    if key:
        (store if store is not None else A)[key] = f"{q(ws.title)}!$C${row}"


nfd24, i_nfd24 = E("Endettement financier net (groupe)", "31/12/2024")
nfd25, i_nfd25 = E("Endettement financier net (groupe)", "31/12/2025")
lev24, i_lev24 = E("Ratio de levier publié", "FY2024")
lev25, i_lev25 = E("Ratio de levier publié", "FY2025")
nfdh126, i_nfdh126 = E("Endettement financier net (groupe), après Clean Earth", "30/06/2026")
nfdh125, i_nfdh125 = E("Endettement financier net (groupe)", "30/06/2025")
ebh125, i_ebh125 = E("EBITDA (groupe)", "S1 2025")
ebh126, i_ebh126 = E("EBITDA (groupe)", "S1 2026")
levq1, i_levq1 = E("Ratio de levier publié, hors Clean Earth", "31/03/2026")
clos26, i_clos26 = E("Provisions de fermeture et post-fermeture (réhabilitation", "30/06/2026")

r = 5
put(ws, f"A{r}", "A", bold=True); put(ws, f"B{r}", "La définition reproduit-elle les ratios publiés ?", bold=True); r += 1
line(ws, r, "nfd24", "Dette financière nette au 31/12/2024", f"={nfd24}", "M EUR", i_nfd24); r += 1
line(ws, r, "eb24", "EBITDA 2024", f"={eb24}", "M EUR", id_eb24); r += 1
line(ws, r, "lev24c", "Levier recalculé 2024", f"=C{r-2}/C{r-1}", "x", "DFN / EBITDA", NF_X, True); r += 1
line(ws, r, "lev24p", "Levier publié 2024", f"={lev24}", "x", i_lev24, NF_X); r += 1
line(ws, r, "nfd25", "Dette financière nette au 31/12/2025", f"={nfd25}", "M EUR", i_nfd25); r += 1
line(ws, r, "eb25", "EBITDA 2025", f"={eb25}", "M EUR", id_eb25); r += 1
line(ws, r, "lev25c", "Levier recalculé 2025", f"=C{r-2}/C{r-1}", "x", "DFN / EBITDA", NF_X, True); r += 1
line(ws, r, "lev25p", "Levier publié 2025", f"={lev25}", "x", i_lev25, NF_X); r += 1
line(ws, r, "head25", "Marge sous le plafond au 31/12/2025", f"={HY('cap')}*C{r-3}-C{r-4}", "M EUR",
     "plafond × EBITDA − DFN", NF_M, True); r += 2

put(ws, f"A{r}", "B", bold=True); put(ws, f"B{r}", "Position de départ au 30/06/2026 — trois lectures", bold=True); r += 1
line(ws, r, "nfdh1", "Dette financière nette au 30/06/2026 (Clean Earth inclus)", f"={nfdh126}", "M EUR", i_nfdh126); r += 1
line(ws, r, "ltm", "EBITDA glissant 12 mois au 30/06/2026", f"={eb25}-{ebh125}+{ebh126}", "M EUR",
     f"{id_eb25} − {i_ebh125} + {i_ebh126}"); r += 1
line(ws, r, "levMech", "① Levier mécanique (à ne pas citer seul)", f"=C{r-2}/C{r-1}", "x",
     "Toute la dette de Clean Earth, un seul mois de son EBITDA", NF_X, True); r += 1
line(ws, r, "cePF", "+ 11 mois d'EBITDA de Clean Earth (pro forma)", f"={HY('ceEb')}*11/12/{HY('fx')}", "M EUR",
     "EBITDA année pleine × 11/12, converti"); r += 1
line(ws, r, "levPF", "② Levier pro forma Clean Earth", f"=C{r-4}/(C{r-3}+C{r-1})", "x", "", NF_X, True); r += 1
line(ws, r, "seas", "− saisonnalité : baisse de la dette du S2 2025", f"={nfdh125}-{nfd25}", "M EUR",
     f"{i_nfdh125} − {i_nfd25}"); r += 1
line(ws, r, "levSeas", "③ Levier pro forma, corrigé de la saisonnalité", f"=(C{r-6}-C{r-1})/(C{r-5}+C{r-3})", "x",
     "La dette de juin est structurellement plus haute qu'en décembre", NF_X, True); r += 1
line(ws, r, "levQ1", "Pour mémoire : levier publié au 31/03/2026, hors Clean Earth", f"={levq1}", "x", i_levq1, NF_X); r += 2
mat25, i_mat25 = E("Maturité moyenne de la dette financière nette", "31/12/2025")
mat24, i_mat24 = E("Maturité moyenne de la dette financière nette", "31/12/2024")
fixed, i_fixed = E("Part de la dette financière nette à taux fixe", "31/12/2025")
tgt8, i_tgt8 = E("Objectif GreenUp : EBITDA", "2027")
put(ws, f"A{r}", "B'", bold=True); put(ws, f"B{r}", "Profil de la dette et plafond en euros", bold=True); r += 1
line(ws, r, "mat25", "Maturité moyenne de la dette nette", f"={mat25}", "années", i_mat25, '0.0'); r += 1
line(ws, r, "mat24", "Maturité moyenne de la dette nette", f"={mat24}", "années", i_mat24, '0.0'); r += 1
line(ws, r, "fixed", "Part à taux fixe après couverture", f"={fixed}/100", "%", i_fixed, NF_P); r += 1
line(ws, r, "maxDebtTgt", "Dette maximale fin 2027 si l'objectif ≥ 8 Md€ est atteint (plafond × 8 000)",
     f"={HY('cap')}*{tgt8}*1000", "M EUR", f"{i_tgt8} × plafond", NF_M, True); r += 1
line(ws, r, "maxDebtMod", "Dette maximale fin 2027 à l'EBITDA du modèle (plafond × EBITDA 2027)",
     f"={HY('cap')}*{q('Trajectoire')}!$D$EB27", "M EUR", "plafond × EBITDA 2027 central", NF_M, True); r += 1
line(ws, r, "perHalf", "Dette libérée par 0,5 Md€ d'EBITDA en plus", f"={HY('cap')}*500", "M EUR", "plafond × 500", NF_M); r += 2

put(ws, f"A{r}", "B''", bold=True); put(ws, f"B{r}", "La vraie limite : les seuils de dégradation des agences (notation BBB / Baa1)", bold=True); r += 1
mo25, i_mo25 = E("Moody's : FFO / dette nette ajustée", "FY2025")
mo24, i_mo24 = E("Moody's : FFO / dette nette ajustée", "FY2024")
mo26lo, i_mo26lo = E("Moody's : FFO / dette nette ajustée attendu, bas", "2026F")
mo26hi, i_mo26hi = E("Moody's : FFO / dette nette ajustée attendu, haut", "2026F")
mond25, i_mond25 = E("Moody's : dette nette ajustée", "FY2025", "Md EUR")
mond26, i_mond26 = E("Moody's : dette nette ajustée, pic", "2026F")
motrig, i_motrig = E("Moody's : seuil de dégradation", "2026-2027")
sptrig, i_sptrig = E("S&P : seuil de dégradation", "2026-2028")
splo, i_splo = E("S&P : FFO / dette ajustés attendu, bas", "2026-2028")
sphi, i_sphi = E("S&P : FFO / dette ajustés attendu, haut", "2026-2028")
line(ws, r, "mo24", "Moody's : FFO / dette nette ajustée", f"={mo24}/100", "%", i_mo24, NF_P); r += 1
line(ws, r, "mo25", "Moody's : FFO / dette nette ajustée", f"={mo25}/100", "%", i_mo25, NF_P); r += 1
line(ws, r, "mo26", "Moody's : attendu en 2026, bas de fourchette", f"={mo26lo}/100", "%", i_mo26lo, NF_P); r += 1
line(ws, r, "mo26h", "Moody's : attendu en 2026, haut de fourchette", f"={mo26hi}/100", "%", i_mo26hi, NF_P); r += 1
line(ws, r, "motrig", "Moody's : dégradation si FFO / dette nette passe sous", f"={motrig}", "", i_motrig + " — libellé publié, pas un nombre"); ws[f"C{r}"].alignment = Alignment(horizontal="right"); r += 1
line(ws, r, "sptrig", "S&P : dégradation si FFO / dette ne reste pas durablement au-dessus de", f"={sptrig}/100", "%", i_sptrig, NF_P, True); r += 1
line(ws, r, "sp2628", "S&P : attendu 2026-2028 (bas)", f"={splo}/100", "%", f"{i_splo}, {i_sphi} (haut : 22 %)", NF_P); r += 1
line(ws, r, "mond25", "Moody's : dette nette ajustée 2025", f"={mond25}*1000", "M EUR", i_mond25, NF_M); r += 1
line(ws, r, "mond26", "Moody's : dette nette ajustée, pic 2026 (environ)", f"={mond26}*1000", "M EUR", i_mond26, NF_M); r += 1
line(ws, r, "ffo25", "FFO 2025 implicite (ratio × dette ajustée)", f"={A['mo25']}*{A['mond25']}", "M EUR", "20,3 % × 25 400", NF_M, True); r += 1
line(ws, r, "ffoReq", "FFO requis en 2026 au seuil S&P (18 % × pic de dette ajustée)", f"={A['sptrig']}*{A['mond26']}", "M EUR", "seuil × dette ajustée 2026", NF_M, True); r += 1
line(ws, r, "ffoGap", "Marge de FFO en 2026 au seuil S&P (FFO 2025 implicite − requis)", f"={A['ffo25']}-{A['ffoReq']}", "M EUR", "négatif = le seuil mord avant le 3x de Veolia", NF_M, True); r += 1
line(ws, r, "debtAt18", "Dette ajustée maximale au seuil S&P avec le FFO 2025 implicite", f"={A['ffo25']}/{A['sptrig']}", "M EUR", "FFO / 18 %", NF_M); r += 1
line(ws, r, "adjGap", "Écart entre dette ajustée Moody's 2025 et dette nette publiée (IFRS 16 incl.)", f"={A['mond25']}-{nfd25}", "M EUR", f"{i_mond25} − {i_nfd25} : décomposé ci-dessous", NF_M); r += 2
put(ws, f"A{r}", "B3", bold=True); put(ws, f"B{r}", "D'où vient l'écart : la réconciliation de Moody's (Exhibit 13), poste par poste", bold=True); r += 1
mo_rep, i_morep = E("Moody's : dette brute publiée", "FY2025")
mo_pens, i_mopens = E("Moody's : ajustement pensions", "FY2025")
mo_hyb, i_mohyb = E("Moody's : ajustement titres hybrides", "FY2025")
mo_sec, i_mosec = E("Moody's : ajustement titrisation", "FY2025")
mo_ns, i_mons = E("Moody's : ajustements non standard", "FY2025")
mo_adj, i_moadj = E("Moody's : dette brute ajustée", "FY2025", "M EUR")
mo_cash, i_mocash = E("Moody's : trésorerie retenue", "FY2025")
mo_net, i_monet = E("Moody's : dette nette ajustée (Exhibit 13)", "FY2025")
urd_gross, i_ugross = E("Sous-total des emprunts", "31/12/2025")
urd_cash, i_ucash = E("Trésorerie et équivalents", "31/12/2025")
urd_liq, i_uliq = E("Actifs liquides et actifs liés au financement", "31/12/2025")
urd_fv, i_ufv = E("Juste valeur des dérivés de couverture de dette", "31/12/2025")
line(ws, r, "moRep", "Dette brute publiée (as reported)", f"={mo_rep}", "M EUR", f"{i_morep} ; URD p.351 sous-total des emprunts : {i_ugross}"); r += 1
line(ws, r, "moPens", "+ Pensions (engagements non financés)", f"={mo_pens}", "M EUR", i_mopens); r += 1
line(ws, r, "moHyb", "+ Hybrides comptés à 50 %", f"={mo_hyb}", "M EUR", f"{i_mohyb} = 50 % × {id_hyb} (4,1 Md€)"); r += 1
line(ws, r, "moSec", "+ Titrisation (créances cédées)", f"={mo_sec}", "M EUR", i_mosec); r += 1
line(ws, r, "moNs", "+ Ajustements non standard", f"={mo_ns}", "M EUR", i_mons); r += 1
line(ws, r, "moAdjC", "Dette brute ajustée recalculée (somme)", f"=SUM(C{r-5}:C{r-1})", "M EUR", "somme", NF_M, True); r += 1
line(ws, r, "moAdjP", "Dette brute ajustée publiée par Moody's", f"={mo_adj}", "M EUR", i_moadj); r += 1
line(ws, r, "moCash", "− Trésorerie retenue par Moody's", f"={mo_cash}", "M EUR", f"{i_mocash} ; Veolia déduit {i_ucash} + {i_uliq} (8 021 + 1 952)"); r += 1
line(ws, r, "moNetC", "Dette nette ajustée recalculée (brute ajustée − trésorerie)", f"=C{r-2}-C{r-1}", "M EUR", "", NF_M, True); r += 1
line(ws, r, "moNetP", "Dette nette ajustée publiée par Moody's", f"={mo_net}", "M EUR", i_monet); r += 1
line(ws, r, "veoNet", "Pour mémoire : dette nette Veolia = brute − trésorerie − actifs liquides + JV dérivés (± PPA)", f"={urd_gross}-{urd_cash}-{urd_liq}+{urd_fv}", "M EUR",
     f"{i_ugross} − {i_ucash} − {i_uliq} + {i_ufv} ; écart résiduel = retraitement PPA Suez", NF_M); r += 1
line(ws, r, "hybShare", "Part des hybrides dans l'écart de dette nette", f"={A['moHyb']}/({A['moNetP']}-{nfd25})", "%", "hybrides / écart total", NF_P); r += 2
put(ws, f"A{r}", "B4", bold=True); put(ws, f"B{r}", "Le FFO tel que Moody's le publie (Exhibit 15)", bold=True); r += 1
mo_ffo24, i_moffo24 = E("Moody's : FFO (funds from operations)", "FY2024")
mo_ffo23, i_moffo23 = E("Moody's : FFO (funds from operations)", "FY2023")
mo_div, i_modiv = E("Moody's : dividendes", "FY2025")
mo_rcf, i_morcf = E("Moody's : RCF (retained", "FY2025")
mo_int, i_moint = E("Moody's : charge d'intérêts ajustée", "FY2025")
mo_ndeb, i_mondeb = E("Moody's : dette nette ajustée / EBITDA ajusté", "FY2025")
mo_f26, i_mof26 = E("Moody's : FFO / dette nette ajustée prévu", "2026F")
mo_f27, i_mof27 = E("Moody's : FFO / dette nette ajustée prévu", "2027F")
line(ws, r, "ffo23", "FFO 2023", f"={mo_ffo23}", "M EUR", i_moffo23); r += 1
line(ws, r, "ffo24", "FFO 2024", f"={mo_ffo24}", "M EUR", i_moffo24); r += 1
line(ws, r, "ffo25p", "FFO 2025 publié", f"={moffo25}", "M EUR", id_moffo25, NF_M, True); r += 1
line(ws, r, "ffoRatio", "FFO 2025 / EBITDA 2025 publié par Veolia (= hypothèse ffo)", f"=C{r-1}/{eb25}", "%", f"{id_moffo25} / {id_eb25}", NF_P); r += 1
line(ws, r, "ffoCheck", "FFO 2025 / dette nette ajustée publiée (doit redonner 20,3 %)", f"=C{r-2}/{A['moNetP']}", "%", "", NF_P); r += 1
line(ws, r, "moDiv", "Dividendes (minoritaires et hybrides compris, selon Moody's)", f"={mo_div}", "M EUR", i_modiv); r += 1
line(ws, r, "moRcf", "RCF = FFO − dividendes", f"={mo_rcf}", "M EUR", i_morcf); r += 1
line(ws, r, "moInt", "Charge d'intérêts ajustée", f"={mo_int}", "M EUR", i_moint); r += 1
line(ws, r, "moNdeb", "Dette nette ajustée / EBITDA ajusté selon Moody's", f"={mo_ndeb}", "x", f"{i_mondeb} — contre 2,79x publié par Veolia : même entreprise, deux définitions", NF_X); r += 1
line(ws, r, "moF26", "Prévision Moody's : FFO / dette nette 2026", f"={mo_f26}/100", "%", i_mof26, NF_P); r += 1
line(ws, r, "moF27", "Prévision Moody's : FFO / dette nette 2027", f"={mo_f27}/100", "%", i_mof27, NF_P); r += 1
put(ws, f"B{r}", "Lecture : l'écart de 5,7 Md€ entre la dette des agences et celle de Veolia tient aux hybrides pour un peu plus d'un tiers, "
    "le reste aux pensions, à la titrisation et aux retraitements de Moody's. Le FFO publié (5 160 M€) redonne le ratio 20,3 % : "
    "c'est lui que le modèle projette.", color=GREY, italic=True, wrap=True)
ws.merge_cells(f"B{r}:E{r}"); ws.row_dimensions[r].height = 44; r += 1
put(ws, f"B{r}", "Lecture : les agences ne regardent pas le 3x de Veolia mais FFO / dette ajustée. Au pic de dette 2026 (~29 Md€), tenir 18 % "
    "demande ~5,2 Md€ de FFO, à peu près le FFO 2025 implicite : la marge est nulle en 2026 et ne revient qu'avec les cessions. "
    "C'est la limite qui mord en premier, avant le plafond de 3x.", color=GREY, italic=True, wrap=True)
ws.merge_cells(f"B{r}:E{r}"); ws.row_dimensions[r].height = 44; r += 2

put(ws, f"A{r}", "B5", bold=True); put(ws, f"B{r}", "Le pont du FFO : reconstitué depuis le tableau de flux (DEU 2025 p.362-363) face au FFO de Moody's", bold=True); r += 1
header(ws, r, ["", "Poste (signes du tableau de flux)", "2025", "2024", "Réf."]); r += 1
FFO_ITEMS = [
    ("Capacité d'autofinancement avant variation du BFR", "Tableau de flux : capacité d'autofinancement avant variation du BFR"),
    ("Impôts payés", "Tableau de flux : impôts payés"),
    ("Intérêts payés", "Tableau de flux : intérêts payés"),
    ("Intérêts sur actifs IFRIC 12", "Tableau de flux : intérêts sur actifs IFRIC 12"),
    ("Intérêts sur dette de loyers IFRS 16", "Tableau de flux : intérêts sur dette de loyers IFRS 16"),
    ("Remboursements d'actifs financiers opérationnels (IFRIC 12)", "Tableau de flux : remboursements d'actifs financiers opérationnels"),
    ("Dividendes reçus (coentreprises et associés)", "Tableau de flux : dividendes reçus"),
]
FF0 = r
for lab, pre in FFO_ITEMS:
    a25, i25 = E(pre, "FY2025"); a24, i24 = E(pre, "FY2024")
    put(ws, f"B{r}", lab); put(ws, f"C{r}", f"={a25}", color=GREEN, nf=NF_M); put(ws, f"D{r}", f"={a24}", color=GREEN, nf=NF_M)
    put(ws, f"E{r}", f"{i25}, {i24}", color=GREY); r += 1
put(ws, f"B{r}", "FFO reconstitué (somme, signes du tableau de flux)", bold=True)
put(ws, f"C{r}", f"=SUM(C{FF0}:C{r-1})", nf=NF_M, bold=True); put(ws, f"D{r}", f"=SUM(D{FF0}:D{r-1})", nf=NF_M, bold=True)
ws[f"C{r}"].border = TOPLINE; ws[f"D{r}"].border = TOPLINE
A["ffoRec25"] = f"{q('Levier')}!$C${r}"; A["ffoRec24"] = f"{q('Levier')}!$D${r}"; FFR = r; r += 1
mo_ffo24b, i_moffo24b = E("Moody's : FFO (funds from operations)", "FY2024")
put(ws, f"B{r}", "FFO publié par Moody's (Exhibit 15)")
put(ws, f"C{r}", f"={moffo25}", color=GREEN, nf=NF_M); put(ws, f"D{r}", f"={mo_ffo24b}", color=GREEN, nf=NF_M)
put(ws, f"E{r}", f"{id_moffo25}, {i_moffo24b}", color=GREY); FFM = r; r += 1
put(ws, f"B{r}", "Écart reconstitué − Moody's"); put(ws, f"C{r}", f"=C{FFR}-C{FFM}", nf=NF_M); put(ws, f"D{r}", f"=D{FFR}-D{FFM}", nf=NF_M); r += 1
put(ws, f"B{r}", "Écart en % du FFO Moody's", bold=True)
put(ws, f"C{r}", f"=C{FFR}/C{FFM}-1", nf=NF_P, bold=True); put(ws, f"D{r}", f"=D{FFR}/D{FFM}-1", nf=NF_P, bold=True)
A["ffoRecGap25"] = f"{q('Levier')}!$C${r}"; A["ffoRecGap24"] = f"{q('Levier')}!$D${r}"; r += 1
mond24b, i_mond24b = E("Moody's : dette nette ajustée (Exhibit 13)", "FY2024")
put(ws, f"B{r}", "FFO reconstitué / EBITDA publié par Veolia (à comparer à l'hypothèse ffo)")
put(ws, f"C{r}", f"=C{FFR}/{eb25}", nf=NF_P); put(ws, f"D{r}", f"=D{FFR}/{eb24}", nf=NF_P)
put(ws, f"E{r}", f"{id_eb25}, {id_eb24}", color=GREY); A["ffoRecEb25"] = f"{q('Levier')}!$C${r}"; A["ffoRecEb24"] = f"{q('Levier')}!$D${r}"; r += 1
put(ws, f"B{r}", "FFO reconstitué / dette nette ajustée Moody's (publié : 20,3 % et 21,3 %)")
put(ws, f"C{r}", f"=C{FFR}/{A['moNetP']}", nf=NF_P); put(ws, f"D{r}", f"=D{FFR}/{mond24b}", nf=NF_P)
put(ws, f"E{r}", f"{i_monet}, {i_mond24b}", color=GREY); r += 1
put(ws, f"B{r}", "Lecture : Moody's ne publie pas sa formule. Lue depuis le tableau de flux de Veolia — capacité d'autofinancement avant BFR, "
    "moins impôts et intérêts payés (dette, IFRIC 12, IFRS 16), plus les remboursements d'actifs financiers opérationnels et les dividendes "
    "reçus — la reconstitution retombe sur le FFO de Moody's à moins de 2 % près sur les deux années. C'est la lecture du groupe, contrôlée "
    "à ± 5 % (Vérifications) ; elle donne un FFO projetable poste par poste, dont les intérêts (Échéancier §B).", color=GREY, italic=True, wrap=True)
ws.merge_cells(f"B{r}:E{r}"); ws.row_dimensions[r].height = 58; r += 2

put(ws, f"A{r}", "B6", bold=True); put(ws, f"B{r}", "Les pairs selon Moody's (Exhibit 12, 12 mois à juin 2025) : où se place Veolia", bold=True); r += 1
header(ws, r, ["", "Société", "Notation", "FFO / dette nette", "RCF / dette nette", "Dette / EBITDA", "EBITDA ajusté (M EUR)", "Réf."]); r += 1
PR = {}
PEER0 = r
for name in ("Veolia", "ACEA", "Hera", "Suez"):
    put(ws, f"B{r}", name, bold=name == "Veolia")
    rating, i_rt = E(f"Moody's, pairs : {name}, notation", "05/2026")
    put(ws, f"C{r}", f"={rating}", color=GREEN, align="center")
    if name == "Veolia":
        put(ws, f"D{r}", f"={A['mo25']}", color=GREEN, nf=NF_P); put(ws, f"E{r}", f"={mo_rcf}/{A['moNetP']}", nf=NF_P)
        put(ws, f"F{r}", f"={E('Moody' + chr(39) + 's : dette brute ajustée / EBITDA ajusté', 'FY2025')[0]}", color=GREEN, nf=NF_X)
        put(ws, f"G{r}", f"={E('Moody' + chr(39) + 's : EBITDA ajusté', 'FY2025')[0]}", color=GREEN, nf=NF_M)
        put(ws, f"H{r}", f"{i_rt} ; FY2025 (Exhibit 15)", color=GREY)
    else:
        refs = [i_rt]
        for col, what, nf in (("D", "FFO / dette nette", NF_P), ("E", "RCF / dette nette", NF_P), ("F", "dette / EBITDA", NF_X), ("G", "EBITDA ajusté", NF_M)):
            a, i = E(f"Moody's, pairs : {name}, {what}", "LTM 06/2025"); refs.append(i)
            put(ws, f"{col}{r}", f"={a}" + ("/100" if nf == NF_P else ""), color=GREEN, nf=nf)
        put(ws, f"H{r}", ", ".join(refs), color=GREY)
    PR[name] = r; r += 1
line(ws, r, "peerGapSuez", "Écart de FFO / dette nette entre Veolia et Suez (même notation à un cran près, perspective négative)", f"=D{PR['Veolia']}-D{PR['Suez']}", "pts", "Veolia − Suez", NF_P, True); r += 1
line(ws, r, "peerGapHera", "Écart entre Veolia et le mieux noté des pairs comparables (Hera)", f"=D{PR['Veolia']}-D{PR['Hera']}", "pts", "négatif = Veolia en dessous", NF_P, True); r += 1
put(ws, f"B{r}", "Lecture : à notation égale (Baa1), Veolia est au niveau d'ACEA et sous Hera ; Suez, Baa2 à perspective négative, est à 11 %. "
    "Le seuil des « high teens » de Moody's n'est pas théorique : c'est la zone où se trouve déjà le concurrent direct. La marge de "
    "sécurité que gardent les pairs sous leur seuil est de l'ordre de 2 à 5 points ; la nôtre, fin 2026, est de 1 point (Trajectoire).",
    color=GREY, italic=True, wrap=True)
ws.merge_cells(f"B{r}:H{r}"); ws.row_dimensions[r].height = 44; r += 2

put(ws, f"A{r}", "B7", bold=True); put(ws, f"B{r}", "Les hybrides : calendrier des premières dates de réinitialisation (DEU 2025 p.429) et ce qu'un rappel change", bold=True); r += 1
header(ws, r, ["", "Tranche", "Nominal (M EUR)", "Coupon jusqu'à la réinitialisation", "Première réinitialisation", "Coupon annuel (M EUR)", "Réf."]); r += 1
HB = {}
H0 = r
for tranche, reset in (("septembre 2019", "09/2026"), ("novembre 2021", "02/2028"), ("novembre 2023", "02/2029"),
                       ("octobre 2020, tranche résiduelle", "04/2029"), ("mai 2025 (hybride vert)", "08/2030"), ("septembre 2025", "01/2033")):
    a, i = E(f"Hybride {tranche} : nominal", "31/12/2025")
    cp, i_cp = E(f"Hybride {tranche} : coupon", "31/12/2025")
    put(ws, f"B{r}", f"Hybride {tranche}"); put(ws, f"C{r}", f"={a}", color=GREEN, nf=NF_M); put(ws, f"D{r}", f"={cp}/100", color=GREEN, nf='0.000%')
    put(ws, f"E{r}", reset, align="center"); put(ws, f"F{r}", f"=C{r}*D{r}", nf=NF_M); put(ws, f"G{r}", f"{i}, {i_cp}", color=GREY)
    HB[reset] = r; r += 1
H1 = r - 1
hyb_tot, i_hybtot = E("Hybrides : encours hors coupons", "31/12/2025")
line(ws, r, "hybSum", "Somme des tranches", f"=SUM(C{H0}:C{H1})", "M EUR", "", NF_M, True); put(ws, f"F{r}", f"=SUM(F{H0}:F{H1})", nf=NF_M, bold=True); HB["sum"] = r; r += 1
line(ws, r, "hybPub", "Encours publié hors coupons", f"={hyb_tot}*1000", "M EUR", i_hybtot); HB["pub"] = r; r += 1
line(ws, r, "hybAvgCpn", "Coupon moyen pondéré", f"=F{HB['sum']}/C{HB['sum']}", "%", "coupons / nominal", '0.00%'); r += 1
line(ws, r, "hybBefore28", "Tranches dont la première réinitialisation tombe avant fin 2027", f"=C{HB['09/2026']}", "M EUR", "septembre 2026 : 500 M€ à 1,625 %", NF_M, True); r += 1
new_cpn, i_newcpn = E("Hybride septembre 2025 : coupon", "31/12/2025")
line(ws, r, "hybRefiCost", "Si cette tranche est remplacée au coupon de la dernière émission (4,322 %) : coupon annuel en plus", f"=C{HB['09/2026']}*({new_cpn}-{E('Hybride septembre 2019 : coupon', '31/12/2025')[0]})/100", "M EUR", "nominal × écart de coupon", NF_M, True); r += 1
line(ws, r, "hybRedeemLev", "Si elle est remboursée sans remplacement : levier fin 2027 (définition Veolia) après +500 M€ de dette nette", f"=({q('Trajectoire')}!$D$NFD27+C{HB['09/2026']})/{q('Trajectoire')}!$D$EB27", "x", "DFN + 500 / EBITDA 2027", NF_X, True); r += 1
line(ws, r, "hybRedeemMoody", "… et dette ajustée Moody's fin 2027 : +500 de dette, −250 d'hybride comptée à 50 %", f"={q('Trajectoire')}!$D$ADJ27+C{HB['09/2026']}*(1-{HY('hybPct')})", "M EUR", "ADJ27 + 500 × (1 − 50 %)", NF_M); r += 1
put(ws, f"B{r}", "Lecture : 4,1 Md€ d'hybrides en capitaux propres chez Veolia, à moitié en dette chez Moody's. Une seule date tombe dans l'horizon "
    "(septembre 2026, 500 M€ à 1,625 %) ; la remplacer coûte une quinzaine de millions de coupon par an, ne pas la remplacer ajoute 500 M€ "
    "à la dette nette et 250 M€ à la dette ajustée. Le levier n'est pas la contrainte : c'est le coût du capital hybride, passé de 1,6 % "
    "à 4,3 % entre 2019 et 2025. Les grosses échéances (2028-2029 : 2,25 Md€) sont hors horizon mais pas hors question.", color=GREY, italic=True, wrap=True)
ws.merge_cells(f"B{r}:G{r}"); ws.row_dimensions[r].height = 58; r += 2

put(ws, f"A{r}", "C", bold=True); put(ws, f"B{r}", "Le plafond selon la définition retenue (scénario actif, fin 2027)", bold=True); r += 1
header(ws, r, ["", "Définition", "Levier 2027", "Marge (M EUR)", "Ce qu'on ajoute à la DFN"]); r += 1
DEF_START = r
defs = [
    ("Veolia (hybrides en capitaux propres)", "0", "rien"),
    ("Hybrides comptés à 50 % en dette", f"{HY('hybPct')}*{hyb}*1000", "50 % des 4,1 Md€ d'hybrides"),
    ("Hybrides comptés à 100 % en dette", f"{hyb}*1000", "les 4,1 Md€"),
    ("Hybrides à 50 % + provisions de fermeture", f"{HY('hybPct')}*{hyb}*1000+{clos26}",
     "lecture « dette élargie » : 50 % des hybrides et 1 290 M€ de provisions"),
]
for lab, add, what in defs:
    put(ws, f"B{r}", lab)
    put(ws, f"C{r}", f"=({q('Trajectoire')}!$D$NFD27+{add})/{q('Trajectoire')}!$D$EB27", nf=NF_X)
    put(ws, f"D{r}", f"={HY('cap')}*{q('Trajectoire')}!$D$EB27-({q('Trajectoire')}!$D$NFD27+{add})", nf=NF_M)
    put(ws, f"E{r}", what, color=GREY)
    r += 1
A["def_rows"] = (DEF_START, r - 1)
put(ws, f"B{r}", "Le plafond de 3x n'a de sens que dans la définition de Veolia. Les lignes suivantes montrent la "
    "sensibilité du ratio à la définition : ce ne sont pas des seuils d'agence.", color=GREY, italic=True, wrap=True)
ws.merge_cells(f"B{r}:E{r}"); ws.row_dimensions[r].height = 28
widths(ws, {"A": 4, "B": 58, "C": 14, "D": 14, "E": 60})

# ================================================================ Pont de dette
ws = sheets["Pont de dette"]
title(ws, "Rôles 2 et 3 — D'où vient la dette : ponts publiés et saisonnalité",
      "Signes du tableau de Veolia : un flux négatif augmente la dette.")
P = {}
r = 4
put(ws, f"A{r}", "A", bold=True); put(ws, f"B{r}", "Pont du S1 2026 (amendement DEU, p.26)", bold=True); r += 1
header(ws, r, ["", "Flux", "M EUR", "", "Réf."]); r += 1
items26 = [
    ("Cash-flow libre net", "Cash-flow libre net avant", "S1 2026"),
    ("Dividendes versés (minoritaires compris)", "Pont de dette S1 2026 : dividendes versés", "S1 2026"),
    ("Investissements financiers nets des cessions", "Pont de dette S1 2026 : investissements financiers", "S1 2026"),
    ("Variation des créances et autres actifs financiers", "Pont de dette S1 2026 : variation des créances", "S1 2026"),
    ("Émission d'actions", "Pont de dette S1 2026 : émission d'actions", "S1 2026"),
    ("Effet de change et de juste valeur", "Pont de dette S1 2026 : effet de change", "S1 2026"),
    ("Autres mouvements", "Pont de dette S1 2026 : autres mouvements", "S1 2026"),
]
line(ws, r, "open26", "Dette financière nette au 31/12/2025", f"={nfd25}", "", i_nfd25, store=P); r += 1
f0 = r
for lab, pre, per in items26:
    a, i = E(pre, per)
    line(ws, r, None, f"   {lab}", f"={a}", "", i, store=P); r += 1
line(ws, r, "sum26", "Somme des flux", f"=SUM(C{f0}:C{r-1})", "", "", bold=True, store=P); r += 1
remeas, i_rem = E("Pont de dette S1 2026 : retraitement", "S1 2026")
line(ws, r, "rem26", "Retraitement des passifs financiers (PPA Suez)", f"={remeas}", "", i_rem, store=P); r += 1
line(ws, r, "close26c", "Dette au 30/06/2026 recalculée", f"=C{f0-1}-C{r-2}-C{r-1}", "", "ouverture − flux − retraitement",
     bold=True, store=P); r += 1
line(ws, r, "close26p", "Dette au 30/06/2026 publiée", f"={nfdh126}", "", i_nfdh126, store=P); r += 1
ce_nfd, i_cenfd = E("Clean Earth : effet sur l'endettement", "S1 2026")
envp, i_envp = E("Enviropacific Services : effet", "S1 2026")
line(ws, r, "ceNfd", "   dont Clean Earth dans les investissements financiers", f"={ce_nfd}", "", i_cenfd, store=P); r += 1
line(ws, r, "envp", "   dont Enviropacific", f"={envp}", "", i_envp, store=P); r += 1
line(ws, r, "ceGap", "   Clean Earth : effet sur la dette moins prix payé", f"=-C{r-2}-{ce_eur}", "",
     f"{i_cenfd} − {id_ceeur} : écart non expliqué dans les documents", bold=True, store=P); r += 2

put(ws, f"A{r}", "B", bold=True); put(ws, f"B{r}", "Pont 2025 (communiqué annuel, p.14-15)", bold=True); r += 1
header(ws, r, ["", "Flux", "M EUR", "", "Réf."]); r += 1
items25 = [
    ("Cash-flow libre net", "Cash-flow libre net (net free cash flow)", "FY2025"),
    ("Investissements financiers nets des cessions", "Pont de dette 2025 : investissements financiers", "FY2025"),
    ("Dividendes versés aux actionnaires", "Pont de dette 2025 : dividendes versés", "FY2025"),
    ("Hybride verte (nette)", "Pont de dette 2025 : émission de la première", "FY2025"),
    ("Augmentation de capital Sequoia", "Pont de dette 2025 : augmentation de capital", "FY2025"),
    ("Annulation d'actions rachetées", "Pont de dette 2025 : réduction de capital", "FY2025"),
    ("Change et juste valeur", "Pont de dette 2025 : effet de change", "FY2025"),
]
line(ws, r, "open25", "Dette financière nette au 31/12/2024", f"={nfd24}", "", i_nfd24, store=P); r += 1
g0 = r
for lab, pre, per in items25:
    a, i = E(pre, per)
    line(ws, r, None, f"   {lab}", f"={a}", "", i, store=P); r += 1
line(ws, r, "sum25", "Somme des flux détaillés", f"=SUM(C{g0}:C{r-1})", "", "", bold=True, store=P); r += 1
line(ws, r, "close25c", "Dette au 31/12/2025 par les seuls flux détaillés", f"=C{g0-1}-C{r-1}", "", "", store=P); r += 1
line(ws, r, "close25p", "Dette au 31/12/2025 publiée", f"={nfd25}", "", i_nfd25, store=P); r += 1
divt25, i_divt25 = E("Pont de dette S1 2025 : dividendes", "S1 2025")
div25, i_div25 = E("Pont de dette 2025 : dividendes versés", "FY2025")
line(ws, r, "resid25", "Flux non détaillés dans le communiqué", f"=C{r-1}-C{r-2}", "", "à expliquer au rôle 3", bold=True,
     store=P); r += 1
line(ws, r, "minor25", "   dont au moins : dividendes aux minoritaires du S1 2025", f"=-{divt25}+{div25}", "",
     f"−{i_divt25} + {i_div25}", store=P); r += 2

put(ws, f"A{r}", "C", bold=True); put(ws, f"B{r}", "Saisonnalité : pourquoi juin n'est pas décembre", bold=True); r += 1
line(ws, r, "fcfH1", "Cash-flow libre net S1 2025", f"={nfcfh125}", "", id_nfcfh125, store=P); r += 1
line(ws, r, "fcfH2", "Cash-flow libre net S2 2025 (année − S1)", f"={nfcf25}-{nfcfh125}", "", f"{id_nfcf25} − {id_nfcfh125}",
     bold=True, store=P); r += 1
wcr26, i_wcr26 = E("Variation du BFR opérationnel", "S1 2026")
wcr25, i_wcr25 = E("Variation du BFR opérationnel", "S1 2025")
line(ws, r, "wcr25", "Variation du BFR S1 2025", f"={wcr25}", "", i_wcr25, store=P); r += 1
line(ws, r, "wcr26", "Variation du BFR S1 2026", f"={wcr26}", "", i_wcr26, store=P); r += 1
line(ws, r, "nfdJun25", "Dette au 30/06/2025", f"={nfdh125}", "", i_nfdh125, store=P); r += 1
line(ws, r, "nfdDec25", "Dette au 31/12/2025", f"={nfd25}", "", i_nfd25, store=P); r += 1
line(ws, r, "dropH2", "Baisse de la dette au S2 2025", f"=C{r-2}-C{r-1}", "", "", bold=True, store=P); r += 1
widths(ws, {"A": 4, "B": 56, "C": 14, "D": 4, "E": 58})

# ================================================================ Échéancier
ws = sheets["Échéancier"]
title(ws, "Rôle 2 — Échéancier de la dette : le mur de refinancement face aux liquidités et au cash-flow",
      "DEU 2025, note 8.1.1 (p.405-409) et note 8.3.2 (p.421-422). Les flux contractuels non actualisés comprennent "
      "le principal et les intérêts, tels que Veolia les publie.")
M, MR = {}, {}
YEARS = ["2026", "2027", "2028", "2029", "2030", "au-delà de 5 ans"]
r = 4
put(ws, f"A{r}", "A", bold=True); put(ws, f"B{r}", "Flux contractuels par année au 31/12/2025 (note 8.3.2.1, p.421)", bold=True); r += 1
header(ws, r, ["", "Instrument", "2026", "2027", "2028", "2029", "2030", "> 5 ans", "Total publié", "Somme des années", "Écart", "Réf."]); r += 1
INSTR = [("bonds", "Emprunts obligataires", "emprunts obligataires"),
         ("other", "Autres passifs et découverts bancaires", "autres passifs et découverts bancaires"),
         ("lease", "Dette de loyers IFRS 16", "dette de loyers IFRS 16")]
F0 = r
for key, lab, pre in INSTR + [("gross", "Passifs financiers bruts publiés", "passifs financiers bruts")]:
    if key == "gross":
        put(ws, f"B{r}", "Somme des trois instruments", bold=True)
        for j in range(7):
            col = L(3 + j)
            put(ws, f"{col}{r}", f"=SUM({col}{F0}:{col}{r-1})", nf=NF_M, bold=True)
        M["sum"] = r; r += 1
    put(ws, f"B{r}", lab, bold=key == "gross")
    for j, y in enumerate(YEARS):
        a, _ = E(f"Flux contractuels non actualisés : {pre}, {y}", "31/12/2025")
        put(ws, f"{L(3 + j)}{r}", f"={a}", color=GREEN, nf=NF_M)
    a, i = E(f"Flux contractuels non actualisés : {pre}, total", "31/12/2025")
    put(ws, f"I{r}", f"={a}", color=GREEN, nf=NF_M)
    put(ws, f"J{r}", f"=SUM(C{r}:H{r})", nf=NF_M)
    put(ws, f"K{r}", f"=J{r}-I{r}", nf=NF_M, color=GREY)
    put(ws, f"L{r}", i, color=GREY)
    M[key] = r; r += 1
put(ws, f"B{r}", "Écart somme recalculée − publié", color=GREY)
for j in range(7):
    col = L(3 + j)
    put(ws, f"{col}{r}", f"={col}{M['sum']}-{col}{M['gross']}", nf=NF_M, color=GREY)
r += 1
put(ws, f"B{r}", "Part de l'année dans le total", color=GREY)
for j in range(6):
    col = L(3 + j)
    put(ws, f"{col}{r}", f"={col}{M['gross']}/$I${M['gross']}", nf=NF_P, color=GREY)
r += 1
cp, i_cp = E("Billets de trésorerie (commercial paper)", "31/12/2025")
hyb863, i_hyb863 = E("Titres super-subordonnés dont le remboursement est notifié", "31/12/2025")
b1y, i_b1y = E("Emprunts obligataires à moins d'un an", "31/12/2025")
put(ws, f"B{r}", "   dont, en 2026 : billets de trésorerie, renouvelés en continu (p.408)")
put(ws, f"C{r}", f"={cp}", color=GREEN, nf=NF_M); put(ws, f"L{r}", i_cp, color=GREY); M["cp"] = r; r += 1
put(ws, f"B{r}", "   dont, en 2026 : hybride dont le remboursement est notifié, payé le 09/02/2026 (p.408)")
put(ws, f"C{r}", f"={hyb863}", color=GREEN, nf=NF_M); put(ws, f"L{r}", i_hyb863, color=GREY); M["hyb"] = r; r += 1
put(ws, f"B{r}", "Flux 2026 hors billets de trésorerie et hybride notifié", bold=True)
put(ws, f"C{r}", f"=C{M['gross']}-C{M['cp']}-C{M['hyb']}", nf=NF_M, bold=True); M["wall26"] = r; r += 1
put(ws, f"B{r}", "Pour mémoire : obligations à moins d'un an en valeur comptable (p.406) — les flux 2026 les dépassent car ils comprennent les intérêts")
put(ws, f"C{r}", f"={b1y}", color=GREEN, nf=NF_M); put(ws, f"L{r}", i_b1y, color=GREY); M["bond1y"] = r; r += 2

put(ws, f"A{r}", "B", bold=True); put(ws, f"B{r}", "Principal des souches euro par année (p.406-407) et coût de leur refinancement", bold=True); r += 1
header(ws, r, ["", "Année", "Nominal à échéance", "Taux facial moyen pondéré", "Taux de l'émission de juin 2025 (7 ans)",
               "Surcoût d'intérêts annuel si refinancé à ce taux", "Souches (registre)"]); r += 1
SER = {
    "2026": [("Souche obligataire euro à échéance du 09/06/2026", False), ("Souche obligataire euro à échéance du 30/11/2026", False)],
    "2027": [("Souche EMTN Series 43, échéance 14/01/2027", True), ("Souche EMTN Series 29 (PEO), échéance 30/03/2027", True),
             ("Souche EMTN Series 23, échéance 02/04/2027", True), ("Souche EMTN Series 3, échéance 08/06/2027", True)],
    "2028": [("Souche EMTN Series 31 (PEO), échéance 10/01/2028", True), ("Souche EMTN Series 41, échéance 15/04/2028", True),
             ("Souche EMTN Series 17, échéance 19/05/2028", True)],
    "2029": [("Souche EMTN Series 34, échéance 04/01/2029", False), ("Souche EMTN Series 19, échéance 03/04/2029", False),
             ("Souche EMTN Series 13, échéance 21/05/2029", False)],
    "2030": [("Souche EMTN Series 38, échéance 07/01/2030", False), ("Souche EMTN Series 15, échéance 01/07/2030", False),
             ("Souche EMTN Series 21, échéance 17/09/2030", False), ("Souche EMTN Series 9 (GBP), échéance 02/12/2030", False)],
}
new7, i_new7 = E("Émission obligataire du 17/06/2025, tranche 2032 : taux", "FY2025")
for year, items in SER.items():
    put(ws, f"B{r}", year, align="center")
    noms, ids = [], []
    for pre, _has_rate in items:
        a, i = E(pre + " : nominal", "31/12/2025"); noms.append(a); ids.append(i)
    put(ws, f"C{r}", "=" + "+".join(noms), color=GREEN, nf=NF_M, bold=True)
    if all(h for _, h in items):
        rates = [E(pre + " : taux", "31/12/2025")[0] for pre, _ in items]
        prod = "+".join(f"{n}*{t}" for n, t in zip(noms, rates))
        put(ws, f"D{r}", f"=({prod})/C{r}/100", nf='0.00%')
        put(ws, f"E{r}", f"={new7}/100", color=GREEN, nf='0.000%')
        put(ws, f"F{r}", f"=C{r}*(E{r}-D{r})", nf=NF_M, bold=True)
    else:
        for col in "DEF":
            put(ws, f"{col}{r}", "—", color=GREY, align="center")
    put(ws, f"G{r}", ", ".join(ids), color=GREY)
    M[f"nom{year}"] = r; r += 1
emtn, i_emtn = E("Souches EMTN : nominal total", "31/12/2025")
put(ws, f"B{r}", "Souches 2027-2030 listées ci-dessus, en part du nominal EMTN total (p.407)")
put(ws, f"C{r}", f"=SUM(C{M['nom2027']}:C{M['nom2030']})/{emtn}", nf=NF_P); put(ws, f"G{r}", i_emtn, color=GREY); r += 1
put(ws, f"B{r}", "Surcoût d'intérêts annuel, en régime, si les souches 2027 et 2028 sont refinancées au taux de juin 2025", bold=True)
put(ws, f"C{r}", f"=F{M['nom2027']}+F{M['nom2028']}", nf=NF_M, bold=True); M["extraInt"] = r; r += 1
put(ws, f"B{r}", "   en % du FFO 2025 publié par Moody's", color=GREY)
put(ws, f"C{r}", f"=C{r-1}/{moffo25}", nf=NF_P, color=GREY); put(ws, f"G{r}", id_moffo25, color=GREY); M["extraIntPct"] = r; r += 1
put(ws, f"B{r}", "Lecture : la dette qui tombe en 2027-2028 porte les coupons des années de taux bas (0 % à 1,6 %, une souche à 4,6 %). "
    "Refinancée au taux de juin 2025, elle coûte quelques dizaines de millions d'intérêts de plus par an : l'effet passe par le FFO que "
    "regardent les agences, pas par la dette nette. Le mur lui-même est petit face aux liquidités (§C).", color=GREY, italic=True, wrap=True)
ws.merge_cells(f"B{r}:L{r}"); ws.row_dimensions[r].height = 44; r += 2

put(ws, f"A{r}", "C", bold=True); put(ws, f"B{r}", "Les liquidités face au mur (note 8.3.2.2, p.421-422)", bold=True); r += 1
header(ws, r, ["", "Libellé", "Valeur", "Unité", "Réf. / calcul"]); r += 1


def ln(key, label, formula, unit, ref, nf=NF_M, bold=False):
    global r
    line(ws, r, key, label, formula, unit, ref, nf, bold, store=MR)
    M[key] = r
    r += 1


C0 = r
for key, lab, pre in (
        ("synd", "Ligne syndiquée non tirée, prolongée jusqu'en 2030", "Ligne de crédit syndiquée non tirée"),
        ("bilat", "Lignes bilatérales MT non tirées (Veolia Environnement)", "Lignes de crédit bilatérales MT non tirées"),
        ("cashVE", "Trésorerie, actifs liquides et de financement (Veolia Environnement)",
         "Trésorerie, équivalents, actifs liquides et de financement (Veolia Environnement)"),
        ("bilatSub", "Lignes bilatérales des filiales", "Lignes de crédit bilatérales des filiales"),
        ("cashSub", "Trésorerie, actifs liquides et de financement (filiales)",
         "Trésorerie, équivalents, actifs liquides et de financement (filiales)")):
    a, i = E(pre, "31/12/2025")
    ln(key, lab, f"={a}", "M EUR", i)
ln("liqC", "Liquidités totales recalculées", f"=SUM(C{C0}:C{r-1})", "M EUR", "somme", NF_M, True)
tl, i_tl = E("Total des liquidités", "31/12/2025")
ln("liqP", "Liquidités totales publiées", f"={tl}", "M EUR", i_tl)
ln("lines", "   dont lignes confirmées non tirées", f"=C{M['synd']}+C{M['bilat']}+C{M['bilatSub']}", "M EUR", "syndiquée + bilatérales")
ln("cash", "   dont trésorerie et actifs liquides", f"=C{M['cashVE']}+C{M['cashSub']}", "M EUR", "")
ln("cov26", "Liquidités / flux contractuels 2026", f"=C{M['liqP']}/C{M['gross']}", "x", "le mur de 2026 est couvert … fois", NF_X, True)
ln("cov26x", "Liquidités / flux 2026 hors billets de trésorerie et hybride notifié", f"=C{M['liqP']}/C{M['wall26']}", "x", "", NF_X, True)
ln("cov2627", "Liquidités / flux contractuels 2026 + 2027", f"=C{M['liqP']}/(C{M['gross']}+D{M['gross']})", "x", "", NF_X)
ln("cov27fcf", "(Liquidités + cash-flow libre net 2027 du modèle) / flux 2026 + 2027",
   f"=(C{M['liqP']}+{q('Trajectoire')}!$D$FCF27)/(C{M['gross']}+D{M['gross']})", "x", "FCF 2027 : Trajectoire, scénario actif", NF_X)
mat25b, i_mat25b = E("Maturité moyenne de la dette financière nette", "31/12/2025")
ln("mat", "Maturité moyenne de la dette nette", f"={mat25b}", "années", i_mat25b, '0.0')
put(ws, f"B{r}", "Covenants : la documentation des financements bancaires et obligataires de Veolia Environnement ne contient aucun covenant "
    "financier (p.422). Des covenants existent sur certains financements de filiales ; le groupe les déclare respectés au 31/12/2025.",
    color=GREY, italic=True, wrap=True)
ws.merge_cells(f"B{r}:L{r}"); ws.row_dimensions[r].height = 30; r += 1
widths(ws, {"A": 4, "B": 70, "C": 14, "D": 14, "E": 14, "F": 14, "G": 14, "H": 11, "I": 13, "J": 14, "K": 10, "L": 40})
ws.freeze_panes = "C4"

# ================================================================ Cessions
ws = sheets["Cessions"]
title(ws, "Rôles 1 et 3 — L'enveloppe : brute ou nette ? Le programme de cessions et son calendrier",
      "GreenUp, allocation du capital 2024-2027 (p.63) ; amendement DEU S1 2026 (p.29) ; présentation S1 2026 (p.4, p.14).")
C = {}
r = 4
put(ws, f"A{r}", "A", bold=True); put(ws, f"B{r}", "Veolia annonce des enveloppes nettes", bold=True); r += 1
header(ws, r, ["", "Composante", "Md EUR / an", "", "Réf."]); r += 1
maint, i_maint = E("GreenUp : investissements de maintenance", "2024-2027")
growth, i_growth = E("GreenUp : croissance sur contrats existants", "2024-2027")
inddisp, i_ind = E("GreenUp : cessions industrielles", "2024-2027")
netcap, i_netcap = E("GreenUp : investissements nets par an", "2024-2027")
line(ws, r, None, "Maintenance", f"={maint}", "", i_maint, NF_D2, store=C); r += 1
line(ws, r, None, "+ Croissance sur contrats existants", f"={growth}", "", i_growth, NF_D2, store=C); r += 1
line(ws, r, None, "− Cessions industrielles (au plus)", f"=-{inddisp}", "", i_ind, NF_D2, store=C); r += 1
line(ws, r, "netcapC", "Investissements nets recalculés (somme)", f"=SUM(C{r-3}:C{r-1})", "", "", NF_D2, True, store=C); r += 1
line(ws, r, "netcapP", "Investissements nets annoncés", f"={netcap}", "", i_netcap, NF_D2, store=C); r += 2
disp_c, i_dc = E("GreenUp : cessions continues", "2024-2027")
netrot, i_nr = E("GreenUp : rotation NETTE", "2024-2027")
line(ws, r, "tLo", "Tuck-ins, bas de fourchette", f"={tuck_lo}", "", id_tlo, NF_D2, store=C); r += 1
line(ws, r, "tHi", "Tuck-ins, haut de fourchette", f"={tuck_hi}", "", id_thi, NF_D2, store=C); r += 1
line(ws, r, "dC", "Cessions d'actifs non stratégiques (~)", f"={disp_c}", "", i_dc, NF_D2, store=C); r += 1
line(ws, r, "netLo", "Rotation nette implicite, bas (tuck-ins bas − cessions)", f"=C{r-3}-C{r-1}", "", "", NF_D2, store=C); r += 1
line(ws, r, "netHi", "Rotation nette implicite, haut (tuck-ins haut − cessions)", f"=C{r-3}-C{r-2}", "", "", NF_D2, True,
     store=C); r += 1
line(ws, r, "netAnn", "Rotation nette annoncée (~)", f"={netrot}", "", i_nr, NF_D2, store=C); r += 1
put(ws, f"B{r}", "Lecture : les ~0,5 Md€/an annoncés sont la rotation NETTE des cessions ; ils correspondent au haut de "
    "la fourchette de tuck-ins (1,0 Md€ brut) moins ~0,5 Md€ de cessions. La même convention vaut pour les "
    "investissements industriels (3,1 Md€ nets).", color=GREY, italic=True, wrap=True)
ws.merge_cells(f"B{r}:E{r}"); ws.row_dimensions[r].height = 42; r += 2

put(ws, f"A{r}", "B", bold=True); put(ws, f"B{r}", "Clean Earth face à l'enveloppe", bold=True); r += 1
prog, i_prog = E("Programme de cessions", "2026-2028")
rot8, i_rot8 = E("Rotation d'actifs sur 4 ans", "2024-2027")
line(ws, r, "ceCash", "Clean Earth : effet sur la dette", f"=-{ce_nfd}/1000", "Md EUR", i_cenfd, NF_D2, store=C); r += 1
line(ws, r, "env4", "Enveloppe nette sur 4 ans (2024-2027)", f"=C{r-3-4}*4", "Md EUR", "rotation nette annoncée × 4", NF_D2,
     store=C)
ws[f"C{r}"].value = f"={C['netAnn'].split('!')[1]}*4"; r += 1
line(ws, r, "ceYears", "Clean Earth en années d'enveloppe nette", f"=C{r-2}/{C['netAnn'].split('!')[1]}", "ans", "",
     '0.0', True, store=C); r += 1
line(ws, r, "prog", "Programme de cessions à réaliser d'ici mi-2028 (au moins)", f"={prog}", "Md EUR", i_prog, NF_D2,
     store=C); r += 1
line(ws, r, "ceNet", "Clean Earth net du programme", f"=C{r-4}-C{r-1}", "Md EUR", "", NF_D2, True, store=C); r += 1
line(ws, r, "ceNetYears", "… soit, en années d'enveloppe nette", f"=C{r-1}/{C['netAnn'].split('!')[1]}", "ans", "", '0.0',
     store=C); r += 1
line(ws, r, "rot8", "Pour mémoire : « plus de 8 Md€ de rotation d'actifs en 4 ans » (brut)", f"={rot8}", "Md EUR", i_rot8,
     NF_D2, store=C); r += 1
put(ws, f"B{r}", "Lecture : Clean Earth sort de l'enveloppe nette ; c'est le programme de cessions > 2 Md€, lancé en 2026, "
    "qui le ramène à peu près dedans. Les 8 Md€ sont un chiffre brut (achats + ventes) : à ne pas comparer aux 0,5 Md€ nets.",
    color=GREY, italic=True, wrap=True)
ws.merge_cells(f"B{r}:E{r}"); ws.row_dimensions[r].height = 42; r += 2

put(ws, f"A{r}", "C", bold=True); put(ws, f"B{r}", "Calendrier du programme (scénario central)", bold=True); r += 1
header(ws, r, ["", "Période", "Encaissé (Md EUR)", "Cumul", "Réf. / hypothèse"]); r += 1
closed, i_closed = E("Cessions réalisées", "S1 2026")
c0 = r
line(ws, r, None, "S1 2026 — réalisé (Mexique, Belgique)", f"={closed}", "", i_closed, NF_D2, store=C)
ws[f"D{r}"] = f"=C{r}"; r += 1
line(ws, r, None, "S2 2026 — part des signatures 2026", f"={sign26}*{HY('s26')}", "", f"{id_sign26} × hypothèse s26", NF_D2,
     store=C)
ws[f"D{r}"] = f"=D{r-1}+C{r}"; r += 1
line(ws, r, None, "2027 — solde des signatures 2026 + part du reste",
     f"={sign26}*(1-{HY('s26')})+MAX(0,{prog}-{closed}-{sign26})*{HY('s27')}", "",
     "hypothèses s26, s27", NF_D2, store=C)
ws[f"D{r}"] = f"=D{r-1}+C{r}"; r += 1
line(ws, r, None, "S1 2028 — solde du programme", f"=MAX(0,{prog}-D{r-1})", "", "pour atteindre le minimum annoncé", NF_D2,
     store=C)
ws[f"D{r}"] = f"=D{r-1}+C{r}"; r += 1
for rr in range(c0, r):
    ws[f"D{rr}"].number_format = NF_D2
    ws[f"D{rr}"].font = font()
line(ws, r, "ebLost", "EBITDA perdu par les cessions 2026-2027 (au multiple central)",
     f"=(C{c0+1}+C{c0+2})*1000/{HY('mDisp')}", "M EUR", "encaissements / multiple", NF_M, True, store=C); r += 2

put(ws, f"A{r}", "D", bold=True); put(ws, f"B{r}", "Ce qui a été dépensé face aux enveloppes annoncées", bold=True); r += 1
header(ws, r, ["", "Flux", "Md EUR", "", "Réf."]); r += 1
boost4, i_b4 = E("GreenUp : investissements nets de croissance 2024-2027 dans les boosters", "2024-2027")
strong2, i_s2 = E("GreenUp : à investir dans les strongholds", "2024-2027")
rot4, i_r4 = E("Rotation d'actifs déjà réalisée", "2024-2025e")
rot85, i_r85 = E("Rotation d'actifs totale GreenUp", "2024-2028")
nfi24d, i_n24d = E("Investissements financiers nets 2024", "FY2024")
nfi25d, i_n25d = E("Investissements financiers nets 2025e", "FY2025e")
nfi24, i_n24 = E("Pont de dette 2024 : investissements financiers nets des cessions", "FY2024")
nfi25, i_n25 = E("Pont de dette 2025 : investissements financiers", "FY2025")
nfih1, i_nh1 = E("Pont de dette S1 2026 : investissements financiers", "S1 2026")
pctB, i_pb = E("Part des acquisitions GreenUp dans les boosters", "2024-2025e")
pctX, i_px = E("Part des acquisitions GreenUp hors d'Europe", "2024-2025e")
line(ws, r, "env4", "Enveloppe GreenUp : investissements NETS de croissance dans les boosters", f"={boost4}", "", i_b4, NF_D2, store=C); r += 1
line(ws, r, "env2", "Enveloppe GreenUp : strongholds", f"={strong2}", "", i_s2, NF_D2, store=C); r += 1
line(ws, r, "spent24", "Sorties nettes 2024 (achats − cessions, pont de dette)", f"=-{nfi24}/1000", "", f"−{i_n24} ; deck Clean Earth : {i_n24d}", NF_D2, store=C); r += 1
line(ws, r, "spent25", "Sorties nettes 2025", f"=-{nfi25}/1000", "", f"−{i_n25} ; deck : {i_n25d}", NF_D2, store=C); r += 1
line(ws, r, "spentH1", "Sorties nettes S1 2026 (Clean Earth, Enviropacific)", f"=-{nfih1}/1000", "", f"−{i_nh1}", NF_D2, store=C); r += 1
line(ws, r, "spentCum", "Sorties nettes cumulées depuis le début de GreenUp", f"=SUM(C{r-3}:C{r-1})", "", "achats nets des cessions, 2024 → 30/06/2026", NF_D2, True, store=C); r += 1
line(ws, r, "spentVsEnv", "… face à l'enveloppe boosters de 4 Md€", f"=C{r-1}-C{r-6}", "", "positif = enveloppe nette déjà dépassée", NF_D2, True, store=C); r += 1
line(ws, r, "rot4", "Rotation déjà réalisée selon Veolia (brut, achats + ventes)", f"={rot4}", "", i_r4, NF_D2, store=C); r += 1
line(ws, r, "rot85", "Rotation totale prévue avec Clean Earth et le plan de cessions (brut)", f"={rot85}", "", i_r85, NF_D2, store=C); r += 1
line(ws, r, "pctB", "Part des acquisitions dans les boosters", f"={pctB}/100", "%", i_pb, NF_P, store=C); r += 1
line(ws, r, "pctX", "Part des acquisitions hors d'Europe", f"={pctX}/100", "%", i_px, NF_P, store=C); r += 1
put(ws, f"B{r}", "Lecture : la diapositive GreenUp p.44 dit elle-même « net growth investments » : les 4 Md€ sont nets des cessions. "
    "Les sorties nettes cumulées dépassent déjà cette enveloppe au 30 juin 2026 ; c'est le programme de cessions qui doit les y ramener.",
    color=GREY, italic=True, wrap=True)
ws.merge_cells(f"B{r}:E{r}"); ws.row_dimensions[r].height = 42; r += 2

put(ws, f"A{r}", "E", bold=True); put(ws, f"B{r}", "L'univers des cessions : ce que Veolia peut vendre, et à quel rythme elle a vendu", bold=True); r += 1
put(ws, f"B{r}", "GreenUp (p.64) : « réduire les activités matures ou banalisées » — construction, facility management sans efficacité "
    "énergétique, sélectivité sur la collecte de déchets. Le DEU 2025 (p.375) donne le chiffre d'affaires par pays ; la présentation 2025 "
    "(p.24) les activités « strongholds » ; deals.csv les cessions passées.", color=GREY, italic=True, wrap=True)
ws.merge_cells(f"B{r}:E{r}"); ws.row_dimensions[r].height = 42; r += 1
header(ws, r, ["", "E1. Chiffre d'affaires par pays (DEU 2025, p.375)", "CA 2025", "Part du groupe", "Réf."]); r += 1
grp, i_grp = E("Chiffre d'affaires du groupe (IFRS 8.33)", "FY2025")
E1 = r
for name in ("France", "États-Unis", "Allemagne", "Espagne", "Pologne", "Royaume-Uni", "République tchèque", "Australie", "Italie",
             "Belgique", "Hongrie", "Maroc", "Chine", "Chili", "Japon", "Hong Kong", "Slovaquie", "autres pays (moins de 500 M EUR chacun)"):
    a, i = E(f"Chiffre d'affaires par pays : {name}", "FY2025")
    put(ws, f"B{r}", name); put(ws, f"C{r}", f"={a}", color=GREEN, nf=NF_M); put(ws, f"D{r}", f"=C{r}/{grp}", nf=NF_P)
    put(ws, f"E{r}", i, color=GREY); r += 1
line(ws, r, "ctrySum", "Somme des pays", f"=SUM(C{E1}:C{r-1})", "M EUR", "", NF_M, True, store=C); r += 1
line(ws, r, "ctryGrp", "Chiffre d'affaires du groupe publié (IFRS 8.33)", f"={grp}", "M EUR", i_grp, store=C); r += 1
mar_b, i_marb = E("Chiffre d'affaires par pays : Maroc", "FY2024")
line(ws, r, "morocco24", "Pour mémoire : Maroc 2024, avant la cession de Lydec (2025 : 932)", f"={mar_b}", "M EUR", i_marb, store=C); r += 2

header(ws, r, ["", "E2. Le programme face au périmètre mature", "Valeur", "Unité", "Réf. / calcul"]); r += 1
marg_b, i_margb = E("Marge d'EBITDA (groupe)", "FY2025")
sh_b, i_shb = E("Strongholds : EBITDA", "FY2025")
sw_b, i_swb = E("Déchets solides : EBITDA", "FY2025")
mw_b, i_mwb = E("Eau municipale : EBITDA", "FY2025")
dh_b, i_dhb = E("Réseaux de chaleur et de froid : EBITDA", "FY2025")
line(ws, r, "progM", "Programme de cessions, minimum annoncé", f"={prog}*1000", "M EUR", i_prog, store=C); r += 1
line(ws, r, "progEb", "EBITDA cédé au multiple central (hypothèse mDisp)", f"={C['progM']}/{HY('mDisp')}", "M EUR", "programme / multiple", NF_M, True, store=C); r += 1
line(ws, r, "progRev", "Chiffre d'affaires cédé si la marge est celle du groupe", f"={C['progEb']}/({marg_b}/100)", "M EUR", f"EBITDA cédé / {i_margb}", NF_M, store=C); r += 1
line(ws, r, "progRevPct", "… en part du chiffre d'affaires du groupe", f"={C['progRev']}/{grp}", "%", "", NF_P, True, store=C); r += 1
line(ws, r, "shEb", "EBITDA des strongholds 2025", f"={sh_b}", "M EUR", i_shb, store=C); r += 1
line(ws, r, "swEb", "   dont déchets solides", f"={sw_b}", "M EUR", i_swb, store=C); r += 1
line(ws, r, "mwEb", "   dont eau municipale", f"={mw_b}", "M EUR", i_mwb, store=C); r += 1
line(ws, r, "dhEb", "   dont réseaux de chaleur et de froid", f"={dh_b}", "M EUR", i_dhb, store=C); r += 1
line(ws, r, "progShareSh", "EBITDA cédé en part des strongholds", f"={C['progEb']}/{C['shEb']}", "%", "", NF_P, True, store=C); r += 1
line(ws, r, "progShareSw", "EBITDA cédé en part des seuls déchets solides", f"={C['progEb']}/{C['swEb']}", "%", "", NF_P, True, store=C); r += 2

header(ws, r, ["", "E3. Cessions réalisées par Veolia depuis 2022 (deals.csv, produit de cession publié)", "Produit (M EUR)", "Année", "Opération, source"]); r += 1
sold = []
for d in DEALS:
    if "Veolia" not in d["seller"] or d["currency"] != "EUR":
        continue
    try:
        v = float(d["value"].replace(",", "").replace(" ", ""))
    except ValueError:
        continue
    year = (d["date_closed"] or d["date_announced"])[:4]
    sold.append((year, v, d))
sold.sort(key=lambda s: (s[0], s[2]["id"]))
D0 = r
for year, v, d in sold:
    put(ws, f"B{r}", d["target"][:60]); put(ws, f"C{r}", v, color=BLUE, nf=NF_M)
    put(ws, f"D{r}", int(year) if year.isdigit() else "n/d", align="center", color=BLACK if year.isdigit() else GREY)
    put(ws, f"E{r}", f"{d['id']} — {(d['file'] or d['source_url']).split('/')[-1][:40]}" + (f", p.{d['page']}" if d.get("page") else ""), color=GREY)
    r += 1
D1 = r - 1
line(ws, r, "soldSum", "Produits de cession cumulés 2022 → S1 2026", f"=SUM(C{D0}:C{D1})", "M EUR", f"{len(sold)} opérations à produit publié", NF_M, True, store=C); r += 1
line(ws, r, "soldN", "Nombre d'opérations", f"=COUNT(C{D0}:C{D1})", "", "", NF_I, store=C); r += 1
line(ws, r, "soldAvg", "Taille moyenne", f"=AVERAGE(C{D0}:C{D1})", "M EUR", "", NF_M, store=C); r += 1
line(ws, r, "soldMax", "La plus grande (régénération d'acide, États-Unis, 2024)", f"=MAX(C{D0}:C{D1})", "M EUR", "", NF_M, store=C); r += 1
line(ws, r, "soldPerYear", "Rythme passé : produits par an (cumul / 4 ans, 2022-2025)", f"=SUMIFS(C{D0}:C{D1},D{D0}:D{D1},\"<2026\")/4", "M EUR", "hors 2026", NF_M, True, store=C); r += 1
line(ws, r, "needAvg", "Opérations de taille moyenne nécessaires pour le programme", f"={C['progM']}/{C['soldAvg']}", "", "programme / taille moyenne", NF_D2, True, store=C); r += 1
line(ws, r, "needMax", "… de la taille de la plus grande", f"={C['progM']}/{C['soldMax']}", "", "", NF_D2, store=C); r += 1
line(ws, r, "needYears", "Années nécessaires au rythme passé", f"={C['progM']}/{C['soldPerYear']}", "années", "programme / rythme ; Veolia se donne deux ans", NF_D2, True, store=C); r += 2

header(ws, r, ["", "E4. Ce que rapporte 1 Md EUR vendu, selon le multiple", "EBITDA cédé", "Capacité de dette perdue", "Marge nette gagnée"]); r += 1
for m in (8, 10, 12):
    put(ws, f"B{r}", f"Cession de 1 000 M EUR à {m}x l'EBITDA")
    put(ws, f"C{r}", f"=1000/{m}", nf=NF_M); put(ws, f"D{r}", f"={HY('cap')}*C{r}", nf=NF_M); put(ws, f"E{r}", f"=1000-D{r}", nf=NF_M, bold=True)
    r += 1
put(ws, f"B{r}", "Lecture : il y a de quoi vendre — au multiple central, le programme retire moins de 4 % de l'EBITDA des strongholds et "
    "3 % du chiffre d'affaires du groupe, dans un portefeuille de 17 pays à plus de 500 M€. Le risque n'est pas l'existence des actifs "
    "mais le rythme : depuis 2022, Veolia a cédé une quinzaine d'activités de taille modeste ; 2 Md€ en deux ans, c'est plusieurs fois le "
    "rythme passé. C'est pour cela que s26 et s27 sont les hypothèses qui comptent, et que la question du calendrier est posée le 16 octobre.",
    color=GREY, italic=True, wrap=True)
ws.merge_cells(f"B{r}:E{r}"); ws.row_dimensions[r].height = 58; r += 1
widths(ws, {"A": 4, "B": 60, "C": 16, "D": 14, "E": 50})

# ================================================================ Trajectoire
ws = sheets["Trajectoire"]
title(ws, "Rôles 3 et 6 — Trajectoire 2026-2027 et marge de manœuvre sous 3x",
      "Colonne D = scénario actif (Hypothèses, colonne Actif ; base partout par défaut). Chaque colonne suivante change UNE hypothèse (cellule saumon). "
      "Les deux dernières combinent tous les bas défavorables, puis tous les hauts favorables.")
cols = [("D", "Scénario actif", None, None)]
ci = 5
for code in VARIED:
    for side in ("Bas", "Haut"):
        cols.append((L(ci), f"{code} {side.lower()}", code, side)); ci += 1
UNFAV, FAV = L(ci), L(ci + 1)
cols.append((UNFAV, "Défavorable combiné", "*", "unfav"))
cols.append((FAV, "Favorable combiné", "*", "fav"))
put(ws, "B4", "Scénario", bold=True)
for col, name, code, side in cols:
    put(ws, f"{col}4", name, bold=True, fill=HEAD, wrap=True, align="center")
ws.row_dimensions[4].height = 30
put(ws, "B5", "Hypothèses", bold=True)
INPUT_ORDER = VARIED + [c for c in MD.FIXED_INPUTS if c not in VARIED]
IN = {}
r = 6
for code in INPUT_ORDER:
    hrow = H[code]
    IN[code] = r
    put(ws, f"A{r}", code, color=GREY)
    put(ws, f"B{r}", sheets["Hypothèses"][f"B{hrow}"].value)
    nf = sheets["Hypothèses"][f"C{hrow}"].number_format
    put(ws, f"C{r}", sheets["Hypothèses"][f"F{hrow}"].value, color=GREY)
    r += 1
RES0 = r + 1
# lignes de calcul : clé -> (libellé, formule avec {c} = colonne, {i[code]} = ligne d'hypothèse, unité, format, gras)
import model_def as MD
NFK = {"M": NF_M, "x": NF_X, "%": NF_P, None: None}
calc = [(key, label, tpl, unit, NFK[kind], bold) for key, label, tpl, unit, kind, bold in MD.CALC]
tgt, i_tgt = E("Objectif GreenUp : EBITDA", "2027")
syn_rr, i_syn = E("Clean Earth : synergies de coûts", "année 4")
put(ws, f"B{RES0-1}", "Calcul", bold=True)
ROW = {}
rr = RES0
for key, *_ in calc:
    ROW[key] = rr; rr += 1
refs = dict(eb25=eb25, sign26=sign26, nfdh126=nfdh126, prog=prog, closed=closed, divsh26=div_sh26, tgt=tgt, synRR=syn_rr,
            mond25=mond25, nfd25=nfd25, sptrig=sptrig, monet=mo_net)
CONST_IDS = dict(eb25=id_eb25, sign26=id_sign26, nfdh126=i_nfdh126, prog=i_prog, closed=i_closed, divsh26=id_divsh26, tgt=i_tgt,
                 synRR=i_syn, mond25=i_mond25, nfd25=i_nfd25, sptrig=i_sptrig, monet=i_monet)
assert set(CONST_IDS) == set(MD.CONSTS), "model_def.CONSTS et build_model.refs divergent"
for key, label, formula, unit, nf, bold in calc:
    rr = ROW[key]
    if label is None:
        continue
    put(ws, f"A{rr}", key, color=GREY)
    put(ws, f"B{rr}", label, bold=bold)
    put(ws, f"C{rr}", unit, color=GREY)
    for col, *_ in cols:
        fmt = {**refs, **{k: f"{col}{v}" for k, v in IN.items()}, **{k: f"{col}{v}" for k, v in ROW.items()}}
        put(ws, f"{col}{rr}", formula.format(**fmt), nf=nf, bold=bold)
        if key == "WHICH27":
            ws[f"{col}{rr}"].alignment = Alignment(horizontal="right")
        if bold:
            ws[f"{col}{rr}"].border = TOPLINE if key in ("EB26", "EB27", "NFD27", "ADJ26") else Border()
# hypothèses par colonne
for code in INPUT_ORDER:
    rr = IN[code]
    hrow = H[code]
    nf = sheets["Hypothèses"][f"C{hrow}"].number_format
    for col, name, dcode, side in cols:
        if col == "D":
            f, color, fill = f"={HY(code)}", GREEN, None
        elif dcode == code:
            f, color, fill = f"={HY(code, 'D' if side == 'Bas' else 'E')}", GREEN, CHANGED
        elif dcode == "*" and code in VARIED:
            lo_col = next(c for c, n, d, s in cols if d == code and s == "Bas")
            hi_col = next(c for c, n, d, s in cols if d == code and s == "Haut")
            hl, hh = f"${lo_col}${ROW['HEAD27']}", f"${hi_col}${ROW['HEAD27']}"
            op = "<=" if side == "unfav" else ">"
            f = f"=IF({hl}{op}{hh},{HY(code, 'D')},{HY(code, 'E')})"
            color, fill = BLACK, CHANGED
        else:
            f, color, fill = f"=$D{rr}", BLACK, None
        put(ws, f"{col}{rr}", f, color=color, nf=nf, fill=fill)
# remplace les jetons NFD27 / EB27 posés dans Levier avant de connaître les lignes
for _sh in ("Levier", "Échéancier"):
    for row in sheets[_sh].iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and "$D$" in cell.value:
                for _key in ("NFD27", "EB27", "FCF27", "ADJ27"):
                    cell.value = cell.value.replace(f"$D${_key}", f"$D${ROW[_key]}")
widths(ws, {"A": 8, "B": 54, "C": 9, **{c: 11 for c, *_ in cols}})
ws.column_dimensions[UNFAV].width = 13
ws.column_dimensions[FAV].width = 13
ws.freeze_panes = "D5"
T = {k: f"{q('Trajectoire')}!$D${v}" for k, v in ROW.items()}
# graphique : levier 2024 → 2027, scénario actif, défavorable et favorable
rr = max(ROW.values()) + 2
put(ws, f"B{rr}", "Levier par scénario (source du graphique)", bold=True); rr += 1
header(ws, rr, ["", "Année", "Scénario actif", "Défavorable combiné", "Favorable combiné"]); rr += 1
G0 = rr
for year, trio in (("2024", (f"={lev24}",) * 3), ("2025", (f"={lev25}",) * 3),
                   ("2026", (f"=$D${ROW['LEV26']}", f"=${UNFAV}${ROW['LEV26']}", f"=${FAV}${ROW['LEV26']}")),
                   ("2027", (f"=$D${ROW['LEV27']}", f"=${UNFAV}${ROW['LEV27']}", f"=${FAV}${ROW['LEV27']}"))):
    put(ws, f"B{rr}", year, align="center")
    for col, f in zip("CDE", trio):
        put(ws, f"{col}{rr}", f, nf=NF_X, color=GREEN if "Entrées" in f else BLACK)
    rr += 1
G1 = rr - 1
lch = BarChart(); lch.type = "col"; lch.grouping = "clustered"
lch.title = "Levier 2024 → 2027 (plafond : 3x fin 2027)"
lch.y_axis.title = "× EBITDA"; lch.y_axis.delete = False; lch.x_axis.delete = False
lch.add_data(Reference(ws, min_col=3, max_col=5, min_row=G0 - 1, max_row=G1), titles_from_data=True)
lch.set_categories(Reference(ws, min_col=2, min_row=G0, max_row=G1))
lch.height, lch.width = 9, 18
ws.add_chart(lch, f"G{G0 - 2}")

# ================================================================ Sensibilité
ws = sheets["Sensibilité"]
title(ws, "Rôle 3 — Ce qui bouge le plus la réponse",
      "Marge de manœuvre fin 2027 sous le plafond, une hypothèse à la fois entre ses bornes. Trié par amplitude.")
header(ws, 4, ["Code", "Hypothèse", "Valeur basse", "Valeur haute", "Marge 3x si bas", "Marge 3x si haut",
               "Δ bas", "Δ haut", "Amplitude", "Levier 2026 bas", "Levier 2026 haut", "clé de tri",
               "Marge S&P si bas", "Marge S&P si haut", "Amplitude S&P"])
S0 = 5
for k, code in enumerate(VARIED):
    rr = S0 + k
    lo_col = next(c for c, n, d, s in cols if d == code and s == "Bas")
    hi_col = next(c for c, n, d, s in cols if d == code and s == "Haut")
    hrow = H[code]
    nf = sheets["Hypothèses"][f"C{hrow}"].number_format
    put(ws, f"A{rr}", code, color=GREY)
    put(ws, f"B{rr}", f"={q('Hypothèses')}!$B${hrow}", color=GREEN)
    put(ws, f"C{rr}", f"={HY(code, 'D')}", color=GREEN, nf=nf)
    put(ws, f"D{rr}", f"={HY(code, 'E')}", color=GREEN, nf=nf)
    put(ws, f"E{rr}", f"={q('Trajectoire')}!${lo_col}${ROW['HEAD27']}", color=GREEN, nf=NF_M)
    put(ws, f"F{rr}", f"={q('Trajectoire')}!${hi_col}${ROW['HEAD27']}", color=GREEN, nf=NF_M)
    put(ws, f"G{rr}", f"=E{rr}-{T['HEAD27']}", nf=NF_M)
    put(ws, f"H{rr}", f"=F{rr}-{T['HEAD27']}", nf=NF_M)
    put(ws, f"I{rr}", f"=ABS(F{rr}-E{rr})", nf=NF_M, bold=True)
    put(ws, f"J{rr}", f"={q('Trajectoire')}!${lo_col}${ROW['LEV26']}", color=GREEN, nf=NF_X)
    put(ws, f"K{rr}", f"={q('Trajectoire')}!${hi_col}${ROW['LEV26']}", color=GREEN, nf=NF_X)
    put(ws, f"L{rr}", f"=I{rr}+ROW()/1000000", color=GREY, nf='0.000000')
    put(ws, f"M{rr}", f"={q('Trajectoire')}!${lo_col}${ROW['HEADSP27']}", color=GREEN, nf=NF_M)
    put(ws, f"N{rr}", f"={q('Trajectoire')}!${hi_col}${ROW['HEADSP27']}", color=GREEN, nf=NF_M)
    put(ws, f"O{rr}", f"=ABS(N{rr}-M{rr})", nf=NF_M, bold=True)
S1 = S0 + len(VARIED) - 1
r = S1 + 2
put(ws, f"B{r}", "Base : marge de manœuvre fin 2027 sous 3x · sous le seuil S&P", bold=True)
put(ws, f"E{r}", f"={T['HEAD27']}", color=GREEN, nf=NF_M, bold=True)
put(ws, f"M{r}", f"={T['HEADSP27']}", color=GREEN, nf=NF_M, bold=True); BASE_ROW = r; r += 2
put(ws, f"A{r}", "Classement", bold=True); r += 1
header(ws, r, ["Rang", "Hypothèse", "Δ bas", "Δ haut", "Amplitude"]); r += 1
R0 = r
for k in range(len(VARIED)):
    rr = R0 + k
    put(ws, f"A{rr}", k + 1, align="center")
    m = f"MATCH(LARGE($L${S0}:$L${S1},A{rr}),$L${S0}:$L${S1},0)"
    put(ws, f"B{rr}", f"=INDEX($B${S0}:$B${S1},{m})")
    put(ws, f"C{rr}", f"=INDEX($G${S0}:$G${S1},{m})", nf=NF_M)
    put(ws, f"D{rr}", f"=INDEX($H${S0}:$H${S1},{m})", nf=NF_M)
    put(ws, f"E{rr}", f"=INDEX($I${S0}:$I${S1},{m})", nf=NF_M, bold=True)
R1 = R0 + len(VARIED) - 1
chart = BarChart()
chart.type = "bar"
chart.grouping = "clustered"
chart.overlap = 100
chart.title = "Écart à la marge centrale fin 2027, M EUR"
chart.y_axis.title = "M EUR"
chart.x_axis.scaling.orientation = "maxMin"
chart.add_data(Reference(ws, min_col=3, max_col=4, min_row=R0 - 1, max_row=R1), titles_from_data=True)
chart.set_categories(Reference(ws, min_col=2, min_row=R0, max_row=R1))
chart.height, chart.width = 10, 24
chart.x_axis.delete = False
chart.x_axis.tickLblPos = "low"
chart.y_axis.delete = False
ws.add_chart(chart, f"G{R0 - 1}")
r = R1 + 2
put(ws, f"A{r}", "Élasticités — ce que vaut 100 M€ de chaque levier sur la marge 2027", bold=True); r += 1
header(ws, r, ["", "Levier (+100 M€)", "Δ marge", "", "Pourquoi"]); r += 1
el = [
    ("EBITDA 2027", f"=100*({HY('cap')}+{HY('conv')})", "× plafond, plus le cash qu'il génère (conversion)"),
    ("Cash-flow libre", "=100", "un pour un"),
    ("Cessions (au multiple central)", f"=100*(1-{HY('cap')}/{HY('mDisp')})", "on encaisse, mais on perd de l'EBITDA × plafond"),
    ("Tuck-ins (au multiple des tuck-ins)", f"=-100*(1-{HY('cap')}/{HY('mTuck')})", "symétrique : on paie plus que 3x ce qu'on achète"),
    ("Dividendes", "=-100", "un pour un"),
    ("Dette requalifiée (hybrides, provisions)", "=-100", "un pour un, sans contrepartie d'EBITDA"),
]
for lab, f, why in el:
    put(ws, f"B{r}", lab); put(ws, f"C{r}", f, nf=NF_M); put(ws, f"E{r}", why, color=GREY); r += 1
r += 1
put(ws, f"A{r}", "Macro — ce que valent l'énergie, le dollar et les taux sur la marge 2027 (DEU 2025 p.406, 409, 416, 418 ; résultats 2025 ; amendement)", bold=True); r += 1
header(ws, r, ["", "Entrée externe", "Effet publié", "Unité", "Effet sur la marge 2027 (M EUR)", "Comment"]); r += 1
en25, i_en25 = E("Effet des prix des commodités (énergie, recyclats) sur l'EBITDA", "FY2025")
enh1, i_enh1 = E("Effet des prix des commodités sur l'EBITDA", "S1 2026")
usd_debt, i_usd = E("Dette libellée en dollars", "31/12/2025")
fx_bond, i_fxb = E("Effet de change sur les obligations en dollars", "FY2025")
fx_oi, i_fxoi = E("Sensibilité au change : résultat opérationnel si les devises se déprécient", "FY2025")
fx_fin, i_fxfin = E("Sensibilité au change : coût de l'endettement financier net", "FY2025")
flt, i_flt = E("Position nette à taux variable après couverture", "31/12/2025")
fix_pct, i_fix = E("Dette brute après couverture : part à taux fixe", "31/12/2025")
MA = {}
for key, lab, f_eff, unit, f_marge, how, i in (
        ("energy25", "Énergie et recyclats : effet sur l'EBITDA 2025", f"={en25}", "M EUR", f"=C{{r}}*{HY('cap')}", "effet × plafond : si 2027 revit 2025", i_en25),
        ("energyH1", "Énergie et recyclats : effet au S1 2026 (à annualiser : × 2)", f"={enh1}", "M EUR", f"=C{{r}}*2*{HY('cap')}", "effet × 2 × plafond : si 2027 revit le S1 2026", i_enh1),
        ("usd10", "Dollar : dette en USD, effet d'une hausse de 10 % du dollar sur la dette nette", f"={usd_debt}", "M EUR de dette", f"=-C{{r}}*0.1", "10 % de la dette en dollars, convertie", i_usd),
        ("fxbond", "Dollar : effet de change réel sur les obligations en USD en 2025 (dollar plus bas)", f"={fx_bond}", "M EUR", f"=-C{{r}}", "signe : une baisse du dollar a réduit la dette", i_fxb),
        ("fxoi", "Toutes devises : résultat opérationnel si les devises baissent de 10 % (sensibilité DEU)", f"={fx_oi}", "M EUR", f"=C{{r}}*{HY('cap')}", "résultat opérationnel ≈ EBITDA à amortissements constants", i_fxoi),
        ("rate1", "Taux : position nette à taux variable ; +1 point de taux = coût en plus", f"={flt}", "M EUR", f"=C{{r}}*0.01", "position × 1 % (la position est un passif net : négatif = coût)", i_flt),
        ("fixpct", "Taux : part de la dette brute à taux fixe après couverture", f"={fix_pct}/100", "%", "", "pour mémoire", i_fix)):
    put(ws, f"A{r}", key, color=GREY); put(ws, f"B{r}", lab)
    put(ws, f"C{r}", f_eff, color=GREEN, nf=NF_P if unit == "%" else NF_M); put(ws, f"D{r}", unit, color=GREY)
    if f_marge:
        put(ws, f"E{r}", f_marge.replace("{r}", str(r)), nf=NF_M, bold=True)
    put(ws, f"F{r}", f"{how} — {i}", color=GREY)
    MA[key] = f"{q('Sensibilité')}!$E${r}"; r += 1
put(ws, f"B{r}", "Lecture : aucune de ces entrées ne déplace la marge autant que les cessions ou l'efficacité, mais elles s'additionnent : une "
    "année d'énergie comme le S1 2026 (−120 M€ d'EBITDA) coûte 360 M€ de marge, un dollar 10 % plus fort 245 M€ de dette. Le DEU ne "
    "donne pas d'élasticité de l'EBITDA au prix de l'énergie : c'est l'effet réalisé qui sert de borne.", color=GREY, italic=True, wrap=True)
ws.merge_cells(f"B{r}:F{r}"); ws.row_dimensions[r].height = 44; r += 1
widths(ws, {"A": 7, "B": 50, "C": 12, "D": 12, "E": 12, "F": 12, "G": 10, "H": 10, "I": 11, "J": 11, "K": 11, "L": 11,
            "M": 13, "N": 13, "O": 12})
ws.column_dimensions["L"].hidden = True

hw24, i_hw24 = E("Déchets dangereux traités", "2024")
hw25, i_hw25 = E("Déchets dangereux traités", "2025")
hwt, i_hwt = E("Objectif : déchets dangereux traités en 2027", "2027")
pf_rev, i_pfrev = E("Déchets dangereux Veolia + Clean Earth : chiffre d'affaires", "2025E")
pf_eb, i_pfeb = E("Déchets dangereux Veolia + Clean Earth : EBITDA", "2025E")
tgt8, i_tgt8 = E("Objectif GreenUp : EBITDA", "2027")
ce_ev_usd, i_ceev = E("Clean Earth : valeur d'entreprise", "annonce 11/2025", "Md USD")

# ================================================================ Simulation (Monte-Carlo en formules)
import random as _random
ws = sheets["Simulation"]
title(ws, "Pari 1 — Monte-Carlo : la trajectoire rejouée sur 1 000 tirages, en formules",
      f"Chaque hypothèse variée est tirée dans une loi triangulaire (bas, base, haut de l'onglet Hypothèses). Uniformes fixées "
      f"par la graine {MD.SEED} : le classeur redonne les mêmes nombres à chaque recalcul. La ligne « base » met chaque hypothèse à sa base.")
SIM_N = MD.N_DRAWS
SIM_HEAD = 14
SIM_BASE = SIM_HEAD + 1
SIM0, SIM1 = SIM_BASE + 1, SIM_BASE + SIM_N
UCOL, XCOL = {}, {}
ci = 4
for code in VARIED:
    UCOL[code], XCOL[code] = L(ci), L(ci + 1); ci += 2
CALC_KEYS = [k for k, lab, tpl, unit, kind, bold in MD.CALC if lab is not None and kind is not None]
CCOL = {}
for key in CALC_KEYS:
    CCOL[key] = L(ci); ci += 1
# paramètres des lois
put(ws, "A4", "Lois triangulaires", bold=True)
for rr, lab in ((5, "bas"), (6, "mode = base"), (7, "haut"), (8, "F(mode)")):
    put(ws, f"A{rr}", lab, color=GREY)
for code in VARIED:
    x = XCOL[code]
    lo, base, hi = HY(code, "D"), HY(code, "C"), HY(code, "E")
    put(ws, f"{x}5", f"=MIN({lo},{base},{hi})", nf='0.0000')
    put(ws, f"{x}6", f"=MEDIAN({lo},{base},{hi})", nf='0.0000')
    put(ws, f"{x}7", f"=MAX({lo},{base},{hi})", nf='0.0000')
    put(ws, f"{x}8", f"=IF({x}7={x}5,0.5,({x}6-{x}5)/({x}7-{x}5))", nf='0.000')
# le facteur « conjoncture » commun (copule gaussienne à un facteur), hypothèse du groupe
put(ws, "A9", "exposition", color=GREY)
for code in VARIED:
    put(ws, f"{XCOL[code]}9", MD.FACTOR_LOADINGS.get(code, 0), color=BLUE, nf='0')
put(ws, "A10", "Facteur « conjoncture » commun (copule gaussienne à un facteur) — hypothèse du groupe", bold=True)
put(ws, "A11", "poids du facteur"); put(ws, "C11", MD.FACTOR_WEIGHT, color=BLUE, fill=YELLOW, nf='0.00')
put(ws, "A12", "Le poids est la corrélation entre deux hypothèses exposées (ligne 9 : 1 = exposée, 0 = indépendante) ; 0 = tirages indépendants.",
    color=GREY, italic=True)
FW = "$C$11"
put(ws, f"{L(6)}11", "Lecture : une ligne = un futur possible. Les colonnes « u » sont des tirages uniformes (nombres fixés) ; "
    "les colonnes suivantes les transforment en hypothèses, puis la trajectoire est recalculée ligne par ligne avec les formules de "
    "Trajectoire (même définition, model_def.py).", color=GREY, italic=True)
# en-têtes
put(ws, f"A{SIM_HEAD}", "Tirage", bold=True, fill=HEAD)
put(ws, f"B{SIM_HEAD}", "u conjoncture", bold=True, fill=HEAD, align="center")
for code in VARIED:
    put(ws, f"{UCOL[code]}{SIM_HEAD}", f"u {code}", bold=True, fill=HEAD, align="center")
    put(ws, f"{XCOL[code]}{SIM_HEAD}", code, bold=True, fill=HEAD, align="center")
for key in CALC_KEYS:
    put(ws, f"{CCOL[key]}{SIM_HEAD}", key, bold=True, fill=HEAD, align="center")


def _x_formula(code, r, base_row=False):
    x = XCOL[code]
    if base_row:
        return f"={x}$6"
    load = f"{x}$9"
    u = f"NORMSDIST({load}*SQRT({FW})*NORMSINV($B{r})+SQRT(1-{load}^2*{FW})*NORMSINV({UCOL[code]}{r}))"
    a_, c_, b_, f_ = f"{x}$5", f"{x}$6", f"{x}$7", f"{x}$8"
    return f"=IF({b_}={a_},{c_},IF({u}<{f_},{a_}+SQRT({u}*({b_}-{a_})*({c_}-{a_})),{b_}-SQRT((1-{u})*({b_}-{a_})*({b_}-{c_}))))"


def _calc_formula(tpl, r):
    mapping = dict(refs)
    for code in VARIED:
        mapping[code] = f"{XCOL[code]}{r}"
    for code in INPUT_ORDER:
        if code not in VARIED:
            mapping[code] = HY(code)
    for key in CALC_KEYS:
        mapping[key] = f"{CCOL[key]}{r}"
    return tpl.format(**mapping)


_rng = _random.Random(MD.SEED)
UNIFORMS = []   # exportés pour le simulateur : mêmes tirages, mêmes nombres
TPL = {k: tpl for k, lab, tpl, unit, kind, bold in MD.CALC}
NFK_SIM = {k: NFK[kind] for k, lab, tpl, unit, kind, bold in MD.CALC}
for i in range(SIM_N + 1):
    r = SIM_BASE + i
    base_row = i == 0
    ws[f"A{r}"] = "base" if base_row else i
    row_u = []
    if base_row:
        ws[f"B{r}"] = None
    else:
        um = min(1 - 1e-9, max(1e-9, _rng.random()))
        row_u.append(round(um, 9))
        ws[f"B{r}"] = round(um, 9)
    for code in VARIED:
        if base_row:
            ws[f"{UCOL[code]}{r}"] = None
        else:
            u = min(1 - 1e-9, max(1e-9, _rng.random()))
            row_u.append(round(u, 9))
            ws[f"{UCOL[code]}{r}"] = round(u, 9)
        ws[f"{XCOL[code]}{r}"] = _x_formula(code, r, base_row)
    if not base_row:
        UNIFORMS.append(row_u)
    for key in CALC_KEYS:
        ws[f"{CCOL[key]}{r}"] = _calc_formula(TPL[key], r)
# formats (une seule fois par colonne, sur la ligne base, pour garder le fichier léger)
for code in VARIED:
    ws[f"{XCOL[code]}{SIM_BASE}"].number_format = '0.0000'
for key in CALC_KEYS:
    ws[f"{CCOL[key]}{SIM_BASE}"].number_format = NFK_SIM[key] or "General"
    ws[f"{CCOL[key]}{SIM_BASE}"].font = font(bold=True)
ws.freeze_panes = f"B{SIM_BASE}"
ws.column_dimensions["A"].width = 10
SIM = {k: f"{q('Simulation')}!${CCOL[k]}${SIM0}:${CCOL[k]}${SIM1}" for k in CALC_KEYS}
SIMB = {k: f"{q('Simulation')}!${CCOL[k]}${SIM_BASE}" for k in CALC_KEYS}

# ================================================================ Distribution
ws = sheets["Distribution"]
title(ws, "Pari 1 — La capacité en probabilités",
      f"{SIM_N} futurs tirés (onglet Simulation). Une fourchette dit « entre » ; ces lignes disent « avec quelle chance ». "
      "Les lois sont triangulaires entre les bornes de l'onglet Hypothèses, la base au sommet.")
DS = {}
r = 4
put(ws, f"A{r}", "A", bold=True); put(ws, f"B{r}", "Ce que valent les résultats sur l'ensemble des tirages", bold=True); r += 1
header(ws, r, ["", "Résultat", "Moyenne", "p5", "p10", "Médiane", "p90", "p95", "Ligne base", "Trajectoire (actif)"]); r += 1
for key, lab, kind in MD.OUTPUTS:
    nf = {"M": NF_M, "x": NF_X, "%": NF_P}[kind]
    rng = SIM[key]
    put(ws, f"A{r}", key, color=GREY); put(ws, f"B{r}", lab)
    put(ws, f"C{r}", f"=AVERAGE({rng})", nf=nf)
    for col, pc in (("D", 0.05), ("E", 0.10), ("F", 0.50), ("G", 0.90), ("H", 0.95)):
        put(ws, f"{col}{r}", f"=PERCENTILE({rng},{pc})", nf=nf, bold=col == "F")
    put(ws, f"I{r}", f"={SIMB[key]}", nf=nf, color=GREEN)
    put(ws, f"J{r}", f"={T[key]}", nf=nf, color=GREEN)
    DS[key] = r; r += 1
r += 1
put(ws, f"A{r}", "B", bold=True); put(ws, f"B{r}", "Les probabilités qui répondent à la question", bold=True); r += 1
header(ws, r, ["", "Événement", "Probabilité", "", "Comment c'est compté"]); r += 1
N_ = f"{SIM_N}"
PROBS = [
    ("pBind", "La marge contraignante fin 2027 est positive (3x ET seuil S&P tenus)", f'=COUNTIF({SIM["BIND27"]},">0")/{N_}', "tirages où la plus petite des deux marges > 0"),
    ("pHead", "Le levier fin 2027 reste sous le plafond (marge 3x > 0)", f'=COUNTIF({SIM["HEAD27"]},">0")/{N_}', ""),
    ("pSP27", "Le ratio FFO / dette reste au-dessus du seuil S&P fin 2027", f'=COUNTIF({SIM["HEADSP27"]},">0")/{N_}', ""),
    ("pSP26", "Le ratio FFO / dette passe sous le seuil S&P fin 2026 (risque de dégradation)", f'=COUNTIF({SIM["RATIO26"]},"<"&{sptrig}/100)/{N_}', "seuil S&P du registre"),
    ("pAgency", "C'est le seuil des agences qui mord en premier fin 2027 (et non le 3x)", f'=SUMPRODUCT(({SIM["HEADSP27"]}<{SIM["HEAD27"]})*1)/{N_}', ""),
    ("pLev26", "Le levier fin 2026 dépasse 3,1x", f'=COUNTIF({SIM["LEV26"]},">3.1")/{N_}', "« égal ou légèrement supérieur à 3x » : 3,1x comme borne de lecture"),
    ("pAcq1", "Veolia peut acheter au moins 1 Md€ fin 2027 sans franchir la contrainte qui mord", f'=COUNTIF({SIM["MAXACQB"]},">=1000")/{N_}', "au multiple des tuck-ins tiré"),
    ("pAcq2", "… au moins 2 Md€", f'=COUNTIF({SIM["MAXACQB"]},">=2000")/{N_}', ""),
    ("pAcq3", "… au moins 3 Md€", f'=COUNTIF({SIM["MAXACQB"]},">=3000")/{N_}', ""),
    ("pGap8", "L'EBITDA 2027 atteint l'objectif de 8 Md€", f'=COUNTIF({SIM["GAP8"]},">=0")/{N_}', ""),
    ("nNeg", "Futurs (sur l'ensemble des tirages) où la marge contraignante devient négative", f'=COUNTIF({SIM["BIND27"]},"<0")', "un nombre de tirages, pas une probabilité"),
]
for key, lab, f, how in PROBS:
    put(ws, f"B{r}", lab, bold=key in ("pBind", "pAcq1"))
    put(ws, f"C{r}", f, nf=NF_I if key == "nNeg" else '0%', bold=True)
    put(ws, f"E{r}", how, color=GREY)
    DS[key] = f"{q('Distribution')}!$C${r}"; r += 1
r += 1
put(ws, f"A{r}", "C", bold=True); put(ws, f"B{r}", "Histogramme de la marge contraignante fin 2027 (M EUR)", bold=True); r += 1
header(ws, r, ["", "Classe (de … à …)", "De", "À", "Tirages"]); r += 1
rngB = SIM["BIND27"]
HB0 = r
NB = 14
for k in range(NB):
    lo = f"MIN({rngB})+(MAX({rngB})-MIN({rngB}))*{k}/{NB}"
    hi = f"MIN({rngB})+(MAX({rngB})-MIN({rngB}))*{k + 1}/{NB}"
    put(ws, f"C{r}", f"={lo}", nf=NF_M); put(ws, f"D{r}", f"={hi}", nf=NF_M)
    put(ws, f"B{r}", f'=TEXT(C{r},"# ##0")&" à "&TEXT(D{r},"# ##0")')
    cond = f'"<="&D{r}' if k == NB - 1 else f'"<"&D{r}'
    put(ws, f"E{r}", f'=COUNTIFS({rngB},">="&C{r},{rngB},{cond})', nf=NF_I)
    r += 1
HB1 = r - 1
hch = BarChart(); hch.type = "col"; hch.grouping = "clustered"; hch.gapWidth = 10
hch.title = "Marge contraignante fin 2027 sur les tirages (M EUR)"
hch.y_axis.title = "tirages"; hch.y_axis.delete = False; hch.x_axis.delete = False
hch.add_data(Reference(ws, min_col=5, min_row=HB0, max_row=HB1), titles_from_data=False)
hch.set_categories(Reference(ws, min_col=2, min_row=HB0, max_row=HB1))
hch.legend = None; hch.height, hch.width = 8, 20
ws.add_chart(hch, f"G{HB0 - 1}")
r += 1
put(ws, f"A{r}", "D", bold=True); put(ws, f"B{r}", "L'éventail du levier 2024 → 2027", bold=True); r += 1
header(ws, r, ["", "Année", "p10", "Médiane", "p90"]); r += 1
FA0 = r
for year, trio in (("2024", (f"={lev24}",) * 3), ("2025", (f"={lev25}",) * 3),
                   ("2026", tuple(f"=PERCENTILE({SIM['LEV26']},{p})" for p in (0.1, 0.5, 0.9))),
                   ("2027", tuple(f"=PERCENTILE({SIM['LEV27']},{p})" for p in (0.1, 0.5, 0.9)))):
    put(ws, f"B{r}", year, align="center")
    for col, f in zip("CDE", trio):
        put(ws, f"{col}{r}", f, nf=NF_X, color=GREEN if "Entrées" in f else BLACK)
    r += 1
FA1 = r - 1
from openpyxl.chart import LineChart
fch = LineChart(); fch.title = "Levier : p10, médiane, p90 des tirages (plafond 3x en 2027)"
fch.y_axis.title = "× EBITDA"; fch.y_axis.delete = False; fch.x_axis.delete = False
fch.add_data(Reference(ws, min_col=3, max_col=5, min_row=FA0 - 1, max_row=FA1), titles_from_data=True)
fch.set_categories(Reference(ws, min_col=2, min_row=FA0, max_row=FA1))
fch.height, fch.width = 8, 16
ws.add_chart(fch, f"G{FA0 - 1}")
r += 1
put(ws, f"B{r}", "Lecture : la fourchette défavorable / favorable de Trajectoire combine tous les extrêmes à la fois, ce qui n'arrive "
    "presque jamais ; ici, chaque futur tire chaque hypothèse dans sa loi, et un facteur « conjoncture » commun (poids en C11 de "
    "Simulation) fait bouger ensemble les hypothèses qui y sont exposées. Les probabilités "
    "dépendent des bornes choisies dans Hypothèses : elles mesurent notre incertitude, pas celle du marché.", color=GREY, italic=True, wrap=True)
ws.merge_cells(f"B{r}:J{r}"); ws.row_dimensions[r].height = 44
widths(ws, {"A": 8, "B": 62, "C": 12, "D": 12, "E": 12, "F": 12, "G": 12, "H": 12, "I": 12, "J": 14})
ws.freeze_panes = "C4"

# ================================================================ Financement (pari 4)
ws = sheets["Financement"]
title(ws, "Pari 4 — Comment acheter plus que la capacité : le menu de financement, et les seuils de bascule",
      "Pour une cible de taille donnée : ce qui manque sous chaque plafond fin 2027, ce que chaque levier rapporte par euro levé, "
      "ce qu'il coûte. Puis, pour chaque conclusion du rapport, la valeur de chaque hypothèse qui la ferait basculer.")
FN = {}
TJ = q("Trajectoire")
price, i_price = E("Cours de clôture de l'action Veolia", "31/12/2025")
shares, i_shares = E("Nombre d'actions composant le capital", "31/12/2025")
treas, i_treas = E("Actions autodétenues", "31/12/2025")
hyb_cpn, i_hybcpn = E("Hybride septembre 2025 : coupon", "31/12/2025")
r = 4
put(ws, f"A{r}", "A", bold=True); put(ws, f"B{r}", "La cible étudiée et ce qu'elle coûte en marge fin 2027", bold=True); r += 1
header(ws, r, ["", "Libellé", "Valeur", "Unité", "Réf. / calcul"]); r += 1
put(ws, f"B{r}", "Taille de la cible (valeur d'entreprise payée)", bold=True)
put(ws, f"C{r}", 3000, color=BLUE, fill=YELLOW, nf=NF_M); put(ws, f"D{r}", "M EUR", color=GREY)
put(ws, f"E{r}", "cellule jaune : une cible « taille Clean Earth » par défaut", color=GREY); FN["S"] = f"$C${r}"; r += 1
put(ws, f"B{r}", "Multiple VE / EBITDA payé", bold=True)
put(ws, f"C{r}", f"={HY('mTuck')}", color=GREEN, fill=YELLOW, nf='0.0"x"'); put(ws, f"D{r}", "x", color=GREY)
put(ws, f"E{r}", "par défaut, l'hypothèse mTuck active ; remplaçable", color=GREY); FN["m"] = f"$C${r}"; r += 1
line(ws, r, "ebT", "EBITDA apporté par la cible", f"={FN['S']}/{FN['m']}", "M EUR", "taille / multiple", store=FN); r += 1
line(ws, r, "h3", "Marge sous 3x fin 2027, avant la cible (scénario actif)", f"={T['HEAD27']}", "M EUR", "Trajectoire", store=FN); r += 1
line(ws, r, "hS", "Marge sous le seuil S&P fin 2027, avant la cible", f"={T['HEADSP27']}", "M EUR", "Trajectoire", store=FN); r += 1
line(ws, r, "h3p", "Marge sous 3x après la cible", f"={FN['h3']}-{FN['S']}+{HY('cap')}*{FN['ebT']}", "M EUR",
     "− dette payée + plafond × EBITDA acquis", NF_M, True, store=FN); r += 1
line(ws, r, "hSp", "Marge sous le seuil S&P après la cible", f"={FN['hS']}-{FN['S']}+{HY('ffo')}*{FN['ebT']}/({sptrig}/100)", "M EUR",
     "− dette ajustée + FFO acquis / seuil", NF_M, True, store=FN); r += 1
line(ws, r, "g3", "Ce qui manque sous 3x", f"=MAX(0,-{FN['h3p']})", "M EUR", "", NF_M, True, store=FN); r += 1
line(ws, r, "gS", "Ce qui manque sous le seuil S&P", f"=MAX(0,-{FN['hSp']})", "M EUR", "", NF_M, True, store=FN); r += 2

put(ws, f"A{r}", "B", bold=True); put(ws, f"B{r}", "Le menu : ce que chaque levier rapporte, ce qu'il faut lever, ce que cela coûte", bold=True); r += 1
header(ws, r, ["", "Levier", "Marge 3x par € levé", "Marge S&P par € levé", "Montant nécessaire (M EUR)", "Coût", "Unité du coût", "Ce que dit la dataroom"]); r += 1
div27 = f"({div_sh26}*(1+{HY('gDiv')}))"
LEVERS = [
    ("hyb", "Émission d'hybrides (capitaux propres pour Veolia, 50 % en dette pour les agences)",
     "1", f"1-{HY('hybPct')}", f"=E{{r}}*{hyb_cpn}/100", "M EUR de coupon par an",
     f"coupon de la dernière émission ({i_hybcpn}) ; plafond d'equity credit des agences non publié dans la dataroom"),
    ("disp", "Cessions accélérées (au multiple mDisp)",
     f"1-{HY('cap')}/{HY('mDisp')}", f"1-{HY('ffo')}/({sptrig}/100)/{HY('mDisp')}", f"=E{{r}}/{HY('mDisp')}", "M EUR d'EBITDA perdu par an",
     "chaque euro encaissé retire 1/multiple d'EBITDA : il faut céder plus que l'écart"),
    ("equity", "Augmentation de capital (au cours du 31/12/2025)",
     "1", "1", f"=E{{r}}*1000000/{price}/({shares}-{treas})", "% du capital (dilution)",
     f"cours {i_price}, actions {i_shares} moins autodétenues {i_treas}"),
    ("scrip", "Dividende 2027 payé en actions (au plus le dividende)",
     "1", "1", f"=IF(E{{r}}>{div27},\"au-delà du dividende\",E{{r}}*1000000/{price}/({shares}-{treas}))", "% du capital (dilution)",
     "plafonné au dividende 2027 du modèle : un levier d'appoint"),
    ("minor", "Partenaire minoritaire dans la cible (part cédée au prix payé)",
     "1", "1", f"=E{{r}}/{FN['m']}", "M EUR d'EBITDA revenant au partenaire",
     "le précédent : WTS, détenue à 70 % puis rachetée (D14) ; au plus 49 % pour garder le contrôle"),
]
L0 = r
for key, lab, e3, eS, cost, unit, note in LEVERS:
    put(ws, f"A{r}", key, color=GREY); put(ws, f"B{r}", lab, wrap=True)
    put(ws, f"C{r}", f"={e3}", nf='0.00'); put(ws, f"D{r}", f"={eS}", nf='0.00')
    need = f'=IF(OR(C{r}<=0,D{r}<=0),"sans effet",MAX({FN["g3"]}/C{r},{FN["gS"]}/D{r}))'
    if key == "minor":
        need = f'=IF(MAX({FN["g3"]}/C{r},{FN["gS"]}/D{r})>0.49*{FN["S"]},"au-delà de 49 %",MAX({FN["g3"]}/C{r},{FN["gS"]}/D{r}))'
    put(ws, f"E{r}", need, nf=NF_M, bold=True)
    put(ws, f"F{r}", cost.format(r=r), nf=NF_P if "capital" in unit else NF_M)
    put(ws, f"G{r}", unit, color=GREY); put(ws, f"H{r}", note, color=GREY, wrap=True)
    ws.row_dimensions[r].height = 30
    FN[key] = r; r += 1
r += 1
put(ws, f"A{r}", "C", bold=True); put(ws, f"B{r}", "La combinaison du groupe (cellules jaunes) : les deux plafonds tiennent-ils ?", bold=True); r += 1
header(ws, r, ["", "Levier", "Montant retenu (M EUR)", "Apport marge 3x", "Apport marge S&P", "Coût"]); r += 1
C0 = r
for key, lab, *_ in LEVERS:
    lr = FN[key]
    put(ws, f"B{r}", f"=B{lr}")
    put(ws, f"C{r}", 0, color=BLUE, fill=YELLOW, nf=NF_M)
    put(ws, f"D{r}", f"=C{r}*C{lr}", nf=NF_M); put(ws, f"E{r}", f"=C{r}*D{lr}", nf=NF_M)
    cost_expr = LEVERS[[k for k, *_ in LEVERS].index(key)][4].replace("E{r}", f"C{r}")
    put(ws, f"F{r}", cost_expr.format(r=r) if "{r}" in cost_expr else cost_expr, nf=NF_P if "capital" in LEVERS[[k for k, *_ in LEVERS].index(key)][5] else NF_M)
    r += 1
C1 = r - 1
line(ws, r, "mix3", "Marge sous 3x après la cible et la combinaison", f"={FN['h3p']}+SUM(D{C0}:D{C1})", "M EUR", "", NF_M, True, store=FN); r += 1
line(ws, r, "mixS", "Marge sous le seuil S&P après la cible et la combinaison", f"={FN['hSp']}+SUM(E{C0}:E{C1})", "M EUR", "", NF_M, True, store=FN); r += 1
put(ws, f"B{r}", "Les deux plafonds tiennent ?", bold=True)
put(ws, f"C{r}", f'=IF(AND({FN["mix3"]}>=0,{FN["mixS"]}>=0),"oui","non")', bold=True, align="center"); FN["mixOk"] = f"$C${r}"; r += 2
put(ws, f"B{r}", "Lecture : avec le multiple de Clean Earth, une cible de 3 Md€ fin 2027 manque de marge sous 3x ; la combler par des hybrides coûte un "
    "coupon de l'ordre de 4 %, par des cessions coûte un EBITDA récurrent, par du capital une dilution. La ligne de partage n'est pas "
    "technique : c'est le prix que Veolia accepte de payer pour la taille. Les agences, elles, ne retiennent qu'à moitié les hybrides.",
    color=GREY, italic=True, wrap=True)
ws.merge_cells(f"B{r}:H{r}"); ws.row_dimensions[r].height = 44; r += 2

put(ws, f"A{r}", "D", bold=True); put(ws, f"B{r}", "Seuils de bascule : la valeur de chaque hypothèse (seule à bouger) qui renverse chaque conclusion", bold=True); r += 1
put(ws, f"B{r}", "Calcul exact : chaque marge est affine dans l'hypothèse (ou dans son inverse, pour les deux multiples) ; le seuil se lit "
    "entre ses valeurs aux bornes basse et haute de Trajectoire. « dans la fourchette » = la conclusion peut basculer sans sortir des bornes.",
    color=GREY, italic=True, wrap=True)
ws.merge_cells(f"B{r}:H{r}"); ws.row_dimensions[r].height = 30; r += 1
put(ws, f"B{r}", "Levier fin 2026 lu « légèrement supérieur à 3x » au plus à", bold=False)
put(ws, f"C{r}", 3.1, color=BLUE, fill=YELLOW, nf=NF_X); FN["lev26max"] = f"$C${r}"; r += 1
CONCL = [
    ("C1", "Levier 2026 ≤ seuil de lecture", lambda c: f"{FN['lev26max']}*{TJ}!{c}{ROW['EB26']}-{TJ}!{c}{ROW['NFD26']}"),
    ("C2", "Plafond de 3x tenu fin 2027", lambda c: f"{TJ}!{c}{ROW['HEAD27']}"),
    ("C3", "Le 3x mord avant les agences", lambda c: f"{TJ}!{c}{ROW['HEADSP27']}-{TJ}!{c}{ROW['HEAD27']}"),
    ("C4", "Cible de la taille étudiée possible sous 3x (au multiple des tuck-ins)", lambda c: f"{TJ}!{c}{ROW['HEAD27']}-{FN['S']}*(1-{HY('cap')}/{TJ}!{c}{IN['mTuck']})"),
]
header(ws, r, ["", "Hypothèse", "Base"] + [f"{k} : {lab}" for k, lab, _ in CONCL] + ["", "Réf."], start=1); r += 1
INV = {"mDisp", "mTuck"}
BK0 = r
for code in VARIED:
    lo_col = next(c for c, n, d, s in cols if d == code and s == "Bas")
    hi_col = next(c for c, n, d, s in cols if d == code and s == "Haut")
    hrow = H[code]
    nf = sheets["Hypothèses"][f"C{hrow}"].number_format
    put(ws, f"A{r}", code, color=GREY); put(ws, f"B{r}", f"={q('Hypothèses')}!$B${hrow}", color=GREEN)
    put(ws, f"C{r}", f"={HY(code, 'C')}", color=GREEN, nf=nf)
    hl, hh = f"{TJ}!{lo_col}{IN[code]}", f"{TJ}!{hi_col}{IN[code]}"
    unit = sheets["Hypothèses"][f"F{hrow}"].value
    tfmt = {"%": "0.0%", "x": '0.0""x""', "Md EUR": "0.00"}.get(unit, "#,##0")   # guillemets doublés : on est dans une formule
    xl, xh = (f"(1/{hl})", f"(1/{hh})") if code in INV else (hl, hh)
    for k, (cid, lab, metric) in enumerate(CONCL):
        col = L(4 + k)
        ml, mh = metric(lo_col), metric(hi_col)
        star = f"({xl}+(0-({ml}))*({xh}-{xl})/(({mh})-({ml})))"
        value = f"1/{star}" if code in INV else star
        num_col = L(11 + k)
        put(ws, f"{num_col}{r}", f'=IF(ABS(({mh})-({ml}))<1E-9,"",{value})', color=GREY, nf='0.0000')
        # à la française : espace fine pour les milliers, virgule décimale, vrai signe moins
        shown = f'SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(TEXT({num_col}{r},"{tfmt}"),",","\u202f"),".",","),"-","−")'
        put(ws, f"{col}{r}", f'=IF({num_col}{r}="","sans effet",IF(AND(({ml})>=0,({mh})>=0),"tient partout (seuil "&{shown}&")",'
                              f'IF(AND(({ml})<0,({mh})<0),"faux partout (il faudrait "&{shown}&")","bascule à "&{shown})))')
    put(ws, f"I{r}", "inverse du multiple" if code in INV else "", color=GREY)
    put(ws, f"J{r}", f'=IF(COUNTIF(D{r}:G{r},"bascule*")>0,1,0)', color=GREY, nf=NF_I)
    r += 1
BK1 = r - 1
put(ws, f"B{r}", "Lecture : « bascule à x » = la conclusion change de signe entre les bornes basse et haute de l'hypothèse, au point x ; "
    "« hors fourchette » = elle tient sur toute la plage, et le seuil (affiché pour mémoire) est en dehors. C'est la liste des hypothèses "
    "qu'il faut défendre à l'oral, et de nulle autre.", color=GREY, italic=True, wrap=True)
ws.merge_cells(f"B{r}:H{r}"); ws.row_dimensions[r].height = 44; r += 1
put(ws, f"B{r}", "Hypothèses qui font basculer au moins une conclusion dans leurs bornes", bold=True)
put(ws, f"C{r}", f"=SUM(J{BK0}:J{BK1})", nf=NF_I, bold=True)
FN["nFragile"] = f"{q('Financement')}!$C${r}"; r += 1
for k, (cid, lab, _m) in enumerate(CONCL):
    put(ws, f"{L(11 + k)}{BK0 - 1}", f"{cid} (valeur)", bold=True, fill=HEAD, color=GREY)
FN["bk0"], FN["bk1"] = BK0, BK1
widths(ws, {"A": 8, "B": 58, "C": 16, "D": 26, "E": 26, "F": 26, "G": 26, "H": 46, "I": 16, "K": 11, "L": 11, "M": 11, "N": 11})
ws.freeze_panes = "C4"
FIN_REF = {k: (f"{q('Financement')}!{v}" if isinstance(v, str) and v.startswith("$") else v) for k, v in FN.items()}

# ================================================================ Segments
ws = sheets["Segments"]
title(ws, "Rôle 4 (question 1 du cours) — D'où viennent les 8 Md€ : segments, boosters, KPI ESG",
      "Présentation des résultats 2025 (p.7, p.23, p.25, p.41), rapport de gestion 2025 (p.12-13), présentation S1 2026 (p.10, p.19, p.22), DEU 2025 (p.12, p.191).")
SG = {}
r = 4
put(ws, f"A{r}", "A", bold=True); put(ws, f"B{r}", "Les segments : qui livre la croissance", bold=True); r += 1
header(ws, r, ["", "Segment", "CA 2024", "CA 2025", "EBITDA 2024", "EBITDA 2025", "Marge 2025", "Croiss. EBITDA publiée",
               "Croiss. organique EBITDA 2025", "CA S1 2026", "EBITDA S1 2026", "Réf."]); r += 1
def seg_row(label, rev24, rev25, eb24, eb25, g25, revh1, ebh1):
    """Chaque argument est (préfixe de libellé, période) ou None ; la cellule reste vide si rien n'est au registre."""
    refs = []
    put(ws, f"B{r}", label)
    for col, spec, nf in (("C", rev24, NF_M), ("D", rev25, NF_M), ("E", eb24, NF_M), ("F", eb25, NF_M),
                          ("I", g25, NF_P), ("J", revh1, NF_M), ("K", ebh1, NF_M)):
        if spec is None:
            put(ws, f"{col}{r}", None); continue
        addr, fid = E(*spec)
        put(ws, f"{col}{r}", f"={addr}" + ("/100" if nf == NF_P else ""), color=GREEN, nf=nf); refs.append(fid)
    put(ws, f"G{r}", f'=IF(AND(ISNUMBER(D{r}),ISNUMBER(F{r}),D{r}<>0),F{r}/D{r},"")', nf=NF_P)
    put(ws, f"H{r}", f'=IF(AND(ISNUMBER(E{r}),ISNUMBER(F{r}),E{r}<>0),F{r}/E{r}-1,"")', nf=NF_P)
    put(ws, f"L{r}", ", ".join(refs), color=GREY)
SEG0 = r
seg_row("Water Technologies", None, ("Chiffre d'affaires Water Technologies", "FY2025"), None,
        ("EBITDA Water Technologies", "FY2025"), ("Croissance organique de l'EBITDA Water Technologies", "FY2025"),
        ("Chiffre d'affaires Water Technologies", "S1 2026"), ("EBITDA Water Technologies", "S1 2026")); r += 1
seg_row("Amériques, Asie-Pacifique, Afrique Moyen-Orient", ("Chiffre d'affaires Amériques", "FY2024"), ("Chiffre d'affaires Amériques", "FY2025"),
        ("EBITDA Amériques", "FY2024"), ("EBITDA Amériques", "FY2025"), ("Croissance organique de l'EBITDA Amériques", "FY2025"),
        ("Chiffre d'affaires Amériques", "S1 2026"), ("EBITDA Amériques", "S1 2026")); r += 1
seg_row("Europe", ("Chiffre d'affaires Europe", "FY2024"), ("Chiffre d'affaires Europe", "FY2025"),
        ("EBITDA Europe", "FY2024"), ("EBITDA Europe", "FY2025"), ("Croissance organique de l'EBITDA Europe", "FY2025"),
        ("Chiffre d'affaires Europe", "S1 2026"), None); r += 1
seg_row("France et Déchets dangereux Europe", ("Chiffre d'affaires France et Déchets dangereux Europe", "FY2024"),
        ("Chiffre d'affaires France et Déchets dangereux Europe", "FY2025"), ("EBITDA France et Déchets dangereux Europe", "FY2024"),
        ("EBITDA France et Déchets dangereux Europe", "FY2025"), None, None, ("EBITDA France et Déchets dangereux Europe", "S1 2026")); r += 1
SEG1 = r - 1
put(ws, f"B{r}", "Somme des quatre segments", bold=True)
for col in "CDEF":
    put(ws, f"{col}{r}", f"=SUM({col}{SEG0}:{col}{SEG1})", nf=NF_M, bold=True)
SG["sum25"] = f"{q('Segments')}!$F${r}"; r += 1
put(ws, f"B{r}", "EBITDA du groupe publié"); put(ws, f"E{r}", f"={eb24}", color=GREEN, nf=NF_M); put(ws, f"F{r}", f"={eb25}", color=GREEN, nf=NF_M)
put(ws, f"L{r}", f"{id_eb24}, {id_eb25}", color=GREY); r += 1
put(ws, f"B{r}", "Autres (holding, éliminations) = groupe − segments"); put(ws, f"E{r}", "n/d", color=GREY, align="right"); put(ws, f"F{r}", f"=F{r-1}-F{r-2}", nf=NF_M)
put(ws, f"L{r}", "2024 : l'EBITDA Water Technologies 2024 n'est pas au registre", color=GREY)
SG["other25"] = f"{q('Segments')}!$F${r}"; r += 2

put(ws, f"A{r}", "B", bold=True); put(ws, f"B{r}", "Les boosters (croissance et marge)", bold=True); r += 1
header(ws, r, ["", "Booster", "CA 2024", "CA 2025", "EBITDA 2024", "EBITDA 2025", "Marge 2025", "Croiss. organique CA 2025",
               "Croiss. organique EBITDA 2025", "CA S1 2026", "EBITDA S1 2026", "Réf."]); r += 1
def booster_row(label, rev24, rev25, eb24_, eb25_, grev, geb, revh1, ebh1):
    refs = []
    put(ws, f"B{r}", label)
    for col, spec, nf in (("C", rev24, NF_M), ("D", rev25, NF_M), ("E", eb24_, NF_M), ("F", eb25_, NF_M), ("H", grev, NF_P),
                          ("I", geb, NF_P), ("J", revh1, NF_M), ("K", ebh1, NF_M)):
        if spec is None:
            put(ws, f"{col}{r}", None); continue
        addr, fid = E(*spec)
        put(ws, f"{col}{r}", f"={addr}" + ("/100" if nf == NF_P else ""), color=GREEN, nf=nf); refs.append(fid)
    put(ws, f"G{r}", f'=IF(AND(ISNUMBER(D{r}),ISNUMBER(F{r}),D{r}<>0),F{r}/D{r},"")', nf=NF_P)
    put(ws, f"L{r}", ", ".join(refs), color=GREY)
BR0 = r
booster_row("Tous les boosters", None, ("Boosters : chiffre d'affaires", "FY2025"), None, ("Boosters : EBITDA", "FY2025"),
            ("Boosters : croissance organique du chiffre d'affaires", "FY2025"), ("Boosters : croissance organique de l'EBITDA", "FY2025"),
            ("Boosters : chiffre d'affaires", "S1 2026"), ("Boosters : EBITDA", "S1 2026")); SG["boostMargin"] = f"{q('Segments')}!$G${r}"; r += 1
booster_row("Bioénergie, flexibilité, efficacité énergétique", ("Bioénergie, flexibilité, efficacité énergétique : chiffre d'affaires", "FY2024"),
            ("Bioénergie, flexibilité, efficacité énergétique : chiffre d'affaires", "FY2025"),
            ("Bioénergie, flexibilité, efficacité énergétique : EBITDA", "FY2024"), ("Bioénergie, flexibilité, efficacité énergétique : EBITDA", "FY2025"),
            None, ("Bioénergie, flexibilité, efficacité énergétique : croissance organique de l'EBITDA", "S1 2026"),
            ("Bioénergie, flexibilité, efficacité énergétique : chiffre d'affaires", "S1 2026"), None); r += 1
put(ws, f"B{r}", "Déchets dangereux (pro forma avec Clean Earth, présentation Clean Earth p.9)")
put(ws, f"D{r}", f"={pf_rev}*1000", color=GREEN, nf=NF_M); put(ws, f"F{r}", f"={pf_eb}*1000", color=GREEN, nf=NF_M)
put(ws, f"G{r}", f"=F{r}/D{r}", nf=NF_P); put(ws, f"L{r}", f"{i_pfrev}, {i_pfeb} (2025E)", color=GREY); r += 1
put(ws, f"B{r}", "Marge publiée des boosters 2025 (pour contrôle)"); bm, i_bm = E("Boosters : marge d'EBITDA", "FY2025")
put(ws, f"G{r}", f"={bm}/100", color=GREEN, nf=NF_P); put(ws, f"L{r}", i_bm, color=GREY); SG["boostMarginPub"] = f"{q('Segments')}!$G${r}"; r += 1
b23, i_b23 = E("GreenUp : chiffre d'affaires bioénergie, flexibilité, efficacité énergétique (base)", "2023")
b27, i_b27 = E("GreenUp : ambition de chiffre d'affaires bioénergie", "2027")
put(ws, f"B{r}", "Ambition GreenUp bioénergie : CA 2023 → 2027 (Md EUR), et rythme annuel requis")
put(ws, f"C{r}", f"={b23}", color=GREEN, nf=NF_D2); put(ws, f"D{r}", f"={b27}", color=GREEN, nf=NF_D2)
put(ws, f"H{r}", f"=(D{r}/C{r})^(1/4)-1", nf=NF_P); put(ws, f"L{r}", f"{i_b23}, {i_b27} — rythme requis sur 4 ans", color=GREY); r += 1
put(ws, f"B{r}", "Rythme constaté 2023 → 2025 (périmètre courant)"); put(ws, f"H{r}", f"=(D{BR0+1}/1000/C{r-1})^(1/2)-1", nf=NF_P)
put(ws, f"L{r}", "CA 2025 / base 2023, sur deux ans — effets de périmètre compris", color=GREY); r += 2

put(ws, f"A{r}", "C", bold=True); put(ws, f"B{r}", "D'où viennent les 8 Md€ : si chaque segment gardait son rythme organique de 2025", bold=True); r += 1
header(ws, r, ["", "Segment", "EBITDA 2025", "Rythme organique 2025", "EBITDA 2027 au même rythme", "", "", "", "", "", "", "Note"]); r += 1
C0 = r
for i, rr in enumerate(range(SEG0, SEG1 + 1)):
    put(ws, f"B{r}", f"=B{rr}", color=BLACK)
    put(ws, f"C{r}", f"=F{rr}", nf=NF_M)
    put(ws, f"D{r}", f'=IF(ISNUMBER(I{rr}),I{rr},H{rr})', nf=NF_P)
    put(ws, f"E{r}", f'=IF(AND(ISNUMBER(C{r}),ISNUMBER(D{r})),C{r}*(1+D{r})^2,"")', nf=NF_M)
    put(ws, f"L{r}", "croissance publiée à défaut d'organique" if i == 3 else "", color=GREY)
    r += 1
put(ws, f"B{r}", "Autres (holding), supposés stables"); put(ws, f"C{r}", f"={SG['other25']}", nf=NF_M); put(ws, f"E{r}", f"=C{r}", nf=NF_M); r += 1
put(ws, f"B{r}", "Clean Earth en année pleine (hypothèse ceEb, convertie)"); put(ws, f"E{r}", f"={HY('ceEb')}/{HY('fx')}", color=GREEN, nf=NF_M); r += 1
put(ws, f"B{r}", "EBITDA 2027 « au rythme de 2025 »", bold=True); put(ws, f"E{r}", f"=SUM(E{C0}:E{r-1})", nf=NF_M, bold=True)
SG["eb27seg"] = f"{q('Segments')}!$E${r}"; r += 1
put(ws, f"B{r}", "Objectif GreenUp ≥"); put(ws, f"E{r}", f"={tgt8}*1000", color=GREEN, nf=NF_M); r += 1
put(ws, f"B{r}", "EBITDA 2027 du modèle (Trajectoire, central)"); put(ws, f"E{r}", f"={T['EB27']}", color=GREEN, nf=NF_M); r += 1
put(ws, f"B{r}", "Lecture : les Amériques (+9 %) et Water Technologies (+14 %) portent la croissance ; l'Europe (+2 %) et la France-DD "
    "(+6 %) non. Prolonger 2025 deux ans mène près de l'objectif, avant cessions : c'est ce qui fonde l'hypothèse g27 du modèle.",
    color=GREY, italic=True, wrap=True)
ws.merge_cells(f"B{r}:L{r}"); ws.row_dimensions[r].height = 30; r += 2

put(ws, f"A{r}", "D", bold=True); put(ws, f"B{r}", "Les KPI environnementaux de GreenUp face à 2027", bold=True); r += 1
header(ws, r, ["", "KPI", "Base 2023", "2024", "2025", "Objectif 2027", "Écart 2025 → objectif", "", "", "", "", "Réf."]); r += 1
co23, i_co23 = E("KPI GreenUp : émissions de GES évitées (scope 4), base 2023", "2023")
co24, i_co24 = E("KPI GreenUp : émissions de GES évitées, progression vs 2023", "2024")
co25, i_co25 = E("KPI GreenUp : émissions de GES évitées, progression vs 2023", "2025")
cot, i_cot = E("KPI GreenUp : émissions de GES évitées, objectif vs 2023", "2027")
co18, i_co18 = E("KPI GreenUp : émissions de GES évitées, objectif", "2027", "Mt CO2e")
put(ws, f"B{r}", "CO2e évité (scope 4), Mt — reconstitué : base × (1 + progression)")
put(ws, f"C{r}", f"={co23}", color=GREEN, nf=NF_D2); put(ws, f"D{r}", f"=C{r}*(1+{co24}/100)", nf=NF_D2)
put(ws, f"E{r}", f"=C{r}*(1+{co25}/100)", nf=NF_D2); put(ws, f"F{r}", f"=C{r}*(1+{cot}/100)", nf=NF_D2)
put(ws, f"G{r}", f"=F{r}-E{r}", nf=NF_D2); put(ws, f"L{r}", f"{i_co23}, {i_co24}, {i_co25}, {i_cot}", color=GREY)
SG["co2target"] = f"{q('Segments')}!$F${r}"; r += 1
put(ws, f"B{r}", "CO2e évité : objectif tel qu'énoncé (18 Mt)"); put(ws, f"F{r}", f"={co18}", color=GREEN, nf=NF_D2)
put(ws, f"L{r}", f"{i_co18} — 13,45 × 1,30 = 17,5, pas 18 : deux formulations", color=GREY); r += 1
w24, i_w24 = E("KPI GreenUp : eau douce économisée", "2024"); w25, i_w25 = E("KPI GreenUp : eau douce économisée", "2025")
wt, i_wt = E("KPI GreenUp : eau douce économisée, objectif", "2027")
put(ws, f"B{r}", "Eau douce économisée, Md m³"); put(ws, f"D{r}", f"={w24}", color=GREEN, nf='0.000'); put(ws, f"E{r}", f"={w25}", color=GREEN, nf='0.000')
put(ws, f"F{r}", f"={wt}", color=GREEN, nf='0.000'); put(ws, f"G{r}", f"=F{r}-E{r}", nf='0.000'); put(ws, f"L{r}", f"{i_w24}, {i_w25}, {i_wt} — déjà dépassé", color=GREY); r += 1
put(ws, f"B{r}", "Déchets dangereux traités, kt"); put(ws, f"D{r}", f"={hw24}", color=GREEN, nf=NF_M); put(ws, f"E{r}", f"={hw25}", color=GREEN, nf=NF_M)
put(ws, f"F{r}", f"={hwt}*1000", color=GREEN, nf=NF_M); put(ws, f"G{r}", f"=F{r}-E{r}", nf=NF_M); put(ws, f"L{r}", f"{i_hw24}, {i_hw25}, {i_hwt} — déjà atteint", color=GREY); r += 1
put(ws, f"B{r}", "Lecture : deux KPI sur trois sont déjà atteints ; seul le CO2 évité reste en chemin (16,6 → 17,5-18 Mt). "
    "Pour le périmètre capacité, la question 1 se réduit à l'EBITDA.", color=GREY, italic=True, wrap=True)
ws.merge_cells(f"B{r}:L{r}"); ws.row_dimensions[r].height = 30; r += 1
widths(ws, {"A": 4, "B": 58, "C": 11, "D": 11, "E": 11, "F": 11, "G": 10, "H": 12, "I": 13, "J": 11, "K": 11, "L": 46})

# ================================================================ Pont EBITDA
ws = sheets["Pont EBITDA"]
title(ws, "Rôles 3 et 4 (question 1 du cours) — Le pont de l'EBITDA : ce que GreenUp promet, ce que 2025 a livré, ce que le modèle projette",
      "GreenUp (p.58, p.69), présentation des résultats 2025 (p.21, p.24), Trajectoire. Une marche = une ligne ; chaque marche "
      "renvoie à une entrée ou à une hypothèse.")
PB = {}
r = 4
put(ws, f"A{r}", "A", bold=True); put(ws, f"B{r}", "Ce que le plan promet (GreenUp 2024-2027)", bold=True); r += 1
header(ws, r, ["", "Libellé", "Valeur", "Unité", "Réf. / calcul"]); r += 1
gu_base, i_gubase = E("GreenUp : EBITDA de base", "2023")
gu_cagr, i_gucagr = E("GreenUp : croissance annuelle moyenne de l'EBITDA", "2023-2027")
gu_eff, i_gueff = E("GreenUp : gains d'efficacité par an", "2024-2027")
gu_syn, i_gusyn = E("GreenUp : synergies de coûts attendues", "2024-2025")
gu_syncum, i_gusyncum = E("GreenUp : synergies Suez cumulées 2022-2025", "2022-2025")
line(ws, r, "base23", "EBITDA de base 2023 (GreenUp)", f"={gu_base}*1000", "M EUR", i_gubase, store=PB); r += 1
line(ws, r, "tgt27", "Objectif 2027 : au moins", f"={tgt8}*1000", "M EUR", i_tgt8, store=PB); r += 1
line(ws, r, "gap", "Écart à combler en quatre ans", f"=C{r-1}-C{r-2}", "M EUR", "objectif − base", NF_M, True, store=PB); r += 1
line(ws, r, "cagrC", "Croissance annuelle implicite 2023 → 2027", f"=(C{r-2}/C{r-3})^(1/4)-1", "%", "(objectif / base)^(1/4) − 1", NF_P, store=PB); r += 1
line(ws, r, "cagrP", "Croissance annuelle annoncée, environ", f"={gu_cagr}/100", "%", i_gucagr, NF_P, store=PB); r += 1
line(ws, r, "eff", "Gains d'efficacité promis, par an", f"={gu_eff}", "M EUR", i_gueff, store=PB); r += 1
line(ws, r, "eff4", "… sur quatre ans (2024-2027)", f"=4*C{r-1}", "M EUR", "4 × 350", NF_M, True, store=PB); r += 1
line(ws, r, "effShare", "Part de l'écart couverte par la seule efficacité", f"=C{r-1}/{PB['gap']}", "%", "efficacité × 4 / écart", NF_P, True, store=PB); r += 1
line(ws, r, "syn2425", "Synergies de coûts Suez attendues en 2024-2025, environ", f"={gu_syn}", "M EUR", i_gusyn, store=PB); r += 1
line(ws, r, "syncum", "Synergies Suez cumulées 2022-2025 confirmées par le plan", f"={gu_syncum}", "M EUR", f"{i_gusyncum} ; réalisé 2025 : voir §B", store=PB); r += 1
line(ws, r, "effSynShare", "Efficacité × 4 + synergies 2024-2025, en part de l'écart", f"=({PB['eff4']}+{PB['syn2425']})/{PB['gap']}", "%", "", NF_P, True, store=PB); r += 1
put(ws, f"B{r}", "Lecture : GreenUp est d'abord un plan de coûts. L'efficacité promise (350 M€ par an) couvre à elle seule l'essentiel des 1,5 Md€ "
    "entre 2023 et 2027 ; la croissance du chiffre d'affaires et les boosters apportent le reste. Un plan de coûts se vérifie année par année (§B).",
    color=GREY, italic=True, wrap=True)
ws.merge_cells(f"B{r}:E{r}"); ws.row_dimensions[r].height = 44; r += 2

put(ws, f"A{r}", "B", bold=True); put(ws, f"B{r}", "Ce que 2024 et 2025 ont livré (présentations des résultats)", bold=True); r += 1
header(ws, r, ["", "Libellé", "2025", "2024", "Réf. / calcul"]); r += 1
eff25, i_eff25 = E("Gains d'efficacité annuels", "FY2025")
eff24, i_eff24 = E("Gains d'efficacité (et non synergies)", "FY2024")
org25, i_org25 = E("Croissance organique de l'EBITDA (groupe)", "FY2025")
syn24, i_syn24 = E("Synergies Suez cumulées", "FY2024")
syn25, i_syn25 = E("Synergies Suez cumulées réalisées", "FY2025")
marg25, i_marg25 = E("Marge d'EBITDA (groupe)", "FY2025")
sh_eb, i_sheb = E("Strongholds : EBITDA", "FY2025")
sh_g, i_shg = E("Strongholds : croissance organique de l'EBITDA", "FY2025")
bo_eb, i_boeb = E("Boosters : EBITDA", "FY2025")
bo_g, i_bog = E("Boosters : croissance organique de l'EBITDA", "FY2025")


def two(key, label, f25, f24, ref, nf=NF_M, bold=False):
    global r
    put(ws, f"B{r}", label, bold=bold)
    for col, f in (("C", f25), ("D", f24)):
        if f is None:
            put(ws, f"{col}{r}", "n/d", color=GREY, align="right")
        else:
            put(ws, f"{col}{r}", f, nf=nf, bold=bold, color=GREEN if (str(f).startswith("=") and "!" in str(f) and re.fullmatch(r"=[^+\-*/()]+(/100)?", str(f))) else BLACK)
    put(ws, f"E{r}", ref, color=GREY)
    PB[key] = f"{q('Pont EBITDA')}!$C${r}"; PB[key + "_24"] = f"{q('Pont EBITDA')}!$D${r}"
    r += 1


two("ebY", "EBITDA publié", f"={eb25}", f"={eb24}", f"{id_eb25}, {id_eb24}")
two("ebY1", "EBITDA de l'année précédente, tel que publié", f"={eb24}", f"={PB['base23']}", f"{id_eb24} ; 2023 : base GreenUp (6,5 Md€, arrondi)")
two("orgPct", "Croissance organique de l'EBITDA publiée", f"={org25}/100", None, i_org25, NF_P)
two("orgM", "… en M EUR, sur la base publiée de l'année précédente (approximation : la base organique n'est pas publiée)",
    f"=C{r-1}*D{r-3}", None, "croissance × EBITDA N−1 publié", NF_M, True)
two("effY", "Gains d'efficacité de l'année", f"={eff25}", f"={eff24}", f"{i_eff25}, {i_eff24}")
two("synInc", "Synergies Suez supplémentaires dans l'année (cumul N − cumul N−1)", f"={syn25}-{syn24}", None, f"{i_syn25} − {i_syn24}")
two("effSynY", "Efficacité + synergies de l'année", f"=C{r-2}+C{r-1}", None, "", NF_M, True)
two("effShareY", "Part de la croissance organique expliquée par l'efficacité seule", f"=C{r-3}/{PB['orgM']}", None, "efficacité / croissance organique en M EUR", NF_P, True)
two("rest", "Reste : volumes, prix, énergie, inflation des coûts (croissance organique − efficacité − synergies)",
    f"={PB['orgM']}-C{r-2}", None, "négatif = hors efficacité et synergies, l'EBITDA organique recule", NF_M, True)
two("margY", "Marge d'EBITDA publiée", f"={marg25}/100", None, i_marg25, NF_P)
put(ws, f"B{r}", "Contre-épreuve par les familles d'activités (présentation 2025, p.24 et p.7)", bold=True); r += 1
two("shOrg", "Strongholds : EBITDA × croissance organique", f"={sh_eb}*{sh_g}/100", None, f"{i_sheb} × {i_shg}")
two("boOrg", "Boosters : EBITDA × croissance organique", f"={bo_eb}*{bo_g}/100", None, f"{i_boeb} × {i_bog}")
two("famOrg", "Somme des deux familles (sur les EBITDA 2025, approximation)", f"=C{r-2}+C{r-1}", None, "à comparer à la croissance organique du groupe en M EUR", NF_M, True)
put(ws, f"B{r}", "Lecture : en 2025, les gains d'efficacité (399 M€) pèsent à peu près autant que toute la croissance organique de l'EBITDA ; "
    "avec les synergies Suez, ils la dépassent. Hors ces deux programmes, l'EBITDA organique ne progresse pas : volumes, prix et énergie "
    "absorbent l'inflation des coûts, pas plus. La capacité d'acquisition de 2027 repose donc sur un programme de coûts qui doit être "
    "livré deux années de plus (§D).", color=GREY, italic=True, wrap=True)
ws.merge_cells(f"B{r}:E{r}"); ws.row_dimensions[r].height = 58; r += 2

put(ws, f"A{r}", "C", bold=True); put(ws, f"B{r}", "Le pont du modèle 2025 → 2027, marche par marche (Trajectoire)", bold=True); r += 1
header(ws, r, ["", "Marche", "Scénario actif", "Défavorable combiné", "Favorable combiné", "Ce qui la porte"]); r += 1
TJ = q('Trajectoire')


def tref(col, key):
    return f"{TJ}!${col}${ROW[key]}"


def href(col, code):
    return f"{TJ}!${col}${IN[code]}"


STEPS = [
    ("eb25", "EBITDA 2025 publié", lambda c: f"={eb25}", id_eb25, True),
    ("org26", "+ croissance organique 2026 (hypothèse g26)", lambda c: f"={tref(c, 'A26')}-{eb25}", "A26 − EBITDA 2025", False),
    ("ce26", "+ Clean Earth 2026 (mois consolidés, converti)", lambda c: f"={tref(c, 'CE26')}", "CE26", False),
    ("disp26", "− EBITDA cédé en 2026", lambda c: f"={tref(c, 'DE26')}", "DE26", False),
    ("eb26", "EBITDA 2026 (sous-total)", lambda c: f"={tref(c, 'EB26')}", "EB26", True),
    ("org27", "+ croissance organique 2027 (hypothèse g27, sur la base 2026 Clean Earth en année pleine)",
     lambda c: f"={tref(c, 'B27')}-({tref(c, 'A26')}+{tref(c, 'DE26')}+{href(c, 'ceEb')}/{href(c, 'fx')})", "B27 − base", False),
    ("ce27", "+ complément Clean Earth (12 mois − mois 2026)", lambda c: f"={href(c, 'ceEb')}/{href(c, 'fx')}-{tref(c, 'CE26')}", "ceEb / fx − CE26", False),
    ("syn27", "+ synergies Clean Earth réalisées en 2027", lambda c: f"={tref(c, 'SYN27')}", "SYN27", False),
    ("tuck27", "+ EBITDA apporté par les tuck-ins 2027", lambda c: f"={tref(c, 'TE27')}", "TE27", False),
    ("disp27", "− EBITDA cédé en 2027", lambda c: f"={tref(c, 'DE27')}", "DE27", False),
    ("eb27", "EBITDA 2027 (somme des marches)", None, "somme", True),
    ("eb27T", "EBITDA 2027 de Trajectoire (contrôle)", lambda c: f"={tref(c, 'EB27')}", "EB27", False),
    ("tgt", "Objectif GreenUp : au moins", lambda c: f"={tgt8}*1000", i_tgt8, False),
    ("gap8", "Écart à l'objectif", lambda c: f"={tref(c, 'GAP8')}", "GAP8", True),
]
S0 = r
for key, label, fn, why, bold in STEPS:
    put(ws, f"B{r}", label, bold=bold)
    for col, tcol in (("C", "D"), ("D", UNFAV), ("E", FAV)):
        if fn is None:
            put(ws, f"{col}{r}", f"={col}{PB['eb25_row']}+SUM({col}{PB['eb25_row'] + 1}:{col}{PB['eb25_row'] + 3})+SUM({col}{PB['eb25_row'] + 5}:{col}{PB['eb25_row'] + 9})", nf=NF_M, bold=True)
        else:
            f = fn(tcol)
            put(ws, f"{col}{r}", f, nf=NF_M, bold=bold, color=GREEN if re.fullmatch(r"=[^+\-*/()]+", f) else BLACK)
    put(ws, f"F{r}", why, color=GREY)
    PB[key] = f"{q('Pont EBITDA')}!$C${r}"
    if key == "eb25":
        PB["eb25_row"] = r
    if key == "eb27":
        PB["eb27_row"] = r
    r += 1
S1 = r - 1
put(ws, f"B{r}", "Lecture : la colonne « Scénario actif » suit le sélecteur de Hypothèses. Les marches positives sont la croissance organique "
    "(dont l'efficacité, §D), Clean Earth et ses synergies, les tuck-ins ; les cessions retirent leur EBITDA sur l'année entière (prudent).",
    color=GREY, italic=True, wrap=True)
ws.merge_cells(f"B{r}:F{r}"); ws.row_dimensions[r].height = 30; r += 1
pch = BarChart(); pch.type = "col"; pch.grouping = "clustered"
pch.title = "Les marches du pont 2025 → 2027, scénario actif (M EUR)"
pch.y_axis.title = "M EUR"; pch.y_axis.delete = False; pch.x_axis.delete = False
pch.add_data(Reference(ws, min_col=3, min_row=S0 + 1, max_row=S0 + 9), titles_from_data=False)
pch.set_categories(Reference(ws, min_col=2, min_row=S0 + 1, max_row=S0 + 9))
pch.legend = None
pch.height, pch.width = 9, 22
ws.add_chart(pch, f"H{S0}")
r += 1

put(ws, f"A{r}", "D", bold=True); put(ws, f"B{r}", "L'efficacité porte tout : que vaut la marge 2027 si le programme de coûts n'est pas livré ?", bold=True); r += 1
header(ws, r, ["", "Libellé", "Valeur", "Unité", "Réf. / calcul"]); r += 1
line(ws, r, "effPlan", "Efficacité promise par an (GreenUp)", f"={gu_eff}", "M EUR", i_gueff, store=PB); r += 1
line(ws, r, "org26M", "Croissance organique 2026 du modèle (scénario actif)", f"={PB['org26']}", "M EUR", "§C", store=PB); r += 1
line(ws, r, "effIn26", "Part de l'organique 2026 que représente l'efficacité promise", f"={PB['effPlan']}/{PB['org26M']}", "%", "350 / organique 2026", NF_P, True, store=PB); r += 1
line(ws, r, "org27M", "Croissance organique 2027 du modèle (scénario actif)", f"={PB['org27']}", "M EUR", "§C", store=PB); r += 1
line(ws, r, "effIn27", "Part de l'organique 2027 que représente l'efficacité promise", f"={PB['effPlan']}/{PB['org27M']}", "%", "350 / organique 2027", NF_P, True, store=PB); r += 1
put(ws, f"B{r}", "Taux de réalisation de l'efficacité en 2026 et 2027 (1 = livrée comme promise)", bold=True)
put(ws, f"C{r}", 1, color=BLUE, nf=NF_P, fill=YELLOW); put(ws, f"D{r}", "%", color=GREY); put(ws, f"E{r}", "cellule jaune : à faire varier (0 = rien, 0,5 = la moitié)", color=GREY)
PB["kEff"] = f"{q('Pont EBITDA')}!$C${r}"; KROW = r; r += 1
line(ws, r, "effLost", "EBITDA 2027 perdu si l'efficacité n'est livrée qu'à ce taux (deux années)", f"=-(1-{PB['kEff']})*2*{PB['effPlan']}", "M EUR", "−(1 − taux) × 2 × 350", NF_M, True, store=PB); r += 1
line(ws, r, "eb27k", "EBITDA 2027 au taux choisi", f"={T['EB27']}+{PB['effLost']}", "M EUR", "EB27 + perte", NF_M, True, store=PB); r += 1
line(ws, r, "lev27k", "Levier fin 2027 au taux choisi (dette inchangée)", f"={T['NFD27']}/{PB['eb27k']}", "x", "DFN 2027 / EBITDA 2027 corrigé", NF_X, True, store=PB); r += 1
line(ws, r, "head27k", "Marge sous 3x fin 2027 au taux choisi", f"={HY('cap')}*{PB['eb27k']}-{T['NFD27']}", "M EUR", "plafond × EBITDA corrigé − DFN", NF_M, True, store=PB); r += 1
line(ws, r, "head27", "Pour mémoire : marge sous 3x fin 2027, scénario actif", f"={T['HEAD27']}", "M EUR", "Trajectoire", NF_M, store=PB); r += 1
put(ws, f"B{r}", "Grille : marge sous 3x fin 2027 selon le taux de réalisation de l'efficacité", bold=True); r += 1
header(ws, r, ["", "Taux de réalisation", "EBITDA 2027", "Levier 2027", "Marge sous 3x (M EUR)"]); r += 1
G0 = r
for k in (0, 0.25, 0.5, 0.75, 1):
    put(ws, f"B{r}", k, color=BLUE, nf=NF_P, align="center")
    put(ws, f"C{r}", f"={T['EB27']}-(1-B{r})*2*{PB['effPlan']}", nf=NF_M)
    put(ws, f"D{r}", f"={T['NFD27']}/C{r}", nf=NF_X)
    put(ws, f"E{r}", f"={HY('cap')}*C{r}-{T['NFD27']}", nf=NF_M, bold=True)
    if k == 0:
        PB["head27k0"] = f"{q('Pont EBITDA')}!$E${r}"
    r += 1
put(ws, f"B{r}", "Lecture : chaque euro d'efficacité non livré coûte trois euros de capacité d'endettement (plafond × EBITDA) et ne rapporte rien "
    "en cash-flow en face. Sans efficacité en 2026-2027, la marge sous 3x disparaît : la question à poser à Veolia le 16 octobre est "
    "moins « combien achèterez-vous ? » que « les 350 M€ d'efficacité par an sont-ils sécurisés ? ».", color=GREY, italic=True, wrap=True)
ws.merge_cells(f"B{r}:F{r}"); ws.row_dimensions[r].height = 44; r += 1
r += 1
put(ws, f"A{r}", "E", bold=True); put(ws, f"B{r}", "Veolia tient-elle ses plans ? Objectifs annoncés et réalisés (Impact 2023, guidance 2025)", bold=True); r += 1
header(ws, r, ["", "Objectif", "Visé", "Réalisé", "Tenu ?", "Réf."]); r += 1
CR = {}
C0 = r
for lab, v_pre, v_per, r_pre, r_per, kind, scale in (
        ("Impact 2023 : EBITDA 2023 (Md EUR, bas de fourchette)", "Impact 2023 : EBITDA 2023 visé, bas", "2023", "Impact 2023 : EBITDA 2023 réalisé", "FY2023", "min", 1),
        ("Impact 2023 : résultat net courant 2023 (M EUR)", "Impact 2023 : résultat net courant 2023 visé", "2023", "Impact 2023 : résultat net courant 2023 réalisé", "FY2023", "min", 1000),
        ("Impact 2023 : CA déchets liquides et dangereux (Md EUR)", "Impact 2023 : chiffre d'affaires déchets liquides et dangereux visé", "2023", "Impact 2023 : chiffre d'affaires déchets liquides et dangereux réalisé", "FY2023", "min", 1),
        ("Impact 2023 : émissions évitées (Mt CO2eq)", "Impact 2023 : émissions évitées visées", "2023", "Impact 2023 : émissions évitées réalisées", "FY2023", "min", 1),
        ("Impact 2023 : plastiques transformés (kt)", "Impact 2023 : plastiques transformés visés", "2023", "Impact 2023 : plastiques transformés réalisés", "FY2023", "min", 1),
        ("Impact 2023 : femmes parmi les 500 cadres dirigeants nommés (%)", "Impact 2023 : part de femmes nommées parmi les 500 cadres dirigeants, visée", "2020-2023", "Impact 2023 : part de femmes nommées parmi les 500 cadres dirigeants, réalisée", "2020-2023", "min", 1),
        ("Guidance 2025 : croissance organique de l'EBITDA (%, bas de fourchette)", "Guidance 2025 : croissance organique de l'EBITDA, bas", "2025", "Croissance organique de l'EBITDA (groupe)", "FY2025", "min", 1),
        ("Guidance 2025 : gains d'efficacité (M EUR, plus de)", "Guidance 2025 : gains d'efficacité, plus de", "2025", "Gains d'efficacité annuels", "FY2025", "min", 1),
        ("Guidance 2025 : synergies cumulées 2022-2025 (M EUR)", "Guidance 2025 : synergies cumulées 2022-2025 visées", "2025", "Synergies Suez cumulées réalisées", "FY2025", "min", 1),
        ("Guidance 2025 : croissance du résultat net courant (%, environ)", "Guidance 2025 : croissance du résultat net courant", "2025", "Résultat net courant 2025 : croissance réalisée", "FY2025", "min", 1),
        ("Guidance 2025 : levier < 3x", "Engagement de levier du groupe : ≤", "2027", "Ratio de levier publié", "FY2025", "max", 1)):
    v, iv = E(v_pre, v_per); rr_, ir = E(r_pre, r_per)
    put(ws, f"B{r}", lab)
    put(ws, f"C{r}", f"={v}*{scale}" if scale != 1 else f"={v}", color=GREEN, nf=NF_D2 if scale == 1 else NF_M)
    put(ws, f"D{r}", f"={rr_}", color=GREEN, nf=NF_D2)
    put(ws, f"E{r}", f'=IF(D{r}{">=" if kind == "min" else "<="}C{r},"oui","non")', bold=True, align="center")
    put(ws, f"F{r}", f"{iv}, {ir}", color=GREY)
    r += 1
C1 = r - 1
line(ws, r, "credHit", "Objectifs tenus (sur ceux du tableau)", f'=COUNTIF(E{C0}:E{C1},"oui")&" / "&COUNTA(E{C0}:E{C1})', "", "", None, True, store=CR); ws[f"C{r}"].alignment = Alignment(horizontal="right"); r += 1
line(ws, r, "credFin", "Objectifs financiers tenus", f'=COUNTIF(E{C0}:E{C0+2},"oui")+COUNTIF(E{C0+6}:E{C1},"oui")&" / "&(COUNTA(E{C0}:E{C0+2})+COUNTA(E{C0+6}:E{C1}))', "", "EBITDA, résultat net, CA déchets dangereux, guidance 2025", None, True, store=CR); ws[f"C{r}"].alignment = Alignment(horizontal="right"); r += 1
put(ws, f"B{r}", "Lecture : sur les objectifs financiers, Veolia a tenu tout ce qu'elle a annoncé depuis 2020 (Impact 2023 dépassé, guidance 2025 dépassée) ; "
    "les objectifs manqués sont non financiers (plastiques, mixité). Cela fonde le poids du scénario central sur le favorable pour la "
    "guidance 2026 — et n'enlève rien à la fragilité propre à l'efficacité (§D) : un plan tenu cinq ans de suite n'est pas tenu la sixième par décret.",
    color=GREY, italic=True, wrap=True)
ws.merge_cells(f"B{r}:F{r}"); ws.row_dimensions[r].height = 44; r += 1
widths(ws, {"A": 4, "B": 76, "C": 15, "D": 15, "E": 15, "F": 30})
ws.freeze_panes = "C4"

# ================================================================ Booster
ws = sheets["Booster"]
title(ws, "Rôle 4 — Où en est le booster Déchets dangereux face à 2027",
      "Volumes (URD 2025 p.226, GreenUp p.9), profil financier (présentation Clean Earth), moteur organique (CP S1 2026 p.7).")
B = {}
r = 4
put(ws, f"A{r}", "A", bold=True); put(ws, f"B{r}", "Volumes traités", bold=True); r += 1
hw24, i_hw24 = E("Déchets dangereux traités", "2024")
hw25, i_hw25 = E("Déchets dangereux traités", "2025")
hwt, i_hwt = E("Objectif : déchets dangereux traités en 2027", "2027")
hwt0, i_hwt0 = E("Objectif GreenUp initial", "2027")
c430, i_430 = E("Capacité ajoutée par 5 usines", "en construction")
line(ws, r, None, "Traités en 2024", f"={hw24}", "kt", i_hw24, store=B); r += 1
line(ws, r, "hw25", "Traités en 2025", f"={hw25}", "kt", i_hw25, store=B); r += 1
line(ws, r, None, "Variation 2024 → 2025", f"=C{r-1}/C{r-2}-1", "%", "", NF_P, store=B); r += 1
line(ws, r, "tgt", "Objectif 2027 (URD 2025)", f"={hwt}*1000", "kt", i_hwt, store=B); r += 1
line(ws, r, "tgt0", "Objectif 2027 initial (GreenUp 2024, « déchets dangereux et polluants »)", f"={hwt0}*1000", "kt",
     i_hwt0, store=B); r += 1
line(ws, r, None, "Écart à l'objectif actuel", f"=C{r-4}-C{r-2}", "kt", "atteint dès 2024", NF_M, True, store=B); r += 1
line(ws, r, "gap0", "Écart à l'objectif initial", f"=C{r-5}-C{r-2}", "kt", "", NF_M, True, store=B); r += 1
line(ws, r, None, "Capacité en construction (5 usines)", f"={c430}", "kt", i_430, store=B); r += 1
line(ws, r, None, "Écart à l'objectif initial, capacité nouvelle comprise", f"=C{r-2}+C{r-1}", "kt", "", NF_M, store=B)
r += 1
put(ws, f"B{r}", "Lecture : l'objectif de volume a été ramené de 10 Mt à 9 Mt, et le périmètre a changé (« et polluants » a "
    "disparu). À expliquer à l'oral avant toute comparaison.", color=GREY, italic=True, wrap=True)
ws.merge_cells(f"B{r}:E{r}"); ws.row_dimensions[r].height = 30; r += 2

put(ws, f"A{r}", "B", bold=True); put(ws, f"B{r}", "Profil financier et ambition", bold=True); r += 1
hwrev24, i_hwrev24 = E("Déchets dangereux Veolia : chiffre d'affaires", "2024")
pf_rev, i_pfrev = E("Déchets dangereux Veolia + Clean Earth : chiffre d'affaires", "2025E")
pf_eb, i_pfeb = E("Déchets dangereux Veolia + Clean Earth : EBITDA", "2025E")
pf_m, i_pfm = E("Déchets dangereux Veolia + Clean Earth : marge", "2025E")
cagr, i_cagr = E("Ambition GreenUp relevée", "2024-2027")
ce_r25, i_cer25 = E("Clean Earth : chiffre d'affaires", "2025E")
ce_r26, i_cer26 = E("Clean Earth : chiffre d'affaires", "2026E")
line(ws, r, None, "Chiffre d'affaires DD de Veolia 2024", f"={hwrev24}", "Md EUR", i_hwrev24, NF_D2, store=B); r += 1
line(ws, r, None, "Chiffre d'affaires DD pro forma 2025E (avec Clean Earth)", f"={pf_rev}", "Md EUR", i_pfrev, NF_D2, store=B)
r += 1
line(ws, r, "pfEb", "EBITDA DD pro forma 2025E", f"={pf_eb}", "Md EUR", i_pfeb, NF_D2, store=B); r += 1
line(ws, r, None, "Marge pro forma recalculée", f"=C{r-1}/C{r-2}", "%", f"publiée : {i_pfm}", NF_P, store=B); r += 1
line(ws, r, "cagr", "Ambition : croissance annuelle de l'EBITDA DD 2024-2027, plus de", f"={cagr}/100", "%", i_cagr, NF_P,
     store=B); r += 1
line(ws, r, None, "Illustration : EBITDA DD 2027 si +10 %/an depuis le pro forma 2025E", f"=C{r-3}*(1+C{r-1})^2", "Md EUR",
     "base pro forma × 1,1² — illustratif", NF_D2, True, store=B); r += 1
line(ws, r, None, "Clean Earth : croissance du CA 2025E → 2026E", f"={ce_r26}/{ce_r25}-1", "%", f"{i_cer25}, {i_cer26}",
     NF_P, store=B); r += 1
line(ws, r, None, "Clean Earth : marge d'EBITDA 2026E", f"={ce_eb}/{ce_r26}", "%", f"{id_ceeb}, {i_cer26}", NF_P, store=B)
r += 2

put(ws, f"A{r}", "C", bold=True); put(ws, f"B{r}", "Le moteur organique (S1 2026)", bold=True); r += 1
eu_rev, i_eurev = E("Chiffre d'affaires Déchets dangereux Europe", "S1 2026")
eu_g, i_eug = E("Déchets dangereux Europe : croissance", "S1 2026")
fh25, i_fh25 = E("EBITDA France et Déchets dangereux Europe", "S1 2025")
fh26, i_fh26 = E("EBITDA France et Déchets dangereux Europe", "S1 2026")
am25, i_am25 = E("EBITDA Amériques", "S1 2025")
am26, i_am26 = E("EBITDA Amériques", "S1 2026")
line(ws, r, None, "CA Déchets dangereux Europe S1 2026", f"={eu_rev}", "M EUR", i_eurev, store=B); r += 1
line(ws, r, "euG", "… croissance à périmètre et change constants", f"={eu_g}/100", "%", i_eug, NF_P, True, store=B); r += 1
line(ws, r, None, "EBITDA France et DD Europe, S1 2026 / S1 2025 (publié)", f"={fh26}/{fh25}-1", "%", f"{i_fh25}, {i_fh26}",
     NF_P, store=B); r += 1
line(ws, r, None, "EBITDA Amériques-Asie-Afrique, S1 2026 / S1 2025 (publié)", f"={am26}/{am25}-1", "%",
     f"{i_am25}, {i_am26}", NF_P, store=B); r += 1
put(ws, f"B{r}", "Lecture : en Europe le booster est à l'arrêt en organique (−0,2 %). Les +10 %/an reposent sur Clean Earth "
    "et les États-Unis : c'est de la croissance achetée.", color=GREY, italic=True, wrap=True)
ws.merge_cells(f"B{r}:E{r}"); ws.row_dimensions[r].height = 30; r += 2

put(ws, f"A{r}", "D", bold=True); put(ws, f"B{r}", "Comparables — cadre à remplir (phase D, sources à verser dans la dataroom)",
                                       bold=True); r += 1
header(ws, r, ["", "Société", "Pays", "Part DD du CA", "CA (M)", "EBITDA (M)", "Marge", "VE / EBITDA", "Fichier", "Page",
               "Pourquoi comparable (ou pas)"]); r += 1
peers = [("Clean Harbors", "États-Unis", "TSDF, incinération ; leader US, concurrent direct de Clean Earth"),
         ("Enviri (ex-Harsco)", "États-Unis", "vendeur de Clean Earth : ses comptes donnent l'historique de la cible"),
         ("Republic Services (US Ecology)", "États-Unis", "généraliste ; DD via US Ecology — mix différent"),
         ("Remondis / SARP", "Allemagne", "non coté : comparable métier, pas de multiple de marché")]
for name, ctry, why in peers:
    put(ws, f"B{r}", name); put(ws, f"C{r}", ctry)
    for col in "DEFGHIJ":
        put(ws, f"{col}{r}", None, fill=YELLOW)
    if name.startswith("Clean Harbors"):
        clh_rev, i_crev = E("Clean Harbors : chiffre d'affaires direct total", "FY2025")
        clh_eb, i_ceb = E("Clean Harbors : EBITDA ajusté total", "FY2025")
        put(ws, f"D{r}", "100 % (déchets dangereux et services associés)", color=GREY)
        put(ws, f"E{r}", f"={clh_rev}", color=GREEN, nf=NF_M1)
        put(ws, f"F{r}", f"={clh_eb}/1000", color=GREEN, nf=NF_M1)
        put(ws, f"I{r}", "01_Financial/clean-harbors-10-k-2025.pdf", color=GREY)
        put(ws, f"J{r}", f"{i_crev} p.33, {i_ceb} p.38", color=GREY)
        why += " — 2025, M USD ; EBITDA ajusté selon la définition de Clean Harbors"
    elif name.startswith("Enviri"):
        en_rev, i_erev = E("Enviri : chiffre d'affaires du segment Clean Earth", "FY2025")
        en_oi, i_eoi = E("Enviri : résultat opérationnel du segment Clean Earth", "FY2025")
        en_dep, i_edep = E("Enviri : dépréciation du segment Clean Earth", "FY2025")
        en_am, i_eam = E("Enviri : amortissement du segment Clean Earth", "FY2025")
        put(ws, f"D{r}", "segment Clean Earth seul", color=GREY)
        put(ws, f"E{r}", f"={en_rev}", color=GREEN, nf=NF_M1)
        put(ws, f"F{r}", f"={en_oi}+({en_dep}+{en_am})/1000", nf=NF_M1)
        put(ws, f"I{r}", "01_Financial/enviri-10-k-2025.pdf", color=GREY)
        put(ws, f"J{r}", f"{i_erev} p.69, {i_eoi} p.70, {i_edep}/{i_eam} p.284", color=GREY)
        why += " — 2025, M USD ; EBITDA = résultat opérationnel + D&A du segment, recalculé ici (Enviri ne publie pas d'EBITDA de segment)"
    put(ws, f"G{r}", f'=IF(AND(ISNUMBER(E{r}),ISNUMBER(F{r}),E{r}<>0),F{r}/E{r},"")', nf=NF_P)
    ws[f"G{r}"].fill = PatternFill()
    put(ws, f"K{r}", why, color=GREY)
    r += 1
clh_m, i_clhm = E("Clean Harbors : marge d'EBITDA ajusté", "FY2025")
put(ws, f"B{r}", "Pour mémoire : marge d'EBITDA ajusté de Clean Harbors publiée")
put(ws, f"G{r}", f"={clh_m}/100", color=BLACK, nf=NF_P); put(ws, f"K{r}", i_clhm, color=GREY); B["clhMargin"] = f"{q('Booster')}!$G${r}"; r += 2

put(ws, f"A{r}", "E", bold=True); put(ws, f"B{r}", "Clean Earth vu du vendeur (10-K 2025 d'Enviri) : quel EBITDA paie-t-on ?", bold=True); r += 1
en_rev24, i_erev24 = E("Enviri : chiffre d'affaires du segment Clean Earth", "FY2024")
en_oi24, i_eoi24 = E("Enviri : résultat opérationnel du segment Clean Earth", "FY2024")
line(ws, r, None, "Chiffre d'affaires du segment", f"={en_rev24}", "M USD", f"{i_erev24} (2024)", NF_M1, store=B); r += 1
line(ws, r, "ceRev25", "Chiffre d'affaires du segment", f"={en_rev}", "M USD", f"{i_erev} (2025)", NF_M1, store=B); r += 1
line(ws, r, None, "Croissance du chiffre d'affaires 2025", f"=C{r-1}/C{r-2}-1", "%", "", NF_P, store=B); r += 1
line(ws, r, None, "Résultat opérationnel du segment", f"={en_oi24}", "M USD", f"{i_eoi24} (2024)", NF_M1, store=B); r += 1
line(ws, r, None, "Résultat opérationnel du segment", f"={en_oi}", "M USD", f"{i_eoi} (2025)", NF_M1, store=B); r += 1
line(ws, r, None, "Dépréciation + amortissement du segment", f"=({en_dep}+{en_am})/1000", "M USD", f"{i_edep}, {i_eam} (2025)", NF_M1, store=B); r += 1
line(ws, r, "ceEbRec", "EBITDA 2025 recalculé (résultat opérationnel + D&A)", f"=C{r-2}+C{r-1}", "M USD", "non publié par Enviri : reconstitué", NF_M1, True, store=B); r += 1
line(ws, r, "ceMargRec", "Marge d'EBITDA 2025 recalculée", f"=C{r-1}/C{r-6}", "%", "", NF_P, store=B); r += 1
line(ws, r, None, "EBITDA 2026E retenu par Veolia (ajusté, post IFRS 16)", f"={ce_eb}", "M USD", id_ceeb, NF_M1, store=B); r += 1
line(ws, r, "ceEbGap", "Écart entre l'EBITDA de Veolia et l'EBITDA 2025 reconstitué", f"=C{r-1}/C{r-3}-1", "%", "croissance, IFRS 16, ajustements : à faire expliquer", NF_P, True, store=B); r += 1
line(ws, r, "mRec", "Multiple sur l'EBITDA 2025 reconstitué (VE / EBITDA)", f"={ce_ev_usd}*1000/C{r-4}", "x", f"{i_ceev} / recalcul", '0.0"x"', True, store=B); r += 1
line(ws, r, None, "Pour mémoire : multiple publié par Veolia après synergies", f"={ce_mult}", "x", id_cemult, '0.0"x"', store=B); r += 1
line(ws, r, None, "Actifs totaux du segment", f"={E('Enviri : actifs totaux du segment Clean Earth', '31/12/2025')[0]}/1000", "M USD",
     E("Enviri : actifs totaux du segment Clean Earth", "31/12/2025")[1], NF_M1, store=B); r += 1
put(ws, f"B{r}", "Lecture : sur l'EBITDA que le vendeur publie, le prix ressort à près de 20x ; le 9,8x de Veolia repose sur un EBITDA "
    "2026E un tiers plus haut, plus 120 M$ de synergies. C'est le pont à faire expliquer le 16 octobre.", color=GREY, italic=True, wrap=True)
ws.merge_cells(f"B{r}:E{r}"); ws.row_dimensions[r].height = 30; r += 1
put(ws, f"B{r}", "Cellules jaunes à remplir uniquement depuis un document versé dans la dataroom (fichier + page).",
    color=GREY, italic=True)
widths(ws, {"A": 4, "B": 60, "C": 13, "D": 12, "E": 40, "F": 11, "G": 9, "H": 11, "I": 20, "J": 7, "K": 60})

# ================================================================ ESG
ws = sheets["ESG"]
title(ws, "Rôle 5 — Le coût ESG : ce qu'une acquisition ajoute en passif et en exposition",
      "Provisions de fermeture et post-fermeture (URD 2025 p.432, amendement p.65) ; Clean Earth (amendement p.25, p.53).")
G = {}
r = 4
put(ws, f"A{r}", "A", bold=True); put(ws, f"B{r}", "Les passifs environnementaux de Veolia", bold=True); r += 1
header(ws, r, ["", "Provisions de fermeture et post-fermeture", "31/12/2024", "31/12/2025", "30/06/2026", "Réf."]); r += 1
cl24, i_cl24 = E("Provisions de fermeture et post-fermeture (total)", "31/12/2024")
cl25, i_cl25 = E("Provisions de fermeture et post-fermeture (total)", "31/12/2025")
put(ws, f"B{r}", "Total", bold=True)
put(ws, f"C{r}", f"={cl24}", color=GREEN, nf=NF_M); put(ws, f"D{r}", f"={cl25}", color=GREEN, nf=NF_M)
put(ws, f"E{r}", f"={clos26}", color=GREEN, nf=NF_M); put(ws, f"F{r}", f"{i_cl24}, {i_cl25}, {i_clos26}", color=GREY)
G["tot"] = r; r += 1
seg = [("Amériques, Asie-Pacifique, Afrique Moyen-Orient", "Provisions de fermeture, dont Amériques", "Provisions de fermeture, dont Amériques"),
       ("France et Déchets dangereux Europe", "Provisions de fermeture, dont France", "dont segment France et Déchets dangereux Europe"),
       ("Europe", "Provisions de fermeture, dont Europe", "Provisions de fermeture, dont Europe")]
for lab, p25, p26 in seg:
    a25, i25 = E(p25, "31/12/2025"); a26, i26 = E(p26, "30/06/2026")
    put(ws, f"B{r}", f"   dont {lab}")
    put(ws, f"D{r}", f"={a25}", color=GREEN, nf=NF_M); put(ws, f"E{r}", f"={a26}", color=GREEN, nf=NF_M)
    put(ws, f"F{r}", f"{i25}, {i26}", color=GREY)
    put(ws, f"G{r}", f"=E{r}-D{r}", nf=NF_M)
    r += 1
put(ws, f"G{G['tot']}", f"=E{G['tot']}-D{G['tot']}", nf=NF_M, bold=True)
put(ws, f"G{G['tot']-1}", "Δ S1 2026", bold=True, fill=HEAD)
G["am"] = G["tot"] + 1
r += 1
site, i_site = E("Provisions pour réhabilitation de sites", "31/12/2025")
envr, i_envr = E("Provisions pour risques environnementaux", "31/12/2025")
dism, i_dism = E("Provisions pour démantèlement", "31/12/2025")
unw, i_unw = E("Désactualisation des provisions", "FY2025")
line(ws, r, None, "Par nature au 31/12/2025 : réhabilitation de sites", f"={site}", "M EUR", i_site, store=G); r += 1
line(ws, r, None, "   + risques environnementaux", f"={envr}", "M EUR", i_envr, store=G); r += 1
line(ws, r, None, "   + démantèlement", f"={dism}", "M EUR", i_dism, store=G); r += 1
line(ws, r, "natSum", "   = total par nature", f"=SUM(C{r-3}:C{r-1})", "M EUR", "doit égaler le total au 31/12/2025", NF_M, True,
     store=G); r += 1
line(ws, r, None, "Désactualisation 2025 (charge financière annuelle)", f"={unw}", "M EUR", i_unw, store=G); r += 1
line(ws, r, None, "Provisions de fermeture en tours d'EBITDA 2025", f"={cl25}/{eb25}", "x", "", NF_X, store=G); r += 2

put(ws, f"A{r}", "B", bold=True); put(ws, f"B{r}", "Ce que Clean Earth apporte — et ce qu'on ne voit pas encore", bold=True); r += 1
gw, i_gw = E("Clean Earth : goodwill préliminaire", "30/06/2026")
g25, i_g25 = E("Engagements reçus liés au périmètre", "31/12/2025")
g26, i_g26 = E("Engagements reçus liés au périmètre", "30/06/2026")
line(ws, r, None, "Prix payé", f"={ce_eur}", "M EUR", id_ceeur, store=G); r += 1
line(ws, r, None, "Goodwill préliminaire", f"={gw}", "M EUR", i_gw, store=G); r += 1
line(ws, r, "gwPct", "Goodwill en % du prix", f"=C{r-1}/C{r-2}", "%", "affectation du prix non réalisée au 30/06/2026",
     NF_P, True, store=G); r += 1
line(ws, r, None, "Hausse des provisions de fermeture, segment Amériques, au S1 2026", f"=G{G['am']}", "M EUR",
     "compatible avec l'entrée de Clean Earth, non attribuée par Veolia", store=G); r += 1
line(ws, r, None, "Hausse des garanties reçues liées aux acquisitions", f"={g26}-{g25}", "M EUR", f"{i_g25}, {i_g26}",
     store=G); r += 1
clh_env, i_cenv = E("Clean Harbors : passifs environnementaux totaux", "31/12/2025")
clh_rev2, i_crev2 = E("Clean Harbors : chiffre d'affaires direct total", "FY2025")
en_cur, i_ecur = E("Enviri : passifs environnementaux, part courante", "31/12/2025")
en_lt, i_elt = E("Enviri : passifs environnementaux, part non courante", "31/12/2025")
en_rev25b, i_erev25b = E("Enviri : chiffre d'affaires du segment Clean Earth", "FY2025")
line(ws, r, "clhEnvPct", "Clean Harbors : passifs environnementaux en % du chiffre d'affaires", f"={clh_env}/1000/{clh_rev2}", "%",
     f"{i_cenv} / {i_crev2} (2025)", NF_P, store=G); r += 1
line(ws, r, "enEnv", "Enviri (groupe entier) : passifs environnementaux", f"=({en_cur}+{en_lt})/1000", "M USD", f"{i_ecur} + {i_elt}", NF_M1, store=G); r += 1
line(ws, r, "enEnvPct", "… en % du chiffre d'affaires de Clean Earth (borne haute : le groupe entier)", f"=C{r-1}/{en_rev25b}", "%", i_erev25b, NF_P, store=G); r += 1
line(ws, r, "ceEnvAnalog", "Ordre de grandeur par analogie Clean Harbors : % × CA Clean Earth (pas une mesure)", f"=C{r-3}*{en_rev25b}", "M USD",
     "analogie, à remplacer par la juste valeur à l'affectation du prix", NF_M1, store=G); r += 1
put(ws, f"B{r}", "Passifs environnementaux propres à Clean Earth (à sourcer : 10-Q T1 2026 d'Enviri, Clean Earth en activité abandonnée)", bold=True)
put(ws, f"C{r}", None, fill=YELLOW, nf=NF_M); put(ws, f"D{r}", "M EUR", color=GREY)
put(ws, f"E{r}", "cellule à remplir : laisser vide tant qu'aucun document ne le dit", color=GREY)
G["ceLiab"] = r; r += 1
line(ws, r, None, "Effet sur le levier 2027 s'ils étaient traités comme de la dette",
     f"=IF(ISNUMBER(C{r-1}),C{r-1}/{T['EB27']},0)", "x", "passifs / EBITDA 2027 central", NF_D2, True, store=G); r += 1
put(ws, f"B{r}", "Lecture : 81 % du prix est encore du goodwill. Tant que l'affectation n'est pas faite (au plus tard "
    "aux comptes 2026), la juste valeur des passifs environnementaux de Clean Earth n'est pas visible chez Veolia. "
    "C'est l'angle mort du rôle 5.", color=GREY, italic=True, wrap=True)
ws.merge_cells(f"B{r}:F{r}"); ws.row_dimensions[r].height = 44; r += 1
widths(ws, {"A": 4, "B": 62, "C": 12, "D": 12, "E": 12, "F": 40, "G": 11})

# ---------------------------------------------------------------- ESG : la taxonomie (recherche 12, question 3 du cours)
ws = sheets["ESG"]
r = ws.max_row + 2


def EXACT(label, period):
    """Comme E(), mais sur le libellé entier : « capex éligible » n'est pas « capex éligible et aligné »."""
    hits = [x for x in rows if x["label"] == label and x["period"] == period]
    if len(hits) != 1:
        raise SystemExit(f"{len(hits)} lignes pour {label!r} / {period!r}")
    USED.append({"id": hits[0]["id"], "prefix": label, "period": period, "unit": ""})
    return f"{q('Entrées')}!$C${ENTREE_ROW[hits[0]['id']]}", hits[0]["id"]


put(ws, f"A{r}", "T", bold=True); put(ws, f"B{r}", "La taxonomie européenne 2025 : où vont les euros « verts » (DEU 2025, p.250)", bold=True); r += 1
header(ws, r, ["", "Secteur d'activités éligibles", "CA éligible", "CA éligible et aligné", "Capex éligible", "Capex éligible et aligné",
               "Capex aligné / éligible", "Part du capex aligné du groupe", "Réf."]); r += 1
TX = {}
T0 = r
for key, lab, pre in (("water", "Eau et assainissement", "eau et assainissement"),
                      ("waste", "Déchets pour l'économie circulaire (déchets dangereux compris)", "collecte et traitement des déchets pour l'économie circulaire (déchets dangereux compris)"),
                      ("hw", "Déchets dangereux, prévention de la pollution (PPC 2.1, 2.2)", "collecte et traitement des déchets dangereux, prévention de la pollution (PPC 2.1, 2.2)"),
                      ("energy", "Énergie produite et distribuée", "énergie produite et distribuée"),
                      ("eserv", "Services énergétiques aux infrastructures", "services énergétiques aux infrastructures"),
                      ("other", "Autres activités éligibles", "autres activités éligibles")):
    refs = []
    put(ws, f"B{r}", lab)
    for col, what in (("C", "chiffre d'affaires éligible"), ("D", "chiffre d'affaires éligible et aligné"), ("E", "capex éligible"), ("F", "capex éligible et aligné")):
        a, i = EXACT(f"Taxonomie 2025, {pre} : {what}", "FY2025")
        put(ws, f"{col}{r}", f"={a}*1000", color=GREEN, nf=NF_M); refs.append(i)
    put(ws, f"G{r}", f'=IF(E{r}>0,F{r}/E{r},"")', nf=NF_P)
    put(ws, f"H{r}", f"=F{r}/$F${T0 + 7}", nf=NF_P)
    put(ws, f"I{r}", ", ".join(refs), color=GREY)
    TX[key] = r; r += 1
T1 = r - 1
put(ws, f"B{r}", "Somme des secteurs", bold=True)
for col in "CDEF":
    put(ws, f"{col}{r}", f"=SUM({col}{T0}:{col}{T1})", nf=NF_M, bold=True)
TX["sum"] = r; r += 1
put(ws, f"B{r}", "Total publié par Veolia (éligible ; éligible et aligné)", bold=True)
for col, what in (("C", "chiffre d'affaires éligible"), ("D", "chiffre d'affaires éligible et aligné"), ("E", "capex éligible"), ("F", "capex éligible et aligné")):
    a, i = EXACT(f"Taxonomie 2025 : {what}", "FY2025")
    put(ws, f"{col}{r}", f"={a}*1000", color=GREEN, nf=NF_M, bold=True)
TX["pub"] = r; r += 1
for rr in range(T0, T1 + 1):
    ws[f"H{rr}"] = f"=F{rr}/$F${TX['pub']}"; ws[f"H{rr}"].number_format = NF_P; ws[f"H{rr}"].font = font()
cx_tot, i_cxtot = EXACT("Taxonomie 2025 : capex total", "FY2025")
cx_pct, i_cxpct = E("Taxonomie 2025 : part du capex éligible et aligné", "FY2025")
cx_al24, i_cxal24 = EXACT("Taxonomie 2024 : capex éligible et aligné", "FY2024")
cx_tot24, i_cxtot24 = EXACT("Taxonomie 2024 : capex total", "FY2024")
line(ws, r, "cxTot", "Capex total du groupe (base taxonomie)", f"={cx_tot}*1000", "M EUR", i_cxtot, store=TX); r += 1
line(ws, r, "cxAlignedShare", "Part du capex éligible et aligné, recalculée", f"=F{TX['pub']}/C{r-1}", "%", "aligné / total", NF_P, True, store=TX); r += 1
line(ws, r, "cxAlignedPub", "… publiée", f"={cx_pct}/100", "%", i_cxpct, NF_P, store=TX); r += 1
line(ws, r, "cxAligned24", "Capex aligné 2024, pour mémoire (total 2024 : 4,1 Md€)", f"={cx_al24}*1000", "M EUR", f"{i_cxal24}, {i_cxtot24}", store=TX); r += 1
line(ws, r, "hwShare", "Déchets dangereux (PPC) : part du capex aligné du groupe", f"=F{TX['hw']}/F{TX['pub']}", "%", "", NF_P, True, store=TX); r += 1
line(ws, r, "hwAlignRate", "Déchets dangereux (PPC) : capex aligné / éligible", f"=G{TX['hw']}", "%", "", NF_P, True, store=TX); r += 1
line(ws, r, "nonEligible", "Capex non éligible (hors taxonomie)", f"=C{TX['cxTot'].split('$')[-1]}-E{TX['pub']}", "M EUR", "total − éligible", NF_M, store=TX); r += 1
put(ws, f"B{r}", "Lecture : près de la moitié du capex du groupe est « vert » au sens de la taxonomie, et l'eau en porte la plus grande part. "
    "Les déchets dangereux sont une petite ligne (0,3 Md€ alignés) mais la mieux alignée de toutes. Ce que la taxonomie ne dit pas, et que la "
    "question 3 du cours demande : l'impact par euro (tonnes traitées, CO2 évité, m³ économisés par M€ investi). Veolia publie ses KPI au "
    "niveau du groupe, pas par booster : c'est une question pour le 16 octobre, pas un calcul.", color=GREY, italic=True, wrap=True)
ws.merge_cells(f"B{r}:I{r}"); ws.row_dimensions[r].height = 58; r += 1
widths(ws, {"C": 14, "D": 16, "E": 14, "F": 16, "G": 14, "H": 16, "I": 40})

# ================================================================ Cibles
ws = sheets["Cibles"]
title(ws, "Rôle 6 — Calibrer la recommandation sur l'enveloppe",
      "Taille maximale d'une acquisition payée fin 2027 sans dépasser le plafond : marge / (1 − plafond / multiple).")
K = {}
r = 4
put(ws, f"A{r}", "A", bold=True); put(ws, f"B{r}", "Capacité selon le multiple payé et le scénario", bold=True); r += 1
header(ws, r, ["", "Multiple VE / EBITDA payé", "Défavorable", "Central", "Favorable", "Note"]); r += 1
for lab, key in (("Marge de manœuvre fin 2027 sous le levier ≤ 3x (M EUR)", "HEAD27"),
                 ("Marge fin 2027 sous le seuil S&P (FFO / dette ≥ 18 %)", "HEADSP27"),
                 ("Marge contraignante fin 2027 (la plus petite des deux)", "BIND27")):
    put(ws, f"B{r}", lab, bold=key == "BIND27")
    put(ws, f"C{r}", f"={q('Trajectoire')}!${UNFAV}${ROW[key]}", color=GREEN, nf=NF_M, bold=key == "BIND27")
    put(ws, f"D{r}", f"={T[key]}", color=GREEN, nf=NF_M, bold=key == "BIND27")
    put(ws, f"E{r}", f"={q('Trajectoire')}!${FAV}${ROW[key]}", color=GREEN, nf=NF_M, bold=key == "BIND27")
    put(ws, f"F{r}", "" if key != "BIND27" else "la grille ci-dessous part de cette ligne", color=GREY)
    HR = r; r += 1
mults = [(8, "bas de marché supposé"), (f"={ce_mult}", "Clean Earth après synergies"), (12, "hypothèse"),
         (f"={ce_ev_usd}*1000/{ce_eb}", "Clean Earth avant synergies (VE / EBITDA 2026E)")]
for m, note in mults:
    put(ws, f"B{r}", m, color=BLUE if not isinstance(m, str) else GREEN, nf='0.0"x"', fill=YELLOW if not isinstance(m, str) else None)
    for col in "CDE":
        put(ws, f"{col}{r}", f"=IF($B{r}>{HY('cap')},MAX(0,{col}${HR})/(1-{HY('cap')}/$B{r}),0)", nf=NF_M)
    put(ws, f"F{r}", note, color=GREY)
    r += 1
put(ws, f"B{r}", "Une marge négative donne 0 : il n'y a alors aucune place sans cession supplémentaire ou levée de fonds propres.",
    color=GREY, italic=True); r += 2

put(ws, f"A{r}", "B", bold=True); put(ws, f"B{r}", "Le multiple de Clean Earth se recalcule-t-il ?", bold=True); r += 1
line(ws, r, None, "Valeur d'entreprise", f"={ce_ev_usd}*1000", "M USD", i_ceev, store=K); r += 1
line(ws, r, None, "EBITDA 2026E", f"={ce_eb}", "M USD", id_ceeb, store=K); r += 1
line(ws, r, None, "Synergies en régime de croisière", f"={syn_rr}", "M USD", i_syn, store=K); r += 1
line(ws, r, "mPre", "Multiple avant synergies recalculé", f"=C{r-3}/C{r-2}", "x", "", '0.0"x"', True, store=K); r += 1
line(ws, r, "mPost", "Multiple après synergies recalculé", f"=C{r-4}/(C{r-3}+C{r-2})", "x", "", '0.0"x"', True, store=K); r += 1
line(ws, r, "mPub", "Multiple publié après synergies", f"={ce_mult}", "x", id_cemult, '0.0"x"', store=K); r += 1
put(ws, f"B{r}", "Écart à expliquer : le 9,8x publié ne se retrouve pas avec 3,0 Md$ / (200 + 120). Base d'EBITDA différente "
    "(pré IFRS 16 ?) ou synergies nettes des coûts ? Question pour le rôle 6.", color=GREY, italic=True, wrap=True)
ws.merge_cells(f"B{r}:F{r}"); ws.row_dimensions[r].height = 30; r += 2

put(ws, f"A{r}", "D", bold=True); put(ws, f"B{r}", "Les multiples que Veolia dit avoir payés (2024-2026)", bold=True); r += 1
header(ws, r, ["", "Opération", "Multiple VE / EBITDA", "Base", "Réf.", "Source"]); r += 1
m_t25, id_mt25 = E("Tuck-ins 2025 (États-Unis, Brésil, Japon) : multiple", "FY2025")
a_t25, id_at25 = E("Tuck-ins 2025 (États-Unis, Brésil, Japon) : montant", "FY2025")
M0 = r
for lab, f, base, ref, srcdoc in (
        ("Tuck-ins en Espagne, 13 opérations, 87 M€ de VE (2024-2025)", f"={m_es}", "moyenne publiée", id_mes, "Résultats 2025, p.19"),
        ("Tuck-ins 2025 : États-Unis, Brésil, Japon (370 M€)", f"={m_t25}", "moyenne publiée, environ", id_mt25, "Résultats 2025, p.11"),
        ("Clean Earth, après synergies en régime de croisière", f"={ce_mult}", "VE / EBITDA 2026e + synergies", id_cemult, "Présentation Clean Earth"),
        ("WTS, rachat des 30 % (1,75 Md$)", f"={m_wts}", "VE / EBITDA 2025e après synergies, environ", id_mwts, "Résultats T1 2025, p.11"),
        ("Clean Earth, avant synergies (recalculé)", f"={K['mPre']}", "VE / EBITDA 2026E de Veolia", "§B", "classeur"),
        ("Clean Earth, sur l'EBITDA 2025 du vendeur (recalculé)", f"={B['mRec']}", "VE / EBITDA 2025 reconstitué d'Enviri", "Booster §E", "classeur")):
    put(ws, f"B{r}", lab); put(ws, f"C{r}", f, color=GREEN, nf='0.0"x"'); put(ws, f"D{r}", base, color=GREY)
    put(ws, f"E{r}", ref, color=GREY); put(ws, f"F{r}", srcdoc, color=GREY); r += 1
M1 = r - 1
line(ws, r, "mMin", "Le moins cher payé (après synergies, tel que publié)", f"=MIN(C{M0}:C{M0+3})", "x", "min des quatre multiples publiés", '0.0"x"', True, store=K); r += 1
line(ws, r, "mMed", "Médiane des quatre multiples publiés", f"=MEDIAN(C{M0}:C{M0+3})", "x", "", '0.0"x"', store=K); r += 1
line(ws, r, "mMax", "Le plus cher payé (après synergies, tel que publié)", f"=MAX(C{M0}:C{M0+3})", "x", "max des quatre", '0.0"x"', True, store=K); r += 1
line(ws, r, "mSpread", "Écart avant / après synergies sur Clean Earth (recalculé)", f"=C{M0+4}-C{M0+2}", "x", "ce que les synergies « achètent » de multiple", '0.0"x"', store=K); r += 1
put(ws, f"B{r}", "Lecture : Veolia publie ses multiples après synergies, entre 7x et 11x ; l'hypothèse mTuck (Hypothèses) est bornée par "
    "ces deux points, Clean Earth au centre. Avant synergies, le même Clean Earth coûte près de deux fois plus : c'est la base qu'il faut "
    "préciser à chaque comparaison — et la question à poser pour toute cible.", color=GREY, italic=True, wrap=True)
ws.merge_cells(f"B{r}:F{r}"); ws.row_dimensions[r].height = 44; r += 2

put(ws, f"A{r}", "E", bold=True); put(ws, f"B{r}", "Les multiples du secteur (communiqués des acquéreurs, versés dans la dataroom)", bold=True); r += 1
header(ws, r, ["", "Opération (acquéreur → cible)", "VE (M USD)", "EBITDA de référence (M USD)", "VE / EBITDA avant synergies", "Multiple publié après synergies", "Base de l'EBITDA", "Réf."]); r += 1
SECTOR = [
    # (libellé, préfixe VE, période VE, ×1000 ?, préfixe EBITDA, période EBITDA, préfixe multiple publié, période, base)
    ("Clean Harbors → EnviroServe (2026)", "EnviroServe (Clean Harbors) : prix d'acquisition", "08/2026", False,
     "EnviroServe (Clean Harbors) : EBITDA ajusté annuel attendu", "2026e", "EnviroServe et ES&H (Clean Harbors) : multiple combiné après synergies", "10/2026",
     "EBITDA ajusté annuel attendu ; 8,9x après synergies publié pour EnviroServe et ES&H ensemble"),
    ("Clean Harbors → HEPACO (2024)", "HEPACO (Clean Harbors) : prix d'acquisition", "03/2024", False,
     "HEPACO (Clean Harbors) : EBITDA ajusté 2023", "FY2023", "HEPACO (Clean Harbors) : multiple après synergies", "2024", "EBITDA ajusté 2023 ; synergies ~20 M$"),
    ("EQT → Covanta (2021)", "Covanta (EQT) : valeur de la transaction", "annonce 07/2021", True,
     "Covanta (EQT) : EBITDA ajusté 2021 attendu, bas", "2021e", None, None, "bas de la fourchette d'EBITDA 2021 attendu (460-480)"),
    ("WM → Stericycle (2024)", "Stericycle (WM) : valeur d'entreprise", "annonce 06/2024", True,
     None, None, None, None, "EBITDA non publié dans le communiqué (synergies > 125 M$/an)"),
    ("Republic Services → US Ecology (2022)", "US Ecology (Republic Services) : valeur totale", "annonce 02/2022", True,
     "US Ecology (Republic Services) : EBITDA ajusté 12 mois", "au 30/09/2021", None, None, "EBITDA ajusté des 12 mois au 30/09/2021"),
]
S0 = r
for lab, ev_pre, ev_per, ev_bn, eb_pre, eb_per, mp_pre, mp_per, base in SECTOR:
    put(ws, f"B{r}", lab)
    try:
        ev, i_ev = E(ev_pre, ev_per)
    except SystemExit:
        ev, i_ev = None, "non versé"
    refs = [i_ev]
    put(ws, f"C{r}", (f"={ev}*1000" if ev_bn else f"={ev}") if ev else "n/d", color=GREEN if ev else GREY, nf=NF_M, align=None if ev else "right")
    if eb_pre:
        try:
            eb, i_eb = E(eb_pre, eb_per); refs.append(i_eb)
            put(ws, f"D{r}", f"={eb}", color=GREEN, nf=NF_M)
            put(ws, f"E{r}", f"=C{r}/D{r}", nf='0.0"x"', bold=True)
        except SystemExit:
            put(ws, f"D{r}", "non versé", color=GREY, align="right"); put(ws, f"E{r}", "—", color=GREY, align="center")
    else:
        put(ws, f"D{r}", "n/d", color=GREY, align="right"); put(ws, f"E{r}", "—", color=GREY, align="center")
    if mp_pre:
        mp, i_mp = E(mp_pre, mp_per); refs.append(i_mp)
        put(ws, f"F{r}", f"={mp}", color=GREEN, nf='0.0"x"')
    else:
        put(ws, f"F{r}", "—", color=GREY, align="center")
    put(ws, f"G{r}", base, color=GREY); put(ws, f"H{r}", ", ".join(refs), color=GREY)
    r += 1
S1 = r - 1
line(ws, r, "secMin", "Secteur : multiple avant synergies le plus bas (opérations à EBITDA publié)", f"=MIN(E{S0}:E{S1})", "x", "", '0.0"x"', True, store=K); r += 1
line(ws, r, "secMax", "Secteur : multiple avant synergies le plus haut", f"=MAX(E{S0}:E{S1})", "x", "", '0.0"x"', True, store=K); r += 1
line(ws, r, "secMed", "Secteur : médiane des multiples avant synergies", f"=MEDIAN(E{S0}:E{S1})", "x",
     f"{sum(1 for x in SECTOR if x[4])} opérations à EBITDA publié", '0.0"x"', True, store=K); r += 1
line(ws, r, "secPostMin", "Secteur : multiple publié après synergies, le plus bas", f"=MIN(F{S0}:F{S1})", "x", "", '0.0"x"', store=K); r += 1
line(ws, r, "secPostMax", "Secteur : multiple publié après synergies, le plus haut", f"=MAX(F{S0}:F{S1})", "x", "", '0.0"x"', store=K); r += 1
line(ws, r, "ceVsSec", "Clean Earth avant synergies (recalculé) face au haut du secteur", f"={K['mPre']}-{K['secMax']}", "x", "positif = Veolia a payé plus que le plus cher du secteur, avant synergies", '0.0"x"', True, store=K); r += 1
def _x1(ref):
    return f'SUBSTITUTE(TEXT({ref},"0.0"),".",",")&"x"'
_k = {k: K[k] for k in ("secMin", "secMax", "secPostMin", "secPostMax", "mPre", "ceVsSec", "mPub")}
_lect = ('="Lecture : avant synergies, le secteur paie de "&' + _x1(_k["secMin"]) + '&" à "&' + _x1(_k["secMax"])
         + f'&" (le plus cher : "&INDEX(B{S0}:B{S1},MATCH({_k["secMax"]},E{S0}:E{S1},0))&") ; après synergies, les acquéreurs publient "&'
         + _x1(_k["secPostMin"]) + '&" à "&' + _x1(_k["secPostMax"]) + '&". Clean Earth, "&' + _x1(_k["mPre"])
         + '&" avant synergies sur l\'EBITDA 2026E de Veolia, "&'
         + f'IF({_k["ceVsSec"]}>0,"est au-dessus de tout ce que le secteur a payé","reste sous le plus cher du secteur, de "&' + _x1("-" + _k["ceVsSec"]) + ')'
         + '&" ; à "&' + _x1(_k["mPub"]) + '&" après synergies, il est dans la norme. Une petite cible se paie cher avant synergies : '
         'ce sont les synergies qui ramènent le prix sous 10x. Stericycle n\'a pas d\'EBITDA publié dans son communiqué."')
put(ws, f"B{r}", _lect, color=GREY, italic=True, wrap=True)
ws.merge_cells(f"B{r}:H{r}"); ws.row_dimensions[r].height = 58; r += 2

put(ws, f"A{r}", "C", bold=True); put(ws, f"B{r}", "L'univers de cibles, noté (00_Admin/targets.csv : une ligne par cible, sourcée)", bold=True); r += 1
put(ws, f"B{r}", "Chaque cible est confrontée aux 1 000 futurs de l'onglet Simulation : elle « tient » dans un tirage si son prix n'excède pas "
    "marge contraignante / (1 − 3 / multiple) dans ce tirage. Valeurs = dernière valeur publiée (année dans Propriétaire).", color=GREY, italic=True); r += 1
TARGETS_CSV = admin_file("targets.csv")
TARGETS = list(csv.DictReader(TARGETS_CSV.open(encoding="utf-8"))) if TARGETS_CSV.exists() else []
line(ws, r, "mRet", "Multiple payé retenu pour une cible (avant synergies) — jugement du groupe", f"={K['secMed']}", "x",
     "par défaut : médiane du secteur (§E) ; une cible peut avoir le sien (colonne J)", '0.0"x"', True, store=K)
ws[f"C{r}"].fill = YELLOW; r += 1
cad, i_cad = E("Taux de référence BCE : dollars canadiens pour un euro", "06/10/2026")
line(ws, r, "cad", "Dollars canadiens pour un euro (BCE)", f"={cad}", "CAD / EUR", i_cad, '0.0000', store=K); r += 1
put(ws, f"B{r}", "Le dollar américain est converti au taux de l'hypothèse fx (Hypothèses), celui de Clean Earth à la clôture.",
    color=GREY, italic=True); r += 2
CRIT = [("Traitement", "actifs de traitement (incinération, TSDF), pas de collecte seule"),
        ("Géographie", "États-Unis ou zone où Veolia veut croître en déchets dangereux"),
        ("PFAS", "offre PFAS ou polluants émergents"),
        ("Concurrence", "pas de doublon avec Veolia (2 = aucun chevauchement)"),
        ("Passifs", "passifs environnementaux connus et limités")]
CW = ["N", "O", "P", "Q", "R"]
put(ws, f"M{r}", "Poids", bold=True, align="right")
for col, (name, _) in zip(CW, CRIT):
    put(ws, f"{col}{r}", 1, color=BLUE, fill=YELLOW, nf="0", align="center")
WROW = r; r += 1
header(ws, r, ["", "Cible", "Pays", "Activité", "Propriétaire (depuis)", "Valeur publiée", "Unité", "Base", "VE (M EUR)",
               "Multiple retenu", "Tient au central ?", "Tient dans X % des tirages", "Années de détention fin 2027"]
       + [c[0] + " (0-2)" for c in CRIT] + ["Score (0-2)", "Réf."]); r += 1
GR0 = r
N_ = f"{SIM_N}"
def _fits(row, scen_cell):
    return (f'=IF(AND(ISNUMBER(I{row}),ISNUMBER(J{row})),IF(I{row}<=IF(J{row}>{HY("cap")},MAX(0,{scen_cell})/(1-{HY("cap")}/J{row}),0),'
            f'"oui","non"),"")')
def _prob(row):
    return (f'=IF(AND(ISNUMBER(I{row}),ISNUMBER(J{row})),IF(J{row}>{HY("cap")},'
            f'COUNTIF({SIM["BIND27"]},">="&I{row}*(1-{HY("cap")}/J{row}))/{N_},0),"")')
# l'étalon : Clean Earth, au prix payé et au multiple publié
put(ws, f"B{r}", "Clean Earth (étalon, déjà acquis)", italic=True); put(ws, f"C{r}", "États-Unis", italic=True)
put(ws, f"D{r}", "TSDF, 82 sites", italic=True); put(ws, f"E{r}", "Veolia (depuis le 01/06/2026)", italic=True)
put(ws, f"F{r}", f"=-{ce_nfd}", color=GREEN, nf=NF_M); put(ws, f"G{r}", "M EUR", color=GREY); put(ws, f"H{r}", "effet sur la dette", color=GREY)
put(ws, f"I{r}", f"=F{r}", nf=NF_M); put(ws, f"J{r}", f"={ce_mult}", color=GREEN, nf='0.0"x"')
put(ws, f"K{r}", _fits(r, f"$D${HR}")); put(ws, f"L{r}", _prob(r), nf="0%", bold=True)
put(ws, f"T{r}", f"{i_cenfd}, {id_cemult}", color=GREY)
CE_ROW = r; r += 1
SCALE = {"M": 1, "Md": 1000}
for t in TARGETS:
    if "Veolia" in (t.get("owner") or ""):
        continue
    put(ws, f"B{r}", t["name"]); put(ws, f"C{r}", t["country"]); put(ws, f"D{r}", t["activity"])
    put(ws, f"E{r}", t.get("owner", ""))
    unit = (t.get("currency") or "").split()
    scale, cur = (SCALE.get(unit[0]), unit[1]) if len(unit) == 2 else (None, unit[0] if unit else "")
    try:
        v = float((t.get("ev_estimate") or "").replace(",", ""))
    except ValueError:
        v = None
    if v is not None:
        put(ws, f"F{r}", v, color=BLUE, nf="#,##0.0"); put(ws, f"G{r}", t["currency"], color=GREY)
        put(ws, f"H{r}", {"enterprise_value": "VE", "equity_value": "fonds propres", "other": "valorisation"}.get(t.get("ev_basis"), t.get("ev_basis", "")), color=GREY)
        rate = {"EUR": "1", "USD": HY("fx"), "CAD": K["cad"]}.get(cur)
        if scale and rate:
            put(ws, f"I{r}", f"=F{r}*{scale}/{rate}", nf=NF_M)
        else:
            put(ws, f"I{r}", "devise non convertie", color=GREY, align="right")
    else:
        put(ws, f"F{r}", "non publié", color=GREY, align="right"); put(ws, f"I{r}", "n/d", color=GREY, align="right")
    put(ws, f"J{r}", f"={K['mRet']}", nf='0.0"x"', fill=YELLOW)
    put(ws, f"K{r}", _fits(r, f"$D${HR}")); put(ws, f"L{r}", _prob(r), nf="0%", bold=True)
    yr = re.search(r"depuis[^0-9]*?(?:\d{1,2}/\d{1,2}/)?(\d{4})", t.get("owner", ""))
    if yr:
        put(ws, f"M{r}", f"=2027-{yr.group(1)}", nf="0", align="center")
    for col in CW:
        put(ws, f"{col}{r}", None, fill=YELLOW, nf="0", align="center")
    put(ws, f"S{r}", f'=IF(COUNT(N{r}:R{r})=0,"",SUMPRODUCT($N${WROW}:$R${WROW},N{r}:R{r})/SUMPRODUCT($N${WROW}:$R${WROW}*(N{r}:R{r}<>"")))',
        nf="0.0", bold=True)
    put(ws, f"T{r}", t["id"] + " — " + (t.get("file") or t.get("source_url") or "").split("/")[-1][:44], color=GREY)
    r += 1
GR1 = r - 1
r += 1
line(ws, r, "tN", "Cibles sourcées (hors étalon)", f"=COUNTA(B{CE_ROW + 1}:B{GR1})", "", "targets.csv", NF_I, store=K); r += 1
line(ws, r, "tValued", "… dont avec une valeur publiée convertie en euros", f"=COUNT(I{CE_ROW + 1}:I{GR1})", "", "", NF_I, store=K); r += 1
line(ws, r, "tFit", "… dont qui tiennent dans la marge centrale", f'=COUNTIF(K{CE_ROW + 1}:K{GR1},"oui")', "", "", NF_I, True, store=K); r += 1
line(ws, r, "tBestP", "Meilleure probabilité de tenir parmi les cibles", f"=MAX(L{CE_ROW + 1}:L{GR1})", "", "", "0%", store=K); r += 1
line(ws, r, "tMaxFit", "La plus grande cible qui tient au central (VE, M EUR)",
     f'=_xlfn.MAXIFS(I{CE_ROW + 1}:I{GR1},K{CE_ROW + 1}:K{GR1},"oui")', "M EUR", "", NF_M, store=K); r += 1
put(ws, f"B{r}", "Cellules jaunes : le multiple payé (par défaut la médiane du secteur) et la notation qualitative, à remplir et "
    "défendre par le groupe ; le score est la moyenne pondérée des critères renseignés. Ajouter une cible : register_target, la ligne apparaît "
    "au prochain passage de la chaîne.", color=GREY, italic=True, wrap=True)
ws.merge_cells(f"B{r}:L{r}"); ws.row_dimensions[r].height = 30; r += 1
for col, (name, why) in zip(CW, CRIT):
    put(ws, f"B{r}", f"{name} : {why}", color=GREY); r += 1
widths(ws, {"A": 4, "B": 46, "C": 16, "D": 30, "E": 30, "F": 14, "G": 9, "H": 12, "I": 12, "J": 12, "K": 12, "L": 14, "M": 12,
            "N": 11, "O": 11, "P": 11, "Q": 11, "R": 11, "S": 10, "T": 40})

# ================================================================ Vérifications
ws = sheets["Vérifications"]
title(ws, "Vérifications — le modèle se contrôle lui-même",
      "Bloquant : doit être OK avant toute citation. Info : écart documenté, à expliquer, pas une erreur du modèle.")
header(ws, 4, ["#", "Contrôle", "Calculé", "Attendu", "Tolérance", "Statut", "Type"])
checks = [
    ("Levier 2024 recalculé = publié", A["lev24c"], A["lev24p"], 0.006, NF_X, "Bloquant"),
    ("Levier 2025 recalculé = publié", A["lev25c"], A["lev25p"], 0.006, NF_X, "Bloquant"),
    ("Pont S1 2026 : dette recalculée = publiée", P["close26c"], P["close26p"], 2, NF_M, "Bloquant"),
    ("Investissements nets GreenUp : 1,9 + 1,5 − 0,3 = 3,1", C["netcapC"], C["netcapP"], 0.001, NF_D2, "Bloquant"),
    ("Provisions par nature = total au 31/12/2025", G["natSum"], f"{cl25}", 0.5, NF_M, "Bloquant"),
    ("Dividende 2026 : 1,50 € × nombre d'actions ≈ 1 099 M€",
     f"{E('Dividende proposé au titre de 2025', '2025')[0]}*{E('Nombre moyen pondéré', 'S1 2026')[0]}/1000000",
     div_sh26, 15, NF_M, "Bloquant"),
    ("Levier fin 2026 central dans la guidance « égal ou légèrement supérieur à 3x »", T["LEV26"], 3.075, 0.075, NF_X,
     "Bloquant"),
    ("Change implicite de Clean Earth plausible (1,05 à 1,30)", HY("fx"), 1.175, 0.125, '0.0000', "Bloquant"),
    ("Enveloppe nette annoncée = haut de la fourchette implicite", C["netHi"], C["netAnn"], 0.001, NF_D2, "Bloquant"),
    ("Investissements financiers nets 2024 : deck Clean Earth (−0,4 Md€) = communiqué (+397 M€)",
     f"-{E('Investissements financiers nets 2024', 'FY2024')[0]}*1000",
     E("Pont de dette 2024 : investissements financiers nets des cessions", "FY2024")[0], 60, NF_M, "Bloquant"),
    ("Dette maximale 2027 à l'objectif 8 Md€ ≈ 24 Md€ (cadrage du cours)", A["maxDebtTgt"], 24000, 1, NF_M, "Bloquant"),
    ("FFO / dette ajustée 2026 du modèle face à la prévision Moody's (18,5 %, ± 2 pts)", T["RATIO26"], f"{mo_f26}/100", 0.02, NF_P, "Bloquant"),
    ("FFO / dette ajustée 2027 du modèle face à la prévision Moody's (19,5 %, ± 2,5 pts)", T["RATIO27"], f"{mo_f27}/100", 0.025, NF_P, "Bloquant"),
    ("Dette brute publiée selon Moody's = sous-total des emprunts du DEU (p.351)", A["moRep"], f"{urd_gross}", 1, NF_M, "Bloquant"),
    ("Réconciliation Moody's : brute + ajustements = dette brute ajustée publiée", A["moAdjC"], A["moAdjP"], 1, NF_M, "Bloquant"),
    ("Réconciliation Moody's : brute ajustée − trésorerie = dette nette ajustée publiée", A["moNetC"], A["moNetP"], 1, NF_M, "Bloquant"),
    ("FFO implicite (20,3 % × dette nette) = FFO publié par Moody's (5 160)", A["ffo25"], A["ffo25p"], 60, NF_M, "Bloquant"),
    ("Dette nette Veolia recalculée (brute − trésorerie − actifs liquides + JV) face à la publiée : l'écart est le retraitement PPA Suez", A["veoNet"], f"{nfd25}", None, NF_M, "Info"),
    ("Marge des boosters 2025 recalculée = publiée (12,6 %)", SG["boostMargin"], SG["boostMarginPub"], 0.002, NF_P, "Bloquant"),
    ("Clean Harbors : marge d'EBITDA ajusté recalculée = publiée (19,4 %)",
     f"{E('Clean Harbors : EBITDA ajusté total', 'FY2025')[0]}/1000/{E('Clean Harbors : chiffre d\'affaires direct total', 'FY2025')[0]}",
     B["clhMargin"], 0.002, NF_P, "Bloquant"),
    ("Segments 2025 + autres = EBITDA du groupe (autres ≈ 139 M€, note IFRS 8)", SG["other25"], 139, 60, NF_M, "Bloquant"),
    ("CO2 évité : objectif +30 % vs 2023 (17,5 Mt) face aux 18 Mt énoncés", SG["co2target"], 18, None, NF_D2, "Info"),
    ("EBITDA 2027 « au rythme de 2025 » face à l'objectif ≥ 8 Md€", SG["eb27seg"], 8000, None, NF_M, "Info"),
    ("Dette ajustée 2026 du modèle face au pic attendu par Moody's (~29 Md€, ± 1,5 Md€)", T["ADJ26"], 29000, 1500, NF_M, "Bloquant"),
    ("EBITDA 2027 central face à l'objectif ≥ 8 Md€", T["EB27"], f"{tgt}*1000", None, NF_M, "Info"),
    ("Marge de FFO 2026 au seuil S&P de 18 % (négatif : le seuil d'agence mord avant le 3x)", A["ffoGap"], 0, None, NF_M, "Info"),
    ("Pont 2025 : flux non détaillés dans le communiqué", P["resid25"], 0, None, NF_M, "Info"),
    ("Clean Earth : effet sur la dette − prix payé", P["ceGap"], 0, None, NF_M, "Info"),
    ("Multiple Clean Earth recalculé après synergies vs publié", K["mPost"], K["mPub"], None, '0.0"x"', "Info"),
    ("Multiples publiés par Veolia : le minimum (Espagne) est bien la borne basse de mTuck", K["mMin"], HY("mTuck", "D"), 0.001, '0.0"x"', "Bloquant"),
    ("Multiples publiés par Veolia : le maximum (WTS) est bien la borne haute de mTuck", K["mMax"], HY("mTuck", "E"), 0.001, '0.0"x"', "Bloquant"),
    ("Monte-Carlo : la ligne « base » redonne la marge contraignante de Trajectoire (quand toutes les hypothèses sont à la base)",
     f'IF(COUNTIF({q("Hypothèses")}!$I$5:$I${H_LAST},"Bas")+COUNTIF({q("Hypothèses")}!$I$5:$I${H_LAST},"Haut")=0,{SIMB["BIND27"]}-{T["BIND27"]},0)',
     0, 0.5, NF_M, "Bloquant"),
    ("Monte-Carlo : moyenne des uniformes tirées (≈ 0,5 si les tirages sont sains)",
     "AVERAGE(" + ",".join(f"{q('Simulation')}!${c}${SIM0}:${c}${SIM1}" for c in ["B"] + [UCOL[k] for k in VARIED]) + ")", 0.5, 0.02, '0.000', "Bloquant"),
    ("Pont EBITDA : somme des marches 2025 → 2027 = EBITDA 2027 de Trajectoire", PB["eb27"], PB["eb27T"], 1, NF_M, "Bloquant"),
    ("Pont EBITDA : croissance annuelle implicite de GreenUp (6,5 → 8) face au « ~5 % » annoncé", PB["cagrC"], PB["cagrP"], 0.01, NF_P, "Bloquant"),
    ("Pont EBITDA : strongholds + boosters (EBITDA × organique) face à la croissance organique du groupe en M EUR (± 60)", PB["famOrg"], PB["orgM"], None, NF_M, "Info"),
    ("Pont EBITDA : part de l'écart 2023 → 2027 couverte par l'efficacité annoncée (info)", PB["effShare"], 1, None, NF_P, "Info"),
    ("Pont EBITDA : hors efficacité et synergies, croissance organique 2025 de l'EBITDA (négatif = recul)", PB["rest"], 0, None, NF_M, "Info"),
    ("Univers des cessions : somme des pays = chiffre d'affaires du groupe (DEU p.375, arrondis de la page : ± 5)", C["ctrySum"], C["ctryGrp"], 5, NF_M, "Bloquant"),
    ("Hybrides : somme des tranches = encours publié hors coupons (4,1 Md€)", A["hybSum"], A["hybPub"], 1, NF_M, "Bloquant"),
    ("Taxonomie 2025 : somme des secteurs = total publié, capex éligible et aligné (arrondis de la page : ± 150)", f"{q('ESG')}!$F${TX['sum']}", f"{q('ESG')}!$F${TX['pub']}", 150, NF_M, "Bloquant"),
    ("Taxonomie 2025 : part du capex aligné recalculée = publiée (47,2 %, ± 1 pt)", TX["cxAlignedShare"], TX["cxAlignedPub"], 0.01, NF_P, "Bloquant"),
    ("Univers des cessions : programme (au multiple central) en part de l'EBITDA des déchets solides (info)", C["progShareSw"], 0, None, NF_P, "Info"),
    ("FFO 2025 reconstitué depuis le tableau de flux = FFO Moody's (± 5 %)", A["ffoRecGap25"], 0, 0.05, NF_P, "Bloquant"),
    ("FFO 2024 reconstitué depuis le tableau de flux = FFO Moody's (± 5 %)", A["ffoRecGap24"], 0, 0.05, NF_P, "Bloquant"),
    ("Échéancier : somme des trois instruments = passifs financiers bruts publiés, flux 2026", f"{q('Échéancier')}!$C${M['sum']}",
     f"{q('Échéancier')}!$C${M['gross']}", 1, NF_M, "Bloquant"),
    ("Échéancier : somme des années = total publié (obligations)", f"{q('Échéancier')}!$J${M['bonds']}",
     f"{q('Échéancier')}!$I${M['bonds']}", 1, NF_M, "Bloquant"),
    ("Échéancier : somme des années = total publié (passifs financiers bruts)", f"{q('Échéancier')}!$J${M['gross']}",
     f"{q('Échéancier')}!$I${M['gross']}", 1, NF_M, "Bloquant"),
    ("Liquidités totales recalculées = publiées (p.421)", MR["liqC"], MR["liqP"], 1, NF_M, "Bloquant"),
    ("Flux contractuels 2026 des obligations − obligations à moins d'un an (positif : les flux comprennent les intérêts)",
     f"{q('Échéancier')}!$C${M['bonds']}-{q('Échéancier')}!$C${M['bond1y']}", 0, None, NF_M, "Info"),
    ("Chiffres de l'onglet Entrées non relus par un tiers", f"COUNTIF({q('Entrées')}!$L${FIRST}:$L${LAST_ENTREE},\"à relire\")",
     0, None, NF_I, "Info"),
]
r = 5
for n, (lab, calc_ref, exp, tol, nf, kind) in enumerate(checks, 1):
    put(ws, f"A{r}", n, align="center")
    put(ws, f"B{r}", lab, wrap=True)
    put(ws, f"C{r}", f"={calc_ref}", nf=nf, color=GREEN if "!" in str(calc_ref) and "*" not in str(calc_ref) else BLACK)
    put(ws, f"D{r}", f"={exp}" if isinstance(exp, str) else exp, nf=nf,
        color=GREEN if isinstance(exp, str) else BLUE)
    if tol is None:
        put(ws, f"E{r}", "—", color=GREY, align="center")
        put(ws, f"F{r}", "INFO", bold=True, align="center")
    else:
        put(ws, f"E{r}", tol, color=BLUE, nf=nf)
        put(ws, f"F{r}", f'=IF(ABS(C{r}-D{r})<=E{r},"OK","ÉCART")', bold=True, align="center")
    put(ws, f"G{r}", kind, color=GREY)
    ws.row_dimensions[r].height = 28
    r += 1
V1 = r - 1
r += 1
put(ws, f"B{r}", "Statut global des contrôles bloquants", bold=True)
put(ws, f"F{r}", f'=IF(COUNTIFS($G$5:$G${V1},"Bloquant",$F$5:$F${V1},"ÉCART")=0,"OK","ÉCART")', bold=True, align="center")
VSTAT = f"{q('Vérifications')}!$F${r}"
from openpyxl.formatting.rule import CellIsRule
ws.conditional_formatting.add(f"F5:F{r}", CellIsRule(operator="equal", formula=['"OK"'], fill=OKF))
ws.conditional_formatting.add(f"F5:F{r}", CellIsRule(operator="equal", formula=['"ÉCART"'], fill=KOF))
widths(ws, {"A": 4, "B": 64, "C": 13, "D": 13, "E": 11, "F": 10, "G": 10})

# ================================================================ Lisez-moi
ws = ws0
title(ws, "Modèle GreenUp 2027 — Veolia, sujet 2 du capstone EDHEC",
      "Combien Veolia peut-elle encore acheter en déchets dangereux sans dépasser son plafond de levier ?")
r = 4
put(ws, f"B{r}", "Résultats du scénario actif (base partout, sauf choix dans Hypothèses, colonne Actif)", bold=True, size=12); r += 1
res = [
    ("Probabilité que Veolia tienne fin 2027 ses deux plafonds (3x et seuil S&P), sur 1 000 futurs tirés", DS["pBind"], '0%'),
    ("Probabilité de pouvoir acheter au moins 1 Md€ fin 2027 sous la contrainte qui mord", DS["pAcq1"], '0%'),
    ("Cible de 3 Md€ fin 2027 : hybrides nécessaires pour tenir les deux plafonds, en M EUR (Financement §B)", f"{q('Financement')}!$E${FN['hyb']}", NF_M),
    ("… ou cessions accélérées nécessaires, en M EUR", f"{q('Financement')}!$E${FN['disp']}", NF_M),
    ("Hypothèses qui font basculer au moins une conclusion à l'intérieur de leurs bornes (Financement §D)", FN["nFragile"], NF_I),
    ("Probabilité que le ratio FFO / dette passe sous le seuil S&P fin 2026", DS["pSP26"], '0%'),
    ("Acquisition maximale fin 2027 sous la contrainte qui mord : médiane des tirages, en M EUR", f"{q('Distribution')}!$F${DS['MAXACQB']}", NF_M),
    ("Levier fin 2026 (guidance : égal ou légèrement supérieur à 3x)", T["LEV26"], NF_X),
    ("Levier fin 2027 (engagement : ≤ 3x)", T["LEV27"], NF_X),
    ("Marge de manœuvre fin 2027, en M EUR de dette", T["HEAD27"], NF_M),
    ("Acquisition maximale fin 2027 au multiple de Clean Earth, en M EUR", T["MAXACQ"], NF_M),
    ("EBITDA 2027 moins l'objectif ≥ 8 Md€, en M EUR", T["GAP8"], NF_M),
    ("Marge 2027 dans le scénario défavorable combiné, en M EUR", f"{q('Trajectoire')}!${UNFAV}${ROW['HEAD27']}", NF_M),
    ("Levier fin 2027 si 50 % des hybrides comptent en dette", f"{q('Levier')}!$C${A['def_rows'][0] + 1}", NF_X),
    ("Sorties nettes d'acquisitions cumulées depuis 2024, en Md EUR (enveloppe boosters : 4)", C["spentCum"], NF_D2),
    ("Marge de FFO en 2026 au seuil S&P (18 % de la dette ajustée au pic), en M EUR", A["ffoGap"], NF_M),
    ("Marge de dette fin 2027 sous le seuil S&P (18 % FFO / dette ajustée), en M EUR", T["HEADSP27"], NF_M),
    ("Contrainte qui mord en premier fin 2027", T["WHICH27"], None),
    ("EBITDA 2027 si chaque segment garde son rythme organique de 2025, en M EUR", SG["eb27seg"], NF_M),
    ("Clean Earth : multiple du prix sur l'EBITDA 2025 publié par le vendeur (reconstitué)", B["mRec"], '0.0"x"'),
    ("Acquisition maximale fin 2027 sous la contrainte qui mord, au multiple de Clean Earth, en M EUR", T["MAXACQB"], NF_M),
    ("Part de l'écart GreenUp 2023 → 2027 que couvre la seule efficacité promise (350 M€ par an)", PB["effShare"], NF_P),
    ("Part de la croissance organique 2025 de l'EBITDA expliquée par les gains d'efficacité", PB["effShareY"], NF_P),
    ("Marge sous 3x fin 2027 si l'efficacité 2026-2027 n'est pas livrée du tout, en M EUR", PB["head27k0"], NF_M),
    ("Multiples payés par Veolia, après synergies : du moins cher au plus cher (Cibles §D)", f'=TEXT({K["mMin"]},"0.0")&"x à "&TEXT({K["mMax"]},"0.0")&"x"', None),
    ("Secteur, avant synergies : du moins cher au plus cher (Cibles §E)", f'=TEXT({K["secMin"]},"0.0")&"x à "&TEXT({K["secMax"]},"0.0")&"x"', None),
    ("Clean Earth avant synergies, écart au plus cher du secteur (x EBITDA)", K["ceVsSec"], '0.0"x"'),
    ("Pairs Moody's : FFO / dette nette de Veolia moins celui de Suez (Baa2, perspective négative), en points", A["peerGapSuez"], NF_P),
    ("Objectifs financiers annoncés et tenus depuis 2020 (Impact 2023, guidance 2025)", CR["credFin"], None),
    ("Hybrides : tranche dont la réinitialisation tombe avant fin 2027 (septembre 2026), en M EUR", A["hybBefore28"], NF_M),
    ("Macro : marge 2027 perdue si 2027 revit l'énergie du S1 2026, en M EUR", MA["energyH1"], NF_M),
    ("Macro : dette nette en plus si le dollar monte de 10 %, en M EUR", MA["usd10"], NF_M),
    ("Taxonomie 2025 : part du capex du groupe éligible et aligné (DEU p.250)", TX["cxAlignedShare"], NF_P),
    ("Taxonomie 2025 : déchets dangereux, part du capex aligné du groupe", TX["hwShare"], NF_P),
    ("Programme de cessions au multiple central, en part de l'EBITDA des déchets solides", C["progShareSw"], NF_P),
    ("Années nécessaires pour céder 2 Md EUR au rythme des cessions 2022-2025", C["needYears"], NF_D2),
    ("Liquidités / flux contractuels de dette 2026 (DEU p.421)", MR["cov26"], NF_X),
    ("Surcoût d'intérêts annuel si les souches 2027-2028 sont refinancées au taux de juin 2025, en M EUR", f"{q('Échéancier')}!$C${M['extraInt']}", NF_M),
    ("FFO 2025 reconstitué depuis le tableau de flux, en M EUR (Moody's publie 5 160)", A["ffoRec25"], NF_M),
    ("Contrôles bloquants", VSTAT, None),
    ("Scénario", HSTAT, None),
]
ROOTS = []   # (libellé, feuille, cellule) des résultats chiffrés : racines du traçage des chiffres critiques
for lab, ref, nf in res:
    if ref not in (VSTAT, HSTAT):
        ROOTS.append((lab, ws.title, f"C{r}"))
    put(ws, f"B{r}", lab)
    put(ws, f"C{r}", f"={ref}", color=GREEN, nf=nf, bold=True, align="right")
    r += 1
r += 1
put(ws, f"B{r}", "Comment lire", bold=True, size=12); r += 1
legend = [
    ("Bleu", "nombre saisi : soit un chiffre publié (onglet Entrées, avec fichier, page et URL), soit une borne d'hypothèse",
     Font(name=F, color=BLUE), None),
    ("Noir", "formule", Font(name=F, color=BLACK), None),
    ("Vert", "lien vers un autre onglet", Font(name=F, color=GREEN), None),
    ("Jaune", "hypothèse du groupe : seules ces cellules se modifient (onglet Hypothèses, et cadres à remplir)", Font(name=F),
     YELLOW),
    ("Saumon", "dans Trajectoire : la seule hypothèse qui change dans ce scénario", Font(name=F), CHANGED),
    ("Actif", "colonne Actif de l'onglet Hypothèses : Base, Bas ou Haut par hypothèse ; tous les onglets lisent la « Valeur active »",
     Font(name=F, color=BLUE), YELLOW),
]
for k, v, fnt, fill in legend:
    c = ws[f"B{r}"]; c.value = k; c.font = fnt
    if fill:
        c.fill = fill
    put(ws, f"C{r}", v, wrap=True); ws.merge_cells(f"C{r}:F{r}")
    r += 1
r += 1
put(ws, f"B{r}", "Les six rôles", bold=True, size=12); r += 1
roles = [
    ("1. Périmètre et sources", "Entrées, Cessions §A-B", "L'enveloppe annoncée est nette des cessions ; Clean Earth en sort."),
    ("2. La contrainte", "Levier, Pont de dette, Échéancier", "Définition de Veolia, ratios publiés reproduits, trois lectures du départ au 30/06/2026 ; "
     "seuils des agences, pont du FFO, mur de refinancement."),
    ("3. Capacité et sensibilité", "Trajectoire, Sensibilité, Cessions §C, Pont EBITDA", "FCF, calendrier des cessions, ce qui bouge le plus la marge."),
    ("4. L'écart et les comparables", "Segments, Pont EBITDA, Booster", "D'où viennent les 8 Md€ (un plan d'efficacité) ; volumes face à 9 et 10 Mt ; Clean Harbors en comparable ; Clean Earth vu du vendeur."),
    ("5. Le coût ESG", "ESG", "Provisions de fermeture ; ce que l'affectation du prix de Clean Earth ne montre pas encore."),
    ("6. Cibles, puis synthèse", "Cibles", "Taille maximale selon le multiple ; univers de cibles à remplir."),
]
header(ws, r, ["", "Rôle", "Onglets", "Ce que l'onglet tranche"], start=1); r += 1
for a, b, c in roles:
    put(ws, f"B{r}", a, bold=True); put(ws, f"C{r}", b); put(ws, f"D{r}", c, wrap=True); ws.merge_cells(f"D{r}:F{r}")
    ws.row_dimensions[r].height = 28
    r += 1
r += 1
put(ws, f"B{r}", "Règles et limites", bold=True, size=12); r += 1
rules = [
    "Aucun chiffre publié n'est saisi ailleurs que dans Entrées. Les seules constantes dans les formules sont 1 000 (Md → M) et 12 (mois).",
    "Le levier est celui de Veolia : dette financière nette de clôture, IFRS 16 incluse, sur EBITDA IFRS 16 inclus. Les hybrides "
    "(4,1 Md€) sont en capitaux propres (IAS 32.11).",
    "Les cessions retirent leur EBITDA sur l'année entière où elles sont encaissées : c'est prudent.",
    "Le change sur la dette en dollars n'est pas modélisé ; aucun dividende n'est supposé versé au S2.",
    "Le levier 2026 « publié » compte 7 mois de Clean Earth ; la ligne pro forma en compte 12. Veolia ne dit pas laquelle elle retient.",
    "Le sélecteur de scénario (Hypothèses, colonne Actif) change tout le classeur ; les colonnes de Trajectoire restent des variations "
    "une à une autour du scénario actif. Les noms h_<code> (ex. h_s27) pointent sur la valeur active de chaque hypothèse.",
    f"Registre : {len(rows)} chiffres versés dans dataroom.is42.fr, état au {__import__('datetime').date.today().strftime('%d/%m/%Y')}. La relecture par un tiers se suit dans Vérifications "
    "(dernière ligne) et sur la page Relecture de la dataroom.",
]
for t in rules:
    put(ws, f"B{r}", f"•  {t}", wrap=True); ws.merge_cells(f"B{r}:F{r}")
    ws.row_dimensions[r].height = 28
    r += 1
widths(ws, {"A": 3, "B": 60, "C": 24, "D": 30, "E": 30, "F": 30})

for s in wb.worksheets:
    s.sheet_view.showGridLines = False
    s.page_setup.orientation = "landscape"
    s.page_setup.paperSize = s.PAPERSIZE_A4
    s.sheet_properties.pageSetUpPr.fitToPage = True
    s.page_setup.fitToWidth = 1
    s.page_setup.fitToHeight = 0
    for row in s.iter_rows():
        for cell in row:
            if cell.font and cell.font.name != F:
                cell.font = Font(name=F, color=cell.font.color, bold=cell.font.bold, size=cell.font.size or 10,
                                 italic=cell.font.italic)

OUT.parent.mkdir(parents=True, exist_ok=True)
wb.save(OUT)
import json as _json
from datetime import date as _date


# ---------------------------------------------------------------- chiffres critiques
# Remonter les formules depuis chaque résultat de Lisez-moi jusqu'aux lignes d'Entrées : les chiffres
# du registre dont dépend un résultat sont ceux qu'il faut relire d'abord.
from openpyxl.formula import Tokenizer
from openpyxl.utils.cell import range_boundaries, coordinate_from_string, column_index_from_string

ROW_TO_ID = {row: fid for fid, row in ENTREE_ROW.items()}


def _cells_of(ref: str, here: str):
    ref = ref.replace("$", "")
    sheet = here
    if "!" in ref:
        sheet, ref = ref.rsplit("!", 1)
        sheet = sheet.strip("'")
    if sheet not in wb.sheetnames:
        return []
    if ":" in ref:
        try:
            c1, r1, c2, r2 = range_boundaries(ref)
        except ValueError:
            return []
        if None in (c1, r1, c2, r2) or (c2 - c1 + 1) * (r2 - r1 + 1) > 5000:
            return []
        return [(sheet, f"{L(c)}{rr}") for c in range(c1, c2 + 1) for rr in range(r1, r2 + 1)]
    try:
        coordinate_from_string(ref)
    except ValueError:
        return []
    return [(sheet, ref)]


def trace(sheet: str, cell: str, seen: set) -> set:
    """Ids d'Entrées atteints depuis (feuille, cellule), en suivant les références des formules."""
    found, stack = set(), [(sheet, cell)]
    while stack:
        sh, ce = stack.pop()
        if (sh, ce) in seen:
            continue
        seen.add((sh, ce))
        if sh == "Entrées":
            col, row = coordinate_from_string(ce)
            if col == "C" and row in ROW_TO_ID:
                found.add(ROW_TO_ID[row])
            continue
        value = wb[sh][ce].value
        if not (isinstance(value, str) and value.startswith("=")):
            continue
        try:
            tokens = Tokenizer(value).items
        except Exception:
            continue
        for tok in tokens:
            if tok.type == "OPERAND" and tok.subtype == "RANGE":
                stack.extend(_cells_of(tok.value, sh))
    return found


CRITICAL: dict[str, list[str]] = {}
for lab, sh, ce in ROOTS:
    for fid in trace(sh, ce, set()):
        CRITICAL.setdefault(fid, []).append(lab)
OUT.with_name("model-critical.json").write_text(_json.dumps({
    "generated": _date.today().isoformat(), "model": OUT.name,
    "roots": [lab for lab, _, _ in ROOTS],
    "ids": {fid: CRITICAL[fid] for fid in sorted(CRITICAL, key=lambda s: (s[0], int(s[1:].split("-")[0])))},
}, ensure_ascii=False, indent=1), encoding="utf-8")
print("chiffres critiques :", len(CRITICAL), "lignes du registre alimentent", len(ROOTS), "résultats")

OUT.with_name("model-def.json").write_text(_json.dumps({
    "generated": _date.today().isoformat(), "model": OUT.name, "seed": MD.SEED, "n_draws": SIM_N,
    "varied": VARIED, "fixed": [c for c in INPUT_ORDER if c not in VARIED],
    "inputs": {code: {"label": sheets["Hypothèses"][f"B{H[code]}"].value, "unit": sheets["Hypothèses"][f"F{H[code]}"].value,
                      "row": H[code]} for code in INPUT_ORDER},
    "consts": {name: {"id": CONST_IDS[name], "label": MD.CONSTS[name],
                      "value": parse(next(x for x in rows if x["id"] == CONST_IDS[name])["value"])} for name in CONST_IDS},
    "calc": [{"key": k, "label": lab, "template": tpl, "unit": unit, "kind": kind, "bold": bold}
             for k, lab, tpl, unit, kind, bold in MD.CALC if lab is not None],
    "outputs": [{"key": k, "label": lab, "kind": kind} for k, lab, kind in MD.OUTPUTS],
    "factor": {"weight": MD.FACTOR_WEIGHT, "loadings": {k: MD.FACTOR_LOADINGS.get(k, 0) for k in VARIED},
               "note": "uniforms[i][0] est le tirage du facteur commun, puis une uniforme par hypothèse variée, dans l'ordre de varied"},
    "uniforms": UNIFORMS,
    "bascules": {"sheet": "Financement", "first_row": FN["bk0"], "last_row": FN["bk1"], "value_cols": [L(11 + k) for k in range(len(CONCL))],
                 "conclusions": [cid for cid, _l, _m in CONCL], "size_cell": FN["S"].replace("$", ""), "lev26max_cell": FN["lev26max"].replace("$", "")},
    "trajectoire": {"columns": [{"col": col, "name": name, "code": code, "side": side} for col, name, code, side in cols],
                    "input_rows": IN, "calc_rows": {k: v for k, v in ROW.items()}},
}, ensure_ascii=False), encoding="utf-8")

OUT.with_name("model-uses.json").write_text(_json.dumps({
    "generated": _date.today().isoformat(), "model": OUT.name,
    "ids": sorted({u["id"] for u in USED}, key=lambda s: (s[0], int(s[1:].split("-")[0]))),
    "lookups": USED,
}, ensure_ascii=False, indent=1), encoding="utf-8")
print("écrit", OUT, "—", len(rows), "entrées,", len(cols), "scénarios")
