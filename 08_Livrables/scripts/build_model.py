"""Construit modele-greenup-2027.xlsx à partir du registre de la dataroom.

Chaque nombre bleu de l'onglet Entrées est une ligne du registre (fichier, page, URL).
Les cellules jaunes sont les hypothèses du groupe. Tout le reste est formule.
"""
import csv
import re
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L

HERE = Path(__file__).parent
REG = HERE / "register.csv"
OUT = Path(sys.argv[1] if len(sys.argv) > 1 else "/home/ubuntu/capstone-livrables/modele-greenup-2027.xlsx")

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
    if re.fullmatch(r"-?\d{1,3}(,\d{3})+", v):
        return float(v.replace(",", ""))
    if re.fullmatch(r"-?\d+,\d+", v):
        return float(v.replace(",", "."))
    try:
        return float(v)
    except ValueError:
        return value.strip()  # libellé publié sans nombre (« high teens »)


rows = list(csv.DictReader(REG.open(encoding="utf-8")))
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
for name in ["Entrées", "Hypothèses", "Levier", "Pont de dette", "Cessions", "Trajectoire",
             "Sensibilité", "Segments", "Booster", "ESG", "Cibles", "Vérifications"]:
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
header(ws, 4, ["Code", "Hypothèse", "Base", "Bas", "Haut", "Unité", "Justification et ancrage", "Réf."])
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


g_lo, id_glo = E("Guidance 2026 : croissance organique de l'EBITDA, bas", "2026")
g_hi, id_ghi = E("Guidance 2026 : croissance organique de l'EBITDA, haut", "2026")
cni, id_cni = E("Guidance 2026 : croissance du résultat net courant", "2026")
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
mo25h, id_mo25h = E("Moody's : FFO / dette nette ajustée", "FY2025")
mond25h, id_mond25h = E("Moody's : dette nette ajustée", "FY2025")
hyp(r, "ffo", "FFO (mesure des agences) en % de l'EBITDA", f"={mo25h}/100*{mond25h}*1000/{eb25}", f"=C{r}*0.95", f"=C{r}*1.03", "%",
    "FFO 2025 implicite chez Moody's (20,3 % × 25,4 Md€ = 5,2 Md€) rapporté à l'EBITDA 2025 publié. Le bas couvre le surcoût d'intérêts de la dette Clean Earth.",
    f"{id_mo25h}, {id_mond25h}, {id_eb25}", NF_P); r += 1
VARIED = ["g26", "g27", "ceEb", "syn", "fcfH2", "conv", "gDiv", "s26", "s27", "mDisp", "tuck", "ffo"]
r += 1
put(ws, f"A{r}", "Hypothèses fixes (non soumises à la sensibilité)", bold=True); r += 1
hyp(r, "fx", "Change USD par EUR", f"={ce_usd}/{ce_eur}", None, None, "USD/EUR",
    "Taux implicite du prix de Clean Earth à la clôture (2 989 M$ = 2 542 M€). Le change sur la dette en dollars n'est pas modélisé.",
    f"{id_ceusd}, {id_ceeur}", '0.0000'); r += 1
hyp(r, "minDiv", "Dividendes versés aux minoritaires en 2027", f"=-{div_t26}-{div_sh26}", None, None, "M EUR",
    "Écart du S1 2026 entre dividendes totaux (1 394) et dividende Veolia (1 099), supposé constant.",
    f"{id_divt26}, {id_divsh26}", NF_M); r += 1
hyp(r, "mTuck", "Multiple VE / EBITDA des tuck-ins", f"={ce_mult}", None, None, "x",
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
widths(ws, {"A": 8, "B": 44, "C": 11, "D": 11, "E": 11, "F": 9, "G": 78, "H": 22})
ws.freeze_panes = "C5"
for rr in range(5, r + 1):
    ws.row_dimensions[rr].height = 30


def HY(code, col="C"):
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
mond25, i_mond25 = E("Moody's : dette nette ajustée", "FY2025")
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
line(ws, r, "adjGap", "Écart entre dette ajustée Moody's 2025 et dette nette publiée (IFRS 16 incl.)", f"={A['mond25']}-{nfd25}", "M EUR", f"{i_mond25} − {i_nfd25} : retraitements d'agence (hybrides à 50 %, pensions, etc.)", NF_M); r += 1
put(ws, f"B{r}", "Lecture : les agences ne regardent pas le 3x de Veolia mais FFO / dette ajustée. Au pic de dette 2026 (~29 Md€), tenir 18 % "
    "demande ~5,2 Md€ de FFO, à peu près le FFO 2025 implicite : la marge est nulle en 2026 et ne revient qu'avec les cessions. "
    "C'est la limite qui mord en premier, avant le plafond de 3x.", color=GREY, italic=True, wrap=True)
ws.merge_cells(f"B{r}:E{r}"); ws.row_dimensions[r].height = 44; r += 2

put(ws, f"A{r}", "C", bold=True); put(ws, f"B{r}", "Le plafond selon la définition retenue (scénario central, fin 2027)", bold=True); r += 1
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
ws.merge_cells(f"B{r}:E{r}"); ws.row_dimensions[r].height = 42; r += 1
widths(ws, {"A": 4, "B": 60, "C": 16, "D": 12, "E": 50})

# ================================================================ Trajectoire
ws = sheets["Trajectoire"]
title(ws, "Rôles 3 et 6 — Trajectoire 2026-2027 et marge de manœuvre sous 3x",
      "Colonne D = scénario central. Chaque colonne suivante change UNE hypothèse (cellule saumon). "
      "Les deux dernières combinent tous les bas défavorables, puis tous les hauts favorables.")
cols = [("D", "Base", None, None)]
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
INPUT_ORDER = VARIED + ["fx", "minDiv", "mTuck", "tuckH2", "other", "months", "synC", "cap"]
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
calc = [
    ("A26", "EBITDA organique 2026 (hors Clean Earth, hors cessions)", "={eb25}*(1+{c}{g26})", "M EUR", NF_M, False),
    ("CE26", "Clean Earth en 2026 (mois consolidés)", "={c}{ceEb}*{c}{months}/12/{c}{fx}", "M EUR", NF_M, False),
    ("D26", "Cessions encaissées au S2 2026", "={sign26}*{c}{s26}*1000", "M EUR", NF_M, False),
    ("DE26", "EBITDA cédé en 2026 (année pleine, prudent)", "=-{c}{D26}/{c}{mDisp}", "M EUR", NF_M, False),
    ("EB26", "EBITDA 2026", "={c}{A26}+{c}{CE26}+{c}{DE26}", "M EUR", NF_M, True),
    ("NFD26", "Dette financière nette au 31/12/2026",
     "={nfdh126}-{c}{fcfH2}+{c}{tuckH2}*1000-{c}{D26}+{c}{other}", "M EUR", NF_M, True),
    ("LEV26", "Levier fin 2026", "={c}{NFD26}/{c}{EB26}", "x", NF_X, True),
    ("LEV26PF", "Levier fin 2026, Clean Earth pro forma 12 mois",
     "={c}{NFD26}/({c}{EB26}+{c}{ceEb}*(12-{c}{months})/12/{c}{fx})", "x", NF_X, False),
    ("SEP", None, None, None, None, None),
    ("B27", "EBITDA 2027 avant synergies, tuck-ins et cessions 2027",
     "=({c}{A26}+{c}{DE26}+{c}{ceEb}/{c}{fx})*(1+{c}{g27})", "M EUR", NF_M, False),
    ("SYN27", "Synergies Clean Earth réalisées", "={synRR}*{c}{syn}/{c}{fx}", "M EUR", NF_M, False),
    ("TE27", "EBITDA apporté par les tuck-ins 2027", "={c}{tuck}*1000/{c}{mTuck}", "M EUR", NF_M, False),
    ("D27", "Cessions encaissées en 2027",
     "=({sign26}*(1-{c}{s26})+MAX(0,{prog}-{closed}-{sign26})*{c}{s27})*1000", "M EUR", NF_M, False),
    ("DE27", "EBITDA cédé en 2027 (année pleine, prudent)", "=-{c}{D27}/{c}{mDisp}", "M EUR", NF_M, False),
    ("EB27", "EBITDA 2027", "={c}{B27}+{c}{SYN27}+{c}{TE27}+{c}{DE27}", "M EUR", NF_M, True),
    ("FCF27", "Cash-flow libre net 2027", "={c}{conv}*{c}{EB27}-{c}{synC}/{c}{fx}", "M EUR", NF_M, False),
    ("DIV27", "Dividendes versés en 2027", "={divsh26}*(1+{c}{gDiv})+{c}{minDiv}", "M EUR", NF_M, False),
    ("NFD27", "Dette financière nette au 31/12/2027",
     "={c}{NFD26}-{c}{FCF27}+{c}{DIV27}+{c}{tuck}*1000-{c}{D27}+{c}{other}", "M EUR", NF_M, True),
    ("LEV27", "Levier fin 2027", "={c}{NFD27}/{c}{EB27}", "x", NF_X, True),
    ("HEAD27", "Marge de manœuvre fin 2027 (plafond × EBITDA − DFN)", "={c}{cap}*{c}{EB27}-{c}{NFD27}", "M EUR", NF_M,
     True),
    ("MAXACQ", "Acquisition maximale au multiple des tuck-ins",
     "=IF({c}{mTuck}>{c}{cap},MAX(0,{c}{HEAD27})/(1-{c}{cap}/{c}{mTuck}),0)", "M EUR", NF_M, True),
    ("GAP8", "EBITDA 2027 moins l'objectif ≥ 8 Md€", "={c}{EB27}-{tgt}*1000", "M EUR", NF_M, False),
    ("SEP2", None, None, None, None, None),
    ("ADJ26", "Dette nette ajustée par les agences fin 2026 (DFN + écart Moody's 2025)", "={c}{NFD26}+({mond25}*1000-{nfd25})", "M EUR", NF_M, False),
    ("FFO26", "FFO 2026 (hypothèse ffo × EBITDA 2026)", "={c}{ffo}*{c}{EB26}", "M EUR", NF_M, False),
    ("RATIO26", "FFO / dette ajustée fin 2026", "={c}{FFO26}/{c}{ADJ26}", "%", NF_P, True),
    ("HEADSP26", "Marge de dette fin 2026 sous le seuil S&P (FFO / 18 % − dette ajustée)", "={c}{FFO26}/({sptrig}/100)-{c}{ADJ26}", "M EUR", NF_M, True),
    ("ADJ27", "Dette nette ajustée fin 2027", "={c}{NFD27}+({mond25}*1000-{nfd25})", "M EUR", NF_M, False),
    ("FFO27", "FFO 2027", "={c}{ffo}*{c}{EB27}", "M EUR", NF_M, False),
    ("RATIO27", "FFO / dette ajustée fin 2027", "={c}{FFO27}/{c}{ADJ27}", "%", NF_P, True),
    ("HEADSP27", "Marge de dette fin 2027 sous le seuil S&P 18 %", "={c}{FFO27}/({sptrig}/100)-{c}{ADJ27}", "M EUR", NF_M, True),
    ("BIND27", "Marge contraignante fin 2027 (la plus petite des deux)", "=MIN({c}{HEAD27},{c}{HEADSP27})", "M EUR", NF_M, True),
    ("WHICH27", "Contrainte qui mord en premier", '=IF({c}{HEADSP27}<{c}{HEAD27},"agences : FFO / dette ≥ 18 %","Veolia : levier ≤ 3x")', "", None, True),
    ("MAXACQB", "Acquisition maximale fin 2027 sous la contrainte qui mord, au multiple des tuck-ins",
     "=IF({c}{mTuck}>{c}{cap},MAX(0,{c}{BIND27})/(1-{c}{cap}/{c}{mTuck}),0)", "M EUR", NF_M, True),
]
tgt, i_tgt = E("Objectif GreenUp : EBITDA", "2027")
syn_rr, i_syn = E("Clean Earth : synergies de coûts", "année 4")
put(ws, f"B{RES0-1}", "Calcul", bold=True)
ROW = {}
rr = RES0
for key, *_ in calc:
    ROW[key] = rr; rr += 1
refs = dict(eb25=eb25, sign26=sign26, nfdh126=nfdh126, prog=prog, closed=closed, divsh26=div_sh26, tgt=tgt, synRR=syn_rr,
            mond25=mond25, nfd25=nfd25, sptrig=sptrig)
for key, label, formula, unit, nf, bold in calc:
    rr = ROW[key]
    if label is None:
        continue
    put(ws, f"A{rr}", key, color=GREY)
    put(ws, f"B{rr}", label, bold=bold)
    put(ws, f"C{rr}", unit, color=GREY)
    for col, *_ in cols:
        fmt = {**refs, "c": col, **{k: v for k, v in IN.items()}, **{k: v for k, v in ROW.items()}}
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
for row in sheets["Levier"].iter_rows():
    for cell in row:
        if isinstance(cell.value, str) and ("$NFD27" in cell.value or "$EB27" in cell.value):
            cell.value = cell.value.replace("$D$NFD27", f"$D${ROW['NFD27']}").replace("$D$EB27", f"$D${ROW['EB27']}")
widths(ws, {"A": 8, "B": 54, "C": 9, **{c: 11 for c, *_ in cols}})
ws.column_dimensions[UNFAV].width = 13
ws.column_dimensions[FAV].width = 13
ws.freeze_panes = "D5"
T = {k: f"{q('Trajectoire')}!$D${v}" for k, v in ROW.items()}

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
         ("Séché Environnement", "France", "spécialiste DD européen, coté à Paris"),
         ("Enviri (ex-Harsco)", "États-Unis", "vendeur de Clean Earth : ses comptes donnent l'historique de la cible"),
         ("Republic Services (US Ecology)", "États-Unis", "généraliste ; DD via US Ecology — mix différent"),
         ("Remondis / SARP", "Allemagne", "non coté : comparable métier, pas de multiple de marché")]
sec_rev, i_srev = E("Séché : chiffre d'affaires contributif", "FY2025")
sec_eb, i_seb = E("Séché : EBITDA", "FY2025")
sec_m, i_sm = E("Séché : marge d'EBITDA", "FY2025")
for name, ctry, why in peers:
    put(ws, f"B{r}", name); put(ws, f"C{r}", ctry)
    for col in "DEFGHIJ":
        put(ws, f"{col}{r}", None, fill=YELLOW)
    if name.startswith("Séché"):
        put(ws, f"E{r}", f"={sec_rev}", color=GREEN, nf=NF_M1)
        put(ws, f"F{r}", f"={sec_eb}", color=GREEN, nf=NF_M1)
        put(ws, f"I{r}", "01_Financial/26-02_cp-bn-25_en.pdf", color=GREY)
        put(ws, f"J{r}", f"{i_srev} p.1, {i_seb} p.3", color=GREY)
        why += " — 2025, M EUR ; levier 2,3x calculé sur la dette moyenne selon la doc bancaire, pas comme Veolia"
    elif name.startswith("Clean Harbors"):
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
put(ws, f"B{r}", "Pour mémoire : marge d'EBITDA de Séché publiée")
put(ws, f"G{r}", f"={sec_m}/100", color=BLACK, nf=NF_P); put(ws, f"K{r}", i_sm, color=GREY); r += 1
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

put(ws, f"A{r}", "C", bold=True); put(ws, f"B{r}", "Univers de cibles — à remplir (phase F)", bold=True); r += 1
header(ws, r, ["", "Cible", "Pays", "Activité DD", "VE estimée (M EUR)", "Multiple VE / EBITDA", "Tient dans la marge centrale ?",
               "Fichier", "Page"]); r += 1
put(ws, f"B{r}", "Exemple : Clean Earth (pour calibrer)", italic=True)
put(ws, f"C{r}", "États-Unis", italic=True); put(ws, f"D{r}", "TSDF, 82 sites", italic=True)
put(ws, f"E{r}", f"=-{ce_nfd}", color=GREEN, nf=NF_M); put(ws, f"F{r}", f"={ce_mult}", color=GREEN, nf='0.0"x"')
put(ws, f"G{r}", f'=IF(AND(ISNUMBER(E{r}),ISNUMBER(F{r})),IF(E{r}<=IF(F{r}>{HY("cap")},MAX(0,$D${HR})/(1-{HY("cap")}/F{r}),0),"oui","non"),"")')
put(ws, f"H{r}", "01_Financial/veolia_finance_amendment_urd_2025.pdf", color=GREY); put(ws, f"I{r}", 25, color=GREY)
r += 1
for _ in range(8):
    for col in "BCDEFHI":
        put(ws, f"{col}{r}", None, fill=YELLOW)
    ws[f"E{r}"].number_format = NF_M; ws[f"F{r}"].number_format = '0.0"x"'
    put(ws, f"G{r}", f'=IF(AND(ISNUMBER(E{r}),ISNUMBER(F{r})),IF(E{r}<=IF(F{r}>{HY("cap")},MAX(0,$D${HR})/(1-{HY("cap")}/F{r}),0),"oui","non"),"")')
    r += 1
put(ws, f"B{r}", "Cellules jaunes : une ligne par cible, VE et multiple sourcés dans la dataroom. La colonne G répond seule.",
    color=GREY, italic=True)
widths(ws, {"A": 4, "B": 46, "C": 14, "D": 16, "E": 16, "F": 16, "G": 18, "H": 46, "I": 7})

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
    ("FFO / dette ajustée 2026 du modèle dans l'attente de Moody's (18-19 %, ± 2 pts)", T["RATIO26"], 0.185, 0.02, NF_P, "Bloquant"),
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
put(ws, f"B{r}", "Résultats du scénario central", bold=True, size=12); r += 1
res = [
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
    ("Contrôles bloquants", VSTAT, None),
]
for lab, ref, nf in res:
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
    ("2. La contrainte", "Levier, Pont de dette", "Définition de Veolia, ratios publiés reproduits, trois lectures du départ au 30/06/2026."),
    ("3. Capacité et sensibilité", "Trajectoire, Sensibilité, Cessions §C", "FCF, calendrier des cessions, ce qui bouge le plus la marge."),
    ("4. L'écart et les comparables", "Segments, Booster", "D'où viennent les 8 Md€ ; volumes face à 9 et 10 Mt ; Clean Harbors et Séché en comparables ; Clean Earth vu du vendeur."),
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
    "Registre : chiffres versés dans dataroom.is42.fr le 1er octobre 2026. Aucun n'a encore été relu par un tiers (Vérifications, ligne 14).",
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
OUT.with_name("model-uses.json").write_text(_json.dumps({
    "generated": _date.today().isoformat(), "model": OUT.name,
    "ids": sorted({u["id"] for u in USED}, key=lambda s: (s[0], int(s[1:].split("-")[0]))),
    "lookups": USED,
}, ensure_ascii=False, indent=1), encoding="utf-8")
print("écrit", OUT, "—", len(rows), "entrées,", len(cols), "scénarios")
