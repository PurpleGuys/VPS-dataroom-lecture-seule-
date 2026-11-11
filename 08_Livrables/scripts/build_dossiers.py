"""Génère les dossiers des rôles (HTML) et report-data.json depuis les textes du rapport, les registres et le classeur.

Les dossiers se DÉRIVENT des textes de l'édition courte (08_Livrables/texte/*.md) : même dérivation que les fiches d'oral
(build_fiches.py, importé plus bas), même contrôle : une partie absente ou un jeton inconnu se marque dans la page, dans
un bandeau de tête et sur la sortie d'erreur, et tout s'écrit quand même (code 0) ; avec « --strict » ou
DATAROOM_FICHES_STRICT=1, le script s'arrête (code 1) sans rien écrire.
Aucune phrase ni aucun chiffre des dossiers n'est tapé ici. Ce script écrit aussi fourchette-capacite.md,
questions-classees.md et report-data.json (forme inchangée : le rapport et les autres scripts le lisent), dont les
valeurs viennent de register.csv ou d'une cellule du classeur recalculé.
"""
import csv
import os
import html
import re
from pathlib import Path

from openpyxl import load_workbook

HERE = Path(__file__).parent
# Où sont les registres et où écrire : la chaîne (dataroom livrables) passe DATAROOM_ADMIN et DATAROOM_LIVRABLES ;
# à défaut, le vault autour du script (08_Livrables/scripts → ../../00_Admin), puis une copie à côté du script.
ADMIN = Path(os.environ.get("DATAROOM_ADMIN") or (HERE.parent.parent / "00_Admin"))
LIV = Path(os.environ.get("DATAROOM_LIVRABLES") or HERE.parent)
def admin_file(name):
    return ADMIN / name if (ADMIN / name).exists() else HERE / name
REG = list(csv.DictReader(admin_file("register.csv").open(encoding="utf-8")))
DEALS = list(csv.DictReader(admin_file("deals.csv").open(encoding="utf-8"))) if admin_file("deals.csv").exists() else []
WB = load_workbook(os.environ.get("DATAROOM_MODEL") or LIV / "modele-greenup-2027.xlsx", data_only=True)
OUT = LIV / "dossiers-sujet-2.html"
NB = " "
import datetime as _dt


def fig(prefix, period, unit=None):
    hits = [r for r in REG if r["label"].startswith(prefix) and r["period"] == period
            and (unit is None or r["unit"] == unit)]
    if len(hits) != 1:
        raise SystemExit(f"{len(hits)} lignes pour {prefix!r} {period!r}")
    return hits[0]


def mv(sheet, label, col="C"):
    ws = WB[sheet]
    for row in ws.iter_rows(min_col=2, max_col=2):
        c = row[0]
        if isinstance(c.value, str) and c.value.strip().startswith(label):
            return ws[f"{col}{c.row}"].value
    raise SystemExit(f"{sheet}: pas de ligne {label!r}")


def cell(sheet, ref):
    return WB[sheet][ref].value


def fr(v, d=0, unit=""):
    if v is None:
        return "—"
    s = f"{abs(v):,.{d}f}".replace(",", NB).replace(".", ",")
    s = ("−" if v < 0 else "") + s
    return s + (f"{NB}{unit}" if unit else "")


def x(v):
    return fr(v, 2) + "x"


def pct(v, d=1):
    return fr(v * 100, d) + NB + "%"


def esc(s):
    return html.escape(str(s), quote=True)


def parse(value):
    v = value.replace("\u00a0", "").replace("\u202f", "").replace(" ", "").strip()
    if re.fullmatch(r"[-+]?\d{1,3}(,\d{3})+(\.\d+)?", v):
        return float(v.replace(",", "")), (len(v.split(".")[1]) if "." in v else 0)
    if re.fullmatch(r"-?\d+,\d+", v):
        return float(v.replace(",", ".")), len(v.split(",")[1])
    try:
        return float(v), (len(v.split(".")[1]) if "." in v else 0)
    except ValueError:
        return value.strip(), 0


UNITS = {"M EUR": "M€", "Md EUR": "Md€", "Md EUR / an": "Md€/an", "M USD": "M$", "Md USD": "Md$", "% / an": "%/an"}


def shown(r):
    v, d = parse(r["value"])
    if isinstance(v, str):
        return f"« {esc(v)} »"
    unit = UNITS.get(r["unit"], r["unit"])
    return fr(v, d) + ("x" if unit == "x" else NB + unit)


# Une valeur du registre telle qu'un lecteur français l'attend (« 7,050 » imprimé à l'anglaise ne doit pas se lire
# sept virgule zéro cinq) : 19 657 M€, 0,5 Md€/an, 2,8x, 5,1 % — espace fine insécable (U+202F) entre les milliers et
# avant l'unité. Même logique que french_number() et display_value() de src/dataroom/report.py, que les scripts du vault
# n'importent pas.
UNITS_FR = {"M EUR": "M€", "Md EUR": "Md€", "k EUR": "k€", "EUR": "€", "M USD": "M$", "Md USD": "Md$", "USD": "$",
            "M GBP": "M£", "GBP": "£", "Md GBP": "Md£", "Md EUR / an": "Md€/an", "M EUR / an": "M€/an",
            "Md EUR par an": "Md€/an", "% / an": "%/an"}


def _group(integer):
    out = ""
    while len(integer) > 3:
        out = NB + integer[-3:] + out
        integer = integer[:-3]
    return integer + out


def french_number(value):
    """« 25,423 » → « 25 423 », « 5.1 » → « 5,1 », « -200 » → « −200 » ; un texte (« Baa1 Stable ») reste tel quel."""
    v = (value or "").strip()
    compact = v.replace("\u202f", "").replace("\xa0", "").replace(" ", "")
    m = re.fullmatch(r"([-+−]?)(\d{1,3}(?:,\d{3})+)(\.\d+)?", compact)          # 25,423 ou 1,234.5
    if m:
        sign, integer, dec = m.group(1), m.group(2).replace(",", ""), (m.group(3) or "")[1:]
    else:
        m = re.fullmatch(r"([-+−]?)(\d+)(?:[.,](\d+))?", compact)              # 5.1, 0.2, -200, 3,8, 24548
        if not m or (" " in v and not re.fullmatch(r"[-+−]?\d{1,3}(?: \d{3})+(?:,\d+)?", v)):
            return v
        sign, integer, dec = m.group(1), m.group(2), m.group(3) or ""
    sign = "−" if sign in ("-", "−") else sign
    return sign + _group(integer) + ("," + dec if dec else "")


def display_value(value, unit):
    """« valeur unité » en notation française : 25 423 M€, 5,1 %, 10,4x, A-2/BBB."""
    unit = (unit or "").strip()
    if unit == "année" and re.fullmatch(r"\d{4}", str(value).strip()):
        return str(value).strip()   # une année ne se groupe pas : 2019, pas 2 019
    number = french_number(value)
    if not unit or not re.search(r"\d", number) or re.search(r"[A-Za-z]", number):
        return number
    if unit == "x":
        return number + "x"
    if unit == "%":
        return number + NB + "%"
    return number + NB + UNITS_FR.get(unit, unit)


def N(r):
    return parse(r["value"])[0]


def P(r, d=None):
    v, dd = parse(r["value"])
    return fr(v, dd if d is None else d)



def plain(s):
    """HTML → text, for the Word report. Les espaces insécables (U+00A0, U+202F : milliers, avant % et les unités)
    restent insécables : sans elles, le PDF coupe « 24 » et « % des futurs » sur deux lignes."""
    return re.sub(r"[^\S\u00a0\u202f]+", " ", html.unescape(re.sub(r"<[^>]+>", "", str(s)))).strip()


REPORT_ROLES: list = []


# ------------------------------------------------------------------ valeurs du modèle
_wt = WB["Trajectoire"]
def traj(label, col="D", exact=False):
    for c in _wt["B"]:
        if isinstance(c.value, str) and (c.value == label if exact else c.value.startswith(label)):
            return _wt[f"{col}{c.row}"].value
    raise SystemExit(f"Trajectoire : pas de ligne {label!r}")
def _colname(i):
    return chr(ord("A") + i) if i < 26 else chr(ord("A") + i // 26 - 1) + chr(ord("A") + i % 26)
_cols = [c.value for c in _wt[4]]
UNFC = _colname(_cols.index("Défavorable combiné")); FAVC = _colname(_cols.index("Favorable combiné"))
LEV26 = traj("Levier fin 2026", exact=True); LEV26PF = traj("Levier fin 2026, Clean Earth pro forma")
EB27 = traj("EBITDA 2027", exact=True); NFD27 = traj("Dette financière nette au 31/12/2027", exact=True)
LEV27 = traj("Levier fin 2027", exact=True); FCF27 = traj("Cash-flow libre net 2027", exact=True)
DIV27 = traj("Dividendes versés en 2027", exact=True); NFD26 = traj("Dette financière nette au 31/12/2026", exact=True)
EB26 = traj("EBITDA 2026", exact=True)
HEAD = traj("Marge de manœuvre fin 2027"); MAXACQ = traj("Acquisition maximale au multiple"); GAP8 = traj("EBITDA 2027 moins")
HEADSP26 = traj("Marge de dette fin 2026 sous le seuil"); HEADSP27 = traj("Marge de dette fin 2027 sous le seuil")
HEADSP27_UNF = traj("Marge de dette fin 2027 sous le seuil", UNFC); HEADSP26_UNF = traj("Marge de dette fin 2026 sous le seuil", UNFC)
RATIO26 = traj("FFO / dette ajustée fin 2026"); RATIO27 = traj("FFO / dette ajustée fin 2027")
ADJ26 = traj("Dette nette ajustée par les agences fin 2026")
WHICH27 = traj("Contrainte qui mord"); WHICH27_UNF = traj("Contrainte qui mord", UNFC)
WHICH27_FAV = traj("Contrainte qui mord", FAVC)
def constraint(label):
    """La contrainte qui mord, en bref pour un tableau : « levier ≤ 3x (plafond de Veolia) » → « levier ≤ 3x »."""
    return str(label or "").split(" (")[0]
BIND27 = traj("Marge contraignante"); MAXACQB = traj("Acquisition maximale fin 2027 sous la contrainte")
UNF = mv("Cibles", "Marge de manœuvre fin 2027", "C"); FAV = mv("Cibles", "Marge de manœuvre fin 2027", "E")
_wc = WB["Cibles"]
_hr = next(c.row for c in _wc["B"] if isinstance(c.value, str) and c.value.startswith("Marge contraignante"))
grid = {k: [_wc[f"{c}{r}"].value for c in "BCDE"] for k, r in zip((8, 9, 10, 11), range(_hr + 1, _hr + 5))}
assert abs(grid[9][0] - N(fig("Clean Earth : multiple", "2026e"))) < 1e-9 and grid[8][0] == 8 and grid[10][0] == 12, grid
LTM = mv("Levier", "EBITDA glissant"); LMECH = mv("Levier", "① Levier mécanique")
LPF = mv("Levier", "② Levier pro forma"); LSEAS = mv("Levier", "③ Levier pro forma")
HEAD25 = mv("Levier", "Marge sous le plafond au 31/12/2025")
D50 = mv("Levier", "Hybrides comptés à 50"); D100 = mv("Levier", "Hybrides comptés à 100")
DPROV = mv("Levier", "Hybrides à 50 % + provisions")
_ws = WB["Sensibilité"]
_r0 = next(c.row for c in _ws["A"] if c.value == "Rang") + 1
rank = []
for r in range(_r0, _r0 + 40):
    if not _ws[f"B{r}"].value or not isinstance(_ws[f"E{r}"].value, (int, float)):
        break
    rank.append((_ws[f"B{r}"].value, _ws[f"E{r}"].value))
RANK_LH = {_ws[f"B{r}"].value: (_ws[f"C{r}"].value, _ws[f"D{r}"].value) for r in range(_r0, _r0 + len(rank))}
assert rank[0][1] >= rank[-1][1]
H1_LOW = cell("Sensibilité", "E13")      # marge si s27 = 0
_ws = WB["Sensibilité"]
_e0 = next(c.row for c in _ws["B"] if c.value == "Levier (+100 M€)") + 1
ELAST = {_ws[f"B{r}"].value: _ws[f"C{r}"].value for r in range(_e0, _e0 + 6)}
assert len(ELAST) == 6 and None not in ELAST, ELAST
HW_GAP0 = mv("Booster", "Écart à l'objectif initial")
HW_ILL = mv("Booster", "Illustration")
CE_MARGIN = mv("Booster", "Clean Earth : marge")
EU_G = mv("Booster", "… croissance à périmètre")
AM_G = mv("Booster", "EBITDA Amériques-Asie-Afrique")
GW_PCT = mv("ESG", "Goodwill en % du prix")
PROV_X = mv("ESG", "Provisions de fermeture en tours")
M_PRE = mv("Cibles", "Multiple avant synergies"); M_POST = mv("Cibles", "Multiple après synergies recalculé")
CE_YEARS = mv("Cessions", "Clean Earth en années"); CE_NET = mv("Cessions", "Clean Earth net du programme")
CE_NET_Y = mv("Cessions", "… soit, en années")
CAL = [cell("Cessions", f"C{r}") for r in range(32, 36)]
FCF_H2 = mv("Pont de dette", "Cash-flow libre net S2 2025")
DROP_H2 = mv("Pont de dette", "Baisse de la dette au S2 2025")
FCF_Q4 = mv("Pont de dette", "Cash-flow libre net T4 2025")
DROP_Q4 = mv("Pont de dette", "Baisse de la dette au T4 2025")
RESID25 = mv("Pont de dette", "Flux non détaillés")
CE_GAP = mv("Pont de dette", "Clean Earth : effet sur la dette moins")
MINOR25 = mv("Pont de dette", "dont au moins")
_wv = WB["Vérifications"]
checks_ok = next(_wv[f"F{c.row}"].value for c in _wv["B"] if c.value == "Statut global des contrôles bloquants")
SPENT = {k: mv("Cessions", lab) for k, lab in [("s24", "Sorties nettes 2024"), ("s25", "Sorties nettes 2025"),
         ("sH1", "Sorties nettes 9M 2026"), ("cum", "Sorties nettes cumulées")]}
MAXDEBT_TGT = mv("Levier", "Dette maximale fin 2027 si l'objectif"); MAXDEBT_MOD = mv("Levier", "Dette maximale fin 2027 à l'EBITDA du modèle")
PERHALF = mv("Levier", "Dette libérée par 0,5")
FFO25 = mv("Levier", "FFO 2025 implicite"); FFOREQ = mv("Levier", "FFO requis en 2026"); FFOGAP = mv("Levier", "Marge de FFO en 2026")
DEBT18 = mv("Levier", "Dette ajustée maximale au seuil"); ADJGAP = mv("Levier", "Écart entre dette ajustée")
FFO_REC = mv("Levier", "FFO reconstitué (somme"); FFO_GAP_PCT = mv("Levier", "Écart en % du FFO Moody's")
COV26 = mv("Échéancier", "Liquidités / flux contractuels 2026"); COV26X = mv("Échéancier", "Liquidités / flux 2026 hors")
EXTRA_INT = mv("Échéancier", "Surcoût d'intérêts annuel"); EXTRA_PCT = mv("Échéancier", "en % du FFO 2025")
EFF_SHARE = mv("Pont EBITDA", "Part de l'écart couverte par la seule efficacité")
PROG_EB = mv("Cessions", "EBITDA cédé au multiple central"); PROG_SW = mv("Cessions", "EBITDA cédé en part des seuls déchets solides")
PROG_SH = mv("Cessions", "EBITDA cédé en part des strongholds"); PROG_REVPCT = mv("Cessions", "… en part du chiffre d'affaires du groupe")
SOLD_SUM = mv("Cessions", "Produits de cession cumulés"); SOLD_N = mv("Cessions", "Nombre d'opérations"); SOLD_AVG = mv("Cessions", "Taille moyenne")
SOLD_MAX = mv("Cessions", "La plus grande"); SOLD_PY = mv("Cessions", "Rythme passé"); NEED_AVG = mv("Cessions", "Opérations de taille moyenne")
NEED_YEARS = mv("Cessions", "Années nécessaires au rythme passé")
M_MIN = mv("Cibles", "Le moins cher payé"); M_MAX = mv("Cibles", "Le plus cher payé"); M_MED = mv("Cibles", "Médiane des quatre")
M_SPREAD = mv("Cibles", "Écart avant / après synergies sur Clean Earth")
M_SEC_MIN = mv("Cibles", "Secteur : multiple avant synergies le plus bas"); M_SEC_MAX = mv("Cibles", "Secteur : multiple avant synergies le plus haut")
CE_VS_SEC = mv("Cibles", "Clean Earth avant synergies (recalculé) face")
M_SEC_MED = mv("Cibles", "Secteur : médiane des multiples avant synergies")
M_SEC_TOP = next(_wc[f"B{c.row}"].value for c in _wc["E"] if c.value == M_SEC_MAX and isinstance(_wc[f"B{c.row}"].value, str) and "→" in _wc[f"B{c.row}"].value)
CE_VS_SEC_TXT = (f"Clean Earth avant synergies dépasse de {x(CE_VS_SEC)} le plus cher du secteur" if CE_VS_SEC > 0 else
                 f"Clean Earth avant synergies reste {x(-CE_VS_SEC)} sous le plus cher du secteur ({M_SEC_TOP})")
M_RET = mv("Cibles", "Multiple payé retenu")
T_N = mv("Cibles", "Cibles sourcées"); T_PUB = mv("Cibles", "… dont avec un prix publié")
T_VAL = mv("Cibles", "… dont avec une valeur publiée convertie")
T_FIT = mv("Cibles", "… dont qui tiennent dans la marge centrale"); T_BESTP = mv("Cibles", "Meilleure probabilité de tenir")
T_MAXFIT = mv("Cibles", "La plus grande cible qui tient")
_t0 = next(c.row for c in _wc["B"] if c.value == "Cible") + 1
TARGETS = []
for _r in range(_t0, _t0 + 60):
    _name = _wc[f"B{_r}"].value
    if not _name:
        break
    TARGETS.append({"name": _name, "country": _wc[f"C{_r}"].value, "activity": _wc[f"D{_r}"].value, "owner": _wc[f"E{_r}"].value,
                    "value": _wc[f"F{_r}"].value, "unit": _wc[f"G{_r}"].value, "basis": _wc[f"H{_r}"].value,
                    "ev_eur": _wc[f"I{_r}"].value if isinstance(_wc[f"I{_r}"].value, (int, float)) else None,
                    "mult": _wc[f"J{_r}"].value, "fits": _wc[f"K{_r}"].value,
                    "p": _wc[f"L{_r}"].value if isinstance(_wc[f"L{_r}"].value, (int, float)) else None,
                    "years": _wc[f"M{_r}"].value, "score": _wc[f"S{_r}"].value or None, "ref": _wc[f"T{_r}"].value,
                    "benchmark": _name.startswith("Clean Earth")})
CANDIDATES = [t for t in TARGETS if not t["benchmark"]]
assert len(CANDIDATES) == T_N, (len(CANDIDATES), T_N)
TARGET_UNITS = {"M EUR": "M€", "Md EUR": "Md€", "M USD": "M$", "Md USD": "Md$", "M GBP": "M£", "Md GBP": "Md£"}


def published(t):
    """La valeur publiée d'une cible, dans sa devise (« 390 M£ », « 9,4 Md AUD »), ou None."""
    v = t["value"]
    if not isinstance(v, (int, float)) or isinstance(v, bool):
        return None
    return fr(v, 0 if float(v).is_integer() or abs(v) >= 100 else 1) + NB + TARGET_UNITS.get(t["unit"] or "", t["unit"] or "")


# Les cibles dont le prix est publié dans une devise que le registre ne sait pas convertir (pas de taux BCE) : ni en
# euros, ni confrontées à la grille, ce qui n'est pas « prix non publié ».
UNCONVERTED = [t for t in CANDIDATES if t["ev_eur"] is None and published(t)]
UNCONVERTED_CUR = sorted({(t["unit"] or "").split()[-1] for t in UNCONVERTED if t["unit"]})
UNCONVERTED_NAMES = " et ".join(t["name"].split(" — ")[0] for t in UNCONVERTED)
# Le texte de la valeur telle que la source l'imprime (targets.csv) : il dit sa portée (« pour l'ensemble d'Urbaser »).
PRINTED = {t["id"]: (t.get("ev_estimate") or "").strip() for t in csv.DictReader(admin_file("targets.csv").open(encoding="utf-8"))} \
    if admin_file("targets.csv").exists() else {}


def printed(t):
    """« imprimé : … » quand la valeur publiée est une phrase et non un nombre seul."""
    text = PRINTED.get(str(t["ref"] or "").split(" — ")[0], "")
    return f" ; imprimé : « {text} »" if text and not re.fullmatch(r"[\d .,]+", text) else ""
assert T_PUB == T_VAL + len(UNCONVERTED), (T_PUB, T_VAL, len(UNCONVERTED))
# Pari 1 : la capacité en probabilités (onglet Distribution, 1 000 futurs tirés)
_wd = WB["Distribution"]
def dist(key, col="F"):
    row = next(c.row for c in _wd["A"] if c.value == key)
    return _wd[f"{col}{row}"].value
def prob(label):
    row = next(c.row for c in _wd["B"] if isinstance(c.value, str) and c.value.startswith(label))
    return _wd[f"C{row}"].value
P_BIND = prob("La marge contraignante fin 2027 est positive"); P_ACQ1 = prob("Veolia peut acheter au moins 1 Md€")
P_ACQ2 = prob("… au moins 2 Md€"); P_ACQ3 = prob("… au moins 3 Md€"); P_SP26 = prob("Le ratio FFO / dette passe sous le seuil S&P fin 2026")
P_GAP8 = prob("L'EBITDA 2027 atteint l'objectif"); N_NEG = prob("Futurs (sur l'ensemble des tirages)")
ACQ_P10, ACQ_P50, ACQ_P90 = dist("MAXACQB", "E"), dist("MAXACQB", "F"), dist("MAXACQB", "G")
BIND_P5, BIND_P50 = dist("BIND27", "D"), dist("BIND27", "F")
FW = WB["Simulation"]["C11"].value
# Pari 4 : le menu de financement et les seuils de bascule (onglet Financement)
_wf = WB["Financement"]
def fin_row(prefix, col):
    row = next(c.row for c in _wf["B"] if isinstance(c.value, str) and c.value.startswith(prefix))
    return _wf[f"{col}{row}"].value
FIN_S = fin_row("Taille de la cible", "C"); FIN_M = fin_row("Multiple VE / EBITDA payé", "C")
FIN_GAP3 = fin_row("Ce qui manque sous 3x", "C"); FIN_GAPS = fin_row("Ce qui manque sous le seuil S&P", "C")
FIN_HYB, FIN_HYB_COST = fin_row("Émission d'hybrides", "E"), fin_row("Émission d'hybrides", "F")
FIN_DISP, FIN_DISP_COST = fin_row("Cessions accélérées", "E"), fin_row("Cessions accélérées", "F")
FIN_EQ, FIN_EQ_DIL = fin_row("Augmentation de capital", "E"), fin_row("Augmentation de capital", "F")
FIN_MIN, FIN_MIN_COST = fin_row("Partenaire minoritaire", "E"), fin_row("Partenaire minoritaire", "F")
def bascule(code, col):
    row = next(c.row for c in _wf["A"] if c.value == code)
    return _wf[f"{col}{row}"].value
BK_MTUCK_C4 = bascule("mTuck", "N"); BK_G26_C4 = bascule("g26", "N"); BK_FRAGILE = fin_row("Hypothèses qui font basculer", "C")
# Le tableau des seuils tel que le classeur l'écrit (Financement §D) : une ligne par hypothèse, un verdict par conclusion
# (« bascule à x », « tient partout (seuil x) », « faux partout (il faudrait x) », « hors domaine : … », « sans effet »).
_bk_head = next(c.row for c in _wf["B"] if c.value == "Hypothèse" and _wf[f"C{c.row}"].value == "Base")
BK_CONCL = [str(_wf[f"{col}{_bk_head}"].value or "").split(" : ", 1)[-1] for col in "DEFG"]
BK_ROWS = []
for _r in range(_bk_head + 1, _wf.max_row + 1):
    if not _wf[f"A{_r}"].value:
        break
    BK_ROWS.append((_wf[f"A{_r}"].value, str(_wf[f"B{_r}"].value or ""), [str(_wf[f"{col}{_r}"].value or "") for col in "DEFG"]))
def bk_verdict(code, i):
    return next(v[i] for c, _lab, v in BK_ROWS if c == code)
def _low(lab):
    return lab[0].lower() + lab[1:] if lab[1:2].islower() else lab
BK_FLIPS = {}   # conclusion → les hypothèses qui la renversent sans sortir de leurs bornes, avec leur seuil
for _code, _lab, _v in BK_ROWS:
    for _i, _s in enumerate(_v):
        if _s.startswith("bascule à "):
            BK_FLIPS.setdefault(BK_CONCL[_i], []).append(f"{_low(_lab)} (seuil {_s[len('bascule à '):]})")
assert sum(1 for _c, _lab, _v in BK_ROWS if any(s.startswith("bascule à ") for s in _v)) == BK_FRAGILE, BK_FRAGILE
MTUCK_C4 = bk_verdict("mTuck", 3)
MTUCK_C4_TXT = (f"elle ne tient que payée moins de {fr(BK_MTUCK_C4, 1)}x" if MTUCK_C4.startswith(("faux partout", "bascule à"))
                else "elle ne tient à aucun multiple possible" if MTUCK_C4.startswith("hors domaine") and "fausse" in MTUCK_C4
                else "elle tient à tout multiple de la fourchette")


def bk_answer():
    """« Quelle hypothèse ferait tomber vos conclusions ? », lu dans Financement §D."""
    n, src = int(BK_FRAGILE or 0), "onglet Financement §D, chaque seuil vérifié en le remettant dans le moteur"
    if not n:
        text = f"Aucune, prise seule et dans ses bornes : aucun seuil ne tombe dans la fourchette d'une hypothèse ({src})."
    else:
        who = "Ce sont elles" if n > 1 else "C'est elle"
        text = (f"{n} hypothèse{'s' if n > 1 else ''}, chacune seule et dans ses bornes, {'peuvent' if n > 1 else 'peut'} renverser "
                f"une conclusion ({src}). " + " ; ".join(f"« {c} » : {', '.join(items)}" for c, items in BK_FLIPS.items())
                + f". {who} qu'il faut défendre à l'oral.")
    if bk_verdict("g26", 3).startswith("faux partout") and MTUCK_C4.startswith("faux partout"):
        text += (f" Pour une cible de {fr(FIN_S / 1000, 0)} Md€, il faudrait une croissance 2026 de {pct(BK_G26_C4, 1)} "
                 f"ou un multiple sous {fr(BK_MTUCK_C4, 1)}x.")
    return text
TAX_ALIGNED = mv("ESG", "Part du capex éligible et aligné, recalculée"); TAX_HW_SHARE = mv("ESG", "Déchets dangereux (PPC) : part du capex aligné")
TAX_HW_RATE = mv("ESG", "Déchets dangereux (PPC) : capex aligné / éligible"); TAX_NONELIG = mv("ESG", "Capex non éligible")
PEER_GAP_SUEZ = mv("Levier", "Écart de FFO / dette nette entre Veolia et Suez"); PEER_GAP_HERA = mv("Levier", "Écart entre Veolia et le mieux noté")
MA_ENERGY_H1 = mv("Sensibilité", "Énergie et recyclats : effet au S1 2026", "E"); MA_USD10 = mv("Sensibilité", "Dollar : dette en USD, effet", "E")
MA_RATE1 = mv("Sensibilité", "Taux : position nette à taux variable", "E")
CRED_FIN = mv("Pont EBITDA", "Objectifs financiers tenus"); CRED_HIT = mv("Pont EBITDA", "Objectifs tenus (sur ceux")
HYB_AVG = mv("Levier", "Coupon moyen pondéré"); HYB_REFI = mv("Levier", "Si cette tranche est remplacée"); HYB_LEV = mv("Levier", "Si elle est remboursée sans remplacement")
EFF_SHARE_Y = mv("Pont EBITDA", "Part de la croissance organique expliquée"); REST = mv("Pont EBITDA", "Reste : volumes")
ORG_M = mv("Pont EBITDA", "… en M EUR, sur la base publiée"); ORG26 = mv("Pont EBITDA", "+ croissance organique 2026")
ORG27 = mv("Pont EBITDA", "+ croissance organique 2027"); EFF_IN26 = mv("Pont EBITDA", "Part de l'organique 2026")
_wp = WB["Pont EBITDA"]
HEAD_K0 = next(_wp[f"E{c.row}"].value for c in _wp["B"] if c.value == 0)
LEV_K0 = next(_wp[f"D{c.row}"].value for c in _wp["B"] if c.value == 0)
HEAD_K50 = next(_wp[f"E{c.row}"].value for c in _wp["B"] if c.value == 0.5)

# ------------------------------------------------------------------ registre
R = lambda *a: fig(*a)
f = dict(
    rot=R("GreenUp : rotation NETTE", "2024-2027"), tlo=R("GreenUp : tuck-ins, bas", "2024-2027"),
    thi=R("GreenUp : tuck-ins, haut", "2024-2027"), dcont=R("GreenUp : cessions continues", "2024-2027"),
    netcap=R("GreenUp : investissements nets par an", "2024-2027"),
    maint=R("GreenUp : investissements de maintenance", "2024-2027"),
    grow=R("GreenUp : croissance sur contrats existants", "2024-2027"),
    inddisp=R("GreenUp : cessions industrielles", "2024-2027"), rot8=R("Rotation d'actifs sur 4 ans", "2024-2027"),
    prog=R("Programme de cessions", "2026-2028"), closed=R("Cessions réalisées", "S1 2026"),
    sign=R("Signatures de cessions attendues", "2026"), ce_nfd=R("Clean Earth : effet sur l'endettement", "S1 2026"),
    nfd25=R("Endettement financier net (groupe)", "31/12/2025"), eb25=R("EBITDA (groupe)", "FY2025"),
    lev25=R("Ratio de levier publié", "FY2025"), nfd24=R("Endettement financier net (groupe)", "31/12/2024"),
    eb24=R("EBITDA (groupe)", "FY2024"), lev24=R("Ratio de levier publié", "FY2024"),
    flow26=R("Flux contractuels non actualisés : passifs financiers bruts, 2026", "31/12/2025"),
    flow27=R("Flux contractuels non actualisés : passifs financiers bruts, 2027", "31/12/2025"),
    cpap=R("Billets de trésorerie (commercial paper)", "31/12/2025"), liq=R("Total des liquidités", "31/12/2025"),
    synd=R("Ligne de crédit syndiquée non tirée", "31/12/2025"),
    cfo25=R("Tableau de flux : capacité d'autofinancement avant variation du BFR", "FY2025"),
    int25=R("Tableau de flux : intérêts payés", "FY2025"),
    gu_eff=R("GreenUp : gains d'efficacité par an", "2024-2027"), gu_base=R("GreenUp : EBITDA de base", "2023"),
    eff25=R("Gains d'efficacité annuels", "FY2025"), eff24=R("Gains d'efficacité (et non synergies)", "FY2024"),
    org25=R("Croissance organique de l'EBITDA (groupe)", "FY2025"), syn25=R("Synergies Suez cumulées réalisées", "FY2025"),
    syn24=R("Synergies Suez cumulées", "FY2024"), sh_g=R("Strongholds : croissance organique de l'EBITDA", "FY2025"),
    es_mult=R("Espagne : multiple moyen des tuck-ins", "2024-2025"), es_n=R("Espagne : nombre de tuck-ins", "2024-2025"),
    es_ev=R("Espagne : valeur d'entreprise totale des tuck-ins", "2024-2025"),
    sw_eb=R("Déchets solides : EBITDA", "FY2025"), sh_eb=R("Strongholds : EBITDA", "FY2025"),
    rev_fr=R("Chiffre d'affaires par pays : France", "FY2025"), rev_ma=R("Chiffre d'affaires par pays : Maroc", "FY2025"),
    rev_other=R("Chiffre d'affaires par pays : autres pays", "FY2025"),
    m_es=R("Espagne : multiple moyen des tuck-ins", "2024-2025"), m_t25=R("Tuck-ins 2025 (États-Unis, Brésil, Japon) : multiple", "FY2025"),
    m_wts=R("WTS, rachat des 30 % : multiple", "2025e"), hep_m=R("HEPACO (Clean Harbors) : multiple après synergies", "2024"),
    hep_ev=R("HEPACO (Clean Harbors) : prix d'acquisition", "03/2024"), hep_eb=R("HEPACO (Clean Harbors) : EBITDA ajusté 2023", "FY2023"),
    cov_ev=R("Covanta (EQT) : valeur de la transaction", "annonce 07/2021"), cov_eb=R("Covanta (EQT) : EBITDA ajusté 2021 attendu, bas", "2021e"),
    ste_ev=R("Stericycle (WM) : valeur d'entreprise", "annonce 06/2024"), ste_syn=R("Stericycle (WM) : synergies annuelles", "régime de croisière"),
    tax_cx=R("Taxonomie 2025 : capex total", "FY2025"), tax_al=R("Taxonomie 2025 : capex éligible et aligné", "FY2025"),
    tax_pct=R("Taxonomie 2025 : part du capex éligible et aligné", "FY2025"),
    tax_hw=R("Taxonomie 2025, collecte et traitement des déchets dangereux, prévention de la pollution (PPC 2.1, 2.2) : capex éligible et aligné", "FY2025"),
    tax_water=R("Taxonomie 2025, eau et assainissement : capex éligible et aligné", "FY2025"),
    env_ev=R("EnviroServe (Clean Harbors) : prix d'acquisition", "08/2026"), env_eb=R("EnviroServe (Clean Harbors) : EBITDA ajusté annuel attendu", "2026e"),
    env_syn=R("EnviroServe (Clean Harbors) : synergies de coûts attendues", "années 1-2"),
    env_post=R("EnviroServe et ES&H (Clean Harbors) : multiple combiné après synergies", "10/2026"),
    use_ev=R("US Ecology (Republic Services) : valeur totale", "annonce 02/2022"), use_eb=R("US Ecology (Republic Services) : EBITDA ajusté 12 mois", "au 30/09/2021"),
    p_acea=R("Moody's, pairs : ACEA, FFO / dette nette", "LTM 06/2025"), p_hera=R("Moody's, pairs : Hera, FFO / dette nette", "LTM 06/2025"),
    p_suez=R("Moody's, pairs : Suez, FFO / dette nette", "LTM 06/2025"), p_suez_rt=R("Moody's, pairs : Suez, notation", "05/2026"),
    en25=R("Effet des prix des commodités (énergie, recyclats) sur l'EBITDA", "FY2025"), enh1=R("Effet des prix des commodités sur l'EBITDA", "S1 2026"),
    usd_debt=R("Dette libellée en dollars", "31/12/2025"), flt=R("Position nette à taux variable", "31/12/2025"),
    fx_oi=R("Sensibilité au change : résultat opérationnel si les devises se déprécient", "FY2025"),
    hyb_tot=R("Hybrides : encours hors coupons", "31/12/2025"), hyb26=R("Hybride septembre 2019 : nominal", "31/12/2025"),
    hyb26_cpn=R("Hybride septembre 2019 : coupon", "31/12/2025"), hyb33_cpn=R("Hybride septembre 2025 : coupon", "31/12/2025"),
    i23_eb=R("Impact 2023 : EBITDA 2023 réalisé", "FY2023"), i23_pl_t=R("Impact 2023 : plastiques transformés visés", "2023"),
    i23_pl=R("Impact 2023 : plastiques transformés réalisés", "FY2023"), g25_eb=R("Guidance 2025 : croissance organique de l'EBITDA, haut", "2025"),
    nfdh1=R("Endettement financier net (groupe), après Clean Earth", "30/06/2026"),
    nfdh125=R("Endettement financier net (groupe)", "30/06/2025"), ebh1=R("EBITDA (groupe)", "S1 2026"),
    ebh125=R("EBITDA (groupe)", "S1 2025"), hyb=R("Dettes hybrides", "30/06/2026"),
    cap=R("Engagement de levier du groupe : ≤", "2027"), guid=R("Guidance : levier égal", "31/12/2026"),
    q1=R("Ratio de levier publié, hors Clean Earth", "31/03/2026"),
    hybg=R("Pont de dette 2025 : émission de la première", "FY2025"),
    nfcf24=R("Cash-flow libre net (net free cash flow)", "FY2024"),
    nfcf25=R("Cash-flow libre net (net free cash flow)", "FY2025"),
    nfcfh125=R("Cash-flow libre net avant", "S1 2025"), nfcfh126=R("Cash-flow libre net avant", "S1 2026"),
    div=R("Dividende versé aux actionnaires, approuvé", "2026"),
    divt=R("Pont de dette S1 2026 : dividendes versés", "S1 2026"),
    wcr=R("Variation du BFR opérationnel", "S1 2026"), capex=R("Investissements nets (net capex)", "FY2025"),
    g_lo=R("Guidance 2026 : croissance organique de l'EBITDA, bas", "2026"),
    g_hi=R("Guidance 2026 : croissance organique de l'EBITDA, haut", "2026"),
    hw24=R("Déchets dangereux traités", "2024"), hw25=R("Déchets dangereux traités", "2025"),
    hwt=R("Objectif : déchets dangereux traités en 2027", "2027"), hwt0=R("Objectif GreenUp initial", "2027"),
    c430=R("Capacité ajoutée par 5 usines", "en construction"),
    pfrev=R("Déchets dangereux Veolia + Clean Earth : chiffre d'affaires", "2025E"),
    pfeb=R("Déchets dangereux Veolia + Clean Earth : EBITDA", "2025E"),
    pfm=R("Déchets dangereux Veolia + Clean Earth : marge", "2025E"), cagr=R("Ambition GreenUp relevée", "2024-2027"),
    eurev=R("Chiffre d'affaires Déchets dangereux Europe", "S1 2026"), eug=R("Déchets dangereux Europe : croissance", "S1 2026"),
    am25=R("EBITDA Amériques", "S1 2025"), am26=R("EBITDA Amériques", "S1 2026"),
    cer26=R("Clean Earth : chiffre d'affaires", "2026E"), ceeb=R("Clean Earth : EBITDA", "2026E"),
    cl24=R("Provisions de fermeture et post-fermeture (total)", "31/12/2024"),
    cl25=R("Provisions de fermeture et post-fermeture (total)", "31/12/2025"),
    cl26=R("Provisions de fermeture et post-fermeture (réhabilitation", "30/06/2026"),
    am_cl25=R("Provisions de fermeture, dont Amériques", "31/12/2025"),
    am_cl26=R("Provisions de fermeture, dont Amériques", "30/06/2026"),
    unw=R("Désactualisation des provisions", "FY2025"), gw=R("Clean Earth : goodwill préliminaire", "30/06/2026"),
    price=R("Clean Earth : prix d'acquisition", "01/06/2026", "M EUR"),
    gar25=R("Engagements reçus liés au périmètre", "31/12/2025"), gar26=R("Engagements reçus liés au périmètre", "30/06/2026"),
    site=R("Provisions pour réhabilitation de sites", "31/12/2025"), envr=R("Provisions pour risques environnementaux", "31/12/2025"),
    ev=R("Clean Earth : valeur d'entreprise", "annonce 11/2025", "Md USD"), mult=R("Clean Earth : multiple", "2026e"),
    syn=R("Clean Earth : synergies de coûts", "année 4"),
    sites=R("Clean Earth : sites spécialisés", "annonce 11/2025"), permits=R("Clean Earth : permis", "annonce 11/2025"),
    boost4=R("GreenUp : investissements nets de croissance 2024-2027", "2024-2027"),
    strong2=R("GreenUp : à investir dans les strongholds", "2024-2027"), rot4=R("Rotation d'actifs déjà réalisée", "2024-2025e"),
    rot85=R("Rotation d'actifs totale GreenUp", "2024-2028"), nfi24=R("Pont de dette 2024 : investissements financiers nets des cessions", "FY2024"),
    nfi25=R("Pont de dette 2025 : investissements financiers", "FY2025"), nfih1=R("Pont de dette S1 2026 : investissements financiers", "S1 2026"),
    pctB=R("Part des acquisitions GreenUp dans les boosters", "2024-2025e"), pctX=R("Part des acquisitions GreenUp hors d'Europe", "2024-2025e"),
    mat25=R("Maturité moyenne de la dette financière nette", "31/12/2025"), fixed=R("Part de la dette financière nette à taux fixe", "31/12/2025"),
    nfdh125b=R("Endettement financier net (groupe)", "30/06/2025"),
    mo25=R("Moody's : FFO / dette nette ajustée", "FY2025"), mo24=R("Moody's : FFO / dette nette ajustée", "FY2024"),
    mo26lo=R("Moody's : FFO / dette nette ajustée attendu, bas", "2026F"), mo26hi=R("Moody's : FFO / dette nette ajustée attendu, haut", "2026F"),
    mond25=R("Moody's : dette nette ajustée", "FY2025", "Md EUR"), mond26=R("Moody's : dette nette ajustée, pic", "2026F"),
    motrig=R("Moody's : seuil de dégradation", "2026-2027"), moup=R("Moody's : condition de relèvement", "2026-2027"),
    sptrig=R("S&P : seuil de dégradation", "2026-2028"), splo=R("S&P : FFO / dette ajustés attendu, bas", "2026-2028"),
    sphi=R("S&P : FFO / dette ajustés attendu, haut", "2026-2028"), molev=R("Moody's : levier net selon la définition de Veolia", "FY2025"),
    wt_g=R("Croissance organique de l'EBITDA Water Technologies", "FY2025"), am_g=R("Croissance organique de l'EBITDA Amériques", "FY2025"),
    eu_g=R("Croissance organique de l'EBITDA Europe", "FY2025"), wt_eb=R("EBITDA Water Technologies", "FY2025"),
    boost_rev=R("Boosters : chiffre d'affaires", "FY2025"), boost_eb=R("Boosters : EBITDA", "FY2025"),
    boost_g=R("Boosters : croissance organique de l'EBITDA", "FY2025"), bfee_rev=R("Bioénergie, flexibilité, efficacité énergétique : chiffre d'affaires", "FY2025"),
    clh_rev=R("Clean Harbors : chiffre d'affaires direct total", "FY2025"), clh_eb=R("Clean Harbors : EBITDA ajusté total", "FY2025"),
    clh_m=R("Clean Harbors : marge d'EBITDA ajusté", "FY2025"), clh_env=R("Clean Harbors : passifs environnementaux totaux", "31/12/2025"),
    en_rev=R("Enviri : chiffre d'affaires du segment Clean Earth", "FY2025"), en_oi=R("Enviri : résultat opérationnel du segment Clean Earth", "FY2025"),
    en_dep=R("Enviri : dépréciation du segment Clean Earth", "FY2025"), en_am=R("Enviri : amortissement du segment Clean Earth", "FY2025"),
    en_cur=R("Enviri : passifs environnementaux, part courante", "31/12/2025"), en_lt=R("Enviri : passifs environnementaux, part non courante", "31/12/2025"),
    eff398=R("Gains d'efficacité (et non synergies)", "FY2024"), syn435=R("Synergies Suez cumulées", "FY2024"),
    syn530=R("Synergies Suez : objectif cumulé relevé", "fin 2025"), syn534=R("Synergies Suez cumulées réalisées", "FY2025"),
    co25=R("KPI GreenUp : émissions de GES évitées, progression vs 2023", "2025"), co18=R("KPI GreenUp : émissions de GES évitées, objectif", "2027", "Mt CO2e"),
    w25=R("KPI GreenUp : eau douce économisée", "2025"),
    mo_rep=R("Moody's : dette brute publiée", "FY2025"), mo_pens=R("Moody's : ajustement pensions", "FY2025"),
    mo_hyb=R("Moody's : ajustement titres hybrides", "FY2025"), mo_sec=R("Moody's : ajustement titrisation", "FY2025"),
    mo_ns=R("Moody's : ajustements non standard", "FY2025"), mo_adj=R("Moody's : dette brute ajustée", "FY2025", "M EUR"),
    mo_cash=R("Moody's : trésorerie retenue", "FY2025"), mo_net=R("Moody's : dette nette ajustée (Exhibit 13)", "FY2025"),
    mo_ffo25=R("Moody's : FFO (funds from operations)", "FY2025"), mo_ffo24=R("Moody's : FFO (funds from operations)", "FY2024"),
    mo_div=R("Moody's : dividendes", "FY2025"), mo_rcf=R("Moody's : RCF (retained", "FY2025"), mo_ndeb=R("Moody's : dette nette ajustée / EBITDA ajusté", "FY2025"),
    mo_f26=R("Moody's : FFO / dette nette ajustée prévu", "2026F"), mo_f27=R("Moody's : FFO / dette nette ajustée prévu", "2027F"),
    urd_gross=R("Sous-total des emprunts", "31/12/2025"),
)
# Le T3 2026 (publié le 05/11/2026) et les chiffres versés pour corriger l'audit du 09/11.
f.update(dict(
    nfdq3=R("Endettement financier net (groupe), après Clean Earth", "30/09/2026"),
    nfdq325=R("Endettement financier net (groupe)", "30/09/2025"), levq325=R("Ratio de levier publié", "30/09/2025"),
    ebq325=R("EBITDA (groupe)", "9M 2025"), ebq326=R("EBITDA (groupe)", "9M 2026"),
    g9m=R("Croissance organique de l'EBITDA (groupe)", "9M 2026"),
    nfcf9m25=R("Cash-flow libre net avant", "9M 2025"), nfcf9m26=R("Cash-flow libre net avant", "9M 2026"),
    divt9m=R("Pont de dette 9M 2026 : dividendes versés", "9M 2026"),
    nfi9m=R("Pont de dette 9M 2026 : investissements financiers", "9M 2026"),
    ce_nfd9m=R("Clean Earth : effet sur l'endettement", "9M 2026"),
    hyb_rep=R("Pont de dette 9M 2026 : remboursement de dette hybride", "9M 2026"),
    sign570=R("Programme de cessions : montant déjà signé", "au 05/11/2026"),
    mdisp=R("Programme de cessions : multiple moyen", "au 05/11/2026"),
    closed9m=R("Cessions réalisées", "9M 2026"), rook=R("Cession de 20 % de Rookery South", "25/09/2026"),
    lev_guid=R("Guidance 2026 confirmée au T3 : ratio de levier", "31/12/2026"),
    lev_traj=R("Trajectoire de levier présentée au T3", "2027"),
    hw_eb24=R("Déchets dangereux (booster) : EBITDA", "FY2024"), hw_eb25=R("Déchets dangereux (booster) : EBITDA", "FY2025"),
    hw_g25=R("Déchets dangereux (booster) : croissance de l'EBITDA", "FY2025"),
    growth_net=R("GreenUp : croissance, montant net par an", "2024-2027"),
))
HYB_SHARE = mv("Levier", "Part des hybrides dans l'écart")
SEG_EB27 = mv("Segments", "EBITDA 2027 « au rythme", "E"); CE_EBREC = mv("Booster", "EBITDA 2025 recalculé"); CE_EBGAP = mv("Booster", "Écart entre l'EBITDA de Veolia")
M_REC = mv("Booster", "Multiple sur l'EBITDA 2025 reconstitué"); CE_MREC = mv("Booster", "Marge d'EBITDA 2025 recalculée")
CLH_ENVPCT = mv("ESG", "Clean Harbors : passifs environnementaux en %"); EN_ENV = mv("ESG", "Enviri (groupe entier)"); CE_ANALOG = mv("ESG", "Ordre de grandeur par analogie")

# ------------------------------------------------------------------ les dossiers, dérivés des textes du rapport
# Un dossier par rôle, tous sur le même plan : la thèse (le titre de la partie), la réponse (sa section d'ouverture), le
# raisonnement (ses intertitres), ce que nous n'avons pas pu établir, les questions probables du jury (la réponse est la
# section du rapport), puis chaque chiffre que la partie cite, avec sa source. Aucune phrase n'est écrite ici : la
# dérivation, son contrôle (partie absente, jeton inconnu : marqués, ou arrêt en mode strict) et le gabarit sont ceux de
# build_fiches.py.
import build_fiches as F

CORPUS = F.Corpus(LIV, ADMIN)
PARTIES = F.deriver(CORPUS)
# Où chaque rôle travaille dans le classeur : des noms d'onglets, vérifiés contre le classeur recalculé.
ONGLETS = {1: "Entrées, Cessions §A et §B", 2: "Levier, Pont de dette, Échéancier",
           3: "Trajectoire, Sensibilité, Cessions §C, Pont de dette, Pont EBITDA", 4: "Segments, Pont EBITDA, Booster",
           5: "ESG", 6: "Cibles"}
for _n, _tabs in ONGLETS.items():
    for _t in _tabs.split(", "):
        if _t.split(" §")[0] not in WB.sheetnames:
            raise SystemExit(f"build_dossiers : l'onglet « {_t.split(' §')[0]} » cité pour le rôle {_n} n'existe pas dans le classeur.")


def _ouverture_de(cle):
    """Les paragraphes de la première section d'un texte (la conclusion : la réponse du rapport en quatre phrases)."""
    parsed, out, dedans = CORPUS.textes.get(cle), [], False
    for b in (parsed.blocks if parsed else []):
        if b.kind == "heading" and b.level == 2:
            if dedans:
                break
            dedans = True
        elif dedans and b.kind == "paragraph":
            out.append(b)
    return out


FIL = _ouverture_de("conclusion") or [b for b in PARTIES[-1]["ouverture"]["blocs"] if b.kind == "paragraph"]
FIL_JETONS = F._jetons(FIL)
for _t in FIL_JETONS:
    if not CORPUS.connu(_t):
        CORPUS.anomalies.append(("08_Livrables/texte/conclusion.md", f"jeton inconnu des registres et du classeur : [{_t}]"))
if F.STRICT and CORPUS.anomalies:
    F.echec(F.resume(CORPUS) + ". Mode strict : corriger le texte ou verser la ligne au registre ; rien n'est écrit.")


def _figures(p):
    """Les lignes « chiffres » d'un rôle pour report-data.json : (id, libellé, période, valeur, source, relu par)."""
    rows = []
    for t in p["jetons"]:
        r = CORPUS.res[t]
        if not CORPUS.connu(t):
            rows.append((t, CORPUS.libelle(t), "", CORPUS.valeur(t), CORPUS.reference(t), ""))
        elif r.kind == "model":
            rows.append(("calcul", r.label, "", CORPUS.valeur(t), CORPUS.reference(t), ""))
        else:
            rows.append((t, F.R.units_in_label(r.label), CORPUS.periodes.get(t, ""), CORPUS.valeur(t),
                         f"{Path(r.file).name}, p. {r.page}" if r.file else r.url, r.checked_by))
    return rows


def dossier_html(p):
    if p["manque"]:
        return F.absente_html(CORPUS, p, "dossier")

    def etape(s):
        sous = "".join(f'<span class="sous">{CORPUS.html(t)}</span>' for t in s["sous"])
        return f"<li>{CORPUS.html(s['titre'])}{sous}</li>"
    onglets = f'<p class="onglets">Dans le classeur : {esc(ONGLETS[p["role"]])}</p>' if p["role"] in ONGLETS else ""
    return f"""
<article class="fiche dossier" id="dossier-{p['rang']}" aria-labelledby="dossier-{p['rang']}-t">
  {F.tete_html(CORPUS, p, "dossier")}
  <div class="f-grille">
    <section class="bloc b-rep"><h3>La réponse</h3>{''.join(CORPUS.bloc(b) for b in p['ouverture']['blocs'])}</section>
    <section class="bloc b-msg"><h3>Le raisonnement</h3><ol class="messages">{''.join(etape(s) for s in p['sections'])}</ol>{onglets}</section>
    <section class="bloc b-lim"><h3>Ce que nous n'avons pas pu établir</h3>{F.limites_html(CORPUS, p)}</section>
    <section class="bloc b-qa"><h3>Questions probables du jury</h3>{F.questions_html(CORPUS, p['sections'][1:], entier=True)}</section>
    <section class="bloc b-tous"><h3>Les chiffres de la partie</h3>{F.chiffres_html(CORPUS, p['jetons'], ident=True)}</section>
  </div>
</article>"""


for _p in PARTIES:
    if not _p["role"]:
        continue   # report-data.json porte les six rôles ; la recommandation a son texte, pas de rôle
    REPORT_ROLES.append({
        "n": _p["role"], "name": _p["nom"], "brief": CORPUS.texte(_p["these"]),
        "answer": " ".join(CORPUS.texte(b.inlines) for b in _p["ouverture"]["blocs"] if b.kind == "paragraph"),
        "reasoning": [CORPUS.texte(s["titre"]) + (" — " + CORPUS.texte(s["premier"].inlines) if s["premier"] else "") for s in _p["sections"]],
        "open": [CORPUS.texte(item) for item in _p["limites"]],
        "qa": [(F.question(CORPUS, s), CORPUS.texte(s["premier"].inlines)) for s in _p["sections"][1:] if s["premier"]],
        "tabs": ONGLETS[_p["role"]], "figures": _figures(_p),
    })

# Les indicateurs du rapport (tableau de la recommandation, build_report.js) : forme et contenu inchangés.
tiles = [
    ("Levier fin 2026", x(LEV26), "guidance : égal ou légèrement au-dessus de 3x"),
    ("Levier fin 2027", x(LEV27), "engagement : au plus 3x"),
    ("Marge sous 3x fin 2027", fr(HEAD / 1000, 2, "Md€"), f"défavorable {fr(UNF / 1000, 2)} · favorable {fr(FAV / 1000, 2)}"),
    ("Marge sous 18 % S&P fin 2027", fr(HEADSP27 / 1000, 2, "Md€"), f"FFO / dette ajustée {pct(RATIO26)} en 2026, {pct(RATIO27)} en 2027"),
    ("Acquisition max. fin 2027", fr(MAXACQB / 1000, 2, "Md€"), f"sous la contrainte qui mord : {WHICH27}"),
    ("Sorties nettes depuis 2024", fr(SPENT["cum"], 2, "Md€"), "investissements financiers nets des cessions, au 30/09/2026"),
]

# Ce qu'il reste à faire : des comptes, pris aux registres et aux en-têtes des textes ; rien qui vieillisse.
_cites = [t for p in PARTIES for t in p["jetons"] if CORPUS.res[t].kind != "model"]
_cites = list(dict.fromkeys(_cites))
_non_relus = [t for t in _cites if not CORPUS.res[t].checked]
_sans_signature = [f"{p['etiquette'].lower()} ({p['nom']})" for p in PARTIES if not p["signe"]]
_sans_relecteur = [f"{p['etiquette'].lower()} ({p['nom']})" for p in PARTIES if not p["relu"]]
_todo = [f"Relire en croisé les {len(REG)} chiffres du registre ({sum(1 for _r in REG if _r.get('checked_by'))} relus) "
         f"et les {len(DEALS)} opérations ({sum(1 for _d in DEALS if _d.get('checked_by'))} relues)."]
if _non_relus:
    _todo.append(f"En priorité, les {len(_non_relus)} lignes que le rapport cite et qu'aucun second membre n'a relues (sur {len(_cites)} citées).")
if _sans_signature:
    _todo.append("Signer (frontmatter signe_par du texte) : " + ", ".join(_sans_signature) + ".")
if _sans_relecteur:
    _todo.append("Faire relire (frontmatter relu_par du texte) : " + ", ".join(_sans_relecteur) + ".")

CSS_DOSSIERS = """
.fil-texte { max-width: var(--f-mesure); font-size: 1.08rem; line-height: 1.5; margin: 0 0 var(--s5, 16px); }
.fil-texte p { margin: 0 0 var(--s2, 8px); }
.tuiles { display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 0 var(--s7, 26px); border-top: 1px solid var(--ink, #151512); }
.tuile { padding: var(--s3, 10px) 0 var(--s4, 13px); border-bottom: 1px solid var(--hair, #15151224); break-inside: avoid; }
.tuile p { margin: 0; }
.tuile .t-val { font: 700 1.9rem/1.1 var(--sans, system-ui, sans-serif); letter-spacing: -.02em; margin: var(--s1, 6px) 0; font-variant-numeric: tabular-nums; }
.tuile .t-src { font: 400 .74rem/1.5 var(--mono, ui-monospace, monospace); color: var(--mut, #55524c); }
.dossier .b-rep { font-size: 1.02rem; }
.dossier .b-tous { margin-top: var(--s3, 10px); }
@media print { .dossier .b-tous { break-before: auto; } .tuile .t-val { font-size: 16pt; } }
"""
_tuiles = "".join(
    f'<div class="tuile"><p class="anno">{esc(CORPUS.libelle(t))}</p><p class="t-val">{esc(CORPUS.valeur(t))}</p>'
    f'<p class="t-src">' + (f'<a href="{esc(CORPUS.lien(t))}">{esc(CORPUS.reference(t))}</a>' if CORPUS.lien(t) else esc(CORPUS.reference(t)))
    + "</p></div>" for t in FIL_JETONS)
_suivi = "".join(f"<li>{esc(x)}</li>" for x in _todo)
TITRE = "Dossiers des rôles — sujet 2"
corps = f"""<main class="wrap">
<header class="f-couv">
  <p class="anno">Capstone EDHEC · Sujet 2 · Veolia, GreenUp 2027</p>
  <h1>Dossiers des rôles</h1>
  <p class="lede">Un dossier par rôle, tous sur le même plan : la thèse, la réponse, le raisonnement, ce que nous n'avons pas pu établir,
  les questions probables du jury, puis chaque chiffre que la partie cite, avec sa source. Tout vient du texte du rapport et des registres.</p>
  <p class="prov">{esc(F.provenance(CORPUS))} Contrôles bloquants du classeur : {esc(checks_ok)}.</p>{F.alertes_html(CORPUS)}
  <h2>Le fil rouge</h2>
  <div class="fil-texte">{''.join(CORPUS.bloc(b) for b in FIL)}</div>
  <div class="tuiles">{_tuiles}</div>
  <h2>D'une partie à l'autre</h2>
  {F.fil_html(CORPUS, PARTIES, "dossier")}
</header>
{F.nav_html(PARTIES, "dossier")}
{''.join(dossier_html(p) for p in PARTIES)}
<section class="suivi" aria-labelledby="h-todo">
  <h2 id="h-todo">Ce qu'il reste à faire avant l'oral</h2>
  <ul>{_suivi}</ul>
</section>
</main>"""
# Deux sorties, comme avant : le fragment habillé (tête et barre is42, puis la page) et la version autonome, servie
# telle quelle depuis le vault (doctype, charset).
from dataroom import is42 as _is42
page = (_is42.head() + _is42.bar("livrables") + '<meta name="robots" content="noindex, nofollow">\n'
        f"<title>{esc(TITRE)}</title>\n<style>{F.CSS}{CSS_DOSSIERS}{F.css_alerte(CORPUS)}</style>\n{corps}\n{F.JS}\n")
OUT.write_text(page, encoding="utf-8")
standalone = F.page_html(TITRE, corps, CSS_DOSSIERS + F.css_alerte(CORPUS))

OUT_RANGE = OUT.with_name("fourchette-capacite.md")
try:
    _hist = __import__("json").loads(OUT.with_name("fourchette-historique.json").read_text(encoding="utf-8")).get("versions", [])
except (OSError, ValueError):
    _hist = []
_version_line = ""
if _hist:
    _v = _hist[-1]
    _version_line = (f"\nVersion {_v.get('version')} du {_v.get('date')}" + (f", {_v['note']}" if _v.get("note") else "") + "."
                     + (f" La version {_hist[-2].get('version')} du {_hist[-2].get('date')} reste dans fourchette-historique.json." if len(_hist) > 1 else ""))
rank_txt = "\n".join(f"{i + 1}. {n} — amplitude {fr(v, 0)} M€" for i, (n, v) in enumerate(rank[:5]))
OUT_RANGE.write_text(f"""# Fourchette de capacité d'acquisition — GreenUp 2027

Périmètre « capacité financière », pour les groupes des trois autres périmètres. Généré depuis le classeur
`modele-greenup-2027.xlsx` (contrôles bloquants : {checks_ok}) le {__import__('datetime').date.today().isoformat()}.
Chaque chiffre d'entrée vient de `00_Admin/register.csv` ; les hypothèses sont dans l'onglet Hypothèses.
{_version_line}

## Le nombre à respecter

| Fin 2027 | Défavorable | Central | Favorable |
|---|---|---|---|
| Marge de dette sous le levier ≤ 3x (M EUR) | {fr(UNF, 0)} | {fr(HEAD, 0)} | {fr(FAV, 0)} |
| Marge de dette sous le seuil S&P 18 % FFO / dette ajustée (M EUR) | {fr(HEADSP27_UNF, 0)} | {fr(HEADSP27, 0)} | {fr(traj("Marge de dette fin 2027 sous le seuil", FAVC), 0)} |
| Contrainte qui mord en premier | {constraint(WHICH27_UNF)} | {constraint(WHICH27)} | {constraint(WHICH27_FAV)} |
| Acquisition maximale au multiple des tuck-ins du scénario (central : celui de Clean Earth, {P(f['mult'])}x), M EUR | {fr(traj("Acquisition maximale fin 2027 sous la contrainte", UNFC), 0)} | {fr(MAXACQB, 0)} | {fr(traj("Acquisition maximale fin 2027 sous la contrainte", FAVC), 0)} |

Lecture : une acquisition ajoute son prix à la dette et seulement prix ÷ multiple à l'EBITDA ; elle tient tant que
prix ≤ marge ÷ (1 − 3 ÷ multiple). Au multiple de Clean Earth, 1 Md€ de marge vaut environ {fr(1000 / (1 - 3 / N(f['mult'])) / 1000, 2)} Md€ d'acquisition.

## Les repères de départ

- Levier fin 2026 (définition Veolia) : {x(LEV26)} dans le scénario central — guidance « égal ou légèrement supérieur à 3x ».
- Levier fin 2027 : {x(LEV27)} (engagement ≤ 3x).
- FFO / dette ajustée (mesure des agences) : {pct(RATIO26)} fin 2026 (Moody's attend {P(f['mo26lo'])}-{P(f['mo26hi'])} %), {pct(RATIO27)} fin 2027 ; S&P abaisse la note sous {P(f['sptrig'])} %.
- Sorties nettes d'investissements financiers depuis 2024 : {fr(SPENT['cum'], 2)} Md€ au 30/09/2026 (à ne pas comparer aux {P(f['boost4'])} Md€ de GreenUp, organique et strongholds compris).

## Ce qui fait bouger la fourchette (une hypothèse à la fois, marge sous 3x)

{rank_txt}

## Ce que les autres périmètres doivent faire de ce nombre

Toute recommandation d'acquisition se compare à la marge du scénario central et doit survivre au scénario
défavorable, ou dire explicitement quelle cession ou quelle levée de fonds propres la finance. Les hypothèses sont
discutables : elles sont en jaune dans le classeur, et la fourchette se régénère en une commande.
""", encoding="utf-8")
import json as _json
# Le fil rouge de report-data.json : l'ouverture de la conclusion, en texte nu. Le « reste à faire » : les comptes ci-dessus.
_thread = " ".join(CORPUS.texte(b.inlines) for b in FIL)
# Les questions du cours et les blocs du chantier ne s'affichent plus dans les dossiers ; le rapport les garde en repli,
# pour un brouillon sans introduction. Leurs nombres viennent du registre ; la citation du sujet reste entre guillemets.
_COURS = f"""      <div><h3>Les quatre questions de recherche</h3><ol class="steps">
        <li><strong>Écart stratégique</strong> entre la trajectoire et 2027 (EBITDA, croissance, KPI ESG) ; segments et géographies qui sur- ou sous-performent → rôles 2 et 4.</li>
        <li><strong>Benchmark et opportunités</strong> : qui contribue le plus à l'ESG dans chaque booster ; multiples et synergies des deals récents → rôles 4 et 6.</li>
        <li><strong>Allocation du capital restant</strong> : capex organique ou croissance externe, boosters ou strongholds, critère d'impact ESG par euro → rôles 1, 3 et 6. « Tout le reste s'y alimente. »</li>
        <li><strong>Orientations</strong> : géographies, technologies, synergies (« 398 M€ en 2024 → objectif 530 M€ ») → rôle 6.
        Piège repéré : les {P(f['eff398'])} M€ de 2024 sont des <em>gains d'efficacité</em> ; les synergies Suez cumulées étaient de
        {P(f['syn435'])} M€ fin 2024, objectif relevé à {P(f['syn530'])} M€, réalisé {P(f['syn534'])} M€ fin 2025. À dire, pas à corriger en silence.</li>
      </ol></div>
      <div><h3>Les cinq blocs du chantier « capacité »</h3><ol class="steps">
        <li>Reconstituer ce qui a été dépensé, et dire si les {P(f['boost4'])} Md€ étaient bruts ou nets → rôle 1, Cessions §D.</li>
        <li>La position de départ : dette, EBITDA, profil d'échéances, définition du levier propre à Veolia → rôle 2, Levier.</li>
        <li>Traduire l'engagement en nombre, et lire les seuils de dégradation des agences, « la vraie limite » → rôle 2, Levier §B'.</li>
        <li>Les trois sources de capacité : cash-flow libre, cessions, marge de levier → rôle 3, Trajectoire.</li>
        <li>La sensibilité : prouver quel paramètre bouge le plus la réponse → rôle 3, Sensibilité.</li>
      </ol>"""
_qr = [plain(x) for x in re.findall(r"<li>(.*?)</li>", _COURS.split("Les quatre questions de recherche", 1)[1].split("</ol>", 1)[0], re.S)]
_blocks = [plain(x) for x in re.findall(r"<li>(.*?)</li>", _COURS.split("Les cinq blocs du chantier", 1)[1].split("</ol>", 1)[0], re.S)]
_range_rows = []
for line in OUT_RANGE.read_text(encoding="utf-8").splitlines():
    if line.startswith("| ") and not line.startswith("|---") and "Défavorable" not in line:
        _range_rows.append([c.strip() for c in line.strip("|").split("|")])
# ------------------------------------------------------------------ les questions pour Veolia, classées
# Ce que chaque réponse de Veolia déplacerait : l'amplitude de la marge sous 3x fin 2027 entre les bornes de
# l'hypothèse (Sensibilité, classement), et les questions de questions-veolia.md qui la visent. Le lien
# hypothèse → questions est une lecture du groupe, écrite ici ; les nombres viennent du classeur.
QUESTIONS_FOR = {
    "Part du reste du programme encaissée en 2027": [4],
    "Croissance organique de l'EBITDA en 2027": [18, 19, 11],
    "Tuck-ins payés en 2027": [14, 15],
    "Cash-flow libre net du T4 2026": [],
    "Conversion cash-flow libre net / EBITDA en 2027": [20],
    "Croissance organique de l'EBITDA en 2026": [18, 11],
    "Part des synergies réalisée en 2027": [10],
    "Multiple VE / EBITDA des actifs cédés": [5],
    "EBITDA de Clean Earth en année pleine": [10],
    "Croissance du dividende versé en 2027": [6],
    "Multiple VE / EBITDA des tuck-ins": [14, 15],
    "Part des cessions signées non encaissées qui l'est avant le 31/12/2026": [4],
    "FFO (mesure des agences) en % de l'EBITDA": [2, 21],
}
_fcf_h1_26 = R("Cash-flow libre net avant investissements financiers et dividendes", "S1 2026")
_fcf_h1_25 = R("Cash-flow libre net avant investissements financiers et dividendes", "S1 2025")
_fcf_fy25 = R("Cash-flow libre net (net free cash flow)", "FY2025")
TO_ADD = {
    "Cash-flow libre net du T4 2026": f"Quel cash-flow libre net attendez-vous au quatrième trimestre 2026, après {P(f['nfcf9m26'])} M€ "
        f"sur neuf mois ({f['nfcf9m26']['id']}) ? En 2025, le quatrième trimestre avait apporté {fr(N(_fcf_fy25) - N(f['nfcf9m25']), 0)} M€ "
        f"({_fcf_fy25['id']} − {f['nfcf9m25']['id']}) : est-ce la bonne référence, avec Clean Earth consolidé et le BFR du Moyen-Orient ?",
}
unknown = [lab for lab, _ in rank if lab not in QUESTIONS_FOR]
assert not unknown, f"hypothèse sans correspondance de questions : {unknown}"
_q_rows, _q_order, _seen = [], [], set()
for i, (lab, amp) in enumerate(rank, 1):
    lo, hi = RANK_LH[lab]
    qs = QUESTIONS_FOR[lab]
    _q_rows.append(f"| {i} | {lab} | {fr(amp, 0)} | {fr(lo, 0)} / {fr(hi, 0)} | "
                   + (", ".join(f"Q{n}" for n in qs) if qs else "**aucune — à poser**") + " |")
    for n in qs:
        if n not in _seen:
            _seen.add(n); _q_order.append(n)
_q_add = [f"- **{lab}** ({fr(dict(rank)[lab], 0)} M€ d'amplitude) : {TO_ADD.get(lab, 'question à rédiger.')}"
          for lab, _ in rank if not QUESTIONS_FOR[lab] and dict(rank)[lab] > 0]
OUT.with_name("questions-classees.md").write_text(f"""# Questions pour Veolia, classées par ce que la réponse déplacerait

Généré le {__import__('datetime').date.today().strftime('%d/%m/%Y')} depuis le classeur (onglet Sensibilité, classement). **Amplitude** : écart de la marge
sous 3x fin 2027 entre la borne basse et la borne haute de l'hypothèse, toutes les autres à leur base (M€).
Les numéros renvoient à `questions-veolia.md`. Une réponse de Veolia vaut ce qu'elle resserre : commencer par le haut.

| Rang | Hypothèse | Amplitude | Δ bas / Δ haut | Questions |
|---:|---|---:|---:|---|
""" + "\n".join(_q_rows) + f"""

## Ordre proposé pour les questions à Veolia

{" → ".join(f"Q{n}" for n in _q_order)} ; puis les questions de méthode (Q16, Q17) si le temps le permet.
Les questions de définition et de périmètre ({", ".join(f"Q{n}" for n in range(1, 22) if n not in _seen and n not in (16, 17))}) ne déplacent pas
une hypothèse du modèle : elles cadrent la lecture des chiffres.

## Hypothèses qu'aucune question ne vise encore

""" + ("\n".join(_q_add) if _q_add else "Aucune : chaque hypothèse qui déplace la marge a sa question.") + """

Le FFO en % de l'EBITDA ne déplace pas la marge sous 3x (amplitude nulle) mais toute la contrainte des agences
(onglet Levier §B5) : Q2 et Q21 restent prioritaires pour le seuil S&P.
""", encoding="utf-8")

OUT.with_name("report-data.json").write_text(_json.dumps({
    "generated": __import__("datetime").date.today().isoformat(),
    "checks": str(checks_ok), "n_figures": len(REG),
    "title": "GreenUp 2027 : combien Veolia peut-elle encore acheter sans dépasser ses engagements de levier ?",
    "subtitle": "Capstone Innovations in Finance — sujet 2, périmètre « capacité financière »",
    "thread": _thread,
    "tiles": [(a, plain(b), c) for a, b, c in tiles],
    "research_questions": _qr, "capacity_blocks": _blocks,
    "range_rows": _range_rows,
    "roles": REPORT_ROLES,
    "todo": _todo,
    "register": [[r["id"], r["label"], display_value(r["value"], r["unit"]), r["period"],
                  f"{Path(r['file']).name}, p. {r['page']}, {r['source_url']} (consulté le {r['date_consulted']})",
                  r.get("workstream", ""), r.get("checked_by", "")] for r in REG],
}, ensure_ascii=False, indent=1), encoding="utf-8")
OUT_STANDALONE = OUT.with_name("dossiers-sujet-2.standalone.html")
OUT_STANDALONE.write_text(standalone, encoding="utf-8")
print("écrit", OUT, len(page) // 1024, "Ko ; version autonome", OUT_STANDALONE.name, ";", len(PARTIES), "dossiers dérivés des textes")
F.signaler(CORPUS, "build_dossiers")
