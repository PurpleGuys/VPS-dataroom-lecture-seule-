"""Génère les dossiers des six rôles (HTML) depuis le registre et le classeur recalculé.

Aucun chiffre n'est tapé ici : chaque valeur vient de register.csv (avec son ID, son
fichier et sa page) ou d'une cellule du modèle recalculé.
"""
import csv
import html
import re
from pathlib import Path

from openpyxl import load_workbook

HERE = Path(__file__).parent
REG = list(csv.DictReader((HERE / "register.csv").open(encoding="utf-8")))
WB = load_workbook("/home/ubuntu/capstone-livrables/modele-greenup-2027.xlsx", data_only=True)
OUT = HERE / "dossiers-sujet-2.html"
NB = " "


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


def N(r):
    return parse(r["value"])[0]


def P(r, d=None):
    v, dd = parse(r["value"])
    return fr(v, dd if d is None else d)


def src(r):
    return (f'<span class="id">{r["id"]}</span> {esc(Path(r["file"]).name)} '
            f'<span class="pg">p.{r["page"]}</span>')


class Row(str):
    """A table row: the HTML, plus the cells the Word report needs."""
    cells: tuple = ()


def row_fig(label, r, shown=None):  # noqa: F811
    value = shown or globals()["shown"](r)
    row = Row(f'<tr><td>{esc(label)}<span class="per">{esc(r["period"])}</span></td>'
              f'<td class="num" title="imprimé : {esc(r["value"])}">{value}</td><td class="src">{src(r)}</td></tr>')
    row.cells = (r["id"], label, r["period"], html.unescape(value), f"{Path(r['file']).name}, p. {r['page']}", r.get("checked_by", ""))
    return row


def row_calc(label, value, how):
    row = Row(f'<tr class="calc"><td>{esc(label)}</td><td class="num">{value}</td>'
              f'<td class="src">{esc(how)}</td></tr>')
    row.cells = ("calcul", label, "", value, how, "")
    return row


LAST_ROWS: list = []


def table(rows):
    LAST_ROWS[:] = [r.cells for r in rows if getattr(r, "cells", None)]
    return ('<div class="tbl"><table><thead><tr><th>Chiffre</th><th class="num">Valeur</th>'
            '<th>Source</th></tr></thead><tbody>' + "".join(rows) + "</tbody></table></div>")


def plain(s):
    """HTML → text, for the Word report."""
    return re.sub(r"\s+", " ", html.unescape(re.sub(r"<[^>]+>", "", str(s)))).strip()


REPORT_ROLES: list = []


def steps(items):
    return "<ol class=\"steps\">" + "".join(f"<li>{i}</li>" for i in items) + "</ol>"


def bullets(items):
    return "<ul>" + "".join(f"<li>{i}</li>" for i in items) + "</ul>"


def jury(pairs):
    return "".join(f"<details><summary>{esc(q)}</summary><p>{a}</p></details>" for q, a in pairs)


def role(n, name, brief, answer, figures, reasoning, open_items, qa, tabs):
    REPORT_ROLES.append({
        "n": n, "name": name, "brief": brief, "answer": plain(answer), "reasoning": [plain(x) for x in reasoning],
        "open": [plain(x) for x in open_items], "qa": [(plain(q), plain(a)) for q, a in qa], "tabs": tabs,
        "figures": list(LAST_ROWS),
    })
    return f"""
<section class="role" id="role-{n}" aria-labelledby="h-{n}">
  <header class="role-head">
    <span class="role-n">{n}</span>
    <div><h2 id="h-{n}">{esc(name)}</h2><p class="brief">{esc(brief)}</p></div>
  </header>
  <p class="answer">{answer}</p>
  <div class="role-grid">
    <div class="col-text">
      <h3>Le raisonnement</h3>{steps(reasoning)}
      <h3>Ce qui reste ouvert</h3>{bullets(open_items)}
      <h3>Questions probables du jury</h3>{jury(qa)}
      <p class="tabs">Dans le modèle : {esc(tabs)}</p>
    </div>
    <div class="col-fig"><h3>Les chiffres</h3>{figures}</div>
  </div>
</section>"""


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
BIND27 = traj("Marge contraignante"); MAXACQB = traj("Acquisition maximale fin 2027 sous la contrainte")
UNF = mv("Cibles", "Marge de manœuvre fin 2027", "C"); FAV = mv("Cibles", "Marge de manœuvre fin 2027", "E")
_wc = WB["Cibles"]
_hr = next(c.row for c in _wc["B"] if isinstance(c.value, str) and c.value.startswith("Marge contraignante"))
grid = {k: [_wc[f"{c}{r}"].value for c in "BCDE"] for k, r in zip((8, 9, 10, 11), range(_hr + 1, _hr + 5))}
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
RESID25 = mv("Pont de dette", "Flux non détaillés")
CE_GAP = mv("Pont de dette", "Clean Earth : effet sur la dette moins")
MINOR25 = mv("Pont de dette", "dont au moins")
_wv = WB["Vérifications"]
checks_ok = next(_wv[f"F{c.row}"].value for c in _wv["B"] if c.value == "Statut global des contrôles bloquants")
SPENT = {k: mv("Cessions", lab) for k, lab in [("s24", "Sorties nettes 2024"), ("s25", "Sorties nettes 2025"),
         ("sH1", "Sorties nettes S1 2026"), ("cum", "Sorties nettes cumulées"), ("vs", "… face à l'enveloppe boosters")]}
MAXDEBT_TGT = mv("Levier", "Dette maximale fin 2027 si l'objectif"); MAXDEBT_MOD = mv("Levier", "Dette maximale fin 2027 à l'EBITDA du modèle")
PERHALF = mv("Levier", "Dette libérée par 0,5")
FFO25 = mv("Levier", "FFO 2025 implicite"); FFOREQ = mv("Levier", "FFO requis en 2026"); FFOGAP = mv("Levier", "Marge de FFO en 2026")
DEBT18 = mv("Levier", "Dette ajustée maximale au seuil"); ADJGAP = mv("Levier", "Écart entre dette ajustée")
FFO_REC = mv("Levier", "FFO reconstitué (somme"); FFO_GAP_PCT = mv("Levier", "Écart en % du FFO Moody's")
COV26 = mv("Échéancier", "Liquidités / flux contractuels 2026"); COV26X = mv("Échéancier", "Liquidités / flux 2026 hors")
EXTRA_INT = mv("Échéancier", "Surcoût d'intérêts annuel"); EXTRA_PCT = mv("Échéancier", "en % du FFO 2025")

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
    boost4=R("GreenUp : investissements nets de croissance 2024-2027 dans les boosters", "2024-2027"),
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
HYB_SHARE = mv("Levier", "Part des hybrides dans l'écart")
SEG_EB27 = mv("Segments", "EBITDA 2027 « au rythme", "E"); CE_EBREC = mv("Booster", "EBITDA 2025 recalculé"); CE_EBGAP = mv("Booster", "Écart entre l'EBITDA de Veolia")
M_REC = mv("Booster", "Multiple sur l'EBITDA 2025 reconstitué"); CE_MREC = mv("Booster", "Marge d'EBITDA 2025 recalculée")
CLH_ENVPCT = mv("ESG", "Clean Harbors : passifs environnementaux en %"); EN_ENV = mv("ESG", "Enviri (groupe entier)"); CE_ANALOG = mv("ESG", "Ordre de grandeur par analogie")

ROLES = []
ROLES.append(role(
    1, "Périmètre et sources", "Le registre, et si l'enveloppe annoncée était brute ou nette des cessions.",
    f"<strong>Nette.</strong> GreenUp annonce environ {P(f['rot'])} Md€ par an de rotation <em>nette</em> d'actifs : "
    f"{P(f['tlo'])} à {P(f['thi'])} Md€ de tuck-ins, moins environ {P(f['dcont'])} Md€ de cessions. "
    f"Clean Earth ({fr(-N(f['ce_nfd']) / 1000, 2)} Md€ de dette en plus) vaut {fr(CE_YEARS, 1)} années "
    f"de cette enveloppe ; c'est le programme de cessions de plus de {P(f['prog'])} Md€ qui le ramène dedans. "
    f"Les {P(f['boost4'])} Md€ d'investissements de croissance dans les boosters sont eux aussi annoncés <em>nets</em> (GreenUp p.44) : "
    f"au 30 juin 2026, les sorties nettes d'acquisitions depuis 2024 atteignent {fr(SPENT['cum'], 2)} Md€, soit {fr(SPENT['vs'], 2)} Md€ de plus.",
    table([
        row_fig("GreenUp : investissements nets de croissance, boosters", f["boost4"]),
        row_fig("GreenUp : strongholds", f["strong2"]),
        row_fig("Investissements financiers nets des cessions (entrée)", f["nfi24"]),
        row_fig("Investissements financiers nets des cessions", f["nfi25"]),
        row_fig("Investissements financiers nets des cessions", f["nfih1"]),
        row_calc("Sorties nettes cumulées 2024 → 30/06/2026", fr(SPENT['cum'], 2, "Md EUR"), "somme des trois ponts, signe inversé"),
        row_calc("… face aux 4 Md€ de l'enveloppe boosters", fr(SPENT['vs'], 2, "Md EUR"), "positif : enveloppe nette dépassée"),
        row_fig("Rotation déjà réalisée selon Veolia (brut)", f["rot4"]), row_fig("Rotation totale prévue (brut)", f["rot85"]),
        row_fig("Part des acquisitions dans les boosters", f["pctB"]), row_fig("Part des acquisitions hors d'Europe", f["pctX"]),
        row_fig("Rotation nette d'actifs annoncée", f["rot"]), row_fig("Tuck-ins, bas de fourchette", f["tlo"]),
        row_fig("Tuck-ins, haut de fourchette", f["thi"]), row_fig("Cessions d'actifs non stratégiques", f["dcont"]),
        row_calc("Rotation nette implicite (haut − cessions)", f"0,5{NB}Md EUR", "recalcul : égale l'annonce"),
        row_fig("Investissements nets annoncés", f["netcap"]), row_fig("dont maintenance", f["maint"]),
        row_fig("dont croissance sur contrats", f["grow"]), row_fig("moins cessions industrielles (au plus)", f["inddisp"]),
        row_fig("Rotation d'actifs sur 4 ans (brut)", f["rot8"]),
        row_fig("Clean Earth : effet sur la dette", f["ce_nfd"]), row_fig("Programme de cessions (au moins)", f["prog"]),
        row_calc("Clean Earth net du programme", fr(CE_NET, 2, "Md EUR"), f"{fr(CE_NET_Y, 1)} année d'enveloppe nette"),
    ]),
    [
        "Lire la diapositive d'allocation du capital de GreenUp (p.63) : chaque colonne est exprimée en net.",
        f"Recalculer les deux enveloppes : {P(f['maint'])} + {P(f['grow'])} − {P(f['inddisp'])} = "
        f"{P(f['netcap'])} Md€ d'investissements nets ; {P(f['thi'])} − {P(f['dcont'])} = {fr(N(f['thi']) - N(f['dcont']), 1)} Md€ de rotation "
        "nette. L'annonce correspond au haut de la fourchette de tuck-ins.",
        f"Ne pas confondre avec les « plus de {P(f['rot8'])} Md€ de rotation d'actifs en 4 ans » de la présentation du S1 2026 : "
        "ce chiffre additionne achats et ventes, il est brut.",
        f"Placer Clean Earth : {fr(CE_YEARS, 1)} années d'enveloppe nette en une opération. Net du programme de cessions, "
        f"il en reste {fr(CE_NET, 2)} Md€, soit {fr(CE_NET_Y, 1)} année.",
        f"Reconstituer ce qui est sorti : {fr(-SPENT['s24'], 2)} Md€ rentrés en 2024 (cessions SADE, RGS, Lydec, Haikou), "
        f"{fr(SPENT['s25'], 2)} Md€ sortis en 2025 (minoritaires WTS, déchets dangereux), {fr(SPENT['sH1'], 2)} Md€ au S1 2026 (Clean Earth). "
        f"Cumul net : {fr(SPENT['cum'], 2)} Md€, contre une enveloppe boosters de {P(f['boost4'])} Md€ sur quatre ans.",
        "Le registre : chaque chiffre a son fichier, sa page, son URL et sa date de consultation, et un contrôle automatique "
        "vérifie que la valeur figure bien sur la page citée.",
    ],
    [
        "GreenUp ne dit pas si la fourchette de tuck-ins est annuelle ; la colonne l'est (« / year »). À formuler prudemment.",
        "« Plus de 2 Md€ » est un plancher : un programme plus gros changerait la lecture.",
        "Aucun chiffre ni aucun deal n'a encore été relu par un deuxième membre du groupe.",
    ],
    [
        ("Pourquoi dites-vous que l'enveloppe est nette ?",
         "Parce que la diapositive la présente comme une rotation nette et que les composantes publiées se recalculent : "
         f"{P(f['thi'])} Md€ de tuck-ins moins {P(f['dcont'])} Md€ de cessions donnent les {P(f['rot'])} Md€ annoncés. "
         "Veolia applique la même convention aux investissements industriels."),
        ("Les 4 Md€ annoncés étaient-ils bruts ou nets des cessions ?",
         f"Nets : la diapositive d'allocation du capital parle de « net growth investments » pour les boosters. Le brut, Veolia l'appelle "
         f"« rotation d'actifs » : {P(f['rot4'])} Md€ déjà faits, {P(f['rot85'])} Md€ au total avec Clean Earth et le plan de cessions."),
        ("Les 8 Md€ de rotation ne contredisent-ils pas les 0,5 Md€ par an ?",
         "Non : les 8 Md€ sont un montant brut (achats et ventes additionnés) sur quatre ans, les 0,5 Md€ un solde annuel."),
        ("Comment avez-vous constitué le registre ?",
         "À partir des documents publiés par Veolia (DEU, amendements, communiqués, présentations), une ligne par chiffre avec fichier, page, "
         "URL et date. Les opérations viennent des notes « Principales évolutions de périmètre » des comptes consolidés."),
    ],
    "Entrées, Cessions §A et §B",
))

ROLES.append(role(
    2, "La contrainte", "La position de départ, la définition du levier propre à Veolia, et le calcul jusqu'au plafond.",
    f"Veolia mesure son levier comme <strong>dette financière nette de clôture (IFRS 16 incluse) sur EBITDA (IFRS 16 inclus)</strong>, "
    f"hybrides exclus. La définition redonne les ratios publiés. Au 30 juin 2026, le ratio mécanique de {x(LMECH)} surestime : "
    f"corrigé de Clean Earth en année pleine et de la saisonnalité, on est à {x(LSEAS)}. Le modèle donne {x(LEV26)} fin 2026, "
    f"dans la guidance, et {x(LEV27)} fin 2027, soit {fr(HEAD / 1000, 2)} Md€ de marge sous 3x. "
    f"<strong>Mais la vraie limite est celle des agences</strong> : S&P abaisse la note si FFO / dette ajustée ne reste pas au-dessus de "
    f"{P(f['sptrig'])} %, Moody's si le ratio passe sous « {esc(f['motrig']['value'])} ». Au pic de dette ajustée de 2026 "
    f"(~{P(f['mond26'])} Md€), tenir 18 % demande {fr(FFOREQ / 1000, 1)} Md€ de FFO, soit le FFO 2025 ({fr(FFO25 / 1000, 1)} Md€) : "
    f"à FFO constant il n'y a pas de marge. Si le FFO suit l'EBITDA, le modèle donne {pct(RATIO26)} fin 2026 (Moody's attend "
    f"{P(f['mo26lo'])}-{P(f['mo26hi'])} %) et {pct(RATIO27)} fin 2027 : <strong>le seuil des agences mord en 2026, le 3x de Veolia "
    f"mord en 2027</strong> ({fr(HEADSP27 / 1000, 1)} Md€ de marge sous 18 % contre {fr(HEAD / 1000, 2)} Md€ sous 3x). "
    f"Dans le scénario défavorable, la marge d'agence 2026 tombe à {fr(HEADSP26_UNF / 1000, 2)} Md€. "
    f"L'écart de {fr(N(f['mo_net']) - N(f['nfd25']), 0)} M€ entre la dette des agences et celle de Veolia se réconcilie poste par poste "
    f"(Moody's, Exhibit 13) : hybrides comptés à 50 % ({P(f['mo_hyb'])}), pensions ({P(f['mo_pens'])}), titrisation ({P(f['mo_sec'])}), "
    f"retraitements ({P(f['mo_ns'])}), et une trésorerie retenue plus basse. Moody's publie aussi le FFO : {P(f['mo_ffo25'])} M€ en 2025, "
    f"et prévoit {P(f['mo_f26'])} % de FFO / dette nette en 2026, {P(f['mo_f27'])} % en 2027.",
    table([
        row_fig("Dette financière nette", f["nfd24"]), row_fig("EBITDA", f["eb24"]), row_fig("Levier publié", f["lev24"]),
        row_fig("Dette financière nette", f["nfd25"]), row_fig("EBITDA", f["eb25"]), row_fig("Levier publié", f["lev25"]),
        row_calc("Levier recalculé 2025", x(mv("Levier", "Levier recalculé 2025")), "DFN / EBITDA"),
        row_fig("Dette au 30/06/2026, Clean Earth inclus", f["nfdh1"]), row_fig("Dette au 30/06/2025", f["nfdh125"]),
        row_calc("EBITDA glissant 12 mois au 30/06/2026", fr(LTM, 0, "M EUR"), "2025 − S1 2025 + S1 2026"),
        row_calc("① Levier mécanique", x(LMECH), "toute la dette, un mois d'EBITDA de Clean Earth"),
        row_calc("② Pro forma Clean Earth 12 mois", x(LPF), "+ 11 mois d'EBITDA de Clean Earth"),
        row_calc("③ Corrigé de la saisonnalité", x(LSEAS), "− la baisse de dette du S2 2025"),
        row_fig("Levier publié hors Clean Earth", f["q1"]),
        row_calc("Dette maximale fin 2027 si l'objectif ≥ 8 Md€ est atteint", fr(MAXDEBT_TGT, 0, "M EUR"), "3 × 8 000"),
        row_calc("Dette maximale fin 2027 à l'EBITDA du modèle", fr(MAXDEBT_MOD, 0, "M EUR"), "3 × EBITDA 2027 central"),
        row_calc("Dette libérée par 0,5 Md€ d'EBITDA en plus", fr(PERHALF, 0, "M EUR"), "3 × 500"),
        row_fig("Maturité moyenne de la dette nette", f["mat25"]), row_fig("Part à taux fixe après couverture", f["fixed"]),
        row_fig("Dettes hybrides (capitaux propres)", f["hyb"]),
        row_fig("Hybride vert 2025 : baisse de la dette nette", f["hybg"]),
        row_fig("Moody's : FFO / dette nette ajustée", f["mo24"]), row_fig("Moody's : FFO / dette nette ajustée", f["mo25"]),
        row_fig("Moody's : attendu 2026, bas", f["mo26lo"]), row_fig("Moody's : attendu 2026, haut", f["mo26hi"]),
        row_fig("Moody's : dégradation si FFO / dette nette sous", f["motrig"]), row_fig("Moody's : relèvement si dans", f["moup"]),
        row_fig("Moody's : dette nette ajustée", f["mond25"]), row_fig("Moody's : dette nette ajustée, pic", f["mond26"]),
        row_fig("S&P : dégradation si FFO / dette durablement sous", f["sptrig"]),
        row_fig("S&P : attendu 2026-2028, bas", f["splo"]), row_fig("S&P : attendu 2026-2028, haut", f["sphi"]),
        row_calc("FFO 2025 implicite", fr(FFO25, 0, "M EUR"), "20,3 % × 25 400"),
        row_calc("FFO requis en 2026 au seuil S&P", fr(FFOREQ, 0, "M EUR"), "18 % × ~29 000"),
        row_calc("Marge de FFO en 2026 au seuil S&P", fr(FFOGAP, 0, "M EUR"), "FFO 2025 − requis"),
        row_calc("Écart dette ajustée Moody's − dette nette publiée 2025", fr(ADJGAP, 0, "M EUR"), "retraitements d'agence"),
        row_fig("Moody's : dette brute publiée", f["mo_rep"]), row_fig("Sous-total des emprunts (DEU)", f["urd_gross"]),
        row_fig("+ pensions", f["mo_pens"]), row_fig("+ hybrides à 50 %", f["mo_hyb"]), row_fig("+ titrisation", f["mo_sec"]),
        row_fig("+ ajustements non standard", f["mo_ns"]), row_fig("= dette brute ajustée", f["mo_adj"]),
        row_fig("− trésorerie retenue", f["mo_cash"]), row_fig("= dette nette ajustée", f["mo_net"]),
        row_calc("Part des hybrides dans l'écart", pct(HYB_SHARE), "hybrides / (ajustée − publiée)"),
        row_fig("Moody's : FFO", f["mo_ffo24"]), row_fig("Moody's : FFO", f["mo_ffo25"]), row_fig("Moody's : dividendes", f["mo_div"]),
        row_fig("Moody's : RCF", f["mo_rcf"]), row_fig("Moody's : dette nette / EBITDA ajustés", f["mo_ndeb"]),
        row_fig("Moody's : FFO / dette nette prévu", f["mo_f26"]), row_fig("Moody's : FFO / dette nette prévu", f["mo_f27"]),
        row_fig("Guidance fin 2026 : égal ou légèrement au-dessus de", f["guid"]),
        row_fig("Engagement 2027 : au plus", f["cap"]),
        row_calc("Levier fin 2026 (modèle)", x(LEV26), f"pro forma 12 mois : {x(LEV26PF)}"),
        row_calc("Levier fin 2027 (modèle)", x(LEV27), f"marge {fr(HEAD, 0)} M EUR"),
        row_calc("… si 50 % des hybrides comptent en dette", x(D50), "lecture d'analyste"),
        row_fig("Flux contractuels de dette 2026, principal et intérêts", f["flow26"]), row_fig("Flux contractuels de dette 2027", f["flow27"]),
        row_fig("Liquidités totales", f["liq"]), row_fig("Ligne syndiquée non tirée, jusqu'en 2030", f["synd"]),
        row_calc("Liquidités / flux de dette 2026", x(COV26), f"{x(COV26X)} hors billets de trésorerie et hybride notifié"),
        row_calc("Surcoût d'intérêts annuel, souches 2027-2028 refinancées au taux de juin 2025", fr(EXTRA_INT, 0, "M EUR"),
                 f"{pct(EXTRA_PCT)} du FFO 2025"),
        row_fig("Capacité d'autofinancement avant BFR (tableau de flux)", f["cfo25"]), row_fig("Intérêts payés", f["int25"]),
        row_calc("FFO 2025 reconstitué depuis le tableau de flux", fr(FFO_REC, 0, "M EUR"),
                 f"Moody's : {P(f['mo_ffo25'])} ; écart {pct(FFO_GAP_PCT)}"),
    ]),
    [
        f"Prendre la définition dans le DEU 2025 (p.355) et la tester : elle redonne {P(f['lev24'])}x pour 2024 et {P(f['lev25'])}x pour 2025.",
        f"Les {P(f['hyb'])} Md€ d'hybrides sont des capitaux propres en IFRS (IAS 32.11). L'hybride vert de 2025 a réduit la dette "
        f"nette de {P(f['hybg'])} M€ : c'est un levier sur le ratio, pas sur l'économie.",
        f"Au 30 juin 2026, trois lectures : {x(LMECH)} mécanique, {x(LPF)} avec douze mois de Clean Earth, {x(LSEAS)} en retirant "
        f"la saisonnalité (la dette baisse de {fr(DROP_H2, 0)} M€ au second semestre). Seule la troisième se compare à un ratio de décembre.",
        f"Le plafond en euros : si l'objectif de 8 Md€ est atteint, la dette ne peut pas dépasser {fr(MAXDEBT_TGT / 1000, 0)} Md€ fin 2027 ; "
        f"la dette de juin 2026 ({P(f['nfdh1'])} M€) est déjà au-dessus. Chaque 0,5 Md€ d'EBITDA en plus libère {fr(PERHALF / 1000, 1)} Md€ : "
        "la capacité dépend plus de l'EBITDA que de n'importe quelle décision d'allocation.",
        f"Fin 2026, le modèle donne {x(LEV26)} (dette {fr(NFD26, 0)} M€, EBITDA {fr(EB26, 0)} M€). Il retombe dans la guidance sans "
        "réglage : c'est le test de calibrage.",
        f"Fin 2027 : {x(LEV27)}, marge de {fr(HEAD, 0)} M€ sous 3x. Avec la moitié des hybrides en dette, le ratio passe à {x(D50)} "
        f"et la marge disparaît ; avec tous les hybrides, {x(D100)}.",
        f"Les agences, publiées par Veolia elle-même sur sa page « Debt and ratings » : Moody's (Baa1 stable, mai 2026) mesure FFO / dette "
        f"nette ajustée à {P(f['mo25'])} % en 2025 et attend {P(f['mo26lo'])}-{P(f['mo26hi'])} % en 2026 avec un pic de dette ajustée vers "
        f"{P(f['mond26'])} Md€ ; dégradation si le ratio passe sous « {esc(f['motrig']['value'])} ». S&P (BBB stable, avril 2026) : "
        f"dégradation si FFO / dette ne reste pas durablement au-dessus de {P(f['sptrig'])} %. La dette ajustée des agences dépasse de "
        f"{fr(ADJGAP / 1000, 1)} Md€ la dette nette publiée : hybrides comptés en partie, pensions, provisions.",
        f"La réconciliation, poste par poste (Moody's Exhibit 13) : dette brute publiée {P(f['mo_rep'])} M€ (c'est le sous-total des "
        f"emprunts du DEU, {P(f['urd_gross'])}), plus pensions {P(f['mo_pens'])}, hybrides à 50 % {P(f['mo_hyb'])}, titrisation "
        f"{P(f['mo_sec'])}, retraitements {P(f['mo_ns'])} = {P(f['mo_adj'])} ; moins une trésorerie retenue de {P(f['mo_cash'])} = "
        f"{P(f['mo_net'])}. Les hybrides pèsent {pct(HYB_SHARE, 0)} de l'écart. Moody's mesure 3,7x de dette nette / EBITDA ajustés "
        f"là où Veolia publie 2,79x : même entreprise, deux définitions.",
        f"Le mur de refinancement n'en est pas un : les flux contractuels de dette 2026 ({P(f['flow26'])} M€, dont {P(f['cpap'])} de billets "
        f"de trésorerie renouvelés en continu) sont couverts {x(COV26)} par {P(f['liq'])} M€ de liquidités, dont une ligne syndiquée de "
        f"{P(f['synd'])} M€ non tirée jusqu'en 2030 ; 2027 pèse {P(f['flow27'])} M€ (DEU p.421). Aucun covenant financier sur la dette de "
        f"Veolia Environnement (p.422). Ce que le mur change, c'est le FFO : les souches de 2027-2028 portent des coupons de 0 à 1,6 % et "
        f"se refinanceront vers 3,3 % (émission de juin 2025), soit ~{fr(EXTRA_INT, 0)} M€ d'intérêts de plus par an ({pct(EXTRA_PCT)} du FFO).",
        f"Le FFO des agences se reconstitue depuis le tableau de flux : capacité d'autofinancement avant BFR {P(f['cfo25'])} M€, moins impôts "
        f"et intérêts payés, plus remboursements d'actifs financiers opérationnels et dividendes reçus = {fr(FFO_REC, 0)} M€, contre "
        f"{P(f['mo_ffo25'])} publiés par Moody's (écart {pct(FFO_GAP_PCT)}). Le classeur le contrôle à ± 5 % sur 2024 et 2025.",
    ],
    [
        f"Veolia ne dit pas si son ratio de fin 2026 comptera douze mois de Clean Earth ({x(LEV26PF)}) ou sept ({x(LEV26)}).",
        "Les coupons des deux souches euro de 2026 (750 et 650 M€) ne figurent pas au tableau p.407 (lignes courantes) : "
        "le surcoût de refinancement de 2026 n'est pas chiffré.",
        "Le change sur la dette en dollars n'est pas modélisé.",
    ],
    [
        ("Pourquoi ne pas citer 3,39x au 30 juin ?",
         f"Parce qu'il rapporte toute la dette de Clean Earth à un seul mois de son EBITDA, et une dette de juin, toujours plus haute, à un "
         f"plafond qui se mesure en décembre. Corrigé des deux, on est à {x(LSEAS)}."),
        ("Quel est le vrai plafond : 3x ou les agences ?",
         f"Les deux, à des dates différentes. Veolia s'engage à garder BBB / Baa1, et S&P abaisse la note sous 18 % de FFO / dette. "
         f"Au pic de dette de 2026, ce seuil demande {fr(FFOREQ / 1000, 1)} Md€ de FFO pour {fr(FFO25 / 1000, 1)} Md€ générés en 2025 : "
         f"c'est 2026 qui est tendu, et c'est pour cela que les cessions sont annoncées dans les deux ans. Fin 2027, avec un FFO qui suit "
         f"l'EBITDA, le ratio remonte à {pct(RATIO27)} et c'est le 3x de Veolia qui limite la capacité ({WHICH27.split(' : ')[0]} dans le modèle)."),
        ("D'où vient l'écart de 5,7 Md€ entre la dette des agences et la dette publiée ?",
         f"De la réconciliation que Moody's publie : la dette brute est la même ({P(f['mo_rep'])} M€), Moody's y ajoute pensions, "
         f"titrisation, retraitements et la moitié des hybrides ({P(f['mo_hyb'])} M€), puis retient moins de trésorerie "
         f"({P(f['mo_cash'])} contre 8 021 + 1 952 chez Veolia). Chaque poste est au registre avec sa page."),
        ("Les hybrides sont-ils de la dette ?",
         "Pas pour Veolia ni en IFRS. Mais ils portent un coupon et une date de rappel : un lecteur prudent en compte une partie, "
         f"et le plafond de 2027 n'est alors plus tenu ({x(D50)} à 50 %)."),
        ("Y a-t-il un mur de refinancement avant 2027 ?",
         f"Non. {P(f['flow26'])} M€ de flux contractuels en 2026, dont {P(f['cpap'])} de billets de trésorerie qui se renouvellent, face à "
         f"{P(f['liq'])} M€ de liquidités : couvert {x(COV26)}. Le mur agit sur le FFO, pas sur la dette : environ {fr(EXTRA_INT, 0)} M€ "
         f"d'intérêts de plus par an quand les souches à bas coupon de 2027-2028 se refinancent au taux de 2025."),
        ("Comment reconstituez-vous le FFO de Moody's, qui ne publie pas sa formule ?",
         f"Depuis le tableau de flux du DEU (p.362-363) : capacité d'autofinancement avant BFR, moins impôts et intérêts payés, plus "
         f"remboursements d'actifs financiers opérationnels et dividendes reçus : {fr(FFO_REC, 0)} M€ contre {P(f['mo_ffo25'])} publiés, "
         f"écart {pct(FFO_GAP_PCT)}. Même lecture en 2024 à moins de 2 % près : c'est un contrôle bloquant du classeur."),
        ("Pourquoi votre fin 2026 est-il crédible ?",
         f"Il n'utilise que la guidance de croissance (+{P(f['g_lo'])} à +{P(f['g_hi'])} %), le second semestre 2025 reproduit et la moitié des cessions signées : "
         f"il tombe à {x(LEV26)}, ce que Veolia annonce."),
    ],
    "Levier, Pont de dette, Échéancier",
))

top3 = "; ".join(f"{esc(n)} ({fr(v / 1000, 2)} Md€)" for n, v in rank[:3])
ROLES.append(role(
    3, "Capacité et sensibilité",
    "Le cash-flow libre, le programme de cessions et son calendrier, et ce qui fait le plus bouger la réponse.",
    f"Le cash-flow libre net ({P(f['nfcf25'])} M€ en 2025) paie à peine les dividendes : la marge vient de la croissance de "
    f"l'EBITDA et des cessions. Ce qui bouge le plus la marge de fin 2027 : {top3}. Si tout tourne mal en même temps, "
    f"la marge devient négative ({fr(UNF, 0)} M€).",
    table([
        row_fig("Cash-flow libre net", f["nfcf24"]), row_fig("Cash-flow libre net", f["nfcf25"]),
        row_fig("Cash-flow libre net, S1", f["nfcfh125"]), row_fig("Cash-flow libre net, S1", f["nfcfh126"]),
        row_calc("Cash-flow libre net du S2 2025", fr(FCF_H2, 0, "M EUR"), "année − S1"),
        row_fig("Variation du BFR, S1", f["wcr"]), row_fig("Investissements nets", f["capex"]),
        row_fig("Dividende versé aux actionnaires", f["div"]), row_fig("Dividendes versés, minoritaires compris", f["divt"]),
        row_fig("Programme de cessions (au moins)", f["prog"]), row_fig("Cessions réalisées", f["closed"]),
        row_fig("Signatures attendues", f["sign"]),
        row_calc("Calendrier central S2 2026 / 2027 / S1 2028", " / ".join(fr(v, 2) for v in CAL[1:]) + f"{NB}Md EUR",
                 "hypothèses s26 et s27"),
        row_fig("Guidance croissance organique EBITDA, bas", f["g_lo"]), row_fig("… haut", f["g_hi"]),
        row_calc("Cash-flow libre net 2027 (modèle)", fr(FCF27, 0, "M EUR"), "conversion moyenne 2024-2025"),
        row_calc("Dividendes 2027 (modèle)", fr(DIV27, 0, "M EUR"), "+8 %, minoritaires constants"),
        row_calc("Marge centrale fin 2027", fr(HEAD, 0, "M EUR"), f"défavorable {fr(UNF, 0)} · favorable {fr(FAV, 0)}"),
    ]),
    [
        f"La saisonnalité d'abord : le premier semestre consomme du cash (BFR {P(f['wcr'])} M€), le second en rend "
        f"({fr(FCF_H2, 0)} M€ en 2025). Toute analyse en juin doit en tenir compte.",
        f"Le calendrier : {P(f['closed'])} Md€ encaissés, environ {P(f['sign'])} Md€ de signatures attendues en 2026, "
        "le solde d'ici mi-2028. Signer n'est pas encaisser : le scénario central encaisse la moitié des signatures en 2026 et la moitié du reste en 2027.",
        "Ce que vaut chaque levier sur la marge 2027 : " + ", ".join(f"{esc(k)} {fr(v, 0)}" for k, v in ELAST.items()) + " (M€ pour +100 M€).",
        "Le classement : " + "; ".join(f"{i + 1}. {esc(n)} ({fr(v, 0)} M€)" for i, (n, v) in enumerate(rank[:6])) + ".",
        f"Les combinaisons : tout défavorable donne {fr(UNF, 0)} M€, tout favorable {fr(FAV, 0)} M€. La réponse dépend moins du "
        "point central que du calendrier des cessions.",
    ],
    [
        "Aucun multiple de cession n'est publié : on suppose 10x, entre 8x et 12x.",
        "Aucun dividende n'est supposé versé au second semestre ; le change n'est pas prolongé.",
        f"Le pont de dette 2025 du communiqué laisse {fr(RESID25, 0)} M€ de flux non détaillés (dont au moins {fr(MINOR25, 0)} M€ de dividendes aux minoritaires).",
    ],
    [
        ("Et si les cessions glissent en 2028 ?",
         f"Sans rien encaisser du reste du programme en 2027, la marge tombe de {fr(HEAD, 0)} à {fr(H1_LOW, 0)} M€. C'est le premier facteur."),
        ("Pourquoi ne pas couper le dividende ?",
         "GreenUp s'engage sur un dividende qui suit le bénéfice par action (p.66). Le modèle le traite comme un engagement ; "
         "chaque 100 M€ économisés vaudrait 100 M€ de marge."),
        ("Pourquoi le multiple de cession compte-t-il ?",
         "Chaque euro encaissé réduit la dette d'un euro mais retire 1/multiple d'EBITDA, soit trois fois plus de capacité : à 10x, "
         "100 M€ de cessions ne libèrent que 70 M€ de marge."),
    ],
    "Trajectoire, Sensibilité, Cessions §C, Pont de dette",
))

seche_rows, seche_text = [], ""
ROLES.append(role(
    4, "L'écart et les comparables", "Où se situe le booster face à 2027, et qui est réellement comparable.",
    f"En volume, l'objectif est passé de {P(f['hwt0'])} Mt (« déchets dangereux et polluants ») à {P(f['hwt'])} Mt : "
    f"Veolia, à {P(f['hw25'])} kt, tient l'objectif révisé mais reste à {fr(-HW_GAP0, 0)} kt de l'initial. "
    f"En valeur, l'ambition de plus de {P(f['cagr'])} % par an d'EBITDA repose sur Clean Earth : en Europe, le chiffre "
    f"d'affaires recule de {fr(-N(f['eug']), 1)} % en organique. " + seche_text +
    f"Côté segments, la croissance vient des Amériques (EBITDA organique +{P(f['am_g'])} %) et de Water Technologies (+{P(f['wt_g'])} %), "
    f"pas de l'Europe (+{P(f['eu_g'])} %) : au rythme de 2025, l'EBITDA 2027 atteindrait {fr(SEG_EB27 / 1000, 2)} Md€, l'objectif tient. "
    f"Chez Clean Harbors, le comparable direct, la marge est de {P(f['clh_m'])} % ; chez Clean Earth vu par son vendeur, "
    f"{pct(CE_MREC)} sur un EBITDA 2025 reconstitué de {fr(CE_EBREC, 0)} M$, un tiers sous les {P(f['ceeb'])} M$ 2026E de Veolia.",
    table([
        row_fig("Déchets dangereux traités", f["hw24"]), row_fig("Déchets dangereux traités", f["hw25"]),
        row_fig("Objectif 2027 actuel", f["hwt"]), row_fig("Objectif 2027 initial", f["hwt0"]),
        row_fig("Capacité en construction", f["c430"]),
        row_fig("CA pro forma avec Clean Earth", f["pfrev"]), row_fig("EBITDA pro forma", f["pfeb"]),
        row_fig("Marge pro forma", f["pfm"]), row_fig("Ambition EBITDA, par an, plus de", f["cagr"]),
        row_calc("EBITDA 2027 si +10 %/an depuis le pro forma", fr(HW_ILL, 2, "Md EUR"), "illustratif"),
        row_fig("CA Déchets dangereux Europe", f["eurev"]), row_fig("… croissance organique", f["eug"]),
        row_calc("EBITDA Amériques-Asie-Afrique, S1/S1", pct(AM_G), "publié, périmètre courant"),
        row_fig("Clean Earth : chiffre d'affaires", f["cer26"]), row_fig("Clean Earth : EBITDA", f["ceeb"]),
        row_calc("Clean Earth : marge 2026E", pct(CE_MARGIN), "EBITDA / CA"),
        row_fig("EBITDA organique Water Technologies", f["wt_g"]), row_fig("EBITDA organique Amériques-Asie-Afrique", f["am_g"]),
        row_fig("EBITDA organique Europe", f["eu_g"]), row_fig("Boosters : chiffre d'affaires", f["boost_rev"]),
        row_fig("Boosters : EBITDA", f["boost_eb"]), row_fig("Boosters : croissance organique de l'EBITDA", f["boost_g"]),
        row_calc("EBITDA 2027 si chaque segment garde son rythme de 2025", fr(SEG_EB27, 0, "M EUR"), "onglet Segments §C"),
        row_fig("Clean Harbors : chiffre d'affaires", f["clh_rev"]), row_fig("Clean Harbors : EBITDA ajusté", f["clh_eb"]),
        row_fig("Clean Harbors : marge", f["clh_m"]),
        row_fig("Enviri : CA du segment Clean Earth", f["en_rev"]), row_fig("Enviri : résultat opérationnel Clean Earth", f["en_oi"]),
        row_calc("Clean Earth : EBITDA 2025 reconstitué (RO + D&A)", fr(CE_EBREC, 0, "M USD"), "Enviri 10-K p.70 + p.284"),
        row_calc("Clean Earth : marge 2025 reconstituée", pct(CE_MREC), ""),
    ] + seche_rows),
    [
        "Volumes : l'objectif a changé de niveau et de périmètre entre GreenUp (2024) et le DEU 2025. Le dire avant de comparer.",
        f"Profil financier : avec Clean Earth, le métier pèse {P(f['pfrev'])} Md€ de chiffre d'affaires et {P(f['pfeb'])} Md€ "
        f"d'EBITDA. À +10 % par an depuis ce pro forma, on viserait environ {fr(HW_ILL, 2)} Md€ en 2027 (illustratif).",
        f"Moteur organique : l'Europe est à {pct(EU_G)} au S1 2026, les Amériques progressent de {pct(AM_G)} en publié. "
        "La croissance du booster est achetée plus que générée.",
        f"Les segments (question 1 du cours) : Amériques +{P(f['am_g'])} % et Water Technologies +{P(f['wt_g'])} % d'EBITDA organique en 2025, "
        f"Europe +{P(f['eu_g'])} %, France-DD +6 % publié. Les boosters font {P(f['boost_rev'])} M€ de CA et {P(f['boost_eb'])} M€ d'EBITDA "
        f"(+{P(f['boost_g'])} % organique). Prolongé deux ans, ce rythme donne {fr(SEG_EB27 / 1000, 2)} Md€ en 2027 : l'objectif tient, "
        "mais par les Amériques et les technologies de l'eau, pas par l'Europe.",
        f"Clean Earth vu du vendeur : {P(f['en_rev'])} M$ de chiffre d'affaires et {P(f['en_oi'])} M$ de résultat opérationnel en 2025, "
        f"soit {fr(CE_EBREC, 0)} M$ d'EBITDA une fois les {fr((N(f['en_dep']) + N(f['en_am'])) / 1000, 1)} M$ de D&A rajoutés. Veolia paie 9,8x "
        f"un EBITDA 2026E de {P(f['ceeb'])} M$, {pct(CE_EBGAP, 0)} plus haut : sur l'EBITDA du vendeur, le prix ressort à {fr(M_REC, 1)}x.",
        "Comparabilité : même mix (incinération, traitement physico-chimique, centres de stockage spécialisés), même géographie "
        "que le booster américain, et même définition d'EBITDA et de levier — Clean Harbors publie un EBITDA « ajusté » selon sa propre "
        "définition, Veolia un EBITDA IFRS 16 inclus.",
        "Candidats : Clean Harbors (États-Unis, concurrent direct de Clean Earth, dans la dataroom), Enviri (le vendeur : ses comptes "
        "donnent l'historique de la cible). Republic Services n'est comparable que par sa filiale US Ecology.",
    ],
    [
        "Les comptes de Clean Harbors et d'Enviri ne se téléchargent pas depuis le serveur (refus 403) : à verser à la main.",
        "Aucune valeur de marché des pairs (capitalisation, VE) n'est dans la dataroom.",
        "L'EBITDA déchets dangereux de Veolia seul en 2024, base de l'ambition de +10 %, n'est pas publié.",
    ],
    [
        ("Le booster est-il en retard ?",
         "En volume, sur l'objectif initial, oui ; sur l'objectif révisé, non. En valeur, il n'est dans les clous que grâce à Clean Earth."),
        ("Pourquoi Clean Harbors est-il le bon comparable ?",
         f"Même métier aux États-Unis, là où Clean Earth opère : {P(f['clh_rev'])} M$ de chiffre d'affaires, marge {P(f['clh_m'])} %, "
         "comptes 10-K audités. Mais son EBITDA est « ajusté » selon sa définition, et il porte ses propres passifs environnementaux : "
         "comparer les marges demande la même base."),
        ("L'objectif de 9 Mt est-il une révision à la baisse ?",
         "Le chiffre baisse et le libellé perd « et polluants ». Le DEU ne l'explique pas : c'est une question à poser, pas à trancher."),
    ],
    "Booster",
))

ROLES.append(role(
    5, "Le coût ESG", "Ce qu'une acquisition ajoute en passif et en exposition.",
    f"Veolia porte {P(f['cl26'])} M€ de provisions de fermeture et post-fermeture ({x(PROV_X)} d'EBITDA), en hausse de "
    f"{fr(N(f['cl26']) - N(f['cl25']), 0)} M€ au S1 2026, dont "
    f"+{fr(N(f['am_cl26']) - N(f['am_cl25']), 0)} M€ dans le segment où est entré Clean Earth. "
    f"Mais <strong>{pct(GW_PCT, 0)} du prix de Clean Earth est encore du goodwill</strong> : l'affectation du prix n'est pas faite, "
    "donc la juste valeur de ses passifs environnementaux n'apparaît pas encore chez Veolia. "
    f"Les comparables donnent l'ordre de grandeur : Clean Harbors porte {pct(CLH_ENVPCT)} de son chiffre d'affaires en passifs "
    f"environnementaux ; appliqué au chiffre d'affaires de Clean Earth, cela ferait {fr(CE_ANALOG, 0)} M$. Le groupe Enviri entier en porte "
    f"{fr(EN_ENV, 0)} M$. C'est une analogie, pas une mesure.",
    table([
        row_fig("Provisions de fermeture et post-fermeture", f["cl24"]), row_fig("Provisions de fermeture et post-fermeture", f["cl25"]),
        row_fig("Provisions de fermeture et post-fermeture", f["cl26"]), row_fig("dont réhabilitation de sites", f["site"]),
        row_fig("dont risques environnementaux", f["envr"]), row_fig("Désactualisation (charge annuelle)", f["unw"]),
        row_fig("Segment Amériques-Asie-Afrique", f["am_cl25"]), row_fig("Segment Amériques-Asie-Afrique", f["am_cl26"]),
        row_fig("Clean Earth : prix payé", f["price"]), row_fig("Clean Earth : goodwill préliminaire", f["gw"]),
        row_calc("Goodwill en % du prix", pct(GW_PCT), "aucune affectation au 30/06/2026"),
        row_fig("Garanties reçues liées aux acquisitions", f["gar25"]), row_fig("Garanties reçues liées aux acquisitions", f["gar26"]),
        row_calc("Provisions de fermeture comptées en dette, levier 2027", x(DPROV), "avec 50 % des hybrides"),
        row_fig("Clean Harbors : passifs environnementaux", f["clh_env"]), row_fig("Clean Harbors : chiffre d'affaires", f["clh_rev"]),
        row_calc("Clean Harbors : passifs environnementaux / CA", pct(CLH_ENVPCT), ""),
        row_fig("Enviri : passifs environnementaux courants", f["en_cur"]), row_fig("Enviri : passifs environnementaux non courants", f["en_lt"]),
        row_calc("Ordre de grandeur pour Clean Earth par analogie Clean Harbors", fr(CE_ANALOG, 0, "M USD"), "analogie, pas une mesure"),
    ]),
    [
        f"Le stock : {P(f['cl25'])} M€ fin 2025, surtout de la réhabilitation de sites ; il coûte {P(f['unw'])} M€ par an "
        "de désactualisation, une charge financière qui ne passe pas dans l'EBITDA.",
        "Ce que Clean Earth a ajouté jusqu'ici : la hausse du segment Amériques est compatible avec son entrée, mais Veolia ne l'attribue pas.",
        f"Ce qu'on ne voit pas : {pct(GW_PCT, 0)} du prix reste en goodwill préliminaire (amendement p.53). IFRS 3 laisse douze mois pour "
        "affecter le prix : la juste valeur des passifs de la cible arrivera au plus tard mi-2027.",
        f"L'exposition : {P(f['sites'])} sites spécialisés et {P(f['permits'])} permis aux États-Unis, une croissance attendue sur les PFAS. "
        "Ce qui est une opportunité de marché est aussi un risque de passif.",
        f"Lu en dette élargie (provisions de fermeture comprises), le levier de 2027 gagne environ {x(N(f['cl26']) / EB27)}.",
    ],
    [
        "Les passifs environnementaux propres à Clean Earth sont dans les comptes d'Enviri (activité cédée) : à verser à la main.",
        f"Les garanties reçues passent de {P(f['gar25'])} à {P(f['gar26'])} M€, sans être attribuées à Clean Earth.",
        "La dataroom ne contient aucun rapport d'agence sur le traitement des provisions environnementales.",
    ],
    [
        ("Combien Clean Earth ajoute-t-il en passifs environnementaux ?",
         "On ne peut pas encore le dire depuis les comptes de Veolia : l'affectation du prix n'est pas faite. La seule trace est la hausse "
         "des provisions du segment Amériques, que Veolia n'attribue pas. La réponse est dans les comptes d'Enviri."),
        ("Les provisions sont-elles de la dette ?",
         f"Pas dans la définition de Veolia. Un prêteur prudent en tient compte : elles ajouteraient environ {x(N(f['cl26']) / EB27)} au levier de 2027."),
        ("Pourquoi 81 % de goodwill est-il un sujet ESG ?",
         "Parce que tant que le prix n'est pas affecté, un passif environnemental découvert plus tard viendra réduire l'actif net acquis "
         "et gonfler le goodwill, sans que personne l'ait vu à l'achat."),
    ],
    "ESG",
))

ROLES.append(role(
    6, "Cibles, puis synthèse", "L'univers de cibles, puis une recommandation calibrée sur l'enveloppe propre.",
    f"Fin 2027, la marge centrale permet une acquisition d'environ <strong>{fr(MAXACQ / 1000, 1)} Md€</strong> au multiple de Clean Earth, "
    f"moins que Clean Earth lui-même ({fr(-N(f['ce_nfd']) / 1000, 1)} Md€). Dans le scénario défavorable, "
    "rien. Recommandation proposée : pas de nouvelle grande opération avant l'encaissement du programme de cessions ; des tuck-ins "
    "au bas de la fourchette en 2027 ; une cible de 1 à 2 Md€ seulement une fois les cessions encaissées.",
    table([
        row_calc(f"Acquisition max. à {fr(g[0], 1)}x : défavorable / central / favorable",
                 " / ".join(fr(v / 1000, 2) for v in g[1:]) + f"{NB}Md EUR", "marge / (1 − 3 / multiple)")
        for g in grid.values()
    ] + [
        row_fig("Clean Earth : valeur d'entreprise", f["ev"]), row_fig("Clean Earth : EBITDA", f["ceeb"]),
        row_fig("Clean Earth : synergies, année 4", f["syn"]),
        row_calc("Multiple avant synergies recalculé", fr(M_PRE, 1) + "x", "VE / EBITDA 2026E"),
        row_calc("Multiple après synergies recalculé", fr(M_POST, 1) + "x", "VE / (EBITDA + synergies)"),
        row_fig("Multiple publié après synergies", f["mult"]),
    ]),
    [
        "La formule : une acquisition ajoute son prix à la dette mais seulement prix / multiple à l'EBITDA. Elle tient sous 3x tant que "
        "prix ≤ marge / (1 − 3 / multiple).",
        f"La grille : au multiple de Clean Earth, {fr(grid[8][2] / 1000, 2)} Md€ en central, 0 en défavorable, "
        f"{fr(grid[8][3] / 1000, 2)} Md€ en favorable ; à {fr(grid[10][0], 0)}x, {fr(grid[10][2] / 1000, 2)} Md€ en central.",
        f"Le calibrage : Clean Earth ne passerait plus. Quatre multiples circulent pour lui : {fr(M_REC, 1)}x sur l'EBITDA 2025 publié par "
        f"le vendeur (reconstitué, onglet Booster §E), {fr(M_PRE, 1)}x sur l'EBITDA 2026E avant synergies, "
        f"{P(f['mult'])}x publié après synergies (non recalculable : {fr(M_POST, 1)}x avec les chiffres publiés), et "
        "18,6x l'EBITDA ajusté des douze derniers mois selon Enviri (document à verser).",
        "Les critères de cible : actifs de traitement (pas de collecte seule), géographie des boosters (États-Unis, Asie), PFAS et nouveaux "
        "polluants, taille compatible avec la grille, et pas de doublon de concurrence en Europe, où Veolia est déjà numéro un "
        "(présentation Clean Earth p.9).",
        "L'univers reste à construire : aucun document de la dataroom ne nomme une cible disponible ; les candidats viendront de la "
        "presse et des rapports de pairs, chacun avec sa valeur publiée ou « non communiqué ».",
    ],
    [
        "L'univers de cibles reste à construire (onglet Cibles §C) : nom, pays, activité, valeur d'entreprise sourcée, multiple.",
        "Aucune valeur d'entreprise de cible non cotée n'est dans la dataroom ; une estimation devra le dire.",
        "La recommandation est une proposition du modèle : au groupe de la défendre ou de l'amender.",
    ],
    [
        ("Pourquoi ne pas recommander une acquisition de 3 Md€ ?",
         f"Parce qu'au multiple de Clean Earth elle demanderait une marge de {fr(3000 * (1 - 3 / N(f['mult'])), 0)} M€ et "
         f"que le scénario central n'en dégage que {fr(HEAD, 0)}, sans même compter un glissement des cessions."),
        ("Quel multiple retenez-vous ?",
         "Celui de Clean Earth après synergies, parce que c'est le seul qu'un acheteur comme Veolia ait publié. Plus le multiple monte, "
         f"plus la taille possible baisse : à {fr(grid[10][0], 0)}x, {fr(grid[10][2] / 1000, 2)} Md€."),
        ("Et si le programme de cessions glisse ?",
         f"Sans encaissement du reste du programme en 2027, la marge baisse de {fr(HEAD - H1_LOW, 0)} M€ et la taille possible au "
         f"multiple de Clean Earth tombe à environ {fr(H1_LOW / (1 - 3 / N(f['mult'])) / 1000, 2)} Md€. La recommandation conditionne "
         "toute opération à l'encaissement des cessions."),
    ],
    "Cibles",
))

nav = "".join(f'<a href="#role-{n}"><span>{n}</span>{esc(t)}</a>' for n, t in
              [(1, "Périmètre"), (2, "Contrainte"), (3, "Capacité"), (4, "Écart"), (5, "ESG"), (6, "Cibles")])
tiles = [
    ("Levier fin 2026", x(LEV26), "guidance : égal ou légèrement au-dessus de 3x"),
    ("Levier fin 2027", x(LEV27), "engagement : au plus 3x"),
    ("Marge sous 3x fin 2027", fr(HEAD / 1000, 2, "Md€"), f"défavorable {fr(UNF / 1000, 2)} · favorable {fr(FAV / 1000, 2)}"),
    ("Marge sous 18 % S&P fin 2027", fr(HEADSP27 / 1000, 2, "Md€"), f"FFO / dette ajustée {pct(RATIO26)} en 2026, {pct(RATIO27)} en 2027"),
    ("Acquisition max. fin 2027", fr(MAXACQB / 1000, 2, "Md€"), f"sous la contrainte qui mord : {esc(WHICH27.split(' : ')[0])}"),
    ("Sorties nettes depuis 2024", fr(SPENT["cum"], 2, "Md€"), f"enveloppe boosters annoncée : {P(f['boost4'])} Md€, nets"),
]
tiles_html = "".join(f'<div class="tile"><p class="t-label">{esc(a)}</p><p class="t-val">{b}</p><p class="t-note">{esc(c)}</p></div>'
                     for a, b, c in tiles)

page = f"""<title>Dossiers du sujet 2</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500&family=IBM+Plex+Sans+Condensed:wght@500;600;700&family=IBM+Plex+Sans:ital,wght@0,400;0,500;0,600;1,400&display=swap">
<style>
/* Layout : une note d'analyste — synthèse chiffrée en tête, puis un dossier par rôle, texte à gauche, chiffres sourcés à droite. */
:root {{
  --bg: #F5F7F9; --surface: #FFFFFF; --fg: #16212C; --muted: #5A6875; --line: #D8DEE4;
  --accent: #1E5A8E; --accent-soft: #E3EDF6; --flag: #9A5B00;
  --display: "IBM Plex Sans Condensed", "Arial Narrow", Arial, sans-serif;
  --body: "IBM Plex Sans", "Segoe UI", Arial, sans-serif;
  --mono: "IBM Plex Mono", ui-monospace, Menlo, Consolas, monospace;
}}
@media (prefers-color-scheme: dark) {{ :root:not([data-theme="light"]) {{
  --bg: #0E141A; --surface: #151D25; --fg: #E4E9EE; --muted: #96A3AF; --line: #27323D;
  --accent: #82B4E3; --accent-soft: #1A2B3B; --flag: #E2AA5F; color-scheme: dark; }} }}
:root[data-theme="dark"] {{
  --bg: #0E141A; --surface: #151D25; --fg: #E4E9EE; --muted: #96A3AF; --line: #27323D;
  --accent: #82B4E3; --accent-soft: #1A2B3B; --flag: #E2AA5F; color-scheme: dark; }}
* {{ box-sizing: border-box; }}
body {{ background: var(--bg); color: var(--fg); font: 15px/1.6 var(--body); }}
.wrap {{ max-width: 1120px; margin: 0 auto; padding-inline: 20px; padding-block: 32px 64px; }}
h1, h2, h3 {{ font-family: var(--display); text-wrap: balance; margin: 0; }}
h1 {{ font-size: 2.1rem; font-weight: 700; letter-spacing: -0.01em; }}
.kicker {{ font: 500 .75rem/1 var(--mono); letter-spacing: .08em; text-transform: uppercase; color: var(--muted); margin: 0 0 10px; }}
.lede {{ max-width: 68ch; color: var(--muted); margin: 10px 0 0; font-size: 1.02rem; }}
.thread {{ margin-top: 28px; background: var(--surface); border: 1px solid var(--line); border-radius: 6px; padding: 22px 24px; }}
.thread h2 {{ font-size: 1.15rem; font-weight: 600; margin-bottom: 8px; }}
.thread p {{ max-width: 75ch; margin: 0; }}
.course .two {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 24px; margin-top: 12px; }}
.course .two > div {{ min-width: 0; }}
.course p + p {{ margin-top: 12px; }}
.tiles {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(210px, 1fr)); gap: 12px; margin-top: 18px; }}
.tile {{ border-top: 3px solid var(--accent); background: var(--bg); padding: 12px 14px; }}
.tile p {{ margin: 0; }}
.t-label {{ font: 500 .72rem/1.3 var(--mono); text-transform: uppercase; letter-spacing: .06em; color: var(--muted); }}
.t-val {{ font: 600 1.9rem/1.2 var(--display); font-variant-numeric: tabular-nums; margin-top: 4px !important; }}
.t-note {{ font-size: .82rem; color: var(--muted); }}
nav.roles {{ position: sticky; top: env(safe-area-inset-top, 0px); z-index: 5; background: var(--bg); border-bottom: 1px solid var(--line);
  margin-top: 28px; padding-block: 10px; display: flex; flex-wrap: wrap; gap: 6px; }}
nav.roles a {{ display: inline-flex; align-items: center; gap: 8px; text-decoration: none; color: var(--fg); font: 500 .9rem/1 var(--display);
  padding: 7px 12px 7px 7px; border: 1px solid var(--line); border-radius: 999px; background: var(--surface); }}
nav.roles a span {{ display: inline-grid; place-items: center; width: 22px; height: 22px; border-radius: 50%; background: var(--accent);
  color: var(--surface); font: 600 .78rem/1 var(--mono); }}
nav.roles a:hover, nav.roles a:focus-visible {{ border-color: var(--accent); outline: none; }}
.role {{ margin-top: 40px; scroll-margin-top: 70px; }}
.role-head {{ display: flex; gap: 16px; align-items: flex-start; }}
.role-n {{ font: 700 2.6rem/1 var(--display); color: var(--accent); min-width: 1.2ch; }}
.role-head h2 {{ font-size: 1.6rem; font-weight: 700; }}
.brief {{ margin: 2px 0 0; color: var(--muted); font-style: italic; }}
.answer {{ margin: 16px 0 0; padding: 14px 18px; background: var(--accent-soft); border-radius: 4px; max-width: 92ch; font-size: 1.02rem; }}
.role-grid {{ display: grid; grid-template-columns: minmax(0, 1fr) minmax(0, 1.05fr); gap: 32px; margin-top: 18px; }}
.role-grid > div {{ min-width: 0; }}
h3 {{ font-size: .8rem; font-family: var(--mono); font-weight: 500; text-transform: uppercase; letter-spacing: .08em; color: var(--muted);
  margin: 18px 0 8px; }}
.col-text h3:first-child, .col-fig h3:first-child {{ margin-top: 0; }}
ol.steps {{ margin: 0; padding-left: 1.3em; display: grid; gap: 6px; }}
ul {{ margin: 0; padding-left: 1.2em; display: grid; gap: 4px; }}
ul li::marker {{ color: var(--flag); }}
details {{ border-top: 1px solid var(--line); padding: 8px 0; }}
details:last-of-type {{ border-bottom: 1px solid var(--line); }}
summary {{ cursor: pointer; font-weight: 500; }}
summary:focus-visible {{ outline: 2px solid var(--accent); outline-offset: 2px; }}
details p {{ margin: 6px 0 0; color: var(--fg); }}
.tabs {{ margin-top: 14px; font: .82rem var(--mono); color: var(--muted); }}
.tbl {{ overflow-x: auto; border: 1px solid var(--line); border-radius: 4px; background: var(--surface); }}
table {{ width: 100%; border-collapse: collapse; font-size: .86rem; }}
th, td {{ text-align: left; padding: 6px 10px; border-bottom: 1px solid var(--line); vertical-align: top; }}
th {{ font: 500 .7rem/1.3 var(--mono); text-transform: uppercase; letter-spacing: .06em; color: var(--muted); background: var(--bg); }}
tr:last-child td {{ border-bottom: 0; }}
td.num, th.num {{ text-align: right; white-space: nowrap; font-variant-numeric: tabular-nums; }}
td.num {{ font-weight: 500; }}
.per {{ display: block; font: .72rem var(--mono); color: var(--muted); }}
td.src {{ font: .74rem/1.4 var(--mono); color: var(--muted); }}
.id {{ color: var(--accent); font-weight: 500; }}
.pg {{ color: var(--fg); }}
tr.calc td {{ background: var(--accent-soft); }}
tr.calc td:first-child {{ font-style: italic; }}
.todo {{ margin-top: 48px; border-top: 1px solid var(--line); padding-top: 20px; }}
.todo h2 {{ font-size: 1.2rem; margin-bottom: 8px; }}
.foot {{ margin-top: 24px; font: .78rem var(--mono); color: var(--muted); max-width: 90ch; }}
@media (max-width: 820px) {{ .role-grid {{ grid-template-columns: minmax(0, 1fr); gap: 20px; }} h1 {{ font-size: 1.7rem; }} }}
@media (prefers-reduced-motion: reduce) {{ html {{ scroll-behavior: auto; }} }}
@media (prefers-reduced-motion: no-preference) {{ html {{ scroll-behavior: smooth; }} }}
</style>
<div class="wrap">
  <p class="kicker">Capstone EDHEC · Sujet 2 · Veolia, GreenUp 2027</p>
  <h1>Combien Veolia peut-elle encore acheter sans dépasser 3x ?</h1>
  <p class="lede">Un dossier par rôle : la réponse en une phrase, le raisonnement, les chiffres avec leur source (identifiant du registre,
  fichier, page), ce qui reste ouvert et les questions probables du jury. Les valeurs calculées viennent du classeur
  <span style="font-family:var(--mono)">modele-greenup-2027.xlsx</span> (contrôles : {esc(checks_ok)}).</p>
  <section class="thread" aria-labelledby="h-thread">
    <h2 id="h-thread">Le fil rouge</h2>
    <p>Veolia annonce des enveloppes nettes. Clean Earth en est sorti ; le programme de cessions de plus de 2 Md€ d'ici mi-2028 l'y
    ramène. Si ce programme est encaissé à temps, il reste fin 2027 environ {fr(HEAD / 1000, 1)} Md€ de marge sous 3x, soit une acquisition
    d'environ {fr(MAXACQ / 1000, 1)} Md€ au multiple de Clean Earth. Cette marge dépend d'abord du calendrier des cessions, de la croissance
    organique de 2027 et des tuck-ins. Elle disparaît si tout tourne mal en même temps, ou si l'on compte la moitié des hybrides en dette.
    Le seuil des agences (FFO / dette ajustée ≥ 18 % chez S&P) est la contrainte de 2026 : {pct(RATIO26)} dans le scénario central, dans
    l'attente de Moody's ; fin 2027 il laisse {fr(HEADSP27 / 1000, 1)} Md€ et c'est le 3x qui mord.
    Le booster déchets dangereux tient son objectif de volume révisé, pas l'initial, et sa croissance est surtout achetée. Le passif
    environnemental de Clean Earth n'est pas encore visible dans les comptes.</p>
    <div class="tiles">{tiles_html}</div>
  </section>
  <section class="thread course" aria-labelledby="h-course">
    <h2 id="h-course">Ce que le cours attend du sujet 2</h2>
    <p>Rapport de 50 à 60 pages, une section par rôle signée par son auteur, chaque chiffre renvoyant à un fichier joint. Sélection
    pour l'oral sur le rapport seul. Note : 30 % rapport, 25 % profondeur du cœur d'analyse, 20 % oral, 25 % défense individuelle du rôle.</p>
    <div class="two">
      <div><h3>Les quatre questions de recherche</h3><ol class="steps">
        <li><strong>Écart stratégique</strong> entre la trajectoire et 2027 (EBITDA, croissance, KPI ESG) ; segments et géographies qui sur- ou sous-performent → rôles 2 et 4.</li>
        <li><strong>Benchmark et opportunités</strong> : qui contribue le plus à l'ESG dans chaque booster ; multiples et synergies des deals récents → rôles 4 et 6.</li>
        <li><strong>Allocation du capital restant</strong> : capex organique ou croissance externe, boosters ou strongholds, critère d'impact ESG par euro → rôles 1, 3 et 6. « Tout le reste s'y alimente. »</li>
        <li><strong>Orientations</strong> : géographies, technologies, synergies (« 398 M€ en 2024 → objectif 530 M€ ») → rôle 6.
        Piège repéré : les {P(f['eff398'])} M€ de 2024 sont des <em>gains d'efficacité</em> ; les synergies Suez cumulées étaient de
        {P(f['syn435'])} M€ fin 2024, objectif relevé à {P(f['syn530'])} M€, réalisé {P(f['syn534'])} M€ fin 2025. À dire, pas à corriger en silence.</li>
      </ol></div>
      <div><h3>Les cinq blocs du chantier « capacité »</h3><ol class="steps">
        <li>Reconstituer ce qui a été dépensé, et dire si les 4 Md€ étaient bruts ou nets → rôle 1, Cessions §D.</li>
        <li>La position de départ : dette, EBITDA, profil d'échéances, définition du levier propre à Veolia → rôle 2, Levier.</li>
        <li>Traduire l'engagement en nombre, et lire les seuils de dégradation des agences, « la vraie limite » → rôle 2, Levier §B'.</li>
        <li>Les trois sources de capacité : cash-flow libre, cessions, marge de levier → rôle 3, Trajectoire.</li>
        <li>La sensibilité : prouver quel paramètre bouge le plus la réponse → rôle 3, Sensibilité.</li>
      </ol><p class="brief">Ce chantier publie sa fourchette début novembre ; les autres groupes se calent dessus.</p></div>
    </div>
    <p><strong>Jalons</strong> : 7 octobre, sources chargées et lues (la dataroom est l'objet à montrer) · 16 octobre, questions à Veolia ·
    17 novembre, brouillon complet · 25 novembre, remise · 2 décembre, soutenance.
    <strong>Règle sur l'IA</strong> : elle construit l'outil qui calcule, elle ne calcule pas ; une page d'annexe décrit les outils utilisés.
    C'est exactement la construction retenue ici : les chiffres sont saisis depuis les documents, le classeur calcule, ce dossier lit le classeur.</p>
  </section>
  <nav class="roles" aria-label="Rôles">{nav}</nav>
  {''.join(ROLES)}
  <section class="todo" aria-labelledby="h-todo">
    <h2 id="h-todo">Ce qu'il reste à faire avant l'oral</h2>
    <ul>
      <li>Relire en croisé les {len(REG)} chiffres du registre et les 26 deals : aucun n'a encore de relecteur.</li>
      <li>Lire les 10-K 2025 d'Enviri et de Clean Harbors, désormais dans la dataroom, et en tirer les chiffres des rôles 4 et 5
      (passifs environnementaux de Clean Earth, marges et levier de Clean Harbors). Reste à verser : le 10-Q du T1 2026 d'Enviri et le
      communiqué de vente du 20 novembre 2025.</li>
      <li>Avant le 7 octobre : montrer la dataroom comme « sources chargées et lues », et noter dans gaps.md ce qui manque encore.</li>
      <li>Rédiger la page d'annexe sur les outils d'IA : dataroom MCP (recherche, registre, contrôles), Claude pour construire le classeur
      et ces dossiers, LibreOffice pour le recalcul. Aucun chiffre produit par le modèle de langage.</li>
      <li>Construire l'univers de cibles (onglet Cibles §C) : une ligne par cible, valeur d'entreprise et multiple sourcés.</li>
      <li>Chacun ouvre son rôle avec l'outil <span style="font-family:var(--mono)">workstream_status</span> de la dataroom pour voir
      ses chiffres et ce qui manque.</li>
    </ul>
  </section>
  <p class="foot">Sources : registre de la dataroom (00_Admin/register.csv), {len(REG)} chiffres au 1er octobre 2026, documents publics
  uniquement. Chaque chiffre vient du registre ou d'une cellule du classeur ; seule exception, le multiple de 18,6x d'Enviri, cité
  d'après sa publication et signalé comme non encore versé. Les hypothèses du groupe sont dans l'onglet Hypothèses du classeur.</p>
</div>
"""
OUT.write_text(page, encoding="utf-8")
# Version autonome pour le vault (servie telle quelle par Caddy) : doctype, charset, et le classeur en lien.
head, body = page.split('<div class="wrap">', 1)
body = body.replace('<span style="font-family:var(--mono)">modele-greenup-2027.xlsx</span>',
                    '<a href="modele-greenup-2027.xlsx" style="font-family:var(--mono)">modele-greenup-2027.xlsx</a>', 1)
standalone = ('<!doctype html>\n<html lang="fr">\n<head>\n<meta charset="utf-8">\n'
              '<meta name="viewport" content="width=device-width, initial-scale=1">\n' + head +
              '<style>body{margin:0}</style>\n</head>\n<body>\n<div class="wrap">' + body + '\n</body>\n</html>\n')
OUT_RANGE = OUT.with_name("fourchette-capacite.md")
rank_txt = "\n".join(f"{i + 1}. {n} — amplitude {fr(v, 0)} M€" for i, (n, v) in enumerate(rank[:5]))
OUT_RANGE.write_text(f"""# Fourchette de capacité d'acquisition — GreenUp 2027

Périmètre « capacité financière », pour les groupes des trois autres périmètres. Généré depuis le classeur
`modele-greenup-2027.xlsx` (contrôles bloquants : {checks_ok}) le {__import__('datetime').date.today().isoformat()}.
Chaque chiffre d'entrée vient de `00_Admin/register.csv` ; les hypothèses sont dans l'onglet Hypothèses.

## Le nombre à respecter

| Fin 2027 | Défavorable | Central | Favorable |
|---|---|---|---|
| Marge de dette sous le levier ≤ 3x (M EUR) | {fr(UNF, 0)} | {fr(HEAD, 0)} | {fr(FAV, 0)} |
| Marge de dette sous le seuil S&P 18 % FFO / dette ajustée (M EUR) | {fr(HEADSP27_UNF, 0)} | {fr(HEADSP27, 0)} | {fr(traj("Marge de dette fin 2027 sous le seuil", FAVC), 0)} |
| Contrainte qui mord en premier | {WHICH27_UNF.split(' : ')[0]} | {WHICH27.split(' : ')[0]} | {traj("Contrainte qui mord", FAVC).split(' : ')[0]} |
| Acquisition maximale au multiple de Clean Earth ({P(f['mult'])}x), M EUR | {fr(traj("Acquisition maximale fin 2027 sous la contrainte", UNFC), 0)} | {fr(MAXACQB, 0)} | {fr(traj("Acquisition maximale fin 2027 sous la contrainte", FAVC), 0)} |

Lecture : une acquisition ajoute son prix à la dette et seulement prix ÷ multiple à l'EBITDA ; elle tient tant que
prix ≤ marge ÷ (1 − 3 ÷ multiple). Au multiple de Clean Earth, 1 Md€ de marge vaut environ {fr(1000 / (1 - 3 / N(f['mult'])) / 1000, 2)} Md€ d'acquisition.

## Les repères de départ

- Levier fin 2026 (définition Veolia) : {x(LEV26)} dans le scénario central — guidance « égal ou légèrement supérieur à 3x ».
- Levier fin 2027 : {x(LEV27)} (engagement ≤ 3x).
- FFO / dette ajustée (mesure des agences) : {pct(RATIO26)} fin 2026 (Moody's attend {P(f['mo26lo'])}-{P(f['mo26hi'])} %), {pct(RATIO27)} fin 2027 ; S&P abaisse la note sous {P(f['sptrig'])} %.
- Sorties nettes d'acquisitions depuis 2024 : {fr(SPENT['cum'], 2)} Md€, pour une enveloppe boosters annoncée de {P(f['boost4'])} Md€ nets.

## Ce qui fait bouger la fourchette (une hypothèse à la fois, marge sous 3x)

{rank_txt}

## Ce que les autres périmètres doivent faire de ce nombre

Toute recommandation d'acquisition se compare à la marge du scénario central et doit survivre au scénario
défavorable, ou dire explicitement quelle cession ou quelle levée de fonds propres la finance. Les hypothèses sont
discutables : elles sont en jaune dans le classeur, et la fourchette se régénère en une commande.
""", encoding="utf-8")
import json as _json
_thread = plain(page.split('<h2 id="h-thread">Le fil rouge</h2>', 1)[1].split("<div class=\"tiles\">", 1)[0])
_qr = [plain(x) for x in re.findall(r"<li>(.*?)</li>", page.split("Les quatre questions de recherche", 1)[1].split("</ol>", 1)[0], re.S)]
_blocks = [plain(x) for x in re.findall(r"<li>(.*?)</li>", page.split("Les cinq blocs du chantier", 1)[1].split("</ol>", 1)[0], re.S)]
_todo = [plain(x) for x in re.findall(r"<li>(.*?)</li>", page.split('id="h-todo"', 1)[1].split("</ul>", 1)[0], re.S)]
_range_rows = []
for line in OUT_RANGE.read_text(encoding="utf-8").splitlines():
    if line.startswith("| ") and not line.startswith("|---") and "Défavorable" not in line:
        _range_rows.append([c.strip() for c in line.strip("|").split("|")])
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
    "register": [[r["id"], r["label"], f"{r['value']} {r['unit']}".strip(), r["period"],
                  f"{Path(r['file']).name}, p. {r['page']}, {r['source_url']} (consulté le {r['date_consulted']})",
                  r.get("workstream", ""), r.get("checked_by", "")] for r in REG],
}, ensure_ascii=False, indent=1), encoding="utf-8")
OUT_STANDALONE = OUT.with_name("dossiers-sujet-2.standalone.html")
OUT_STANDALONE.write_text(standalone, encoding="utf-8")
print("écrit", OUT, len(page) // 1024, "Ko ; version autonome", OUT_STANDALONE.name)
