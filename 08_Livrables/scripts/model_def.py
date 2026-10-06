"""La définition du modèle de trajectoire 2026-2027 : une seule source, plusieurs moteurs.

Chaque ligne est une formule nommée, écrite une fois. Dans les gabarits, `{nom}` désigne
soit une autre ligne ou une hypothèse (une variable, qui change d'un scénario à l'autre),
soit une constante du registre (CONSTS : une cellule d'Entrées). Trois moteurs lisent
cette liste et doivent donner les mêmes nombres :

- build_model.py, onglet Trajectoire : une colonne par scénario ;
- build_model.py, onglet Simulation : une ligne par tirage du Monte-Carlo ;
- build_simulator.py : le moteur JavaScript du simulateur web.

Le test de parité compare les trois. Rien ici n'est un chiffre publié.
"""

# Hypothèses fixes (non tirées au sort) lues par la trajectoire, après les hypothèses variées.
FIXED_INPUTS = ["fx", "minDiv", "mTuck", "tuckH2", "other", "months", "synC", "cap"]

# Constantes : nom dans les gabarits -> ce qu'elles sont (la cellule d'Entrées est fixée par build_model.py).
CONSTS = {
    "eb25": "EBITDA 2025 du groupe (M EUR)",
    "sign26": "Signatures de cessions attendues en 2026 (Md EUR)",
    "nfdh126": "Dette financière nette au 30/06/2026, Clean Earth inclus (M EUR)",
    "prog": "Programme de cessions d'ici mi-2028, au moins (Md EUR)",
    "closed": "Cessions réalisées au S1 2026 (Md EUR)",
    "divsh26": "Dividende versé aux actionnaires en 2026 (M EUR)",
    "tgt": "Objectif GreenUp : EBITDA 2027, au moins (Md EUR)",
    "synRR": "Clean Earth : synergies de coûts en année 4 (M USD)",
    "mond25": "Moody's : dette nette ajustée 2025 (Md EUR)",
    "nfd25": "Dette financière nette au 31/12/2025 (M EUR)",
    "sptrig": "S&P : seuil de dégradation FFO / dette (%)",
    "monet": "Moody's : dette nette ajustée 2025, Exhibit 13 (M EUR)",
}

# (clé, libellé, gabarit, unité, format : "M" millions | "x" multiple | "%" | None texte, gras)
CALC = [
    ("A26", "EBITDA organique 2026 (hors Clean Earth, hors cessions)", "={eb25}*(1+{g26})", "M EUR", "M", False),
    ("CE26", "Clean Earth en 2026 (mois consolidés)", "={ceEb}*{months}/12/{fx}", "M EUR", "M", False),
    ("D26", "Cessions encaissées au S2 2026", "={sign26}*{s26}*1000", "M EUR", "M", False),
    ("DE26", "EBITDA cédé en 2026 (année pleine, prudent)", "=-{D26}/{mDisp}", "M EUR", "M", False),
    ("EB26", "EBITDA 2026", "={A26}+{CE26}+{DE26}", "M EUR", "M", True),
    ("NFD26", "Dette financière nette au 31/12/2026",
     "={nfdh126}-{fcfH2}+{tuckH2}*1000-{D26}+{other}", "M EUR", "M", True),
    ("LEV26", "Levier fin 2026", "={NFD26}/{EB26}", "x", "x", True),
    ("LEV26PF", "Levier fin 2026, Clean Earth pro forma 12 mois",
     "={NFD26}/({EB26}+{ceEb}*(12-{months})/12/{fx})", "x", "x", False),
    ("SEP", None, None, None, None, None),
    ("B27", "EBITDA 2027 avant synergies, tuck-ins et cessions 2027",
     "=({A26}+{DE26}+{ceEb}/{fx})*(1+{g27})", "M EUR", "M", False),
    ("SYN27", "Synergies Clean Earth réalisées", "={synRR}*{syn}/{fx}", "M EUR", "M", False),
    ("TE27", "EBITDA apporté par les tuck-ins 2027", "={tuck}*1000/{mTuck}", "M EUR", "M", False),
    ("D27", "Cessions encaissées en 2027",
     "=({sign26}*(1-{s26})+MAX(0,{prog}-{closed}-{sign26})*{s27})*1000", "M EUR", "M", False),
    ("DE27", "EBITDA cédé en 2027 (année pleine, prudent)", "=-{D27}/{mDisp}", "M EUR", "M", False),
    ("EB27", "EBITDA 2027", "={B27}+{SYN27}+{TE27}+{DE27}", "M EUR", "M", True),
    ("FCF27", "Cash-flow libre net 2027", "={conv}*{EB27}-{synC}/{fx}", "M EUR", "M", False),
    ("DIV27", "Dividendes versés en 2027", "={divsh26}*(1+{gDiv})+{minDiv}", "M EUR", "M", False),
    ("NFD27", "Dette financière nette au 31/12/2027",
     "={NFD26}-{FCF27}+{DIV27}+{tuck}*1000-{D27}+{other}", "M EUR", "M", True),
    ("LEV27", "Levier fin 2027", "={NFD27}/{EB27}", "x", "x", True),
    ("HEAD27", "Marge de manœuvre fin 2027 (plafond × EBITDA − DFN)", "={cap}*{EB27}-{NFD27}", "M EUR", "M",
     True),
    ("MAXACQ", "Acquisition maximale au multiple des tuck-ins",
     "=IF({mTuck}>{cap},MAX(0,{HEAD27})/(1-{cap}/{mTuck}),0)", "M EUR", "M", True),
    ("GAP8", "EBITDA 2027 moins l'objectif ≥ 8 Md€", "={EB27}-{tgt}*1000", "M EUR", "M", False),
    ("SEP2", None, None, None, None, None),
    ("ADJ26", "Dette nette ajustée par les agences fin 2026 (DFN + écart Moody's 2025)", "={NFD26}+({monet}-{nfd25})", "M EUR", "M", False),
    ("FFO26", "FFO 2026 (hypothèse ffo × EBITDA 2026)", "={ffo}*{EB26}", "M EUR", "M", False),
    ("RATIO26", "FFO / dette ajustée fin 2026", "={FFO26}/{ADJ26}", "%", "%", True),
    ("HEADSP26", "Marge de dette fin 2026 sous le seuil S&P (FFO / 18 % − dette ajustée)", "={FFO26}/({sptrig}/100)-{ADJ26}", "M EUR", "M", True),
    ("ADJ27", "Dette nette ajustée fin 2027", "={NFD27}+({monet}-{nfd25})", "M EUR", "M", False),
    ("FFO27", "FFO 2027", "={ffo}*{EB27}", "M EUR", "M", False),
    ("RATIO27", "FFO / dette ajustée fin 2027", "={FFO27}/{ADJ27}", "%", "%", True),
    ("HEADSP27", "Marge de dette fin 2027 sous le seuil S&P 18 %", "={FFO27}/({sptrig}/100)-{ADJ27}", "M EUR", "M", True),
    ("BIND27", "Marge contraignante fin 2027 (la plus petite des deux)", "=MIN({HEAD27},{HEADSP27})", "M EUR", "M", True),
    ("WHICH27", "Contrainte qui mord en premier", '=IF({HEADSP27}<{HEAD27},"agences : FFO / dette ≥ 18 %","Veolia : levier ≤ 3x")', "", None, True),
    ("MAXACQB", "Acquisition maximale fin 2027 sous la contrainte qui mord, au multiple des tuck-ins",
     "=IF({mTuck}>{cap},MAX(0,{BIND27})/(1-{cap}/{mTuck}),0)", "M EUR", "M", True),
]

# Ce que le Monte-Carlo résume (onglet Distribution) et ce que le simulateur affiche.
OUTPUTS = [
    ("LEV26", "Levier fin 2026", "x"),
    ("LEV27", "Levier fin 2027", "x"),
    ("HEAD27", "Marge sous 3x fin 2027", "M"),
    ("RATIO26", "FFO / dette ajustée fin 2026", "%"),
    ("HEADSP26", "Marge sous le seuil S&P fin 2026", "M"),
    ("HEADSP27", "Marge sous le seuil S&P fin 2027", "M"),
    ("BIND27", "Marge contraignante fin 2027", "M"),
    ("MAXACQB", "Acquisition maximale fin 2027 sous la contrainte qui mord", "M"),
    ("EB27", "EBITDA 2027", "M"),
    ("NFD27", "Dette financière nette fin 2027", "M"),
]

# Monte-Carlo : graine des tirages (reproductibles) et nombre de tirages.
SEED = 20261016
N_DRAWS = 1000

# Un facteur « conjoncture » commun (copule gaussienne à un facteur) : dans une récession, la croissance
# ralentit, le cash-flow se tasse, les cessions glissent et se vendent moins cher, les cibles aussi
# deviennent moins chères. Exposition 1 = l'hypothèse baisse quand la conjoncture baisse ; 0 = elle n'en
# dépend pas (politique de dividende, enveloppe de tuck-ins décidée). Le poids est la corrélation entre
# deux hypothèses exposées : hypothèse du groupe, cellule jaune de l'onglet Simulation (0 = indépendance).
FACTOR_WEIGHT = 0.4
FACTOR_LOADINGS = {
    "g26": 1, "g27": 1, "ceEb": 1, "syn": 1, "fcfH2": 1, "conv": 1, "gDiv": 0,
    "s26": 1, "s27": 1, "mDisp": 1, "tuck": 0, "ffo": 1, "mTuck": 1,
}
CORRELATIONS: list = []   # remplacé par le facteur commun ; gardé pour la compatibilité des exports


def variables(template: str) -> list[str]:
    """Les noms `{x}` d'un gabarit, dans l'ordre, sans doublon."""
    import re

    seen: list[str] = []
    for name in re.findall(r"{(\w+)}", template or ""):
        if name not in seen:
            seen.append(name)
    return seen
