"""La définition du modèle de trajectoire 2026-2027 (et sa prolongation exploratoire 2028-2030) : une seule source, plusieurs moteurs.

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
    "sign26": "Cessions déjà signées au 05/11/2026, Rookery comprise (M EUR)",
    "rook": "dont déjà encaissé au 30/09/2026 : 20 % de Rookery South (M EUR)",
    "nfdq326": "Dette financière nette au 30/09/2026, Clean Earth inclus (M EUR)",
    "prog": "Programme de cessions d'ici mi-2028, au moins (Md EUR)",
    "closed": "Cessions réalisées sur neuf mois 2026 (Md EUR)",
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
    ("D26", "Cessions encaissées au T4 2026", "=({sign26}-{rook})*{s26}", "M EUR", "M", False),
    ("DE26", "EBITDA cédé en 2026 (année pleine, prudent)", "=-{D26}/{mDisp}", "M EUR", "M", False),
    ("EB26", "EBITDA 2026", "={A26}+{CE26}+{DE26}", "M EUR", "M", True),
    ("NFD26", "Dette financière nette au 31/12/2026",
     "={nfdq326}-{fcfQ4}+{tuckH2}*1000-{D26}+{other}", "M EUR", "M", True),
    ("LEV26", "Levier fin 2026", "={NFD26}/{EB26}", "x", "x", True),
    ("LEV26PF", "Levier fin 2026, Clean Earth pro forma 12 mois",
     "={NFD26}/({EB26}+{ceEb}*(12-{months})/12/{fx})", "x", "x", False),
    ("SEP", None, None, None, None, None),
    ("B27", "EBITDA 2027 avant synergies, tuck-ins et cessions 2027",
     "=({A26}+{DE26}+{ceEb}/{fx})*(1+{g27})", "M EUR", "M", False),
    ("SYN27", "Synergies Clean Earth réalisées", "={synRR}*{syn}/{fx}", "M EUR", "M", False),
    ("TE27", "EBITDA apporté par les tuck-ins 2027", "={tuck}*1000/{mTuck}", "M EUR", "M", False),
    ("D27", "Cessions encaissées en 2027",
     "=({sign26}-{rook})*(1-{s26})+MAX(0,({prog}-{closed})*1000-({sign26}-{rook}))*{s27}", "M EUR", "M", False),
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
    # Le libellé nomme la contrainte (et à qui elle appartient), pas seulement son auteur : « Veolia » n'est pas une contrainte.
    ("WHICH27", "Contrainte qui mord en premier",
     '=IF({HEADSP27}<{HEAD27},"FFO / dette ajustée ≥ 18 % (seuil des agences)","levier ≤ 3x (plafond de Veolia)")', "", None, True),
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

# ---------------------------------------------------------------- prolongation 2028-2030 (hors engagements GreenUp)
# Les mêmes grandeurs que 2027, année par année, avec les hypothèses de 2027 reconduites telles quelles :
# mêmes bornes, mêmes lois, mêmes tirages, même facteur de conjoncture. Aucune hypothèse nouvelle, aucun
# nombre nouveau : chaque ligne ne lit que des hypothèses et des constantes déjà dans le modèle. Bloc SÉPARÉ :
# il ne change aucune ligne de CALC, n'entre pas dans OUTPUTS et ne fait pas partie de la fourchette publiée.
HORIZON_YEARS = [2028, 2029, 2030]
HORIZON_LABEL = "Prolongation exploratoire, hors engagements GreenUp"
HORIZON_CAP_NOTE = ("Le plafond de levier n'est un engagement de Veolia que pour 2027 : au-delà, il est prolongé par hypothèse. "
                    "Le seuil des agences est publié pour 2026-2028 : au-delà, il est prolongé par hypothèse.")
HORIZON_READING = ("Chaque année se lit seule : la capacité de fin d'année suppose que rien d'autre que les tuck-ins de "
                   "l'hypothèse n'a été acheté avant. Les capacités de deux années ne s'additionnent pas.")

# Gabarits d'une année : @Y = l'année (deux chiffres), @P = l'année d'avant, @N = le rang de l'année depuis 2026
# (2027 = 1), @A = l'année en quatre chiffres. Écrits une fois, déroulés pour chaque année de HORIZON_YEARS.
# (clé générique, clé du moteur, libellé, gabarit, unité, format, gras)
_HORIZON_LINES = [
    ("B", "B@Y", "EBITDA @A avant synergies, tuck-ins et cessions de l'année (croissance de 2027 reconduite)",
     "=({EB@P}-{SYN@P})*(1+{g27})", "M EUR", "M", False),
    ("SYN", "SYN@Y", "Synergies Clean Earth réalisées en @A (montée linéaire de 2027 prolongée, au plus le régime de croisière)",
     "={synRR}*MIN(1,{syn}*@N)/{fx}", "M EUR", "M", False),
    ("TE", "TE@Y", "EBITDA apporté par les tuck-ins @A (rythme de 2027 reconduit)", "={tuck}*1000/{mTuck}", "M EUR", "M", False),
    ("D", "D@Y", None, None, "M EUR", "M", False),   # cessions : voir _HORIZON_DISPOSALS
    ("DE", "DE@Y", "EBITDA cédé en @A (année pleine, prudent)", "=-{D@Y}/{mDisp}", "M EUR", "M", False),
    ("EB", "EB@Y", "EBITDA @A", "={B@Y}+{SYN@Y}+{TE@Y}+{DE@Y}", "M EUR", "M", True),
    ("FCF", "FCF@Y", "Cash-flow libre net @A (conversion de 2027 reconduite)", "={conv}*{EB@Y}-{synC}/{fx}", "M EUR", "M", False),
    ("DIV", "DIV@Y", "Dividendes versés en @A (croissance de 2027 reconduite)", "={divsh26}*(1+{gDiv})^@N+{minDiv}", "M EUR", "M", False),
    ("NFD", "NFD@Y", "Dette financière nette au 31/12/@A",
     "={NFD@P}-{FCF@Y}+{DIV@Y}+{tuck}*1000-{D@Y}+{other}", "M EUR", "M", True),
    ("LEV", "LEV@Y", "Levier fin @A", "={NFD@Y}/{EB@Y}", "x", "x", True),
    ("HEAD", "HEAD@Y", "Marge sous le plafond fin @A (plafond de 2027 prolongé par hypothèse)", "={cap}*{EB@Y}-{NFD@Y}", "M EUR", "M", True),
    ("ADJ", "ADJ@Y", "Dette nette ajustée fin @A", "={NFD@Y}+({monet}-{nfd25})", "M EUR", "M", False),
    ("FFO", "FFO@Y", "FFO @A", "={ffo}*{EB@Y}", "M EUR", "M", False),
    ("RATIO", "RATIO@Y", "FFO / dette ajustée fin @A", "={FFO@Y}/{ADJ@Y}", "%", "%", False),
    ("HEADSP", "HEADSP@Y", "Marge sous le seuil S&P fin @A", "={FFO@Y}/({sptrig}/100)-{ADJ@Y}", "M EUR", "M", True),
    ("BIND", "BIND@Y", "Marge contraignante fin @A (la plus petite des deux)", "=MIN({HEAD@Y},{HEADSP@Y})", "M EUR", "M", True),
    ("WHICH", "WHICH@Y", "Contrainte qui mord en premier fin @A",
     '=IF({HEADSP@Y}<{HEAD@Y},"seuil des agences (FFO / dette ajustée)","plafond de levier (prolongé par hypothèse)")', "", None, True),
    ("MAXACQ", "MAXACQ@Y", "Acquisition maximale fin @A sous la contrainte qui mord, au multiple des tuck-ins",
     "=IF({mTuck}>{cap},MAX(0,{BIND@Y})/(1-{cap}/{mTuck}),0)", "M EUR", "M", True),
]
# Cessions : le programme publié court jusqu'à mi-2028 ; ce que 2027 n'a pas encaissé (1 − s27) l'est en 2028.
# Au-delà, rien n'est publié : aucune cession n'est supposée.
_HORIZON_DISPOSALS = {
    2028: ("Cessions encaissées en @A (solde du programme publié, à réaliser d'ici mi-@A)",
           "=MAX(0,({prog}-{closed})*1000-({sign26}-{rook}))*(1-{s27})"),
    None: ("Cessions encaissées en @A (aucune publiée au-delà du programme)", "=0"),
}


def _horizon_calc():
    out, keys = [], {}
    for year in HORIZON_YEARS:
        subst = {"@Y": f"{year % 100:02d}", "@P": f"{(year - 1) % 100:02d}", "@N": str(year - 2026), "@A": str(year)}

        def fill(text):
            for token, value in subst.items():
                text = text.replace(token, value)
            return text

        keys[str(year)] = {}
        for generic, key, label, tpl, unit, kind, bold in _HORIZON_LINES:
            if generic == "D":
                label, tpl = _HORIZON_DISPOSALS.get(year, _HORIZON_DISPOSALS[None])
            out.append((fill(key), fill(label), fill(tpl), unit, kind, bold))
            keys[str(year)][generic] = fill(key)
    return out, keys


# HORIZON_CALC : mêmes sextuplets que CALC ; HORIZON_KEYS : année -> {clé générique -> clé du moteur}.
HORIZON_CALC, HORIZON_KEYS = _horizon_calc()

# Les grandeurs que les pages lisent pour chaque année (clé générique, libellé, format).
HORIZON_METRICS = [
    ("EB", "EBITDA", "M"),
    ("FCF", "Cash-flow libre net", "M"),
    ("DIV", "Dividendes versés", "M"),
    ("NFD", "Dette financière nette en fin d'année", "M"),
    ("LEV", "Levier en fin d'année", "x"),
    ("HEAD", "Marge sous le plafond (prolongé par hypothèse)", "M"),
    ("HEADSP", "Marge sous le seuil S&P", "M"),
    ("BIND", "Marge contraignante", "M"),
    ("WHICH", "Contrainte qui mord en premier", None),
    ("MAXACQ", "Acquisition maximale sous la contrainte qui mord", "M"),
]
# Ce que le Monte-Carlo résume pour chaque année (onglet Distribution, bloc E) : mêmes tirages que 2027.
HORIZON_OUTPUTS = ["BIND", "LEV", "MAXACQ", "HEAD", "HEADSP", "EB", "FCF", "DIV", "NFD"]
# Les probabilités de chaque année (bloc F) : (clé, grandeur, comparaison, seuil en M EUR, libellé). Mêmes seuils de
# lecture que l'onglet Distribution pour 2027.
HORIZON_PROBS = [
    ("pBind", "BIND", ">", 0, "La marge contraignante fin @A est positive (plafond prolongé ET seuil S&P tenus)"),
    ("pHead", "HEAD", ">", 0, "Le levier fin @A reste sous le plafond prolongé (marge > 0)"),
    ("pSP", "HEADSP", ">", 0, "Le ratio FFO / dette reste au-dessus du seuil S&P fin @A"),
    ("pAcq1", "MAXACQ", ">=", 1000, "Veolia peut acheter au moins 1 Md€ fin @A sans franchir la contrainte qui mord"),
    ("pAcq2", "MAXACQ", ">=", 2000, "… au moins 2 Md€"),
    ("pAcq3", "MAXACQ", ">=", 3000, "… au moins 3 Md€"),
]

# Ce que la prolongation suppose, ligne par ligne : (clé, libellé, hypothèses reconduites, constantes du registre lues).
# Les références [F…] sont résolues par build_model.py (CONST_IDS, colonne Réf. d'Hypothèses) : aucune n'est tapée ici.
HORIZON_ASSUMPTIONS = [
    ("croissance", "Croissance organique de l'EBITDA : celle de 2027 reconduite chaque année (mêmes bornes, même tirage)", ["g27"], []),
    ("synergies", "Synergies Clean Earth : la montée linéaire supposée pour 2027 se poursuit jusqu'au régime de croisière publié, "
                  "sans le dépasser", ["syn", "fx"], ["synRR"]),
    ("couts", "Coûts de mise en œuvre des synergies : la même part annuelle qu'en 2027, sur les quatre années publiées", ["synC", "fx"], []),
    ("tuckins", "Tuck-ins : le rythme et le multiple de 2027 reconduits chaque année (l'enveloppe GreenUp s'arrête en 2027 : "
                "aucun engagement au-delà)", ["tuck", "mTuck"], []),
    ("cessions", "Cessions : le solde du programme publié est encaissé en 2028 (ce que 2027 n'a pas encaissé) ; aucune cession "
                 "supposée ensuite", ["s27", "mDisp"], ["prog", "closed", "sign26", "rook"]),
    ("conversion", "Cash-flow libre net : la conversion de 2027 reconduite chaque année", ["conv"], []),
    ("dividendes", "Dividendes : la croissance de 2027 reconduite chaque année ; minoritaires constants", ["gDiv", "minDiv"], ["divsh26"]),
    ("autres", "Autres flux de dette : comme en 2027", ["other"], []),
    ("plafond", "Plafond de levier : engagement de Veolia pour 2027 seulement, prolongé par hypothèse", ["cap"], []),
    ("agences", "Seuil des agences : seuil S&P publié pour 2026-2028, prolongé par hypothèse ensuite ; même écart entre dette "
                "ajustée et dette financière nette, même part de FFO dans l'EBITDA qu'en 2027", ["ffo"], ["sptrig", "monet", "nfd25"]),
    ("tirages", "Monte-Carlo : les mêmes futurs que 2027 (mêmes uniformes, mêmes lois triangulaires, même facteur de conjoncture) ; "
                "un futur garde ses hypothèses d'une année à l'autre", [], []),
    ("scenarios", "Scénarios défavorable et favorable combinés : les colonnes de 2027 (le sens de chaque borne est choisi sur la "
                  "marge fin 2027), prolongées", [], []),
]
# Ce qui est publié et daté après 2027 mais reste HORS modèle : (clé, pourquoi, lignes du registre). build_sim_data.py
# refuse une ligne absente du registre ; les nombres affichés sont ceux du registre, jamais retapés.
HORIZON_EXCLUDED = [
    ("hybrides", "Réinitialisation des hybrides après 2027 : Veolia les compte en capitaux propres ; ni rappel ni refinancement "
                 "n'est modélisé", ["F568", "F570", "F564", "F572"]),
    ("echeances", "Échéances obligataires 2028-2030 : supposées refinancées, sans effet sur la dette nette ; le coût du "
                  "refinancement n'est pas modélisé", ["F350", "F351", "F352"]),
    ("previsions_sp", "Prévisions de S&P jusqu'en 2028 : un point de comparaison, pas une entrée du modèle",
     ["F166", "F167", "F838", "F829"]),
    ("lignes", "Lignes de crédit à échéance 2030 : liquidité, sans effet sur la dette nette", ["F414", "F416"]),
    ("pfas", "Ambition 2030 dans les PFAS : un chiffre d'affaires visé, sans EBITDA ni calendrier publié", ["F772"]),
]
HORIZON_NOT_MODELLED = [
    "Aucune guidance, aucun objectif d'EBITDA ni aucun engagement de levier n'est publié au-delà de 2027.",
    "Aucune acquisition autre que les tuck-ins de l'hypothèse, aucune cession au-delà du programme publié.",
    "Change, taux d'intérêt à refinancement, rachats d'actions, inflation des contrats : non modélisés, comme en 2027.",
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
    "g26": 1, "g27": 1, "ceEb": 1, "syn": 1, "fcfQ4": 1, "conv": 1, "gDiv": 0,
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
