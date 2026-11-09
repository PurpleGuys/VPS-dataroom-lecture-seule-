"""La fourchette de référence pour la promo : un instantané daté, ajouté à l'historique quand les nombres changent.

Lit simulateur-data.json (valeurs calculées par LibreOffice dans le classeur), n'en calcule
aucune, et tient 08_Livrables/fourchette-historique.json : la liste des versions publiées,
la plus récente en dernier. Une nouvelle version n'est ajoutée que si un nombre publié change.

Usage : python build_fourchette.py [simulateur-data.json] [fourchette-historique.json]
"""
import json
import sys
from datetime import date
from pathlib import Path

HERE = Path(__file__).parent
DATA = Path(sys.argv[1]) if len(sys.argv) > 1 else HERE.parent / "simulateur-data.json"
HIST = Path(sys.argv[2]) if len(sys.argv) > 2 else HERE.parent / "fourchette-historique.json"
KEYS = ["MAXACQB", "BIND27", "HEAD27", "HEADSP27", "LEV26", "LEV27", "EB27", "NFD27", "RATIO26", "WHICH27"]
# Ce sur quoi repose la version publiée : la date d'arrêté des données du modèle. À changer avec le point de départ.
DATA_CUTOFF = ("données arrêtées au 30/09/2026 (rapport du T3 2026 de Veolia, publié le 05/11/2026) : dette de départ au 30/09, "
               "cash-flow du T4 2025 reproduit, 570 M€ de cessions signées à 10,4x, croissance organique publiée à fin septembre")

data = json.loads(DATA.read_text(encoding="utf-8"))
cols = {c["name"]: c for c in data["columns"]}


def pick(name):
    col = cols[name]
    return {k: col["outputs"].get(k) for k in KEYS} | {"mTuck": col["inputs"].get("mTuck"), "cap": col["inputs"].get("cap")}


snapshot = {
    "date": date.today().isoformat(),
    "generated": data["generated"],
    "model": data["model"],
    "central": pick("Scénario actif"),
    "unfavourable": pick("Défavorable combiné"),
    "favourable": pick("Favorable combiné"),
    "distribution": {k: data["distribution"][k] for k in ("MAXACQB", "BIND27", "LEV26", "LEV27") if k in data["distribution"]},
    "probabilities": data["probabilities"],
    "factor_weight": data.get("factor_weight"),
    "note": DATA_CUTOFF,
}


def binding(label):
    """La contrainte qui mord, quel que soit son libellé dans le classeur : « 3x » (le plafond de levier de Veolia)
    ou « agences » (le seuil FFO / dette ajustée de 18 %). Anciens libellés : « Veolia : levier ≤ 3x », « agences : … »."""
    if not isinstance(label, str):
        return label
    return "agences" if ("18 %" in label or "agences" in label.casefold() or "FFO" in label) else "3x"


def signature(s):
    """Ce qui fait une nouvelle version : les nombres publiés, arrondis à ce qui se lit, et la contrainte qui mord
    (laquelle, pas son libellé : reformuler un libellé ne publie pas une nouvelle version)."""
    def r(v):
        return round(v, 0) if isinstance(v, float) and abs(v) >= 10 else round(v, 3) if isinstance(v, float) else v
    return json.dumps({k: {kk: (binding(vv) if kk == "WHICH27" else r(vv)) for kk, vv in s[k].items()}
                       for k in ("central", "unfavourable", "favourable")}
                      | {"p": {k: r(v) for k, v in s["probabilities"].items()},
                         "d": {k: {kk: r(vv) for kk, vv in v.items()} for k, v in s["distribution"].items()}}, sort_keys=True)


# Le libellé de la contrainte suit celui du classeur dans toutes les versions publiées : c'est un libellé, pas un
# nombre publié (les nombres d'une version ne changent jamais ; la page de la fourchette affiche celui-ci).
LABELS = {binding(c["outputs"].get("WHICH27")): c["outputs"]["WHICH27"] for c in data["columns"] if c["outputs"].get("WHICH27")}


def relabel(version):
    changed = 0
    for k in ("central", "unfavourable", "favourable"):
        old = (version.get(k) or {}).get("WHICH27")
        new = LABELS.get(binding(old))
        if old and new and new != old:
            version[k]["WHICH27"] = new
            changed = 1
    return changed


history = json.loads(HIST.read_text(encoding="utf-8")) if HIST.exists() else {"versions": []}
versions = history.get("versions", [])
relabeled = sum(relabel(v) for v in versions)
if versions and signature(versions[-1]) == signature(snapshot):
    if relabeled:
        HIST.write_text(json.dumps(history, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"fourchette inchangée (version {len(versions)} du {versions[-1]['date']})"
          + (f" ; libellé de la contrainte qui mord repris du classeur dans {relabeled} version(s)" if relabeled else ""))
    sys.exit(0)
snapshot["version"] = len(versions) + 1
versions.append(snapshot)
history["versions"] = versions
HIST.write_text(json.dumps(history, ensure_ascii=False, indent=1), encoding="utf-8")
print(f"fourchette : version {snapshot['version']} publiée ({snapshot['date']})")
