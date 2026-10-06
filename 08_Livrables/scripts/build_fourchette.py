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
}


def signature(s):
    """Ce qui fait une nouvelle version : les nombres publiés, arrondis à ce qui se lit."""
    def r(v):
        return round(v, 0) if isinstance(v, float) and abs(v) >= 10 else round(v, 3) if isinstance(v, float) else v
    return json.dumps({k: {kk: r(vv) for kk, vv in s[k].items()} for k in ("central", "unfavourable", "favourable")}
                      | {"p": {k: r(v) for k, v in s["probabilities"].items()},
                         "d": {k: {kk: r(vv) for kk, vv in v.items()} for k, v in s["distribution"].items()}}, sort_keys=True)


history = json.loads(HIST.read_text(encoding="utf-8")) if HIST.exists() else {"versions": []}
versions = history.get("versions", [])
if versions and signature(versions[-1]) == signature(snapshot):
    print(f"fourchette inchangée (version {len(versions)} du {versions[-1]['date']})")
    sys.exit(0)
snapshot["version"] = len(versions) + 1
versions.append(snapshot)
history["versions"] = versions
HIST.write_text(json.dumps(history, ensure_ascii=False, indent=1), encoding="utf-8")
print(f"fourchette : version {snapshot['version']} publiée ({snapshot['date']})")
