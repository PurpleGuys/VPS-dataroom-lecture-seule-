"""Fiches d'oral : une fiche par partie du rapport (les six rôles, puis la recommandation), dans l'ordre du rapport.

Rien n'est écrit ici que le rapport ne dise déjà. Chaque fiche se DÉRIVE des textes de l'édition courte
(08_Livrables/texte/*.md) et des registres, par les fonctions mêmes du rapport (dataroom.report) :
  - la thèse            = le titre de la partie (son « # ») ;
  - les messages        = ses intertitres (« ## »), cinq au plus : l'ouverture, puis ceux qui reprennent le plus de
                          chiffres de l'ouverture, dans l'ordre du texte ;
  - les chiffres clés   = les jetons [F…], [D…], [T…], [M:…] que cite la section d'ouverture, avec leur source ;
  - ce que nous n'avons pas pu établir = la liste de la section du même nom ;
  - les questions du jury = chaque message retourné en question, et pour réponse le premier paragraphe de sa section.
Si le texte du rapport change, la fiche suit.

Le contrôle de cohérence : une partie absente des textes, une partie sans la structure attendue, un jeton que ni les
registres ni le classeur ne connaissent. Par défaut l'anomalie se MARQUE — « [F9999 ? jeton inconnu] » à sa place, un
bandeau en tête de page qui les compte, un message sur la sortie d'erreur — et tout s'écrit quand même, code 0 : les
textes se modifient dans l'éditeur web, une faute de frappe ne doit arrêter ni le classeur ni le rapport. Le contrôle dur
est une option, « --strict » ou DATAROOM_FICHES_STRICT=1 : code 1, message qui nomme le jeton et le fichier, rien d'écrit.

build_dossiers.py importe ce module (corpus, dérivation, gabarit et feuille de style) : une seule dérivation pour les
deux pages. Usage : python build_fiches.py [--strict] [report-data.json] [fiches-oral.html] — le premier argument, que la
chaîne passe toujours, n'est plus lu : les fiches ne dépendent plus des dossiers.
"""
import datetime as _dt
import html
import os
import re
import sys
from pathlib import Path
from urllib.parse import quote

HERE = Path(__file__).parent
LIV = Path(os.environ.get("DATAROOM_LIVRABLES") or HERE.parent)
ADMIN = Path(os.environ.get("DATAROOM_ADMIN") or (LIV.parent / "00_Admin"))
DOCS = os.environ.get("DATAROOM_DOCS_BASE") or "/documents/"      # la route des documents, comme les autres pages
CLASSEUR = "modele-greenup-2027.xlsx"
PARTIES = ("role-1", "role-2", "role-3", "role-4", "role-5", "role-6", "recommandation")   # l'ordre du rapport
NON_ETABLI = "ce que nous n'avons pas pu établir"
MESSAGES_MIN, MESSAGES_MAX = 3, 5      # « trois à cinq messages à tenir »
CHIFFRES_MAX = 8                       # une fiche tient sur une page : les chiffres de l'ouverture que la partie reprend le plus
STRICT = "--strict" in sys.argv[1:] or (os.environ.get("DATAROOM_FICHES_STRICT") or "").strip().lower() in ("1", "true", "oui", "yes")
esc = html.escape


def echec(message):
    """Un arrêt net : un message lisible, et un code de sortie non nul pour la chaîne."""
    raise SystemExit(f"build_fiches : {message}")


try:
    from dataroom import report as R
    from dataroom.vault import read_deals, read_register, read_reviews, read_targets
except ImportError as exc:   # sans le paquet du serveur, pas de résolution des jetons : mieux vaut ne rien écrire
    echec(f"le paquet « dataroom » est introuvable ({exc}) ; lancer par « dataroom livrables » ou avec PYTHONPATH=…/src")


# ------------------------------------------------------------------ le corpus : textes, registres, classeur
class Corpus:
    def __init__(self, liv=LIV, admin=ADMIN):
        self.liv, self.vault = Path(liv), Path(admin).parent
        self.textes = R.read_texts(self.vault)
        self.modele = R.load_model(self.liv)
        self.registre = read_register(self.vault)
        ids = [t for parsed in self.textes.values() for t in parsed.tokens]   # les parties, et la conclusion pour les dossiers
        self.res = R.resolve_tokens(ids, register=self.registre, deals=read_deals(self.vault), targets=read_targets(self.vault),
                                    reviews=read_reviews(self.vault), model=self.modele)
        self.refs = R.short_refs(self.vault, self.res)                     # « Auteur (année), p. N », comme la bibliographie
        self.periodes = {(r.get("id") or "").strip(): (r.get("period") or "").strip() for r in self.registre}
        self.version = (self.modele or {}).get("version", "")              # « fourchette version 2 du 9 novembre 2026 »
        self.anomalies = []                                                # (fichier, ce qui ne va pas) : le bandeau, la sortie d'erreur

    def connu(self, jeton):
        r = self.res.get(jeton)
        return r is not None and r.status == "ok"

    # -------------------------------------------------------------- un jeton : sa valeur, son libellé, sa source
    def valeur(self, jeton):
        if not self.connu(jeton):
            return f"[{jeton} ? jeton inconnu]"
        r = self.res[jeton]
        return r.display or r.value

    def libelle(self, jeton):
        if not self.connu(jeton):
            return "Jeton inconnu des registres et du classeur"
        r = self.res[jeton]
        periode = self.periodes.get(jeton, "")
        return R.units_in_label(r.label) + (f" ({periode})" if periode else "")

    def reference(self, jeton):
        """La source en clair : « Veolia Environnement SA (2026c), p. 5 », ou l'onglet du classeur pour un calcul."""
        if not self.connu(jeton):
            return "aucune ligne au registre ni au classeur"
        r = self.res[jeton]
        if r.kind == "model":
            ou = r.citation.split(" — ", 1)[-1].rstrip(".")
            return ou[:1].upper() + ou[1:]
        ref = self.refs.get(jeton) or self.refs.get(jeton.split("-")[0])
        return ref or ((Path(r.file).name + R.page_ref(r.file, r.page)) if r.file else (r.url or "source au registre"))

    def lien(self, jeton):
        """Où la preuve s'ouvre : la page du document versé, à défaut son adresse ; le classeur pour un calcul."""
        if not self.connu(jeton):
            return ""
        r = self.res[jeton]
        if r.kind == "model":
            return DOCS + "08_Livrables/" + CLASSEUR
        if r.file:
            ancre = f"#page={r.page}" if r.page.isdigit() and r.file.lower().endswith(".pdf") else ""
            return DOCS + quote(r.file) + ancre
        return r.url

    def etat(self, jeton):
        if not self.connu(jeton):
            return "à corriger dans le texte"
        r = self.res[jeton]
        return r.state or "non relu"

    # -------------------------------------------------------------- le texte du rapport, rendu
    def html(self, inlines):
        out = []
        for i in inlines:
            if i.kind == "token" and not self.connu(i.id):
                s = f'<span class="alerte" title="{esc(self.libelle(i.id))}">{esc(self.valeur(i.id))}</span>'
            elif i.kind == "token":
                href, titre = self.lien(i.id), f"{i.id} — {self.libelle(i.id)} — {self.reference(i.id)}"
                s = (f'<a class="ch" href="{esc(href)}" title="{esc(titre)}">{esc(self.valeur(i.id))}</a>' if href
                     else f'<span class="ch" title="{esc(titre)}">{esc(self.valeur(i.id))}</span>')
            elif i.kind == "link":
                s = f'<a href="{esc(i.url)}">{esc(i.text)}</a>'
            elif i.kind == "code":
                s = f"<code>{esc(i.text)}</code>"
            else:
                s = esc(i.text)
            if i.bold:
                s = f"<strong>{s}</strong>"
            if i.italic:
                s = f"<em>{s}</em>"
            out.append(s)
        return "".join(out).strip()

    def texte(self, inlines):
        """Le même passage en texte nu (report-data.json, attributs)."""
        return "".join(self.valeur(i.id) if i.kind == "token" else i.text for i in inlines).strip()

    def bloc(self, b):
        if b.kind == "paragraph":
            return f"<p>{self.html(b.inlines)}</p>"
        if b.kind == "quote":
            return f"<blockquote>{self.html(b.inlines)}</blockquote>"
        if b.kind == "list":
            tag = "ol" if b.ordered else "ul"
            return f"<{tag}>" + "".join(f"<li>{self.html(item)}</li>" for item in b.items) + f"</{tag}>"
        if b.kind == "table" and b.rows:
            tete = "".join(f"<th>{self.html(c)}</th>" for c in b.rows[0])
            corps = "".join("<tr>" + "".join(f"<td>{self.html(c)}</td>" for c in row) + "</tr>" for row in b.rows[1:])
            return f'<div class="tbl"><table class="texte"><thead><tr>{tete}</tr></thead><tbody>{corps}</tbody></table></div>'
        if b.kind == "heading":
            return f"<h4>{self.html(b.inlines)}</h4>"
        return ""


def _jetons(blocs):
    """Les jetons que citent ces blocs (paragraphes, listes, tableaux), dans l'ordre de lecture, sans répétition."""
    vus = []
    for b in blocs:
        morceaux = [b.inlines, *b.items, *(cell for row in b.rows for cell in row)]
        for i in (i for morceau in morceaux for i in morceau):
            if i.kind == "token" and i.id not in vus:
                vus.append(i.id)
    return vus


# ------------------------------------------------------------------ la dérivation, et son contrôle
def resume(corpus):
    """Les anomalies en une ligne, pour la sortie d'erreur : chacune nomme son fichier."""
    n = len(corpus.anomalies)
    return (f"{n} anomalie{'s' if n > 1 else ''} dans les textes — " + " ; ".join(f"{f} : {m}" for f, m in corpus.anomalies))


def deriver(corpus, strict=None):
    """Une entrée par partie du rapport. Tout y vient du texte. Ce qui manque (partie absente, titre, intertitres, section
    des limites, jeton inconnu) est noté dans corpus.anomalies et se marque dans la page ; en mode strict, le script s'arrête."""
    strict = STRICT if strict is None else strict
    parties, noter = [], corpus.anomalies.append
    for rang, cle in enumerate(PARTIES, 1):
        parsed = corpus.textes.get(cle)
        fichier = f"08_Livrables/texte/{cle}.md"
        manque = "" if parsed is not None else f"la partie « {cle} » est absente des textes ({fichier} manque)"
        if manque:
            noter((fichier, manque + " : pas de fiche sans texte."))
        blocs = parsed.blocks if parsed is not None else []
        these, corps = [], blocs
        if blocs and blocs[0].kind == "heading" and blocs[0].level == 1:
            these, corps = [i for i in blocs[0].inlines], blocs[1:]
            if these and these[0].kind == "text":   # « Rôle 3 — » est le gabarit, pas la thèse
                these[0] = type(these[0])("text", re.sub(r"^Rôle\s*\d\s*[—–-]\s*", "", these[0].text), bold=these[0].bold, italic=these[0].italic)
        elif not manque:
            noter((fichier, "le texte ne commence pas par son titre (« # … ») : la thèse de la fiche en vient."))
        sections, courant = [], None
        for b in corps:
            if b.kind == "heading" and b.level == 2:
                courant = {"titre": b.inlines, "blocs": [], "sous": []}
                sections.append(courant)
            elif courant is None:
                continue
            else:
                if b.kind == "heading":
                    courant["sous"].append(b.inlines)
                courant["blocs"].append(b)
        limites = [s for s in sections if corpus.texte(s["titre"]).casefold().startswith(NON_ETABLI)]
        sections = [s for s in sections if s not in limites]
        sans_limites = not manque and (not limites or not any(b.kind == "list" for b in limites[0]["blocs"]))
        if not manque and len(sections) < MESSAGES_MIN:
            noter((fichier, f"{len(sections)} intertitre(s) « ## » seulement : une fiche tient {MESSAGES_MIN} à {MESSAGES_MAX} messages."))
        if sans_limites:
            noter((fichier, "pas de section « ## Ce que nous n'avons pas pu établir » avec sa liste : la fiche la reprend."))
        for s in sections:
            s["jetons"] = _jetons(s["blocs"])
            s["premier"] = next((b for b in s["blocs"] if b.kind == "paragraph"), None)
        ouverture = sections[0] if sections else {"titre": [], "blocs": [], "sous": [], "jetons": [], "premier": None}
        # Les messages : l'ouverture, puis les sections qui portent le plus de ses chiffres ; l'ordre reste celui du texte.
        if len(sections) <= MESSAGES_MAX:
            retenus = list(range(len(sections)))
        else:
            commun = lambda k: len(set(ouverture["jetons"]) & set(sections[k]["jetons"]))  # noqa: E731
            retenus = [0] + sorted(sorted(range(1, len(sections)), key=lambda k: (-commun(k), k))[:MESSAGES_MAX - 1])
        tous = _jetons(blocs)
        # Les chiffres clés : ceux de l'ouverture ; s'il y en a trop pour une fiche, ceux que le plus de sections reprennent.
        reprises = lambda t: sum(1 for s in sections if t in s["jetons"])  # noqa: E731
        candidats = ouverture["jetons"] or tous
        garde = set(sorted(candidats, key=lambda t: (-reprises(t), candidats.index(t)))[:CHIFFRES_MAX])
        cles = [t for t in candidats if t in garde]
        for t in tous:
            if not corpus.connu(t):
                noter((fichier, f"jeton inconnu des registres et du classeur : [{t}]"))
        numero = R.TEXT_FILES.get(cle) or 0
        parties.append({
            "cle": cle, "rang": rang, "role": numero,
            "nom": R.ROLE_NAMES.get(numero, "Recommandation"),
            "etiquette": f"Rôle {numero}" if numero else "Synthèse",
            "these": these, "sections": sections, "ouverture": ouverture,
            "messages": [sections[k] for k in retenus],
            "autres": [s for k, s in enumerate(sections) if k not in retenus],
            "cles": cles, "jetons": tous,
            "limites": [item for b in (limites[0]["blocs"] if limites else []) if b.kind == "list" for item in b.items],
            "signe": str(parsed.meta.get("signe_par") or "").strip() if parsed is not None else "",
            "relu": str(parsed.meta.get("relu_par") or "").strip() if parsed is not None else "",
            "manque": manque, "sans_limites": sans_limites,
        })
    if strict and corpus.anomalies:
        echec(resume(corpus) + ". Mode strict : corriger le texte ou verser la ligne au registre ; rien n'est écrit.")
    return parties


def signaler(corpus, script="build_fiches"):
    """Hors mode strict : les anomalies sur la sortie d'erreur, pour les journaux de la chaîne ; le code reste 0."""
    if corpus.anomalies:
        print(f"{script} : {resume(corpus)}. Les pages sont écrites, chaque anomalie y est marquée (contrôle dur : --strict).", file=sys.stderr)


def alertes_html(corpus):
    """Le bandeau de tête : combien d'anomalies, lesquelles, dans quel fichier. Rien quand tout va bien."""
    if not corpus.anomalies:
        return ""
    n = len(corpus.anomalies)
    return (f'<aside class="f-alerte" role="alert"><p class="anno">{n} anomalie{"s" if n > 1 else ""} dans les textes du rapport</p><ul>'
            + "".join(f"<li><code>{esc(f)}</code> : {esc(m)}</li>" for f, m in corpus.anomalies) + "</ul></aside>")


def these_html(corpus, p):
    if p["these"]:
        return corpus.html(p["these"])
    return f'<span class="alerte">[{"partie absente des textes" if p["manque"] else "titre « # » manquant"}]</span>'


# ------------------------------------------------------------------ les pièces d'une fiche (les dossiers les reprennent)
def question(corpus, section):
    return f"Sur quoi vous appuyez-vous pour dire : « {corpus.texte(section['titre'])} » ?"


def tete_html(corpus, p, prefixe="fiche"):
    signe = " · ".join(x for x in (f"signé par {p['signe']}" if p["signe"] else "", f"relu par {p['relu']}" if p["relu"] else "") if x)
    return (f'<header class="f-tete"><span class="f-n" aria-hidden="true">{p["rang"]}</span><div>'
            f'<p class="anno">{esc(p["etiquette"])} · {esc(p["nom"])}{" · " + esc(signe) if signe else ""}</p>'
            f'<h2 id="{prefixe}-{p["rang"]}-t">{these_html(corpus, p)}</h2></div></header>')


def chiffres_html(corpus, jetons, ident=False):
    """Le tableau des chiffres : le libellé et, dessous, la source cliquable et l'état de relecture ; la valeur à droite."""
    lignes = []
    for t in jetons:
        r = corpus.res[t]
        href, ref = corpus.lien(t), corpus.reference(t)
        source = f'<a href="{esc(href)}">{esc(ref)}</a>' if href else esc(ref)
        idc = f'<span class="id">{esc(t)}</span> ' if (ident or r.kind != "model") else ""
        etat = "" if r.kind == "model" else f' <span class="etat{" ok" if r.checked else ""}">{esc(corpus.etat(t))}</span>'
        classe = "alerte" if not corpus.connu(t) else ("calc" if r.kind == "model" else "")
        lignes.append(f'<tr{" class=\"" + classe + "\"" if classe else ""}><td class="lib"><span class="l">{esc(corpus.libelle(t))}</span>'
                      f'<span class="src">{idc}{source}{etat}</span></td><td class="num">{esc(corpus.valeur(t))}</td></tr>')
    return ('<div class="tbl"><table class="chiffres"><thead><tr><th>Chiffre et source</th><th class="num">Valeur</th></tr></thead>'
            f'<tbody>{"".join(lignes)}</tbody></table></div>')


def limites_html(corpus, p):
    absente = '<li class="alerte">[section « Ce que nous n\'avons pas pu établir » absente du texte]</li>' if p["sans_limites"] else ""
    return '<ul class="limites">' + "".join(f"<li>{corpus.html(item)}</li>" for item in p["limites"]) + absente + "</ul>"


def absente_html(corpus, p, prefixe="fiche"):
    """La place d'une partie que les textes n'ont pas : son numéro, son nom, et ce qui manque."""
    return (f'\n<article class="fiche" id="{prefixe}-{p["rang"]}" aria-labelledby="{prefixe}-{p["rang"]}-t">\n  {tete_html(corpus, p, prefixe)}\n'
            f'  <p class="alerte">{esc(p["manque"][:1].upper() + p["manque"][1:])}.</p>\n</article>')


def questions_html(corpus, sections, entier=False):
    """Une question par section ; la réponse est le texte du rapport : son premier paragraphe (fiche) ou la section (dossier)."""
    out = []
    for s in sections:
        if s["premier"] is None:
            continue
        reponse = "".join(corpus.bloc(b) for b in s["blocs"]) if entier else corpus.bloc(s["premier"])
        out.append(f'<details class="qa"><summary>{esc(question(corpus, s))}</summary><div class="rep">{reponse}</div></details>')
    return "".join(out)


def fiche_html(corpus, p):
    if p["manque"]:
        return absente_html(corpus, p)
    messages = "".join(f"<li>{corpus.html(s['titre'])}</li>" for s in p["messages"])
    autres = ("<p class=\"autres\"><span class=\"anno\">Aussi dans la partie</span> "
              + " · ".join(corpus.html(s["titre"]) for s in p["autres"]) + "</p>") if p["autres"] else ""
    return f"""
<article class="fiche" id="fiche-{p['rang']}" aria-labelledby="fiche-{p['rang']}-t">
  {tete_html(corpus, p)}
  <div class="f-grille">
    <section class="bloc b-msg"><h3>Les messages à tenir</h3><ol class="messages">{messages}</ol>{autres}</section>
    <section class="bloc b-chf"><h3>Les chiffres clés</h3>{chiffres_html(corpus, p['cles'])}</section>
    <section class="bloc b-lim"><h3>Ce que nous n'avons pas pu établir</h3>{limites_html(corpus, p)}</section>
    <section class="bloc b-qa"><h3>Questions probables du jury</h3>{questions_html(corpus, p['messages'])}</section>
  </div>
</article>"""


def fil_html(corpus, parties, prefixe="fiche"):
    """Le fil de l'oral : les thèses des parties, l'une après l'autre. C'est aussi le sommaire."""
    return ('<ol class="fil">' + "".join(
        f'<li><a href="#{prefixe}-{p["rang"]}"><span class="fil-n">{p["rang"]}</span><span class="fil-r">{esc(p["etiquette"])} · {esc(p["nom"])}</span>'
        f'<span class="fil-t">{esc(corpus.texte(p["these"]) or ("[partie absente des textes]" if p["manque"] else "[titre manquant]"))}</span></a></li>'
        for p in parties) + "</ol>")


def nav_html(parties, prefixe="fiche"):
    return ('<nav class="f-nav" aria-label="Parties"><div class="f-nav-in">' + "".join(
        f'<a href="#{prefixe}-{p["rang"]}" title="{esc(p["nom"])}"><span>{p["rang"]}</span></a>' for p in parties)
        + '<button type="button" id="f-ouvrir" aria-pressed="false">Tout ouvrir</button>'
          '<button type="button" id="f-imprimer">Imprimer</button></div></nav>')


def provenance(corpus):
    """D'où vient la page, daté : les textes, le registre, la fourchette publiée. Aucun nombre tapé."""
    jour = _dt.date.today().strftime("%d/%m/%Y")
    return (f"Généré le {jour} depuis les textes du rapport (08_Livrables/texte, édition courte), les registres "
            f"(00_Admin : {len(corpus.registre)} chiffres) et le classeur du groupe" + (f" — {corpus.version}" if corpus.version else "") + ".")


# Les couleurs, les polices et le pas sont ceux d'is42 (src/dataroom/templates/is42/is42.css) : encre sur papier, filets,
# annotations en Plex Mono capitales, module de 42 px. La feuille sert aux fiches et aux dossiers.
CSS = """
:root, :root[data-theme] { --f-mesure: 72ch; }
body { margin: 0; background: var(--paper, #eeebe4); color: var(--ink, #151512); font: 15px/1.55 var(--sans, system-ui, sans-serif); }
.wrap { max-width: var(--maxw, 1240px); margin: 0 auto; padding: var(--s8, 33px) var(--pad, 20px) var(--s12, 86px); }
.anno, .fiche h3, .f-couv h2, th { font: 400 .72rem/1.4 var(--mono, ui-monospace, monospace); letter-spacing: .06em; text-transform: uppercase; color: var(--faint, #66635a); margin: 0; }
h1 { font: 800 clamp(2rem, 5vw, 3.3rem)/.92 var(--sans, system-ui, sans-serif); font-stretch: 112%; letter-spacing: -.04em; text-transform: uppercase; margin: var(--s3, 10px) 0 var(--s5, 16px); text-wrap: balance; }
.lede { max-width: var(--f-mesure); margin: 0; font-size: 1.02rem; }
.prov { max-width: var(--f-mesure); margin: var(--s4, 13px) 0 0; font: 400 .74rem/1.5 var(--mono, ui-monospace, monospace); color: var(--mut, #55524c); }
.f-couv { padding-bottom: var(--s8, 33px); }
.f-couv h2 { margin: var(--s8, 33px) 0 var(--s3, 10px); }
ol.fil { list-style: none; margin: 0; padding: 0; border-top: 1px solid var(--ink, #151512); }
ol.fil a { display: grid; grid-template-columns: var(--s9, 42px) minmax(0, 15rem) minmax(0, 1fr); gap: 0 var(--s5, 16px); align-items: baseline; padding: var(--s3, 10px) 0;
           border-bottom: 1px solid var(--hair, #15151224); text-decoration: none; }
ol.fil a:hover .fil-t, ol.fil a:focus-visible .fil-t { text-decoration: underline; }
.fil-n { font: 500 .95rem/1 var(--mono, ui-monospace, monospace); }
.fil-r { font: 400 .72rem/1.4 var(--mono, ui-monospace, monospace); letter-spacing: .06em; text-transform: uppercase; color: var(--faint, #66635a); }
.fil-t { font-weight: 600; text-wrap: pretty; }
.f-nav { position: sticky; top: var(--bar-h, 0px); z-index: 5; background: var(--paper, #eeebe4); border-bottom: 1px solid var(--ink, #151512); }
.f-nav-in { display: flex; gap: var(--s2, 8px); align-items: center; overflow-x: auto; padding: var(--s2, 8px) 0; scrollbar-width: none; }
.f-nav a { min-width: var(--s8, 33px); justify-content: center; }
.f-nav a, .f-nav button { flex: none; display: inline-flex; align-items: center; gap: var(--s2, 8px); font: 400 .72rem/1 var(--mono, ui-monospace, monospace); letter-spacing: .06em;
           text-transform: uppercase; text-decoration: none; padding: var(--s2, 8px) var(--s3, 10px); border: 1px solid var(--hair-2, #1515124d); background: transparent; color: var(--ink, #151512); cursor: pointer; }
.f-nav a span { font-weight: 500; }
.f-nav a:hover, .f-nav button:hover { border-color: var(--ink, #151512); }
.f-nav button[aria-pressed="true"] { background: var(--ink, #151512); color: var(--paper, #eeebe4); border-color: var(--ink, #151512); }
#f-ouvrir { margin-left: auto; }
.fiche { padding: var(--s10, 53px) 0 var(--s8, 33px); border-bottom: 1px solid var(--ink, #151512); scroll-margin-top: calc(var(--bar-h, 0px) + var(--s9, 42px)); }
.f-tete { display: grid; grid-template-columns: var(--s9, 42px) minmax(0, 1fr); gap: var(--s5, 16px); align-items: start; margin-bottom: var(--s7, 26px); }
.f-n { display: grid; place-items: center; width: var(--s9, 42px); height: var(--s9, 42px); background: var(--ink, #151512); color: var(--paper, #eeebe4); font: 500 1.05rem/1 var(--mono, ui-monospace, monospace); }
.fiche h2 { font: 700 clamp(1.35rem, 2.6vw, 1.9rem)/1.12 var(--sans, system-ui, sans-serif); letter-spacing: -.015em; text-transform: none; color: var(--ink, #151512); margin: var(--s1, 6px) 0 0; max-width: 34ch; text-wrap: balance; }
.f-grille { display: grid; grid-template-columns: minmax(0, 4fr) minmax(0, 7fr); gap: var(--s8, 33px) var(--s9, 42px); align-items: start; }
.b-qa, .b-rep, .b-tous { grid-column: 1 / -1; }
.b-msg { grid-column: 1; } .b-lim { grid-column: 1; } .b-chf { grid-column: 2; grid-row: 1 / span 2; }
.dossier .b-msg { grid-row: 2; } .dossier .b-lim { grid-column: 2; grid-row: 2; }
.bloc { min-width: 0; }
.fiche h3 { padding-bottom: var(--s2, 8px); border-bottom: 1px solid var(--ink, #151512); margin-bottom: var(--s3, 10px); color: var(--ink, #151512); }
ol.messages { list-style: none; counter-reset: m; margin: 0; padding: 0; }
ol.messages > li { counter-increment: m; display: grid; grid-template-columns: var(--s7, 26px) minmax(0, 1fr); padding: var(--s2, 8px) 0; border-bottom: 1px solid var(--hair, #15151224); font-weight: 600; line-height: 1.35; text-wrap: pretty; }
ol.messages > li::before { content: counter(m); font: 400 .78rem/1.7 var(--mono, ui-monospace, monospace); color: var(--faint, #66635a); }
ol.messages > li > .sous { grid-column: 2; font-weight: 400; color: var(--mut, #55524c); font-size: .9rem; margin-top: 2px; }
.autres { margin: var(--s3, 10px) 0 0; font-size: .86rem; color: var(--mut, #55524c); line-height: 1.45; }
.autres .anno { display: block; margin-bottom: 2px; }
ul.limites { margin: 0; padding: 0; list-style: none; }
ul.limites li { position: relative; padding: var(--s2, 8px) 0 var(--s2, 8px) var(--s7, 26px); border-bottom: 1px solid var(--hair, #15151224); line-height: 1.45; }
ul.limites li::before { content: "?"; position: absolute; left: 0; top: var(--s2, 8px); font: 500 .8rem/1.7 var(--mono, ui-monospace, monospace); color: var(--faint, #66635a); }
.tbl { overflow-x: auto; }
table { width: 100%; border-collapse: collapse; font-size: .88rem; }
th, td { text-align: left; padding: var(--s2, 8px) var(--s3, 10px) var(--s2, 8px) 0; border-bottom: 1px solid var(--hair, #15151224); vertical-align: top; }
th { padding-top: 0; border-bottom-color: var(--hair-2, #1515124d); font-size: .68rem; }
th:last-child, td:last-child { padding-right: 0; }
td.num, th.num { text-align: right; white-space: nowrap; font-variant-numeric: tabular-nums; }
td.num { font: 500 .9rem/1.5 var(--mono, ui-monospace, monospace); }
td.lib { line-height: 1.4; }
td.lib .l { display: block; }
.src { display: block; margin-top: 3px; font: 400 .72rem/1.5 var(--mono, ui-monospace, monospace); color: var(--mut, #55524c); }
.src a { color: var(--ink, #151512); text-decoration-color: var(--hair-2, #1515124d); }
.src .etat::before { content: " · "; }
@media print, (min-width: 901px) { .b-tous td.lib { display: grid; grid-template-columns: minmax(0, 1fr) minmax(0, 26rem); gap: 0 var(--s7, 26px); align-items: baseline; } .b-tous .src { margin-top: 0; } }
.id { display: inline-block; padding: 0 4px; border: 1px solid var(--hair-2, #1515124d); color: var(--ink, #151512); font-size: .68rem; margin-right: 2px; }
.etat { white-space: nowrap; color: var(--faint, #66635a); }
tr.calc td.lib { font-style: italic; }
table.texte td, table.texte th { padding-right: var(--s4, 13px); }
a.ch, span.ch { white-space: nowrap; font-variant-numeric: tabular-nums; font-weight: 600; text-decoration: underline dotted 1px; text-underline-offset: 3px; }
a.ch:hover { color: var(--pig, #1a22c4); text-decoration-style: solid; }
details.qa { border-bottom: 1px solid var(--hair, #15151224); }
details.qa > summary { cursor: pointer; list-style: none; display: grid; grid-template-columns: var(--s7, 26px) minmax(0, 1fr); padding: var(--s3, 10px) 0; font-weight: 600; line-height: 1.4; text-wrap: pretty; }
details.qa > summary::-webkit-details-marker { display: none; }
details.qa > summary::before { content: "+"; font: 400 .95rem/1.4 var(--mono, ui-monospace, monospace); color: var(--faint, #66635a); }
details.qa[open] > summary::before { content: "−"; }
details.qa > summary:hover { color: var(--pig, #1a22c4); }
.rep { padding: 0 0 var(--s4, 13px) var(--s7, 26px); max-width: var(--f-mesure); }
.rep p, .b-rep p { margin: 0 0 var(--s2, 8px); }
.rep ul, .rep ol, .b-rep ul { margin: 0 0 var(--s2, 8px); padding-left: 1.2em; }
.rep h4 { font: 600 .92rem/1.4 var(--sans, system-ui, sans-serif); margin: var(--s4, 13px) 0 var(--s1, 6px); }
.rep .tbl { margin: var(--s2, 8px) 0 var(--s3, 10px); }
.b-rep { max-width: var(--f-mesure); font-size: 1.02rem; }
.onglets { margin: var(--s3, 10px) 0 0; font: 400 .74rem/1.5 var(--mono, ui-monospace, monospace); color: var(--mut, #55524c); }
.pied { margin-top: var(--s8, 33px); max-width: var(--f-mesure); font: 400 .74rem/1.5 var(--mono, ui-monospace, monospace); color: var(--mut, #55524c); }
.suivi { padding: var(--s8, 33px) 0 0; }
.suivi h2 { font: 700 1.2rem/1.2 var(--sans, system-ui, sans-serif); text-transform: none; letter-spacing: -.01em; color: var(--ink, #151512); margin: 0 0 var(--s3, 10px); }
.suivi ul { margin: 0; padding-left: 1.2em; max-width: var(--f-mesure); }
@media screen and (max-width: 900px) {
  .f-grille { grid-template-columns: minmax(0, 1fr); gap: var(--s7, 26px); }
  .f-grille > .bloc, .dossier .b-lim { grid-column: 1; grid-row: auto; }
  ol.fil a { grid-template-columns: var(--s8, 33px) minmax(0, 1fr); }
  .fil-r, .fil-t { grid-column: 2; }
}
@media screen and (max-width: 560px) {
  body { font-size: 15px; }
  .fiche { padding-top: var(--s8, 33px); }
  .f-tete { grid-template-columns: var(--s8, 33px) minmax(0, 1fr); gap: var(--s4, 13px); }
  .f-n { width: var(--s8, 33px); height: var(--s8, 33px); font-size: .9rem; }
  .f-nav-in { gap: var(--s1, 6px); }
  .f-nav a, .f-nav button { padding: var(--s2, 8px) var(--s1, 6px); }
  .f-nav a { min-width: 30px; }
  .rep { padding-left: 0; }
  #f-imprimer { display: none; }
}
/* L'impression : A4, marges 2,5 cm (is42.css pose aussi @page) ; chaque fiche commence sur sa page, la couverture porte le fil ;
   un bloc court ne se coupe pas ; les en-têtes de tableau se répètent ; les réponses s'impriment ouvertes. */
@page { size: A4; margin: 2.5cm; }
@media print {
  html { font-size: 9pt; }
  body { font-size: 9pt; line-height: 1.38; background: #fff; }
  .wrap { padding: 0; max-width: none; }
  .f-nav { display: none !important; }
  h1 { font-size: 26pt; }
  .fiche { break-before: page; page-break-before: always; padding: 0; border: 0; }
  .f-tete { margin-bottom: 12pt; break-after: avoid; }
  .f-n { background: none; color: #151512; border: 1px solid #151512; }
  .f-grille { grid-template-columns: minmax(0, 5fr) minmax(0, 6fr); gap: 12pt 18pt; }
  .dossier .f-grille { display: block; }
  .b-tous td.lib { grid-template-columns: minmax(0, 6fr) minmax(0, 5fr); gap: 0 12pt; }
  /* une fiche = une feuille : au recto la thèse, les messages, les chiffres et les limites ; au verso les questions du jury */
  .fiche:not(.dossier) .b-qa { break-before: page; page-break-before: always; }
  .fiche:not(.dossier) details.qa > summary { padding-top: 9pt; }
  .fiche:not(.dossier) .rep { padding-bottom: 9pt; font-size: 10pt; }
  .dossier .bloc { margin-bottom: 14pt; }
  .f-couv { padding-bottom: 0; }
  .f-couv h2 { margin-top: 16pt; }
  ol.fil a { padding: 5pt 0; grid-template-columns: 22pt minmax(0, 11rem) minmax(0, 1fr); }
  .fil-texte { font-size: 10.5pt; }
  ol.messages > li, ul.limites li, th, td { padding-top: 4pt; padding-bottom: 4pt; }
  details.qa > summary { padding: 5pt 0 2pt; }
  .rep { padding-bottom: 5pt; }
  .rep p { margin-bottom: 3pt; }
  .b-msg, .b-lim, ol.messages > li, ul.limites li, details.qa, tr, ol.fil li, .f-tete { break-inside: avoid; page-break-inside: avoid; }
  .fiche h3, summary { break-after: avoid; page-break-after: avoid; }
  thead { display: table-header-group; }
  .tbl { overflow: visible !important; }
  details.qa::details-content { content-visibility: visible !important; display: block !important; height: auto !important; }
  details.qa > summary::before { content: "Q"; }
  details.qa[open] > summary::before { content: "Q"; }
  a { color: inherit; }
  a.ch, span.ch { text-decoration: none; }
  td.src a, .rep a:not(.ch) { text-decoration: underline; }
  .fiche h2 { max-width: none; font-size: 15pt; }
  .rep { max-width: none; }
}
"""

# L'alerte : le pigment d'is42, seule couleur de la page. Jointe à la feuille seulement quand une anomalie existe.
CSS_ALERTE = """
.alerte { color: var(--pig, #1a22c4); font-weight: 600; }
span.alerte { font-family: var(--mono, ui-monospace, monospace); font-size: .9em; border: 1px solid var(--pig, #1a22c4); padding: 0 4px; white-space: nowrap; }
tr.alerte td, tr.alerte .src, tr.alerte .src .etat { color: var(--pig, #1a22c4); }
p.alerte { margin: 0; padding: var(--s3, 10px) var(--s5, 16px); border-left: 2px solid var(--pig, #1a22c4); max-width: var(--f-mesure); }
ul.limites li.alerte::before { content: "!"; color: var(--pig, #1a22c4); }
.f-alerte { margin: var(--s7, 26px) 0 0; padding: var(--s4, 13px) var(--s5, 16px); border: 1px solid var(--pig, #1a22c4); border-left-width: var(--s1, 6px); max-width: var(--f-mesure); }
.f-alerte .anno { color: var(--pig, #1a22c4); font-weight: 500; }
.f-alerte ul { margin: var(--s2, 8px) 0 0; padding-left: 1.2em; font-size: .9rem; line-height: 1.45; }
.f-alerte code { font-size: .8rem; }
@media print { .f-alerte, p.alerte, tr.alerte { break-inside: avoid; } }
"""


def css_alerte(corpus):
    return CSS_ALERTE if corpus.anomalies else ""


# Tout ouvrir, tout refermer ; imprimer. Les <details> s'ouvrent seuls le temps de l'impression (is42.PRINT_DETAILS).
JS = """<script>(function(){var o=document.getElementById("f-ouvrir"),i=document.getElementById("f-imprimer");
if(o)o.addEventListener("click",function(){var on=o.getAttribute("aria-pressed")!=="true";
document.querySelectorAll("details.qa").forEach(function(d){d.open=on});o.setAttribute("aria-pressed",on?"true":"false");
o.textContent=on?"Tout refermer":"Tout ouvrir"});if(i)i.addEventListener("click",function(){window.print()})})();</script>"""


def page_html(titre, corps, css=""):
    """La page entière, habillée par is42 (barre, polices, thème clair ou sombre, règles d'impression)."""
    page = ('<!doctype html>\n<html lang="fr">\n<head>\n<meta charset="utf-8">\n'
            '<meta name="viewport" content="width=device-width, initial-scale=1">\n<meta name="robots" content="noindex, nofollow">\n'
            f"<title>{esc(titre)}</title>\n<style>{CSS}{css}</style>\n</head>\n<body>\n{corps}\n{JS}\n</body>\n</html>\n")
    from dataroom import is42
    return is42.apply(page, "livrables")


def main(argv):
    argv = [a for a in argv if a != "--strict"]
    sortie = Path(argv[2] if len(argv) > 2 else LIV / "fiches-oral.html")
    corpus = Corpus()
    parties = deriver(corpus)
    corps = f"""<main class="wrap">
<header class="f-couv">
  <p class="anno">Capstone EDHEC · Sujet 2 · Veolia, GreenUp 2027</p>
  <h1>Fiches d'oral</h1>
  <p class="lede">Une fiche par partie du rapport, dans son ordre : la thèse, les messages à tenir, les chiffres clés et leur source,
  ce que nous n'avons pas pu établir, les questions probables du jury. Chaque ligne vient du texte du rapport ; chaque chiffre s'ouvre sur sa page.</p>
  <p class="prov">{esc(provenance(corpus))}</p>{alertes_html(corpus)}
  <h2>Le fil de l'oral</h2>
  {fil_html(corpus, parties)}
</header>
{nav_html(parties)}
{''.join(fiche_html(corpus, p) for p in parties)}
</main>"""
    sortie.write_text(page_html("Fiches d'oral — sujet 2", corps, css_alerte(corpus)), encoding="utf-8")
    signaler(corpus)
    print("écrit", sortie, len(parties), "fiches,", sum(len(p["cles"]) for p in parties), "chiffres clés,",
          sum(len(p["messages"]) for p in parties), "messages")


if __name__ == "__main__":
    main(sys.argv)
