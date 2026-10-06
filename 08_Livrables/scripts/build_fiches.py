"""Fiches d'oral : question → réponse et chiffre → source, par rôle, depuis report-data.json.

Rien n'est écrit ici que le dossier ne dise déjà : les réponses sont celles des dossiers,
les chiffres ceux du registre. Usage : python build_fiches.py [report-data.json] [fiches-oral.html]
"""
import html
import json
import sys
from pathlib import Path

HERE = Path(__file__).parent
DATA = json.loads(Path(sys.argv[1] if len(sys.argv) > 1 else HERE / "report-data.json").read_text(encoding="utf-8"))
OUT = Path(sys.argv[2] if len(sys.argv) > 2 else HERE / "fiches-oral.html")
esc = html.escape

cards = []
for role in DATA["roles"]:
    n, name = role["n"], role["name"]
    for q, a in role["qa"]:
        cards.append({"role": n, "kind": "question", "front": q, "back": a, "tag": f"Rôle {n} — {name}"})
    for fid, label, period, value, source, checked in role["figures"]:
        cards.append({"role": n, "kind": "chiffre", "front": f"{label}" + (f" ({period})" if period else ""),
                      "back": f"{value} — {source}" + (f" — relu par {checked}" if checked else " — pas encore relu"),
                      "tag": f"{fid} · rôle {n}"})
for q in DATA.get("research_questions", []):
    cards.append({"role": 0, "kind": "cours", "front": q, "back": "Où le rapport répond : voir la section « Les quatre questions » et le rôle concerné.", "tag": "Question du cours"})

PAGE = """<!doctype html>
<html lang="fr">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="robots" content="noindex, nofollow">
<title>Fiches d'oral — sujet 2</title>
<style>
  :root { --ground: #F4F5F2; --panel: #FBFCFA; --ink: #1A211D; --muted: #6E7873; --rule: #DDE1DB; --accent: #1F6F50; }
  @media (prefers-color-scheme: dark) { :root { --ground: #131815; --panel: #1A201C; --ink: #E4E9E5; --muted: #97A29C; --rule: #2B332E; --accent: #6FBF95; } }
  body { margin: 0; background: var(--ground); color: var(--ink); font: 15px/1.5 "IBM Plex Sans", system-ui, sans-serif; }
  header { padding: 1rem 1.2rem .6rem; border-bottom: 1px solid var(--rule); background: var(--panel); position: sticky; top: 0; z-index: 2; }
  h1 { margin: 0 0 .3rem; font-size: 1.15rem; }
  .bar { display: flex; flex-wrap: wrap; gap: .5rem .9rem; align-items: center; font-size: .85rem; color: var(--muted); }
  .bar button, .bar label { font: inherit; cursor: pointer; }
  .bar button { border: 1px solid var(--rule); background: var(--panel); color: var(--ink); border-radius: 4px; padding: .25rem .6rem; }
  .bar button.on { background: var(--accent); color: #fff; border-color: var(--accent); }
  .grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(300px, 1fr)); gap: .8rem; padding: 1rem 1.2rem 3rem; }
  .card { background: var(--panel); border: 1px solid var(--rule); border-radius: 8px; padding: .8rem .9rem; cursor: pointer; min-height: 7rem;
          display: flex; flex-direction: column; gap: .4rem; outline: none; }
  .card:focus-visible { box-shadow: 0 0 0 2px var(--accent); }
  .card .tag { font-size: .72rem; color: var(--muted); font-family: "IBM Plex Mono", ui-monospace, monospace; }
  .card .front { font-weight: 600; }
  .card .back { color: var(--ink); border-top: 1px dashed var(--rule); padding-top: .4rem; display: none; font-size: .92rem; }
  .card.open .back { display: block; }
  .card.chiffre .front::before { content: "Chiffre · "; color: var(--muted); font-weight: 400; }
  .card.question .front::before { content: "Q · "; color: var(--accent); }
  .hidden { display: none !important; }
  .count { margin-left: auto; }
  @media (max-width: 720px) { .grid { padding: .8rem; grid-template-columns: 1fr; } }
</style>
</head>
<body>
<header>
  <h1>Fiches d'oral — sujet 2, capacité financière</h1>
  <div class="bar">
    <span>Rôle :</span>
    __ROLE_BUTTONS__
    <span>·</span>
    <button type="button" data-kind="question" class="on">Questions</button>
    <button type="button" data-kind="chiffre" class="on">Chiffres</button>
    <button type="button" data-kind="cours" class="on">Cours</button>
    <span>·</span>
    <button type="button" id="recto">Recto seul</button>
    <button type="button" id="shuffle">Mélanger</button>
    <button type="button" id="reveal">Tout révéler</button>
    <span class="count" id="count"></span>
  </div>
</header>
<main class="grid" id="grid"></main>
<script>
(function () {
  "use strict";
  var CARDS = __CARDS__;
  var state = { role: "all", kinds: { question: true, chiffre: true, cours: true }, recto: false, order: CARDS.map(function (_, i) { return i; }) };
  function esc(s) { return String(s == null ? "" : s).replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; }); }
  function render() {
    var grid = document.getElementById("grid"), html = "", shown = 0;
    state.order.forEach(function (i) {
      var c = CARDS[i];
      if (state.role !== "all" && String(c.role) !== state.role) return;
      if (!state.kinds[c.kind]) return;
      shown++;
      html += '<article class="card ' + c.kind + '" tabindex="0" data-i="' + i + '"><div class="tag">' + esc(c.tag) + '</div>' +
              '<div class="front">' + esc(c.front) + '</div><div class="back">' + esc(c.back) + "</div></article>";
    });
    grid.innerHTML = html;
    document.getElementById("count").textContent = shown + " fiche" + (shown > 1 ? "s" : "");
    grid.querySelectorAll(".card").forEach(function (el) {
      var flip = function () { if (!state.recto) el.classList.toggle("open"); };
      el.addEventListener("click", flip);
      el.addEventListener("keydown", function (e) { if (e.key === " " || e.key === "Enter") { e.preventDefault(); flip(); } });
    });
  }
  document.querySelectorAll("[data-role]").forEach(function (b) {
    b.addEventListener("click", function () {
      state.role = b.getAttribute("data-role");
      document.querySelectorAll("[data-role]").forEach(function (x) { x.classList.toggle("on", x === b); });
      render();
    });
  });
  document.querySelectorAll("[data-kind]").forEach(function (b) {
    b.addEventListener("click", function () { var k = b.getAttribute("data-kind"); state.kinds[k] = !state.kinds[k]; b.classList.toggle("on", state.kinds[k]); render(); });
  });
  document.getElementById("recto").addEventListener("click", function (e) {
    state.recto = !state.recto; e.target.classList.toggle("on", state.recto);
    if (state.recto) document.querySelectorAll(".card.open").forEach(function (c) { c.classList.remove("open"); });
  });
  document.getElementById("shuffle").addEventListener("click", function () {
    for (var i = state.order.length - 1; i > 0; i--) { var j = Math.floor(Math.random() * (i + 1)); var t = state.order[i]; state.order[i] = state.order[j]; state.order[j] = t; }
    render();
  });
  document.getElementById("reveal").addEventListener("click", function () { if (!state.recto) document.querySelectorAll(".card").forEach(function (c) { c.classList.add("open"); }); });
  render();
})();
</script>
</body>
</html>
"""
buttons = '<button type="button" data-role="all" class="on">tous</button>' + "".join(
    f'<button type="button" data-role="{r["n"]}" title="{esc(r["name"])}">{r["n"]}</button>' for r in DATA["roles"])
page = PAGE.replace("__ROLE_BUTTONS__", buttons).replace("__CARDS__", json.dumps(cards, ensure_ascii=False).replace("</", "<\\/"))
OUT.write_text(page, encoding="utf-8")
print("écrit", OUT, len(cards), "fiches")
