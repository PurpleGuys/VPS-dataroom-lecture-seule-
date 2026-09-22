# Premiers pas — vault de la dataroom

Ce dossier est le vault Obsidian du groupe. Il est lu par le serveur MCP
`https://dataroom.is42.fr/mcp`, qui l'indexe et nous renvoie des passages citables.

## Une fois pour toutes

1. Ouvrir ce dossier comme vault dans Obsidian (*Open folder as vault*).
2. Installer le plugin communautaire **Obsidian Git** et le régler sur commit + push
   automatique (toutes les 10 minutes, par exemple).

## À chaque document ajouté

Deux fichiers, toujours ensemble : le PDF, et une note portant **exactement le nom du
PDF suivi de `.md`**. Voir `00_Admin/README.md` pour les champs obligatoires.

Sans `source_url` ni `date_consulted`, le serveur refuse le document — il apparaît
alors dans `dataroom_status()` avec la raison. Ce n'est pas une perte, c'est un rappel :
un document qu'on ne peut pas citer ne sert à rien dans le rapport.

## Le serveur se met à jour tout seul

Il fait un `git pull` toutes les cinq minutes et réindexe ce qui a changé.
Rien à lancer à la main.
