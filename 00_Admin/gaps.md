# Gaps — what a real buyer would see and we cannot

Outside-in dataroom: public documents only. This file is the honest list of what a
genuine buy-side data room would contain and ours does not. Say what is missing, why
it matters, and what we used instead. Keep one `##` section per target so
`list_gaps(target=...)` can filter.

## Veolia — hazardous waste perimeter

- **Site-level P&L for the hazardous waste business** — not disclosed; the segment is
  reported only at group level. We proxy with segment revenue and published capacity.
  *Impact: margin by site is an estimate, never a sourced figure.*
- **Customer contracts and pricing terms** — commercially confidential. We use tender
  notices and published tariffs where they exist.
- **Environmental liability provisions per site** — only the aggregate provision is
  published. Per-site exposure is inferred from permit conditions and incident records.

## (add a section per target)

## Général

- **Comptes de Clean Harbors : 10-K 2025** — Rôle 4. Premier comparable américain de Clean Earth (TSDF, incinération). ir.cleanharbors.com et sec.gov renvoient 403 au serveur : à télécharger depuis un navigateur et à verser à la main. *(signalé par dataroom-bearer)*
- **Univers de cibles disponibles en déchets dangereux (hors Clean Earth)** — Rôle 6. Aucun document de la dataroom ne nomme une cible encore disponible : les candidats viennent de la presse et des rapports de pairs. À sourcer par submit_url (communiqués, presse spécialisée) puis register_target, avec la valeur publiée ou « non communiqué ». *(signalé par dataroom-bearer)*
- **Sources réglementaires officielles impossibles à capturer (10/10/2026) : (1) toutes les pages ECHA sur la restriction universelle des PFAS (communiqué « ECHA supports PFAS restriction with targeted derogations », page hot-topics PFAS, registre des intentions de restriction, documents 10162/…), qui renvoient une erreur 403 (pare-feu Azure) ; (2) Légifrance (JORFTEXT000051236798, loi n° 2025-188), erreur 403 ; (3) la notice de retrait de la proposition RCRA « Listing of Specific PFAS as Hazardous Constituents » : aucune notice officielle trouvée, seul le retrait de la règle compagne (91 FR 25266, 8 mai 2026) est publié ; (4) les montants 2026 de la TGAP déchets dans le CIBS (art. L. 433-57 et L. 433-86), non vérifiés.** — Ces textes fixent le calendrier PFAS de l'UE (avis final SEAC attendu fin 2026, décision en 2027), la loi PFAS française et le coût fiscal de l'élimination. Remplacements utilisés : présentation EMA du 4/02/2026 qui reprend le calendrier ECHA ; texte de la loi sur AIDA (INERIS) ; agenda réglementaire reginfo.gov (règle finale RCRA annoncée pour 01/2027) cité dans la note du document ; rescrit BOFiP BOI-RES-TCA-000259 pour le transfert de la TGAP vers le CIBS. La capture de la page thématique AIDA 2790 (02_Permits_Regulatory/fr-icpe-rubrique-2790-aida.html) est presque vide (JavaScript) : utiliser fr-icpe-rubrique-2790-texte-aida.html. *(signalé par dataroom-bearer)*

## Clean Earth

- **Comptes d'Enviri, vendeur de Clean Earth : 10-K 2025, 10-Q du T1 2026, communiqué de vente du 20/11/2025** — Rôles 5 et 6. Ils donnent les passifs environnementaux de Clean Earth (activité abandonnée chez Enviri) et le multiple côté vendeur (18,6x l'EBITDA ajusté des 12 derniers mois selon Enviri, contre 9,8x après synergies selon Veolia). investors.enviri.com et sec.gov renvoient 403 au serveur : à télécharger depuis un navigateur et à verser à la main. *(signalé par dataroom-bearer)*

## Enviracore Services

- **Permis RCRA / TSDF d'EMI (Guthrie, Oklahoma) et capacités de traitement réellement détenues** — C'est le principal actif dur possible du groupe ; le profil SAM.gov (NAICS 562211) indique une activité de traitement et d'élimination de déchets dangereux mais ne prouve pas l'existence d'un permis TSDF. À vérifier dans EPA RCRAInfo / ECHO et auprès de l'Oklahoma DEQ. *(signalé par noa)*
- **Chiffre d'affaires, EBITDA, effectifs et prix d'acquisition d'Enviracore et de ses trois filiales (Boomer Environmental, Environmental Cleanup Inc., EMI)** — Aucune valeur publiée : la valorisation reste une hypothèse (fourchette indicative d'environ 35 à 120 M$ sur 5 à 12 M$ d'EBITDA supposé et 7x à 10x), à ne pas présenter comme un fait. Pistes : PitchBook, dépôts UCC, effectifs LinkedIn, dossiers de l'Oklahoma DEQ. *(signalé par noa)*
- **Communiqués originaux non capturés : acquisition d'EMI (Business Wire mai 2026 / Guthrie News Leader avril 2026), acquisition de Boomer Environmental, article du San Diego Business Journal, dépôt de marque USPTO n° 99349742** — Les sites ont refusé la capture automatique (403, 429, 503). Les faits restent citables via les URL publiques, mais il faut télécharger ces pages à la main et les déposer dans l'inbox pour les figer. *(signalé par noa)*

## Envisol

- **Prix payé par Géotec en février 2025, EBITDA et marge d'Envisol** — Aucun montant publié ; la valorisation d'environ 5 à 15 M€ repose sur des hypothèses (CA 8 à 10 M€, marge EBITDA 8 à 12 %, 7x à 10x, prime de contrôle). Les comptes déposés au greffe (Infogreffe, Pappers) donneraient le résultat d'exploitation réel. *(signalé par noa)*
- **Processus de cession éventuel de Géotec (titre CFNews « Géotec entame sa transmission ») et sources non capturées : note Techleap/Fusacq sur le rachat, profil PitchBook, fiche SCAN 360 de la Revue EIN, brochure élus 2025 de l'AMRF** — Si Géotec est en vente, Envisol devient accessible avec le groupe ou en carve-out ; seul le titre de l'article a été vu. Les quatre autres sources ont refusé la capture (403, 404, format) et doivent être téléchargées à la main puis déposées dans l'inbox. *(signalé par noa)*
- **Brevets, dépôts de marque (EnviRisk®, SCAN 360) et certification LNE sites et sols pollués (norme NF X31-620, domaines A/B/C)** — Aucun brevet trouvé par recherche web, non concluant : à vérifier dans INPI et Espacenet. La certification conditionne l'argument d'indépendance entre diagnostic et travaux si Veolia, qui traite les terres, en devenait propriétaire. *(signalé par noa)*
