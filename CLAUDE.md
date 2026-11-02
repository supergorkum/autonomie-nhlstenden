# CLAUDE.md, Autonomie NHL Stenden

Geheugen van dit project. Lees dit eerst, daarna pas aan het werk.
Aangemaakt op 1 november 2026 bij de overstap van de chat naar Claude Code. Opgesteld op basis van de projectmap, de git-historie en Edgars geheugen, niet op basis van de oude chat zelf. De overdracht staat in `.overstap/OVERDRACHT-UIT-CHAT.md` (niet in git); wat daarvan blijvend geldt, staat hier.

## Eerste handeling van elke sessie

```
cd "<pad naar CODE - Claude>/autonomie-nhlstenden"
git pull
grep -n 'const VERSION' src/App.jsx
git log --oneline | head -5
git status
git config core.hooksPath                 # hoort scripts/git-hooks te zijn
ls -l scripts/git-hooks/pre-push          # hoort uitvoerbaar te zijn (-rwx)
```

Google Drive bewaart het uitvoerbit niet altijd. Is `pre-push` niet uitvoerbaar, zet het terug met `chmod +x scripts/git-hooks/pre-push`; anders slaat git de controle stil over. Draai nooit `npm install` in de projectmap; zie "Werken in Claude Code".

## Wat de site is

- Naam: Portfolioanalyse Digitale Autonomie, in de app "Digitale Autonomie-assessment" (paginatitel in index.html). Het thema heette tot september 2026 Digitale Soevereiniteit; die naam blijft alleen historisch staan (map, package-naam, expertisegroep).
- Een werkend prototype waarmee NHL Stenden het applicatielandschap beoordeelt op digitale autonomie: 15 tot 20 applicaties, scoring volgens de DAAF (Universiteit Utrecht) en DICTU-kaders. Met dashboard, vergelijking, portfoliokaart, spindiagram, PDF- en Excel-export, handleiding, transparantiepagina en een beheeromgeving.
- Voor wie: de kerngroep en expertisegroep Digitale Autonomie, ambassadeurs (J. Haije, E. Rolf, J. Blom), portefeuillehouder Carlo Segers, en de collega's van contractbeheer en inkoop. Kwartiermaker: Edgar van Gorkum.
- Taal: de interface is Nederlands (`<html lang="nl">`). Of er ook Engels in zit, is niet gecontroleerd.
- Context: de Autonomie-tool volgt de aanpak van de Universiteit Utrecht. Een SharePoint-versie ("Autonomiescan") bestaat los van deze site en wordt niet hier onderhouden. Een interne versie draait op https://autonomie.nhlstenden.com/ (alleen via het netwerk van NHL Stenden); of dat dezelfde code is, staat niet vast.

## Techniek

- React 18 met Vite 5, Tailwind 3, Recharts, d3 en xlsx. Vrijwel alle code staat in één bestand: `src/App.jsx` (circa 7300 regels). Daarnaast `src/main.jsx` en `src/index.css`.
- Opslag: Netlify Blobs, store `nhl-assessment`, sleutel `apps` (hele lijst als JSON). Back-ups staan in dezelfde store onder `backup:<tijdstempel>`.
- Netlify Functions in `netlify/functions/` (.mjs): load-data, save-data, list-backups, restore-backup, force-backup, scheduled-backup, load-export-timestamp, save-export-timestamp. `/api/*` wordt doorgestuurd naar `/.netlify/functions/*`.
- Bouwen: Netlify draait `npm run build` (publish-map `dist`). Zelf bouwen en bekijken gaat via `npm run controle` en `npm run lokaal` in de bouwmap; zie "Werken in Claude Code".
- Hosting: Netlify. Live adres: https://autonomie-nhlstenden.netlify.app
- Netlify: site `autonomie-nhlstenden`, site-id `6d6a1b8a-9c17-469c-9f23-c8911db814d6`, team `6a047d5b4965c04ab89fec76` (hetzelfde team als autonomie-frisius-mc), plan Pro. Beheer: https://app.netlify.com/projects/autonomie-nhlstenden. Het id in `.netlify/state.json` (c225450e...) hoort bij een andere, oude site; niet gebruiken.
- De hele site zit achter het Netlify-wachtwoord (Password Protection, alle deploys). Wie de site opent, ziet eerst het Netlify-scherm en daarna het inlogscherm van de app. Het Netlify-wachtwoord staat niet in deze map.
- Blobs en functions draaien in de VS (us-east-2); de transparantiepagina noemt dat sinds v2.6 ook zo.
- GitHub: https://github.com/supergorkum/autonomie-nhlstenden, productietak `main`. De repo is PUBLIEK (zie Aandachtspunten onderaan).
- Geplande back-up: `0 8,11,16 * * *` in UTC. In de zomer is dat 10:00, 13:00 en 18:00 Nederlandse tijd, in de winter een uur eerder.
- Plan: migratie naar eigen NHL Stenden-infrastructuur (NVIDIA Spark, Leeuwarden). De tekst op de transparantiepagina en de autonomiescores van het instrument zelf gaan daarvan uit. Nog niet uitgevoerd.

## Werken in Claude Code (sinds 1 november 2026)

**Een push naar `main` publiceert de site.** Netlify is aan de GitHub-repo gekoppeld en bouwt zelf bij elke push naar `main` (vastgesteld op 1 november 2026: commit e9d5c67 stond binnen een minuut live als deploy 6ae755e7b850b9f21525f998). Mislukt die bouw, dan blijft de vorige versie live. Er is geen deploy-script en er is er ook geen nodig. GitHub krijgt van Netlify geen status of check terug; kijk dus in Netlify zelf (zie "Live controleren").

**Bouwen gebeurt buiten Google Drive.** In de projectmap staat geen `node_modules`, en dat blijft zo. Alles wat node_modules nodig heeft draait in een kopie, standaard `~/Developer/autonomie-nhlstenden-build`, die `scripts/bouwmap.mjs` bijhoudt (andere plek: `ANHL_BOUWMAP`). npm gebruikt daar een eigen cache in de bouwmap. Draai nooit `npm install` of `npm ci` in de projectmap. De oude `dist/` uit de projectmap staat sinds 1 november 2026 in de Prullenmand (map autonomie-nhlstenden-opruimen-20261101).

- `npm run controle`: versie en changelog (VERSION gelijk aan de bovenste CHANGELOG-entry, nummer niet al eerder gebruikt, package.json loopt gelijk), nergens een gedachtestreepje in een tekstbestand dat git meeneemt, de onderhoudsregel op het inlogscherm, geen codes uit `.sleutels/` in een ander bestand dan `src/App.jsx`, en `vite build` in de bouwmap met de controle dat de onderhoudsregel in de bundel zit. Er is geen eslint en geen test; de bouw vangt alleen syntaxfouten en kapotte imports. `--snel` slaat de bouw over.
- `npm run lokaal`: de site op http://127.0.0.1:5173, ook als preview in Code (`.claude/launch.json`, naam `autonomie-nhlstenden`). Wijzigingen in de projectmap lopen vanzelf mee. `/api` gaat via een proxy naar de live site (`ANHL_LIVE_PROXY` in `vite.config.js`). Zolang het Netlify-wachtwoord erop zit, geeft dat een 401 en laadt de proef geen gegevens; het inlogscherm en de opmaak zijn wel te bekijken.
- `npm run bouwmap`: alleen de bouwmap bijwerken.

**Elke push gaat eerst door de controle.** `scripts/git-hooks/pre-push` draait `node scripts/controle.mjs --push`, met een extra stap: raakt de push `src/`, `netlify/`, `public/` of `index.html`, dan moet `VERSION` in `src/App.jsx` anders zijn dan op `origin/main`. Ingesteld met `git config core.hooksPath scripts/git-hooks` (staat in `.git/config`). `--no-verify` alleen als Edgar daar zelf om vraagt.

**Een wijziging** gaat zo:

1. Bestanden rechtstreeks aanpassen in de projectmap. `src/App.jsx` is groot: zoek met grep, wijzig gericht.
2. `const VERSION` met 0.1 ophogen en een nieuwe bovenste entry in `CHANGELOG` (beide bovenin `src/App.jsx`), datum voluit ("1 november 2026"). `version` in package.json mee laten lopen (v2.6 wordt 2.6.0). Het label "Huidige versie" zet de app zelf bij de bovenste entry.
3. `npm run controle`, en bij iets zichtbaars de proef bekijken met `npm run lokaal`.
4. Commit met een bericht in de vorm `v2.6: wat de bezoeker merkt`, dan `git push`.
5. Ongeveer een minuut later de deploy nakijken en Edgar melden wat er live staat.

**Live controleren.** Met `curl` lukt het niet: de site en de bundel geven een 401 door het Netlify-wachtwoord. Gebruik de Netlify-koppeling in Code: `get-project` met site-id `6d6a1b8a-9c17-469c-9f23-c8911db814d6` geeft de huidige deploy-id, en `get-deploy-for-site` geeft daarvan `commit_ref` en `state`. Is `commit_ref` jouw commit en `state` ready, dan staat hij live. De onderhoudsregel en het versienummer zelf zie je alleen met het wachtwoord, dus in Edgars browser.

**Niet tegelijk op twee computers.** Git draait in een Google Drive-map; werk niet tegelijk op de MacBook en de HUB in deze map. Bij wisselen eerst `git pull`.

**Bestanden die alleen online staan.** Google Drive geeft soms een bestand met de juiste omvang en alleen nullen. `scripts/bouwmap.mjs` stopt dan met een melding. Lees zo'n bestand nooit als leeg en schrijf er niets overheen.

**Geleerd bij de overstap (1 november 2026).**
- Het site-id in `.netlify/state.json` was verkeerd; de echte site is gevonden met `get-projects` en zoekterm "autonomie".
- Live nakijken via `curl` en de ingebouwde browser loopt vast op het Netlify-wachtwoord. Het wachtwoord intikken doet Claude niet. Edgars Chrome (Claude in Chrome) is aangemeld op Netlify en, sinds de avond van 1 november 2026, ook op de site. Live nakijken: open de site in Chrome en lees versie en onderhoudsregel op het inlogscherm (`get_page_text`). Klik daar NIET op Inloggen: na het inloggen slaat de app de lijst automatisch op. Op die manier is v2.7 live gezien.
- autonomie.nhlstenden.com is ook vanuit Chrome op deze computer niet bereikbaar; alleen via het netwerk van NHL Stenden.
- Versie v2.5 is bewust gekozen: v2.2 tot en met v2.4 stonden al in de changelog. Vanaf nu gewoon 0.1 per release.
- `npm ci` meldt dat de install-scripts van esbuild en fsevents niet zijn goedgekeurd; de bouw werkt toch.
- Voorbeeld voor deze opzet: autonomie-frisius-mc (dezelfde codebasis) en Sepupu.

## Sleutels en wachtwoorden (alleen de plek)

- **Toegang via A.I. 2 Business (sinds v2.8, 2 november 2026).** Het
  Netlify-sitewachtwoord staat uit. `netlify/edge-functions/a2b-poort.js`
  bewaakt nu de hele site, ook de serverfuncties onder `/.netlify/functions/`
  (nagekeken: zonder pas 401). Binnen kom je alleen met een doorgangspas van de
  portfolio (https://edgar-van-gorkum.netlify.app, ingelogd); anders stuurt het
  slot naar het inlogscherm van de portfolio. Functies die de eigen site
  aanroepen, moeten het koekje van de bezoeker meesturen (het koekje heet nu
  `a2b`). Bron: `portfolio-ai2business/demo-slot/a2b-poort.js` en
  `ontwerp-doorgang-demos.md`.

- `APP_API_TOKEN`: omgevingsvariabele in Netlify. Alle functions controleren die tegen de header `x-api-token`. Zonder de variabele geven ze een 500.
- `VITE_API_TOKEN`: staat in Netlify (alle scopes, nagekeken op 1 november 2026, alleen de naam). De build zet hem in de bundel en de app stuurt hem mee. Zonder deze variabele valt de app terug op wat er op het inlogscherm is ingetikt; in de lokale proef is dat zo.
- **Geen eigen codes meer in de site (v2.7, 1 november 2026).** Het inlogscherm en het toegangsscherm van Beheer blijven staan, maar laten iedereen door, ook met een leeg veld. Reden: de site is een demo en staat in zijn geheel achter het sitewachtwoord van Netlify (`requiresPassword: true`). Besluit van Edgar. Niet terugzetten tenzij Edgar erom vraagt; in `src/App.jsx` staat bij beide plekken een opmerking. De oude codes waren constanten in de code, niet in de opslag; er staat dus niets ouds in Netlify Blobs.
- De oude codes staan nog in `.sleutels/autonomie-nhlstenden.env` (`NHL_TOEGANGSCODE`, `NHL_BEHEERPIN`) en in de git-geschiedenis. `npm run controle` houdt sinds v2.7 tegen dat ze in welk bestand dan ook terugkomen, ook in `src/App.jsx`.
- De servercontrole blijft: de functions eisen nog steeds `x-api-token` gelijk aan `APP_API_TOKEN`. Besluit van Edgar op 1 november 2026: laten staan. Niet weghalen tenzij Edgar erom vraagt.
- Het Netlify-sitewachtwoord staat nergens in deze map.
- Geen sleutels in git: `.env`, `.env.local`, `.netlify`, `.sleutels/` en `.overstap/` staan in .gitignore.

## Vaste afspraken

- Alle tekst in het Nederlands, in gewone taal, korte zinnen.
- Geen streepjes, gedachtestreepjes of em-dashes als stijlmiddel in Nederlandse tekst of in tekst in de app. Geen "A – B" of "kan - voor". Gebruik komma's of punten. Controleer dit bij elke oplevering; het teken U+2014 (em-dash) mag in geen enkel document voorkomen. Let op: de bestaande changelog bevat er nog veel (zie overdracht).
- Het thema heet Digitale Autonomie, niet Digitale Soevereiniteit. Geen "DNA" in publieke schermen.
- Methodische regels: B1 meet het aantal producten van één leverancier (geen exclusiviteit). De DICTU-schalen zijn omgekeerd (hoger is beter). Een leverancier met een moederbedrijf buiten de EU heeft jurisdictierisico, ook als de uitvoerende entiteit Nederlands of Europees is.
- In externe stukken DAAF en DICTU niet uitvoerig als eigen prestatie beschrijven: NHL Stenden volgt de Universiteit Utrecht.
- De tool is een prototype ("ter inspiratie en visualisatie"). Dat moet zichtbaar blijven (header, loginscherm, Over de tool, transparantiepagina).
- Huisstijl gelijk aan het AI-Netwerk: diepblauw verloop, NHL Stenden-logo, lettertype Inter, roze accenten (#E91E8C), golf onder de navigatie.
- Versienummer: bij elke release 0.1 omhoog in `VERSION` en een entry in `CHANGELOG`; package.json loopt mee.
- Onderhoudsregel "onderhoud: Claude Code · 2026" uit `src/onderhoudsregel.js`, in kleine letters onder de versiebadge op het inlogscherm (afspraak van 1 november 2026, voor al Edgars sites). Niet weghalen; de controle let erop.
- Wijzigingen gericht en met behoud van de bestaande structuur; geen grote herschrijving van App.jsx zonder overleg.
- Git-commits eindigen met de attributieregels die de omgeving voorschrijft.
- Edgar verwacht showcasekwaliteit, werkt het liefst autonoom, en wil een werkdag afsluiten zonder losse eindjes.
- Er wordt aan de SharePoint-variant (Autonomiescan) sinds 3 oktober 2026 niets meer gebouwd behalve foutherstel. Voor deze site is die afspraak niet vastgelegd; vraag het na voordat je veel gaat bouwen.

## Bekende zaken om op te letten

- De changelog onder v2.5 staat niet op volgorde (v2.0, v2.4, v2.3, v2.2, v2.1, v2.0, v1.9...) en v2.0 komt twee keer voor. Dat is geschiedenis en blijft zo; de echte volgorde is niet meer te achterhalen. De v2.0-entry zegt dat het loginscherm is verwijderd; het is later teruggekomen.
- Sinds v2.6 staat er geen gedachtestreepje meer in de code (199 vervangen door een komma, punt, dubbele punt of ·). De controle houdt dat zo.
- Het back-upschema is `0 8,11,16 * * *` UTC: in de zomertijd 10:00, 13:00 en 18:00, in de wintertijd een uur eerder. Back-ups worden na 7 dagen verwijderd. De data staat alleen in Netlify Blobs: maak voor grote wijzigingen eerst een export of force-backup.
- Open vragen uit de overstap: draait autonomie.nhlstenden.com dezelfde code, had de oude chat nog openstaande wensen, moet de interface ook Engels?
- `BACKUP 29 juni 2026/` bevat een JSON-export van de database en staat in .gitignore.

## Aandachtspunten

- De GitHub-repo is publiek. De oude inlogcode en beheerpincode staan sinds v2.7 niet meer in de code, maar nog wel in de git-geschiedenis. Besluit van Edgar op 1 november 2026: zo laten. Breng het pas opnieuw in als er gevoelige data in de tool komt.
- Het API-token zit via `VITE_API_TOKEN` in de JavaScript van de site. Wie langs het sitewachtwoord komt, kan het lezen. De servercontrole beschermt dus vooral tegen aanroepen van buiten die het sitewachtwoord niet hebben.
- De site heeft geen AI-hulp (geen aanroep naar api.anthropic.com).
