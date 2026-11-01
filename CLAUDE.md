# CLAUDE.md, Autonomie NHL Stenden

Geheugen van dit project. Lees dit eerst, daarna pas aan het werk.
Aangemaakt op 1 november 2026 bij de overstap van de chat naar Claude Code. Opgesteld op basis van de projectmap, de git-historie en Edgars geheugen, niet op basis van de oude chat zelf (zie OVERDRACHT-UIT-CHAT.md).

## Eerste handeling van elke sessie

```
cd "<pad naar CODE - Claude>/autonomie-nhlstenden"
git pull
grep -n 'const VERSION' src/App.jsx
git log --oneline | head -5
git status
```

Staat er geen node_modules (na de verhuizing naar Google Drive zijn die niet meegekomen), draai dan eerst `npm install`.

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
- Bouwen: `npm run build` (publish-map `dist`). Lokaal: `npm run dev`. Zie netlify.toml.
- Hosting: Netlify. Live adres: https://autonomie-nhlstenden.netlify.app
- Netlify site-id (uit .netlify/state.json): c225450e-a935-48e4-981a-187e8ffdd8e6. De sitenaam en het team in Netlify zijn niet bevestigd: een opvraging via de Netlify-koppeling gaf een 404, waarschijnlijk zit de site in een ander account dan die koppeling.
- GitHub: https://github.com/supergorkum/autonomie-nhlstenden, productietak `main`. Op 1 november 2026 is de werkmap schoon en gelijk aan origin/main (laatste commit 392dea6).
- Geplande back-up: `0 8,11,16 * * *` in UTC. In de zomer is dat 10:00, 13:00 en 18:00 Nederlandse tijd, in de winter een uur eerder.
- Plan: migratie naar eigen NHL Stenden-infrastructuur (NVIDIA Spark, Leeuwarden). De tekst op de transparantiepagina en de autonomiescores van het instrument zelf gaan daarvan uit. Nog niet uitgevoerd.

## Hoe een wijziging live ging

In de map staan geen patchscripts, geen deploy-script en geen .command-bestand. De git-historie bestaat uit korte Nederlandse commitberichten zonder prefix. Hoe de laatste releases daadwerkelijk live gingen (push naar `main` met automatische Netlify-build, of `netlify deploy` via de CLI) is niet vastgelegd. Zie OVERDRACHT-UIT-CHAT.md. Controleer dit in de eerste sessie en schrijf het hier op.

Stappen die bij een release horen:
1. Wijzig de code, bij voorkeur gericht en met behoud van de bestaande structuur.
2. Verhoog `VERSION` in `src/App.jsx` met 0.1 en voeg een blok toe aan `CHANGELOG`.
3. `npm run build` en controleer lokaal.
4. Commit met een korte Nederlandse boodschap en push naar `main`.
5. Controleer de live site.

## Sleutels en wachtwoorden (alleen de plek)

- `APP_API_TOKEN`: omgevingsvariabele in Netlify. Alle functions controleren die tegen de header `x-api-token`. Zonder de variabele geven ze een 500.
- `VITE_API_TOKEN`: optionele omgevingsvariabele bij de build; de app valt anders terug op wat de gebruiker op het loginscherm intikt.
- De inlogcode van het loginscherm staat hard in `src/App.jsx` (zoek op `LOGIN_CODE`, rond regel 2144) en dus in de publieke bundel en in de repository. De pincode voor de beheerstand staat ook in de code (zoek op `ADMIN_PIN`, rond regel 2231). Verplaats dit naar omgevingsvariabelen voordat er gevoelige data in komt.
- Geen sleutels in git: `.env`, `.env.local` en `.netlify` staan in .gitignore.

## Vaste afspraken

- Alle tekst in het Nederlands, in gewone taal, korte zinnen.
- Geen streepjes, gedachtestreepjes of em-dashes als stijlmiddel in Nederlandse tekst of in tekst in de app. Geen "A – B" of "kan - voor". Gebruik komma's of punten. Controleer dit bij elke oplevering; het teken — mag in geen enkel document voorkomen. Let op: de bestaande changelog bevat er nog veel (zie overdracht).
- Het thema heet Digitale Autonomie, niet Digitale Soevereiniteit. Geen "DNA" in publieke schermen.
- Methodische regels: B1 meet het aantal producten van één leverancier (geen exclusiviteit). De DICTU-schalen zijn omgekeerd (hoger is beter). Een leverancier met een moederbedrijf buiten de EU heeft jurisdictierisico, ook als de uitvoerende entiteit Nederlands of Europees is.
- In externe stukken DAAF en DICTU niet uitvoerig als eigen prestatie beschrijven: NHL Stenden volgt de Universiteit Utrecht.
- De tool is een prototype ("ter inspiratie en visualisatie"). Dat moet zichtbaar blijven (header, loginscherm, Over de tool, transparantiepagina).
- Huisstijl gelijk aan het AI-Netwerk: diepblauw verloop, NHL Stenden-logo, lettertype Inter, roze accenten (#E91E8C), golf onder de navigatie.
- Versienummer: bij elke release 0.1 omhoog in `VERSION` en een regel in `CHANGELOG`.
- Wijzigingen gericht en met behoud van de bestaande structuur; geen grote herschrijving van App.jsx zonder overleg.
- Git-commits eindigen met de attributieregels die de omgeving voorschrijft.
- Edgar verwacht showcasekwaliteit, werkt het liefst autonoom, en wil een werkdag afsluiten zonder losse eindjes.
- Er wordt aan de SharePoint-variant (Autonomiescan) sinds 3 oktober 2026 niets meer gebouwd behalve foutherstel. Voor deze site is die afspraak niet vastgelegd; vraag het na voordat je veel gaat bouwen.

## Bekende zaken om op te letten

- Versienummers kloppen niet onderling: `VERSION` is v2.1, package.json zegt 2.0.0, en de changelog staat niet op volgorde (bovenaan v2.0 met "Huidige versie", daaronder v2.4, v2.3, v2.2, v2.1). Zie overdracht.
- `src/files.zip` is een oud archief (App_met_beheer.jsx en main.jsx van 26 mei). Niet gebruiken.
- `BACKUP 29 juni 2026/` bevat een JSON-export van de database en staat in .gitignore.
