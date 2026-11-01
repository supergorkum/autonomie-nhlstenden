// DE CONTROLE VOOR ELKE PUBLICATIE (1 november 2026).
//
// Sinds de site in Claude Code wordt onderhouden, gaat elke wijziging via git,
// en een push naar main publiceert de site (Netlify bouwt zelf bij elke push
// naar main). Dit script is de poort: het draait vanzelf bij elke git push
// (scripts/git-hooks/pre-push) en met de hand met npm run controle.
//
// Wat het nakijkt, in deze volgorde:
// 1. VERSION en de bovenste CHANGELOG-entry in src/App.jsx horen bij elkaar,
//    het versienummer in package.json loopt mee, en de onderhoudsregel staat
//    op het inlogscherm;
// 2. bij een push: raakt de push src/, netlify/, public/ of index.html, dan
//    hoort VERSION in src/App.jsx mee te veranderen;
// 3. geen codes uit .sleutels/ in een bestand dat git meeneemt, ook niet in
//    src/App.jsx (sinds v2.7 heeft de app geen eigen codes meer; de site staat
//    achter het sitewachtwoord van Netlify), en nergens een gedachtestreepje
//    (U+2014), afspraak van Edgar;
// 4. de bouwmap buiten Google Drive wordt bijgewerkt (scripts/bouwmap.mjs);
// 5. vite build slaagt, en de onderhoudsregel staat in de bundel.
//
// Deze site heeft geen eslint en geen tests. Dat kan dit script dus niet
// controleren; de bouw vangt alleen syntaxfouten en kapotte imports op.

import { execFileSync, spawnSync } from "node:child_process";
import fs from "node:fs";
import path from "node:path";
import { REPO, BOUW, synchroniseer, afhankelijkheden } from "./bouwmap.mjs";

const PUSH = process.argv.includes("--push");
const SNEL = process.argv.includes("--snel");
const fouten = [];
const meld = (t) => console.log("  " + t);
const lees = (rel) => fs.readFileSync(path.join(REPO, rel), "utf8");
const git = (...a) => execFileSync("git", a, { cwd: REPO, maxBuffer: 64 * 1024 * 1024 }).toString();

const versieUit = (bron) => (bron.match(/const VERSION = "v(\d+\.\d+)";/) || [])[1];

function inhoud() {
  console.log("1. Versie, changelog en onderhoudsregel");
  const app = lees("src/App.jsx");
  const versie = versieUit(app);
  if (!versie) return fouten.push('const VERSION = "vX.Y" niet gevonden in src/App.jsx');
  const start = app.indexOf("const CHANGELOG = [");
  const blok = app.slice(start, app.indexOf("\n];", start));
  const entries = [...blok.matchAll(/versie:\s*"v(\d+\.\d+)",\s*datum:\s*"([^"]*)"/g)];
  if (!entries.length) return fouten.push("geen entries gevonden in CHANGELOG");
  if (entries[0][1] !== versie) fouten.push("VERSION zegt " + versie + ", de bovenste CHANGELOG-entry v" + entries[0][1]);
  // Alleen de nieuwste entry: v2.0 staat er van vroeger twee keer in. Dat is
  // geschiedenis en blijft zo.
  if (entries.slice(1).some((e) => e[1] === entries[0][1])) fouten.push("versie " + entries[0][1] + " staat al eerder in CHANGELOG");
  const pkg = JSON.parse(lees("package.json")).version;
  if (pkg !== versie + ".0") fouten.push("package.json zegt versie " + pkg + ", hoort " + versie + ".0 te zijn");
  if (!fs.existsSync(path.join(REPO, "src/onderhoudsregel.js"))) fouten.push("src/onderhoudsregel.js ontbreekt");
  else if (!/\{ONDERHOUDSREGEL\}/.test(app)) fouten.push("de onderhoudsregel staat niet meer op het inlogscherm (src/App.jsx)");
  meld("versie " + versie + " (" + entries[0][2] + ")");
  return versie;
}

function pushInhoud(versie) {
  console.log("2. Wat deze push meeneemt");
  let basis;
  try {
    basis = git("rev-parse", "--verify", "-q", "origin/main").trim();
  } catch {
    meld("origin/main nog onbekend, stap overgeslagen");
    return;
  }
  const namen = git("diff", "--name-only", basis + "..HEAD").split("\n").filter(Boolean);
  if (!namen.length) return meld("niets nieuws ten opzichte van origin/main");
  meld(namen.length + " bestand(en) gewijzigd");
  const zichtbaar = namen.filter((n) => /^(src|netlify|public)\//.test(n) || n === "index.html");
  if (!zichtbaar.length) return;
  const oud = versieUit(git("show", basis + ":src/App.jsx"));
  if (oud === versie) {
    fouten.push("deze push wijzigt de site (" + zichtbaar.slice(0, 3).join(", ") + (zichtbaar.length > 3 ? ", ..." : "") + ") maar VERSION staat nog op " + versie + ". Verhoog met 0.1 en zet een entry in CHANGELOG.");
  }
}

function sleutels() {
  console.log("3. Geen codes uit .sleutels/ en geen gedachtestreepjes in wat git meeneemt");
  const bestand = path.join(REPO, ".sleutels/autonomie-nhlstenden.env");
  if (!fs.existsSync(bestand)) meld(".sleutels/autonomie-nhlstenden.env ontbreekt op deze computer, codes niet nagekeken");
  const waarden = (fs.existsSync(bestand) ? fs.readFileSync(bestand, "utf8") : "").split("\n")
    .map((r) => r.match(/^[A-Z_]+=(.+)$/))
    .filter(Boolean)
    .map((m) => m[1].trim())
    .filter((w) => w.length >= 6);
  const bestanden = git("ls-files", "-co", "--exclude-standard", "-z").split("\0").filter(Boolean);
  let gevonden = 0;
  for (const rel of bestanden) {
    if (!/\.(m?jsx?|json|css|html|md|toml|txt|svg)$/i.test(rel) && !rel.includes("git-hooks/")) continue;
    let tekst;
    try {
      tekst = fs.readFileSync(path.join(REPO, rel), "utf8");
    } catch {
      continue;
    }
    if (tekst.includes("\u2014")) {
      const regel = tekst.slice(0, tekst.indexOf("\u2014")).split("\n").length;
      fouten.push(rel + ":" + regel + " bevat een gedachtestreepje. Gebruik een komma, punt of dubbele punt.");
    }
    if (waarden.some((w) => tekst.includes(w))) {
      fouten.push(rel + " bevat een code uit .sleutels/. Haal hem eruit voor de commit.");
      gevonden++;
    }
  }
  if (!gevonden) meld("in orde (" + bestanden.length + " bestanden nagekeken)");
}

function bouw() {
  console.log("5. vite build");
  const r = spawnSync(process.execPath, ["node_modules/vite/bin/vite.js", "build", "--logLevel", "error"], { cwd: BOUW, maxBuffer: 256 * 1024 * 1024 });
  if (r.status !== 0) {
    const uit = (r.stdout.toString() + r.stderr.toString()).trim().split("\n").slice(-15).join("\n      ");
    return fouten.push("vite build is mislukt:\n      " + uit);
  }
  const regel = (lees("src/onderhoudsregel.js").match(/ONDERHOUDSREGEL = "([^"]+)"/) || [])[1];
  const assets = path.join(BOUW, "dist/assets");
  const bundel = fs.readdirSync(assets).filter((f) => f.endsWith(".js")).map((f) => fs.readFileSync(path.join(assets, f), "utf8")).join("\n");
  if (regel && !bundel.includes(regel)) fouten.push("de onderhoudsregel staat niet in de gebouwde bundel");
  meld("gelukt, uitvoer in " + path.join(BOUW, "dist"));
}

console.log("Controle autonomie-nhlstenden" + (PUSH ? " (voor de push)" : ""));
try {
  const versie = inhoud();
  if (PUSH && versie) pushInhoud(versie);
  sleutels();
  if (!SNEL) {
    console.log("4. Bouwmap bijwerken");
    synchroniseer();
    afhankelijkheden();
    bouw();
  }
} catch (e) {
  fouten.push(e.anhl ? e.message : "onverwachte fout: " + (e.stack || e.message || e));
}

if (fouten.length) {
  console.log("");
  console.log("NIET IN ORDE:");
  for (const f of fouten) console.log("  - " + f);
  if (PUSH) console.log("\nDe push is tegengehouden. Er is niets gepubliceerd.");
  process.exit(1);
}
console.log("\nAlles in orde." + (PUSH ? " De push gaat door; Netlify publiceert zodra de bouw daar klaar is." : ""));
