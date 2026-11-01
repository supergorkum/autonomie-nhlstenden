// DE LOKALE PROEF (1 november 2026).
//
// Start de site op http://127.0.0.1:5173 vanuit de bouwmap buiten Google
// Drive, en houdt die bouwmap bij terwijl je in de projectmap werkt: elke
// opgeslagen wijziging staat binnen een seconde in de proef.
//
// LET OP: aanroepen naar /api gaan via een proxy naar de LIVE site
// (vite.config.js, alleen als ANHL_LIVE_PROXY aan staat). Zolang de live site
// achter het Netlify-wachtwoord zit (zo op 1 november 2026), komt daar een
// 401 terug en laadt de proef geen gegevens: het inlogscherm en de opmaak zijn
// wel te bekijken. Gaat het wachtwoord er ooit af, dan verandert opslaan of
// een actie in Beheer de echte gegevens. Kijken mag; opslaan, terugzetten en
// verwijderen alleen met opzet.

import { spawn } from "node:child_process";
import fs from "node:fs";
import path from "node:path";
import { REPO, BOUW, synchroniseer, afhankelijkheden } from "./bouwmap.mjs";

const POORT = process.env.PORT || "5173";

try {
  synchroniseer();
  afhankelijkheden();
} catch (e) {
  console.error("GESTOPT: " + e.message);
  process.exit(1);
}

const vite = spawn(process.execPath, ["node_modules/vite/bin/vite.js", "--host", "127.0.0.1", "--port", POORT, "--strictPort"], {
  cwd: BOUW,
  stdio: "inherit",
  env: { ...process.env, ANHL_LIVE_PROXY: "1" },
});
vite.on("exit", (code) => process.exit(code ?? 0));

// Bijhouden. Een korte pauze vangt een reeks opslagacties in een keer op.
let klok = null;
const OVERSLAAN = /(^|\/)(\.git|node_modules|dist|\.netlify|\.npm-cache|\.sleutels|\.overstap)(\/|$)/;
try {
  fs.watch(REPO, { recursive: true }, (_soort, naam) => {
    if (naam && OVERSLAAN.test(String(naam).split(path.sep).join("/"))) return;
    clearTimeout(klok);
    klok = setTimeout(() => {
      try {
        const r = synchroniseer({ stil: true });
        if (r.gekopieerd || r.verwijderd) console.log("[anhl] proef bijgewerkt: " + r.gekopieerd + " bestand(en)");
        if (afhankelijkheden()) console.log("[anhl] afhankelijkheden opnieuw geinstalleerd; herstart de proef als iets niet laadt");
      } catch (e) {
        console.error("[anhl] bijwerken mislukt: " + e.message);
      }
    }, 300);
  });
} catch (e) {
  console.error("[anhl] meekijken met de projectmap lukt niet (" + e.message + "); draai na een wijziging npm run bouwmap.");
}

for (const sein of ["SIGINT", "SIGTERM"]) process.on(sein, () => vite.kill(sein));
