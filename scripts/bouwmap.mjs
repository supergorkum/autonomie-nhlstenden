// DE BOUWMAP: WAAR NODE_MODULES WEL MAG STAAN (1 november 2026).
//
// De projectmap staat in Google Drive. Daar hoort node_modules niet in:
// tienduizenden kleine bestanden die Drive allemaal wil uploaden, en die op
// elke computer anders zijn. Na de verhuizing van 28 oktober 2026 waren de
// startbestanden in node_modules/.bin daardoor stuk.
//
// Dus bouwen, de lokale proef en de publicatie gebeuren in een kopie buiten
// Drive, standaard ~/Developer/autonomie-nhlstenden-build. Dit bestand houdt
// die kopie gelijk aan de projectmap: het neemt precies de bestanden over die
// git kent of zou kennen (alles wat niet in .gitignore staat), haalt weg wat
// in de projectmap niet meer bestaat, en draait npm ci zodra
// package-lock.json anders is dan bij de vorige keer. npm gebruikt daarbij
// een eigen cache in de bouwmap, omdat ~/.npm op Edgars MacBook bestanden van
// root bevat (EACCES).
//
// Werken doe je altijd in de projectmap. De bouwmap is wegwerp: weggooien
// mag altijd, de volgende ronde bouwt hem opnieuw op.

import { execFileSync, spawnSync } from "node:child_process";
import fs from "node:fs";
import path from "node:path";
import os from "node:os";
import crypto from "node:crypto";
import { fileURLToPath } from "node:url";

export const REPO = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
export const BOUW = path.resolve(process.env.ANHL_BOUWMAP || path.join(os.homedir(), "Developer", "autonomie-nhlstenden-build"));

const MERKTEKEN = ".anhl-bouwmap";
const LOCKHASH = ".anhl-lock-hash";
// Wat in de bouwmap thuishoort en nooit uit de projectmap komt.
const EIGEN = new Set(["node_modules", "dist", MERKTEKEN, LOCKHASH, ".vite", ".npm-cache", ".netlify"]);
// Bestanden waarvan we na het kopieren nakijken of er echt iets in staat.
// Google Drive geeft voor een bestand dat alleen online staat soms geen fout
// maar een bestand met de juiste omvang en alleen nullen erin.
const TEKST = /\.(m?jsx?|json|css|html|md|toml|txt|svg|py|command)$/i;

function fout(t) {
  const e = new Error(t);
  e.anhl = true;
  throw e;
}

function veiligheid() {
  const rel = path.relative(REPO, BOUW);
  if (!rel || !rel.startsWith("..")) fout("De bouwmap ligt in of op de projectmap (" + BOUW + "). Dat mag niet.");
  if (/CloudStorage|Google ?Drive|Mijn Drive|My Drive/i.test(BOUW)) fout("De bouwmap ligt in een cloudmap (" + BOUW + "). Kies een gewone map, bijvoorbeeld ~/Developer/autonomie-nhlstenden-build.");
  if (fs.existsSync(BOUW)) {
    const inhoud = fs.readdirSync(BOUW);
    if (inhoud.length && !inhoud.includes(MERKTEKEN)) {
      fout("De map " + BOUW + " bestaat al en is niet door dit script gemaakt. Ik raak hem niet aan. Kies een andere map met ANHL_BOUWMAP.");
    }
  } else {
    fs.mkdirSync(BOUW, { recursive: true });
  }
  fs.writeFileSync(path.join(BOUW, MERKTEKEN), "Kopie van de projectmap autonomie-nhlstenden om in te bouwen. Weggooien mag altijd.\n");
}

function lijst() {
  const uit = execFileSync("git", ["ls-files", "-co", "--exclude-standard", "-z"], { cwd: REPO, maxBuffer: 64 * 1024 * 1024 });
  return uit
    .toString("utf8")
    .split("\0")
    .filter(Boolean)
    .filter((p) => !p.split("/").some((d) => EIGEN.has(d)));
}

function alleenNullen(buf) {
  if (!buf.length) return false;
  for (let i = 0; i < buf.length; i++) if (buf[i] !== 0) return false;
  return true;
}

function doorloop(map, rel, uit) {
  for (const naam of fs.readdirSync(map, { withFileTypes: true })) {
    if (!rel && EIGEN.has(naam.name)) continue;
    const r = rel ? rel + "/" + naam.name : naam.name;
    const vol = path.join(map, naam.name);
    if (naam.isDirectory()) doorloop(vol, r, uit);
    else uit.push(r);
  }
  return uit;
}

export function synchroniseer({ stil = false } = {}) {
  veiligheid();
  const bestanden = lijst();
  const bekend = new Set();
  let gekopieerd = 0;
  for (const rel of bestanden) {
    const bron = path.join(REPO, rel);
    let st;
    try {
      st = fs.statSync(bron);
    } catch {
      continue; // in git maar weggehaald in de werkmap
    }
    if (!st.isFile()) continue;
    bekend.add(rel);
    const doel = path.join(BOUW, rel);
    let gelijk = false;
    try {
      const d = fs.statSync(doel);
      gelijk = d.size === st.size && Math.abs(d.mtimeMs - st.mtimeMs) < 1;
    } catch {
      gelijk = false;
    }
    if (gelijk) continue;
    const buf = fs.readFileSync(bron);
    if (TEKST.test(rel) && alleenNullen(buf)) {
      fout("Het bestand " + rel + " bestaat alleen uit nullen. Google Drive heeft het waarschijnlijk nog niet echt binnengehaald. Open de projectmap in Finder, wacht tot Drive klaar is, en probeer opnieuw.");
    }
    fs.mkdirSync(path.dirname(doel), { recursive: true });
    fs.writeFileSync(doel, buf);
    fs.utimesSync(doel, st.atime, st.mtime);
    gekopieerd++;
  }
  let verwijderd = 0;
  for (const rel of doorloop(BOUW, "", [])) {
    if (!bekend.has(rel)) {
      fs.unlinkSync(path.join(BOUW, rel));
      verwijderd++;
    }
  }
  if (!stil) console.log("Bouwmap " + BOUW + ": " + gekopieerd + " bijgewerkt, " + verwijderd + " weggehaald, " + bekend.size + " bestanden in totaal.");
  return { gekopieerd, verwijderd, totaal: bekend.size };
}

export function afhankelijkheden() {
  const lock = path.join(BOUW, "package-lock.json");
  if (!fs.existsSync(lock)) fout("package-lock.json ontbreekt in de bouwmap.");
  const hash = crypto.createHash("sha256").update(fs.readFileSync(lock)).update(process.platform + process.arch).digest("hex");
  const opgeslagen = path.join(BOUW, LOCKHASH);
  const klaar = fs.existsSync(path.join(BOUW, "node_modules")) && fs.existsSync(opgeslagen) && fs.readFileSync(opgeslagen, "utf8") === hash;
  if (klaar) return false;
  console.log("npm ci in de bouwmap (package-lock.json is nieuw of veranderd)...");
  const npm = process.platform === "win32" ? "npm.cmd" : "npm";
  const r = spawnSync(npm, ["ci", "--no-audit", "--no-fund"], {
    cwd: BOUW,
    stdio: "inherit",
    env: { ...process.env, npm_config_cache: path.join(BOUW, ".npm-cache") },
  });
  if (r.status !== 0) fout("npm ci is mislukt in " + BOUW + ". Zie de meldingen hierboven.");
  fs.writeFileSync(opgeslagen, hash);
  return true;
}

if (process.argv[1] && fileURLToPath(import.meta.url) === path.resolve(process.argv[1])) {
  try {
    synchroniseer();
    if (!process.argv.includes("--zonder-npm")) afhankelijkheden();
  } catch (e) {
    console.error("GESTOPT: " + e.message);
    process.exit(1);
  }
}
