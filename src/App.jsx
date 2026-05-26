import { useState, useEffect } from "react";
import * as XLSX from "xlsx";
import {
  RadarChart, Radar, PolarGrid, PolarAngleAxis, PolarRadiusAxis,
  ResponsiveContainer, BarChart, Bar, XAxis, YAxis, CartesianGrid,
  Tooltip, Legend, ScatterChart, Scatter, Cell
} from "recharts";

// ──────────────────────────────────────────────────────────────
// FRAMEWORK DATA
// ──────────────────────────────────────────────────────────────

const DAAF = [
  // ── Niveau 1: Risico-exposure ──────────────────────────────
  { key:"A1", dim:"A", level:"Risico", dimName:"Geopolitiek risico",
    name:"Jurisdictie leverancier",
    hint:"Score 1 = weinig risico · Score 5 = hoog risico",
    question:"Onder welke jurisdictie valt de leverancier en waar staat de data? Geopolitieke en jurisdictierisico's.",
    toelichting:"Waar is de leverancier gevestigd? Extraterritoriale wetgeving? Adequaatheidsbesluit? Let op: het EU-US Data Privacy Framework (2023) is het derde adequaatheidsbesluit voor de VS; de vorige twee zijn door het Europese Hof vernietigd. De onderliggende wetgeving (CLOUD Act, FISA 702) is echter niet veranderd. Weeg zelf of je dit als score 3 of 4 beschouwt.",
    scores:[
      {s:1,label:"EU/EER volledig",    desc:"EU/EER-jurisdictie. Geen extraterritoriale claims. Volledige EU-bescherming."},
      {s:2,label:"EU/EER beperkt",     desc:"EU/EER met beperkte extraterritoriale claims. CLOUD Act n.v.t."},
      {s:3,label:"Adequaatheid + risico",desc:"Adequaatheidsbesluit, maar extraterritoriale wetgeving (CLOUD Act, FISA 702) geeft buitenlandse overheid potentieel toegang."},
      {s:4,label:"SCCs, geen adequaat",desc:"Geen adequaatheidsbesluit maar contractuele waarborgen (SCCs). Juridische bescherming beperkt."},
      {s:5,label:"Geen waarborgen",    desc:"Geen adequaatheidsbesluit, geen waarborgen. Directe toegang buitenlandse overheden."}
    ]
  },
  { key:"A3", dim:"A", level:"Risico", dimName:"Geopolitiek risico",
    name:"Hosting en datalocatie",
    hint:"Score 1 = weinig risico · Score 5 = hoog risico",
    question:"Waar worden data en applicatie fysiek gehost en hoe is dat geborgd?",
    toelichting:"Waar staan de servers fysiek? Is de data- en applicatielocatie contractueel vastgelegd? Kunnen backups of replicatie buiten de EU terechtkomen? Let op: data en applicatie kunnen op verschillende locaties staan (bijv. applicatie in Ierland, database in Frankfurt, beheerconsole in de VS). Check of er ergens in de keten een niet-Europese jurisdictie meespeelt en neem de slechtste jurisdictie als uitgangspunt.",
    scores:[
      {s:1,label:"EU/EER contractueel", desc:"Data uitsluitend in EU/EER. Contractueel vastgelegd."},
      {s:2,label:"EU/EER + adequaat",   desc:"Data in EU/EER. Backups of replicatie mogelijk in land met adequaatheidsbesluit."},
      {s:3,label:"EU/EER, geen garantie",desc:"Data in EU/EER, maar geen contractuele garantie over locatie. Kan wijzigen."},
      {s:4,label:"Deels buiten EU",     desc:"Data deels buiten EU, met contractuele waarborgen (SCCs of adequaatheidsbesluit)."},
      {s:5,label:"Buiten EU",           desc:"Data buiten EU, geen waarborgen, of onduidelijk waar data staat."}
    ]
  },
  { key:"B1", dim:"B", level:"Risico", dimName:"Leveranciersafhankelijkheid",
    name:"Vendor concentratie",
    hint:"Score 1 = weinig risico · Score 5 = hoog risico",
    question:"Hoeveel producten en diensten neem je af bij dezelfde leverancier?",
    toelichting:"Tel alle producten, diensten, licenties en platformen die je van deze leverancier gebruikt, zowel afgenomen diensten als gekochte software. Denk ook aan onderliggende platformen (bijv. Azure AD, SharePoint, Intune naast Office 365).",
    scores:[
      {s:1,label:"1 product",          desc:"1 product of dienst bij deze leverancier."},
      {s:2,label:"2–3 producten",       desc:"2–3 producten/diensten bij deze leverancier."},
      {s:3,label:"4–6 producten",       desc:"4–6 producten/diensten bij deze leverancier."},
      {s:4,label:"7–15 producten",      desc:"7–15 producten/diensten bij deze leverancier."},
      {s:5,label:"Heel ecosysteem",     desc:"Heel ecosysteem (>15 producten/diensten) bij deze leverancier."}
    ]
  },
  { key:"C1", dim:"C", level:"Mitigatie", dimName:"Technische weerbaarheid",
    name:"Alternatief beschikbaar", hint:"Score 1 = zwakke mitigatie · Score 5 = sterke mitigatie",
    question:"Is er een realistisch en praktisch inzetbaar alternatief voor deze applicatie?",
    scores:[
      {s:1,label:"Geen alternatief",    desc:"Geen alternatief beschikbaar of realistisch denkbaar."},
      {s:2,label:"Theoretisch",         desc:"Theoretisch alternatief aanwezig, niet praktisch inzetbaar."},
      {s:3,label:"Met inspanning",      desc:"Alternatief beschikbaar, maar migratie is complex en tijdrovend."},
      {s:4,label:"Goed alternatief",    desc:"Goed alternatief beschikbaar. Migratie is haalbaar."},
      {s:5,label:"Meerdere alt.",       desc:"Meerdere volwaardige alternatieven. Eenvoudige en snelle migratie mogelijk."}
    ]
  },
  { key:"D1", dim:"D", level:"Mitigatie", dimName:"Kennisweerbaarheid",
    name:"Interne kennis & exitstrategie", hint:"Score 1 = zwakke mitigatie · Score 5 = sterke mitigatie",
    question:"Is er voldoende interne kennis aanwezig en een concrete exitstrategie?",
    scores:[
      {s:1,label:"Geen",         desc:"Geen interne kennis. Geen exitstrategie aanwezig."},
      {s:2,label:"Minimaal",     desc:"Minimale kennis aanwezig. Globale exitstrategie beschikbaar."},
      {s:3,label:"Basis",        desc:"Basiskennis aanwezig. Exitstrategie in ontwikkeling."},
      {s:4,label:"Goed",         desc:"Goede interne kennis. Exitstrategie gedocumenteerd en actueel."},
      {s:5,label:"Uitgebreid",   desc:"Uitgebreide expertise. Exitstrategie getest en periodiek bijgewerkt."}
    ]
  },
  { key:"E1", dim:"E", level:"Mitigatie", dimName:"Contractuele weerbaarheid",
    name:"Contractuele bescherming", hint:"Score 1 = zwakke mitigatie · Score 5 = sterke mitigatie",
    question:"In hoeverre biedt het contract bescherming bij geopolitieke risico's en onwenselijke toegang?",
    scores:[
      {s:1,label:"Geen",          desc:"Geen contractuele bescherming. Aanbieder bepaalt alles."},
      {s:2,label:"Basis",         desc:"Standaard contract. Geen autonomie- of databeschermingsbepalingen."},
      {s:3,label:"Beperkt",       desc:"Contract met enige beschermingsclausules. Beperkte afdwingbaarheid."},
      {s:4,label:"Sterk",         desc:"Sterk contract met EU-rechtskeuze, databepalingen en meldplicht."},
      {s:5,label:"Maximaal",      desc:"Maximale bescherming incl. escrow, actief juridisch verzet en MLAT."}
    ]
  },
  { key:"F1", dim:"F", level:"Belang", dimName:"Operationeel belang",
    name:"Kriticiteit primair proces", hint:"Score 1 = minder urgent · Score 5 = hoogste urgentie",
    question:"Hoe kritiek is de applicatie voor het primaire onderwijs- of onderzoeksproces van NHL Stenden?",
    scores:[
      {s:1,label:"Ondersteunend",   desc:"Uitval heeft minimale impact. Makkelijk op te vangen."},
      {s:2,label:"Beperkt",         desc:"Tijdelijke vervanging of workaround eenvoudig mogelijk."},
      {s:3,label:"Relevant",        desc:"Uitval leidt tot merkbare verstoring van processen."},
      {s:4,label:"Kritiek",         desc:"Uitval verstoort het primaire proces significant."},
      {s:5,label:"Mission-critical",desc:"Uitval stopt primaire processen direct. Geen alternatief."}
    ]
  },
  { key:"G1", dim:"G", level:"Belang", dimName:"Data-kritikaliteit",
    name:"Gevoeligheid verwerkte data", hint:"Score 1 = minder urgent · Score 5 = hoogste urgentie",
    question:"Hoe gevoelig of strategisch waardevol zijn de gegevens die de applicatie verwerkt?",
    scores:[
      {s:1,label:"Openbaar",     desc:"Openbare of anonieme data. Geen privacyrisico of strategische waarde."},
      {s:2,label:"Intern",       desc:"Interne data. Beperkt privacygevoelig. Geen bijzondere persoonsgegevens."},
      {s:3,label:"Gevoelig",     desc:"Gevoelige persoonsgegevens. AVG-relevant."},
      {s:4,label:"Bijzonder",    desc:"Bijzondere persoonsgegevens of strategisch waardevolle bedrijfsdata."},
      {s:5,label:"Hoogst gev.",  desc:"Hoogst gevoelige data. Inbreuk heeft ernstige gevolgen."}
    ]
  },
  { key:"H1", dim:"H", level:"Belang", dimName:"Strategisch belang",
    name:"Bijdrage aan kernmissie", hint:"Score 1 = minder urgent · Score 5 = hoogste urgentie",
    question:"In hoeverre is de applicatie essentieel voor de strategische doelen van NHL Stenden?",
    scores:[
      {s:1,label:"Geen bijdrage", desc:"Geen directe bijdrage aan strategische doelen van NHL Stenden."},
      {s:2,label:"Marginaal",     desc:"Marginale bijdrage. Makkelijk vervangbaar zonder strategisch verlies."},
      {s:3,label:"Relevant",      desc:"Relevante bijdrage aan doelen. Niet uniek maar waardevol."},
      {s:4,label:"Significant",   desc:"Significante bijdrage aan realisatie van strategische doelen."},
      {s:5,label:"Onmisbaar",     desc:"Onmisbaar voor realisatie van de kernmissie van NHL Stenden."}
    ]
  }
];

const DICTU = [
  { key:"2.1", cat:"Data & AI", name:"Data residency",
    question:"Worden alle klantdata (incl. back-ups) uitsluitend opgeslagen en verwerkt binnen het Europese grondgebied (EU, EER, EFTA)?",
    norm:"Alle data moeten fysiek en logisch binnen Europa blijven. Geen enkele vorm van overdracht buiten EU is toegestaan.",
    scores:[
      {s:1,label:"Geen EU-garantie",     desc:"Geen garantie dat data, back-ups, metadata of logs binnen de EU worden opgeslagen of verwerkt."},
      {s:2,label:"Gedeeltelijk EU",       desc:"Hoofddata in EU, maar back-ups, metadata of logs kunnen buiten EU staan of verwerkt worden."},
      {s:3,label:"EU + uitzonderingen",   desc:"Hoofddata en back-ups in EU, maar systeemdata (metadata, logs) kan buiten EU. Analytics/training mogelijk buiten EU."},
      {s:4,label:"Volledig EU",           desc:"Alle data (incl. back-ups, metadata, logs) in EU. Geen datatransfer buiten EU, ook niet voor analytics/training."},
      {s:5,label:"Lokaal/EU gegarandeerd",desc:"Alle data in EU + optie voor lokale (on-premise) opslag. Geen enkele datatransfer buiten EU."}
    ]
  },
  { key:"2.2", cat:"Data & AI", name:"Technische Toegangsbeveiliging",
    question:"Zijn er aantoonbare technische maatregelen die ervoor zorgen dat ongecodeerde klantdata voor niemand toegankelijk is, inclusief de aanbieder zelf?",
    norm:"Juridische beloftes zijn niet voldoende. Er moeten verifieerbare technische garanties zijn (bijv. confidential computing of klant-exclusief sleutelbeheer).",
    scores:[
      {s:1,label:"Geen bescherming",     desc:"Aanbieder heeft volledige toegang tot data. Geen encryptie of sleutelbeheer door klant."},
      {s:2,label:"Platform Managed Keys",desc:"Data versleuteld maar aanbieder beheert alle sleutels en kan toegang verkrijgen."},
      {s:3,label:"BYOK / CMK",           desc:"Klant levert of beheert sleutels, maar aanbieder kan rechten aanpassen via het cloudplatform."},
      {s:4,label:"External HSM",         desc:"Sleutelbeheer volledig onder controle van klant via externe HSM. Aanbieder heeft geen/beperkte toegang."},
      {s:5,label:"HYOK + Confidential",  desc:"Klant houdt sleutels. Data cryptografisch geïsoleerd zelfs tijdens verwerking. Aanbieder nooit toegang tot ongecodeerde data."}
    ]
  },
  { key:"2.3", cat:"Data & AI", name:"Juridische Toegangsbeveiliging",
    question:"Is de aanbieder contractueel verplicht om niet-Europese verzoeken tot datatoegang juridisch aan te vechten en de klant hierover te informeren?",
    norm:"De aanbieder moet optreden als juridisch beschermingsmechanisme voor de klant en actief verzet bieden tegen niet-EU dataverzoeken, inclusief verplichte melding.",
    scores:[
      {s:1,label:"Geen verplichting",       desc:"Aanbieder voldoet aan buitenlandse verzoeken zonder verzet."},
      {s:2,label:"Vrijwillige melding",      desc:"Aanbieder informeert klant vrijwillig of op aanvraag, maar vecht verzoeken niet aan."},
      {s:3,label:"Contractuele melding",     desc:"Verplichte melding van verzoeken, maar geen juridisch verzet."},
      {s:4,label:"Actief juridisch verzet",  desc:"Verplichte melding + weigering + altijd doorverwijzing naar klant."},
      {s:5,label:"Volledige bescherming",    desc:"Melding + weigering + doorverwijzing naar internationale mechanismen zoals MLAT."}
    ]
  },
  { key:"4.1", cat:"EU-Infrastructuur", name:"EU-Infrastructuur en Control Plane",
    question:"Bevindt de volledige technische infrastructuur (datacenters, netwerk) en control plane zich fysiek binnen de EU/EER/EFTA?",
    norm:"De gehele fysieke infrastructuur en de control plane om Cloud services te managen en orkestreren moeten zich bevinden en uitgevoerd worden binnen de EU/EER/EFTA.",
    scores:[
      {s:1,label:"Volledig non-EU",       desc:"Infrastructuur en control plane volledig buiten EU. Beheer door niet-Europese partijen."},
      {s:2,label:"EU-infra, non-EU CP",   desc:"Fysieke infra in EU, maar control plane extern beheerd buiten EU."},
      {s:3,label:"EU + deels extern CP",  desc:"Infra en control plane in EU, maar control plane (deels) extern beheerd."},
      {s:4,label:"Volledig EU",           desc:"Infra en control plane gegarandeerd volledig in EU."},
      {s:5,label:"EU + audit",            desc:"Volledig EU-operatie gegarandeerd + periodieke audit door erkende EU-partner."}
    ]
  }
];

// ──────────────────────────────────────────────────────────────
// UTILITIES
// ──────────────────────────────────────────────────────────────

function scoreColor(s, max = 10) {
  if (!s) return "#9ca3af";
  const p = s / max;
  if (p >= 0.7) return "#16a34a";
  if (p >= 0.5) return "#ca8a04";
  if (p >= 0.3) return "#ea580c";
  return "#dc2626";
}

function scoreLabel(s) {
  if (!s) return { text: "Onvolledig", bg: "#f3f4f6", fg: "#6b7280" };
  if (s >= 7) return { text: "Goed",        bg: "#dcfce7", fg: "#15803d" };
  if (s >= 5) return { text: "Acceptabel",  bg: "#fef9c3", fg: "#a16207" };
  if (s >= 3) return { text: "Zorgwekkend", bg: "#ffedd5", fg: "#c2410c" };
  return            { text: "Kritiek",       bg: "#fee2e2", fg: "#b91c1c" };
}

function calcScores(scores) {
  const avg = keys => {
    const vals = keys.filter(k => (scores[k] || 0) > 0).map(k => scores[k]);
    return vals.length ? vals.reduce((a, b) => a + b, 0) / vals.length : null;
  };
  const risico    = avg(["A1","A3","B1"]);
  const mitigatie = avg(["C1","D1","E1"]);
  const belang    = avg(["F1","G1","H1"]);
  const dictuAvg  = avg(["2.1","2.2","2.3","4.1"]);

  let autonomyScore = null;
  if (risico && mitigatie && belang) {
    const raw  = mitigatie / (risico * belang);
    const lMin = Math.log(0.04);
    const lMax = Math.log(5);
    autonomyScore = Math.max(1, Math.min(10,
      1 + 9 * (Math.log(Math.max(raw, 0.04)) - lMin) / (lMax - lMin)
    ));
  }
  const allKeys = [...DAAF.map(d => d.key), ...DICTU.map(q => q.key)];
  const filled  = allKeys.filter(k => (scores[k] || 0) > 0).length;
  return { risico, mitigatie, belang, autonomyScore, dictuAvg,
           completeness: Math.round(100 * filled / allKeys.length) };
}

// ──────────────────────────────────────────────────────────────
// COMPONENTS
// ──────────────────────────────────────────────────────────────

function Gauge({ score, size = 88 }) {
  const cx = 50, cy = 46, r = 36, sw = 5.5;
  const rad = d => d * Math.PI / 180;
  const arc = (sd, sweep) => {
    const sx = cx + r * Math.cos(rad(sd)), sy = cy + r * Math.sin(rad(sd));
    const ex = cx + r * Math.cos(rad(sd + sweep)), ey = cy + r * Math.sin(rad(sd + sweep));
    return `M ${sx.toFixed(1)} ${sy.toFixed(1)} A ${r} ${r} 0 ${sweep >= 180 ? 1 : 0} 1 ${ex.toFixed(1)} ${ey.toFixed(1)}`;
  };
  const pct   = score ? (score - 1) / 9 : 0;
  const sweep = 270 * pct;
  const color = scoreColor(score);
  const nd    = 135 + sweep;
  const nl    = r * 0.72;
  const nx    = cx + nl * Math.cos(rad(nd));
  const ny    = cy + nl * Math.sin(rad(nd));
  return (
    <svg width={size} height={size * 0.83} viewBox="0 0 100 84">
      <path d={arc(135, 270)} stroke="#e5e7eb" strokeWidth={sw} fill="none" strokeLinecap="round" />
      {score > 0 && <path d={arc(135, sweep)} stroke={color} strokeWidth={sw} fill="none" strokeLinecap="round" />}
      {score > 0 && <>
        <line x1={cx} y1={cy} x2={nx.toFixed(1)} y2={ny.toFixed(1)} stroke={color} strokeWidth={2.5} strokeLinecap="round" />
        <circle cx={cx} cy={cy} r={3.5} fill={color} />
      </>}
      <text x={cx} y={cy + 20} textAnchor="middle" fontSize={score ? 19 : 13} fontWeight="700" fill={score ? color : "#9ca3af"}>
        {score ? score.toFixed(1) : "–"}
      </text>
      {score > 0 && <text x={cx} y={cy + 29} textAnchor="middle" fontSize={9} fill="#9ca3af">/10</text>}
    </svg>
  );
}

function SovBar({ score5 }) {
  const pct   = score5 ? ((score5 - 1) / 4) * 100 : 0;
  const color = scoreColor(score5, 5);
  return (
    <div>
      <div className="flex justify-between mb-1" style={{ fontSize: 10, color: "#9ca3af" }}>
        <span>Afhankelijk</span><span>Soeverein</span>
      </div>
      <div className="relative h-2 rounded-full" style={{ background: "linear-gradient(to right, #fca5a5, #fde68a, #86efac)" }}>
        {score5 > 0 && (
          <div className="absolute top-1/2 -translate-y-1/2 w-3.5 h-3.5 rounded-full border-2 border-white shadow"
            style={{ left: `calc(${pct}% - 7px)`, background: color, transition: "left 0.3s" }} />
        )}
      </div>
      <div className="flex justify-between mt-0.5" style={{ fontSize: 10, color: "#d1d5db" }}>
        {[1, 2, 3, 4, 5].map(n => <span key={n}>{n}</span>)}
      </div>
    </div>
  );
}

const BTN_COLORS_INV = ["#16a34a","#84cc16","#ca8a04","#ea580c","#dc2626"];
const BTN_COLORS_FWD = ["#dc2626","#ea580c","#ca8a04","#84cc16","#16a34a"];

function ScoreBtn({ s, selected, label, desc, dir, onClick }) {
  const colors = dir === "fwd" ? BTN_COLORS_FWD : BTN_COLORS_INV;
  const c = colors[s - 1];
  return (
    <button onClick={onClick} title={desc} className="flex-1 py-2 px-1 rounded-lg border-2 text-center transition-all"
      style={selected
        ? { borderColor: c, background: c, color: "#fff" }
        : { borderColor: "#e5e7eb", background: "#f9fafb", color: "#374151" }
      }>
      <div style={{ fontWeight: 700, fontSize: 14 }}>{s}</div>
      <div style={{ fontSize: 10, lineHeight: 1.2, marginTop: 2 }}>{label}</div>
    </button>
  );
}

function QuestionCard({ q, value, onChange, dir }) {
  return (
    <div className="mb-3" style={{ background:"#fff", border:"1px solid #D0E4F7", borderRadius:4, padding:16 }}>
      <div className="flex items-start gap-2 mb-3">
        <span className="text-xs font-semibold px-2 py-0.5 flex-shrink-0"
          style={{ borderRadius:3, background:"#EBF3FF", color:"#1A56A0" }}>{q.key}</span>
        <div>
          <p className="font-semibold text-sm leading-snug" style={{ color:"#0C2340" }}>{q.name}</p>
          <p className="text-xs text-gray-500 mt-0.5 leading-relaxed">{q.question || q.description}</p>
        </div>
      </div>
      {q.norm && (
        <div className="rounded p-2 mb-3" style={{ background:"#E6F7F7", border:"1px solid #26B5AE" }}>
          <p style={{ fontSize:11, color:"#0C6B68" }}><span className="font-semibold">Norm: </span>{q.norm}</p>
        </div>
      )}
      {q.toelichting && (
        <div className="rounded p-2 mb-3" style={{ background:"#fffbeb", border:"1px solid #fde68a" }}>
          <p style={{ fontSize:11, color:"#78350f" }}><span className="font-semibold">Toelichting: </span>{q.toelichting}</p>
        </div>
      )}
      {q.hint && (
        <p style={{ fontSize: 10, color: "#9ca3af", marginBottom: 6 }}>{q.hint}</p>
      )}
      <div className="flex gap-1.5">
        {q.scores.map(({ s, label, desc }) => (
          <ScoreBtn key={s} s={s} selected={value === s} label={label} desc={desc} dir={dir} onClick={() => onChange(s)} />
        ))}
      </div>
      {value > 0 && (
        <p className="mt-2 text-gray-400 italic" style={{ fontSize: 11 }}>
          {q.scores.find(sc => sc.s === value)?.desc}
        </p>
      )}
    </div>
  );
}

// ──────────────────────────────────────────────────────────────
// APP
// ──────────────────────────────────────────────────────────────

export default function App() {
  // ── Login state ─────────────────────────────────────────────
  const [loggedIn,   setLoggedIn]   = useState(() => sessionStorage.getItem("nhl_auth") === "ok");
  const [loginInput, setLoginInput] = useState("");
  const [loginError, setLoginError] = useState(false);
  const LOGIN_CODE = "Geheim";

  // ── App state ────────────────────────────────────────────────
  const [apps,       setApps]      = useState([]);
  const [ready,      setReady]     = useState(false);
  const [saving,     setSaving]    = useState(false);
  const [saveError,  setSaveError] = useState(false);
  const [view,       setView]      = useState("about");
  const [selId,      setSelId]     = useState(null);
  const [step,       setStep]      = useState(0);
  const [showModal,  setShowModal] = useState(false);
  const [form,       setForm]      = useState({ name:"", cat:"", supplier:"", owner:"", notes:"" });
  const [hiddenApps, setHiddenApps] = useState(new Set()); // IDs verborgen in dashboard

  // Beheer (admin) state
  const [adminUnlocked, setAdminUnlocked] = useState(false);
  const [adminPin,      setAdminPin]      = useState("");
  const [adminPinError, setAdminPinError] = useState(false);
  const [editAppId,     setEditAppId]     = useState(null);
  const [editForm,      setEditForm]      = useState({});
  const ADMIN_PIN = "nhl2026";

  // Ref voor scroll-naar-boven bij stapwissel in Assess
  const assessScrollRef = React.useRef(null);

  // ── Laden van gedeelde data via Netlify Blobs API ───────────
  useEffect(() => {
    if (!loggedIn) return;
    async function load() {
      try {
        const r = await fetch("/api/load-data");
        if (r.ok) {
          const data = await r.json();
          if (Array.isArray(data)) setApps(data);
        }
      } catch {
        // Fallback: lokale opslag als API niet bereikbaar is
        try {
          const stored = localStorage.getItem("nhl_sov_v2");
          if (stored) setApps(JSON.parse(stored));
        } catch {}
      }
      setReady(true);
    }
    load();
  }, [loggedIn]);

  // ── Opslaan naar gedeelde API (debounced 800ms) ─────────────
  useEffect(() => {
    if (!ready) return;
    const timer = setTimeout(async () => {
      setSaveError(false);
      setSaving(true);
      try {
        const r = await fetch("/api/save-data", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(apps)
        });
        if (!r.ok) throw new Error("save failed");
        // Lokale backup
        localStorage.setItem("nhl_sov_v2", JSON.stringify(apps));
      } catch {
        setSaveError(true);
        // Sla toch lokaal op als fallback
        try { localStorage.setItem("nhl_sov_v2", JSON.stringify(apps)); } catch {}
      } finally {
        setSaving(false);
      }
    }, 800);
    return () => clearTimeout(timer);
  }, [apps, ready]);

  // ── Login handler ────────────────────────────────────────────
  function handleLogin() {
    if (loginInput === LOGIN_CODE) {
      sessionStorage.setItem("nhl_auth", "ok");
      setLoggedIn(true);
      setLoginError(false);
    } else {
      setLoginError(true);
      setLoginInput("");
    }
  }

  // ── Login scherm ─────────────────────────────────────────────
  if (!loggedIn) {
    return (
      <div className="flex items-center justify-center h-screen" style={{ background:"#EBF3FF", fontFamily:"system-ui,sans-serif" }}>
        <div className="bg-white w-full max-w-sm p-0 overflow-hidden" style={{ borderRadius:4, boxShadow:"0 8px 32px rgba(12,35,64,0.2)" }}>
          {/* Header */}
          <div className="px-8 py-6 text-white" style={{ background:"#0C2340" }}>
            <div className="flex items-center gap-3 mb-3">
              <div className="px-2.5 py-1.5 border-2 border-white" style={{ borderRadius:2 }}>
                <span className="font-bold leading-none text-white" style={{ fontSize:10, letterSpacing:1 }}>NHL<br/>STENDEN</span>
              </div>
              <div className="w-px self-stretch" style={{ background:"#26B5AE", margin:"2px 0" }}/>
              <div>
                <p className="font-bold text-white" style={{ fontSize:12 }}>Digitale Soevereiniteitsassessment</p>
                <p style={{ fontSize:10, color:"#7DD3D0" }}>Project Digitale Soevereiniteit</p>
              </div>
            </div>
          </div>
          {/* Form */}
          <div className="px-8 py-6">
            <div className="flex items-center gap-2 mb-5">
              <div className="w-8 h-8 flex items-center justify-center text-white" style={{ background:"#1A56A0", borderRadius:4, fontSize:16 }}>🔒</div>
              <div>
                <p className="font-semibold text-sm" style={{ color:"#0C2340" }}>Toegangscode vereist</p>
                <p className="text-xs text-gray-400">Voer de code in om toegang te krijgen</p>
              </div>
            </div>
            <input
              type="password"
              value={loginInput}
              onChange={e => { setLoginInput(e.target.value); setLoginError(false); }}
              onKeyDown={e => e.key === "Enter" && handleLogin()}
              placeholder="Toegangscode"
              autoFocus
              className="w-full border px-3 py-2.5 text-sm focus:outline-none mb-2"
              style={{ borderColor: loginError ? "#dc2626" : "#D0E4F7", borderRadius:4, letterSpacing:3 }}
            />
            {loginError && (
              <p className="text-xs mb-3" style={{ color:"#dc2626" }}>Toegangscode onjuist. Probeer opnieuw.</p>
            )}
            <button
              onClick={handleLogin}
              className="w-full text-white py-2.5 text-sm font-semibold mt-1"
              style={{ background:"#1A56A0", borderRadius:4 }}>
              Inloggen
            </button>
            <p className="text-xs text-center mt-4" style={{ color:"#9ca3af" }}>
              Neem contact op met de applicatiebeheerder voor de toegangscode.
            </p>
          </div>
        </div>
      </div>
    );
  }

  const selApp = apps.find(a => a.id === selId);

  function addApp() {
    if (!form.name.trim()) return;
    const a = {
      id: Date.now() + "", name: form.name, cat: form.cat,
      supplier: form.supplier, owner: form.owner, notes: form.notes,
      scores: {}, createdAt: new Date().toISOString()
    };
    setApps(p => [...p, a]);
    setForm({ name:"", cat:"", supplier:"", owner:"", notes:"" });
    setShowModal(false);
    setSelId(a.id);
    setStep(0);
    setView("assess");
  }

  function setScore(appId, key, val) {
    setApps(p => p.map(a => a.id === appId ? { ...a, scores: { ...a.scores, [key]: val } } : a));
  }

  function delApp(id) {
    if (!confirm("Applicatie verwijderen? Dit kan niet ongedaan worden gemaakt.")) return;
    setApps(p => p.filter(a => a.id !== id));
    if (selId === id) { setSelId(null); setView("dashboard"); }
  }

  function exportXlsx() {
    const wb = XLSX.utils.book_new();

    const ws1 = XLSX.utils.aoa_to_sheet([
      ["Applicatie","Leverancier","Categorie","Eigenaar","Autonomiescore (1-10)","Risico","Mitigatie","Belang","DICTU Score (1-5)","Volledigheid (%)"],
      ...apps.map(a => {
        const s = calcScores(a.scores);
        return [a.name, a.supplier, a.cat, a.owner,
          s.autonomyScore ? +s.autonomyScore.toFixed(2) : "",
          s.risico    ? +s.risico.toFixed(2)    : "",
          s.mitigatie ? +s.mitigatie.toFixed(2) : "",
          s.belang    ? +s.belang.toFixed(2)    : "",
          s.dictuAvg  ? +s.dictuAvg.toFixed(2)  : "",
          s.completeness
        ];
      })
    ]);
    XLSX.utils.book_append_sheet(wb, ws1, "Overzicht");

    const ws2 = XLSX.utils.aoa_to_sheet([
      ["Applicatie","Leverancier", ...DAAF.map(d => `${d.key} ${d.name}`)],
      ...apps.map(a => [a.name, a.supplier, ...DAAF.map(d => a.scores[d.key] || "")])
    ]);
    XLSX.utils.book_append_sheet(wb, ws2, "DAAF Scores");

    const ws3 = XLSX.utils.aoa_to_sheet([
      ["Applicatie","Leverancier", ...DICTU.map(q => `${q.key} ${q.name}`)],
      ...apps.map(a => [a.name, a.supplier, ...DICTU.map(q => a.scores[q.key] || "")])
    ]);
    XLSX.utils.book_append_sheet(wb, ws3, "DICTU Scores");

    const ws4 = XLSX.utils.aoa_to_sheet([
      ["Vraag","Beschrijving","Norm (min. vereiste score)"],
      ...DAAF.map(d  => [d.key,  d.name + " — " + d.question, ""]),
      ...DICTU.map(q => [q.key,  q.name + " — " + q.question, q.norm])
    ]);
    XLSX.utils.book_append_sheet(wb, ws4, "Framework vragen");

    const out  = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    const blob = new Blob([out], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const url  = URL.createObjectURL(blob);
    const el   = document.createElement("a");
    el.href = url; el.download = "NHL_Stenden_Soevereiniteitsassessment.xlsx"; el.click();
    URL.revokeObjectURL(url);
  }

  // ── VIEWS ──────────────────────────────────────────────────

  function Dashboard() {
    // ── Zichtbare applicaties (gefilterd op hiddenApps) ─────────
    const visibleApps = apps.filter(a => !hiddenApps.has(a.id));
    const minVisible  = 2;

    function toggleApp(id) {
      setHiddenApps(prev => {
        const next = new Set(prev);
        if (next.has(id)) {
          next.delete(id); // altijd zichtbaar maken mag
        } else {
          if (visibleApps.length <= minVisible) return prev; // minimum bewaken
          next.add(id);
        }
        return next;
      });
    }

    const scored = visibleApps.map(a => ({ ...a, sc: calcScores(a.scores) }));
    const allScored = apps.map(a => ({ ...a, sc: calcScores(a.scores) })); // voor filter-strip
    const withSc = scored.filter(a => a.sc.autonomyScore);
    const avgA   = withSc.length ? withSc.map(a => a.sc.autonomyScore).reduce((x,y)=>x+y,0)/withSc.length : null;
    const COLORS = ["#1e40af","#7c3aed","#065f46","#92400e","#991b1b","#0f766e"];

    const radarKey = n => n.substring(0, 13);

    // Dedupleer per dim-letter: gemiddeld van alle vragen in die dimensie
    const dimLetters = [...new Set(DAAF.map(d => d.dim))];
    const dimLabel = letter => {
      const first = DAAF.find(d => d.dim === letter);
      return first ? first.dimName.substring(0, 14) : letter;
    };
    const radarData = dimLetters.map(letter => {
      const qs = DAAF.filter(d => d.dim === letter);
      const entry = { dim: dimLabel(letter) };
      visibleApps.slice(0, 5).forEach(a => {
        const vals = qs.map(q => a.scores[q.key] || 0).filter(v => v > 0);
        entry[radarKey(a.name)] = vals.length ? vals.reduce((x,y)=>x+y,0)/vals.length : 0;
      });
      return entry;
    });

    // Kwadrant data: X = risico × belang (1–25), Y = mitigatie (1–5)
    const kwData = scored
      .filter(a => a.sc.risico && a.sc.mitigatie && a.sc.belang)
      .map(a => ({
        name: a.name,
        x: +((a.sc.risico * a.sc.belang).toFixed(2)),
        y: +a.sc.mitigatie.toFixed(2),
        score: a.sc.autonomyScore,
        id: a.id
      }));

    // SVG Autonomie-kwadrant
    const KwadrantSVG = () => {
      const W = 700, H = 420;
      const pad = { top:32, right:24, bottom:52, left:52 };
      const iW  = W - pad.left - pad.right;
      const iH  = H - pad.top  - pad.bottom;

      // Scales: X = 1..25, Y = 1..5
      const xMin=1, xMax=25, yMin=1, yMax=5;
      const mx = 13; // midpoint X
      const my = 3;  // midpoint Y

      const toX = v => pad.left + (v - xMin) / (xMax - xMin) * iW;
      const toY = v => pad.top  + (yMax - v) / (yMax - yMin) * iH;
      const midX = toX(mx);
      const midY = toY(my);

      // Quadrant bg colors (matching the image)
      const quads = [
        { x1:pad.left, y1:pad.top,  x2:midX,      y2:midY,         fill:"#e8f5e9", label:"OPTIMAAL",      sub:"Behoud huidige situatie,\nmonitor periodiek",        color:"#2e7d5e" },
        { x1:midX,     y1:pad.top,  x2:pad.left+iW,y2:midY,         fill:"#fff8e1", label:"BEHEERSBAAR",   sub:"Risico's geaccepteerd\nmet goede fallback",           color:"#e07b20" },
        { x1:pad.left, y1:midY,     x2:midX,       y2:pad.top+iH,   fill:"#fff3e0", label:"AANDACHTSPUNT", sub:"Bouw mitigatie op of\naccep­teer risico bewust",       color:"#e07b20" },
        { x1:midX,     y1:midY,     x2:pad.left+iW,y2:pad.top+iH,   fill:"#fce4ec", label:"KRITIEK",       sub:"Urgente actie vereist:\nmigreer of mitigeer",         color:"#c0392b" },
      ];

      // Axis ticks
      const xTicks = [1,5,10,15,20,25];
      const yTicks = [1,2,3,4,5];

      return (
        <svg viewBox={`0 0 ${W} ${H}`} width="100%" style={{ fontFamily:"system-ui,sans-serif", display:"block" }}>
          {/* Quadrant backgrounds */}
          {quads.map((q,i) => (
            <rect key={i} x={q.x1} y={q.y1} width={q.x2-q.x1} height={q.y2-q.y1} fill={q.fill} />
          ))}

          {/* Quadrant labels */}
          {quads.map((q,i) => {
            const cx = (q.x1 + q.x2) / 2;
            const cy = (q.y1 + q.y2) / 2;
            const lines = q.sub.split("\n");
            return (
              <g key={i}>
                <text x={cx} y={cy - 14} textAnchor="middle" fill={q.color}
                  style={{ fontSize:15, fontWeight:700, fontStyle:"italic", letterSpacing:1 }}>
                  {q.label}
                </text>
                {lines.map((l,j) => (
                  <text key={j} x={cx} y={cy + 8 + j*16} textAnchor="middle" fill="#555"
                    style={{ fontSize:11 }}>{l}</text>
                ))}
              </g>
            );
          })}

          {/* Grid lines */}
          {xTicks.map(v => (
            <line key={v} x1={toX(v)} y1={pad.top} x2={toX(v)} y2={pad.top+iH}
              stroke="#fff" strokeWidth={v===mx?0:1} strokeDasharray="3 3" />
          ))}
          {yTicks.map(v => (
            <line key={v} x1={pad.left} y1={toY(v)} x2={pad.left+iW} y2={toY(v)}
              stroke="#fff" strokeWidth={v===my?0:1} strokeDasharray="3 3" />
          ))}

          {/* Midpoint divider lines */}
          <line x1={midX} y1={pad.top} x2={midX} y2={pad.top+iH} stroke="#aaa" strokeWidth={1.5} />
          <line x1={pad.left} y1={midY} x2={pad.left+iW} y2={midY} stroke="#aaa" strokeWidth={1.5} />

          {/* Border */}
          <rect x={pad.left} y={pad.top} width={iW} height={iH}
            fill="none" stroke="#ccc" strokeWidth={1} />

          {/* X axis ticks & labels */}
          {xTicks.map(v => (
            <g key={v}>
              <line x1={toX(v)} y1={pad.top+iH} x2={toX(v)} y2={pad.top+iH+5} stroke="#999" strokeWidth={1}/>
              <text x={toX(v)} y={pad.top+iH+17} textAnchor="middle" fill="#888" style={{ fontSize:10 }}>{v}</text>
            </g>
          ))}

          {/* Y axis ticks & labels */}
          {yTicks.map(v => (
            <g key={v}>
              <line x1={pad.left-5} y1={toY(v)} x2={pad.left} y2={toY(v)} stroke="#999" strokeWidth={1}/>
              <text x={pad.left-10} y={toY(v)+4} textAnchor="end" fill="#888" style={{ fontSize:10 }}>{v}</text>
            </g>
          ))}

          {/* Axis labels */}
          <text x={pad.left + iW/2} y={H-4} textAnchor="middle" fill="#444"
            style={{ fontSize:12, fontWeight:600 }}>
            Risico-exposure × Strategisch belang
          </text>
          <text x={14} y={pad.top + iH/2} textAnchor="middle" fill="#444"
            transform={`rotate(-90, 14, ${pad.top + iH/2})`}
            style={{ fontSize:12, fontWeight:600 }}>
            Mitigatie
          </text>

          {/* Axis direction hints */}
          <text x={pad.left+6}  y={pad.top-10} fill="#aaa" style={{ fontSize:9 }}>Laag</text>
          <text x={pad.left+iW-24} y={pad.top-10} fill="#aaa" style={{ fontSize:9 }}>Hoog</text>
          <text x={pad.left-46} y={pad.top+iH-4} fill="#aaa" style={{ fontSize:9 }}>Laag</text>
          <text x={pad.left-46} y={pad.top+10}   fill="#aaa" style={{ fontSize:9 }}>Hoog</text>

          {/* App dots */}
          {kwData.length === 0 && (
            <text x={W/2} y={H/2+10} textAnchor="middle" fill="#bbb" style={{ fontSize:13 }}>
              Vul minimaal alle dimensies in om applicaties in het kwadrant te plotten
            </text>
          )}
          {kwData.map((d, i) => {
            const cx = toX(d.x);
            const cy = toY(d.y);
            const col = scoreColor(d.score);
            return (
              <g key={d.id} style={{ cursor:"pointer" }}
                onClick={() => { setSelId(d.id); setStep(0); setView("assess"); }}>
                <circle cx={cx} cy={cy} r={13} fill={col} fillOpacity={0.2} stroke={col} strokeWidth={2} />
                <circle cx={cx} cy={cy} r={5}  fill={col} />
                {/* Label — shift to avoid overlap */}
                <rect x={cx+10} y={cy-10} width={Math.min(d.name.length*6.5+8,120)} height={16} rx={3}
                  fill="white" fillOpacity={0.88} />
                <text x={cx+14} y={cy+2} fill={col} style={{ fontSize:10, fontWeight:600 }}>
                  {d.name.substring(0,18)}
                </text>
              </g>
            );
          })}
        </svg>
      );
    };

    return (
      <div className="h-full overflow-y-auto" style={{ background:"#EBF3FF" }}>
        <div className="p-5" style={{ maxWidth:"100%", margin:"0 auto" }}>

          {/* ── Stat row ── */}
          <div className="grid grid-cols-4 gap-3 mb-4">
            {[
              { label:"Applicaties",        val: apps.length,     color: "#1A56A0" },
              { label:"Gem. autonomiescore", val: avgA ? avgA.toFixed(1) : "–", color: scoreColor(avgA) },
              { label:"Goed (≥7)",           val: withSc.filter(a=>a.sc.autonomyScore>=7).length, color:"#26B5AE" },
              { label:"Aandacht nodig (<5)", val: withSc.filter(a=>a.sc.autonomyScore<5).length,  color:"#E87722" },
            ].map(({ label, val, color }) => (
              <div key={label} className="rounded text-center px-3 py-3"
                style={{ background:"#fff", border:"1px solid #D0E4F7" }}>
                <div style={{ fontSize: 26, fontWeight: 700, color }}>{val}</div>
                <div className="text-xs mt-0.5" style={{ color:"#6b7280" }}>{label}</div>
              </div>
            ))}
          </div>

          {/* ── Filter strip — alleen tonen als er >2 apps zijn ── */}
          {apps.length > 2 && (
            <div className="rounded p-3 mb-4 flex items-center gap-3 flex-wrap"
              style={{ background:"#fff", border:"1px solid #D0E4F7" }}>
              <div className="flex items-center gap-1.5 flex-shrink-0">
                <span style={{ fontSize:11, fontWeight:600, color:"#0C2340" }}>Dashboard filter</span>
                <span style={{ fontSize:10, color:"#9ca3af" }}>
                  — {visibleApps.length} van {apps.length} zichtbaar
                  {hiddenApps.size > 0 && ` · ${hiddenApps.size} verborgen`}
                </span>
              </div>
              <div className="flex gap-2 flex-wrap flex-1">
                {allScored.map((a, i) => {
                  const hidden  = hiddenApps.has(a.id);
                  const isLast2 = !hidden && visibleApps.length <= minVisible;
                  const col     = COLORS[i % COLORS.length];
                  return (
                    <button key={a.id}
                      onClick={() => toggleApp(a.id)}
                      disabled={isLast2 && !hidden}
                      title={isLast2 && !hidden ? "Minimaal 2 applicaties moeten zichtbaar blijven" : hidden ? "Klik om zichtbaar te maken" : "Klik om te verbergen"}
                      className="flex items-center gap-1.5 text-xs px-2.5 py-1.5 font-medium transition-all"
                      style={{
                        borderRadius: 4,
                        border: `2px solid ${hidden ? "#e5e7eb" : col}`,
                        background: hidden ? "#f9fafb" : `${col}18`,
                        color: hidden ? "#9ca3af" : col,
                        opacity: isLast2 && !hidden ? 0.5 : 1,
                        cursor: isLast2 && !hidden ? "not-allowed" : "pointer",
                        textDecoration: hidden ? "line-through" : "none"
                      }}>
                      <span className="w-2 h-2 rounded-full flex-shrink-0"
                        style={{ background: hidden ? "#d1d5db" : col }} />
                      {a.name.substring(0, 22)}
                      <span style={{ fontSize:10, marginLeft:2, opacity:0.7 }}>
                        {hidden ? "＋" : "✕"}
                      </span>
                    </button>
                  );
                })}
              </div>
              {hiddenApps.size > 0 && (
                <button onClick={() => setHiddenApps(new Set())}
                  className="text-xs px-2.5 py-1.5 flex-shrink-0 font-medium"
                  style={{ borderRadius:4, background:"#EBF3FF", color:"#1A56A0", border:"1px solid #D0E4F7" }}>
                  Alles tonen
                </button>
              )}
            </div>
          )}

          {apps.length === 0 ? (

          {/* ── Rij 1: Kwadrant (links) + App-kaarten (rechts, 2 cols) ── */}
          <div className="grid gap-4 mb-4" style={{ gridTemplateColumns:"1fr 1fr" }}>

            {/* Autonomie-kwadrant */}
            <div className="rounded p-4" style={{ background:"#fff", border:"1px solid #D0E4F7" }}>
              <div className="flex items-center justify-between mb-2 flex-wrap gap-2">
                <div className="flex items-center gap-2">
                  <span className="text-xs font-bold px-2 py-0.5" style={{ background:"#1A56A0", color:"#fff", borderRadius:3 }}>DAAF</span>
                  <h3 className="font-bold" style={{ color:"#0C2340", fontSize:14 }}>Autonomie-kwadrant</h3>
                </div>
                <div className="flex gap-1.5 flex-wrap">
                  {[
                    { label:"OPTIMAAL",      color:"#2e7d5e", bg:"#e8f5e9" },
                    { label:"BEHEERSBAAR",   color:"#e07b20", bg:"#fff8e1" },
                    { label:"AANDACHTSPUNT", color:"#e07b20", bg:"#fff3e0" },
                    { label:"KRITIEK",       color:"#c0392b", bg:"#fce4ec" },
                  ].map(l => (
                    <span key={l.label} style={{ background:l.bg, color:l.color, borderRadius:3, fontSize:10, padding:"2px 6px", fontWeight:600 }}>{l.label}</span>
                  ))}
                </div>
              </div>
              <div className="rounded p-2 mb-3 text-xs" style={{ background:"#EBF3FF", border:"1px solid #D0E4F7", color:"#374151" }}>
                <strong>Horizontale as</strong> = Risico × Belang (rechts = meer urgentie) ·{" "}
                <strong>Verticale as</strong> = Mitigatie (hoger = beter) ·{" "}
                <span style={{ color:"#2e7d5e", fontWeight:600 }}>Linksboven</span> = ideaal ·{" "}
                <span style={{ color:"#c0392b", fontWeight:600 }}>Rechtsboven</span> = actie vereist.
                Klik op een punt om naar het assessment te gaan.
              </div>
              <KwadrantSVG />
              <div className="mt-2 px-2 py-1 rounded text-xs" style={{ background:"#f8fafc", border:"1px solid #e5e7eb", color:"#9ca3af" }}>
                Risico = gem. A1+A3+B1 · Mitigatie = gem. C1+D1+E1 · Belang = gem. F1+G1+H1 · Grens: X=13, Y=3
              </div>
            </div>

            {/* App-kaarten rechts — 2 kolommen */}
            <div>
              <div className="flex items-center justify-between mb-3">
                <h3 className="text-sm font-bold" style={{ color:"#0C2340" }}>Applicaties ({apps.length})</h3>
                <div className="flex gap-2">
                  <button onClick={() => setShowModal(true)}
                    className="text-white text-xs px-3 py-1.5 font-medium"
                    style={{ background:"#1A56A0", borderRadius:4 }}>
                    + Toevoegen
                  </button>
                  <button onClick={exportXlsx} disabled={apps.length===0}
                    className="text-xs px-3 py-1.5 font-medium"
                    style={{ background:"#E6F7F7", color:"#26B5AE", borderRadius:4, border:"1px solid #26B5AE55", opacity:apps.length===0?0.5:1 }}>
                    📥 Excel
                  </button>
                </div>
              </div>
              <div className="grid grid-cols-2 gap-3">
                {scored.map(a => {
                  const lbl = scoreLabel(a.sc.autonomyScore);
                  return (
                    <div key={a.id}
                      onClick={() => { setSelId(a.id); setStep(0); setView("assess"); }}
                      className="cursor-pointer transition-all"
                      style={{ background:"#fff", borderRadius:4, padding:14, border:"1px solid #D0E4F7", borderLeft:`4px solid ${scoreColor(a.sc.autonomyScore)}` }}
                      onMouseEnter={e=>e.currentTarget.style.boxShadow="0 2px 10px rgba(26,86,160,0.15)"}
                      onMouseLeave={e=>e.currentTarget.style.boxShadow="none"}>
                      <div className="flex items-start gap-2 mb-2">
                        <Gauge score={a.sc.autonomyScore} size={58} />
                        <div className="flex-1 min-w-0">
                          <h3 className="font-bold text-xs truncate" style={{ color:"#0C2340" }}>{a.name}</h3>
                          {a.supplier && <p style={{ fontSize:10, color:"#9ca3af" }}>{a.supplier}</p>}
                          <span className="text-xs px-1.5 py-0.5 font-medium mt-1 inline-block"
                            style={{ borderRadius:3, background: lbl.bg, color: lbl.fg, fontSize:10 }}>{lbl.text}</span>
                        </div>
                      </div>
                      <div className="mb-2">
                        <div className="flex justify-between" style={{ fontSize:9, color:"#9ca3af", marginBottom:2 }}>
                          <span>DICTU</span><span>{a.sc.dictuAvg ? a.sc.dictuAvg.toFixed(1)+"/5" : "–"}</span>
                        </div>
                        <SovBar score5={a.sc.dictuAvg} />
                      </div>
                      <div className="grid grid-cols-3 gap-1">
                        {[
                          { lbl:"Risico",    val:a.sc.risico,    color:"#dc2626" },
                          { lbl:"Mitigatie", val:a.sc.mitigatie, color:"#26B5AE" },
                          { lbl:"Belang",    val:a.sc.belang,    color:"#E87722" },
                        ].map(s => (
                          <div key={s.lbl} className="text-center rounded py-1" style={{ background:"#f8fafc" }}>
                            <div style={{ fontSize:12, fontWeight:700, color:s.val ? s.color : "#d1d5db" }}>{s.val ? s.val.toFixed(1) : "–"}</div>
                            <div style={{ fontSize:9, color:"#9ca3af" }}>{s.lbl}</div>
                          </div>
                        ))}
                      </div>
                    </div>
                  );
                })}
              </div>

              {/* Leeswijzer compact onderaan */}
              <div className="grid grid-cols-2 gap-2 mt-3">
                <div className="rounded p-2.5" style={{ background:"#EBF3FF", border:"1px solid #D0E4F7" }}>
                  <div className="flex items-center gap-1.5 mb-1">
                    <span className="text-xs font-bold px-1.5 py-0.5" style={{ background:"#1A56A0", color:"#fff", borderRadius:3, fontSize:9 }}>DAAF</span>
                    <span className="font-semibold" style={{ fontSize:10, color:"#0C2340" }}>Snelheidsmeter (1–10)</span>
                  </div>
                  <div className="flex gap-1.5 flex-wrap">
                    {[{t:"≥7 Goed",bg:"#dcfce7",fg:"#15803d"},{t:"5-7 OK",bg:"#fef9c3",fg:"#a16207"},{t:"3-5 ⚠️",bg:"#ffedd5",fg:"#c2410c"},{t:"<3 🔴",bg:"#fee2e2",fg:"#b91c1c"}].map(s=>(
                      <span key={s.t} style={{ fontSize:9, background:s.bg, color:s.fg, borderRadius:2, padding:"1px 4px", fontWeight:600 }}>{s.t}</span>
                    ))}
                  </div>
                </div>
                <div className="rounded p-2.5" style={{ background:"#E6F7F7", border:"1px solid #26B5AE44" }}>
                  <div className="flex items-center gap-1.5 mb-1">
                    <span className="text-xs font-bold px-1.5 py-0.5" style={{ background:"#26B5AE", color:"#fff", borderRadius:3, fontSize:9 }}>DICTU</span>
                    <span className="font-semibold" style={{ fontSize:10, color:"#0C2340" }}>Kleurenbalk (1–5)</span>
                  </div>
                  <p style={{ fontSize:9, color:"#6b7280" }}>Rood = afhankelijk → groen = soeverein. Gem. van 2.1, 2.2, 2.3 en 4.1.</p>
                </div>
              </div>
            </div>
          </div>

          {/* ── Rij 2: Radar full-width ── */}
          <div className="rounded p-4 mb-4" style={{ background:"#fff", border:"1px solid #D0E4F7" }}>
            <div className="flex items-center gap-2 mb-1">
              <span className="text-xs font-bold px-2 py-0.5" style={{ background:"#1A56A0", color:"#fff", borderRadius:3 }}>DAAF</span>
              <h3 className="font-bold" style={{ color:"#0C2340", fontSize:14 }}>Spindiagram — dimensies per applicatie</h3>
            </div>
            <p className="text-xs leading-relaxed mb-3" style={{ color:"#6b7280" }}>
              Elke as toont één DAAF-dimensie (gemiddelde score 1–5).{" "}
              <span style={{ color:"#26B5AE", fontWeight:600 }}>Mitigatie-assen</span>: groter is beter.{" "}
              <span style={{ color:"#dc2626", fontWeight:600 }}>Risico-assen</span> en{" "}
              <span style={{ color:"#E87722", fontWeight:600 }}>Belang-assen</span>: kleiner is beter.
              Vergelijk de vormen van applicaties om te zien waar de grootste verschillen zitten.
            </p>
            {apps.length >= 1 ? (
              <ResponsiveContainer width="100%" height={360}>
                <RadarChart data={radarData} margin={{ top:15, right:90, bottom:15, left:90 }}>
                  <PolarGrid stroke="#e5e7eb" />
                  <PolarAngleAxis dataKey="dim" tick={{ fontSize: 12, fill:"#374151" }} />
                  <PolarRadiusAxis domain={[0, 5]} tick={{ fontSize:9, fill:"#9ca3af" }} tickCount={6} />
                  {apps.slice(0, 5).map((a, i) => (
                    <Radar key={a.id} name={a.name} dataKey={radarKey(a.name)}
                      stroke={COLORS[i % COLORS.length]} fill={COLORS[i % COLORS.length]} fillOpacity={0.15} strokeWidth={2} />
                  ))}
                  <Tooltip wrapperStyle={{ fontSize: 11 }} />
                  <Legend wrapperStyle={{ fontSize: 11 }} />
                </RadarChart>
              </ResponsiveContainer>
            ) : (
              <div className="text-center py-10 text-sm text-gray-400">Voeg een applicatie toe om het spindiagram te zien.</div>
            )}
          </div>

          </>)}
        </div>
      </div>
    );
  }

          {/* ── Stat row ── */}
          <div className="grid grid-cols-4 gap-3 mb-4">
            {[
              { label:"Applicaties",        val: apps.length,     color: "#1A56A0" },
              { label:"Gem. autonomiescore", val: avgA ? avgA.toFixed(1) : "\u2013", color: scoreColor(avgA) },
              { label:"Goed (\u22657)",      val: withSc.filter(a=>a.sc.autonomyScore>=7).length, color:"#26B5AE" },
              { label:"Aandacht nodig (<5)", val: withSc.filter(a=>a.sc.autonomyScore<5).length,  color:"#E87722" },
            ].map(({ label, val, color }) => (
              <div key={label} className="rounded text-center px-3 py-3"
                style={{ background:"#fff", border:"1px solid #D0E4F7" }}>
                <div style={{ fontSize: 26, fontWeight: 700, color }}>{val}</div>
                <div className="text-xs mt-0.5" style={{ color:"#6b7280" }}>{label}</div>
              </div>
            ))}
          </div>


  function AppsList() {
    return (
      <div className="h-full overflow-y-auto" style={{ background:"#EBF3FF" }}>
        <div className="p-5 max-w-4xl mx-auto">
          <div className="flex justify-between items-center mb-4">
            <h2 className="text-lg font-semibold" style={{ color:"#0C2340" }}>Applicaties ({apps.length})</h2>
            <button onClick={() => setShowModal(true)}
              className="text-white text-sm px-4 py-2 font-medium"
              style={{ background:"#1A56A0", borderRadius:4 }}>
              + Toevoegen
            </button>
          </div>

          {apps.length === 0 ? (
            <div className="rounded text-center p-10" style={{ background:"#fff", border:"2px dashed #D0E4F7", color:"#9ca3af" }}>
              Nog geen applicaties toegevoegd.
            </div>
          ) : (
            <div className="space-y-3">
              {apps.map(a => {
                const sc  = calcScores(a.scores);
                const lbl = scoreLabel(sc.autonomyScore);
                return (
                  <div key={a.id} className="flex items-center gap-4 rounded p-4"
                    style={{ background:"#fff", border:"1px solid #D0E4F7",
                             borderLeft:`4px solid ${scoreColor(sc.autonomyScore)}` }}>
                    <Gauge score={sc.autonomyScore} size={64} />
                    <div className="flex-1 min-w-0">
                      <h3 className="font-semibold" style={{ color:"#0C2340" }}>{a.name}</h3>
                      <p className="text-xs text-gray-500">{a.supplier}{a.cat ? ` · ${a.cat}` : ""}</p>
                      <div className="flex items-center gap-2 mt-1">
                        <span className="text-xs px-2 py-0.5 font-medium" style={{ borderRadius:3, background: lbl.bg, color: lbl.fg }}>
                          {lbl.text}
                        </span>
                        <span className="text-xs text-gray-400">{sc.completeness}% van vragen ingevuld</span>
                      </div>
                    </div>
                    <div className="flex gap-2 flex-shrink-0">
                      <button onClick={() => { setSelId(a.id); setStep(0); setView("assess"); }}
                        className="text-white text-xs px-3 py-1.5 font-medium"
                        style={{ background:"#1A56A0", borderRadius:4 }}>
                        Assessment
                      </button>
                      <button onClick={() => delApp(a.id)}
                        className="text-xs px-3 py-1.5"
                        style={{ borderRadius:4, border:"1px solid #fecaca", color:"#dc2626" }}>
                        Verwijder
                      </button>
                    </div>
                  </div>
                );
              })}
            </div>
          )}
        </div>
      </div>
    );
  }

  function Assess() {
    if (!selApp) return (
      <div className="h-full flex items-center justify-center text-gray-400">
        Selecteer een applicatie om te beginnen.
      </div>
    );
    const sc  = calcScores(selApp.scores);
    const lbl = scoreLabel(sc.autonomyScore);

    return (
      <div className="flex h-full overflow-hidden">
        {/* Main form */}
        <div ref={assessScrollRef} className="flex-1 overflow-y-auto p-5" style={{ background:"#EBF3FF" }}>
          <div className="max-w-2xl mx-auto">
            <div className="flex items-center gap-2 mb-4">
              <button onClick={() => setView("apps")} className="text-sm hover:underline font-medium" style={{ color:"#1A56A0" }}>
                ← Applicaties
              </button>
              <span className="text-gray-300">›</span>
              <h2 className="font-semibold truncate" style={{ color:"#0C2340" }}>{selApp.name}</h2>
            </div>

            {/* Step tabs */}
            <div className="flex gap-2 mb-5">
              {[
                { i:0, label:"1 · DAAF Quick Scan",           bg:"#1A56A0" },
                { i:1, label:"2 · DICTU Soevereiniteitscheck", bg:"#26B5AE" }
              ].map(t => (
                <button key={t.i} onClick={() => {
                  setStep(t.i);
                  if (assessScrollRef.current) assessScrollRef.current.scrollTop = 0;
                }}
                  className="text-sm px-4 py-2 font-medium transition-all"
                  style={step === t.i
                    ? { background: t.bg, color:"#fff", borderRadius:4 }
                    : { background:"rgba(255,255,255,0.8)", color:"#6b7280", borderRadius:4, border:"1px solid #D0E4F7" }
                  }>
                  {t.label}
                </button>
              ))}
            </div>

            {step === 0 && <>
              <div className="rounded p-3 mb-4 text-xs"
                style={{ background:"#EBF3FF", border:"1px solid #D0E4F7", color:"#1A56A0" }}>
                <strong>DAAF Framework — Quick Scan</strong> — Niveau 1 (Risico): lager = minder risico.
                Niveau 2 (Mitigatie): hoger = betere weerbaarheid. Niveau 3 (Belang): lager = minder urgent autonomieprobleem.
              </div>
              {["Risico","Mitigatie","Belang"].map(lv => {
                const colors = { Risico:"#dc2626", Mitigatie:"#26B5AE", Belang:"#E87722" };
                const labels = { Risico:"Niveau 1: Risico-exposure", Mitigatie:"Niveau 2: Mitigatie-capaciteit", Belang:"Niveau 3: Strategisch belang" };
                return (
                  <div key={lv} className="mb-5">
                    <h3 className="text-sm font-semibold mb-2 flex items-center gap-2" style={{ color: colors[lv] }}>
                      <span className="w-1 inline-block rounded-full" style={{ background:colors[lv], height:14 }}/>
                      {labels[lv]}
                    </h3>
                    {DAAF.filter(d => d.level === lv).map(q => (
                      <QuestionCard key={q.key} q={q} value={selApp.scores[q.key] || 0}
                        dir={lv === "Mitigatie" ? "fwd" : "inv"}
                        onChange={v => setScore(selApp.id, q.key, v)} />
                    ))}
                  </div>
                );
              })}
              <button onClick={() => {
                setStep(1);
                if (assessScrollRef.current) assessScrollRef.current.scrollTop = 0;
              }}
                className="w-full text-white py-3 text-sm font-medium mb-4"
                style={{ background:"#26B5AE", borderRadius:4 }}>
                Verder: DICTU Soevereiniteitscheck →
              </button>
            </>}

            {step === 1 && <>
              <div className="rounded p-3 mb-4 text-xs"
                style={{ background:"#E6F7F7", border:"1px solid #26B5AE", color:"#0C6B68" }}>
                <strong>DICTU Soevereiniteitscheck</strong> — Geselecteerde DICTU-vragen voor digitale soevereiniteit.
                Score 1 = minst soeverein · Score 5 = maximaal soeverein (hoger is beter).
              </div>
              {["Data & AI","EU-Infrastructuur"].map(cat => (
                <div key={cat} className="mb-5">
                  <h3 className="text-sm font-semibold mb-2 flex items-center gap-2" style={{ color:"#1A56A0" }}>
                    <span className="w-1 inline-block rounded-full" style={{ background:"#26B5AE", height:14 }}/>
                    {cat}
                  </h3>
                  {DICTU.filter(q => q.cat === cat).map(q => (
                    <QuestionCard key={q.key} q={q} value={selApp.scores[q.key] || 0}
                      dir="fwd"
                      onChange={v => setScore(selApp.id, q.key, v)} />
                  ))}
                </div>
              ))}
              <button onClick={() => setView("dashboard")}
                className="w-full text-white py-3 text-sm font-medium mb-4"
                style={{ background:"#1A56A0", borderRadius:4 }}>
                ✓ Afronden en naar Dashboard →
              </button>
            </>}
          </div>
        </div>

        {/* Score sidebar */}
        <div className="w-64 flex-shrink-0 border-l overflow-y-auto p-4"
          style={{ background:"#fff", borderColor:"#D0E4F7" }}>
          <div className="flex items-center gap-2 mb-3 pb-2" style={{ borderBottom:"2px solid #1A56A0" }}>
            <div className="w-1 h-4 rounded-full" style={{ background:"#26B5AE" }}/>
            <p className="text-sm font-semibold" style={{ color:"#0C2340" }}>Live scoreoverzicht</p>
          </div>

          <div className="flex justify-center mb-3">
            <div className="text-center">
              <Gauge score={sc.autonomyScore} size={96} />
              <p className="text-xs text-gray-500 mt-0.5">Autonomiescore</p>
              <span className="text-xs px-2 py-0.5 rounded-full font-medium mt-1 inline-block"
                style={{ background: lbl.bg, color: lbl.fg }}>{lbl.text}</span>
            </div>
          </div>

          {/* DAAF levels */}
          <div className="rounded p-3 mb-3" style={{ border:"1px solid #D0E4F7" }}>
            {[
              { label:"Risico-exposure",   val:sc.risico,    hint:"↓ goed", color:"#dc2626" },
              { label:"Mitigatie-cap.",     val:sc.mitigatie, hint:"↑ goed", color:"#26B5AE" },
              { label:"Strategisch belang", val:sc.belang,    hint:"↓ goed", color:"#E87722" },
            ].map(({ label,val,hint,color }) => (
              <div key={label} className="mb-2 last:mb-0">
                <div className="flex justify-between text-xs mb-0.5">
                  <span style={{ color, fontWeight:600 }}>{label}</span>
                  <span className="text-gray-400">{val ? val.toFixed(1) : "–"} <span style={{ fontSize:9 }}>{hint}</span></span>
                </div>
                <div className="bg-gray-100 rounded-full h-1.5">
                  <div className="h-1.5 rounded-full" style={{ width:`${(val||0)*20}%`, background: color, transition:"width 0.3s" }} />
                </div>
              </div>
            ))}
          </div>

          {/* DICTU sovereignty */}
          <div className="rounded p-3 mb-3" style={{ border:"1px solid #26B5AE", background:"#E6F7F7" }}>
            <p className="text-xs font-semibold mb-2" style={{ color:"#0C6B68" }}>DICTU Soevereiniteit</p>
            {sc.dictuAvg ? (
              <>
                <div className="text-center mb-2">
                  <span style={{ fontSize:22, fontWeight:700, color:"#26B5AE" }}>{sc.dictuAvg.toFixed(1)}</span>
                  <span className="text-xs text-gray-400">/5</span>
                </div>
                <SovBar score5={sc.dictuAvg} />
              </>
            ) : (
              <p className="text-xs text-center" style={{ color:"#26B5AE" }}>Nog geen DICTU scores ingevuld</p>
            )}
          </div>

          {/* Completeness */}
          <div className="rounded p-3 mb-3" style={{ border:"1px solid #D0E4F7" }}>
            <div className="flex justify-between text-xs mb-1">
              <span className="font-medium text-gray-600">Volledigheid</span>
              <span style={{ fontWeight:700, color:"#1A56A0" }}>{sc.completeness}%</span>
            </div>
            <div className="rounded-full h-2" style={{ background:"#EBF3FF" }}>
              <div className="h-2 rounded-full transition-all" style={{ width:`${sc.completeness}%`, background:"#1A56A0" }} />
            </div>
          </div>

          {/* Per-question dots */}
          <div className="rounded p-3" style={{ border:"1px solid #D0E4F7" }}>
            <p className="text-xs font-semibold mb-2" style={{ color:"#0C2340" }}>Per vraag</p>
            <div className="space-y-1">
              {[...DAAF,...DICTU].map(q => (
                <div key={q.key} className="flex items-center gap-1.5">
                  <span className="w-8 flex-shrink-0" style={{ fontSize:10, color:"#9ca3af" }}>{q.key}</span>
                  <div className="flex gap-0.5">
                    {[1,2,3,4,5].map(s => (
                      <div key={s} className="w-2.5 h-2.5 rounded-sm"
                        style={{ background:(selApp.scores[q.key]||0) >= s ? "#1A56A0" : "#e5e7eb" }} />
                    ))}
                  </div>
                  {(selApp.scores[q.key]||0) > 0 && (
                    <span style={{ fontSize:9, color:"#9ca3af" }}>{selApp.scores[q.key]}</span>
                  )}
                </div>
              ))}
            </div>
          </div>
        </div>
      </div>
    );
  }

  function Compare() {
    if (apps.length < 2) return (
      <div className="h-full flex items-center justify-center text-gray-400">
        <div className="text-center">
          <p className="text-lg mb-1">Voeg minimaal 2 applicaties toe</p>
          <p className="text-sm">om een vergelijking te kunnen maken.</p>
        </div>
      </div>
    );

    const COLORS = ["#1e40af","#7c3aed","#065f46","#92400e","#991b1b","#0f766e"];
    const name14 = n => n.substring(0, 14);

    const radarData = DAAF.map(d => ({
      dim: d.dimName.substring(0, 12),
      ...Object.fromEntries(apps.map(a => [name14(a.name), a.scores[d.key] || 0]))
    }));

    const barData = apps.map(a => {
      const s = calcScores(a.scores);
      return {
        name: a.name.substring(0, 16),
        Autonomiescore: s.autonomyScore ? +s.autonomyScore.toFixed(1) : 0,
        "DICTU ×2":     s.dictuAvg     ? +(s.dictuAvg * 2).toFixed(1) : 0
      };
    });

    return (
      <div className="h-full overflow-y-auto" style={{ background:"#EBF3FF" }}>
        <div className="p-5 max-w-5xl mx-auto">
          <h2 className="text-lg font-semibold mb-4" style={{ color:"#0C2340" }}>Applicatievergelijking</h2>

          <div className="grid grid-cols-2 gap-4 mb-4">
            <div className="rounded p-4" style={{ background:"#fff", border:"1px solid #D0E4F7" }}>
              <p className="text-sm font-semibold text-gray-700 mb-3">Scores — alle applicaties</p>
              <ResponsiveContainer width="100%" height={260}>
                <BarChart data={barData} margin={{ top:5, right:10, bottom:65, left:0 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#f3f4f6" />
                  <XAxis dataKey="name" tick={{ fontSize:9 }} angle={-35} textAnchor="end" interval={0} />
                  <YAxis domain={[0,10]} tick={{ fontSize:9 }} />
                  <Tooltip wrapperStyle={{ fontSize:11 }} />
                  <Legend wrapperStyle={{ fontSize:10 }} />
                  <Bar dataKey="Autonomiescore" radius={[3,3,0,0]} name="Autonomiescore (1-10)">
                    {barData.map((d,i) => <Cell key={i} fill={scoreColor(d.Autonomiescore)} />)}
                  </Bar>
                  <Bar dataKey="DICTU ×2" fill="#7c3aed" fillOpacity={0.7} radius={[3,3,0,0]} name="DICTU score ×2 (schaal 0-10)" />
                </BarChart>
              </ResponsiveContainer>
            </div>

            <div className="rounded p-4" style={{ background:"#fff", border:"1px solid #D0E4F7" }}>
              <p className="text-sm font-semibold text-gray-700 mb-3">DAAF Radar — per dimensie</p>
              <ResponsiveContainer width="100%" height={260}>
                <RadarChart data={radarData}>
                  <PolarGrid stroke="#f3f4f6" />
                  <PolarAngleAxis dataKey="dim" tick={{ fontSize:9 }} />
                  <PolarRadiusAxis domain={[0,5]} tick={false} axisLine={false} />
                  {apps.slice(0,5).map((a,i) => (
                    <Radar key={a.id} name={name14(a.name)} dataKey={name14(a.name)}
                      stroke={COLORS[i%COLORS.length]} fill={COLORS[i%COLORS.length]} fillOpacity={0.12} strokeWidth={2} />
                  ))}
                  <Tooltip wrapperStyle={{ fontSize:11 }} />
                  <Legend wrapperStyle={{ fontSize:10 }} />
                </RadarChart>
              </ResponsiveContainer>
            </div>
          </div>

          {/* Table */}
          <div className="rounded p-4 overflow-x-auto" style={{ background:"#fff", border:"1px solid #D0E4F7" }}>
            <p className="text-sm font-semibold mb-3" style={{ color:"#0C2340" }}>Vergelijkingstabel</p>
            <table className="w-full" style={{ fontSize:12 }}>
              <thead>
                <tr style={{ borderBottom:`2px solid #1A56A0` }}>
                  {["Applicatie","Leverancier","Auto-score","Risico","Mitigatie","Belang","DICTU","Volledigheid"].map(h => (
                    <th key={h} className="text-left py-1.5 px-2 font-semibold" style={{ color:"#1A56A0" }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {apps.map(a => {
                  const s   = calcScores(a.scores);
                  const lbl = scoreLabel(s.autonomyScore);
                  return (
                    <tr key={a.id} style={{ borderBottom:"1px solid #EBF3FF" }}
                      className="cursor-pointer transition-colors"
                      style={{ borderBottom:"1px solid #EBF3FF", cursor:"pointer" }}
                      onClick={() => { setSelId(a.id); setStep(0); setView("assess"); }}>
                      <td className="py-2 px-2 font-medium" style={{ color:"#0C2340" }}>{a.name}</td>
                      <td className="py-2 px-2 text-gray-500">{a.supplier}</td>
                      <td className="py-2 px-2">
                        <span className="px-2 py-0.5 font-semibold text-xs"
                          style={{ borderRadius:3, background:lbl.bg, color:lbl.fg }}>
                          {s.autonomyScore ? s.autonomyScore.toFixed(1) : "–"}
                        </span>
                      </td>
                      <td className="py-2 px-2 text-gray-600">{s.risico    ? s.risico.toFixed(1)    : "–"}</td>
                      <td className="py-2 px-2 text-gray-600">{s.mitigatie ? s.mitigatie.toFixed(1) : "–"}</td>
                      <td className="py-2 px-2 text-gray-600">{s.belang    ? s.belang.toFixed(1)    : "–"}</td>
                      <td className="py-2 px-2 text-gray-600">{s.dictuAvg  ? s.dictuAvg.toFixed(1)  : "–"}/5</td>
                      <td className="py-2 px-2 text-gray-600">{s.completeness}%</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    );
  }

  // ── ADMIN ─────────────────────────────────────────────────

  function Admin() {
    // PIN lock screen
    if (!adminUnlocked) {
      return (
        <div className="h-full flex items-center justify-center" style={{ background:"#EBF3FF" }}>
          <div className="bg-white p-8 w-full max-w-sm" style={{ borderRadius:4, boxShadow:"0 4px 24px rgba(12,35,64,0.15)", border:"1px solid #D0E4F7" }}>
            <div className="flex items-center gap-3 mb-6">
              <div className="w-10 h-10 flex items-center justify-center" style={{ background:"#0C2340", borderRadius:4 }}>
                <span style={{ fontSize:20 }}>🔐</span>
              </div>
              <div>
                <h2 className="font-bold" style={{ color:"#0C2340" }}>Beheeromgeving</h2>
                <p className="text-xs text-gray-400">Voer de beheerpincode in</p>
              </div>
            </div>
            <input
              type="password"
              value={adminPin}
              onChange={e => { setAdminPin(e.target.value); setAdminPinError(false); }}
              onKeyDown={e => {
                if (e.key === "Enter") {
                  if (adminPin === ADMIN_PIN) { setAdminUnlocked(true); setAdminPin(""); }
                  else { setAdminPinError(true); setAdminPin(""); }
                }
              }}
              placeholder="Pincode"
              className="w-full border px-3 py-2.5 text-sm focus:outline-none mb-2"
              style={{ borderColor: adminPinError ? "#dc2626" : "#D0E4F7", borderRadius:4, letterSpacing:4 }}
              autoFocus
            />
            {adminPinError && (
              <p className="text-xs mb-3" style={{ color:"#dc2626" }}>Pincode onjuist. Probeer opnieuw.</p>
            )}
            <button
              onClick={() => {
                if (adminPin === ADMIN_PIN) { setAdminUnlocked(true); setAdminPin(""); }
                else { setAdminPinError(true); setAdminPin(""); }
              }}
              className="w-full text-white py-2.5 text-sm font-semibold"
              style={{ background:"#1A56A0", borderRadius:4 }}>
              Toegang
            </button>
            <button onClick={() => setView("dashboard")}
              className="w-full py-2 text-sm mt-2"
              style={{ color:"#9ca3af" }}>
              Terug naar dashboard
            </button>
          </div>
        </div>
      );
    }

    // Edit modal
    const editApp = editAppId ? apps.find(a => a.id === editAppId) : null;

    return (
      <div className="h-full overflow-y-auto" style={{ background:"#EBF3FF" }}>
        <div className="p-5 max-w-5xl mx-auto">

          {/* Header strip */}
          <div className="flex items-center justify-between mb-5">
            <div className="flex items-center gap-3">
              <div className="w-8 h-8 flex items-center justify-center" style={{ background:"#0C2340", borderRadius:4 }}>
                <span style={{ fontSize:16 }}>🔐</span>
              </div>
              <div>
                <h2 className="font-bold" style={{ color:"#0C2340" }}>Beheeromgeving</h2>
                <p className="text-xs text-gray-400">{apps.length} applicatie{apps.length !== 1 ? "s" : ""} in het systeem</p>
              </div>
            </div>
            <button onClick={() => { setAdminUnlocked(false); setView("dashboard"); }}
              className="text-xs px-3 py-1.5 font-medium"
              style={{ border:"1px solid #D0E4F7", borderRadius:4, color:"#6b7280", background:"#fff" }}>
              🔒 Vergrendelen
            </button>
          </div>

          {apps.length === 0 ? (
            <div className="text-center py-16 bg-white rounded" style={{ border:"2px dashed #D0E4F7", color:"#9ca3af" }}>
              Nog geen applicaties in het systeem.
            </div>
          ) : (
            <div className="space-y-3">
              {apps.map(a => {
                const sc  = calcScores(a.scores);
                const lbl = scoreLabel(sc.autonomyScore);
                const allQ = [...DAAF, ...DICTU];
                const filled = allQ.filter(q => (a.scores[q.key] || 0) > 0);

                return (
                  <div key={a.id} className="bg-white rounded" style={{ border:"1px solid #D0E4F7", borderLeft:`4px solid ${scoreColor(sc.autonomyScore)}` }}>
                    {/* App header row */}
                    <div className="flex items-center gap-4 p-4">
                      <Gauge score={sc.autonomyScore} size={60} />
                      <div className="flex-1 min-w-0">
                        <div className="flex items-center gap-2 flex-wrap">
                          <h3 className="font-bold" style={{ color:"#0C2340" }}>{a.name}</h3>
                          {a.supplier && <span className="text-xs text-gray-400">· {a.supplier}</span>}
                          {a.cat && <span className="text-xs px-2 py-0.5 rounded" style={{ background:"#EBF3FF", color:"#1A56A0" }}>{a.cat}</span>}
                        </div>
                        <div className="flex items-center gap-3 mt-1 flex-wrap">
                          {a.owner && <span className="text-xs text-gray-400">👤 {a.owner}</span>}
                          <span className="text-xs px-2 py-0.5 font-medium" style={{ borderRadius:3, background:lbl.bg, color:lbl.fg }}>{lbl.text}</span>
                          <span className="text-xs text-gray-400">{sc.completeness}% ingevuld · {filled.length}/{allQ.length} vragen</span>
                          <span className="text-xs text-gray-400">Aangemaakt: {new Date(a.createdAt).toLocaleDateString("nl-NL")}</span>
                        </div>
                        {a.notes && <p className="text-xs text-gray-400 mt-1 italic">"{a.notes}"</p>}
                      </div>
                      {/* Action buttons */}
                      <div className="flex gap-2 flex-shrink-0">
                        <button
                          onClick={() => { setSelId(a.id); setStep(0); setView("assess"); }}
                          className="text-white text-xs px-3 py-1.5 font-medium"
                          style={{ background:"#1A56A0", borderRadius:4 }}>
                          ✏️ Invullen
                        </button>
                        <button
                          onClick={() => {
                            setEditAppId(a.id);
                            setEditForm({ name:a.name, cat:a.cat, supplier:a.supplier, owner:a.owner, notes:a.notes });
                          }}
                          className="text-xs px-3 py-1.5 font-medium"
                          style={{ border:"1px solid #D0E4F7", borderRadius:4, color:"#1A56A0", background:"#fff" }}>
                          ⚙️ Gegevens
                        </button>
                        <button
                          onClick={() => {
                            if (confirm(`"${a.name}" definitief verwijderen? Dit kan niet ongedaan worden gemaakt.`)) {
                              setApps(p => p.filter(x => x.id !== a.id));
                              if (selId === a.id) setSelId(null);
                            }
                          }}
                          className="text-xs px-3 py-1.5 font-medium"
                          style={{ border:"1px solid #fecaca", borderRadius:4, color:"#dc2626", background:"#fff" }}>
                          🗑 Verwijder
                        </button>
                      </div>
                    </div>

                    {/* Score detail strip */}
                    <div className="grid grid-cols-6 gap-0" style={{ borderTop:"1px solid #EBF3FF" }}>
                      {[
                        { label:"Risico",    val:sc.risico,    color:"#dc2626", hint:"↓ beter" },
                        { label:"Mitigatie", val:sc.mitigatie, color:"#26B5AE", hint:"↑ beter" },
                        { label:"Belang",    val:sc.belang,    color:"#E87722", hint:"↓ beter" },
                        { label:"DICTU",     val:sc.dictuAvg,  color:"#1A56A0", hint:"/5" },
                        { label:"Autonomie", val:sc.autonomyScore, color:scoreColor(sc.autonomyScore), hint:"/10" },
                        { label:"Volledig",  val:sc.completeness, color:"#6b7280", hint:"%" },
                      ].map(({ label, val, color, hint }) => (
                        <div key={label} className="py-2 px-3 text-center" style={{ borderRight:"1px solid #EBF3FF" }}>
                          <div style={{ fontSize:15, fontWeight:700, color: val ? color : "#d1d5db" }}>
                            {val ? (label === "Volledig" ? val : val.toFixed(1)) : "–"}
                            <span style={{ fontSize:9, color:"#9ca3af", marginLeft:1 }}>{hint}</span>
                          </div>
                          <div style={{ fontSize:10, color:"#9ca3af" }}>{label}</div>
                        </div>
                      ))}
                    </div>

                    {/* Ingevulde scores per vraag */}
                    <div className="px-4 py-3" style={{ borderTop:"1px solid #EBF3FF" }}>
                      <p className="text-xs font-semibold mb-2" style={{ color:"#0C2340" }}>Scores per vraag</p>
                      <div className="flex flex-wrap gap-1.5">
                        {allQ.map(q => {
                          const s = a.scores[q.key] || 0;
                          return (
                            <div key={q.key} title={`${q.key}: ${q.name}\nScore: ${s || "niet ingevuld"}`}
                              className="flex items-center gap-1 px-2 py-1 text-xs"
                              style={{ borderRadius:3, background: s ? "#EBF3FF" : "#f9fafb", border:"1px solid #D0E4F7", color:"#0C2340" }}>
                              <span style={{ color:"#1A56A0", fontWeight:600 }}>{q.key}</span>
                              <span style={{ fontWeight:700, color: s ? scoreColor(s, 5) : "#d1d5db" }}>{s || "–"}</span>
                            </div>
                          );
                        })}
                      </div>
                    </div>
                  </div>
                );
              })}
            </div>
          )}

          {/* Danger zone */}
          {apps.length > 0 && (
            <div className="mt-6 p-4 rounded" style={{ border:"1px solid #fecaca", background:"#fff5f5" }}>
              <p className="text-xs font-semibold mb-1" style={{ color:"#dc2626" }}>⚠️ Gevaarlijke zone</p>
              <p className="text-xs text-gray-500 mb-3">Hiermee worden ALLE applicaties en scores definitief gewist. Niet terug te draaien.</p>
              <button
                onClick={() => {
                  if (confirm("ALLE applicaties verwijderen? Dit wist alle assessmentdata permanent.")) {
                    if (confirm("Weet u het zeker? Dit kan NIET ongedaan worden gemaakt.")) {
                      setApps([]);
                      setSelId(null);
                    }
                  }
                }}
                className="text-xs px-4 py-2 font-semibold text-white"
                style={{ background:"#dc2626", borderRadius:4 }}>
                🗑 Alles wissen
              </button>
            </div>
          )}
        </div>

        {/* Edit metadata modal */}
        {editApp && (
          <div className="fixed inset-0 flex items-center justify-center z-50 p-4"
            style={{ background:"rgba(12,35,64,0.6)" }}>
            <div className="bg-white w-full max-w-md" style={{ borderRadius:4, boxShadow:"0 8px 32px rgba(12,35,64,0.3)" }}>
              <div className="px-6 py-4" style={{ borderBottom:"3px solid #1A56A0" }}>
                <div className="flex items-center justify-between">
                  <div className="flex items-center gap-3">
                    <div className="w-1 h-5" style={{ background:"#26B5AE", borderRadius:2 }}/>
                    <h2 className="text-base font-semibold" style={{ color:"#0C2340" }}>Applicatiegegevens aanpassen</h2>
                  </div>
                  <button onClick={() => setEditAppId(null)} style={{ color:"#9ca3af", fontSize:18 }}>✕</button>
                </div>
              </div>
              <div className="p-6">
                <div className="space-y-3">
                  {[
                    { k:"name",     l:"Applicatienaam *" },
                    { k:"supplier", l:"Leverancier" },
                    { k:"cat",      l:"Categorie" },
                    { k:"owner",    l:"Applicatie-eigenaar" },
                  ].map(f => (
                    <div key={f.k}>
                      <label className="text-xs font-semibold block mb-1" style={{ color:"#0C2340" }}>{f.l}</label>
                      <input
                        value={editForm[f.k] || ""}
                        onChange={e => setEditForm(p => ({ ...p, [f.k]: e.target.value }))}
                        className="w-full border px-3 py-2 text-sm focus:outline-none"
                        style={{ borderColor:"#D0E4F7", borderRadius:4 }}
                        onFocus={e => e.target.style.borderColor="#1A56A0"}
                        onBlur={e => e.target.style.borderColor="#D0E4F7"}
                      />
                    </div>
                  ))}
                  <div>
                    <label className="text-xs font-semibold block mb-1" style={{ color:"#0C2340" }}>Toelichting</label>
                    <textarea
                      value={editForm.notes || ""}
                      onChange={e => setEditForm(p => ({ ...p, notes: e.target.value }))}
                      rows={2}
                      className="w-full border px-3 py-2 text-sm focus:outline-none"
                      style={{ borderColor:"#D0E4F7", borderRadius:4 }}/>
                  </div>
                </div>
                <div className="flex gap-3 mt-5">
                  <button onClick={() => setEditAppId(null)}
                    className="flex-1 py-2 text-sm"
                    style={{ border:"1px solid #D0E4F7", borderRadius:4, color:"#6b7280" }}>
                    Annuleren
                  </button>
                  <button
                    disabled={!editForm.name?.trim()}
                    onClick={() => {
                      setApps(p => p.map(a => a.id === editAppId
                        ? { ...a, name:editForm.name, cat:editForm.cat, supplier:editForm.supplier, owner:editForm.owner, notes:editForm.notes }
                        : a
                      ));
                      setEditAppId(null);
                    }}
                    className="flex-1 text-white py-2 text-sm font-semibold"
                    style={{ background: editForm.name?.trim() ? "#1A56A0" : "#d1d5db", borderRadius:4 }}>
                    Opslaan
                  </button>
                </div>
              </div>
            </div>
          </div>
        )}
      </div>
    );
  }

  // ── ABOUT ─────────────────────────────────────────────────

  function About() {
    const Section = ({ title, children, accent="#1A56A0" }) => (
      <div className="mb-6">
        <div className="flex items-center gap-2 mb-3">
          <div className="w-1 h-5 rounded-full flex-shrink-0" style={{ background:accent }}/>
          <h2 className="font-bold text-base" style={{ color:"#0C2340" }}>{title}</h2>
        </div>
        {children}
      </div>
    );

    const Tip = ({ label, children, color="#1A56A0", bg="#EBF3FF" }) => (
      <div className="rounded p-3 mb-2" style={{ background:bg, border:`1px solid ${color}33` }}>
        <p className="text-xs font-bold mb-1" style={{ color }}>{label}</p>
        <p className="text-xs leading-relaxed" style={{ color:"#374151" }}>{children}</p>
      </div>
    );

    return (
      <div className="h-full overflow-y-auto" style={{ background:"#EBF3FF" }}>
        <div className="p-5 max-w-4xl mx-auto">

          {/* Hero */}
          <div className="rounded p-6 mb-6 text-white" style={{ background:"linear-gradient(135deg, #0C2340 0%, #1A56A0 100%)" }}>
            <div className="flex items-center gap-3 mb-3">
              <div className="px-3 py-2 border-2 border-white" style={{ borderRadius:2 }}>
                <span className="font-bold leading-none" style={{ fontSize:10, letterSpacing:1 }}>NHL<br/>STENDEN</span>
              </div>
              <div className="w-px self-stretch" style={{ background:"#26B5AE", margin:"2px 0" }}/>
              <div>
                <h1 className="font-bold" style={{ fontSize:17 }}>Digitale Soevereiniteitsassessment</h1>
                <p style={{ fontSize:12, color:"#7DD3D0" }}>Project Digitale Soevereiniteit · Ambassadeurslijn Digitale Soevereiniteit</p>
              </div>
            </div>
            <p className="text-sm leading-relaxed" style={{ color:"rgba(255,255,255,0.85)" }}>
              Dit instrument helpt het Transitieteam Digitalisering van NHL Stenden Hogeschool om per applicatie
              te beoordelen hoe urgent het autonomieprobleem is en in hoeverre de instelling digitaal soeverein opereert.
              De tool combineert twee erkende frameworks: <strong>DAAF</strong> (Universiteit Utrecht) en <strong>DICTU</strong> (Rijksoverheid).
            </p>
          </div>

          {/* Frameworks */}
          <Section title="Gebruikte frameworks en wat ze meten">
            <div className="grid grid-cols-2 gap-3 mb-2">
              <div className="rounded p-4" style={{ background:"#fff", border:"1px solid #D0E4F7" }}>
                <div className="flex items-center gap-2 mb-2">
                  <span className="text-xs font-bold px-2 py-0.5" style={{ background:"#1A56A0", color:"#fff", borderRadius:3 }}>DAAF</span>
                  <span className="font-semibold text-sm" style={{ color:"#0C2340" }}>Digital Autonomy Assessment Framework</span>
                </div>
                <p className="text-xs leading-relaxed mb-2" style={{ color:"#6b7280" }}>
                  Ontwikkeld door de Universiteit Utrecht. Het DAAF Framework beoordeelt de digitale autonomie
                  van een organisatie ten opzichte van haar leveranciers. Het werkt met drie niveaus:
                </p>
                <div className="space-y-1">
                  {[
                    { lv:"Niveau 1 · Risico (A, B)",      txt:"Hoe groot is het externe risico — jurisdictie leverancier (A1), hosting & datalocatie (A3) en vendor concentratie (B1)?", c:"#dc2626" },
                    { lv:"Niveau 2 · Mitigatie (C, D, E)",txt:"Hoe goed kan NHL Stenden risico's beheersen — alternatieven, kennis, contracten?", c:"#26B5AE" },
                    { lv:"Niveau 3 · Belang (F, G, H)",   txt:"Hoe kritiek is de applicatie — operationeel, data en strategisch belang?",         c:"#E87722" },
                  ].map(r => (
                    <div key={r.lv} className="flex gap-2 text-xs">
                      <span className="w-2 h-2 rounded-full mt-1 flex-shrink-0" style={{ background:r.c }}/>
                      <div><strong style={{ color:r.c }}>{r.lv}:</strong> {r.txt}</div>
                    </div>
                  ))}
                </div>
                <p className="text-xs mt-2 italic" style={{ color:"#9ca3af" }}>Elke dimensie scoort 1–5. De autonomiescore (1–10) volgt uit de formule Mitigatie ÷ (Risico × Belang).</p>
              </div>

              <div className="rounded p-4" style={{ background:"#fff", border:"1px solid #26B5AE44" }}>
                <div className="flex items-center gap-2 mb-2">
                  <span className="text-xs font-bold px-2 py-0.5" style={{ background:"#26B5AE", color:"#fff", borderRadius:3 }}>DICTU</span>
                  <span className="font-semibold text-sm" style={{ color:"#0C2340" }}>Dienst ICT Uitvoering</span>
                </div>
                <p className="text-xs leading-relaxed mb-2" style={{ color:"#6b7280" }}>
                  Framework van het Ministerie van EZK (Rijksoverheid) voor digitale soevereiniteit van clouddiensten.
                  Vier vragen beoordelen juridische en technische bescherming van klantdata:
                </p>
                <div className="space-y-1">
                  {[
                    { k:"2.1 Dataresidency",               txt:"Staat alle data (incl. back-ups) uitsluitend in de EU?" },
                    { k:"2.2 Technische beveiliging",       txt:"Zijn er verifieerbare garanties dat niemand — ook de aanbieder niet — de data kan inzien?" },
                    { k:"2.3 Juridische bescherming",       txt:"Verzet de aanbieder zich actief tegen niet-EU dataverzoeken en meldt hij dit?" },
                    { k:"4.1 EU-infrastructuur",            txt:"Bevindt ook de control plane (beheer van de clouddienst) zich volledig in de EU?" },
                  ].map(r => (
                    <div key={r.k} className="flex gap-2 text-xs">
                      <span className="font-bold flex-shrink-0" style={{ color:"#26B5AE", minWidth:32 }}>{r.k.split(" ")[0]}</span>
                      <div><strong>{r.k.split(" ").slice(1).join(" ")}:</strong> {r.txt}</div>
                    </div>
                  ))}
                </div>
                <p className="text-xs mt-2 italic" style={{ color:"#9ca3af" }}>Score 1–5 per vraag. Gemiddelde = soevereiniteitsscore op de kleurenbalk (rood → groen).</p>
              </div>
            </div>
          </Section>

          {/* Wat zegt de score */}
          <Section title="Wat zegt de score? — De urgentie van het autonomieprobleem" accent="#E87722">
            <div className="rounded p-4 mb-3" style={{ background:"#fff", border:"2px solid #E87722" }}>
              <p className="text-sm font-semibold mb-2" style={{ color:"#0C2340" }}>
                De autonomiescore is géén maat voor "hoe digitaal autonoom is deze applicatie" —
                maar voor <em>"hoe urgent is het autonomieprobleem?"</em>
              </p>
              <p className="text-xs leading-relaxed mb-3" style={{ color:"#374151" }}>
                De score combineert drie dingen: hoe groot is het risico, hoe goed kun je het beheersen, en hoe belangrijk is de applicatie.
                Twee applicaties met dezelfde score kunnen een heel verschillende situatie beschrijven.
              </p>
              <div className="grid grid-cols-2 gap-3">
                <div className="rounded p-3" style={{ background:"#dcfce7", border:"1px solid #86efac" }}>
                  <p className="text-xs font-bold mb-1" style={{ color:"#15803d" }}>✓ Hoge score (7+) — geen acute actie</p>
                  <p className="text-xs" style={{ color:"#166534" }}>
                    Kan twee dingen betekenen: de risico's zijn goed gemitigeerd (alternatieven beschikbaar,
                    sterke contracten, kennis in huis), óf de applicatie heeft weinig strategisch belang.
                    Beide situaties zijn positief: er is geen acute actie nodig. Monitor periodiek.
                  </p>
                </div>
                <div className="rounded p-3" style={{ background:"#fee2e2", border:"1px solid #fca5a5" }}>
                  <p className="text-xs font-bold mb-1" style={{ color:"#b91c1c" }}>✗ Lage score (&lt;5) — actie vereist</p>
                  <p className="text-xs" style={{ color:"#7f1d1d" }}>
                    Er is een combinatie van hoog risico, zwakke mitigatie en/of hoog strategisch belang.
                    Kies één van drie acties: <strong>migreren</strong> (lagere risicoleverancier),
                    <strong> mitigeren</strong> (weerbaarheid opbouwen) of <strong>bewust accepteren</strong> met besluitvorming.
                  </p>
                </div>
              </div>
            </div>
            <Tip label="Waar zit de verbetermarge?" color="#1A56A0" bg="#EBF3FF">
              Strategisch belang (Niveau 3) ligt grotendeels vast — een applicatie wordt niet minder belangrijk.
              Verbeteringen zitten in Risico-exposure verlagen (Niveau 1, bijv. migreren naar EU-aanbieder)
              en Mitigatie-capaciteit verhogen (Niveau 2, bijv. alternatieven ontwikkelen, contracten versterken, kennis opbouwen).
              Gebruik de drie niveauscores apart in het dashboard om te bepalen waar de grootste winst te behalen is.
            </Tip>
          </Section>

          {/* Hoe werkt de applicatie */}
          <Section title="Hoe werkt de applicatie — stap voor stap">
            <div className="space-y-2">
              {[
                { n:"1", title:"Applicatie toevoegen", icon:"➕",
                  txt:'Klik op "+ Applicatie toevoegen" in het dashboard of via het tabblad Applicaties. Vul naam, leverancier, categorie en eigenaar in. De applicatie wordt direct opgeslagen.' },
                { n:"2", title:"Assessment invullen — DAAF Quick Scan", icon:"📋",
                  txt:'Ga naar het tabblad van de applicatie. Stap 1 toont de 8 DAAF-vragen verdeeld over drie niveaus. Klik op de gewenste score (1–5) per vraag. De live sidebar rechtsboven toont direct hoe de scores uitpakken.' },
                { n:"3", title:"Assessment invullen — DICTU Soevereiniteitscheck", icon:"🔍",
                  txt:'Stap 2 toont de 4 DICTU-vragen over datasoevereiniteit. Per vraag staat de norm vermeld (wat is het minimale vereiste niveau). Scores worden direct verwerkt in de kleurenbalk.' },
                { n:"4", title:"Dashboard lezen", icon:"📊",
                  txt:'Het dashboard toont per applicatie de autonomiescore (snelheidsmeter, DAAF) en soevereiniteitsscore (kleurenbalk, DICTU). Het spindiagram vergelijkt alle applicaties per DAAF-dimensie. Het kwadrant plaatst applicaties op de assen Risico×Belang vs. Mitigatie.' },
                { n:"5", title:"Vergelijken", icon:"🔎",
                  txt:'Het tabblad Vergelijking toont een staafdiagram en tabel met alle applicaties naast elkaar. Klik op een rij om direct naar het assessment te gaan.' },
                { n:"6", title:"Excel exporteren", icon:"📥",
                  txt:'Klik op "Exporteer Excel" in de header. Het bestand bevat vier tabbladen: Overzicht (alle scores), DAAF scores, DICTU scores en de volledige vragenlijst met normen.' },
                { n:"7", title:"Beheer", icon:"🔐",
                  txt:'Via het tabblad Beheer (pincode vereist) kun je applicatiegegevens aanpassen, applicaties verwijderen en een volledig overzicht per applicatie bekijken inclusief alle ingevulde scores. Data wordt opgeslagen in de browser (localStorage).' },
              ].map(s => (
                <div key={s.n} className="flex gap-3 rounded p-3" style={{ background:"#fff", border:"1px solid #D0E4F7" }}>
                  <div className="w-8 h-8 flex items-center justify-center flex-shrink-0 text-white text-sm font-bold"
                    style={{ background:"#1A56A0", borderRadius:4 }}>{s.n}</div>
                  <div className="flex-1">
                    <p className="text-sm font-semibold" style={{ color:"#0C2340" }}>{s.icon} {s.title}</p>
                    <p className="text-xs leading-relaxed mt-0.5" style={{ color:"#6b7280" }}>{s.txt}</p>
                  </div>
                </div>
              ))}
            </div>
          </Section>

          {/* Berekening */}
          <Section title="De berekening in vier stappen" accent="#26B5AE">
            <div className="rounded p-4" style={{ background:"#fff", border:"1px solid #D0E4F7" }}>
              <div className="grid grid-cols-4 gap-3 mb-3">
                {[
                  { n:"1", lbl:"Dimensiescore", txt:"Per dimensie (A t/m H): gemiddelde van de ingevulde indicatorscores (1–5).", c:"#1A56A0" },
                  { n:"2", lbl:"Niveauscore",   txt:"Per niveau (Risico, Mitigatie, Belang): gemiddelde van de bijbehorende dimensiescores.", c:"#26B5AE" },
                  { n:"3", lbl:"Ruwe score",    txt:"Mitigatie ÷ (Risico × Belang). Bereik: 0,04 (slechtst) tot 5,0 (best).", c:"#E87722" },
                  { n:"4", lbl:"Normalisatie",  txt:"Logaritmische schaal zet de ruwe score om naar 1–10, zodat verschillen zichtbaar en vergelijkbaar zijn.", c:"#6d28d9" },
                ].map(s => (
                  <div key={s.n} className="text-center rounded p-3" style={{ background:"#f8fafc", border:`2px solid ${s.c}33` }}>
                    <div className="w-7 h-7 rounded-full flex items-center justify-center text-white text-xs font-bold mx-auto mb-2"
                      style={{ background:s.c }}>{s.n}</div>
                    <p className="text-xs font-bold mb-1" style={{ color:s.c }}>{s.lbl}</p>
                    <p className="text-xs" style={{ color:"#6b7280" }}>{s.txt}</p>
                  </div>
                ))}
              </div>
              <div className="rounded p-2 text-xs" style={{ background:"#EBF3FF", color:"#374151" }}>
                <strong>Waarom logaritmisch?</strong> De formule deelt Mitigatie (max 5) door Risico×Belang (max 25).
                Daardoor liggen de meeste ruwe scores tussen 0,1 en 0,5. Zonder correctie zou 90% van de applicaties
                een eindscore tussen 1 en 3 krijgen — onbruikbaar voor vergelijking.
                De logaritmische schaal spreidt scores uit over het volledige bereik van 1 tot 10.
              </div>
            </div>
          </Section>

          {/* Praktische tips */}
          <Section title="Praktische tips voor het team">
            <div className="grid grid-cols-2 gap-3">
              <Tip label="💾 Data opslaan" color="#1A56A0" bg="#EBF3FF">
                Alle data wordt automatisch opgeslagen in de browser (localStorage). Je kunt het venster sluiten en later verdergaan.
                Exporteer regelmatig een Excel-backup. Let op: data is browsergebonden — gebruik altijd dezelfde browser op hetzelfde apparaat.
              </Tip>
              <Tip label="👥 Meerdere beoordelaars" color="#26B5AE" bg="#E6F7F7">
                Laat elke beoordelaar het assessment onafhankelijk invullen. Exporteer afzonderlijk naar Excel en
                vergelijk de scores. Bespreek grote afwijkingen in het team — die geven vaak de meest waardevolle inzichten.
              </Tip>
              <Tip label="📋 Welke vragen gebruik je?" color="#E87722" bg="#fff8e1">
                Deze tool gebruikt een selectie van het volledige DAAF en DICTU framework: de DAAF Quick Scan
                (8 kernindicatoren, één per dimensie) en 4 relevante DICTU-vragen voor soevereiniteit.
                Voor een volledig DAAF assessment met wegingen raadpleeg de Utrecht University tool.
              </Tip>
              <Tip label="🎯 Wat doe je met de uitkomst?" color="#6d28d9" bg="#faf5ff">
                Bespreek het kwadrant in het team. Applicaties rechtsonder (KRITIEK) vragen urgente actie.
                Bepaal per applicatie: migreren, mitigeren of bewust accepteren. Leg de keuze vast met een besluitdocument.
                Herhaal het assessment na significante contractwijzigingen of leveranciersveranderingen.
              </Tip>
            </div>
          </Section>

          {/* Footer */}
          <div className="rounded p-3 text-center" style={{ background:"#fff", border:"1px solid #D0E4F7" }}>
            <p className="text-xs" style={{ color:"#9ca3af" }}>
              NHL Stenden Hogeschool · Project Digitale Soevereiniteit · Ambassadeurslijn Digitale Soevereiniteit
              <br/>Ambassadeurs: J. Haije · E. Rolf · J. Blom · Kwartiermaker: E. van Gorkum · v0.5
            </p>
          </div>

        </div>
      </div>
    );
  }

  // ── SHELL ──────────────────────────────────────────────────

  if (!ready) return (
    <div className="flex items-center justify-center h-screen text-sm" style={{ background:"#EBF3FF", color:"#6b7280" }}>
      Laden...
    </div>
  );

  return (
    <div className="flex flex-col h-screen" style={{ fontFamily:"system-ui,sans-serif", background:"#EBF3FF" }}>
      {/* Header — NHL Stenden huisstijl */}
      <header className="flex items-center justify-between px-5 flex-shrink-0"
        style={{ background:"#0C2340", minHeight:56 }}>
        {/* Logo box (kenmerkend NHL Stenden) */}
        <div className="flex items-center gap-4">
          <div className="flex items-center justify-center px-3 py-2 border-2 border-white"
            style={{ minWidth:72 }}>
            <span className="text-white font-bold leading-none" style={{ fontSize:11, letterSpacing:1 }}>
              NHL<br/>STENDEN
            </span>
          </div>
          {/* Teal accent bar */}
          <div className="w-1 self-stretch" style={{ background:"#26B5AE", margin:"8px 0" }} />
          <div>
            <h1 className="font-bold text-white" style={{ fontSize:13 }}>Digitale Soevereiniteitsassessment</h1>
            <p style={{ fontSize:11, color:"#7DD3D0" }}>
              Project Digitale Soevereiniteit · Ambassadeurslijn Digitale Soevereiniteit
            </p>
            <p style={{ fontSize:10, color:"rgba(125,211,208,0.7)", marginTop:1 }}>
              Ambassadeurs: J. Haije · E. Rolf · J. Blom · Kwartiermaker: E. van Gorkum · <span style={{ color:"#26B5AE", fontWeight:600 }}>v0.5</span>
            </p>
          </div>
        </div>
        <div className="flex items-center gap-3">
          {/* Opslaan status */}
          {saving && (
            <span className="text-xs px-2 py-1 rounded" style={{ background:"rgba(255,255,255,0.15)", color:"#7DD3D0" }}>
              ⏳ Opslaan…
            </span>
          )}
          {saveError && (
            <span className="text-xs px-2 py-1 rounded" style={{ background:"rgba(220,38,38,0.3)", color:"#fca5a5" }}>
              ⚠️ Opslaan mislukt
            </span>
          )}
          {!saving && !saveError && ready && (
            <span className="text-xs px-2 py-1 rounded" style={{ background:"rgba(38,181,174,0.2)", color:"#7DD3D0" }}>
              ✓ Gesynchroniseerd
            </span>
          )}
          <button onClick={exportXlsx} disabled={apps.length === 0}
            className="flex items-center gap-2 text-white text-xs px-4 py-2 font-medium transition-all"
            style={{
              background: apps.length === 0 ? "rgba(255,255,255,0.1)" : "#26B5AE",
              opacity: apps.length === 0 ? 0.5 : 1,
              borderRadius: 4
            }}>
            📥 Exporteer Excel
          </button>
        </div>
      </header>

      {/* Sub-header nav — lichte blauwe balk (NHS Stenden stijl) */}
      <div className="flex-shrink-0" style={{ background:"#1A56A0" }}>
        <nav className="px-5 flex gap-0">
          {[
            { k:"dashboard", label:"Dashboard" },
            { k:"apps",      label:"Applicaties" },
            ...(selApp ? [{ k:"assess", label:selApp.name.substring(0,20) }] : []),
            { k:"compare",   label:"Vergelijking" },
            { k:"about",     label:"ℹ️ Over & uitleg" },
            { k:"admin",     label:"🔐 Beheer" },
          ].map(t => (
            <button key={t.k} onClick={() => setView(t.k)}
              className="px-4 py-2.5 text-sm font-medium border-b-2 transition-colors whitespace-nowrap"
              style={view === t.k
                ? { borderColor:"#26B5AE", color:"#fff", background:"rgba(255,255,255,0.12)" }
                : { borderColor:"transparent", color:"rgba(255,255,255,0.75)" }}>
              {t.label}
            </button>
          ))}
        </nav>
      </div>

      {/* Content */}
      <main className="flex-1 overflow-hidden" style={{ background:"#EBF3FF" }}>
        {view === "dashboard" && Dashboard()}
        {view === "apps"      && AppsList()}
        {view === "assess"    && Assess()}
        {view === "compare"   && Compare()}
        {view === "about"     && About()}
        {view === "admin"     && Admin()}
      </main>

      {/* Add modal */}
      {showModal && (
        <div className="fixed inset-0 flex items-center justify-center z-50 p-4"
          style={{ background:"rgba(12,35,64,0.6)" }}>
          <div className="bg-white w-full max-w-md" style={{ borderRadius:4, boxShadow:"0 8px 32px rgba(12,35,64,0.3)" }}>
            {/* Modal header */}
            <div className="px-6 py-4" style={{ borderBottom:"3px solid #1A56A0" }}>
              <div className="flex items-center gap-3">
                <div className="w-1 h-5" style={{ background:"#26B5AE", borderRadius:2 }}/>
                <h2 className="text-base font-semibold" style={{ color:"#0C2340" }}>Applicatie toevoegen</h2>
              </div>
            </div>
            <div className="p-6">
              <div className="space-y-3">
                {[
                  { k:"name",     l:"Applicatienaam *", p:"bijv. Microsoft 365" },
                  { k:"supplier", l:"Leverancier",       p:"bijv. Microsoft" },
                  { k:"cat",      l:"Categorie",          p:"bijv. Productiviteit, ERP, SIS" },
                  { k:"owner",    l:"Applicatie-eigenaar",p:"bijv. Functioneel beheerder" },
                ].map(f => (
                  <div key={f.k}>
                    <label className="text-xs font-semibold block mb-1" style={{ color:"#0C2340" }}>{f.l}</label>
                    <input value={form[f.k]} onChange={e => setForm(p => ({ ...p, [f.k]: e.target.value }))}
                      placeholder={f.p}
                      className="w-full border px-3 py-2 text-sm focus:outline-none"
                      style={{ borderColor:"#D0E4F7", borderRadius:4 }}
                      onFocus={e => e.target.style.borderColor="#1A56A0"}
                      onBlur={e => e.target.style.borderColor="#D0E4F7"}
                    />
                  </div>
                ))}
                <div>
                  <label className="text-xs font-semibold block mb-1" style={{ color:"#0C2340" }}>Toelichting</label>
                  <textarea value={form.notes} onChange={e => setForm(p => ({ ...p, notes: e.target.value }))}
                    rows={2} placeholder="Optionele context of notities"
                    className="w-full border px-3 py-2 text-sm focus:outline-none"
                    style={{ borderColor:"#D0E4F7", borderRadius:4 }}/>
                </div>
              </div>
              <div className="flex gap-3 mt-5">
                <button onClick={() => setShowModal(false)}
                  className="flex-1 py-2 text-sm"
                  style={{ border:"1px solid #D0E4F7", borderRadius:4, color:"#6b7280" }}>
                  Annuleren
                </button>
                <button onClick={addApp} disabled={!form.name.trim()}
                  className="flex-1 text-white py-2 text-sm font-semibold"
                  style={{ background: form.name.trim() ? "#1A56A0" : "#d1d5db", borderRadius:4 }}>
                  Toevoegen &amp; starten
                </button>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
