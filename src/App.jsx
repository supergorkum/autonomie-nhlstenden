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
  { key:"A1", dim:"A", level:"Risico", dimName:"Hosting & Infrastructuur",
    name:"Hostinglocatie", hint:"Score 1 = weinig risico · Score 5 = hoog risico",
    question:"Waar bevinden de servers en klantdata van deze applicatie zich fysiek?",
    scores:[
      {s:1,label:"Volledig EU",    desc:"Servers en data volledig in de EU. Geen overdracht buiten EU."},
      {s:2,label:"Primair EU",     desc:"Primair EU-gehost, enkele uitzonderingen buiten EU."},
      {s:3,label:"Gemengd",        desc:"Gemengde hosting: zowel EU- als niet-EU-locaties."},
      {s:4,label:"Primair non-EU", desc:"Primair buiten EU. Deel van data mogelijk in EU."},
      {s:5,label:"Volledig non-EU",desc:"Servers en data volledig buiten de EU. Geen EU-garantie."}
    ]
  },
  { key:"B1", dim:"B", level:"Risico", dimName:"Vendor Lock-in",
    name:"Leveranciersafhankelijkheid", hint:"Score 1 = weinig risico · Score 5 = hoog risico",
    question:"In hoeverre is NHL Stenden afhankelijk van deze specifieke leverancier?",
    scores:[
      {s:1,label:"Geen lock-in",      desc:"Open standaarden. Eenvoudig en goedkoop te vervangen."},
      {s:2,label:"Lichte afh.",        desc:"Lichte afhankelijkheid. Alternatieven beschikbaar, lage switching costs."},
      {s:3,label:"Matige afh.",        desc:"Matige afhankelijkheid. Alternatieven mogelijk maar tijdrovend."},
      {s:4,label:"Hoge afh.",          desc:"Hoge afhankelijkheid. Beperkte alternatieven, hoge migratie-inspanning."},
      {s:5,label:"Volledige lock-in",  desc:"Volledige vendor lock-in. Geen realistisch alternatief beschikbaar."}
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
  const risico    = avg(["A1","B1"]);
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
  const [apps,      setApps]      = useState([]);
  const [ready,     setReady]     = useState(false);
  const [view,      setView]      = useState("dashboard");
  const [selId,     setSelId]     = useState(null);
  const [step,      setStep]      = useState(0);
  const [showModal, setShowModal] = useState(false);
  const [form,      setForm]      = useState({ name:"", cat:"", supplier:"", owner:"", notes:"" });

  useEffect(() => {
    try {
      const stored = localStorage.getItem("nhl_sov_v2");
      if (stored) setApps(JSON.parse(stored));
    } catch {}
    setReady(true);
  }, []);

  useEffect(() => {
    if (!ready) return;
    try {
      localStorage.setItem("nhl_sov_v2", JSON.stringify(apps));
    } catch {}
  }, [apps, ready]);

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
    const scored = apps.map(a => ({ ...a, sc: calcScores(a.scores) }));
    const withSc = scored.filter(a => a.sc.autonomyScore);
    const avgA   = withSc.length ? withSc.map(a => a.sc.autonomyScore).reduce((x,y)=>x+y,0)/withSc.length : null;
    const COLORS = ["#1e40af","#7c3aed","#065f46","#92400e","#991b1b","#0f766e"];

    const radarKey = n => n.substring(0, 13);
    const radarData = DAAF.map(d => ({
      dim: d.dimName.substring(0, 12),
      ...Object.fromEntries(apps.slice(0, 5).map(a => [radarKey(a.name), a.scores[d.key] || 0]))
    }));

    const kwData = scored
      .filter(a => a.sc.risico && a.sc.mitigatie && a.sc.belang)
      .map(a => ({
        name: a.name,
        x: +((a.sc.risico * a.sc.belang).toFixed(2)),
        y: a.sc.mitigatie,
        score: a.sc.autonomyScore
      }));

    return (
      <div className="h-full overflow-y-auto">
        <div className="p-5 max-w-5xl mx-auto">
            {/* Stat row */}
          <div className="grid grid-cols-4 gap-3 mb-5">
            {[
              { label:"Applicaties",       val: apps.length,     color: "#1A56A0" },
              { label:"Gem. autonomiescore",val: avgA ? avgA.toFixed(1) : "–", color: scoreColor(avgA) },
              { label:"Status: Goed (≥7)", val: withSc.filter(a=>a.sc.autonomyScore>=7).length, color:"#26B5AE" },
              { label:"Aandacht nodig (<5)",val:withSc.filter(a=>a.sc.autonomyScore<5).length,  color:"#E87722" },
            ].map(({ label, val, color }) => (
              <div key={label} className="rounded text-center px-3 py-4"
                style={{ background:"#fff", border:"1px solid #D0E4F7" }}>
                <div style={{ fontSize: 28, fontWeight: 700, color }}>{val}</div>
                <div className="text-xs mt-0.5" style={{ color:"#6b7280" }}>{label}</div>
              </div>
            ))}
          </div>

          {apps.length === 0 ? (
            <div className="bg-white rounded-2xl border-2 border-dashed border-gray-200 p-16 text-center">
              <div style={{ fontSize: 48, marginBottom: 12 }}>📊</div>
              <h3 className="text-xl font-semibold text-gray-700 mb-2">Start met uw assessment</h3>
              <p className="text-gray-400 text-sm mb-5">Voeg een applicatie toe om te beginnen met het soevereiniteitsassessment.</p>
              <button onClick={() => setShowModal(true)}
                className="text-white text-sm px-5 py-2.5 font-medium"
                style={{ background:"#1A56A0", borderRadius:4 }}>
                + Applicatie toevoegen
              </button>
            </div>
          ) : (
            <div className="grid grid-cols-2 gap-4">
              {/* App cards */}
              <div className="space-y-3">
                {scored.map(a => {
                  const lbl = scoreLabel(a.sc.autonomyScore);
                  return (
                    <div key={a.id} onClick={() => { setSelId(a.id); setStep(0); setView("assess"); }}
                      className="flex items-center gap-3 cursor-pointer transition-all"
                      style={{
                        background:"#fff", borderRadius:4, padding:16,
                        border:"1px solid #D0E4F7",
                        borderLeft:`4px solid ${scoreColor(a.sc.autonomyScore)}`
                      }}
                      onMouseEnter={e=>e.currentTarget.style.boxShadow="0 2px 8px rgba(26,86,160,0.15)"}
                      onMouseLeave={e=>e.currentTarget.style.boxShadow="none"}>
                      <Gauge score={a.sc.autonomyScore} size={72} />
                      <div className="flex-1 min-w-0">
                        <h3 className="font-semibold text-gray-900 text-sm truncate">{a.name}</h3>
                        <p className="text-xs text-gray-400">{a.supplier}</p>
                        <div className="flex items-center gap-2 mt-1">
                          <span className="text-xs px-2 py-0.5 rounded-full font-medium"
                            style={{ background: lbl.bg, color: lbl.fg }}>{lbl.text}</span>
                          <span className="text-xs text-gray-400">{a.sc.completeness}% ingevuld</span>
                        </div>
                        <div className="mt-2">
                          <SovBar score5={a.sc.dictuAvg} />
                        </div>
                      </div>
                    </div>
                  );
                })}
              </div>

              {/* Charts */}
              <div className="space-y-4">
                {apps.length >= 2 && (
                  <div className="rounded p-4" style={{ background:"#fff", border:"1px solid #D0E4F7" }}>
                    <p className="text-sm font-semibold text-gray-700 mb-2">DAAF Radar — alle applicaties</p>
                    <ResponsiveContainer width="100%" height={200}>
                      <RadarChart data={radarData}>
                        <PolarGrid stroke="#f3f4f6" />
                        <PolarAngleAxis dataKey="dim" tick={{ fontSize: 9 }} />
                        <PolarRadiusAxis domain={[0, 5]} tick={false} axisLine={false} />
                        {apps.slice(0, 5).map((a, i) => (
                          <Radar key={a.id} name={radarKey(a.name)} dataKey={radarKey(a.name)}
                            stroke={COLORS[i % COLORS.length]} fill={COLORS[i % COLORS.length]} fillOpacity={0.1} strokeWidth={1.5} />
                        ))}
                        <Tooltip wrapperStyle={{ fontSize: 11 }} />
                        <Legend wrapperStyle={{ fontSize: 10 }} />
                      </RadarChart>
                    </ResponsiveContainer>
                  </div>
                )}

                {kwData.length >= 2 && (
                  <div className="rounded p-4" style={{ background:"#fff", border:"1px solid #D0E4F7" }}>
                    <p className="text-sm font-semibold text-gray-700 mb-1">Kwadrant: Risico × Belang vs. Mitigatie</p>
                    <p className="text-xs text-gray-400 mb-2">Rechtsboven = hoog risico/belang. Verticaal = mitigatiekracht.</p>
                    <ResponsiveContainer width="100%" height={180}>
                      <ScatterChart margin={{ top:5, right:15, bottom:25, left:5 }}>
                        <CartesianGrid strokeDasharray="3 3" stroke="#f3f4f6" />
                        <XAxis dataKey="x" name="Risico×Belang" domain={[0,25]} tick={{ fontSize:9 }}
                          label={{ value:"Risico × Belang →", position:"insideBottom", offset:-12, fontSize:10 }} />
                        <YAxis dataKey="y" name="Mitigatie" domain={[1,5]} tick={{ fontSize:9 }}
                          label={{ value:"Mitigatie ↑", angle:-90, position:"insideLeft", fontSize:10 }} />
                        <Tooltip content={({ payload }) => {
                          if (!payload?.[0]) return null;
                          const d = payload[0].payload;
                          return (
                            <div className="bg-white border border-gray-200 rounded shadow p-2 text-xs">
                              <strong>{d.name}</strong><br/>Score: {d.score?.toFixed(1)}
                            </div>
                          );
                        }} />
                        <Scatter data={kwData}>
                          {kwData.map((d, i) => <Cell key={i} fill={scoreColor(d.score)} />)}
                        </Scatter>
                      </ScatterChart>
                    </ResponsiveContainer>
                  </div>
                )}

                {apps.length === 1 && (
                  <div className="rounded p-4 text-center py-16" style={{ background:"#fff", border:"1px solid #D0E4F7" }}>
                    <p className="text-sm text-gray-400">Voeg meer applicaties toe om vergelijkingen te zien.</p>
                  </div>
                )}
              </div>
            </div>
          )}
        </div>
      </div>
    );
  }

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
        <div className="flex-1 overflow-y-auto p-5" style={{ background:"#EBF3FF" }}>
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
                <button key={t.i} onClick={() => setStep(t.i)}
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
                <strong>DAAF Quick Scan (Utrecht University)</strong> — Niveau 1 (Risico): lager = minder risico.
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
              <button onClick={() => setStep(1)}
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
            <p style={{ fontSize:11, color:"#7DD3D0" }}>Transitieteam Digitalisering · Kwartiermaker Digitale Samenhang</p>
          </div>
        </div>
        <button onClick={exportXlsx} disabled={apps.length === 0}
          className="flex items-center gap-2 text-white text-xs px-4 py-2 font-medium transition-all"
          style={{
            background: apps.length === 0 ? "rgba(255,255,255,0.1)" : "#26B5AE",
            opacity: apps.length === 0 ? 0.5 : 1,
            borderRadius: 4
          }}>
          📥 Exporteer Excel
        </button>
      </header>

      {/* Sub-header nav — lichte blauwe balk (NHS Stenden stijl) */}
      <div className="flex-shrink-0" style={{ background:"#1A56A0" }}>
        <nav className="px-5 flex gap-0">
          {[
            { k:"dashboard", label:"Dashboard" },
            { k:"apps",      label:"Applicaties" },
            ...(selApp ? [{ k:"assess", label:selApp.name.substring(0,20) }] : []),
            { k:"compare",   label:"Vergelijking" },
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
        {view === "dashboard" && <Dashboard />}
        {view === "apps"      && <AppsList />}
        {view === "assess"    && <Assess />}
        {view === "compare"   && <Compare />}
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
