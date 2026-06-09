import React, { useState, useEffect, useRef } from "react";
import * as d3 from "d3";
import * as XLSX from "xlsx";
import {
  RadarChart, Radar, PolarGrid, PolarAngleAxis, PolarRadiusAxis,
  ResponsiveContainer, BarChart, Bar, XAxis, YAxis, CartesianGrid,
  Tooltip, Legend, ScatterChart, Scatter, Cell
} from "recharts";

// ── Error Boundary — toont de echte fout in het scherm ────────
export class ErrorBoundary extends React.Component {
  constructor(props) { super(props); this.state = { error: null }; }
  static getDerivedStateFromError(err) { return { error: err }; }
  render() {
    if (this.state.error) {
      return (
        <div style={{ padding:32, fontFamily:"monospace", background:"#fff1f2", minHeight:"100vh" }}>
          <h2 style={{ color:"#b91c1c", marginBottom:16 }}>⚠️ Applicatiefout</h2>
          <pre style={{ background:"#fff", border:"1px solid #fecaca", padding:16, borderRadius:4,
            fontSize:12, overflow:"auto", whiteSpace:"pre-wrap", color:"#374151" }}>
            {String(this.state.error)}{"\n\n"}{this.state.error?.stack}
          </pre>
          <p style={{ marginTop:16, fontSize:12, color:"#6b7280" }}>
            Kopieer bovenstaande foutmelding en stuur naar de ontwikkelaar.
          </p>
        </div>
      );
    }
    return this.props.children;
  }
}

// ──────────────────────────────────────────────────────────────
// VERSIE — verhoog met 0.1 bij elke release
// ──────────────────────────────────────────────────────────────
const VERSION = "v2.2.1"; // UPLOAD-CHECK: kwadrant r=7 fontSize=12
const MAX_VISIBLE = 10; // maximaal zichtbare applicaties in grafieken
const appColor = (i) => `hsl(${Math.round((i * 137.508) % 360)}, 65%, 42%)`; // unieke kleur per app-index

// Module-level naam helper — wordt aangeroepen met useSecondaryName als parameter
function dn(app, useSecondary) {
  if (!app) return "";
  if (useSecondary && app.nameSecondary && app.nameSecondary.trim()) return app.nameSecondary.trim();
  return app.name || "";
}

const CHANGELOG = [
  {
    versie: "v2.2",
    datum: "Juni 2026",
    wijzigingen: [
      "Versienummer verhoogd naar v2.2 vanwege omvang van wijzigingen",
      "Header: app-teller toegevoegd met aantal apps in database en statuslampje",
      "Header: groen lampje = gesynchroniseerd, geel = opslaan, rood = fout",
      "Dashboard: autonomie-kwadrant volledig breed over de pagina",
      "Dashboard: app-kaarten in aparte 2-koloms rij onder het kwadrant",
      "Spindiagram: kleiner gemaakt (560x440) zodat het binnen de browser past",
      "Bug opgelost: visibleApps scope-fout op Dashboard bij meer dan 10 apps",
      "Bug opgelost: appColor scope-fout na toevoegen applicatie",
      "Bug opgelost: Toevoegen-knop weer zichtbaar voor alle gebruikers",
      "Bug opgelost: MAX_VISIBLE naar module-level zodat alle paginas hem kennen",
      "Grafieken schaalbaar naar 50+ applicaties met dynamische HSL-kleuren",
      "Spindiagram: overlappende stippen gespreide weergave via jitter",
      "Grafiek-selectie: max 10 apps tegelijk voor leesbaarheid, handmatig instelbaar",
    ]
  },
  {
    versie: "v2.1",
    datum: "Juni 2026",
    wijzigingen: [
      "Prototype-status vastgelegd: applicatie is officieel een prototype — zichtbaar in header, Over de tool en Over dit product",
      "Header: subtitel bijgewerkt naar 'Prototype · Ambassadeurslijn Digitale Soevereiniteit'",
      "Over de tool: migratieparagraaf toegevoegd — product draait nu op Netlify, gaat over naar eigen NVIDIA DGX Spark",
      "Over de tool: inspiratiebron-paragraaf toegevoegd — instrument als voorbeeld voor beleidsontwikkeling in eigen organisatie",
      "Over dit product: prototype-vermelding en migratieplan naar eigen NHL Stenden-server toegevoegd",
      "Aan de slag stap 2: verwijzing naar 'live sidebar' verwijderd (niet meer aanwezig in huidige versie)",
      "Aan de slag stap 4: verwijzing naar 'autonomiekwadrant' verwijderd (niet meer aanwezig in huidige versie)",
      "Excel-exportknop verplaatst van header naar Dashboard (was al zo maar tekst klopte niet)",
      "Assessment: alleen-lezen voor gewone gebruikers, bewerken alleen via Beheer met pincode",
      "Geopolitiek tabblad verwijderd — geopolitieke kaart verplaatst naar Portfolio-pagina als compact overzicht",
      "Consistentiecheck gehele applicatie: beschrijvingen, labels en uitleg actueel gemaakt voor v2.1",
    ]
  },
  {
    versie: "v2.0",
    datum: "Juni 2026",
    wijzigingen: [
      "Database export: volledig JSON-bestand met alle apps, scores en motivaties, bestandsnaam met datum en tijd",
      "Database import: selectief importeren van applicaties uit een exportbestand, overzicht toont nieuw vs overschrijven",
      "Import-modal: alles/niets selecteren plus per applicatie aanvinken, ook verwijderde apps terugzetten",
      "Over & uitleg Tips-tabblad uitgebreid met export/import handleiding en stap-voor-stap uitleg",
    ]
  },
  {
    versie: "v1.9",
    datum: "Juni 2026",
    wijzigingen: [
      "PDF: kernapplicatietabel (hoofdstuk 7.2) verwijderd uit het rapport",
      "PDF: bij enkelvoudige selectie verschijnt nu hoofdstuk 7 'Verbeteracties' met een concrete actielijst en afvinkbare tabel per DAAF-dimensie",
      "PDF: bij meervoudige selectie blijft hoofdstuk 7 de vervolgstappen voor eigenaren (zonder vaste kernapplicatietabel)",
      "PDF: voorblad en titel passen zich aan bij enkelvoudige export (naam applicatie prominent zichtbaar)",
      "Alle verwijzingen naar 'bestuur', 'bestuurlijk' en 'CvB' vervangen door NHL Stenden organisatienaamgeving",
    ]
  },
  {
    versie: "v1.8",
    datum: "Juni 2026",
    wijzigingen: [
      "Beveiliging niveau 1: server-side API-tokencheck in Netlify Functions (x-api-token header, HTTP 401 bij ongeldig verzoek)",
      "App.jsx: apiToken wordt meegestuurd bij alle API-aanroepen naar save-data en load-data",
      "Over dit product: uitgebreide beveiligingssectie toegevoegd met versleuteling, toegangsbeveiliging, bekende beperkingen en niveaus van advies",
    ]
  },
  {
    versie: "v1.7",
    datum: "Juni 2026",
    wijzigingen: [
      "Tekst 'Over & uitleg' gecorrigeerd: verouderde verwijzing naar localStorage verwijderd uit de Data opslaan tip",
      "Stap 6 (Excel exporteren) bijgewerkt: beschrijft nu vijf tabbladen (inclusief Motivaties)",
      "'Over & uitleg' pagina volledig omgebouwd met vier tabbladen: Over de tool, Aan de slag, Scores & grafieken, Tips & beheer",
      "Volgorde logischer: tool-uitleg en frameworks eerst, daarna stap-voor-stap, dan grafieken, dan beheer",
      "Nieuw tabblad 'Bestuur' toegevoegd: bestuurssamenvatting met portfoliostatus, risicografiek en top 3 aandachtspunten",
      "Nieuw tabblad 'Transparantie' toegevoegd: DAAF-beoordeling van de eigen technische stack (Netlify, Anthropic, GitHub, React)",
      "Transparantiepagina toont jurisdictie, datalocatie, beveiliging en gegevensstroom per component",
    ]
  },
  {
    versie: "v1.6",
    datum: "Juni 2026",
    wijzigingen: [
      "PDF volledig herschreven: voorblad, inhoudsopgave, inleiding met frameworkuitleg",
      "PDF: paragraafteksten boven elk onderdeel, betere paginabreaks",
      "PDF: datum, versie en naamweergave (primair/secundair) vermeld op voorblad en header",
      "PDF: herhaalende paginaheader op elke pagina met versie en naamstatus",
      "'Alles wissen' knop verwijderd uit de beheeromgeving",
    ]
  },
  {
    versie: "v1.5",
    datum: "Juni 2026",
    wijzigingen: [
      "Secundaire applicatienaam toegevoegd — elke app kan naast een primaire naam ook een alternatieve weergavenaam krijgen",
      "Zichtbaarheidsschakelaar in de header — wissel met één klik tussen primaire en secundaire namen (overal tegelijk: dashboard, grafieken, PDF, Excel)",
      "Beheeromgeving: 'Dummie namen toewijzen' knop voor het automatisch toewijzen van tijdelijke namen aan bestaande apps",
      "Beheeromgeving: beide namen zichtbaar in de app-kaart (primair + geel secundair label)",
      "Add-app formulier en bewerk-scherm uitgebreid met veld voor secundaire naam",
    ]
  },
  {
    versie: "v1.4",
    datum: "Juni 2026",
    wijzigingen: [
      "Motivatieveld toegevoegd bij elke vraag in het assessment — optionele toelichting op de score",
      "Motivaties zichtbaar in de beheeromgeving per applicatie (groen icoontje 💬 als er een motivatie is)",
      "Motivaties opgenomen in de PDF-export als extra kolom in de scoretabel",
      "Excel-export uitgebreid: motivatiekolommen in DAAF en DICTU tabbladen, plus nieuw tabblad 'Motivaties'",
    ]
  },
  {
    versie: "v1.3",
    datum: "Juni 2026",
    wijzigingen: [
      "Changelog toegevoegd in beheeromgeving — overzicht van alle versies en wijzigingen",
      "Versienummer wordt voortaan bijgewerkt bij elke nieuwe release",
    ]
  },
  {
    versie: "v1.2",
    datum: "Juni 2026",
    wijzigingen: [
      "DICTU spindiagram toegevoegd aan dashboard — alle 4 assen met hover-tooltip",
      "Tooltip DICTU spindiagram gefixed via onMouseMove proximiteit (werkt nu voor alle applicaties)",
      "Toevoegen van applicaties hersteld voor gewone gebruikers; verwijderen/bewerken alleen voor beheerder",
      "Versienummer bijgewerkt en changelog toegevoegd",
    ]
  },
  {
    versie: "v1.1",
    datum: "Mei 2026",
    wijzigingen: [
      "Spindiagram vervangen door DivergingChart voor DAAF — semantisch correcte assen (richting heeft betekenis)",
      "Opdrachtskaart toegevoegd bovenaan dashboard met centrale vraagstelling NHL Stenden",
      "Heatmap verwijderd (onleesbaar bij veel applicaties)",
      "Tekst afbreking Quick win en Strategische aanbeveling opgelost — volledige tekst zichtbaar",
      "Label 'Strategisch' hernoemd naar 'Strategische aanbeveling'",
      "Term 'bestuurlijk' en 'CvB' vervangen door NHL Stenden organisatienaamgeving",
      "PDF typografie verbeterd — lopende tekst met koppen en alinea's",
      "PDF dimensieprofiel: spindiagram vervangen door horizontale balkengrafiek",
    ]
  },
  {
    versie: "v1.0",
    datum: "Mei 2026",
    wijzigingen: [
      "Eerste volledige versie live op nhl-soevereiniteitsassessment.netlify.app",
      "DAAF Quick Scan (9 indicatoren) volledig geïmplementeerd conform Utrecht University framework",
      "DICTU Soevereiniteitscheck (4 vragen) toegevoegd",
      "Autonomie-kwadrant (4 kwadranten: OPTIMAAL / BEHEERSBAAR / AANDACHTSPUNT / KRITIEK)",
      "Applicatiekaarten met snelheidsmeter, DICTU-balk en aanbevelingen",
      "Vergelijkingspagina met filter, staafdiagram en vergelijkingstabel",
      "PDF-export met samenvatting, kwadrant, dimensieprofiel en aanbevelingen per app",
      "Excel-export met 4 tabbladen",
      "Beheeromgeving met pincode, bewerken en verwijderen",
      "Netlify Blobs opslag voor gedeelde data tussen gebruikers",
    ]
  },
];

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
  { key:"C1", dim:"C", level:"Mitigatie",
    dimName:"Technische weerbaarheid",
    dimSub:"Kun je technisch gezien overstappen? Zijn er alternatieven, is data exporteerbaar, is de software open?",
    name:"Alternatief beschikbaar",
    hint:"1 = slechte mitigatie ... 5 = sterke mitigatie (goed)",
    question:"Bestaan er vergelijkbare diensten als vervanging?",
    toelichting:"Welke alternatieven? Gebruikt door andere universiteiten? Europees/open source?",
    scores:[
      {s:1,label:"Geen alternatief", desc:"Geen alternatief. Uniek product zonder concurrenten."},
      {s:2,label:"Onvolwassen",      desc:"Alternatief bestaat maar onvolwassen of functioneel beperkt."},
      {s:3,label:"Vergelijkbaar",    desc:"Meerdere alternatieven bij vergelijkbare organisaties. Dekt het meeste."},
      {s:4,label:"Volwassen",        desc:"Volwassen alternatieven incl. Europese optie."},
      {s:5,label:"Breed aanbod",     desc:"Breed aanbod incl. Europese en open source opties. Competitieve markt."}
    ]
  },
  { key:"D1", dim:"D", level:"Mitigatie",
    dimName:"Organisatorische weerbaarheid",
    dimSub:"Heb je de kennis, plannen en backups om een overstap daadwerkelijk uit te voeren?",
    name:"Interne expertise en kennisborging",
    hint:"1 = slechte mitigatie ... 5 = sterke mitigatie (goed)",
    question:"Hoeveel interne kennis hebben we over dit systeem?",
    toelichting:"Denk aan alle relevante kennis: functioneel (wat doet het systeem), technisch (configuratie, koppelingen), data (welke data zit erin, hoe exporteer je die). Niet alleen technisch beheer. Wat gebeurt er als de belangrijkste expert vertrekt?",
    scores:[
      {s:1,label:"Geen kennis",       desc:"Geen interne kennis. Volledig afhankelijk van leverancier."},
      {s:2,label:"Beperkt",           desc:"Kennis bij een of twee personen, niet vastgelegd."},
      {s:3,label:"Basisdocumentatie", desc:"Kennis bij meerdere personen, basisdocumentatie aanwezig."},
      {s:4,label:"Breed gedeeld",     desc:"Kennis breed gedeeld en gedocumenteerd. Niet persoonsafhankelijk."},
      {s:5,label:"Structureel geborgd",desc:"Kennis structureel geborgd. Volledige documentatie, periodiek bijgewerkt."}
    ]
  },
  { key:"E1", dim:"E", level:"Mitigatie",
    dimName:"Contractuele weerbaarheid",
    dimSub:"Wat is er contractueel geregeld over vertrek, data-overdracht en flexibiliteit?",
    name:"Exit-clausules en transitieregeling",
    hint:"1 = slechte mitigatie ... 5 = sterke mitigatie (goed)",
    question:"Wat is contractueel geregeld over data-overdracht bij beëindiging?",
    toelichting:"Contract over beëindiging? Formaat data? Transitieperiode? Consequenties?",
    scores:[
      {s:1,label:"Geen afspraken",    desc:"Geen afspraken over beëindiging."},
      {s:2,label:"Minimaal",          desc:"Minimale bepalingen, zonder termijn of formaat."},
      {s:3,label:"Transitieperiode",  desc:"Exit-clausules met transitieperiode aanwezig."},
      {s:4,label:"Uitgebreid",        desc:"Uitgebreide exit-regeling: open formaten, transitieperiode, actieve medewerking leverancier."},
      {s:5,label:"Afdwingbaar",       desc:"Uitgebreide exit-regeling met afdwingbare consequenties bij niet-nakoming (boeteclausule, escalatieprocedure, verwijdergarantie)."}
    ]
  },
  { key:"F1", dim:"F", level:"Belang",
    dimName:"Organisatorisch belang",
    dimSub:"Hoe belangrijk is deze applicatie voor de organisatie en hoeveel hangt ervan af?",
    name:"Impact bij uitval",
    hint:"1 = laag belang ... 5 = hoog belang (urgent)",
    question:"Wat is de impact als dit systeem uitvalt?",
    toelichting:"Stel dat het systeem 48 uur uitvalt. Wat is de impact? Dit geldt voor alle processen: onderwijs, onderzoek en bedrijfsvoering. Weeg mee: het aantal getroffen gebruikers, de beschikbaarheid van workarounds, en hoe kritiek de getroffen processen zijn.",
    scores:[
      {s:1,label:"Geen impact",        desc:"Geen merkbare impact bij uitval."},
      {s:2,label:"Minimale hinder",    desc:"Minimale hinder. Organisatie draait door."},
      {s:3,label:"Verstoord",          desc:"Meerdere processen verstoord maar niet stilgevallen."},
      {s:4,label:"Stilgevallen",       desc:"Belangrijke processen stilgevallen."},
      {s:5,label:"Organisatiebreed",   desc:"Organisatiebrede stilstand."}
    ]
  },
  { key:"G1", dim:"G", level:"Belang",
    dimName:"Data-gevoeligheid",
    dimSub:"Welke gevoelige gegevens zitten in dit systeem? Persoonsgegevens, onderzoeksdata, intellectueel eigendom.",
    name:"Persoonsgegevens",
    hint:"1 = laag belang ... 5 = hoog belang (urgent)",
    question:"Welke persoonsgegevens worden verwerkt?",
    toelichting:"Welke gegevens? Gevoelige categorieën? Hoeveel personen? DPIA?",
    scores:[
      {s:1,label:"Geen",              desc:"Geen persoonsgegevens."},
      {s:2,label:"Basisgegevens",     desc:"Basisgegevens: naam, e-mail, functie. Beperkt aantal."},
      {s:3,label:"Studie-/werkgeg.",  desc:"Studie-/werkgegevens, beoordelingen, financieel. Significant."},
      {s:4,label:"Bijzondere geg.",   desc:"Bijzondere persoonsgegevens of grootschalige verwerking."},
      {s:5,label:"Grootschalig bijz.",desc:"Grootschalig bijzondere gegevens van kwetsbare groepen."}
    ]
  },
  { key:"H1", dim:"H", level:"Belang",
    dimName:"Academische impact",
    dimSub:"Raakt deze applicatie aan academische vrijheid, samenwerking of langetermijnarchivering?",
    name:"Academische vrijheid",
    hint:"1 = laag belang ... 5 = hoog belang (urgent)",
    question:"Kan de leverancier invloed uitoefenen op academische activiteiten?",
    toelichting:"Kan de leverancier of een buitenlandse overheid inzicht krijgen in onderzoeksactiviteiten? Zijn er eerdere gevallen van censuur of datavordering bij deze leverancier?",
    scores:[
      {s:1,label:"Geen relatie",      desc:"Geen relatie met onderzoek."},
      {s:2,label:"Theoretisch risico",desc:"Theoretisch risico, geen eerdere gevallen bekend."},
      {s:3,label:"Beperkt risico",    desc:"Toegang tot onderzoekspatronen. Beperkt risico op indirecte beïnvloeding."},
      {s:4,label:"Reëel risico",      desc:"Reëel risico. Land waar overheden onderzoek kunnen censureren."},
      {s:5,label:"Aangetoond",        desc:"Aangetoonde censuur of datavordering bij deze leverancier of overheid."}
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
  const sc = k => scores[k] || 0;  // score voor één indicator, 0 = niet ingevuld

  // ── Gewogen gemiddelde binnen een dimensie ──────────────────
  // Alleen meerekenen als de indicator ook daadwerkelijk is ingevuld (>0)
  const weightedDim = (indicators) => {
    const filled = indicators.filter(([k]) => sc(k) > 0);
    if (filled.length === 0) return null;
    const totalWeight = filled.reduce((s, [, w]) => s + w, 0);
    const weightedSum = filled.reduce((s, [k, w]) => s + sc(k) * w, 0);
    return weightedSum / totalWeight;
  };

  // ── Niveau = gemiddelde van dimensiescores ──────────────────
  const levelAvg = dimScores => {
    const filled = dimScores.filter(d => d !== null);
    return filled.length ? filled.reduce((a, b) => a + b, 0) / filled.length : null;
  };

  // ── Niveau 1: Risico-exposure ───────────────────────────────
  // Dimensie A: Geopolitiek risico  (A1 gewicht=3, A3 gewicht=2)
  const dimA = weightedDim([["A1", 3], ["A3", 2]]);
  // Dimensie B: Leveranciersafhankelijkheid  (B1 gewicht=3 — enige indicator)
  const dimB = weightedDim([["B1", 3]]);
  const risico = levelAvg([dimA, dimB]);

  // ── Niveau 2: Mitigatie-capaciteit ─────────────────────────
  // Dimensie C: Technische weerbaarheid  (C1 gewicht=2)
  const dimC = weightedDim([["C1", 2]]);
  // Dimensie D: Organisatorische weerbaarheid  (D1 gewicht=2)
  const dimD = weightedDim([["D1", 2]]);
  // Dimensie E: Contractuele weerbaarheid  (E1 gewicht=3)
  const dimE = weightedDim([["E1", 3]]);
  const mitigatie = levelAvg([dimC, dimD, dimE]);

  // ── Niveau 3: Strategisch belang ────────────────────────────
  // Dimensie F: Organisatorisch belang  (F1 gewicht=3)
  const dimF = weightedDim([["F1", 3]]);
  // Dimensie G: Data-gevoeligheid  (G1 gewicht=2)
  const dimG = weightedDim([["G1", 2]]);
  // Dimensie H: Academische impact  (H1 gewicht=2)
  const dimH = weightedDim([["H1", 2]]);
  const belang = levelAvg([dimF, dimG, dimH]);

  // ── DICTU soevereiniteitsgemiddelde ─────────────────────────
  const dictuKeys = ["2.1","2.2","2.3","4.1"];
  const filledDictu = dictuKeys.filter(k => sc(k) > 0);
  const dictuAvg = filledDictu.length
    ? filledDictu.reduce((s, k) => s + sc(k), 0) / filledDictu.length
    : null;

  // ── Autonomiescore (1–10, logaritmische schaal) ─────────────
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
  const filled  = allKeys.filter(k => sc(k) > 0).length;

  return {
    risico, mitigatie, belang, autonomyScore, dictuAvg,
    // Dimensiescores beschikbaar voor spindiagram
    dims: { A:dimA, B:dimB, C:dimC, D:dimD, E:dimE, F:dimF, G:dimG, H:dimH },
    completeness: Math.round(100 * filled / allKeys.length)
  };
}

// ──────────────────────────────────────────────────────────────
// AANBEVELINGEN — automatisch gegenereerd op basis van scores
// ──────────────────────────────────────────────────────────────
function generateRecommendations(scores) {
  const sc  = calcScores(scores);
  const get = k => scores[k] || 0;

  // ── Quick win: laagst scorende verbeterbare dimensie ─────────
  let quickWin = "";
  const mit = [
    { k:"C1", s:get("C1"), label:"alternatieven", tip:"Verken en documenteer minimaal twee concrete alternatieven voor deze applicatie, inclusief een Europese optie. Breng switching costs en migratietijd in kaart." },
    { k:"D1", s:get("D1"), label:"interne kennis", tip:"Start met het documenteren van interne kennis: functioneel, technisch en datakennis. Stel een exitplan op en verbreed de kennisbasis zodat niet alles bij één persoon ligt." },
    { k:"E1", s:get("E1"), label:"exit-clausules", tip:"Zorg bij de eerstvolgende contractverlenging voor expliciete exit-clausules: open dataformaten, transitieperiode, actieve leveranciersmedewerking en liefst afdwingbare consequenties bij niet-nakoming." },
  ].filter(m => m.s > 0).sort((a,b) => a.s - b.s);

  const a1=get("A1"), a3=get("A3"), dictu21=get("2.1"), dictu22=get("2.2");

  if (mit.length > 0 && mit[0].s <= 2) {
    quickWin = mit[0].tip;
  } else if (a1 >= 4) {
    quickWin = "Verifieer de juridische jurisdictie van de leverancier. Controleer of er een adequaatheidsbesluit van toepassing is en wat de implicaties zijn van de CLOUD Act of FISA 702. Overweeg een juridisch advies bij twijfel.";
  } else if (dictu21 <= 2) {
    quickWin = "Verifieer contractueel of alle data inclusief back-ups en logs uitsluitend binnen de EU worden opgeslagen. Vraag de leverancier om een schriftelijke bevestiging met specifieke datacenterlocaties.";
  } else if (dictu22 <= 2) {
    quickWin = "Bespreek met de leverancier de mogelijkheid van klant-beheerde sleutels (BYOK of CMK). Dit is een relatief eenvoudige stap die de technische toegangsbeveiliging significant verbetert.";
  } else if (mit.length > 0) {
    quickWin = mit[0].tip;
  } else {
    quickWin = "Voer een periodieke hercontrole uit bij de eerstvolgende contractverlenging. Zorg dat de assessmentresultaten worden gedeeld met de contractverantwoordelijke.";
  }

  // ── Strategische aanbeveling: kwadrantpositie ────────────────
  let strategic = "";
  const risB = sc.risico * sc.belang || 0;
  const mit2 = sc.mitigatie || 0;
  const inKritiek      = risB > 13 && mit2 < 3;
  const inBeheersbaar  = risB > 13 && mit2 >= 3;
  const inAandacht     = risB <= 13 && mit2 < 3;
  const inOptimaal     = risB <= 13 && mit2 >= 3;
  const dictuLaag      = sc.dictuAvg && sc.dictuAvg < 3;

  if (inKritiek) {
    strategic = "Deze applicatie staat in het kwadrant KRITIEK: hoog risico én lage weerbaarheid. Urgente actie is vereist. Overweeg drie opties: (1) Migreer naar een Europese aanbieder met lagere risicoscore, (2) versterk de mitigatie door contractonderhandelingen, kennisborging en alternatieven te ontwikkelen, of (3) accepteer het risico bewust via een bewust genomen besluit met onderbouwing. Stel een actieplan op met een concrete deadline.";
  } else if (inBeheersbaar) {
    strategic = "Deze applicatie staat in het kwadrant BEHEERSBAAR: hoog risico maar goede weerbaarheid. De risico's zijn geaccepteerd met een solide fallback-positie. Strategisch advies: bewaken dat de weerbaarheid op peil blijft, met name als er leverancierswijzigingen plaatsvinden. Neem clausules op die NHL Stenden informeren bij overname of beleidswijzigingen van de leverancier.";
  } else if (inAandacht) {
    strategic = "Deze applicatie staat in het kwadrant AANDACHTSPUNT: beperkt risico maar ook beperkte weerbaarheid. Er is ruimte voor verbetering zonder hoge urgentie. Strategisch advies: gebruik de relatief lage druk als momentum om weerbaarheid structureel op te bouwen — begin met kennisborging en contractuele exit-clausules.";
  } else if (inOptimaal) {
    strategic = "Deze applicatie staat in het kwadrant OPTIMAAL: beperkt risico en goede weerbaarheid. Behoud de huidige positie. Strategisch advies: zorg voor periodieke hercontrole (jaarlijks of bij contractverlenging) en houd de weerbaarheidsmaatregelen actueel. Deel de aanpak als voorbeeld voor andere applicaties.";
  } else {
    strategic = "Vul alle DAAF-dimensies in voor een volledige strategische duiding.";
  }

  if (dictuLaag && !inKritiek) {
    strategic += ` Aanvullend: de DICTU-soevereiniteitsscore is laag (${sc.dictuAvg?.toFixed(1)}/5). Prioriteer verbetering van dataresidency en sleutelbeheer bij de volgende leveranciersevaluatie.`;
  }

  return { quickWin, strategic };
}



// About sub-components — module level om React re-mount te voorkomen
function Section({ title, children, accent="#1A56A0" }) {
  return (
    <div className="mb-6">
      <div className="flex items-center gap-2 mb-3">
        <div className="w-1 h-5 rounded-full flex-shrink-0" style={{ background:accent }}/>
        <h2 className="font-bold text-base" style={{ color:"#0C2340" }}>{title}</h2>
      </div>
      {children}
    </div>
  );
}

function Tip({ label, children, color="#1A56A0", bg="#EBF3FF" }) {
  return (
    <div className="rounded p-3 mb-2" style={{ background:bg, border:`1px solid ${color}33` }}>
      <p className="text-xs font-bold mb-1" style={{ color }}>{label}</p>
      <p className="text-xs leading-relaxed" style={{ color:"#374151" }}>{children}</p>
    </div>
  );
}

// Kwadrant SVG — module level, ontvangt data als props
function KwadrantSVG({ kwData, onAppClick }) {
  const W = 700, H = 420;
  const pad = { top:32, right:24, bottom:52, left:52 };
  const iW  = W - pad.left - pad.right;
  const iH  = H - pad.top  - pad.bottom;
  const xMin=1, xMax=25, yMin=1, yMax=5;
  const mx = 13, my = 3;
  const toX = v => pad.left + (v - xMin) / (xMax - xMin) * iW;
  const toY = v => pad.top  + (yMax - v) / (yMax - yMin) * iH;
  const midX = toX(mx), midY = toY(my);
  const quads = [
    { x1:pad.left, y1:pad.top,   x2:midX,        y2:midY,        fill:"#e8f5e9", label:"OPTIMAAL",      sub:"Behoud huidige situatie,\nmonitor periodiek",     color:"#2e7d5e" },
    { x1:midX,     y1:pad.top,   x2:pad.left+iW, y2:midY,        fill:"#fff8e1", label:"BEHEERSBAAR",   sub:"Risico's geaccepteerd\nmet goede fallback",        color:"#e07b20" },
    { x1:pad.left, y1:midY,      x2:midX,        y2:pad.top+iH,  fill:"#fff3e0", label:"AANDACHTSPUNT", sub:"Bouw mitigatie op of\naccepteer risico bewust",    color:"#e07b20" },
    { x1:midX,     y1:midY,      x2:pad.left+iW, y2:pad.top+iH,  fill:"#fce4ec", label:"KRITIEK",       sub:"Urgente actie vereist:\nmigreer of mitigeer",      color:"#c0392b" },
  ];
  const xTicks = [1,5,10,15,20,25];
  const yTicks = [1,2,3,4,5];

  return (
    <svg viewBox={`0 0 ${W} ${H}`} width="100%" style={{ fontFamily:"system-ui,sans-serif", display:"block" }}>
      {quads.map((q,i) => (
        <rect key={i} x={q.x1} y={q.y1} width={q.x2-q.x1} height={q.y2-q.y1} fill={q.fill} />
      ))}
      {quads.map((q,i) => {
        const cx = (q.x1+q.x2)/2, cy = (q.y1+q.y2)/2;
        return (
          <g key={i}>
            <text x={cx} y={cy-10} textAnchor="middle" fill={q.color}
              style={{ fontSize:12, fontWeight:700, fontStyle:"italic", letterSpacing:0.5 }}>{q.label}</text>
            {q.sub.split("\n").map((l,j) => (
              <text key={j} x={cx} y={cy+6+j*13} textAnchor="middle" fill="#777" style={{ fontSize:9 }}>{l}</text>
            ))}
          </g>
        );
      })}
      {xTicks.map(v => <line key={v} x1={toX(v)} y1={pad.top} x2={toX(v)} y2={pad.top+iH} stroke="#fff" strokeWidth={v===mx?0:1} strokeDasharray="3 3" />)}
      {yTicks.map(v => <line key={v} x1={pad.left} y1={toY(v)} x2={pad.left+iW} y2={toY(v)} stroke="#fff" strokeWidth={v===my?0:1} strokeDasharray="3 3" />)}
      <line x1={midX} y1={pad.top} x2={midX} y2={pad.top+iH} stroke="#aaa" strokeWidth={1.5} />
      <line x1={pad.left} y1={midY} x2={pad.left+iW} y2={midY} stroke="#aaa" strokeWidth={1.5} />
      <rect x={pad.left} y={pad.top} width={iW} height={iH} fill="none" stroke="#ccc" strokeWidth={1} />
      {xTicks.map(v => (
        <g key={v}>
          <line x1={toX(v)} y1={pad.top+iH} x2={toX(v)} y2={pad.top+iH+5} stroke="#999" strokeWidth={1}/>
          <text x={toX(v)} y={pad.top+iH+14} textAnchor="middle" fill="#aaa" style={{ fontSize:8 }}>{v}</text>
        </g>
      ))}
      {yTicks.map(v => (
        <g key={v}>
          <line x1={pad.left-5} y1={toY(v)} x2={pad.left} y2={toY(v)} stroke="#999" strokeWidth={1}/>
          <text x={pad.left-8} y={toY(v)+3} textAnchor="end" fill="#aaa" style={{ fontSize:8 }}>{v}</text>
        </g>
      ))}
      <text x={pad.left+iW/2} y={H-4} textAnchor="middle" fill="#666" style={{ fontSize:10, fontWeight:600 }}>Risico-exposure × Strategisch belang</text>
      <text x={14} y={pad.top+iH/2} textAnchor="middle" fill="#666" transform={`rotate(-90,14,${pad.top+iH/2})`} style={{ fontSize:10, fontWeight:600 }}>Mitigatie</text>
      <text x={pad.left+6}    y={pad.top-10} fill="#aaa" style={{ fontSize:9 }}>Laag</text>
      <text x={pad.left+iW-24} y={pad.top-10} fill="#aaa" style={{ fontSize:9 }}>Hoog</text>
      <text x={pad.left-46}  y={pad.top+iH-4} fill="#aaa" style={{ fontSize:9 }}>Laag</text>
      <text x={pad.left-46}  y={pad.top+10}   fill="#aaa" style={{ fontSize:9 }}>Hoog</text>
      {kwData.length === 0 && (
        <text x={W/2} y={H/2+10} textAnchor="middle" fill="#bbb" style={{ fontSize:13 }}>
          Vul minimaal alle dimensies in om applicaties te plotten
        </text>
      )}
      {kwData.map(d => {
        const cx = toX(d.x), cy = toY(d.y), col = scoreColor(d.score);
        return (
          <g key={d.id} style={{ cursor:"pointer" }} onClick={() => onAppClick(d.id)}>
            <circle cx={cx} cy={cy} r={7} fill={col} fillOpacity={0.85} stroke="white" strokeWidth={1.5} />
            <rect x={cx+6} y={cy-7} width={Math.min(d.name.length*5+4,100)} height={12} rx={2} fill="white" fillOpacity={0.88} />
            <text x={cx+8} y={cy+1} fill={col} style={{ fontSize:8, fontWeight:600 }}>{d.name.substring(0,18)}</text>
          </g>
        );
      })}
    </svg>
  );
}

// ── DivergingChart — vervangt het spindiagram ────────────────
// Semantisch correcte weergave: richting van de as heeft betekenis
// Links/rood = risico (laag is beter), rechts/groen = weerbaarheid (hoog is beter)
function DivergingChart({ apps, compact = false, useSecondaryName = false }) {
  const [tip, setTip] = React.useState(null);

  const dimScore = (app, letter) => {
    if (letter === "A") {
      const a1=app.scores["A1"]||0, a3=app.scores["A3"]||0;
      const p=[[a1,3],[a3,2]].filter(([v])=>v>0);
      if (!p.length) return 0;
      const tw=p.reduce((s,[,w])=>s+w,0);
      return p.reduce((s,[v,w])=>s+v*w,0)/tw;
    }
    const qs=DAAF.filter(d=>d.dim===letter);
    const vals=qs.map(q=>app.scores[q.key]||0).filter(v=>v>0);
    return vals.length ? vals.reduce((s,v)=>s+v,0)/vals.length : 0;
  };

  const groups = [
    {
      id:"risico", title:"Niveau 1 — Risico-exposure",
      note:"Lagere score is beter — minder blootstelling",
      tc:"#991b1b", bg:"#fff1f2", border:"#fecaca",
      gradient:"linear-gradient(to right, #dcfce7 0%, #fef9c3 40%, #fca5a5 75%, #dc2626 100%)",
      dims:[
        {l:"A", n:"Geopolitiek risico"},
        {l:"B", n:"Leveranciersafhankelijkheid"},
      ]
    },
    {
      id:"mitigatie", title:"Niveau 2 — Mitigatie-capaciteit",
      note:"Hogere score is beter — meer weerbaarheid",
      tc:"#166534", bg:"#f0fdf4", border:"#bbf7d0",
      gradient:"linear-gradient(to right, #dc2626 0%, #fca5a5 25%, #fde68a 55%, #86efac 80%, #16a34a 100%)",
      dims:[
        {l:"C", n:"Technische weerbaarheid"},
        {l:"D", n:"Organisatorische weerbaarheid"},
        {l:"E", n:"Contractuele weerbaarheid"},
      ]
    },
    {
      id:"belang", title:"Niveau 3 — Strategisch belang",
      note:"Hogere score = meer urgentie — vraagt extra aandacht",
      tc:"#92400e", bg:"#fff7ed", border:"#fed7aa",
      gradient:"linear-gradient(to right, #fffbeb 0%, #fde68a 40%, #f59e0b 70%, #b45309 100%)",
      dims:[
        {l:"F", n:"Organisatorisch belang"},
        {l:"G", n:"Data-gevoeligheid"},
        {l:"H", n:"Academische impact"},
      ]
    },
  ];

  const manyApps = apps.length > 20;
  const barH    = compact ? 20 : 24;
  const labelW  = compact ? 200 : 230;
  const dotSize = compact ? 10 : manyApps ? 10 : 16;

  return (
    <div style={{ fontFamily:"system-ui,sans-serif", position:"relative" }}>
      {/* Legenda — scrollbaar bij veel apps */}
      <div style={{ display:"flex", flexWrap:"wrap", gap:"4px 12px", marginBottom:10,
                    maxHeight: apps.length > 15 ? 72 : "none", overflowY: apps.length > 15 ? "auto" : "visible",
                    padding: apps.length > 15 ? "4px 0" : 0 }}>
        {apps.map((app,i) => (
          <div key={i} style={{ display:"flex", alignItems:"center", gap:4, minWidth:0 }}>
            <div style={{ width:8, height:8, borderRadius:4, background:appColor(i), flexShrink:0 }}/>
            <span style={{ fontSize:9, color:"#374151", whiteSpace:"nowrap", overflow:"hidden",
                           textOverflow:"ellipsis", maxWidth:120 }}>
              {dn(app, useSecondaryName).substring(0,20)}
            </span>
          </div>
        ))}
      </div>

      {groups.map(g => (
        <div key={g.id} style={{ marginBottom:compact?10:14 }}>
          {/* Groepkop */}
          <div style={{
            display:"flex", justifyContent:"space-between", alignItems:"center",
            padding:"3px 10px", marginBottom:6,
            background:g.bg, border:`1px solid ${g.border}`,
            borderLeft:`3px solid ${g.tc}`, borderRadius:3
          }}>
            <span style={{ fontSize:compact?10:10.5, fontWeight:700, color:g.tc }}>{g.title}</span>
            <span style={{ fontSize:9, color:g.tc, opacity:0.75 }}>{g.note}</span>
          </div>

          {g.dims.map(d => {
            const points = apps.map((app,ai) => ({
              name:app.name,
              score:dimScore(app,d.l),
              color:appColor(ai)
            })).filter(p => p.score > 0);

            return (
              <div key={d.l} style={{ display:"flex", alignItems:"center", gap:8, marginBottom:compact?3:5 }}>
                {/* Label */}
                <div style={{ width:labelW, flexShrink:0, display:"flex", gap:4, alignItems:"center" }}>
                  <span style={{ fontSize:9,fontWeight:700,color:g.tc,width:16,flexShrink:0 }}>{d.l}</span>
                  <span style={{ fontSize:9, color:"#374151" }}>{d.n}</span>
                </div>

                {/* Balk + punten */}
                <div style={{ flex:1, position:"relative", height:barH+18 }}>
                  {/* Gradient achtergrond */}
                  <div style={{
                    position:"absolute", top:0, left:0, right:0, height:barH,
                    borderRadius:3, background:g.gradient,
                    border:"1px solid rgba(0,0,0,0.07)"
                  }}/>
                  {/* Tick-lijnen + labels */}
                  {[1,2,3,4,5].map(v => (
                    <React.Fragment key={v}>
                      {v > 1 && <div style={{
                        position:"absolute", left:`${(v-1)/4*100}%`,
                        top:0, height:barH, borderLeft:"1px solid rgba(255,255,255,0.55)"
                      }}/>}
                      <span style={{
                        position:"absolute", top:barH+3,
                        left:`${(v-1)/4*100}%`, transform:"translateX(-50%)",
                        fontSize:8, color:"#9ca3af"
                      }}>{v}</span>
                    </React.Fragment>
                  ))}
                  {/* Applicatie-punten */}
                  {points.map((p,pi) => (
                    <div key={pi}
                      onMouseEnter={e => setTip({
                        x: e.clientX, y: e.clientY,
                        name:p.name, dim:`${d.l}: ${d.n}`, score:p.score, color:p.color
                      })}
                      onMouseLeave={() => setTip(null)}
                      style={{
                        position:"absolute",
                        left:`calc(${(p.score-1)/4*100}% - ${dotSize/2}px)`,
                        top:(barH-dotSize)/2,
                        width:dotSize, height:dotSize, borderRadius:dotSize/2,
                        background:p.color, border:"2.5px solid white",
                        cursor:"crosshair", zIndex:10,
                        boxShadow:"0 1px 4px rgba(0,0,0,0.25)"
                      }}
                    />
                  ))}
                </div>

                {/* Gemiddelde score (voor enkelvoudige app) */}
                {apps.length === 1 && points.length > 0 && (
                  <span style={{ fontSize:11,fontWeight:700,color:g.tc,width:32,textAlign:"right",flexShrink:0 }}>
                    {points[0].score.toFixed(1)}
                  </span>
                )}
              </div>
            );
          })}
        </div>
      ))}

      {/* Tooltip */}
      {tip && (
        <div style={{
          position:"fixed", left:tip.x+14, top:tip.y-10,
          background:"white", border:`2px solid ${tip.color}`,
          borderRadius:6, padding:"6px 10px",
          fontSize:11, zIndex:9999, pointerEvents:"none",
          boxShadow:"0 4px 12px rgba(0,0,0,0.18)", minWidth:150
        }}>
          <div style={{ fontWeight:700, color:tip.color, marginBottom:2 }}>{tip.name}</div>
          <div style={{ fontSize:9, color:"#6b7280", marginBottom:4 }}>{tip.dim}</div>
          <div style={{ fontSize:14, fontWeight:700, color:tip.color }}>{tip.score.toFixed(2)} / 5</div>
        </div>
      )}
    </div>
  );
}

// ── OpdrachtKaart — centrale vraagstelling visueel ───────────
function OpdrachtKaart({ apps, useSecondaryName = false }) {
  const scored = apps.map(a => ({ ...a, sc: calcScores(a.scores) })).filter(a => a.sc.autonomyScore);
  if (apps.length === 0) return null;

  const kritiek  = scored.filter(a => a.sc.autonomyScore < 3);
  const zorg     = scored.filter(a => a.sc.autonomyScore >= 3 && a.sc.autonomyScore < 5);
  const ok       = scored.filter(a => a.sc.autonomyScore >= 5);
  const laagMit  = scored.filter(a => a.sc.mitigatie && a.sc.mitigatie < 2.5);
  const hoogRisk = scored.filter(a => a.sc.risico    && a.sc.risico    > 3.5);

  return (
    <div className="rounded mb-4" style={{ border:"2px solid #1A56A0", overflow:"hidden" }}>
      {/* Koptekst */}
      <div style={{ background:"#0C2340", padding:"14px 20px" }}>
        <div className="flex items-center gap-3">
          <div style={{ background:"#26B5AE", borderRadius:3, padding:"4px 10px",
            fontSize:10, fontWeight:700, color:"white", flexShrink:0 }}>
            VRAAGSTELLING
          </div>
          <p style={{ fontSize:13, fontWeight:700, color:"white", lineHeight:1.3 }}>
            Waar zetten we onze data neer en waar liggen de potentiële problemen?
          </p>
        </div>
        <p style={{ fontSize:10, color:"#7DD3D0", marginTop:4 }}>
          Aanleiding: inventarisatie digitale soevereiniteit — aansluiting bij VH en SURF digitale strategie
        </p>
      </div>

      {/* Inhoud */}
      <div className="grid grid-cols-3 gap-0" style={{ background:"#EBF3FF" }}>

        {/* Kolom 1: Applicatieoverzicht */}
        <div style={{ padding:"14px 16px", borderRight:"1px solid #D0E4F7" }}>
          <p style={{ fontSize:10, fontWeight:700, color:"#0C2340", marginBottom:8 }}>
            📋 Geassesseerd ({apps.length})
          </p>
          {[
            { label:"Kritiek (score &lt;3)",    items:kritiek,  color:"#b91c1c", bg:"#fee2e2" },
            { label:"Aandacht (score 3–5)", items:zorg,     color:"#c2410c", bg:"#ffedd5" },
            { label:"Acceptabel (score ≥5)", items:ok,      color:"#15803d", bg:"#dcfce7" },
          ].map(row => (
            <div key={row.label} className="flex items-start gap-2 mb-2">
              <div style={{ minWidth:10, height:10, borderRadius:2, background:row.color, marginTop:2, flexShrink:0 }}/>
              <div>
                <span style={{ fontSize:9, color:row.color, fontWeight:700 }}
                  dangerouslySetInnerHTML={{ __html: row.label }} />
                {row.items.length > 0 && (
                  <p style={{ fontSize:9, color:"#374151", lineHeight:1.4 }}>
                    {row.items.map(a=>dn(a, useSecondaryName)).join(", ")}
                  </p>
                )}
                {row.items.length === 0 && (
                  <p style={{ fontSize:9, color:"#9ca3af" }}>geen</p>
                )}
              </div>
            </div>
          ))}
        </div>

        {/* Kolom 2: Waar liggen de risico's */}
        <div style={{ padding:"14px 16px", borderRight:"1px solid #D0E4F7" }}>
          <p style={{ fontSize:10, fontWeight:700, color:"#0C2340", marginBottom:8 }}>
            ⚠️ Potentiële problemen
          </p>
          {scored.length > 0 ? (
            <>
              {hoogRisk.length > 0 && (
                <div className="rounded mb-2 px-2 py-1.5" style={{ background:"#fee2e2", border:"1px solid #fca5a5" }}>
                  <p style={{ fontSize:9, fontWeight:700, color:"#b91c1c", marginBottom:2 }}>
                    Hoog risico ({hoogRisk.length})
                  </p>
                  <p style={{ fontSize:9, color:"#7f1d1d", lineHeight:1.4 }}>
                    {hoogRisk.map(a => `${dn(a, useSecondaryName)} (risico ${a.sc.risico?.toFixed(1)}/5)`).join(" · ")}
                  </p>
                </div>
              )}
              {laagMit.length > 0 && (
                <div className="rounded mb-2 px-2 py-1.5" style={{ background:"#ffedd5", border:"1px solid #fed7aa" }}>
                  <p style={{ fontSize:9, fontWeight:700, color:"#c2410c", marginBottom:2 }}>
                    Lage weerbaarheid ({laagMit.length})
                  </p>
                  <p style={{ fontSize:9, color:"#7c2d12", lineHeight:1.4 }}>
                    {laagMit.map(a => `${dn(a, useSecondaryName)} (mitigatie ${a.sc.mitigatie?.toFixed(1)}/5)`).join(" · ")}
                  </p>
                </div>
              )}
              {hoogRisk.length === 0 && laagMit.length === 0 && (
                <div className="rounded px-2 py-1.5" style={{ background:"#dcfce7", border:"1px solid #86efac" }}>
                  <p style={{ fontSize:9, color:"#15803d" }}>Geen urgente problemen geïdentificeerd op basis van huidige assessments.</p>
                </div>
              )}
            </>
          ) : (
            <p style={{ fontSize:9, color:"#9ca3af" }}>Voeg applicaties toe om problemen in kaart te brengen.</p>
          )}
        </div>

        {/* Kolom 3: Waar staat de data */}
        <div style={{ padding:"14px 16px" }}>
          <p style={{ fontSize:10, fontWeight:700, color:"#0C2340", marginBottom:8 }}>
            🌍 Datalocatie & jurisdictie
          </p>
          {apps.length > 0 ? (() => {
            const a1Scores = apps.map(a => ({ name:dn(a, useSecondaryName), v: a.scores["A1"]||0, d: a.scores["A3"]||0 }));
            const nonEU   = a1Scores.filter(a => a.v >= 4);
            const onduidelijk = a1Scores.filter(a => a.v === 0);
            const euOk    = a1Scores.filter(a => a.v > 0 && a.v < 4);
            return (
              <>
                {nonEU.length > 0 && (
                  <div className="rounded mb-2 px-2 py-1.5" style={{ background:"#fee2e2", border:"1px solid #fca5a5" }}>
                    <p style={{ fontSize:9, fontWeight:700, color:"#b91c1c", marginBottom:2 }}>
                      Niet-EU jurisdictie ({nonEU.length})
                    </p>
                    <p style={{ fontSize:9, color:"#7f1d1d", lineHeight:1.4 }}>
                      {nonEU.map(a=>a.name).join(", ")}
                    </p>
                  </div>
                )}
                {euOk.length > 0 && (
                  <div className="rounded mb-2 px-2 py-1.5" style={{ background:"#dcfce7", border:"1px solid #86efac" }}>
                    <p style={{ fontSize:9, fontWeight:700, color:"#15803d", marginBottom:2 }}>
                      EU/beheersbaar ({euOk.length})
                    </p>
                    <p style={{ fontSize:9, color:"#14532d", lineHeight:1.4 }}>
                      {euOk.map(a=>dn(a, useSecondaryName)).join(", ")}
                    </p>
                  </div>
                )}
                {onduidelijk.length > 0 && (
                  <div className="rounded px-2 py-1.5" style={{ background:"#f3f4f6", border:"1px solid #e5e7eb" }}>
                    <p style={{ fontSize:9, color:"#6b7280" }}>
                      Nog niet beoordeeld ({onduidelijk.length}): {onduidelijk.map(a=>dn(a, useSecondaryName)).join(", ")}
                    </p>
                  </div>
                )}
              </>
            );
          })() : (
            <p style={{ fontSize:9, color:"#9ca3af" }}>Nog geen data beschikbaar.</p>
          )}
        </div>
      </div>
    </div>
  );
}

// ── DictuRadarSVG — spindiagram voor DICTU soevereiniteitscheck ──
// Alle 4 assen hebben dezelfde richting: hoger = meer soeverein (goed)
function DictuRadarSVG({ apps, W = 480, H = 380, useSecondaryName = false }) {
  const [tip, setTip] = React.useState(null);
  const [hiddenInRadar, setHiddenInRadar] = React.useState(new Set());
  const svgRef = React.useRef(null);

  const visibleAppsRadar = apps.filter((_, i) => !hiddenInRadar.has(i));
  
  const dims = [
    { key:"2.1", label:"Data residency" },
    { key:"2.2", label:"Technische beveiliging" },
    { key:"2.3", label:"Juridische bescherming" },
    { key:"4.1", label:"EU-infrastructuur" },
  ];

  const N = dims.length, maxV = 5, LEVELS = [1,2,3,4,5];
  const LEGEND_W = 115;
  const chartW = W - LEGEND_W - 10;
  const PAD = 46; // ruimte voor as-labels rondom het web
  const maxR = Math.min(chartW - 20, H - 60) / 2 - PAD;
  const cx = chartW / 2;
  const cy = H / 2;

  const axisAngle = i => (2 * Math.PI * i / N) - Math.PI / 2;
  const pt = (i, v) => {
    const r = (v / maxV) * maxR, a = axisAngle(i);
    return [cx + r * Math.cos(a), cy + r * Math.sin(a)];
  };
  const anchor = i => {
    const x = Math.cos(axisAngle(i));
    return x > 0.3 ? "start" : x < -0.3 ? "end" : "middle";
  };
  const labelPt = i => {
    const r = maxR + 28, a = axisAngle(i);
    return [cx + r * Math.cos(a), cy + r * Math.sin(a)];
  };

  // Mouse move: bereken dichtstbijzijnd punt in SVG-coördinaten
  const handleMouseMove = e => {
    const svgEl = svgRef.current;
    if (!svgEl) return;
    const rect  = svgEl.getBoundingClientRect();
    const scale = W / rect.width;
    const mx = (e.clientX - rect.left) * scale;
    const my = (e.clientY - rect.top)  * scale;

    let closest = null, minDist = 22;
    apps.forEach((app, ai) => {
      const color = appColor(ai);
      dims.forEach((d, i) => {
        const v = app.scores[d.key] || 0;
        if (!v) return;
        const [basePx, basePy] = pt(i, v);
        const jitterAngle = (ai * 2 * Math.PI / Math.max(apps.length, 1)) + axisAngle(i);
        const jitterR = ai === 0 ? 0 : Math.min(ai * 3.5, 12);
        const px = basePx + Math.cos(jitterAngle) * jitterR;
        const py = basePy + Math.sin(jitterAngle) * jitterR;
        const dist = Math.sqrt((mx - px) ** 2 + (my - py) ** 2);
        if (dist < minDist) {
          minDist = dist;
          const appDisplayName = useSecondaryName && app.nameSecondary ? app.nameSecondary : app.name;
          closest = { sx: px, sy: py, appName: appDisplayName, dimLabel: d.label, dimKey: d.key, value: v, color };
        }
      });
    });
    setTip(closest || null);
  };

  // Tooltip positie — altijd binnen SVG
  const TW = 192, TH = 60;
  const tipX = tip ? Math.min(Math.max(tip.sx - TW / 2, 6), W - TW - 6) : 0;
  const tipY = tip ? Math.max(tip.sy - TH - 18, 6) : 0;

  return (
    <svg ref={svgRef} viewBox={`-8 -8 ${W+16} ${H+16}`} width="100%"
      style={{ display:"block", overflow:"visible", cursor:"crosshair" }}
      onMouseMove={handleMouseMove}
      onMouseLeave={() => setTip(null)}>

      {/* Grid — subtiel, strak */}
      <polygon points={dims.map((_, i) => pt(i, 5).join(",")).join(" ")}
        fill="rgba(241,245,249,0.6)" stroke="none" />
      {LEVELS.map(lv => (
        <polygon key={lv}
          points={dims.map((_, i) => pt(i, lv).join(",")).join(" ")}
          fill="none"
          stroke={lv === 5 ? "#cbd5e1" : "#e9ecef"}
          strokeWidth={lv === 5 ? 0.8 : 0.4} />
      ))}
      {/* Schaalcijfers langs bovenas */}
      {LEVELS.map(lv => {
        const [sx, sy] = pt(0, lv);
        return <text key={lv} x={sx + 4} y={sy + 3} fontSize={7} fill="#94a3b8" fontFamily="system-ui">{lv}</text>;
      })}
      {dims.map((_, i) => {
        const [x, y] = pt(i, 5);
        return <line key={i} x1={cx} y1={cy} x2={x} y2={y} stroke="#d1d5db" strokeWidth={1} />;
      })}
      {/* Gridwaarden op eerste as */}
      {LEVELS.map(lv => {
        const [x, y] = pt(0, lv);
        return <text key={lv} x={x + 5} y={y + 3} fill="#bbb" fontSize={8} fontFamily="system-ui">{lv}</text>;
      })}

      {/* Polygonen — strak, dun, licht gevuld */}
      {visibleAppsRadar.map((app, ai) => {
        const origIdx = apps.indexOf(app);
        const color = appColor(origIdx);
        const scores = dims.map(d => app.scores[d.key] || 0);
        if (scores.every(v => v === 0)) return null;
        const poly = dims.map((d, i) => pt(i, scores[i]).join(",")).join(" ");
        return (
          <polygon key={app.id || ai} points={poly}
            fill={color} fillOpacity={0.07}
            stroke={color} strokeWidth={1.5} strokeLinejoin="round"
            strokeDasharray={ai > 4 ? "5 3" : "none"} />
        );
      })}

      {/* Stippen — klein, strak, jitter bij overlap */}
      {visibleAppsRadar.map((app, ai) => {
        const origIdx = apps.indexOf(app);
        const color = appColor(origIdx);
        const scores = dims.map(d => app.scores[d.key] || 0);
        return dims.map((d, i) => {
          const v = scores[i];
          if (v === 0) return null;
          const [basePx, basePy] = pt(i, v);
          const jitterAngle = (ai * 2 * Math.PI / Math.max(visibleAppsRadar.length, 1)) + axisAngle(i);
          const jitterR = visibleAppsRadar.length <= 1 ? 0 : Math.min(ai * 3, 9);
          const px = basePx + Math.cos(jitterAngle) * jitterR;
          const py = basePy + Math.sin(jitterAngle) * jitterR;
          const active = tip && tip.appName === (useSecondaryName && app.nameSecondary ? app.nameSecondary : app.name) && tip.dimKey === d.key;
          return (
            <g key={`${ai}-${i}`}>
              {jitterR > 0 && (
                <line x1={basePx} y1={basePy} x2={px} y2={py}
                  stroke={color} strokeWidth={0.8} strokeOpacity={0.25} strokeDasharray="2 2" />
              )}
              <circle cx={px} cy={py}
                r={active ? 6 : 4}
                fill={color} stroke="white"
                strokeWidth={active ? 2 : 1}
                opacity={0.95} />
              {active && (
                <circle cx={px} cy={py} r={9}
                  fill="none" stroke={color} strokeWidth={1} strokeOpacity={0.35} />
              )}
            </g>
          );
        });
      })}

      {/* As-labels */}
      {dims.map((d, i) => {
        const [lx, ly] = labelPt(i);
        const words = d.label.split(" ");
        const l1 = words.slice(0, Math.ceil(words.length / 2)).join(" ");
        const l2 = words.slice(Math.ceil(words.length / 2)).join(" ");
        return (
          <text key={i} x={lx} y={ly - (l2 ? 6 : 0)}
            textAnchor={anchor(i)} fill="#374151"
            fontSize={11} fontWeight={600} fontFamily="system-ui">
            {l1}
            {l2 && <tspan x={lx} dy={13}>{l2}</tspan>}
          </text>
        );
      })}

      {/* Legenda rechts in SVG als foreignObject — scrollbaar, met toggle per app */}
      <foreignObject x={chartW + 10} y={4} width={LEGEND_W} height={H - 8}>
        <div xmlns="http://www.w3.org/1999/xhtml"
          style={{ height:"100%", overflowY:"auto", display:"flex", flexDirection:"column", gap:3, paddingRight:2 }}>
          {apps.slice(0, 50).map((app, ai) => {
            const color = appColor(ai);
            const scores = dims.map(d => app.scores[d.key] || 0);
            const heeftData = !scores.every(v => v === 0);
            const isHidden = hiddenInRadar.has(ai);
            return (
              <button key={app.id || ai}
                onClick={() => setHiddenInRadar(prev => {
                  const next = new Set(prev);
                  next.has(ai) ? next.delete(ai) : next.add(ai);
                  return next;
                })}
                style={{
                  display:"flex", alignItems:"center", gap:4,
                  padding:"2px 4px", borderRadius:3, cursor:"pointer", textAlign:"left",
                  background: isHidden ? "#f9fafb" : `${color}10`,
                  border: `1px solid ${isHidden ? "#e5e7eb" : color}33`,
                  opacity: !heeftData ? 0.4 : 1,
                  flexShrink:0
                }}>
                <span style={{
                  width:7, height:7, borderRadius:"50%", flexShrink:0,
                  background: isHidden ? "#d1d5db" : color,
                }}/>
                <span style={{
                  fontSize:8, color: isHidden ? "#9ca3af" : "#374151",
                  fontFamily:"system-ui", lineHeight:1.3,
                  textDecoration: isHidden ? "line-through" : "none",
                  overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap", maxWidth:90
                }}>
                  {dn(app, useSecondaryName)}
                </span>
              </button>
            );
          })}
          {hiddenInRadar.size > 0 && (
            <button onClick={() => setHiddenInRadar(new Set())}
              style={{ fontSize:8, color:"#1A56A0", background:"#EBF3FF", border:"1px solid #D0E4F7",
                       borderRadius:3, padding:"2px 6px", cursor:"pointer", marginTop:2 }}>
              Alles tonen
            </button>
          )}
          <div style={{ borderTop:"1px solid #f1f5f9", marginTop:4, paddingTop:4 }}>
            <div style={{ fontSize:8, color:"#9ca3af", lineHeight:1.4 }}>
              Klik om te verbergen
            </div>
          </div>
        </div>
      </foreignObject>

      {/* Noot als er apps zijn zonder DICTU-scores */}
      {apps.some(app => dims.map(d => app.scores[d.key] || 0).every(v => v === 0)) && (
        <text x={cx} y={H - 0} textAnchor="middle" fontSize={9} fill="#9ca3af" fontFamily="system-ui"
          fontStyle="italic">
          * DICTU-vragen nog niet ingevuld — niet zichtbaar in diagram
        </text>
      )}

      {/* Tooltip — puur SVG, altijd bovenop */}
      {tip && (
        <g style={{ pointerEvents:"none" }}>
          <rect x={tipX + 2} y={tipY + 2} width={TW} height={TH} rx={5} fill="rgba(0,0,0,0.12)" />
          <rect x={tipX} y={tipY} width={TW} height={TH} rx={5}
            fill="white" stroke={tip.color} strokeWidth={1.5} />
          <rect x={tipX} y={tipY} width={TW} height={20} rx={5} fill={tip.color} />
          <rect x={tipX} y={tipY + 15} width={TW} height={5} fill={tip.color} />
          <text x={tipX + 10} y={tipY + 14}
            fill="white" fontSize={10} fontWeight={700} fontFamily="system-ui">
            {tip.appName.substring(0, 22)}
          </text>
          <text x={tipX + 10} y={tipY + 35}
            fill="#374151" fontSize={9} fontFamily="system-ui">
            {tip.dimKey} — {tip.dimLabel}
          </text>
          <text x={tipX + 10} y={tipY + 52}
            fill={tip.color} fontSize={14} fontWeight={700} fontFamily="system-ui">
            {tip.value} / 5
          </text>
          <text x={tipX + 46} y={tipY + 52}
            fill="#6b7280" fontSize={9} fontFamily="system-ui">
            {tip.value >= 4 ? "soeverein" : tip.value >= 3 ? "acceptabel" : "aandacht vereist"}
          </text>
        </g>
      )}
    </svg>
  );
}


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

function ScoreBtn({ s, selected, label, desc, dir, onClick, readOnly = false }) {
  const colors = dir === "fwd" ? BTN_COLORS_FWD : BTN_COLORS_INV;
  const c = colors[s - 1];
  return (
    <button
      onClick={readOnly ? undefined : onClick}
      title={readOnly ? "Alleen-lezen — open via Beheer om te wijzigen" : desc}
      disabled={readOnly}
      className="flex-1 py-2 px-1 rounded-lg border-2 text-center transition-all"
      style={selected
        ? { borderColor: c, background: c, color: "#fff",
            cursor: readOnly ? "not-allowed" : "pointer",
            opacity: readOnly ? 0.7 : 1 }
        : { borderColor: "#e5e7eb", background: readOnly ? "#f3f4f6" : "#f9fafb",
            color: readOnly ? "#9ca3af" : "#374151",
            cursor: readOnly ? "not-allowed" : "pointer" }
      }>
      <div style={{ fontWeight: 700, fontSize: 14 }}>{s}</div>
      <div style={{ fontSize: 10, lineHeight: 1.2, marginTop: 2 }}>{label}</div>
    </button>
  );
}

function QuestionCard({ q, value, onChange, dir, note, onNoteChange, useSecondaryName = false, appName = "", appNameSecondary = "", readOnly = false }) {
  // Bereken de weergavenaam op basis van de toggle
  const displayedNote = React.useMemo(() => {
    if (!note || !appName || !appNameSecondary || appName === appNameSecondary) return note;
    if (useSecondaryName) {
      return note.replace(new RegExp(appName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi'), appNameSecondary);
    }
    return note;
  }, [note, useSecondaryName, appName, appNameSecondary]);
  return (
    <div className="mb-3" style={{
        background: readOnly ? "#fafafa" : "#fff",
        border: readOnly ? "1px solid #e5e7eb" : "1px solid #D0E4F7",
        borderRadius:4, padding:16,
        position: "relative"
      }}>
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
          <ScoreBtn key={s} s={s} selected={value === s} label={label} desc={desc} dir={dir} onClick={() => !readOnly && onChange(s)} readOnly={readOnly} />
        ))}
      </div>
      {value > 0 && (
        <p className="mt-2 text-gray-400 italic" style={{ fontSize: 11 }}>
          {q.scores.find(sc => sc.s === value)?.desc}
        </p>
      )}
      {/* Motivatieveld */}
      <div className="mt-3">
        <label style={{ fontSize:10, color:"#6b7280", fontWeight:600, display:"block", marginBottom:3 }}>
          Motivatie / toelichting score <span style={{ fontWeight:400 }}>(optioneel)</span>
          {useSecondaryName && appNameSecondary && appName !== appNameSecondary && (
            <span style={{ marginLeft:6, fontSize:9, color:"#E87722", fontWeight:600 }}>
              🏷 Weergave: secundaire naam
            </span>
          )}
        </label>
        <textarea
          value={displayedNote || ""}
          onChange={e => {
            if (!onNoteChange || readOnly) return;
            // Sla altijd op met de primaire naam — zet secundaire terug als die actief is
            let tekst = e.target.value;
            if (useSecondaryName && appNameSecondary && appName && appName !== appNameSecondary) {
              tekst = tekst.replace(new RegExp(appNameSecondary.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi'), appName);
            }
            onNoteChange(tekst);
          }}
          placeholder="Waarom kies je voor deze score? Voeg context toe voor toekomstig gebruik..."
          rows={2}
          style={{
            width:"100%", fontSize:11, padding:"7px 10px",
            border:"1px solid #D0E4F7", borderRadius:4,
            color:"#374151", lineHeight:1.5, resize:"vertical",
            background: readOnly ? "#f3f4f6" : (note ? "#f0fdf4" : "#f8fafc"),
            borderColor: readOnly ? "#e5e7eb" : (note ? "#86efac" : "#D0E4F7"),
            outline:"none", fontFamily:"inherit", boxSizing:"border-box"
          }}
          onFocus={e => e.target.style.borderColor = "#1A56A0"}
          onBlur={e => e.target.style.borderColor = note ? "#86efac" : "#D0E4F7"}
        />
        {note && (
          <p style={{ fontSize:9, color:"#16a34a", marginTop:2 }}>✓ Motivatie opgeslagen</p>
        )}
      </div>
    </div>
  );
}

// ──────────────────────────────────────────────────────────────
// APP
// ──────────────────────────────────────────────────────────────

// ── MiniGeoKaart — compacte wereldkaart voor Portfolio-pagina ────────────────
function MiniGeoKaart({ geoApps, proj, W, H, REGIO_LON_LAT, REGIO_KLEUR, displayName }) {
  const [worldData, setWorldData] = React.useState(null);

  React.useEffect(function() {
    fetch("https://cdn.jsdelivr.net/npm/world-atlas@2/countries-110m.json")
      .then(function(r) { return r.json(); })
      .then(function(topo) { setWorldData(topo); })
      .catch(function() {});
  }, []);

  const countries = React.useMemo(function() {
    if (!worldData || !worldData.objects) return [];
    try {
      const sc = worldData.transform ? worldData.transform.scale : [1,1];
      const tr = worldData.transform ? worldData.transform.translate : [0,0];
      const arcs = worldData.arcs;
      function decodeArc(idx) {
        const neg = idx < 0;
        const arc = arcs[neg ? (idx * -1) - 1 : idx];
        let x = 0; let y = 0;
        const pts = arc.map(function(p) { x += p[0]; y += p[1]; return [x, y]; });
        if (neg) pts.reverse();
        return pts;
      }
      function toCoord(p) { return [p[0]*sc[0]+tr[0], p[1]*sc[1]+tr[1]]; }
      function ringToCoords(ring) {
        const pts = [];
        ring.forEach(function(i) { decodeArc(i).forEach(function(p) { pts.push(toCoord(p)); }); });
        return pts;
      }
      const EU_CODES = new Set([40,56,100,191,196,203,208,233,246,250,276,300,348,372,380,428,440,442,470,528,616,620,642,703,705,724,752]); // 826=VK bewust weggelaten
      return worldData.objects.countries.geometries.map(function(g) {
        let coords;
        if (g.type === "Polygon") coords = g.arcs.map(function(ring) { return ringToCoords(ring); });
        else if (g.type === "MultiPolygon") coords = g.arcs.map(function(poly) { return poly.map(function(ring) { return ringToCoords(ring); }); });
        else return null;
        return { type:"Feature", id:g.id, isEU: EU_CODES.has(parseInt(g.id)), geometry:{ type:g.type, coordinates:coords } };
      }).filter(Boolean);
    } catch(e) { return []; }
  }, [worldData]);

  // Bereken de bounding box van alle stippen om op in te zoomen
  // met padding eromheen zodat stippen nooit aan de rand zitten
  const PADDING = 60; // pixels padding rondom de stippen

  const allPts = React.useMemo(function() {
    const pts = [];
    geoApps.forEach(function(a) {
      if (a.a1 > 0) {
        const ll = REGIO_LON_LAT[a.jRegio];
        if (ll) { const p = proj(ll); if (p) pts.push(p); }
      }
      if (a.a3 > 0) {
        const ll = REGIO_LON_LAT[a.dRegio];
        if (ll) { const p = proj(ll); if (p) pts.push(p); }
      }
    });
    return pts;
  }, [geoApps, proj, REGIO_LON_LAT]);

  // Bereken viewBox op basis van stippen-bounding box
  const viewBox = React.useMemo(function() {
    if (allPts.length === 0) return "0 0 " + W + " " + H;
    const xs = allPts.map(function(p) { return p[0]; });
    const ys = allPts.map(function(p) { return p[1]; });
    const minX = Math.max(0, Math.min.apply(null, xs) - PADDING);
    const maxX = Math.min(W, Math.max.apply(null, xs) + PADDING);
    const minY = Math.max(0, Math.min.apply(null, ys) - PADDING);
    const maxY = Math.min(H, Math.max.apply(null, ys) + PADDING);
    const vw = maxX - minX;
    const vh = maxY - minY;
    // Zorg voor minimale grootte en behoud aspect ratio
    const minSize = 120;
    const cx = (minX + maxX) / 2;
    const cy = (minY + maxY) / 2;
    const fw = Math.max(vw, minSize);
    const fh = Math.max(vh, minSize);
    return (cx - fw/2) + " " + (cy - fh/2) + " " + fw + " " + fh;
  }, [allPts, W, H]);

  const pathGen = React.useMemo(function() { return d3.geoPath().projection(proj); }, [proj]);
  const sphere  = { type:"Sphere" };

  // Groepeer stippen
  const jurisPerRegio = {};
  const dataPerRegio  = {};
  geoApps.forEach(function(a) {
    if (a.a1 > 0) {
      if (!jurisPerRegio[a.jRegio]) jurisPerRegio[a.jRegio] = [];
      jurisPerRegio[a.jRegio].push(a);
    }
    if (a.a3 > 0) {
      if (!dataPerRegio[a.dRegio]) dataPerRegio[a.dRegio] = [];
      dataPerRegio[a.dRegio].push(a);
    }
  });

  return (
    <svg viewBox={viewBox} style={{ width:"100%", height:"auto", display:"block", borderRadius:4, background:"#bfdbfe" }}>
      {/* Achtergrond oceaan */}
      <path d={pathGen(sphere)} fill="#bfdbfe"/>
      {/* Landen — alleen kleur, geen rand */}
      {countries.map(function(c) {
        return <path key={c.id} d={pathGen(c)}
          fill={c.isEU ? "#dbeafe" : "#f0f4f8"}
          stroke="#d1dde8" strokeWidth="0.25"/>;
      })}
      {/* EU outline subtiel */}
      {countries.filter(function(c) { return c.isEU; }).map(function(c) {
        return <path key={"eu_"+c.id} d={pathGen(c)} fill="#dbeafe" stroke="#93c5fd" strokeWidth="0.4" opacity="0.7"/>;
      })}

      {/* Jurisdictie = kleine gevulde cirkel in app-kleur */}
      {Object.entries(jurisPerRegio).map(function([regio, items]) {
        const ll = REGIO_LON_LAT[regio];
        if (!ll) return null;
        const pt = proj(ll);
        if (!pt) return null;
        return items.map(function(a, idx) {
          const total = items.length;
          const angle = total <= 1 ? -Math.PI/2 : (idx * 2 * Math.PI / total) - Math.PI/2;
          const spread = total <= 1 ? 0 : 10;
          const cx = pt[0] + Math.cos(angle) * spread - 2;
          const cy = pt[1] + Math.sin(angle) * spread - 2;
          return (
            <g key={a.id+"_j"}>
              <circle cx={cx} cy={cy} r="3.5" fill={a.appKleur} stroke="white" strokeWidth="1"/>
            </g>
          );
        });
      })}

      {/* Datalocatie = ring in app-kleur */}
      {Object.entries(dataPerRegio).map(function([regio, items]) {
        const ll = REGIO_LON_LAT[regio];
        if (!ll) return null;
        const pt = proj(ll);
        if (!pt) return null;
        return items.map(function(a, idx) {
          const total = items.length;
          const angle = total <= 1 ? -Math.PI/2 : (idx * 2 * Math.PI / total) - Math.PI/2;
          const spread = total <= 1 ? 0 : 10;
          const cx = pt[0] + Math.cos(angle) * spread + 3;
          const cy = pt[1] + Math.sin(angle) * spread + 4;
          return (
            <g key={a.id+"_d"}>
              <circle cx={cx} cy={cy} r="3.5" fill="white" fillOpacity="0.9" stroke={a.appKleur} strokeWidth="1.8"/>
            </g>
          );
        });
      })}


    </svg>
  );
}

// ── WorldMapD3 — echte wereldkaart via D3 Natural Earth projectie ─────────────
function WorldMapD3({ scored, jurisGroups, dataGroups, geoHoverId, setGeoHoverId,
                      geoTooltip, setGeoTooltip, risicoKleur, displayName, REGIO_COORDS }) {
  const safeJuris = jurisGroups || {};
  const safeData  = dataGroups  || {};
  const [worldData, setWorldData] = React.useState(null);
  const [loading,   setLoading]   = React.useState(true);
  const [transform, setTransform] = React.useState({ k:1, x:0, y:0 });
  const svgRef = React.useRef(null);
  const W = 960, H = 500;

  React.useEffect(() => {
    fetch("https://cdn.jsdelivr.net/npm/world-atlas@2/countries-110m.json")
      .then(r => r.json())
      .then(topo => { setWorldData(topo); setLoading(false); })
      .catch(() => setLoading(false));
  }, []);

  // D3 zoom instellen
  React.useEffect(() => {
    if (!svgRef.current) return;
    const zoom = d3.zoom()
      .scaleExtent([1, 8])
      .translateExtent([[0,0],[W,H]])
      .on("zoom", function(event) {
        setTransform({ k: event.transform.k, x: event.transform.x, y: event.transform.y });
      });
    d3.select(svgRef.current).call(zoom);
    return () => d3.select(svgRef.current).on(".zoom", null);
  }, [loading]);

  const projection = React.useMemo(() =>
    d3.geoNaturalEarth1()
      .scale(153)
      .translate([W / 2, H / 2]),
    []
  );

  const pathGen = React.useMemo(() => d3.geoPath().projection(projection), [projection]);

  const countries = React.useMemo(() => {
    if (!worldData || !worldData.objects) return [];
    try {
      const obj = worldData.objects.countries;
      if (!obj) return [];
      const sc = worldData.transform ? worldData.transform.scale : [1,1];
      const tr = worldData.transform ? worldData.transform.translate : [0,0];
      const arcs = worldData.arcs;

      function decodeArcSafe(idx) {
        const neg = idx < 0;
        const arc = arcs[neg ? (idx * -1) - 1 : idx];
        let x = 0; let y = 0;
        const pts = arc.map(function(p) { x += p[0]; y += p[1]; return [x, y]; });
        if (neg) pts.reverse();
        return pts;
      }
      function toCoordSafe(p) {
        return [p[0] * sc[0] + tr[0], p[1] * sc[1] + tr[1]];
      }
      function ringToCoords(ring) {
        const pts = [];
        ring.forEach(function(i) {
          decodeArcSafe(i).forEach(function(p) { pts.push(toCoordSafe(p)); });
        });
        return pts;
      }

      return obj.geometries.map(function(g) {
        let coords;
        if (g.type === "Polygon") {
          coords = g.arcs.map(function(ring) { return ringToCoords(ring); });
        } else if (g.type === "MultiPolygon") {
          coords = g.arcs.map(function(poly) {
            return poly.map(function(ring) { return ringToCoords(ring); });
          });
        } else { return null; }
        return {
          type: "Feature",
          id: g.id,
          properties: g.properties || {},
          geometry: { type: g.type, coordinates: coords }
        };
      }).filter(Boolean);
    } catch(e) { console.error(e); return []; }
  }, [worldData]);

  const graticule = React.useMemo(() => d3.geoGraticule()(), []);
  const sphere    = { type: "Sphere" };

  // EU landen ISO codes voor highlight
  const EU_CODES = new Set([
    40,56,100,191,196,203,208,233,246,250,276,300,
    348,372,380,428,440,442,470,528,616,620,642,703,
    705,724,752,826,31,8
  ]);

  // Kleur per land
  function landKleur(id) {
    const numId = parseInt(id);
    if (EU_CODES.has(numId)) return "#dbeafe";
    return "#e5e9f0";
  }

  const a1lbl = ["","EU/EER volledig","EU/EER beperkt","Adequaat + risico","SCCs, geen adequaat","Geen waarborgen"];
  const a3lbl = ["","EU/EER contractueel","EU/EER + adequaat","EU/EER, geen garantie","Deels buiten EU","Buiten EU"];

  function resetZoom() {
    setTransform({ k:1, x:0, y:0 });
    if (svgRef.current) {
      const zb = d3.zoom().scaleExtent([1,8]).translateExtent([[0,0],[W,H]])
        .on("zoom", function(ev) { setTransform({ k:ev.transform.k, x:ev.transform.x, y:ev.transform.y }); });
      d3.select(svgRef.current).call(zb.transform, d3.zoomIdentity);
    }
  }

  return (
    <div style={{ position:"relative", background:"#bfdbfe" }}>
      {!loading && transform.k > 1 && (
        <div style={{ position:"absolute", top:8, right:8, zIndex:10, display:"flex", gap:6, alignItems:"center" }}>
          <span style={{ fontSize:10, background:"rgba(255,255,255,0.88)", padding:"2px 8px", borderRadius:3, color:"#6b7280" }}>
            {Math.round(transform.k * 100)}%
          </span>
          <button onClick={resetZoom}
            style={{ fontSize:10, background:"white", border:"1px solid #D0E4F7", borderRadius:4,
                     padding:"3px 10px", color:"#1A56A0", cursor:"pointer", fontWeight:600 }}>
            ↺ Reset
          </button>
        </div>
      )}
      {loading && (
        <div style={{ padding:40, textAlign:"center", color:"#9ca3af", fontSize:12 }}>
          Kaart laden...
        </div>
      )}
      {!loading && (
        <svg ref={svgRef} viewBox={`0 0 ${W} ${H}`}
          style={{ width:"100%", height:"auto", display:"block", cursor: transform.k > 1 ? "grab" : "default" }}>
          {/* Oceaan achtergrond buiten zoom */}
          <rect width={W} height={H} fill="#bfdbfe"/>
          <g transform={`translate(${transform.x},${transform.y}) scale(${transform.k})`}>
          {/* Oceaan sphere */}
          <path d={pathGen(sphere)} fill="#bfdbfe"/>
          {/* Graticule */}
          <path d={pathGen(graticule)} fill="none" stroke="#94a3b8" strokeWidth="0.3" opacity="0.5"/>
          {/* Landen */}
          {countries.map(c => (
            <path key={c.id} d={pathGen(c)}
              fill={landKleur(c.id)}
              stroke="#94a3b8" strokeWidth="0.4"
              opacity="0.95"/>
          ))}
          {/* EU outline extra */}
          {countries.filter(c => EU_CODES.has(parseInt(c.id))).map(c => (
            <path key={"eu_"+c.id} d={pathGen(c)}
              fill="#dbeafe" stroke="#3b82f6" strokeWidth="0.6" opacity="0.7"/>
          ))}

          {/* ── Jurisdictie stippen (gevulde cirkel) ── */}
          {Object.entries(safeJuris).map(([regio, items]) => {
            const rc = REGIO_COORDS[regio] || REGIO_COORDS["Onbekend"];
            // Projecteer vanuit lon/lat naar SVG-pixels
            const [px, py] = projection([rc.lon, rc.lat]);
            return items.map((a, idx) => {
              const totalItems = items.length;
              const angle  = totalItems <= 1 ? 0 : (idx * 2 * Math.PI / totalItems) - Math.PI/2;
              const radius = totalItems <= 1 ? 0 : Math.min(14 + totalItems * 3, 28);
              const cx = px + Math.cos(angle) * radius;
              const cy = py + Math.sin(angle) * radius - 8;
              const kleur = risicoKleur(a.score);
              const isH  = geoHoverId === a.id + "_j";
              return (
                <g key={a.id+"_j"} style={{ cursor:"pointer" }}
                  onMouseEnter={() => { setGeoHoverId(a.id+"_j"); setGeoTooltip({ cx, cy, a, type:"juris" }); }}
                  onMouseLeave={() => { setGeoHoverId(null); setGeoTooltip(null); }}>
                  <circle cx={cx} cy={cy} r={isH ? 11 : 9}
                    fill={kleur} stroke="white" strokeWidth="2" opacity="0.93"/>
                  {isH && <circle cx={cx} cy={cy} r="14" fill="none" stroke={kleur} strokeWidth="1.5" opacity="0.4"/>}
                  <text x={cx} y={cy+3.5} textAnchor="middle" fontSize="7" fill="white" fontWeight="700"
                    style={{ pointerEvents:"none" }}>
                    {displayName(a).substring(0,3).toUpperCase()}
                  </text>
                </g>
              );
            });
          })}

          {/* ── Data-locatie stippen (omrand vierkant) ── */}
          {Object.entries(safeData).map(([regio, items]) => {
            const rc = REGIO_COORDS[regio] || REGIO_COORDS["Onbekend"];
            const [px, py] = projection([rc.lon, rc.lat]);
            return items.map((a, idx) => {
              const totalItems = items.length;
              const angle  = totalItems <= 1 ? 0 : (idx * 2 * Math.PI / totalItems) - Math.PI/2;
              const radius = totalItems <= 1 ? 0 : Math.min(14 + totalItems * 3, 28);
              const cx = px + Math.cos(angle) * radius + 10;
              const cy = py + Math.sin(angle) * radius + 14;
              const sz   = geoHoverId === a.id+"_d" ? 10 : 8;
              const kleur = risicoKleur(a.score);
              const isH  = geoHoverId === a.id+"_d";
              return (
                <g key={a.id+"_d"} style={{ cursor:"pointer" }}
                  onMouseEnter={() => { setGeoHoverId(a.id+"_d"); setGeoTooltip({ cx, cy, a, type:"data" }); }}
                  onMouseLeave={() => { setGeoHoverId(null); setGeoTooltip(null); }}>
                  <rect x={cx-sz} y={cy-sz} width={sz*2} height={sz*2}
                    fill="white" stroke={kleur} strokeWidth="2.5" rx="2" opacity="0.95"/>
                  {isH && <rect x={cx-13} y={cy-13} width="26" height="26" fill="none" stroke={kleur} strokeWidth="1.5" rx="3" opacity="0.4"/>}
                  <text x={cx} y={cy+3} textAnchor="middle" fontSize="6" fill={kleur} fontWeight="700"
                    style={{ pointerEvents:"none" }}>
                    {displayName(a).substring(0,3).toUpperCase()}
                  </text>
                </g>
              );
            });
          })}

          {/* ── Tooltip ── */}
          {geoTooltip && (() => {
            const { cx, cy, a, type } = geoTooltip;
            // Corrigeer voor zoom-transform
            const rawX = cx * transform.k + transform.x;
            const rawY = cy * transform.k + transform.y;
            const tx  = Math.min(rawX + 14, W - 240);
            const ty  = Math.max(rawY - 75, 5);
            const score = type === "juris" ? a.a1 : a.a3;
            const lbl   = type === "juris" ? (a1lbl[score]||"–") : (a3lbl[score]||"–");
            const kleur = risicoKleur(score);
            return (
              <g style={{ pointerEvents:"none" }}>
                <rect x={tx} y={ty} width="230" height="72" rx="5"
                  fill="white" stroke={kleur} strokeWidth="1.5"
                  style={{ filter:"drop-shadow(0 3px 8px rgba(0,0,0,0.18))" }}/>
                <text x={tx+10} y={ty+17} fontSize="11" fontWeight="700" fill="#0C2340">{displayName(a)}</text>
                {a.supplier && <text x={tx+10} y={ty+30} fontSize="9" fill="#9ca3af">{a.supplier}</text>}
                <text x={tx+10} y={ty+46} fontSize="9" fill="#374151">
                  {type === "juris" ? "Jurisdictie leverancier (A1)" : "Datalocatie servers (A3)"}: {score||"–"}/5
                </text>
                <text x={tx+10} y={ty+60} fontSize="10" fill={kleur} fontWeight="600">{lbl}</text>
              </g>
            );
          })()}

          {/* EU label in Europa — geprojecteerde positie voor lon=10, lat=54 */}
          {(() => {
            const euPt = projection([10, 54]);
            if (!euPt) return null;
            return <text x={euPt[0]} y={euPt[1]}
              textAnchor="middle" fontSize={10/transform.k} fill="#1d4ed8" opacity="0.85"
              fontStyle="italic" fontWeight="600">EU</text>;
          })()}

          </g>{/* einde zoom-g */}
          {/* Kaart legenda — buiten zoom */}
          <g transform={`translate(${W-270}, ${H-72})`}>
            <rect width="262" height="66" rx="4" fill="white" opacity="0.93" stroke="#e2e8f0" strokeWidth="1"/>
            <text x="10" y="15" fontSize="9" fontWeight="700" fill="#0C2340">Legenda</text>
            <circle cx="18" cy="29" r="7" fill="#1A56A0" stroke="white" strokeWidth="1.5"/>
            <text x="30" y="33" fontSize="9" fill="#374151">Gevuld cirkel = Jurisdictie (A1)</text>
            <rect x="11" y="44" width="14" height="14" fill="white" stroke="#1A56A0" strokeWidth="2" rx="2"/>
            <text x="30" y="54" fontSize="9" fill="#374151">Omrand vierkant = Datalocatie (A3)</text>
            {[["#16a34a","EU/EER",145],["#ca8a04","VS+DPF",185],["#ea580c","Risico",222],["#dc2626","Kritiek",248]].map(([c,l,x]) => (
              <g key={l}>
                <circle cx={x} cy="29" r="5" fill={c} stroke="white" strokeWidth="1"/>
                <text x={x+8} y="33" fontSize="8" fill="#374151">{l}</text>
              </g>
            ))}
          </g>

        </svg>
      )}
    </div>
  );
}

function GeoKaartCompact({ apps, useSecondaryName, calcScores, appColor, d3 }) {
  const [geoHidden, setGeoHidden] = React.useState(new Set());
  
  function displayName(a) {
    return useSecondaryName && a.nameSecondary ? a.nameSecondary : a.name;
  }
  
  if (!apps || apps.length === 0) return null;


  // Vaste app-kleuren palette — elke app krijgt eigen kleur
  const APP_PALETTE = [
    "#6d28d9","#0891b2","#be185d","#15803d","#b45309",
    "#1d4ed8","#dc2626","#7c3aed","#0f766e","#c2410c",
  ];

  // Bouw geo-data op (gefilterd op selectie)
  const geoApps = apps.map(function(a, i) {
    const a1 = a.scores["A1"] || 0;
    const a3 = a.scores["A3"] || 0;
    // A1: score 1-2=EU, 3=VS+adequaat, 4=VS+SCCs, 5=buiten EU
    // A3: score 1-3=EU/EER (met variaties), 4=deels buiten EU, 5=buiten EU
    // We gebruiken A1 voor jurisdictie-regio en A3 voor data-regio
    // maar de weergave-regio moet per dimensie anders worden bepaald
    function regioA1(sc) {
      if (sc <= 0)  return "Niet ingevuld";
      if (sc <= 2)  return "EU / EER";
      if (sc <= 3)  return "VS";
      if (sc <= 4)  return "Deels buiten EU";
      return "Buiten EU";
    }
    function regioA3(sc) {
      if (sc <= 0)  return "Niet ingevuld";
      if (sc <= 3)  return "EU / EER";   // score 1,2,3 = data in EU/EER
      if (sc <= 4)  return "Deels buiten EU"; // score 4 = deels buiten EU
      return "Buiten EU";        // score 5 = buiten EU
    }
    function regio(sc) { return regioA1(sc); } // default voor A1
    function regioKleur(sc) {
      if (sc <= 0)  return "#9ca3af";
      if (sc <= 2)  return "#16a34a";
      if (sc <= 3)  return "#ca8a04";
      if (sc <= 4)  return "#ea580c";
      return "#dc2626";
    }
    function regioKleurA3(sc) {
      if (sc <= 0)  return "#9ca3af";
      if (sc <= 3)  return "#16a34a";   // score 1-3 = groen (EU)
      if (sc <= 4)  return "#ea580c";   // score 4 = oranje
      return "#dc2626";
    }
    const appKleur = APP_PALETTE[i % APP_PALETTE.length];
    return { ...a, a1, a3,
      jRegio: regioA1(a1), dRegio: regioA3(a3),
      jKleur: regioKleur(a1), dKleur: regioKleurA3(a3),
      appKleur
    };
  });

  // Filter op hidden state
  const visibleGeoApps = geoApps.filter(function(a) { return !geoHidden.has(a.id); });

  // Groepeer per regio
  const REGIO_ORDER = ["EU / EER","VS","Deels buiten EU","Buiten EU","Niet ingevuld"];
  const REGIO_LON_LAT = {
    "EU / EER":     [10, 52],
    "VS":[-95, 38],
    "Deels buiten EU":  [-95, 38],
    "Buiten EU":    [100, 25],
    "Niet ingevuld":[0, 0],
  };
  const REGIO_KLEUR = {
    "EU / EER":     "#16a34a",
    "VS":"#ca8a04",
    "Deels buiten EU":  "#ea580c",
    "Buiten EU":    "#dc2626",
    "Niet ingevuld":"#9ca3af",
  };

  // Tel per regio
  const jTelling = {}; const dTelling = {};
  visibleGeoApps.forEach(function(a) {
    jTelling[a.jRegio] = (jTelling[a.jRegio]||0) + 1;
    dTelling[a.dRegio] = (dTelling[a.dRegio]||0) + 1;
  });

  // Mini SVG kaart met Natural Earth-achtige positionering
  // Gebruik de WorldMapD3 component maar mini
  const W = 600, H = 300;
  const proj = d3.geoNaturalEarth1().scale(90).translate([W/2, H/2]);

  return (
    <div className="rounded mb-4" style={{ background:"#fff", border:"1px solid #D0E4F7" }}>
      <div className="px-4 pt-3 pb-2 flex items-center justify-between">
        <div>
          <h3 className="font-bold text-sm" style={{ color:"#0C2340" }}>Geopolitieke positie — applicatielandschap</h3>
          <p className="text-xs mt-0.5" style={{ color:"#9ca3af" }}>Jurisdictie leverancier en datalocatie servers per regio</p>
        </div>
        <div className="flex items-center gap-2">
          <span className="text-xs px-2 py-0.5 rounded font-semibold"
            style={{ background: geoHidden.size === 0 ? "#f0f9f9" : "#fffbeb",
                     color: geoHidden.size === 0 ? "#0f766e" : "#92400e" }}>
            {apps.length - geoHidden.size} van {apps.length}
          </span>
          {geoHidden.size > 0 && (
            <button onClick={() => setGeoHidden(new Set())}
              className="text-xs px-2 py-0.5 font-medium"
              style={{ border:"1px solid #D0E4F7", borderRadius:3, color:"#1A56A0", background:"#EBF3FF" }}>
              Alles tonen
            </button>
          )}
        </div>
      </div>
      <div className="flex flex-wrap gap-1.5 px-4 pb-2">
        {[...apps].sort((a, b) => displayName(a).localeCompare(displayName(b), "nl", { sensitivity:"base", numeric:true })).map((a) => {
          const col = appColor(apps.indexOf(a));
          const hidden = geoHidden.has(a.id);
          return (
            <button key={a.id}
              onClick={() => setGeoHidden(p => { const n = new Set(p); n.has(a.id) ? n.delete(a.id) : n.add(a.id); return n; })}
              className="flex items-center gap-1.5 text-xs px-2 py-1 font-medium"
              style={{ borderRadius:4, border: hidden ? "1px solid #e5e7eb" : "1px solid "+col+"88",
                       background: hidden ? "#f9fafb" : col+"12", color: hidden ? "#9ca3af" : col,
                       textDecoration: hidden ? "line-through" : "none" }}>
              <span style={{ width:6, height:6, borderRadius:"50%", flexShrink:0, display:"inline-block",
                             background: hidden ? "#d1d5db" : col }}/>
              {displayName(a).substring(0,18)}
            </button>
          );
        })}
      </div>

      <div className="grid grid-cols-2 gap-0" style={{ borderTop:"1px solid #EBF3FF" }}>
        {/* Mini kaart */}
        <div style={{ borderRight:"1px solid #EBF3FF", padding:"8px 12px" }}>
          <MiniGeoKaart geoApps={visibleGeoApps} proj={proj} W={W} H={H} REGIO_LON_LAT={REGIO_LON_LAT} REGIO_KLEUR={REGIO_KLEUR} displayName={displayName}/>
          {/* Legenda onder kaart */}
          <div className="flex flex-wrap gap-x-3 gap-y-1 mt-1">
              {/* Legenda: app-kleuren + vorm-uitleg */}
            <div className="flex flex-wrap gap-x-4 gap-y-1">
    {visibleGeoApps.filter(function(a) { return a.a1 > 0 || a.a3 > 0; }).map(function(a) {
      return (
        <div key={a.id} className="flex items-center gap-1.5">
          <svg width="20" height="10" style={{flexShrink:0}}>
            <circle cx="5"  cy="5" r="3.5" fill={a.appKleur} stroke="white" strokeWidth="1"/>
            <circle cx="15" cy="5" r="3.5" fill="white" stroke={a.appKleur} strokeWidth="1.8"/>
          </svg>
          <span style={{ fontSize:9, color:"#6b7280" }}>{displayName(a)}</span>
        </div>
      );
    })}
            </div>
            <div className="flex gap-4 mt-1.5 pt-1.5" style={{ borderTop:"1px solid #f1f5f9" }}>
    <div className="flex items-center gap-1">
      <svg width="10" height="10" style={{flexShrink:0}}><circle cx="5" cy="5" r="3.5" fill="#9ca3af" stroke="white" strokeWidth="1"/></svg>
      <span style={{ fontSize:9, color:"#9ca3af" }}>Gevuld = jurisdictie</span>
    </div>
    <div className="flex items-center gap-1">
      <svg width="10" height="10" style={{flexShrink:0}}><circle cx="5" cy="5" r="3.5" fill="white" stroke="#9ca3af" strokeWidth="1.8"/></svg>
      <span style={{ fontSize:9, color:"#9ca3af" }}>Ring = datalocatie</span>
    </div>
            </div>
          </div>
        </div>

        {/* Tabel */}
        <div style={{ padding:"8px 12px" }}>
          <table style={{ width:"100%", borderCollapse:"collapse", fontSize:11 }}>
            <thead>
    <tr style={{ background:"#0C2340", color:"white" }}>
      <th style={{ padding:"5px 8px", textAlign:"left", fontSize:10 }}>Regio</th>
      <th style={{ padding:"5px 8px", textAlign:"center", fontSize:10 }}>Jurisdictie</th>
      <th style={{ padding:"5px 8px", textAlign:"center", fontSize:10 }}>Datalocatie</th>
    </tr>
            </thead>
            <tbody>
    {REGIO_ORDER.filter(function(r) { return jTelling[r] || dTelling[r]; }).map(function(r, i) {
      const k = REGIO_KLEUR[r];
      return (
        <tr key={r} style={{ background: i%2===0 ? "#f8fafc" : "white", borderBottom:"1px solid #f1f5f9" }}>
          <td style={{ padding:"5px 8px" }}>
            <div className="flex items-center gap-1.5">
              <span className="w-2 h-2 rounded-full flex-shrink-0" style={{ background:k }}/>
              <span style={{ fontSize:11, color:"#374151", fontWeight:600 }}>{r}</span>
            </div>
          </td>
          <td style={{ padding:"5px 8px", textAlign:"center" }}>
            {jTelling[r] ? (
              <span style={{ fontSize:12, fontWeight:700, color:k }}>{jTelling[r]}</span>
            ) : <span style={{ color:"#d1d5db" }}>–</span>}
          </td>
          <td style={{ padding:"5px 8px", textAlign:"center" }}>
            {dTelling[r] ? (
              <span style={{ fontSize:12, fontWeight:700, color:k }}>{dTelling[r]}</span>
            ) : <span style={{ color:"#d1d5db" }}>–</span>}
          </td>
        </tr>
      );
    })}
            </tbody>
          </table>
          {visibleGeoApps.some(function(a) { return a.a1 === 0 && a.a3 === 0; }) && (
            <p className="text-xs mt-2" style={{ color:"#9ca3af" }}>
    * Applicaties zonder A1/A3-score zijn niet meegenomen.
            </p>
          )}

          {/* Applicatie badges per regio */}
          <div className="mt-3 pt-3" style={{ borderTop:"1px solid #EBF3FF" }}>
            <p className="text-xs font-semibold mb-2" style={{ color:"#6b7280" }}>Applicaties per regio:</p>
            {["EU / EER","VS","Deels buiten EU","Buiten EU","Niet ingevuld"].map(function(regio) {
    const appsInRegio = visibleGeoApps.filter(function(a) {
      return a.jRegio === regio || a.dRegio === regio;
    });
    if (appsInRegio.length === 0) return null;
    const k = REGIO_KLEUR[regio];
    return (
      <div key={regio} className="mb-2">
        <div className="flex items-center gap-1 mb-1">
          <span className="w-2 h-2 rounded-full flex-shrink-0" style={{ background:k }}/>
          <span style={{ fontSize:10, color:k, fontWeight:700 }}>{regio}</span>
        </div>
        <div className="flex flex-wrap gap-1.5 pl-3">
          {appsInRegio.map(function(a) {
            const isJuris = a.jRegio === regio;
            const isData  = a.dRegio === regio;
            return (
              <div key={a.id} className="flex items-center gap-1 px-2 py-1 rounded"
      style={{ background:a.appKleur+"14", border:"1.5px solid "+a.appKleur+"55" }}>
      {/* Gevulde cirkel (jurisdictie) of ring (datalocatie) als mini-icoon */}
      {isJuris && (
        <svg width="10" height="10" style={{flexShrink:0}}>
          <circle cx="5" cy="5" r="4.5" fill={a.appKleur}/>
        </svg>
      )}
      {isData && !isJuris && (
        <svg width="10" height="10" style={{flexShrink:0}}>
          <circle cx="5" cy="5" r="3.5" fill="white" stroke={a.appKleur} strokeWidth="2"/>
        </svg>
      )}
      {isJuris && isData && (
        <svg width="10" height="10" style={{flexShrink:0, marginLeft:2}}>
          <circle cx="5" cy="5" r="3.5" fill="white" stroke={a.appKleur} strokeWidth="2"/>
        </svg>
      )}
      <span style={{ fontSize:10, color:"#0C2340", fontWeight:600 }}>{displayName(a)}</span>
      <span style={{ fontSize:9, color:a.appKleur, opacity:0.85 }}>
        {isJuris && isData ? "juris.+data" : isJuris ? "juris." : "data"}
      </span>
              </div>
            );
          })}
        </div>
      </div>
    );
            })}
          </div>
        </div>
      </div>
    </div>
  );
}


function App() {
  // ── Login state ─────────────────────────────────────────────
  const [loggedIn,   setLoggedIn]   = useState(() => sessionStorage.getItem("nhl_auth") === "ok");
  const [loginInput, setLoginInput] = useState("");
  const [loginError, setLoginError] = useState(false);
  const LOGIN_CODE = "Geheim";
  // API-token: hetzelfde als het inlogwachtwoord — wordt meegestuurd als request-header
  // zodat de Netlify Functions server-side kunnen controleren of het verzoek geldig is.
  const apiToken = sessionStorage.getItem("nhl_api_token") || "";

  // ── App state ────────────────────────────────────────────────
  const [apps,       setApps]      = useState([]);
  const [useSecondaryName, setUseSecondaryName] = useState(false);

  // Helper: geeft de juiste naam terug op basis van de zichtbaarheidsschakelaar
  const displayName = (app) => {
    if (!app) return "";
    if (useSecondaryName && app.nameSecondary && app.nameSecondary.trim()) {
      return app.nameSecondary.trim();
    }
    return app.name || "";
  };

  // Helper: vervangt primaire naam door secondaire (of vice versa) in motivatieteksten
  const adaptNote = (tekst, app) => {
    if (!tekst || !app) return tekst;
    const from = useSecondaryName ? app.name : (app.nameSecondary || "");
    const to   = useSecondaryName ? (app.nameSecondary || "") : app.name;
    if (!from || !to || from === to) return tekst;
    return tekst.replace(new RegExp(from.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi'), to);
  };
  const [ready,      setReady]     = useState(false);
  const [saving,     setSaving]    = useState(false);
  const [saveError,  setSaveError] = useState(false);
  const [lastSaved,  setLastSaved]  = useState(null); // timestamp van laatste succesvolle save
  const [view,       setView]      = useState("about");
  const [aboutTab,   setAboutTab]  = useState("over");
  const [snapshots,  setSnapshots] = useState(() => {
    try {
      const raw = localStorage.getItem("nhl_sov_snapshots");
      return raw ? JSON.parse(raw) : [];
    } catch { return []; }
  });
  const [selId,      setSelId]     = useState(null);
  const [step,       setStep]      = useState(0);
  const [showModal,  setShowModal] = useState(false);
  const [form,       setForm]      = useState({ name:"", nameSecondary:"", cat:"", supplier:"", owner:"", appNotes:"" });
  const [hiddenApps, setHiddenApps] = useState(new Set()); // IDs verborgen in dashboard


  const [compareHidden, setCompareHidden] = useState(new Set());
  const MAX_COMPARE = 5;

  // Auto-select laatste 5 in vergelijking als er meer zijn
  React.useEffect(function() {
    if (apps.length <= MAX_COMPARE) {
      setCompareHidden(new Set());
      return;
    }
    setCompareHidden(function(prev) {
      if (prev.size > 0) return prev;
      return new Set(apps.slice(0, apps.length - MAX_COMPARE).map(function(a) { return a.id; }));
    });
  }, [apps.length]); // IDs verborgen in vergelijking

  // Beheer (admin) state
  const [adminUnlocked, setAdminUnlocked] = useState(false);
  const [assessReadOnly, setAssessReadOnly] = useState(true); // true = alleen lezen, false = bewerken
  const [showChangelog, setShowChangelog] = useState(false);
  const [showDeleteAll, setShowDeleteAll] = useState(false);
  const [deleteConfirmText, setDeleteConfirmText] = useState("");
  const [adminPin,      setAdminPin]      = useState("");
  const [adminPinError, setAdminPinError] = useState(false);
  const [editAppId,     setEditAppId]     = useState(null);
  const [editForm,      setEditForm]      = useState({});
  const [importData,    setImportData]    = useState(null);
  const [importSel,     setImportSel]     = useState(new Set());
  const [showImport,    setShowImport]    = useState(false);
  const [geoHoverId,    setGeoHoverId]    = useState(null);
  const [geoTooltip,    setGeoTooltip]    = useState(null);
  const [geoHidden,     setGeoHidden]     = useState(new Set()); // verborgen apps op de kaart
  const ADMIN_PIN = "nhl2026";

  // Ref voor scroll-naar-boven bij stapwissel in Assess
  const assessScrollRef = useRef(null);

  // ── Laden van gedeelde data via Netlify Blobs API ───────────
  useEffect(() => {
    if (!loggedIn) return;
    async function load() {
      try {
        const r = await fetch("/api/load-data", { headers: { "x-api-token": apiToken } });
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
          headers: { "Content-Type": "application/json", "x-api-token": apiToken },
          body: JSON.stringify(apps)
        });
        if (!r.ok) throw new Error("save failed");
        // Lokale backup
        localStorage.setItem("nhl_sov_v2", JSON.stringify(apps));
        setLastSaved(new Date().toISOString());
        setSaveError(false);
      } catch {
        setSaveError(true);
        // Sla toch lokaal op als fallback
        try { localStorage.setItem("nhl_sov_v2", JSON.stringify(apps)); } catch {}
      } finally {
        setSaving(false);
      }
    }, 200);
    return () => clearTimeout(timer);
  }, [apps, ready]);

  // ── Login handler ────────────────────────────────────────────
  function handleLogin() {
    if (loginInput === LOGIN_CODE) {
      sessionStorage.setItem("nhl_auth", "ok");
      sessionStorage.setItem("nhl_api_token", loginInput); // loginInput === APP_API_TOKEN
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
                <p style={{ fontSize:10, color:"#7DD3D0" }}>Prototype · Ambassadeurslijn Digitale Soevereiniteit</p>
              </div>
            </div>
            <div className="flex items-center justify-between">
              <p style={{ fontSize:10, color:"rgba(255,255,255,0.5)" }}>
                Ambassadeurslijn Digitale Soevereiniteit
              </p>
              <span className="font-bold px-2 py-0.5" style={{ fontSize:10, color:"#0C2340", background:"#26B5AE", borderRadius:3 }}>
                {VERSION}
              </span>
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
      supplier: form.supplier, owner: form.owner, appNotes: form.appNotes, nameSecondary: form.nameSecondary,
      scores: {}, createdAt: new Date().toISOString()
    };
    setApps(p => [...p, a]);
    setForm({ name:"", nameSecondary:"", cat:"", supplier:"", owner:"", appNotes:"" });
    setShowModal(false);
    setSelId(a.id);
    setStep(0);
    setAssessReadOnly(false); // Nieuwe app: direct in bewerkingsmodus
    setView("assess");
  }

  function setScore(appId, key, val) {
    setApps(p => p.map(a => a.id === appId ? { ...a, scores: { ...a.scores, [key]: val } } : a));
  }

  function setNote(appId, key, tekst) {
    setApps(p => p.map(a => a.id === appId
      ? { ...a, notes: { ...(a.notes || {}), [key]: tekst } }
      : a
    ));
  }

  function delApp(id) {
    if (!confirm("Applicatie verwijderen? Dit kan niet ongedaan worden gemaakt.")) return;
    setApps(p => p.filter(a => a.id !== id));
    if (selId === id) { setSelId(null); setView("dashboard"); }
  }

  // ── Database export ─────────────────────────────────────────────────────────
  function exportDatabase() {
    const now   = new Date();
    const ts    = now.getFullYear().toString()
      + String(now.getMonth()+1).padStart(2,"0")
      + String(now.getDate()).padStart(2,"0")
      + "_" + String(now.getHours()).padStart(2,"0")
      + String(now.getMinutes()).padStart(2,"0");
    const payload = {
      exportedAt:  now.toISOString(),
      exportedBy:  "NHL Stenden Portfolioanalyse Digitale Soevereiniteit",
      version:     VERSION,
      appCount:    apps.length,
      apps:        apps,
    };
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
    const url  = URL.createObjectURL(blob);
    const el   = document.createElement("a");
    el.href = url;
    el.download = `NHL_Sov_Database_${ts}.json`;
    el.click();
    URL.revokeObjectURL(url);
  }

  // ── Database import verwerken ────────────────────────────────────────────────
  function handleImportFile(e) {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const parsed = JSON.parse(ev.target.result);
        if (!parsed.apps || !Array.isArray(parsed.apps)) {
          alert("Ongeldig exportbestand. Selecteer een NHL_Sov_Database_*.json bestand.");
          return;
        }
        setImportData(parsed);
        // Standaard alles geselecteerd
        setImportSel(new Set(parsed.apps.map(a => a.id)));
        setShowImport(true);
      } catch {
        alert("Fout bij het lezen van het bestand. Controleer of het een geldig JSON-exportbestand is.");
      }
    };
    reader.readAsText(file);
    e.target.value = ""; // reset input
  }

  // ── Import uitvoeren ─────────────────────────────────────────────────────────
  function executeImport() {
    if (!importData) return;
    const toImport = importData.apps.filter(a => importSel.has(a.id));
    const existingIds = new Set(apps.map(a => a.id));
    let added = 0, updated = 0;
    const newApps = [...apps];
    toImport.forEach(imp => {
      const idx = newApps.findIndex(a => a.id === imp.id);
      if (idx >= 0) {
        newApps[idx] = imp; // overschrijf bestaande
        updated++;
      } else {
        newApps.push(imp); // voeg nieuwe toe
        added++;
      }
    });
    setApps(newApps);
    setShowImport(false);
    setImportData(null);
    setImportSel(new Set());
    alert(`Import voltooid: ${added} applicatie${added !== 1 ? "s" : ""} toegevoegd, ${updated} bijgewerkt.`);
  }

  function exportXlsx() {
    const dName = (a) => (useSecondaryName && a.nameSecondary) ? a.nameSecondary : a.name;
    // Vervang ook in motivatieteksten de primaire naam door de secundaire als toggle actief is
    const dNote = (a, tekst) => {
      if (!tekst || !useSecondaryName || !a.nameSecondary || !a.name) return tekst || "";
      return tekst.replace(new RegExp(a.name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi'), a.nameSecondary);
    };
    const wb = XLSX.utils.book_new();

    const ws1 = XLSX.utils.aoa_to_sheet([
      ["Applicatie","Leverancier","Categorie","Eigenaar","Autonomiescore (1-10)","Risico","Mitigatie","Belang","DICTU Score (1-5)","Volledigheid (%)"],
      ...apps.map(a => {
        const s = calcScores(a.scores);
        return [dName(a), a.supplier, a.cat, a.owner,
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
      ["Applicatie","Leverancier", ...DAAF.map(d => `${d.key} ${d.name}`), ...DAAF.map(d => `${d.key} Motivatie`)],
      ...apps.map(a => [dName(a), a.supplier,
        ...DAAF.map(d => a.scores[d.key] || ""),
        ...DAAF.map(d => dNote(a, (a.notes || {})[d.key]))
      ])
    ]);
    XLSX.utils.book_append_sheet(wb, ws2, "DAAF Scores");

    const ws3 = XLSX.utils.aoa_to_sheet([
      ["Applicatie","Leverancier", ...DICTU.map(q => `${q.key} ${q.name}`), ...DICTU.map(q => `${q.key} Motivatie`)],
      ...apps.map(a => [dName(a), a.supplier,
        ...DICTU.map(q => a.scores[q.key] || ""),
        ...DICTU.map(q => dNote(a, (a.notes || {})[q.key]))
      ])
    ]);
    XLSX.utils.book_append_sheet(wb, ws3, "DICTU Scores");

    // Extra tabblad: alle motivaties op een rij
    const allQ2 = [...DAAF, ...DICTU];
    const ws5 = XLSX.utils.aoa_to_sheet([
      ["Applicatie","Leverancier","Vraag","Naam","Score","Motivatie"],
      ...apps.flatMap(a =>
        allQ2
          .filter(q => (a.notes || {})[q.key])
          .map(q => [
            dName(a), a.supplier || "",
            q.key, q.name,
            a.scores[q.key] || "",
            dNote(a, (a.notes || {})[q.key])
          ])
      )
    ]);
    XLSX.utils.book_append_sheet(wb, ws5, "Motivaties");

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

  function exportDashboardPdf() {
    const dName = (a) => (useSecondaryName && a.nameSecondary) ? a.nameSecondary : a.name;
    const visible = apps.filter(a => !hiddenApps.has(a.id));
    const datum   = new Date().toLocaleDateString("nl-NL", { day:"2-digit", month:"long", year:"numeric" });
    const PCOLORS = ["#1e40af","#7c3aed","#065f46","#92400e","#991b1b","#0f766e"];

    // ── Hulpfunctie: dimensiescore ────────────────────────────
    function pdfDimScore(a, letter) {
      if (letter === "A") {
        const a1=a.scores["A1"]||0, a3=a.scores["A3"]||0;
        const p=[[a1,3],[a3,2]].filter(([v])=>v>0);
        if (!p.length) return null;
        const tw=p.reduce((s,[,w])=>s+w,0);
        return p.reduce((s,[v,w])=>s+v*w,0)/tw;
      }
      const qs=DAAF.filter(d=>d.dim===letter);
      const vals=qs.map(q=>a.scores[q.key]||0).filter(v=>v>0);
      return vals.length ? vals.reduce((s,v)=>s+v,0)/vals.length : null;
    }

    // ── Samenvatting genereren ────────────────────────────────
    function generateSummary(appsArr) {
      const scored = appsArr.map(a => ({ ...a, sc: calcScores(a.scores) })).filter(a => a.sc.autonomyScore);
      if (scored.length === 0) return "<p>Nog geen volledig ingevulde assessments beschikbaar.</p>";

      const sorted    = [...scored].sort((a,b) => (a.sc.autonomyScore||0) - (b.sc.autonomyScore||0));
      const kritiek   = sorted.filter(a => a.sc.autonomyScore < 3);
      const zorg      = sorted.filter(a => a.sc.autonomyScore >= 3 && a.sc.autonomyScore < 5);
      const acceptabel= sorted.filter(a => a.sc.autonomyScore >= 5 && a.sc.autonomyScore < 7);
      const goed      = sorted.filter(a => a.sc.autonomyScore >= 7);
      const avg       = scored.reduce((s,a) => s+(a.sc.autonomyScore||0),0)/scored.length;
      const hoogsteRisico = [...scored].sort((a,b) => (b.sc.risico||0)-(a.sc.risico||0))[0];
      const besteInsol   = [...scored].sort((a,b) => (b.sc.mitigatie||0)-(a.sc.mitigatie||0))[0];
      const laagsteDictu = [...scored].filter(a=>a.sc.dictuAvg).sort((a,b)=>(a.sc.dictuAvg||5)-(b.sc.dictuAvg||5))[0];

      let tekst = `<p>Dit rapport beschrijft de uitkomsten van het digitale soevereiniteitsassessment van <strong>NHL Stenden Hogeschool</strong>, 
        uitgevoerd in het kader van de Ambassadeurslijn Digitale Soevereiniteit. 
        In deze rapportage zijn <strong>${appsArr.length} applicatie${appsArr.length!==1?"s":""}</strong> beoordeeld op basis van twee frameworks: 
        het <strong>DAAF Framework</strong> (Utrecht University) voor digitale autonomie en het <strong>DICTU Framework</strong> (Rijksoverheid) voor technische en juridische soevereiniteit.</p>`;

      tekst += `<p style="margin-top:8px;">De gemiddelde autonomiescore over alle beoordeelde applicaties bedraagt <strong>${avg.toFixed(1)} op een schaal van 1 tot 10</strong>. 
        De autonomiescore is geen maat voor hoe soeverein een applicatie is, maar voor <em>hoe urgent het autonomieprobleem is</em>: 
        een hogere score betekent dat de risico's goed zijn afgedekt of het strategisch belang beperkt is, en er dus minder reden tot zorg bestaat.`;

      if (goed.length)      tekst += ` <strong>${goed.length} applicatie${goed.length!==1?"s":""}</strong> scoort goed (≥7): ${goed.map(a=>dName(a)).join(", ")}.`;
      if (acceptabel.length) tekst += ` <strong>${acceptabel.length}</strong> scoort acceptabel (5–7): ${acceptabel.map(a=>dName(a)).join(", ")}.`;
      if (zorg.length)      tekst += ` <strong>${zorg.length}</strong> vraagt aandacht (3–5): ${zorg.map(a=>dName(a)).join(", ")}.`;
      if (kritiek.length)   tekst += ` <strong style="color:#b91c1c">${kritiek.length} applicatie${kritiek.length!==1?"s":""} scoort kritiek (&lt;3) en vraagt om directe actie: ${kritiek.map(a=>dName(a)).join(", ")}.</strong>`;
      tekst += `</p>`;

      if (hoogsteRisico) tekst += `<p style="margin-top:6px;">De hoogste risico-exposure wordt gemeten bij <strong>${dName(hoogsteRisico)}</strong> 
        (risicoscore ${hoogsteRisico.sc.risico?.toFixed(2)}), wat duidt op een combinatie van geopolitieke blootstelling en leveranciersafhankelijkheid. 
        De sterkste mitigatie-capaciteit toont <strong>${dName(besteInsol)}</strong> 
        (mitigatiescore ${besteInsol.sc.mitigatie?.toFixed(2)}): er zijn alternatieven beschikbaar, de interne kennis is geborgd en de contractuele bescherming is op orde.</p>`;

      if (laagsteDictu) tekst += `<p style="margin-top:6px;">Vanuit het DICTU-perspectief (technische en juridische soevereiniteit, schaal 1–5) verdient 
        <strong>${dName(laagsteDictu)}</strong> extra aandacht met een soevereiniteitsscore van ${laagsteDictu.sc.dictuAvg?.toFixed(1)}. 
        Dit vraagt om nadere controle van datalocatie, sleutelbeheer en juridische beschermingsclausules.</p>`;

      if (appsArr.length > 1) {
        tekst += `<p style="margin-top:6px;">Het autonomie-kwadrant hieronder plaatst elke applicatie op twee assen: 
          de horizontale as toont de gecombineerde druk van risico en strategisch belang; de verticale as toont de weerbaarheid (mitigatie). 
          Applicaties in het kwadrant <em>KRITIEK</em> (rechtsboven) vragen om de meest urgente actie: hoog risico én lage weerbaarheid. 
          Het spindiagram geeft inzicht in welke specifieke dimensies de sterkste en zwakste posities kennen.</p>`;
      }

      return tekst;
    }

    // ── Dimensieprofiel als SVG (horizontale balken) ─────────────
    function generateSpinSVG(appsArr) {
      const groups = [
        { title:"Niveau 1 — Risico-exposure (laag is beter)",   tc:"#991b1b", bg:"#fff1f2",
          grad:"#dcfce7,#fef9c3,#fca5a5,#dc2626",
          dims:[{l:"A",n:"Geopolitiek risico"},{l:"B",n:"Leveranciersafh."}] },
        { title:"Niveau 2 — Mitigatie-capaciteit (hoog is beter)", tc:"#166534", bg:"#f0fdf4",
          grad:"#dc2626,#fca5a5,#fde68a,#86efac,#16a34a",
          dims:[{l:"C",n:"Technische weerbaarheid"},{l:"D",n:"Organisatorische wb."},{l:"E",n:"Contractuele wb."}] },
        { title:"Niveau 3 — Strategisch belang (hoog = meer urgentie)", tc:"#92400e", bg:"#fff7ed",
          grad:"#fffbeb,#fde68a,#f59e0b,#b45309",
          dims:[{l:"F",n:"Organisatorisch belang"},{l:"G",n:"Data-gevoeligheid"},{l:"H",n:"Academische impact"}] },
      ];

      const W=680, barH=18, labelW=180, gapBetweenGroups=16, rowGap=6, headerH=22;
      const totalRows = groups.reduce((s,g)=>s+g.dims.length,0);
      const H = groups.length*(headerH+gapBetweenGroups) + totalRows*(barH+rowGap) + 50;
      const barW = W - labelW - 60;

      // Legenda
      const legendH = 28;
      let svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${W} ${H+legendH}" width="${W}" style="display:block;max-width:100%;font-family:Arial">`;

      // Legenda bovenaan
      appsArr.slice(0,6).forEach((a,ai) => {
        const col = PCOLORS[ai%PCOLORS.length];
        const lx = ai * 105 + 10;
        svg += `<rect x="${lx}" y="6" width="10" height="10" fill="${col}" fill-opacity="0.7" rx="2"/>`;
        svg += `<text x="${lx+14}" y="15" font-size="9" fill="#374151">${dn(a, useSecondaryName).substring(0,14)}</text>`;
      });

      let y = legendH + 4;

      groups.forEach(g => {
        // Groep header
        svg += `<rect x="0" y="${y}" width="${W}" height="${headerH}" fill="${g.bg}"/>`;
        svg += `<rect x="0" y="${y}" width="3" height="${headerH}" fill="${g.tc}"/>`;
        svg += `<text x="8" y="${y+14}" font-size="10" font-weight="700" fill="${g.tc}">${g.title}</text>`;
        y += headerH + 4;

        g.dims.forEach(d => {
          // Dim label
          svg += `<text x="8" y="${y+12}" font-size="9" font-weight="600" fill="${g.tc}">${d.l}</text>`;
          svg += `<text x="22" y="${y+12}" font-size="9" fill="#374151">${d.n}</text>`;

          // Gradient bar background (als linearGradient)
          const gradId = `grad_${d.l}`;
          svg += `<defs><linearGradient id="${gradId}" x1="0%" y1="0%" x2="100%" y2="0%">`;
          const stops = g.grad.split(",");
          stops.forEach((c,si) => {
            svg += `<stop offset="${(si/(stops.length-1)*100).toFixed(0)}%" stop-color="${c.trim()}"/>`;
          });
          svg += `</linearGradient></defs>`;
          svg += `<rect x="${labelW}" y="${y}" width="${barW}" height="${barH}" rx="3" fill="url(#${gradId})" stroke="rgba(0,0,0,0.06)" stroke-width="1"/>`;

          // Tick-labels
          [1,2,3,4,5].forEach(v => {
            const tx = labelW + (v-1)/4*barW;
            svg += `<text x="${tx.toFixed(1)}" y="${y+barH+10}" text-anchor="middle" font-size="7" fill="#9ca3af">${v}</text>`;
            if (v>1) svg += `<line x1="${tx.toFixed(1)}" y1="${y}" x2="${tx.toFixed(1)}" y2="${y+barH}" stroke="rgba(255,255,255,0.5)" stroke-width="1"/>`;
          });

          // App dots
          appsArr.slice(0,6).forEach((a,ai) => {
            const v = pdfDimScore(a,d.l);
            if (!v) return;
            const col = PCOLORS[ai%PCOLORS.length];
            const dx = labelW + (v-1)/4*barW;
            svg += `<circle cx="${dx.toFixed(1)}" cy="${(y+barH/2).toFixed(1)}" r="6" fill="${col}" stroke="white" stroke-width="2"/>`;
            svg += `<text x="${dx.toFixed(1)}" y="${(y+barH/2+3).toFixed(1)}" text-anchor="middle" font-size="6" fill="white" font-weight="700">${v.toFixed(1)}</text>`;
          });

          y += barH + rowGap;
        });
        y += gapBetweenGroups;
      });

      svg += `</svg>`;
      return svg;
    }

    // ── SVG kwadrant ─────────────────────────────────────────
    function generateKwadrantSVG(appsArr) {
      const W=680, H=360, pad={top:28, right:20, bottom:44, left:48};
      const iW=W-pad.left-pad.right, iH=H-pad.top-pad.bottom;
      const xMin=1,xMax=25,yMin=1,yMax=5,mx=13,my=3;
      const toX = v => pad.left+(v-xMin)/(xMax-xMin)*iW;
      const toY = v => pad.top+(yMax-v)/(yMax-yMin)*iH;
      const midX=toX(mx), midY=toY(my);
      let dots="";
      appsArr.forEach((a,i)=>{
        const s=calcScores(a.scores);
        if(!s.risico||!s.mitigatie||!s.belang) return;
        const cx2=toX(s.risico*s.belang), cy2=toY(s.mitigatie);
        const col=PCOLORS[i%PCOLORS.length];
        dots+=`<circle cx="${cx2.toFixed(1)}" cy="${cy2.toFixed(1)}" r="8" fill="${col}" fill-opacity="0.25" stroke="${col}" stroke-width="2"/>`;
        dots+=`<circle cx="${cx2.toFixed(1)}" cy="${cy2.toFixed(1)}" r="4" fill="${col}"/>`;
        dots+=`<rect x="${(cx2+10).toFixed(1)}" y="${(cy2-9).toFixed(1)}" width="${Math.min(dName(a).length*5.5+6,110)}" height="14" rx="2" fill="white" fill-opacity="0.85"/>`;
        dots+=`<text x="${(cx2+13).toFixed(1)}" y="${(cy2+2).toFixed(1)}" fill="${col}" font-size="9" font-weight="bold" font-family="Arial">${dName(a).substring(0,18)}</text>`;
      });
      return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${W} ${H}" width="${W}" style="display:block;max-width:100%">
        <rect x="${pad.left}" y="${pad.top}" width="${midX-pad.left}" height="${midY-pad.top}" fill="#e8f5e9"/>
        <rect x="${midX}" y="${pad.top}" width="${pad.left+iW-midX}" height="${midY-pad.top}" fill="#fff8e1"/>
        <rect x="${pad.left}" y="${midY}" width="${midX-pad.left}" height="${pad.top+iH-midY}" fill="#fff3e0"/>
        <rect x="${midX}" y="${midY}" width="${pad.left+iW-midX}" height="${pad.top+iH-midY}" fill="#fce4ec"/>
        <text x="${((pad.left+midX)/2).toFixed(0)}" y="${((pad.top+midY)/2).toFixed(0)}" text-anchor="middle" fill="#2e7d5e" font-size="12" font-weight="bold" font-style="italic" font-family="Arial">OPTIMAAL</text>
        <text x="${((midX+pad.left+iW)/2).toFixed(0)}" y="${((pad.top+midY)/2).toFixed(0)}" text-anchor="middle" fill="#e07b20" font-size="12" font-weight="bold" font-style="italic" font-family="Arial">BEHEERSBAAR</text>
        <text x="${((pad.left+midX)/2).toFixed(0)}" y="${((midY+pad.top+iH)/2).toFixed(0)}" text-anchor="middle" fill="#e07b20" font-size="12" font-weight="bold" font-style="italic" font-family="Arial">AANDACHTSPUNT</text>
        <text x="${((midX+pad.left+iW)/2).toFixed(0)}" y="${((midY+pad.top+iH)/2).toFixed(0)}" text-anchor="middle" fill="#c0392b" font-size="12" font-weight="bold" font-style="italic" font-family="Arial">KRITIEK</text>
        <line x1="${midX.toFixed(1)}" y1="${pad.top}" x2="${midX.toFixed(1)}" y2="${pad.top+iH}" stroke="#aaa" stroke-width="1.5"/>
        <line x1="${pad.left}" y1="${midY.toFixed(1)}" x2="${pad.left+iW}" y2="${midY.toFixed(1)}" stroke="#aaa" stroke-width="1.5"/>
        <rect x="${pad.left}" y="${pad.top}" width="${iW}" height="${iH}" fill="none" stroke="#ccc" stroke-width="1"/>
        ${[1,5,10,15,20,25].map(v=>`<text x="${toX(v).toFixed(1)}" y="${pad.top+iH+14}" text-anchor="middle" fill="#888" font-size="8" font-family="Arial">${v}</text>`).join("")}
        ${[1,2,3,4,5].map(v=>`<text x="${pad.left-6}" y="${(toY(v)+3).toFixed(1)}" text-anchor="end" fill="#888" font-size="8" font-family="Arial">${v}</text>`).join("")}
        <text x="${(pad.left+iW/2).toFixed(0)}" y="${H-4}" text-anchor="middle" fill="#444" font-size="10" font-weight="bold" font-family="Arial">Risico-exposure x Strategisch belang</text>
        <text x="12" y="${(pad.top+iH/2).toFixed(0)}" text-anchor="middle" fill="#444" font-size="10" font-weight="bold" font-family="Arial" transform="rotate(-90,12,${(pad.top+iH/2).toFixed(0)})">Mitigatie</text>
        ${dots}
      </svg>`;
    }

    // ── Dimensietabel ─────────────────────────────────────────
    function generateDimTable(appsArr) {
      const dimLetters=[...new Set(DAAF.map(d=>d.dim))];
      const dimName=l=>{const f=DAAF.find(d=>d.dim===l);return f?f.dimName:l;};
      const lvlColor=l=>["A","B"].includes(l)?"#dc2626":["C","D","E"].includes(l)?"#166534":"#92400e";
      const lvlLabel=l=>["A","B"].includes(l)?"Risico ↓":["C","D","E"].includes(l)?"Mitigatie ↑":"Belang ↓";
      const headers=appsArr.map(a=>`<th>${dName(a).substring(0,16)}</th>`).join("");
      const dimRows=dimLetters.map(l=>{
        const cells=appsArr.map(a=>{
          const v=pdfDimScore(a,l);
          return `<td style="text-align:center;font-weight:700;color:${lvlColor(l)}">${v!=null?v.toFixed(2):"–"}</td>`;
        }).join("");
        return `<tr><td><strong>${l}</strong></td><td>${dimName(l)}</td><td style="color:${lvlColor(l)};font-size:9px">${lvlLabel(l)}</td>${cells}</tr>`;
      }).join("");
      return `<table class="dim-table"><tr><th>Dim</th><th>Naam</th><th>Richting</th>${headers}</tr>${dimRows}</table>`;
    }

    // ── Algemeen risico-conclusie ─────────────────────────────
    function generateRisicoConclusion(appsArr) {
      const scored = appsArr.map(a => ({ ...a, sc: calcScores(a.scores) })).filter(a => a.sc.autonomyScore);
      if (scored.length < 1) return "";

      // Patronen analyseren
      const n = scored.length;
      const kritiek     = scored.filter(a => (a.sc.risico||0) > 3.5);
      const nonEU       = appsArr.filter(a => (a.scores["A1"]||0) >= 4);
      const laagContr   = appsArr.filter(a => (a.scores["E1"]||0) <= 2);
      const laagKennis  = appsArr.filter(a => (a.scores["D1"]||0) <= 2);
      const geenAlt     = appsArr.filter(a => (a.scores["C1"]||0) <= 2);
      const hoogBelang  = scored.filter(a => (a.sc.belang||0) > 3.5);
      const laagDictu   = scored.filter(a => a.sc.dictuAvg && a.sc.dictuAvg < 3);
      const critApp     = scored.filter(a => a.sc.autonomyScore < 3);
      const avgAuto     = scored.reduce((s,a) => s+(a.sc.autonomyScore||0), 0) / n;
      const avgRisico   = scored.reduce((s,a) => s+(a.sc.risico||0), 0) / n;
      const avgMit      = scored.reduce((s,a) => s+(a.sc.mitigatie||0), 0) / n;

      let html = `<p><strong>Algemeen beeld:</strong> Uit de beoordeling van ${n} applicatie${n!==1?"s":""} komt een gemiddelde autonomiescore 
        van <strong>${avgAuto.toFixed(1)}/10</strong> naar voren, met een gemiddelde risico-exposure van ${avgRisico.toFixed(2)}/5 
        en een gemiddelde mitigatie-capaciteit van ${avgMit.toFixed(2)}/5.`;

      if (critApp.length > 0) {
        html += ` <strong style="color:#b91c1c">${critApp.length} applicatie${critApp.length!==1?"s vereisen":"vereist"} directe 
        directe aandacht</strong> vanwege een kritieke autonomiescore: ${critApp.map(a=>dName(a)).join(", ")}.`;
      }
      html += `</p>`;

      // Strategische risico's benoemen
      html += `<p style="margin-top:8px"><strong>Geopolitieke en juridische risico's:</strong> `;
      if (nonEU.length > 0) {
        html += `Bij ${nonEU.length} van de ${n} applicaties (${nonEU.map(a=>dName(a)).join(", ")}) is de leverancier 
          gevestigd buiten de EU of valt de leverancier onder wetgeving zoals de CLOUD Act of FISA 702 (VS) of vergelijkbare 
          wetgeving in andere jurisdicties. Dit betekent dat een buitenlandse overheid in theorie toegang kan vorderen tot 
          data die NHL Stenden verwerkt via deze applicaties, ook als de data fysiek in Europa staat. `;
      } else {
        html += `De leveranciers in deze selectie zijn allen gevestigd binnen de EU of vallen onder adequaatheidsbesluiten. 
          Het geopolitieke risico is daarmee beheersbaar, mits contractuele bescherming op orde is. `;
      }
      html += `</p>`;

      // Vendor lock-in
      html += `<p style="margin-top:8px"><strong>Vendor lock-in en exitrisico:</strong> `;
      if (geenAlt.length > 0 || laagKennis.length > 0 || laagContr.length > 0) {
        const risks = [];
        if (geenAlt.length > 0)    risks.push(`${geenAlt.length} applicatie${geenAlt.length!==1?"s hebben":"heeft"} geen of nauwelijks reëele alternatieven (${geenAlt.map(a=>dName(a)).join(", ")})`);
        if (laagKennis.length > 0) risks.push(`voor ${laagKennis.length} applicatie${laagKennis.length!==1?"s is de":"is de"} interne kennis onvoldoende geborgd (${laagKennis.map(a=>dName(a)).join(", ")})`);
        if (laagContr.length > 0)  risks.push(`${laagContr.length} applicatie${laagContr.length!==1?"s missen":"mist"} adequate exit-clausules in het contract (${laagContr.map(a=>dName(a)).join(", ")})`);
        html += `Er is sprake van significante lock-in risico's: ${risks.join("; ")}. Dit maakt een ongewenste situatie moeilijk omkeerbaar. `;
      } else {
        html += `De mitigatie-capaciteit voor vendor lock-in is over het algemeen op orde: er zijn alternatieven beschikbaar, de interne kennis is geborgd en contracten bevatten exitbepalingen. `;
      }
      html += `</p>`;

      // Data en DICTU
      if (laagDictu.length > 0) {
        html += `<p style="margin-top:8px"><strong>Technische soevereiniteit (DICTU):</strong> 
          ${laagDictu.length} applicatie${laagDictu.length!==1?"s scoren":"scoort"} laag op de DICTU-soevereiniteitsmaatstaf 
          (${laagDictu.map(a=>dName(a)+" "+a.sc.dictuAvg?.toFixed(1)+"/5").join(", ")}). 
          Dit duidt op onvoldoende waarborgen voor dataresidency, sleutelbeheer of juridische bescherming. 
          Technische soevereiniteit is een noodzakelijke randvoorwaarde: juridische bescherming alleen is onvoldoende 
          als de technische infrastructuur toegang voor derden niet uitsluit.</p>`;
      }

      // Strategisch belang
      if (hoogBelang.length > 0) {
        html += `<p style="margin-top:8px"><strong>Strategisch belang en continuïteit:</strong> 
          ${hoogBelang.length} applicatie${hoogBelang.length!==1?"s zijn":"is"} van hoog strategisch belang voor 
          de organisatie (${hoogBelang.map(a=>dName(a)).join(", ")}). 
          Uitval of ongewenste toegang bij deze applicaties raakt direct aan de continuïteit van onderwijs, 
          onderzoek of bedrijfsvoering van NHL Stenden. De afhankelijkheid van externe partijen bij deze 
          applicaties vraagt om de sterkste contractuele en technische waarborgen.</p>`;
      }

      // Aanbeveling
      html += `<p style="margin-top:10px;padding:10px 14px;background:#fff8f0;border-left:4px solid #f59e0b;border-radius:0 4px 4px 0">
        <strong>Aanbeveling:</strong> `;
      if (critApp.length > 0) {
        html += `Stel voor ${critApp.map(a=>dName(a)).join(" en ")} op korte termijn een actieplan op met concrete 
          maatregelen, een verantwoordelijke en een deadline. `;
      }
      if (nonEU.length > 0 && laagContr.length > 0) {
        html += `Versterk bij niet-EU leveranciers de contractuele bescherming: zorg voor juridisch afdwingbare 
          exit-clausules, transitieregelingen en meldplichten bij datavorderingen. `;
      }
      if (laagKennis.length > 0) {
        html += `Start een kennisborgingsprogramma voor applicaties waar de afhankelijkheid van individuele personen 
          hoog is — dit is een quick win met direct effect op de weerbaarheid. `;
      }
      html += `Bespreek de uitkomsten van dit assessment in het Transitieteam Digitalisering en leg de 
        prioritering vast.</p>`;

      return html;
    }
    const rows = visible.map(a => {
      const s=calcScores(a.scores), lbl=scoreLabel(s.autonomyScore);
      return `<tr>
        <td><strong>${displayName(a)}</strong>${a.supplier?`<br/><span class="sub">${a.supplier}</span>`:""}</td>
        <td style="color:${lbl.fg};font-weight:700">${s.autonomyScore?s.autonomyScore.toFixed(1):"–"}</td>
        <td style="color:#dc2626">${s.risico?s.risico.toFixed(2):"–"}</td>
        <td style="color:#26B5AE">${s.mitigatie?s.mitigatie.toFixed(2):"–"}</td>
        <td style="color:#E87722">${s.belang?s.belang.toFixed(2):"–"}</td>
        <td>${s.dictuAvg?s.dictuAvg.toFixed(1)+"/5":"–"}</td>
        <td>${s.completeness}%</td>
        <td style="color:${lbl.fg};font-weight:600">${lbl.text}</td>
      </tr>`;
    }).join("");

    const kwRows = visible.map(a => {
      const s=calcScores(a.scores), rec=generateRecommendations(a.scores);
      const aN = dName(a); // actieve naam (primair of secundair)
      const daafRows=DAAF.map(q=>{
        const motivatieRaw = (a.notes||{})[q.key] || "";
        const motivatieFull = motivatieRaw ? motivatieRaw.replace(new RegExp(a.name.replace(/[.*+?^${}()|[\]\\]/g,"\\$&"),"gi"), aN) : "";
        // Samenvatting: max 300 tekens, afkappen op laatste woord
        const motivatie = motivatieFull.length > 300
          ? motivatieFull.substring(0, motivatieFull.lastIndexOf(" ", 300)) + "…"
          : motivatieFull;
        const heeftSamenvatting = motivatieFull.length > 300;
        return `<tr>
          <td><strong>${q.key}</strong></td><td>${q.dimName}</td><td>${q.name}</td>
          <td style="text-align:center;font-weight:700">${a.scores[q.key]||"–"}</td>
          <td>${a.scores[q.key]?q.scores.find(sc=>sc.s===a.scores[q.key])?.label||"":""}</td>
          <td style="color:${motivatie?"#374151":"#9ca3af"};font-style:${motivatie?"normal":"italic"}">${motivatie||"–"}${heeftSamenvatting?` <span style="font-size:9px;color:#9ca3af;font-style:italic">(zie tool voor uitgebreide toelichting)</span>`:""}</td>
        </tr>`;
      }).join("");
      const dictuRows=DICTU.map(q=>{
        const motivatieRaw = (a.notes||{})[q.key] || "";
        const motivatieFull = motivatieRaw ? motivatieRaw.replace(new RegExp(a.name.replace(/[.*+?^${}()|[\]\\]/g,"\\$&"),"gi"), aN) : "";
        const motivatie = motivatieFull.length > 300
          ? motivatieFull.substring(0, motivatieFull.lastIndexOf(" ", 300)) + "…"
          : motivatieFull;
        const heeftSamenvatting2 = motivatieFull.length > 300;
        return `<tr>
          <td><strong>${q.key}</strong></td><td>${q.cat}</td><td>${q.name}</td>
          <td style="text-align:center;font-weight:700">${a.scores[q.key]||"–"}</td>
          <td>${a.scores[q.key]?q.scores.find(sc=>sc.s===a.scores[q.key])?.label||"":""}</td>
          <td style="color:${motivatie?"#374151":"#9ca3af"};font-style:${motivatie?"normal":"italic"}">${motivatie||"–"}${heeftSamenvatting2?` <span style="font-size:9px;color:#9ca3af;font-style:italic">(zie tool voor uitgebreide toelichting)</span>`:""}</td>
        </tr>`;
      }).join("");
      const lbl=scoreLabel(s.autonomyScore);
      const heeftMotivaties = [...DAAF,...DICTU].some(q => (a.notes||{})[q.key]);
      return `<div class="app-section">
        <h3>${aN}${a.supplier?` <span class="sub">— ${a.supplier}</span>`:""} 
          <span style="font-size:11px;font-weight:600;color:${lbl.fg};padding:2px 8px;background:${lbl.bg};border-radius:3px;margin-left:8px">${lbl.text} ${s.autonomyScore?s.autonomyScore.toFixed(1):""}</span>
        </h3>
        ${[...DAAF,...DICTU].some(q=>(a.notes||{})[q.key]&&(a.notes||{})[q.key].length>300)?`<p style="font-size:9px;color:#9ca3af;font-family:Arial;margin-bottom:6px;font-style:italic">Motivaties zijn samengevat. De volledige toelichting is terug te lezen in de Digitale Soevereiniteitsassessment Tool.</p>`:""}
        <table class="scores-table">
          <tr><th>Vraag</th><th>Dimensie</th><th>Indicator</th><th>Score</th><th>Label</th><th>Motivatie</th></tr>
          ${daafRows}${dictuRows}
        </table>
        <div class="rec-box rec-qw">
          <div class="rec-label" style="color:#92400e">⚡ Quick win</div>
          <div class="rec-text">${rec.quickWin.replace(new RegExp(a.name.replace(/[.*+?^${}()|[\]\\]/g,"\\$&"),"gi"), aN)}</div>
        </div>
        <div class="rec-box rec-str">
          <div class="rec-label" style="color:#166534">🎯 Strategische aanbeveling</div>
          <div class="rec-text">${rec.strategic.replace(new RegExp(a.name.replace(/[.*+?^${}()|[\]\\]/g,"\\$&"),"gi"), aN)}</div>
        </div>
      </div>`;
    }).join("");

    // ── Slotpagina: pre-berekend zodat fouten zichtbaar zijn ──
    const sortedForSlot = visible
      .map(a => ({ ...a, sc: calcScores(a.scores), rec: generateRecommendations(a.scores) }))
      .filter(a => a.sc.autonomyScore)
      .sort((a,b) => (a.sc.autonomyScore||10) - (b.sc.autonomyScore||10));

    // Aandachtspunten tabel
    const slotAandacht = sortedForSlot.length === 0
      ? `<tr><td colspan="4" style="color:#9ca3af;font-style:italic">Nog geen volledig ingevulde assessments.</td></tr>`
      : sortedForSlot.map(a => {
          const lbl = scoreLabel(a.sc.autonomyScore);
          const punten = [];
          const dimS = (letter) => {
            if (letter === "A") {
              const p = [["A1",3],["A3",2]].filter(([k])=>a.scores[k]>0);
              if (!p.length) return 0;
              return p.reduce((s,[k,w])=>s+a.scores[k]*w,0) / p.reduce((s,[,w])=>s+w,0);
            }
            const qs = DAAF.filter(d=>d.dim===letter);
            const vs = qs.map(q=>a.scores[q.key]||0).filter(v=>v>0);
            return vs.length ? vs.reduce((s,v)=>s+v,0)/vs.length : 0;
          };
          if ((a.sc.risico||0) > 3.5)   punten.push("Hoog geopolitiek of leveranciersrisico");
          if ((a.sc.mitigatie||0) < 2.5) punten.push("Lage weerbaarheid — weinig alternatieven of zwakke contractbescherming");
          if (dimS("C") < 2)             punten.push("Nauwelijks technische exitopties");
          if (dimS("D") < 2)             punten.push("Interne kennis onvoldoende geborgd");
          if (dimS("E") < 2)             punten.push("Exit-clausules ontbreken of zijn zwak");
          if ((a.sc.dictuAvg||5) < 3)    punten.push("Onvoldoende technische soevereiniteit (DICTU)");
          if ((a.scores["A1"]||0) >= 4)  punten.push("Leverancier valt onder niet-EU jurisdictie");
          const punt = punten.length > 0 ? punten[0] : "Geen urgente aandachtspunten gevonden";
          const extra = punten.length > 1 ? `<br/><span style="color:#9ca3af;font-size:9px">+ ${punten.length-1} overig${punten.length>2?"e punten":" punt"}</span>` : "";
          return `<tr>
            <td><strong>${dName(a)}</strong>${a.supplier?`<br/><span style="color:#9ca3af;font-size:9px">${a.supplier}</span>`:""}</td>
            <td style="font-weight:700;color:${lbl.fg}">${a.sc.autonomyScore?.toFixed(1)}</td>
            <td><span style="background:${lbl.bg};color:${lbl.fg};padding:2px 6px;border-radius:3px;font-size:9px;font-weight:600">${lbl.text}</span></td>
            <td style="color:#374151">${punt}${extra}</td>
          </tr>`;
        }).join("");

    // Quick wins (top 5 laagste score)
    const slotQW = sortedForSlot.slice(0,5).filter(a => a.rec.quickWin).map((a,i) => {
      const qw = a.rec.quickWin.replace(new RegExp(a.name.replace(/[.*+?^${}()|[\]\\]/g,"\\$&"),"gi"), dName(a));
      return `<div style="display:flex;gap:10px;margin-bottom:8px;page-break-inside:avoid">
        <div style="background:#f59e0b;color:white;font-family:Arial;font-weight:700;font-size:10px;padding:4px 8px;border-radius:3px;flex-shrink:0;align-self:flex-start">⚡ ${i+1}</div>
        <div style="background:#fffbeb;border-left:3px solid #f59e0b;padding:8px 12px;border-radius:0 4px 4px 0;flex:1">
          <div style="font-family:Arial;font-size:10px;font-weight:700;color:#0C2340;margin-bottom:2px">${dName(a)}</div>
          <div style="font-family:Arial;font-size:10px;color:#374151;line-height:1.5">${qw}</div>
        </div>
      </div>`;
    }).join("") || `<p style="color:#9ca3af;font-style:italic;font-family:Arial;font-size:10px">Nog geen aanbevelingen beschikbaar.</p>`;

    // Strategisch advies (top 5 laagste score)
    const slotStr = sortedForSlot.slice(0,5).filter(a => a.rec.strategic).map((a,i) => {
      const str = a.rec.strategic.replace(new RegExp(a.name.replace(/[.*+?^${}()|[\]\\]/g,"\\$&"),"gi"), dName(a));
      return `<div style="display:flex;gap:10px;margin-bottom:8px;page-break-inside:avoid">
        <div style="background:#1A56A0;color:white;font-family:Arial;font-weight:700;font-size:10px;padding:4px 8px;border-radius:3px;flex-shrink:0;align-self:flex-start">🎯 ${i+1}</div>
        <div style="background:#f0fdf4;border-left:3px solid #22c55e;padding:8px 12px;border-radius:0 4px 4px 0;flex:1">
          <div style="font-family:Arial;font-size:10px;font-weight:700;color:#0C2340;margin-bottom:2px">${dName(a)}</div>
          <div style="font-family:Arial;font-size:10px;color:#374151;line-height:1.5">${str}</div>
        </div>
      </div>`;
    }).join("");

    const slotHTML = `
  <!-- SLOTPAGINA -->
  <div class="page-break">
    <h2>6. Slotbevindingen</h2>
    <div class="section-intro">
      Deze pagina brengt de belangrijkste bevindingen samen. De aandachtspunten zijn direct afgeleid 
      uit de assessmentscores, gesorteerd van meest urgent naar minst urgent. De quick wins zijn 
      per direct uitvoerbaar. Het strategisch advies richt zich op de middellange termijn (6–18 maanden).
    </div>

    <h3>Belangrijkste aandachtspunten</h3>
    <table style="margin-bottom:20px">
      <tr>
        <th style="width:22%">Applicatie</th>
        <th style="width:10%">Score</th>
        <th style="width:14%">Status</th>
        <th style="width:54%">Aandachtspunt</th>
      </tr>
      ${slotAandacht}
    </table>

    <h3>Quick wins — direct uitvoerbaar</h3>
    <p style="font-family:Arial;font-size:10px;color:#6b7280;margin-bottom:10px">
      Onderstaande acties zijn per direct uitvoerbaar zonder grote organisatorische of financiële investering.
    </p>
    ${slotQW}

    ${slotStr ? `<h3 style="margin-top:20px">Strategisch advies</h3>
    <p style="font-family:Arial;font-size:10px;color:#6b7280;margin-bottom:10px">
      De strategische aanbevelingen zijn gericht op structurele verbetering op de middellange termijn.
    </p>
    ${slotStr}` : ""}

    <div style="margin-top:28px;padding:16px 20px;background:#0C2340;border-radius:4px;page-break-inside:avoid">
      <div style="font-family:Arial;font-size:11px;font-weight:700;color:white;margin-bottom:6px">Vervolgstap</div>
      <div style="font-family:Arial;font-size:10px;color:#7DD3D0;line-height:1.7;margin-bottom:10px">
        Bespreek de uitkomsten van deze analyse in het <strong style="color:white">Transitieteam Digitalisering</strong> 
        en leg de prioritering vast in het portfolioplan. De <strong style="color:white">Ambassadeurs</strong> 
        (J. Haije, E. Rolf en J. Blom) coördineren de vervolgacties in afstemming met de 
        Multidisciplinaire Expertisegroep. De rollen zijn als volgt belegd: 
        <strong style="color:white">E. van Gorkum</strong> is verantwoordelijk voor de tool, 
        het beheer en het onderhoud. <strong style="color:white">J. Blom en E. Rolf</strong> 
        richten zich op de methodiek en de inhoudelijke begeleiding van het assessmentproces. 
        Dit is de taakverdeling voor de huidige projectfase — voorlopig, totdat er een 
        officiële inbedding in de organisatie is gerealiseerd als vervolg op dit traject.
      </div>
      <div style="font-family:Arial;font-size:10px;color:white;font-weight:700;margin-bottom:6px">Voor applicatie-eigenaren:</div>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-bottom:8px">
        <div style="background:rgba(255,255,255,0.08);border-radius:3px;padding:8px 10px;font-family:Arial;font-size:9px;color:#7DD3D0;line-height:1.5">
          <div style="color:white;font-weight:700;margin-bottom:3px">① Stap 1 — Bekijk je applicatie</div>
          Lees de score en de aandachtspunten van jouw applicatie(s) in dit rapport. Identificeer de quick win die direct uitvoerbaar is.
        </div>
        <div style="background:rgba(255,255,255,0.08);border-radius:3px;padding:8px 10px;font-family:Arial;font-size:9px;color:#7DD3D0;line-height:1.5">
          <div style="color:white;font-weight:700;margin-bottom:3px">② Stap 2 — Vorm een mini-team</div>
          Betrek je contract-eigenaar en functioneel beheerder. Samen hebben jullie de kennis en mandaat om aanbevelingen uit te voeren.
        </div>
        <div style="background:rgba(255,255,255,0.08);border-radius:3px;padding:8px 10px;font-family:Arial;font-size:9px;color:#7DD3D0;line-height:1.5">
          <div style="color:white;font-weight:700;margin-bottom:3px">③ Stap 3 — Neem contact op</div>
          Neem contact op met het Expertiseteam. Zij begeleiden jullie mini-project: van analyse naar concreet actieplan met eigenaar, maatregel en deadline.
        </div>
        <div style="background:rgba(255,255,255,0.08);border-radius:3px;padding:8px 10px;font-family:Arial;font-size:9px;color:#7DD3D0;line-height:1.5">
          <div style="color:white;font-weight:700;margin-bottom:3px">④ Stap 4 — Herbeoordeel</div>
          Na uitvoering van de aanbevelingen wordt de applicatie opnieuw beoordeeld. Het Expertiseteam verzorgt de herbeoordeling en rapportage.
        </div>
      </div>
      <div style="font-family:Arial;font-size:9px;color:#26B5AE">
        NHL Stenden · Programma Digitale Samenhang · Ambassadeurslijn Digitale Soevereiniteit · Aansluiting VH en SURF
      </div>
    </div>
  </div>`;

    const naamModus = useSecondaryName && visible.some(a => a.nameSecondary)
      ? `Discreet (geanonimiseerd)`
      : `Helder (echte namen)`;


    // ── Verbeteracties HTML (berekend vóór de PDF template) ────────────────
    function buildVerbeteractiesHTML(a) {
      const sc = calcScores(a.scores || {});
      const lbl = !sc.autonomyScore ? "Onvolledig"
        : sc.autonomyScore >= 7 ? "Goed"
        : sc.autonomyScore >= 5 ? "Acceptabel"
        : sc.autonomyScore >= 3 ? "Zorgwekkend"
        : "Kritiek";
      const kleur = !sc.autonomyScore ? "#9ca3af"
        : sc.autonomyScore >= 7 ? "#15803d"
        : sc.autonomyScore >= 5 ? "#a16207"
        : sc.autonomyScore >= 3 ? "#c2410c"
        : "#b91c1c";
      const scoreBg = !sc.autonomyScore ? "#f3f4f6"
        : sc.autonomyScore >= 7 ? "#dcfce7"
        : sc.autonomyScore >= 5 ? "#fef9c3"
        : sc.autonomyScore >= 3 ? "#ffedd5"
        : "#fee2e2";

      const acties = [];
      if (sc.risico > 3.5)
        acties.push({ niveau:"Risico", prio:"Hoog", dim:"A — Geopolitiek", actie:"Inventariseer Europese alternatieven voor deze leverancier. Vraag offertes op bij minimaal twee EU-gevestigde aanbieders.", tip:"Begin met een marktverkenning van 1-2 dagdelen. SURF publiceert regelmatig overzichten van EU-conforme alternatieven per categorie.", kleur:"#b91c1c", bg:"#fee2e2" });
      if ((a.scores||{})["A3"] >= 4)
        acties.push({ niveau:"Risico", prio:"Hoog", dim:"A3 — Datalocatie", actie:"Verzoek de leverancier schriftelijk te bevestigen in welke regio data wordt opgeslagen (incl. back-ups en metadata). Leg dit vast in het contract.", tip:"Vraag ook naar de locatie van de control plane — het beheerpaneel van de dienst.", kleur:"#b91c1c", bg:"#fee2e2" });
      if ((a.scores||{})["B1"] >= 4)
        acties.push({ niveau:"Risico", prio:"Hoog", dim:"B — Leverancier", actie:"Breng de concentratie in kaart: hoeveel kritieke processen zijn afhankelijk van deze leverancier? Stel een maximum vast.", tip:"Gebruik de tool om vergelijkbare applicaties van dezelfde leverancier bij elkaar te zoeken.", kleur:"#ea580c", bg:"#ffedd5" });
      if ((a.scores||{})["C1"] <= 2)
        acties.push({ niveau:"Mitigatie", prio:"Hoog", dim:"C — Technisch", actie:"Documenteer en test een noodprocedure voor het geval deze applicatie uitvalt.", tip:"Een noodprocedure hoeft niet perfect te zijn. Zelfs een A4 met de stappen voor de eerste 4 uur is al waardevol.", kleur:"#1A56A0", bg:"#EBF3FF" });
      if ((a.scores||{})["D1"] <= 2)
        acties.push({ niveau:"Mitigatie", prio:"Middel", dim:"D — Organisatorisch", actie:"Leg de kennis over configuratie en beheer vast bij minimaal twee medewerkers.", tip:"Plan een kennisoverdracht-sessie van een halve dag. Documenteer in een wiki of SharePoint.", kleur:"#1A56A0", bg:"#EBF3FF" });
      if ((a.scores||{})["E1"] <= 2)
        acties.push({ niveau:"Mitigatie", prio:"Hoog", dim:"E — Contractueel", actie:"Voeg bij de eerstvolgende contractverlenging toe: een exit-clausule, dataportabiliteitsgarantie en opzegtermijn van maximaal 3 maanden.", tip:"Vraag de leverancier ook om een data return plan: wat krijg je terug als je stopt, en in welk formaat?", kleur:"#1A56A0", bg:"#EBF3FF" });
      if (sc.dictuAvg && sc.dictuAvg < 3)
        acties.push({ niveau:"DICTU", prio:"Hoog", dim:"2.1–4.1 Soevereiniteit", actie:"Vraag de leverancier schriftelijk naar datalocatie, garantie geen leverancierstoegang, verzet tegen niet-EU dataverzoeken en locatie control plane.", tip:"Gebruik de DICTU-vragenlijst als template voor het gesprek met de leverancier.", kleur:"#6d28d9", bg:"#faf5ff" });
      if (sc.belang >= 4 && sc.risico >= 3)
        acties.push({ niveau:"Belang", prio:"Middel", dim:"F/G/H — Strategisch", actie:"Formeel vastleggen bij NHL Stenden: is het risico bewust aanvaard? Maak een korte risicoafweging en leg de beslissing vast inclusief een herzieningsdatum.", tip:"Een korte notitie met het besluit, de afweging en een jaarlijkse reviewafspraak is voldoende.", kleur:"#E87722", bg:"#fff8e1" });
      if (acties.length === 0)
        acties.push({ niveau:"Onderhoud", prio:"Laag", dim:"Algemeen", actie:"Alle scores zijn acceptabel of goed. Plan een hercontrole bij de eerstvolgende contractverlenging.", tip:"Stel een terugkerende herinnering in op de einddatum van het contract.", kleur:"#15803d", bg:"#dcfce7" });

      const supplierStr = a.supplier ? " — " + a.supplier : "";
      const naam = dName(a);

      let html = "";
      html += '<div style="background:' + scoreBg + ';border:1px solid ' + kleur + '44;border-radius:4px;padding:12px 16px;margin-bottom:16px;font-family:Arial;font-size:10px;page-break-inside:avoid">';
      html += '<div style="display:flex;align-items:center;gap:10px;margin-bottom:4px">';
      html += '<div style="font-weight:700;color:#0C2340;font-size:12px">' + naam + (a.supplier ? ' <span style="font-weight:400;color:#9ca3af;font-size:10px">— ' + a.supplier + "</span>" : "") + "</div>";
      html += '<span style="background:' + kleur + ';color:white;padding:2px 8px;border-radius:2px;font-size:9px;font-weight:700">' + lbl + "</span>";
      if (sc.autonomyScore) html += '<span style="font-size:10px;font-weight:700;color:' + kleur + '">Score: ' + sc.autonomyScore.toFixed(1) + "/10</span>";
      html += "</div>";
      html += '<div style="display:grid;grid-template-columns:1fr 1fr 1fr 1fr;gap:8px;margin-top:8px">';
      [["Risico", sc.risico, "#dc2626"], ["Mitigatie", sc.mitigatie, "#26B5AE"], ["Belang", sc.belang, "#E87722"], ["DICTU", sc.dictuAvg, "#6d28d9"]].forEach(function(s) {
        html += '<div style="text-align:center;background:white;border-radius:3px;padding:6px 4px;border:1px solid #e5e7eb">';
        html += '<div style="font-size:14px;font-weight:700;color:' + (s[1] ? s[2] : "#d1d5db") + '">' + (s[1] ? s[1].toFixed(1) : "–") + "</div>";
        html += '<div style="font-size:8px;color:#9ca3af">' + s[0] + "</div></div>";
      });
      html += "</div></div>";

      html += '<table style="width:100%;border-collapse:collapse;font-family:Arial;font-size:9.5px;margin-bottom:16px">';
      html += '<tr style="background:#0C2340;color:white"><th style="padding:7px 10px;text-align:left;width:10%">Niveau</th><th style="padding:7px 10px;text-align:left;width:8%">Prioriteit</th><th style="padding:7px 10px;text-align:left;width:15%">Dimensie</th><th style="padding:7px 10px;text-align:left;width:5%">Afgevinkt</th><th style="padding:7px 10px;text-align:left;width:37%">Actie</th><th style="padding:7px 10px;text-align:left;width:25%">Tip</th></tr>';
      acties.forEach(function(ac, i) {
        const prioBg = ac.prio === "Hoog" ? "#fee2e2" : ac.prio === "Middel" ? "#fef9c3" : "#dcfce7";
        const prioFg = ac.prio === "Hoog" ? "#b91c1c" : ac.prio === "Middel" ? "#a16207" : "#15803d";
        html += '<tr style="background:' + (i % 2 === 0 ? "#f8fafc" : "white") + ';border-bottom:1px solid #f1f5f9;page-break-inside:avoid">';
        html += '<td style="padding:8px 10px;vertical-align:top"><span style="background:' + ac.kleur + ';color:white;padding:2px 6px;border-radius:2px;font-size:8px;font-weight:700">' + ac.niveau + "</span></td>";
        html += '<td style="padding:8px 10px;vertical-align:top"><span style="background:' + prioBg + ';color:' + prioFg + ';padding:2px 6px;border-radius:2px;font-size:8px;font-weight:600">' + ac.prio + "</span></td>";
        html += '<td style="padding:8px 10px;vertical-align:top;color:#374151">' + ac.dim + "</td>";
        html += '<td style="padding:8px 10px;vertical-align:top;text-align:center"><div style="width:14px;height:14px;border:1.5px solid #D0E4F7;border-radius:2px;display:inline-block"></div></td>';
        html += '<td style="padding:8px 10px;vertical-align:top;color:#0C2340;font-weight:500;line-height:1.5">' + ac.actie + "</td>";
        html += '<td style="padding:8px 10px;vertical-align:top;color:#6b7280;line-height:1.5;font-style:italic">' + ac.tip + "</td>";
        html += "</tr>";
      });
      html += "</table>";

      html += '<div style="background:#EBF3FF;border:1px solid #D0E4F7;border-radius:4px;padding:12px 16px;font-family:Arial;font-size:9.5px">';
      html += '<div style="font-weight:700;color:#0C2340;margin-bottom:6px">Aanbevolen aanpak</div>';
      html += '<div style="color:#374151;line-height:1.7">Bespreek deze actielijst met de applicatie-eigenaar, contract-eigenaar en functioneel beheerder. Zij beschikken over de operationele kennis die nodig is om de acties te prioriteren en uit te voeren. Leg de gemaakte afspraken vast — ook bewuste keuzes om een actie niet op te pakken zijn waardevolle informatie voor de strategische besluitvorming van NHL Stenden.</div>';
      html += "</div>";
      return html;
    }
    const verbeteractiesHTML = visible.length === 1 ? buildVerbeteractiesHTML(visible[0]) : "";

    const html = `<!DOCTYPE html>
<html lang="nl">
<head>
<meta charset="UTF-8"/>
<title>${visible.length === 1 ? "Assessment " + dName(visible[0]) + " — Digitale Soevereiniteit NHL Stenden " + datum : "Portfolioanalyse Digitale Soevereiniteit — NHL Stenden " + datum}</title>
<style>
  *{box-sizing:border-box;margin:0;padding:0}
  body{font-family:Georgia,'Times New Roman',serif;font-size:11px;color:#1a1a1a;line-height:1.7}

  /* ── Vaste header op elke pagina ── */
  .page-header{background:#0C2340;color:white;padding:11px 24px;display:flex;align-items:center;gap:14px;position:running(pageHeader)}
  .logo{border:2px solid white;padding:5px 9px;font-weight:700;font-size:9px;letter-spacing:1px;line-height:1.3;font-family:Arial;flex-shrink:0}
  .header-title{font-size:12px;font-weight:700;font-family:Arial}
  .header-sub{font-size:9px;color:#7DD3D0;margin-top:1px;font-family:Arial}
  .header-right{margin-left:auto;font-family:Arial;font-size:9px;color:#7DD3D0;text-align:right}

  /* ── Voorblad ── */
  .cover{display:flex;flex-direction:column;page-break-after:always}
  .cover-top{background:#0C2340;padding:40px 48px 32px;flex:0;color:white}
  .cover-teal-bar{height:5px;background:#26B5AE}
  .cover-body{padding:40px 48px 32px;flex:1;display:flex;flex-direction:column;justify-content:space-between;background:white;color:#1a1a1a}
  .cover-title{font-size:28px;font-weight:700;font-family:Arial;color:#0C2340;line-height:1.2;margin-bottom:8px}
  .cover-subtitle{font-size:15px;font-family:Arial;color:#1A56A0;margin-bottom:32px}
  .cover-meta-block{background:#EBF3FF;border-left:4px solid #1A56A0;padding:16px 20px;border-radius:0 4px 4px 0;margin-bottom:24px}
  .cover-meta-row{display:flex;gap:8px;margin-bottom:5px;font-family:Arial;font-size:10px}
  .cover-meta-label{color:#6b7280;width:120px;flex-shrink:0}
  .cover-meta-value{color:#0C2340;font-weight:600}
  .cover-disclaimer{font-size:9px;color:#6b7280;font-family:Arial;line-height:1.5;border-top:1px solid #e5e7eb;padding-top:12px;margin-top:auto}

  /* ── Inhoudsopgave ── */
  .toc-page{page-break-after:always;padding:32px 48px}
  .toc-title{font-size:18px;font-weight:700;font-family:Arial;color:#0C2340;margin-bottom:4px}
  .toc-bar{height:3px;background:#26B5AE;width:48px;margin-bottom:24px}
  .toc-section{margin-bottom:6px;display:flex;align-items:baseline;gap:6px;font-family:Arial}
  .toc-nr{font-size:10px;color:#1A56A0;font-weight:700;width:24px;flex-shrink:0}
  .toc-lbl{font-size:11px;color:#0C2340;font-weight:700;flex:1}
  .toc-sub{margin-bottom:3px;display:flex;align-items:baseline;gap:6px;padding-left:24px;font-family:Arial}
  .toc-sub .toc-nr{font-size:9px;color:#6b7280;font-weight:400}
  .toc-sub .toc-lbl{font-size:10px;color:#374151;font-weight:400}
  .toc-dots{flex:1;border-bottom:1px dotted #d1d5db;margin:0 6px;position:relative;top:-2px}
  .toc-pg{font-size:10px;color:#1A56A0;font-weight:600;width:20px;text-align:right;flex-shrink:0}

  /* ── Pagina inhoud ── */
  .content{padding:16px 48px 24px}
  .page-break{page-break-before:always}
  .meta{color:#6b7280;font-size:9px;margin-bottom:16px;font-family:Arial}

  /* ── Typografie ── */
  h2{font-family:Arial;font-size:15px;color:#0C2340;padding-bottom:5px;margin:8px 0 6px;font-weight:700;
     border-bottom:2px solid #1A56A0}
  h3{font-family:Arial;font-size:12px;color:#0C2340;margin:16px 0 5px;font-weight:700;page-break-after:avoid}
  .section-intro{font-size:10.5px;line-height:1.7;color:#374151;margin-bottom:14px;
     border-left:3px solid #D0E4F7;padding-left:12px;font-family:Arial}
  .narrative{font-size:11px;line-height:1.75;color:#1a1a1a}
  .narrative p{margin-bottom:10px}
  .narrative strong{font-weight:700;color:#0C2340}
  .narrative em{font-style:italic}

  /* ── Intro-blokken ── */
  .intro-kader{background:#EBF3FF;border-left:4px solid #1A56A0;padding:12px 16px;border-radius:0 4px 4px 0;margin:14px 0;font-family:Arial;font-size:10.5px;color:#0C2340;line-height:1.6}
  .intro-kader strong{color:#0C2340}
  .framework-grid{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin:14px 0}
  .framework-card{padding:12px 14px;border-radius:4px;font-family:Arial;font-size:10px;line-height:1.55}
  .fw-daaf{background:#f0f4ff;border:1px solid #c7d7ff}
  .fw-dictu{background:#E6F7F7;border:1px solid #26B5AE44}
  .fw-title{font-size:11px;font-weight:700;color:#0C2340;margin-bottom:4px}
  .fw-sub{font-size:9px;color:#6b7280;margin-bottom:6px}
  .fw-body{color:#374151}

  /* ── Tabellen ── */
  table{width:100%;border-collapse:collapse;margin-bottom:14px;font-size:10px;font-family:Arial;page-break-inside:avoid}
  th{background:#0C2340;color:white;padding:5px 8px;text-align:left;font-size:10px}
  td{padding:4px 8px;border-bottom:1px solid #e5e7eb;vertical-align:top;font-family:Arial}
  tr:nth-child(even) td{background:#f8fafc}
  .scores-table th{background:#1A56A0}
  .dim-table th{background:#065f46}

  /* ── Grafieken + aanbevelingen ── */
  .chart-wrap{margin:10px 0 18px;page-break-inside:avoid;text-align:center}
  .app-section{margin-bottom:28px;page-break-inside:avoid;border-bottom:1px solid #e5e7eb;padding-bottom:20px}
  .app-section:last-child{border-bottom:none}
  .rec-box{border-radius:3px;padding:9px 13px;margin-bottom:8px;page-break-inside:avoid}
  .rec-qw{background:#fffbeb;border-left:3px solid #f59e0b}
  .rec-str{background:#f0fdf4;border-left:3px solid #22c55e}
  .rec-label{font-weight:700;font-size:10px;margin-bottom:3px;font-family:Arial}
  .rec-text{font-size:10px;line-height:1.55;color:#374151;font-family:Arial}
  .sub{font-weight:normal;color:#6b7280;font-family:Arial}

  /* ── Footer ── */
  .doc-footer{margin-top:32px;padding-top:12px;border-top:1px solid #e5e7eb;
    font-size:9px;color:#9ca3af;text-align:center;font-family:Arial}

  @page{size:A4;margin:12mm 15mm 12mm 15mm}
  @media print{
    body{-webkit-print-color-adjust:exact;print-color-adjust:exact}
    .page-break{page-break-before:always}
    .app-section{page-break-inside:avoid}
    table{page-break-inside:avoid}
    .chart-wrap{page-break-inside:avoid}
    .framework-grid{page-break-inside:avoid}
  }
</style>
</head>
<body>

<!-- ════════════════════════════════════════════════
     VOORBLAD
════════════════════════════════════════════════ -->
<div class="cover">
  <div class="cover-top">
    <div style="display:flex;align-items:center;gap:16px;margin-bottom:28px">
      <div style="border:2.5px solid white;padding:8px 12px;display:inline-block;line-height:1.2">
        <div style="font-family:Arial;font-size:13px;font-weight:900;color:white;letter-spacing:1px">NHL</div>
        <div style="font-family:Arial;font-size:13px;font-weight:900;color:white;letter-spacing:1px">STENDEN</div>
      </div>
      <div style="width:2px;background:#26B5AE;height:44px"></div>
      <div>
        <div style="font-family:Arial;font-size:12px;font-weight:700;color:white">NHL Stenden Hogeschool</div>
        <div style="font-family:Arial;font-size:9px;color:#7DD3D0;margin-top:2px">Programma Digitale Samenhang · Ambassadeurslijn Digitale Soevereiniteit</div>
      </div>
    </div>
    <div style="height:2px;background:rgba(255,255,255,0.15);margin-bottom:28px"></div>
    <div style="font-family:Arial;font-size:11px;color:#7DD3D0;margin-bottom:8px;letter-spacing:0.05em">AUTOMATISCH GEGENEREERDE ANALYSE</div>
    ${visible.length === 1 ? `
    <div style="font-family:Arial;font-size:13px;font-weight:400;color:#7DD3D0;margin-bottom:8px;letter-spacing:0.02em">Assessment Digitale Soevereiniteit</div>
    <div style="font-family:Arial;font-size:34px;font-weight:700;color:white;line-height:1.1;margin-bottom:6px">` + dName(visible[0]) + `</div>
    <div style="font-family:Arial;font-size:13px;color:#7DD3D0;margin-bottom:4px">` + (visible[0].supplier ? visible[0].supplier + " · " : "") + `NHL Stenden Hogeschool · ` + datum + `</div>
    ` : `
    <div style="font-family:Arial;font-size:30px;font-weight:700;color:white;line-height:1.15;margin-bottom:8px">Portfolioanalyse<br/>Digitale Soevereiniteit</div>
    <div style="font-family:Arial;font-size:13px;color:#7DD3D0">Applicatielandschap NHL Stenden · ` + datum + `</div>
    `}
  </div>
  <div class="cover-teal-bar"></div>
  <div class="cover-body">
    <div>
      <div class="cover-meta-block">
        <div class="cover-meta-row"><span class="cover-meta-label">Datum rapport</span><span class="cover-meta-value">${datum}</span></div>
        <div class="cover-meta-row"><span class="cover-meta-label">Versie tool</span><span class="cover-meta-value">${VERSION}</span></div>
        <div class="cover-meta-row"><span class="cover-meta-label">Applicaties</span><span class="cover-meta-value">${visible.length} applicatie${visible.length!==1?"s":""} in selectie</span></div>
        <div class="cover-meta-row"><span class="cover-meta-label">Naamweergave</span><span class="cover-meta-value">${naamModus}</span></div>
        <div class="cover-meta-row"><span class="cover-meta-label">Kwartiermaker</span><span class="cover-meta-value">E. van Gorkum</span></div>
        <div class="cover-meta-row"><span class="cover-meta-label">Ambassadeurs</span><span class="cover-meta-value">J. Haije · E. Rolf · J. Blom</span></div>
      </div>
      <div style="font-family:Arial;font-size:10px;color:#374151;line-height:1.7;margin-bottom:12px">
        Deze analyse is automatisch gegenereerd door de <strong style="color:#0C2340">Digitale Soevereiniteitsassessment Tool</strong> 
        van NHL Stenden Hogeschool. De tool is ontwikkeld door <strong>kwartiermaker E. van Gorkum</strong> samen met de <strong>Ambassadeurs J. Haije, E. Rolf en J. Blom</strong>, 
        als antwoord op de centrale vraagstelling voor NHL Stenden: <em>"Waar zetten we onze data neer en waar 
        liggen de potentiële problemen?"</em>
      </div>
      <div style="font-family:Arial;font-size:10px;color:#374151;line-height:1.7;margin-bottom:16px">
        De analyse combineert twee erkende normenkaders (DAAF en DICTU) en genereert automatisch 
        scores, aanbevelingen en prioritering op basis van de ingevoerde assessmentdata. 
        De uitkomsten zijn daarmee direct bruikbaar voor gesprekken met applicatie-eigenaren, 
        leveranciers en het management.
      </div>
    </div>
    <div class="cover-disclaimer">
      Dit document is vertrouwelijk en bestemd voor intern gebruik binnen NHL Stenden Hogeschool. 
      Scores en aanbevelingen zijn gebaseerd op de op het moment van assessment beschikbare informatie. 
      NHL Stenden Hogeschool · Programma Digitale Samenhang · ${VERSION}
    </div>
  </div>
</div>

<!-- ════════════════════════════════════════════════
     INHOUDSOPGAVE
════════════════════════════════════════════════ -->
<div class="page-header">
  <div class="logo">NHL<br/>STENDEN</div>
  <div style="width:2px;background:#26B5AE;align-self:stretch"></div>
  <div><div class="header-title">Portfolioanalyse Digitale Soevereiniteit</div>
  <div class="header-sub">Applicatielandschap NHL Stenden · ${datum} · ${VERSION}</div></div>
  <div class="header-right">${naamModus}</div>
</div>
<div class="toc-page">
  <div class="toc-title">Inhoudsopgave</div>
  <div class="toc-bar"></div>

  <div class="toc-section"><span class="toc-nr">1.</span><span class="toc-lbl">Inleiding en kader</span><span class="toc-dots"></span><span class="toc-pg">3</span></div>
  <div class="toc-sub"><span class="toc-nr">1.1</span><span class="toc-lbl">Over deze analyse en de tool</span><span class="toc-dots"></span><span class="toc-pg">3</span></div>
  <div class="toc-sub"><span class="toc-nr">1.2</span><span class="toc-lbl">Het Expertiseteam Digitale Soevereiniteit</span><span class="toc-dots"></span><span class="toc-pg">3</span></div>
  <div class="toc-sub"><span class="toc-nr">1.3</span><span class="toc-lbl">Organisatorische context</span><span class="toc-dots"></span><span class="toc-pg">3</span></div>
  <div class="toc-sub"><span class="toc-nr">1.4</span><span class="toc-lbl">Toegepaste frameworks</span><span class="toc-dots"></span><span class="toc-pg">3</span></div>
  <div class="toc-sub"><span class="toc-nr">1.5</span><span class="toc-lbl">Gebruik in het hoger onderwijs</span><span class="toc-dots"></span><span class="toc-pg">3</span></div>

  <div class="toc-section" style="margin-top:8px"><span class="toc-nr">2.</span><span class="toc-lbl">Samenvatting</span><span class="toc-dots"></span><span class="toc-pg">4</span></div>

  <div class="toc-section" style="margin-top:8px"><span class="toc-nr">3.</span><span class="toc-lbl">Risico-analyse en aanbevelingen</span><span class="toc-dots"></span><span class="toc-pg">4</span></div>

  <div class="toc-section" style="margin-top:8px"><span class="toc-nr">4.</span><span class="toc-lbl">Scoreoverzicht — alle applicaties</span><span class="toc-dots"></span><span class="toc-pg">4</span></div>

  <div class="toc-section" style="margin-top:8px"><span class="toc-nr">5.</span><span class="toc-lbl">Visuele analyse</span><span class="toc-dots"></span><span class="toc-pg">5</span></div>
  <div class="toc-sub"><span class="toc-nr">5.1</span><span class="toc-lbl">Autonomie-kwadrant (DAAF)</span><span class="toc-dots"></span><span class="toc-pg">5</span></div>
  <div class="toc-sub"><span class="toc-nr">5.2</span><span class="toc-lbl">Dimensieprofiel per applicatie</span><span class="toc-dots"></span><span class="toc-pg">5</span></div>
  <div class="toc-sub"><span class="toc-nr">5.3</span><span class="toc-lbl">Dimensiescores tabel</span><span class="toc-dots"></span><span class="toc-pg">5</span></div>

  <div class="toc-section" style="margin-top:8px"><span class="toc-nr">6.</span><span class="toc-lbl">Slotbevindingen — aandachtspunten, quick wins en advies</span><span class="toc-dots"></span><span class="toc-pg">6+</span></div>
  ${visible.length === 1 ? `<div class="toc-section" style="margin-top:8px"><span class="toc-nr">7.</span><span class="toc-lbl">Verbeteracties — ${dName(visible[0])}</span><span class="toc-dots"></span><span class="toc-pg">7+</span></div>` : `<div class="toc-section" style="margin-top:8px"><span class="toc-nr">7.</span><span class="toc-lbl">Vervolgstappen — Review door applicatie-eigenaren</span><span class="toc-dots"></span><span class="toc-pg">7+</span></div>`}
  <div class="toc-section" style="margin-top:8px"><span class="toc-nr">8.</span><span class="toc-lbl">Bijlage — Detailscores per applicatie</span><span class="toc-dots"></span><span class="toc-pg">8+</span></div>
  ${visible.map((a,i) => `<div class="toc-sub"><span class="toc-nr">${i+1}.</span><span class="toc-lbl">${dName(a)}${a.supplier?` <span style="color:#9ca3af;font-weight:400">— ${a.supplier}</span>`:""}</span><span class="toc-dots"></span><span class="toc-pg">${6+i}</span></div>`).join("")}

</div>

<!-- ════════════════════════════════════════════════
     PAGINA 1: INLEIDING
════════════════════════════════════════════════ -->
<div class="page-header">
  <div class="logo">NHL<br/>STENDEN</div>
  <div style="width:2px;background:#26B5AE;align-self:stretch"></div>
  <div><div class="header-title">Portfolioanalyse Digitale Soevereiniteit</div>
  <div class="header-sub">Applicatielandschap NHL Stenden · ${datum} · ${VERSION}</div></div>
  <div class="header-right">${naamModus}</div>
</div>
<div style="padding:16px 48px 24px">

  <h2>1. Inleiding en kader</h2>
  <div class="section-intro">
    Dit hoofdstuk beschrijft de aanleiding voor het rapport, de toegepaste frameworks en de bredere context 
    van digitale soevereiniteit in het Nederlandse hoger onderwijs.
  </div>

  <h3>1.1 Over deze analyse en de tool</h3>
  <div class="narrative">
    <p>Deze portfolioanalyse is automatisch gegenereerd door de <strong>Digitale Soevereiniteitsassessment Tool</strong> 
    van NHL Stenden Hogeschool (${VERSION}). De tool is een levend, realtime instrument — 
    géén statisch rapport. Assessments worden bijgehouden en bijgewerkt naarmate contracten wijzigen, 
    leveranciers hun beleid aanpassen of nieuwe inzichten beschikbaar komen. Dit rapport is een momentopname; 
    de tool zelf biedt altijd de meest actuele stand.</p>
    <p>De analyse geeft antwoord op de centrale vraagstelling voor NHL Stenden: 
    <em>"Waar zetten we onze data neer en waar liggen de potentiële problemen?"</em> 
    Voor deze analyse zijn <strong>${visible.length} applicatie${visible.length!==1?"s":""}</strong> beoordeeld 
    uit het kern-applicatielandschap. De tool is opgezet rondom <strong>23 kernsystemen</strong>, 
    maar is nadrukkelijk open voor elk informatiesysteem dat NHL Stenden inzet. 
    Applicatie-eigenaren en beheerders kunnen de tool <strong>proactief inzetten</strong> 
    om op elk gewenst moment de soevereiniteit van een applicatie te controleren en te documenteren.</p>
  </div>

  <h3>1.2 De Kwartiermaker, Ambassadeurs en Multidisciplinaire Expertisegroep</h3>
  <div class="narrative">
    <p>Het traject Digitale Soevereiniteit wordt uitgevoerd door een combinatie van een 
    <strong>kwartiermaker</strong>, <strong>drie ambassadeurs</strong> en een 
    <strong>multidisciplinaire expertisegroep</strong>. Elk speelt een eigen en duidelijk onderscheiden rol.</p>
  </div>
  <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px;margin:10px 0 12px;page-break-inside:avoid;font-family:Arial;font-size:10px">
    <div style="background:#EBF3FF;border-left:3px solid #1A56A0;padding:11px 13px;border-radius:0 4px 4px 0">
      <div style="font-weight:700;color:#0C2340;margin-bottom:2px">Kwartiermaker</div>
      <div style="color:#1A56A0;font-size:9px;margin-bottom:6px">E. van Gorkum · Programma Digitale Samenhang</div>
      <div style="color:#374151;line-height:1.55;font-size:9px">
        Verantwoordelijk voor de assessmenttool, de methodiek en de begeleiding van het assessmentproces 
        op instellingsniveau. Stuurt de ambassadeurslijn aan en bewaakt de samenhang met het 
        bredere Programma Digitale Samenhang.
      </div>
    </div>
    <div style="background:#f0f9f9;border-left:3px solid #26B5AE;padding:11px 13px;border-radius:0 4px 4px 0">
      <div style="font-weight:700;color:#0C2340;margin-bottom:2px">De Ambassadeurs</div>
      <div style="color:#26B5AE;font-size:9px;margin-bottom:6px">Jooske Haije · Esther Rolf · Jorn Blom</div>
      <div style="color:#374151;line-height:1.55;font-size:9px">
        Verbinden initiatieven, bieden een ordenend kader en zijn de schakel tussen de ambassadeurslijn 
        en de projectorganisatie. Focus op rust, overzicht en gerichte besluitvorming. 
        Maandelijks overleg met de expertisegroep borgt afstemming en voortgang.
      </div>
    </div>
    <div style="background:#fff8f0;border-left:3px solid #E87722;padding:11px 13px;border-radius:0 4px 4px 0">
      <div style="font-weight:700;color:#0C2340;margin-bottom:2px">Multidisciplinaire Expertisegroep</div>
      <div style="color:#E87722;font-size:9px;margin-bottom:6px">Voortgekomen uit het kernteam · 3 domeinen</div>
      <div style="color:#374151;line-height:1.55;font-size:9px">
        <strong>Beleid &amp; Juridisch</strong> (strategie, compliance, juridische kaders) · 
        <strong>Techniek &amp; Beheer</strong> (CISO, architectuur, informatiemanagement) · 
        <strong>Onderwijs &amp; Onderzoek</strong> (onderwijskundige innovatie, lectoraten, SURF). 
        Leden worden flexibel ingezet per product of vraagstuk.
      </div>
    </div>
  </div>
  <div class="narrative">
    <p>De expertisegroep levert input en feedback op de concrete deliverables, borgt de opgebouwde kennis 
    en waarborgt dat digitale soevereiniteit vanuit alle relevante invalshoeken wordt benaderd. 
    Bij applicaties met aandachtspunten adviseert de expertisegroep nadrukkelijk om ook de 
    <strong>applicatie-eigenaar, contract-eigenaar en functioneel beheerder</strong> te betrekken — 
    zij beschikken over de operationele kennis die scores kan nuanceren en aanbevelingen 
    uitvoerbaar maakt (zie hoofdstuk 7 en 8).</p>
  </div>

  <h3>1.3 Organisatorische context</h3>
  <div class="narrative">
    <p>In december 2025 heeft het kernteam Digitale Soevereiniteit een adviesopdracht aangeboden 
    NHL Stenden heeft digitale soevereiniteit verankerd als ambassadeurslijn binnen het 
    Programma Digitale Samenhang — geen apart project, maar een 
    <strong>ordenend perspectief</strong> dat bestaande initiatieven verbindt en richting geeft.</p>
  </div>
  <div style="background:#EBF3FF;border:1px solid #D0E4F7;border-radius:4px;padding:12px 16px;margin:10px 0 14px;font-family:Arial;font-size:10px;page-break-inside:avoid">
    <div style="font-weight:700;color:#0C2340;margin-bottom:8px">Roadmap 2026 — drie fasen</div>
    <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px">
      <div style="background:white;border-left:3px solid #1A56A0;padding:8px 10px;border-radius:0 3px 3px 0">
        <div style="font-weight:700;color:#1A56A0;font-size:9px">FASE 1 · Feb–Mrt</div>
        <div style="font-weight:600;color:#0C2340;margin:2px 0">Verkenning &amp; Analyse</div>
        <div style="color:#6b7280;font-size:9px">Analyse huidige situatie, benchmark en kernwaarden-afwegingskader</div>
      </div>
      <div style="background:white;border-left:3px solid #26B5AE;padding:8px 10px;border-radius:0 3px 3px 0">
        <div style="font-weight:700;color:#26B5AE;font-size:9px">FASE 2 · Apr–Mei</div>
        <div style="font-weight:600;color:#0C2340;margin:2px 0">Strategische Verkenning</div>
        <div style="color:#6b7280;font-size:9px">Strategische dilemma's en 2–3 scenario's met kosten-batenanalyse</div>
      </div>
      <div style="background:white;border-left:3px solid #E87722;padding:8px 10px;border-radius:0 3px 3px 0">
        <div style="font-weight:700;color:#E87722;font-size:9px">FASE 3 · Juni</div>
        <div style="font-weight:600;color:#0C2340;margin:2px 0">Advies &amp; Besluitvorming</div>
        <div style="color:#6b7280;font-size:9px">Onderbouwd eindadvies en implementatie-roadmap voor NHL Stenden</div>
      </div>
    </div>
  </div>

  <h3>1.4 Toegepaste frameworks</h3>
  <div class="section-intro">
    De beoordeling is gebaseerd op twee complementaire en erkende normenkaders voor digitale soevereiniteit.
  </div>
  <div class="framework-grid">
    <div class="framework-card fw-daaf">
      <div class="fw-title">DAAF — Digital Autonomy Assessment Framework</div>
      <div class="fw-sub">Utrecht University · Open source · github.com/utrechtuniversity</div>
      <div class="fw-body">Het DAAF-framework beoordeelt applicaties op drie niveaus: 
      <strong>Risico-exposure</strong> (geopolitieke en leveranciersrisico's), 
      <strong>Mitigatie-capaciteit</strong> (technische, organisatorische en contractuele weerbaarheid) 
      en <strong>Strategisch belang</strong> (impact bij uitval, datagevoeligheid en academische impact). 
      De uitkomst is een gewogen autonomiescore op een schaal van 1 tot 10. 
      NHL Stenden past de Quick Scan variant toe met 9 kernindicatoren.</div>
    </div>
    <div class="framework-card fw-dictu">
      <div class="fw-title">DICTU Soevereiniteitscheck</div>
      <div class="fw-sub">Rijksoverheid / DICTU · Technische soevereiniteit</div>
      <div class="fw-body">De DICTU Soevereiniteitscheck richt zich specifiek op de technische dimensie 
      van soevereiniteit. Vier vragen beoordelen: <strong>data residency</strong> (opslag uitsluitend 
      in de EU), <strong>technische toegangsbeveiliging</strong> (geen leverancierstoegang zonder 
      toestemming), <strong>juridische bescherming</strong> (leverancier bestrijdt niet-EU 
      datavorderingen) en <strong>EU-infrastructuur</strong> (control plane volledig in de EU). 
      Score loopt van 1 (volledig afhankelijk) tot 5 (maximaal soeverein).</div>
    </div>
  </div>

  <h3>1.5 Gebruik in het hoger onderwijs</h3>
  <div class="narrative">
    <p>Digitale soevereiniteit staat breed op de agenda in het Nederlandse hoger onderwijs. 
    De <strong>Vereniging Hogescholen (VH)</strong> en <strong>SURF</strong> — de ICT-samenwerkingsorganisatie 
    van onderwijs en onderzoek — werken aan gezamenlijke kaders en richtlijnen voor instellingen 
    die hun digitale afhankelijkheden in kaart willen brengen. Daarin zijn vier pijlers leidend: 
    juridische soevereiniteit (onder welk recht valt de leverancier?), technische soevereiniteit 
    (is data-portabiliteit en exitbaarheid geborgd?), organisatorische soevereiniteit (is er interne 
    kennis en zijn er exitplannen?) en geopolitieke soevereiniteit (welke risico's brengt de 
    jurisdictie van de leverancier met zich mee?).</p>
    <p>NHL Stenden loopt voorop door het DAAF-framework van Utrecht University en de DICTU 
    Soevereiniteitscheck te combineren in één geïntegreerde assessmenttool. De uitkomsten sluiten 
    aan bij de digitale strategie van VH en SURF en geven het Programma Digitale Samenhang 
    concrete handvatten voor prioritering, leveranciersgesprekken en beleidsvorming.</p>
  </div>

  <!-- PAGINA 2: Samenvatting + analyse -->
  <div class="page-break">
    <h2>2. Samenvatting</h2>
    <div class="section-intro">
      Een beknopt overzicht van de belangrijkste bevindingen, gevolgd door de opbouw van dit rapport.
    </div>
    <div class="narrative">${generateSummary(visible)}</div>
    <div style="background:#EBF3FF;border:1px solid #D0E4F7;border-radius:4px;padding:12px 16px;margin-top:14px;font-family:Arial;font-size:10px;page-break-inside:avoid">
      <div style="font-weight:700;color:#0C2340;margin-bottom:8px">Opbouw van dit rapport</div>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:6px">
        <div style="display:flex;gap:8px;align-items:flex-start"><span style="background:#1A56A0;color:white;font-weight:700;padding:1px 6px;border-radius:2px;font-size:9px;flex-shrink:0">H3</span><span style="color:#374151">Risico-analyse en aanbevelingen — de grootste risico's en concrete aanbevelingen</span></div>
        <div style="display:flex;gap:8px;align-items:flex-start"><span style="background:#1A56A0;color:white;font-weight:700;padding:1px 6px;border-radius:2px;font-size:9px;flex-shrink:0">H4</span><span style="color:#374151">Scoreoverzicht — kerncijfers van alle applicaties op één pagina</span></div>
        <div style="display:flex;gap:8px;align-items:flex-start"><span style="background:#1A56A0;color:white;font-weight:700;padding:1px 6px;border-radius:2px;font-size:9px;flex-shrink:0">H5</span><span style="color:#374151">Visuele analyse — kwadrant, dimensieprofiel en scoretabel</span></div>
        <div style="display:flex;gap:8px;align-items:flex-start"><span style="background:#26B5AE;color:white;font-weight:700;padding:1px 6px;border-radius:2px;font-size:9px;flex-shrink:0">H6</span><span style="color:#374151">Slotbevindingen — aandachtspunten, quick wins en strategisch advies</span></div>
        <div style="display:flex;gap:8px;align-items:flex-start"><span style="background:#26B5AE;color:white;font-weight:700;padding:1px 6px;border-radius:2px;font-size:9px;flex-shrink:0">H7</span><span style="color:#374151">Vervolgacties Fase 2 — review door applicatie-eigenaren en concrete acties</span></div>
        <div style="display:flex;gap:8px;align-items:flex-start"><span style="background:#6b7280;color:white;font-weight:700;padding:1px 6px;border-radius:2px;font-size:9px;flex-shrink:0">H8</span><span style="color:#6b7280">Bijlage — detailscores per applicatie (voor geïnteresseerden en eigenaren)</span></div>
      </div>
    </div>

    <h2>3. Risico-analyse en aanbevelingen</h2>
    <div class="section-intro">
      Op basis van de assessmentresultaten worden hieronder de belangrijkste risicogebieden 
      geduid en concrete aanbevelingen geformuleerd voor het Transitieteam Digitalisering.
    </div>
    <div class="narrative">${generateRisicoConclusion(visible)}</div>
  </div>

  <!-- Scoreoverzicht -->
  <div class="page-break">
    <h2>4. Scoreoverzicht — alle applicaties</h2>
    <div class="section-intro">
      Onderstaande tabel geeft een overzicht van alle beoordeelde applicaties met de 
      kerncijfers per applicatie. De autonomiescore (1–10) is de centrale uitkomst: 
      een lagere score betekent urgentere aandacht. Risico en Belang: laag is beter. 
      Mitigatie: hoog is beter. DICTU: schaal 1–5, hoger is soevereiner.
    </div>
    <table>
      <tr><th>Applicatie</th><th>Autonomie (1-10)</th><th>Risico ↓</th><th>Mitigatie ↑</th><th>Belang ↓</th><th>DICTU ↑</th><th>Volledigheid</th><th>Status</th></tr>
      ${rows}
    </table>
  </div>

  <!-- PAGINA 3: Visuele analyse -->
  <div class="page-break">
    <h2>5. Visuele analyse</h2>
    <div class="section-intro">
      De onderstaande visualisaties geven op drie manieren inzicht in het applicatieportfolio: 
      het kwadrant toont de strategische positie, het dimensieprofiel toont de sterktes en 
      zwaktes per DAAF-dimensie, en de tabel maakt de scores vergelijkbaar.
    </div>

    <!-- 5.1: koptekst + uitleg vastgeplakt aan grafiek via page-break-after:avoid -->
    <div style="page-break-after:avoid">
      <h3>5.1 Autonomie-kwadrant (DAAF)</h3>
      <p style="font-size:10px;color:#374151;margin-bottom:8px;font-family:Arial;line-height:1.6">
        Dit kwadrant plaatst elke applicatie op twee assen en maakt in één oogopslag zichtbaar 
        welke applicaties de meeste aandacht verdienen.<br/>
        <strong>Horizontale as</strong> — Risico × Belang: hoe verder naar rechts, hoe urgenter 
        (combinatie van geopolitiek risico, leveranciersafhankelijkheid en strategisch belang).<br/>
        <strong>Verticale as</strong> — Mitigatie: hoe hoger, hoe beter de organisatie is beschermd 
        (technische exitopties, interne kennis, contractuele weerbaarheid).<br/>
        <strong>Vier kwadranten:</strong> 
        OPTIMAAL (linksboven: laag risico, goede bescherming) · 
        BEHEERSBAAR (rechtsboven: hoog risico maar goede bescherming) · 
        AANDACHTSPUNT (linksonder: laag risico maar weinig bescherming) · 
        KRITIEK (rechtsonder: hoog risico én weinig bescherming — vraagt directe actie).
      </p>
    </div>
    <div class="chart-wrap">${generateKwadrantSVG(visible)}</div>

    <!-- 5.2: altijd nieuwe pagina zodat kop en grafiek samen staan -->
    <div style="page-break-before:always;padding-top:0">
      <h3>5.2 Dimensieprofiel per applicatie</h3>
      <p style="font-size:10px;color:#374151;margin-bottom:8px;font-family:Arial;line-height:1.6">
        Dit profiel toont per DAAF-dimensie hoe alle applicaties scoren op een schaal van 1 tot 5. 
        Elke gekleurde stip op de balk is één applicatie. De kleurovergang laat zien wat 
        goed of slecht is voor die specifieke dimensie:<br/>
        <strong>Risico-assen (A, B)</strong> — groen aan de linkerkant (lage score = weinig risico = goed). 
        Rood aan de rechterkant betekent hoog risico.<br/>
        <strong>Mitigatie-assen (C, D, E)</strong> — groen aan de rechterkant (hoge score = sterke weerbaarheid = goed). 
        Rood aan de linkerkant betekent weinig bescherming.<br/>
        <strong>Belang-assen (F, G, H)</strong> — hoge score betekent hoog strategisch belang. 
        Op zichzelf niet goed of slecht, maar in combinatie met laag risico geeft het prioriteit.<br/>
        Stippen die ver van het groene uiteinde liggen, zijn aandachtspunten.
      </p>
    </div>
    <div class="chart-wrap">${generateSpinSVG(visible)}</div>

    <!-- 5.3: koptekst + uitleg vastgeplakt aan tabel -->
    <div style="page-break-after:avoid">
      <h3>5.3 Dimensiescores per applicatie</h3>
      <p style="font-size:10px;color:#6b7280;margin-bottom:8px;font-family:Arial">
        Gewogen dimensiescores (1–5) per applicatie. Groen = goed voor dat type as. Rood = aandacht vereist.
      </p>
    </div>
    ${generateDimTable(visible)}
  </div>

  ${slotHTML}

  <!-- HOOFDSTUK 7: conditieel ─────────────────────────────────
       Bij 1 app: verbeteracties-checklist
       Bij meerdere: vervolgstappen eigenaren (zonder kernapplicatietabel) -->
  <div class="page-break">
    ${visible.length === 1 ? `
    <h2>7. Verbeteracties — ${dName(visible[0])}</h2>
    <div class="section-intro">
      Dit hoofdstuk vertaalt de assessmentscores van <strong>${dName(visible[0])}</strong> naar concrete,
      uitvoerbare verbeteracties. De acties zijn ingedeeld per DAAF-niveau en gesorteerd op urgentie.
      Gebruik deze lijst als werkdocument voor gesprekken met de applicatie-eigenaar, contract-eigenaar
      en leverancier.
    </div>
    ${verbeteractiesHTML}
    ` : `
    <h2>7. Vervolgstappen — Review door applicatie-eigenaren</h2>
    <div class="section-intro">
      Dit hoofdstuk beschrijft de aanbevolen vervolgstappen voor de review van assessmentscores
      met applicatie-eigenaren en geeft concrete acties voor het verdere traject.
    </div>

    <h3>7.1 Context van de huidige scores</h3>
    <div class="narrative">
      <p>De scores in dit rapport zijn ingevuld op basis van beschikbare contractgegevens,
      publieke documentatie en interne kennis van het
      <strong>Expertiseteam Digitale Soevereiniteit</strong>.
      Dit is een bewuste methodische keuze: door eerst een nulmeting te doen
      op basis van contractinformatie, ontstaat een objectief startpunt dat
      onafhankelijk is van subjectieve perceptie.</p>
      <p><strong>Applicatie-eigenaren, contract-eigenaren en functioneel beheerders</strong>
      beschikken over praktijkkennis die de scores kan nuanceren, verbeteren of
      corrigeren. Scores kunnen daardoor nog wijzigen — en dat is nadrukkelijk de bedoeling.
      De tool is een levend instrument, geen statisch rapport.</p>
    </div>

    <h3>7.2 Concrete acties</h3>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;margin:10px 0;page-break-inside:avoid">
      <div style="background:#EBF3FF;border-left:3px solid #1A56A0;padding:10px 14px;border-radius:0 4px 4px 0;font-family:Arial;font-size:10px">
        <div style="font-weight:700;color:#1A56A0;margin-bottom:4px">① Uitnodiging applicatie-eigenaren</div>
        <div style="color:#374151;line-height:1.6">Elke contactpersoon ontvangt een uitnodiging van het Expertiseteam. Zij worden gevraagd de scores voor hun applicatie(s) te reviewen en aan te vullen met contractuele en operationele kennis.</div>
      </div>
      <div style="background:#E6F7F7;border-left:3px solid #26B5AE;padding:10px 14px;border-radius:0 4px 4px 0;font-family:Arial;font-size:10px">
        <div style="font-weight:700;color:#26B5AE;margin-bottom:4px">② Begeleide review sessies</div>
        <div style="color:#374151;line-height:1.6">Per applicatie wordt een werksessie gepland met eigenaar, contract-eigenaar en functioneel beheerder. Het Expertiseteam faciliteert en zorgt voor een gestructureerde aanpak.</div>
      </div>
      <div style="background:#fff8f0;border-left:3px solid #E87722;padding:10px 14px;border-radius:0 4px 4px 0;font-family:Arial;font-size:10px">
        <div style="font-weight:700;color:#E87722;margin-bottom:4px">③ Bijwerken in de tool</div>
        <div style="color:#374151;line-height:1.6">Na elke review worden de scores bijgewerkt. Motivatieteksten worden aangevuld of gecorrigeerd. De tool genereert automatisch een bijgewerkte rapportage.</div>
      </div>
      <div style="background:#f0fdf4;border-left:3px solid #22c55e;padding:10px 14px;border-radius:0 4px 4px 0;font-family:Arial;font-size:10px">
        <div style="font-weight:700;color:#166534;margin-bottom:4px">④ Input voor strategie en besluitvorming</div>
        <div style="color:#374151;line-height:1.6">De herziene scores en praktijkinzichten vormen directe input voor strategische dilemma's en scenario's met kosten-batenanalyse, als voorbereiding op de strategische besluitvorming van NHL Stenden.</div>
      </div>
    </div>
    `}
  </div>

    <!-- PAGINA 4+: Detail per applicatie -->
  <div class="page-break">
    <h2>8. Detailscores en aanbevelingen per applicatie — Bijlage</h2>
    <div class="section-intro">
      Dit hoofdstuk bevat de volledige detailscores per applicatie — inclusief alle ingevulde scores, 
      motivaties en automatisch gegenereerde aanbevelingen. Het is bedoeld als <strong>naslagwerk</strong> 
      voor geïnteresseerden en applicatie-eigenaren die de volledige onderbouwing willen inzien. 
      De kern van de analyse en de aanbevelingen staan in de voorgaande hoofdstukken.
    </div>
    ${kwRows}
    <div class="doc-footer">
      NHL Stenden Hogeschool · Programma Digitale Samenhang · Ambassadeurslijn Digitale Soevereiniteit ·
      ${VERSION} · ${datum} · Kwartiermaker: E. van Gorkum · Ambassadeurs: J. Haije · E. Rolf · J. Blom
    </div>
  </div>
</div>

${(function(){
  // ── Portfoliostatus pagina (alleen bij meerdere apps) ────────
  if (visible.length <= 1) return "";

  const sc2  = visible.map(a => ({ ...a, sc: calcScores(a.scores || {}) })).filter(a => a.sc.autonomyScore);
  const avg2 = sc2.length ? sc2.reduce((s,a) => s + a.sc.autonomyScore, 0) / sc2.length : null;
  const avgD2= sc2.filter(a=>a.sc.dictuAvg).length ? sc2.filter(a=>a.sc.dictuAvg).reduce((s,a)=>s+a.sc.dictuAvg,0)/sc2.filter(a=>a.sc.dictuAvg).length : null;
  const sorted2 = [...sc2].sort((a,b) => (a.sc.autonomyScore||10)-(b.sc.autonomyScore||10));
  const kritiek2   = sc2.filter(a => a.sc.autonomyScore < 3).length;
  const zorg2      = sc2.filter(a => a.sc.autonomyScore >= 3 && a.sc.autonomyScore < 5).length;
  const acceptabel2= sc2.filter(a => a.sc.autonomyScore >= 5 && a.sc.autonomyScore < 7).length;
  const goed2      = sc2.filter(a => a.sc.autonomyScore >= 7).length;

  const oordeel2 = !sc2.length ? { kleur:"#6b7280", bg:"#f3f4f6", border:"#e5e7eb", tekst:"Nog geen applicaties beoordeeld." }
    : kritiek2 >= 3 || (kritiek2 > 0 && kritiek2/sc2.length > 0.3)
    ? { kleur:"#b91c1c", bg:"#fee2e2", border:"#fca5a5", tekst:"Het portfolio bevat " + kritiek2 + " kritieke applicatie" + (kritiek2!==1?"s":"") + " met een hoog autonomierisico en onvoldoende weerbaarheid. Directe besluitvorming is noodzakelijk." }
    : zorg2 > 0
    ? { kleur:"#c2410c", bg:"#ffedd5", border:"#fed7aa", tekst:"Het portfolio vraagt aandacht: " + (zorg2+kritiek2) + " applicatie" + ((zorg2+kritiek2)!==1?"s":"") + " scoren onder de acceptabele grens. Gerichte maatregelen zijn gewenst." }
    : { kleur:"#15803d", bg:"#dcfce7", border:"#86efac", tekst:"Het portfolio is grotendeels op orde. De meeste applicaties zijn acceptabel tot goed beoordeeld. Periodieke monitoring volstaat." };

  let html2 = '';
  html2 += '<div style="page-break-before:always">';
  html2 += '<div class="page-header"><div class="logo">NHL<br/>STENDEN</div><div style="width:2px;background:#26B5AE;align-self:stretch"></div><div><div class="header-title">Portfolioanalyse Digitale Soevereiniteit</div><div class="header-sub">Applicatielandschap NHL Stenden · ' + datum + ' · ' + VERSION + '</div></div><div class="header-right">' + naamModus + '</div></div>';
  html2 += '<div style="padding:14px 48px 20px">';
  html2 += '<h2 style="font-family:Arial;font-size:14px;color:#0C2340;padding-bottom:4px;margin:0 0 6px;font-weight:700;border-bottom:2px solid #1A56A0">Portfoliostatus — Samenvatting</h2>';
  html2 += '<div style="font-size:10px;line-height:1.5;color:#374151;margin-bottom:10px;border-left:3px solid #D0E4F7;padding-left:10px;font-family:Arial">Overzicht portfoliostatus digitale soevereiniteit NHL Stenden per ' + datum + '.</div>';

  // Oordeel + kerngetallen
  html2 += '<div style="display:grid;grid-template-columns:2fr 1fr;gap:10px;margin-bottom:10px">';
  html2 += '<div style="border:2px solid ' + oordeel2.border + ';border-radius:4px;padding:10px 12px;font-family:Arial">';
  html2 += '<div style="font-size:9px;font-weight:700;color:#9ca3af;text-transform:uppercase;letter-spacing:0.05em;margin-bottom:8px">Waar staat NHL Stenden</div>';
  html2 += '<div style="background:' + oordeel2.bg + ';border:1px solid ' + oordeel2.border + ';border-radius:3px;padding:10px 12px;margin-bottom:10px">';
  html2 += '<div style="font-size:11px;font-weight:600;color:' + oordeel2.kleur + ';line-height:1.5">' + oordeel2.tekst + '</div></div>';
  html2 += '<div style="font-size:9.5px;color:#374151;line-height:1.5;margin-bottom:8px">NHL Stenden heeft ' + visible.length + ' kernapplicatie' + (visible.length!==1?'s':'') + ' in scope genomen voor de portfolioanalyse digitale soevereiniteit. ' + (sc2.length < visible.length ? 'Van ' + (visible.length-sc2.length) + ' applicatie' + (visible.length-sc2.length!==1?'s':'') + ' is het assessment nog niet volledig ingevuld. ' : '') + 'De beoordeling combineert het DAAF-framework (autonomiescore 1–10) en de DICTU soevereiniteitscheck.</div>';
  html2 += '<div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:6px;margin-top:4px">';
  html2 += '<div style="background:#EBF3FF;border-left:3px solid #1A56A0;padding:6px 8px;border-radius:0 3px 3px 0;font-family:Arial">';
  html2 += '<div style="font-size:9px;font-weight:700;color:#1A56A0;margin-bottom:4px">Aanleiding</div>';
  html2 += '<div style="font-size:9px;color:#374151;line-height:1.55">Digitale soevereiniteit staat breed op de agenda in het hoger onderwijs. NHL Stenden brengt systematisch in kaart hoe afhankelijk de instelling is van externe leveranciers en welke risico\'s dat met zich meebrengt.</div>';
  html2 += '</div>';
  html2 += '<div style="background:#f0f9f9;border-left:3px solid #26B5AE;padding:6px 8px;border-radius:0 3px 3px 0;font-family:Arial">';
  html2 += '<div style="font-size:9px;font-weight:700;color:#26B5AE;margin-bottom:4px">Methodiek</div>';
  html2 += '<div style="font-size:9px;color:#374151;line-height:1.55">Twee erkende frameworks: <strong>DAAF</strong> (Utrecht University) beoordeelt autonomie op risico, mitigatie en belang. <strong>DICTU</strong> (Rijksoverheid) toetst technische soevereiniteit op vier dimensies.</div>';
  html2 += '</div>';
  html2 += '<div style="background:#fff8f0;border-left:3px solid #E87722;padding:6px 8px;border-radius:0 3px 3px 0;font-family:Arial">';
  html2 += '<div style="font-size:9px;font-weight:700;color:#E87722;margin-bottom:4px">Wat betekent de score?</div>';
  html2 += '<div style="font-size:9px;color:#374151;line-height:1.55">Score 1–10: hoe lager, hoe urgenter. Onder 5 actie gewenst, boven 7 acceptabel. Combineert risico, weerbaarheid en belang.</div>';
  html2 += '</div>';
  html2 += '</div>';
  html2 += '</div>';

  // Kerngetallen rechts
  html2 += '<div style="display:flex;flex-direction:column;gap:8px">';
  [[avg2 ? avg2.toFixed(1) : "–", "/10", "Gem. autonomiescore", avg2], [avgD2 ? avgD2.toFixed(1) : "–", "/5", "Gem. DICTU-score", avgD2 ? avgD2*2 : null]].forEach(function(k) {
    const c = !k[3] ? "#9ca3af" : k[3]>=7 ? "#16a34a" : k[3]>=5 ? "#ca8a04" : k[3]>=3 ? "#ea580c" : "#dc2626";
    html2 += '<div style="border:1px solid #D0E4F7;border-radius:4px;padding:7px;text-align:center;font-family:Arial;background:white">';
    html2 += '<div style="font-size:9px;color:#9ca3af;text-transform:uppercase;letter-spacing:0.05em;margin-bottom:4px">' + k[2] + '</div>';
    html2 += '<div style="font-size:20px;font-weight:700;color:' + c + ';line-height:1">' + k[0] + '<span style="font-size:11px;color:#9ca3af">' + k[1] + '</span></div>';
    html2 += '</div>';
  });
  // Verdeling
  html2 += '<div style="border:1px solid #D0E4F7;border-radius:4px;padding:10px;font-family:Arial;background:white">';
  html2 += '<div style="font-size:9px;color:#9ca3af;text-transform:uppercase;letter-spacing:0.05em;margin-bottom:6px">Verdeling</div>';
  [["Kritiek", kritiek2, "#fee2e2","#b91c1c","#dc2626"],["Zorgwekkend",zorg2,"#ffedd5","#c2410c","#ea580c"],["Acceptabel",acceptabel2,"#fef9c3","#a16207","#ca8a04"],["Goed",goed2,"#dcfce7","#15803d","#16a34a"]].forEach(function(r){
    html2 += '<div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:4px">';
    html2 += '<div style="display:flex;align-items:center;gap:5px"><div style="width:8px;height:8px;border-radius:50%;background:' + r[4] + '"></div><span style="font-size:10px;color:' + r[3] + ';font-weight:600">' + r[0] + '</span></div>';
    html2 += '<span style="font-size:10px;font-weight:700;background:' + r[2] + ';color:' + r[3] + ';padding:1px 8px;border-radius:2px">' + r[1] + '</span></div>';
  });
  html2 += '</div></div></div>';

  // Horizontale balkgrafiek
  html2 += '<div style="border:1px solid #D0E4F7;border-radius:4px;padding:10px 12px;font-family:Arial;margin-bottom:8px">';
  html2 += '<div style="font-size:11px;font-weight:700;color:#0C2340;margin-bottom:4px">Autonomiescore per applicatie</div>';
  html2 += '<div style="font-size:9px;color:#9ca3af;margin-bottom:10px">Gesorteerd van laagste naar hoogste · Gele lijn = grens acceptabel (5) · Groene lijn = grens goed (7)</div>';
  sorted2.forEach(function(a) {
    const s = a.sc.autonomyScore || 0;
    const klr = s >= 7 ? "#16a34a" : s >= 5 ? "#ca8a04" : s >= 3 ? "#ea580c" : "#dc2626";
    const naam = dName(a).substring(0,24);
    html2 += '<div style="display:flex;align-items:center;gap:8px;margin-bottom:3px">';
    html2 += '<div style="width:140px;text-align:right;font-size:9px;color:#374151;white-space:nowrap;overflow:hidden;text-overflow:ellipsis" title="' + dName(a) + '">' + naam + '</div>';
    html2 += '<div style="flex:1;background:#f1f5f9;border-radius:3px;height:13px;position:relative">';
    html2 += '<div style="position:absolute;left:50%;top:0;bottom:0;width:1px;background:#fbbf24;opacity:0.8"></div>';
    html2 += '<div style="position:absolute;left:70%;top:0;bottom:0;width:1px;background:#4ade80;opacity:0.8"></div>';
    html2 += '<div style="position:absolute;left:0;top:2px;bottom:2px;width:' + ((s/10)*100).toFixed(0) + '%;background:' + klr + ';border-radius:2px;min-width:' + (s>0?3:0) + 'px"></div>';
    html2 += '</div>';
    html2 += '<div style="width:28px;text-align:right;font-size:9px;font-weight:700;color:' + klr + '">' + s.toFixed(1) + '</div></div>';
  });
  html2 += '</div>';

  // Top 3
  const top3pdf = [...sc2].filter(a=>a.sc.autonomyScore<7).sort((a,b)=>{
    const uA=(10-(a.sc.autonomyScore||10))+(a.sc.belang||0);
    const uB=(10-(b.sc.autonomyScore||10))+(b.sc.belang||0);
    return uB-uA;
  }).slice(0,3);

  if (top3pdf.length > 0) {
    html2 += '<div style="border:1px solid #D0E4F7;border-radius:4px;padding:10px 12px;font-family:Arial">';
    html2 += '<div style="font-size:11px;font-weight:700;color:#0C2340;margin-bottom:4px">Top 3 aandachtspunten</div>';
    html2 += '<div style="font-size:9px;color:#9ca3af;margin-bottom:10px">Geselecteerd op combinatie van laagste score en hoogste strategisch belang</div>';
    const mKleur = ["#b91c1c","#c2410c","#a16207"];
    top3pdf.forEach(function(a, i) {
      const s = a.sc;
      const sc = s.autonomyScore || 0;
      const klr = sc>=7?"#16a34a":sc>=5?"#ca8a04":sc>=3?"#ea580c":"#dc2626";
      const lbl = sc>=7?"Goed":sc>=5?"Acceptabel":sc>=3?"Zorgwekkend":"Kritiek";
      const lblBg= sc>=7?"#dcfce7":sc>=5?"#fef9c3":sc>=3?"#ffedd5":"#fee2e2";
      html2 += '<div style="display:flex;gap:8px;padding:7px 10px;background:#f8fafc;border-radius:4px;border-left:3px solid ' + klr + ';margin-bottom:6px;page-break-inside:avoid">';
      html2 += '<div style="width:22px;height:22px;background:' + mKleur[i] + ';color:white;font-weight:700;font-size:11px;border-radius:3px;display:flex;align-items:center;justify-content:center;flex-shrink:0">' + (i+1) + '</div>';
      html2 += '<div style="flex:1">';
      html2 += '<div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin-bottom:6px">';
      html2 += '<span style="font-size:11px;font-weight:700;color:#0C2340">' + dName(a) + '</span>';
      if (a.supplier) html2 += '<span style="font-size:9px;color:#9ca3af">' + a.supplier + '</span>';
      html2 += '<span style="font-size:9px;font-weight:600;background:' + lblBg + ';color:' + klr + ';padding:1px 6px;border-radius:2px">' + lbl + '</span>';
      html2 += '<span style="font-size:10px;font-weight:700;color:' + klr + '">Score: ' + sc.toFixed(1) + '/10</span></div>';
      html2 += '<div style="display:flex;gap:8px">';
      [["Risico",s.risico,"#dc2626"],["Mitigatie",s.mitigatie,"#26B5AE"],["Belang",s.belang,"#E87722"],["DICTU",s.dictuAvg,"#6d28d9"]].forEach(function(d){
        html2 += '<div style="text-align:center;background:white;border:1px solid #e5e7eb;border-radius:3px;padding:3px 6px;min-width:44px">';
        html2 += '<div style="font-size:11px;font-weight:700;color:' + (d[1]?d[2]:"#d1d5db") + '">' + (d[1]?d[1].toFixed(1):"–") + '</div>';
        html2 += '<div style="font-size:8px;color:#9ca3af">' + d[0] + '</div></div>';
      });
      html2 += '</div></div></div>';
    });
    html2 += '</div>';
  }

  html2 += '</div>';
  html2 += '</div>';
  return ""; // portfoliopagina niet opnemen in applicatie-PDF
})()}

${(function(){
  const appNaam = visible.length === 1 ? dName(visible[0]) : "Portfolioanalyse Digitale Soevereiniteit";
  const appSub  = visible.length === 1 ? (visible[0].supplier ? visible[0].supplier + " - " : "") + "NHL Stenden Hogeschool" : "NHL Stenden Hogeschool";
  const hdrTitel = visible.length === 1 ? "Assessment Digitale Soevereiniteit - " + dName(visible[0]) : "Portfolioanalyse Digitale Soevereiniteit";
  let ep = "";
  ep += '<div style="page-break-before:always">';
  ep += '<div class="page-header"><div class="logo">NHL<br/>STENDEN</div><div style="width:2px;background:#26B5AE;align-self:stretch"></div><div><div class="header-title">' + hdrTitel + '</div><div class="header-sub">NHL Stenden Hogeschool - ' + datum + ' - ' + VERSION + '</div></div><div class="header-right">' + naamModus + '</div></div>';
  ep += '<div style="height:calc(100vh - 80px);display:flex;flex-direction:column;align-items:center;justify-content:center;font-family:Arial;text-align:center;padding:48px">';
  ep += '<div style="width:60px;height:4px;background:#26B5AE;border-radius:2px;margin-bottom:32px"></div>';
  ep += '<div style="font-size:9px;color:#9ca3af;letter-spacing:0.15em;text-transform:uppercase;margin-bottom:16px">Einde rapport</div>';
  ep += '<div style="font-size:24px;font-weight:700;color:#0C2340;margin-bottom:6px;line-height:1.2">' + appNaam + '</div>';
  ep += '<div style="font-size:11px;color:#6b7280;margin-bottom:6px">' + appSub + '</div>';
  ep += '<div style="font-size:10px;color:#9ca3af;margin-bottom:32px">' + datum + ' - ' + VERSION + '</div>';
  ep += '<div style="width:60px;height:4px;background:#1A56A0;border-radius:2px;margin-bottom:40px"></div>';
  ep += '<div style="font-size:10px;color:#9ca3af;line-height:1.7;max-width:400px">Dit document is vertrouwelijk en bestemd voor intern gebruik binnen NHL Stenden Hogeschool.<br/>Programma Digitale Samenhang - Ambassadeurslijn Digitale Soevereiniteit<br/>Kwartiermaker: E. van Gorkum - Ambassadeurs: J. Haije - E. Rolf - J. Blom</div>';
  ep += '</div></div>';
  return ep;
})()}
</body>
</html>`;

    // ── Bestandsnaam voor "Opslaan als PDF" ─────────────────────
    const now = new Date();
    const ts  = now.getFullYear().toString()
      + String(now.getMonth()+1).padStart(2,"0")
      + String(now.getDate()).padStart(2,"0")
      + "_" + String(now.getHours()).padStart(2,"0")
      + String(now.getMinutes()).padStart(2,"0");
    const pdfNaam = visible.length === 1
      ? "Assessment_" + dName(visible[0]).replace(/[^a-zA-Z0-9]/g, "_").replace(/_+/g,"_") + "_" + ts
      : "NHL_Stenden_Portfolioanalyse_Soevereiniteit_" + ts;

    const win = window.open("", "_blank");
    win.document.write(html);
    win.document.close();
    win.document.title = pdfNaam;
    win.focus();
    setTimeout(() => { win.document.title = pdfNaam; win.print(); }, 600);
  }

  // ── VIEWS ──────────────────────────────────────────────────

  function Dashboard() {
    // ── Zichtbare applicaties (gefilterd op hiddenApps) ─────────
    const visibleApps = apps.filter(a => !hiddenApps.has(a.id));

    function toggleApp(id) {
      setHiddenApps(prev => {
        const next = new Set(prev);
        if (next.has(id)) {
          // Zichtbaar maken: controleer max
          const currentVisible = apps.filter(a => !next.has(a.id)).length;
          if (currentVisible >= MAX_VISIBLE) return prev; // max voor leesbaarheid grafieken
          next.delete(id);
        } else {
          // Verbergen: minimaal 1 zichtbaar houden
          if (apps.filter(a => !next.has(a.id)).length <= 1) return prev;
          next.add(id);
        }
        return next;
      });
    }

    function selectLaatste10() {
      const toHide = new Set(apps.slice(0, Math.max(0, apps.length - MAX_VISIBLE)).map(a => a.id));
      setHiddenApps(toHide);
    }

    function selectAlles() {
      if (apps.length <= MAX_VISIBLE) {
        setHiddenApps(new Set());
      }
    }

    const scored = visibleApps.map(a => ({ ...a, sc: calcScores(a.scores) }));
    const allScored = apps.map(a => ({ ...a, sc: calcScores(a.scores) })); // voor filter-strip
    const withSc = scored.filter(a => a.sc.autonomyScore);
    const avgA   = withSc.length ? withSc.map(a => a.sc.autonomyScore).reduce((x,y)=>x+y,0)/withSc.length : null;

    const radarKey = n => n.substring(0, 13);

    // Dimensielabels
    const dimLetters = [...new Set(DAAF.map(d => d.dim))];
    const dimLabel = letter => {
      const first = DAAF.find(d => d.dim === letter);
      return first ? first.dimName.substring(0, 14) : letter;
    };

    // Radar: directe berekening per dimensie vanuit ruwe scores
    // (omzeilt dims-veld van calcScores om rendering-bugs te vermijden)
    const dimScore = (a, letter) => {
      if (letter === "A") {
        const a1 = a.scores["A1"] || 0, a3 = a.scores["A3"] || 0;
        const pairs = [[a1, 3], [a3, 2]].filter(([v]) => v > 0);
        if (!pairs.length) return 0;
        const tw = pairs.reduce((s, [,w]) => s+w, 0);
        return pairs.reduce((s, [v,w]) => s+v*w, 0) / tw;
      }
      const qs = DAAF.filter(d => d.dim === letter);
      const vals = qs.map(q => a.scores[q.key] || 0).filter(v => v > 0);
      return vals.length ? vals.reduce((s,v)=>s+v,0)/vals.length : 0;
    };

    const radarData = dimLetters.map(letter => {
      const entry = { dim: dimLabel(letter) };
      scored.forEach(a => {
        entry[radarKey(displayName(a))] = +dimScore(a, letter).toFixed(2);
      });
      return entry;
    });

    // Kwadrant data: X = risico × belang (1–25), Y = mitigatie (1–5)
    const kwData = scored
      .filter(a => a.sc.risico && a.sc.mitigatie && a.sc.belang)
      .map(a => ({
        name: displayName(a),
        x: +((a.sc.risico * a.sc.belang).toFixed(2)),
        y: +a.sc.mitigatie.toFixed(2),
        score: a.sc.autonomyScore,
        id: a.id
      }));

    const handleKwClick = (id) => { setSelId(id); setStep(0); setAssessReadOnly(true); setView("assess"); };

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
              <div className="flex items-center gap-2 flex-shrink-0 flex-wrap">
                <span style={{ fontSize:11, fontWeight:600, color:"#0C2340" }}>Grafiek-selectie</span>
                <span className="text-xs px-2 py-0.5 rounded font-semibold"
                  style={{ background: visibleApps.length >= MAX_VISIBLE ? "#fffbeb" : "#f0f9f9",
                           color: visibleApps.length >= MAX_VISIBLE ? "#92400e" : "#0f766e" }}>
                  {visibleApps.length} van {apps.length} zichtbaar in grafieken
                </span>
                {apps.length > MAX_VISIBLE && visibleApps.length >= MAX_VISIBLE && (
                  <span style={{ fontSize:10, color:"#9ca3af" }}>
                    max {MAX_VISIBLE} voor leesbaarheid
                  </span>
                )}
              </div>
              <div className="flex gap-2 flex-wrap flex-1">
                {allScored.map((a, i) => {
                  const hidden  = hiddenApps.has(a.id);
                  const col     = appColor(i);
                  return (
                    <button key={a.id}
                      onClick={() => toggleApp(a.id)}
                      disabled={(!hidden && visibleApps.length <= 1) || (hidden && visibleApps.length >= MAX_VISIBLE)}
                      title={
                        !hidden && visibleApps.length <= 1 ? "Minimaal 1 applicatie moet zichtbaar blijven" :
                        hidden && visibleApps.length >= MAX_VISIBLE ? `Maximum van ${MAX_VISIBLE} applicaties bereikt — verberg eerst een andere` :
                        hidden ? "Klik om toe te voegen aan selectie" : "Klik om uit selectie te verwijderen"
                      }
                      className="flex items-center gap-1.5 text-xs px-2.5 py-1.5 font-medium transition-all"
                      style={{
                        borderRadius: 4,
                        border: `2px solid ${hidden ? "#e5e7eb" : col}`,
                        background: hidden ? "#f9fafb" : `${col}18`,
                        color: hidden ? "#9ca3af" : col,
                        opacity: ((!hidden && visibleApps.length <= 1) || (hidden && visibleApps.length >= MAX_VISIBLE)) ? 0.4 : 1,
                        cursor: ((!hidden && visibleApps.length <= 1) || (hidden && visibleApps.length >= MAX_VISIBLE)) ? "not-allowed" : "pointer",
                        textDecoration: hidden ? "line-through" : "none"
                      }}>
                      <span className="w-2 h-2 rounded-full flex-shrink-0"
                        style={{ background: hidden ? "#d1d5db" : col }} />
                      {displayName(a).substring(0, 22)}
                      <span style={{ fontSize:10, marginLeft:2, opacity:0.7 }}>
                        {hidden ? "＋" : "✕"}
                      </span>
                    </button>
                  );
                })}
              </div>
              <div className="flex gap-2 flex-shrink-0 flex-wrap">
                {apps.length > MAX_VISIBLE && (
                  <button onClick={selectLaatste10}
                    className="text-xs px-2.5 py-1.5 font-medium"
                    style={{ borderRadius:4, background:"#EBF3FF", color:"#1A56A0", border:"1px solid #D0E4F7" }}>
                    Laatste {MAX_VISIBLE}
                  </button>
                )}
                {apps.length <= MAX_VISIBLE && hiddenApps.size > 0 && (
                  <button onClick={() => setHiddenApps(new Set())}
                    className="text-xs px-2.5 py-1.5 font-medium"
                    style={{ borderRadius:4, background:"#EBF3FF", color:"#1A56A0", border:"1px solid #D0E4F7" }}>
                    Alles tonen
                  </button>
                )}
                <div className="w-px self-stretch" style={{ background:"#D0E4F7", margin:"0 4px" }}/>
                <button onClick={() => setShowModal(true)}
                  className="text-white text-xs px-3 py-1.5 font-medium"
                  style={{ background:"#1A56A0", borderRadius:4 }}>
                  + Toevoegen
                </button>
                <button onClick={exportDashboardPdf} disabled={visibleApps.length===0}
                  className="text-xs px-2.5 py-1.5 font-medium"
                  style={{ background:"#fee2e2", color:"#b91c1c", borderRadius:4, border:"1px solid #fecaca", opacity:visibleApps.length===0?0.5:1 }}>
                  📄 PDF
                </button>
                <button onClick={exportXlsx} disabled={apps.length===0}
                  className="text-xs px-2.5 py-1.5 font-medium"
                  style={{ background:"#E6F7F7", color:"#26B5AE", borderRadius:4, border:"1px solid #26B5AE55", opacity:apps.length===0?0.5:1 }}>
                  📥 Excel
                </button>
              </div>
            </div>
          )}

          {apps.length === 0 ? (
            <div className="bg-white rounded border-2 border-dashed border-gray-200 p-16 text-center">
              <div style={{ fontSize: 48, marginBottom: 12 }}>📊</div>
              <h3 className="text-xl font-semibold text-gray-700 mb-2">Start met uw assessment</h3>
              <p className="text-gray-400 text-sm mb-5">Voeg een applicatie toe om te beginnen.</p>
              <button onClick={() => setShowModal(true)}
                className="text-white text-sm px-5 py-2.5 font-medium"
                style={{ background:"#1A56A0", borderRadius:4 }}>
                + Applicatie toevoegen
              </button>
            </div>
          ) : (<>

          {/* ── Opdrachtskaart — centrale vraagstelling ── */}
          <OpdrachtKaart apps={apps} useSecondaryName={useSecondaryName} />

          {/* ── Rij 1: Kwadrant volledig breed ── */}
          <div className="grid gap-4 mb-4" style={{ gridTemplateColumns:"1fr" }}>

            {/* Autonomie-kwadrant — volledig breed */}
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
              <KwadrantSVG kwData={kwData} onAppClick={handleKwClick} />
              <div className="mt-2 px-2 py-1 rounded text-xs" style={{ background:"#f8fafc", border:"1px solid #e5e7eb", color:"#9ca3af" }}>
                Risico = gem. A1+A3+B1 · Mitigatie = gem. C1+D1+E1 · Belang = gem. F1+G1+H1 · Grens: X=13, Y=3
              </div>
            </div>
          </div>

          {/* ── Rij 2: Leeswijzer + App-kaarten ── */}
          <div className="mb-4">
              {/* Leeswijzer bovenaan app-kaarten */}
              <div className="grid grid-cols-2 gap-2 mb-3">
                <div className="rounded p-2" style={{ background:"#EBF3FF", border:"1px solid #D0E4F7" }}>
                  <div className="flex items-center gap-1.5 mb-1">
                    <span style={{ fontSize:9, fontWeight:700, padding:"1px 5px", background:"#1A56A0", color:"#fff", borderRadius:3 }}>DAAF</span>
                    <span style={{ fontSize:10, fontWeight:600, color:"#0C2340" }}>Snelheidsmeter (1–10)</span>
                  </div>
                  <div className="flex gap-1 flex-wrap">
                    {[{t:"≥7 Goed",bg:"#dcfce7",fg:"#15803d"},{t:"5–7 OK",bg:"#fef9c3",fg:"#a16207"},{t:"3–5 Let op",bg:"#ffedd5",fg:"#c2410c"},{t:"<3 Kritiek",bg:"#fee2e2",fg:"#b91c1c"}].map(s=>(
                      <span key={s.t} style={{ fontSize:9, background:s.bg, color:s.fg, borderRadius:2, padding:"1px 5px", fontWeight:600 }}>{s.t}</span>
                    ))}
                  </div>
                </div>
                <div className="rounded p-2" style={{ background:"#E6F7F7", border:"1px solid #26B5AE44" }}>
                  <div className="flex items-center gap-1.5 mb-1">
                    <span style={{ fontSize:9, fontWeight:700, padding:"1px 5px", background:"#26B5AE", color:"#fff", borderRadius:3 }}>DICTU</span>
                    <span style={{ fontSize:10, fontWeight:600, color:"#0C2340" }}>Kleurenbalk (1–5)</span>
                  </div>
                  <div className="flex gap-1 flex-wrap">
                    {[{t:"1–2 Afhankelijk",bg:"#fee2e2",fg:"#b91c1c"},{t:"3 Deels",bg:"#fef9c3",fg:"#a16207"},{t:"4–5 Soeverein",bg:"#dcfce7",fg:"#15803d"}].map(s=>(
                      <span key={s.t} style={{ fontSize:9, background:s.bg, color:s.fg, borderRadius:2, padding:"1px 5px", fontWeight:600 }}>{s.t}</span>
                    ))}
                  </div>
                </div>
              </div>

              <h3 className="text-sm font-bold mb-3" style={{ color:"#0C2340" }}>
                Applicaties ({visibleApps.length})
              </h3>
              {scored.length > 20 && (
                <p className="text-xs mb-2" style={{ color:"#9ca3af" }}>
                  {scored.length} applicaties — scroll voor het volledige overzicht
                </p>
              )}
              <div className="grid gap-3" style={{ gridTemplateColumns:"repeat(auto-fill, minmax(280px, 1fr))" }}>
                {scored.map(a => {
                  const lbl = scoreLabel(a.sc.autonomyScore);
                  return (
                    <div key={a.id}
                      onClick={() => { setSelId(a.id); setStep(0); setAssessReadOnly(false); setView("assess"); }}
                      className="cursor-pointer transition-all"
                      style={{ background:"#fff", borderRadius:4, padding:14, border:"1px solid #D0E4F7", borderLeft:`4px solid ${scoreColor(a.sc.autonomyScore)}` }}
                      onMouseEnter={e=>e.currentTarget.style.boxShadow="0 2px 10px rgba(26,86,160,0.15)"}
                      onMouseLeave={e=>e.currentTarget.style.boxShadow="none"}>
                      <div className="flex items-start gap-2 mb-2">
                        <div className="flex flex-col items-center gap-0.5 flex-shrink-0">
                          <Gauge score={a.sc.autonomyScore} size={58} />
                          <span style={{ fontSize:8, fontWeight:700, color:"#9ca3af", letterSpacing:"0.04em" }}>DAAF</span>
                        </div>
                        <div className="flex-1 min-w-0">
                          <h3 className="font-bold text-xs truncate" style={{ color:"#0C2340" }}>{displayName(a)}</h3>
                          {a.supplier && <p style={{ fontSize:10, color:"#9ca3af" }}>{a.supplier}</p>}
                          <span className="text-xs px-1.5 py-0.5 font-medium mt-1 inline-block"
                            style={{ borderRadius:3, background: lbl.bg, color: lbl.fg, fontSize:10 }}>{lbl.text}</span>
                        </div>
                      </div>
                      <div className="mb-2">
                        <div className="flex justify-between items-center" style={{ fontSize:9, marginBottom:2 }}>
                          <div className="flex items-center gap-1">
                            <span style={{ fontSize:8, fontWeight:700, padding:"0 4px", background:"#26B5AE", color:"#fff", borderRadius:2 }}>DICTU</span>
                            <span style={{ color:"#9ca3af" }}>Soevereiniteit</span>
                          </div>
                          <span style={{ color: a.sc.dictuAvg ? (a.sc.dictuAvg >= 4 ? "#15803d" : a.sc.dictuAvg >= 3 ? "#a16207" : "#b91c1c") : "#9ca3af", fontWeight:600 }}>
                            {a.sc.dictuAvg ? a.sc.dictuAvg.toFixed(1)+"/5" : "–"}
                          </span>
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
                      {/* Aanbevelingen — toon alleen als app volledig genoeg is */}
                      {a.sc.autonomyScore && (() => {
                        const rec = generateRecommendations(a.scores);
                        return (
                          <div className="mt-2 space-y-1">
                            <div className="rounded px-2 py-1.5" style={{ background:"#fffbeb", border:"1px solid #fde68a" }}>
                              <p style={{ fontSize:9, fontWeight:700, color:"#92400e", marginBottom:2 }}>⚡ Quick win</p>
                              <p style={{ fontSize:9, color:"#78350f", lineHeight:1.45 }}>{adaptNote(rec.quickWin, a)}</p>
                            </div>
                            <div className="rounded px-2 py-1.5" style={{ background:"#f0fdf4", border:"1px solid #86efac" }}>
                              <p style={{ fontSize:9, fontWeight:700, color:"#166534", marginBottom:2 }}>🎯 Strategische aanbeveling</p>
                              <p style={{ fontSize:9, color:"#14532d", lineHeight:1.45 }}>{adaptNote(rec.strategic, a)}</p>
                            </div>
                          </div>
                        );
                      })()}
                    </div>
                  );
                })}
              </div>
          </div>

          {/* ── Rij 2: Dimensieprofiel full-width ── */}
          <div className="rounded p-4 mb-4" style={{ background:"#fff", border:"1px solid #D0E4F7" }}>
            <div className="flex items-center gap-2 mb-1">
              <span className="text-xs font-bold px-2 py-0.5" style={{ background:"#1A56A0", color:"#fff", borderRadius:3 }}>DAAF</span>
              <h3 className="font-bold" style={{ color:"#0C2340", fontSize:14 }}>Dimensieprofiel — per applicatie vergelijken</h3>
            </div>
            <p className="text-xs leading-relaxed mb-3" style={{ color:"#6b7280" }}>
              Elke balk toont de gewogen score (1–5) op één DAAF-dimensie. De kleurovergang geeft direct de kwaliteit aan:
              voor <span style={{ color:"#dc2626", fontWeight:600 }}>risico-assen (A, B)</span> is links/groen beter.
              Voor <span style={{ color:"#16a34a", fontWeight:600 }}>mitigatie-assen (C, D, E)</span> is rechts/groen beter.
              Hover over een punt voor applicatienaam en exacte score.
            </p>
            {scored.length >= 1
              ? <DivergingChart apps={scored} useSecondaryName={useSecondaryName} />
              : <div className="text-center py-10 text-sm text-gray-400">Voeg een applicatie toe om het spindiagram te zien.</div>
            }
            {/* Compacte leeswijzer */}
            <div className="grid grid-cols-3 gap-2 mt-3">
              {[
                { label:"Punt op de balk", text:"Elk gekleurde punt = één applicatie. De positie op de balk (1–5) is de gewogen dimensiescore. Hover voor applicatienaam en exacte waarde." },
                { label:"Kleurovergang", text:"De achtergrondkleur van de balk geeft direct de richting aan: groen = gewenste kant. Risico-assen: groen links (laag risico). Mitigatie-assen: groen rechts (hoog = meer weerbaarheid)." },
                { label:"Belang-assen (F, G, H)", text:"Hoge belang-scores (oranje/bruin) betekenen meer strategische urgentie. Dit is niet per se slecht — maar vraagt dat de risico- en mitigatiescores voor die applicatie goed op orde zijn." },
              ].map(t => (
                <div key={t.label} className="rounded p-2" style={{ background:"#f8fafc", border:"1px solid #e5e7eb" }}>
                  <p style={{ fontSize:9, fontWeight:700, color:"#0C2340", marginBottom:2 }}>{t.label}</p>
                  <p style={{ fontSize:9, color:"#6b7280", lineHeight:1.4 }}>{t.text}</p>
                </div>
              ))}
            </div>
          </div>

          {/* ── Rij 3: DICTU spindiagram ── */}
          {scored.some(a => a.sc.dictuAvg) && (
            <div className="rounded p-4 mb-4" style={{ background:"#fff", border:"1px solid #D0E4F7" }}>
              <div className="flex items-center gap-2 mb-1">
                <span className="text-xs font-bold px-2 py-0.5" style={{ background:"#26B5AE", color:"#fff", borderRadius:3 }}>DICTU</span>
                <h3 className="font-bold" style={{ color:"#0C2340", fontSize:14 }}>Soevereiniteitscheck — spindiagram per applicatie</h3>
              </div>
              <p className="text-xs mb-3" style={{ color:"#6b7280" }}>
                Het spindiagram toont de vier DICTU-dimensies per applicatie op een schaal van 1 tot 5. Alle assen hebben dezelfde richting: hoe groter de gekleurde vorm, hoe soeverein de applicatie scoort.
                Een kleine vorm dicht bij het centrum betekent <span style={{ color:"#dc2626", fontWeight:600 }}>volledig afhankelijk</span> van de leverancier.
                Een grote vorm die de buitenste ring raakt is <span style={{ color:"#26B5AE", fontWeight:600 }}>maximaal soeverein</span>.
              </p>
              <div style={{ maxWidth:640, margin:"0 auto" }}>
                <DictuRadarSVG apps={scored} W={660} H={420} useSecondaryName={useSecondaryName} />
              </div>
              {/* Legenda dimensies */}
              <div className="grid grid-cols-4 gap-2 mt-3">
                {[
                  { key:"2.1", label:"Data residency", uitleg:"Worden data en back-ups uitsluitend binnen de EU opgeslagen?" },
                  { key:"2.2", label:"Technische beveiliging", uitleg:"Zijn er technische maatregelen die ongeoorloofde toegang uitsluiten?" },
                  { key:"2.3", label:"Juridische bescherming", uitleg:"Is de aanbieder verplicht niet-EU datavorderingen te bestrijden?" },
                  { key:"4.1", label:"EU-infrastructuur", uitleg:"Bevinden infrastructuur en control plane zich volledig in de EU?" },
                ].map(d => (
                  <div key={d.key} className="rounded p-2" style={{ background:"#E6F7F7", border:"1px solid #26B5AE44" }}>
                    <p style={{ fontSize:9, fontWeight:700, color:"#0C6B68", marginBottom:2 }}>{d.key} — {d.label}</p>
                    <p style={{ fontSize:9, color:"#374151", lineHeight:1.4 }}>{d.uitleg}</p>
                  </div>
                ))}
              </div>
            </div>
          )}

          </>)}
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
              {[...apps].sort((a, b) => {
                const na = (useSecondaryName && a.nameSecondary ? a.nameSecondary : a.name).toLowerCase();
                const nb = (useSecondaryName && b.nameSecondary ? b.nameSecondary : b.name).toLowerCase();
                return na.localeCompare(nb, "nl", { sensitivity:"base", numeric:true });
              }).map(a => {
                const sc  = calcScores(a.scores);
                const lbl = scoreLabel(sc.autonomyScore);
                return (
                  <div key={a.id} className="flex items-center gap-4 rounded p-4"
                    style={{ background:"#fff", border:"1px solid #D0E4F7",
                             borderLeft:`4px solid ${scoreColor(sc.autonomyScore)}` }}>
                    <Gauge score={sc.autonomyScore} size={64} />
                    <div className="flex-1 min-w-0">
                      <h3 className="font-semibold" style={{ color:"#0C2340" }}>{displayName(a)}</h3>
                      <p className="text-xs text-gray-500">{a.supplier}{a.cat ? ` · ${a.cat}` : ""}</p>
                      <div className="flex items-center gap-2 mt-1">
                        <span className="text-xs px-2 py-0.5 font-medium" style={{ borderRadius:3, background: lbl.bg, color: lbl.fg }}>
                          {lbl.text}
                        </span>
                        <span className="text-xs text-gray-400">{sc.completeness}% van vragen ingevuld</span>
                      </div>
                    </div>
                    <div className="flex gap-2 flex-shrink-0">
                      <button onClick={() => { setSelId(a.id); setStep(0); setAssessReadOnly(!adminUnlocked); setView("assess"); }}
                        className="text-white text-xs px-3 py-1.5 font-medium"
                        style={{ background: adminUnlocked ? "#1A56A0" : "#6b7280", borderRadius:4 }}>
                        {adminUnlocked ? "✏️ Bewerken" : "👁 Bekijken"}
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
              <h2 className="font-semibold truncate" style={{ color:"#0C2340" }}>{displayName(selApp)}</h2>
            </div>

            {/* Readonly banner */}
            {assessReadOnly && (
              <div className="flex items-center justify-between px-4 py-2.5 rounded mb-4"
                style={{ background:"#fffbeb", border:"1px solid #fde68a" }}>
                <div className="flex items-center gap-2">
                  <span style={{ fontSize:16 }}>🔒</span>
                  <span className="text-xs font-semibold" style={{ color:"#92400e" }}>
                    Alleen-lezen — scores wijzigen kan via Beheer
                  </span>
                </div>
                <button onClick={() => setView("admin")}
                  className="text-xs px-3 py-1.5 font-semibold"
                  style={{ background:"#1A56A0", color:"white", borderRadius:4 }}>
                  Naar Beheer →
                </button>
              </div>
            )}

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
                <strong>DAAF Framework — Quick Scan</strong> — Niveau 1 (Risico): score 1 = weinig risico, score 5 = hoog risico.
                Niveau 2 (Mitigatie): score 1 = slechte mitigatie, score 5 = sterke mitigatie (goed).
                Niveau 3 (Belang): score 1 = laag belang, score 5 = hoog belang (urgent).
              </div>
              {["Risico","Mitigatie","Belang"].map(lv => {
                const colors  = { Risico:"#dc2626", Mitigatie:"#26B5AE", Belang:"#E87722" };
                const bgColors = { Risico:"#0C2340", Mitigatie:"#166534", Belang:"#92400e" };
                const hints   = {
                  Risico:    "1 = weinig risico ... 5 = hoog risico",
                  Mitigatie: "1 = slechte mitigatie ... 5 = sterke mitigatie (goed)",
                  Belang:    "1 = laag belang ... 5 = hoog belang (urgent)"
                };
                const dims = DAAF.filter(d => d.level === lv);
                // Groepeer per unieke dimensie
                const uniqueDims = [...new Set(dims.map(d => d.dimName))];
                return (
                  <div key={lv} className="mb-5">
                    {/* Niveau header — stijl conform framework */}
                    <div className="rounded px-3 py-2 mb-3 flex items-center justify-between"
                      style={{ background: bgColors[lv] }}>
                      <span className="text-white font-semibold text-sm">Niveau {lv === "Risico" ? "1" : lv === "Mitigatie" ? "2" : "3"}: {lv === "Risico" ? "Risico-exposure" : lv === "Mitigatie" ? "Mitigatie-capaciteit" : "Strategisch belang"}</span>
                      <span className="text-xs" style={{ color:"rgba(255,255,255,0.7)" }}>{hints[lv]}</span>
                    </div>
                    {uniqueDims.map(dimName => {
                      const dimQuestions = dims.filter(d => d.dimName === dimName);
                      const firstQ = dimQuestions[0];
                      return (
                        <div key={dimName} className="mb-4">
                          {/* Dimensie header */}
                          <div className="rounded px-3 py-2 mb-2" style={{ background:"#f3f4f6", border:"1px solid #e5e7eb" }}>
                            <p className="font-semibold text-sm" style={{ color:"#0C2340" }}>
                              Dimensie {firstQ.dim}: {dimName}
                            </p>
                            {firstQ.dimSub && <p className="text-xs text-gray-500 mt-0.5">{firstQ.dimSub}</p>}
                          </div>
                          {dimQuestions.map(q => (
                            <QuestionCard key={q.key} q={q} value={selApp.scores[q.key] || 0}
                              dir={lv === "Mitigatie" ? "fwd" : "inv"}
                              note={(selApp.notes || {})[q.key] || ""}
                              onNoteChange={t => !assessReadOnly && setNote(selApp.id, q.key, t)}
                              onChange={v => !assessReadOnly && setScore(selApp.id, q.key, v)}
                              readOnly={assessReadOnly}
                              useSecondaryName={useSecondaryName}
                              appName={selApp.name}
                              appNameSecondary={selApp.nameSecondary || ""} />
                          ))}
                        </div>
                      );
                    })}
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
                      note={(selApp.notes || {})[q.key] || ""}
                      onNoteChange={t => !assessReadOnly && setNote(selApp.id, q.key, t)}
                      onChange={v => !assessReadOnly && setScore(selApp.id, q.key, v)}
                      readOnly={assessReadOnly}
                      useSecondaryName={useSecondaryName}
                      appName={selApp.name}
                      appNameSecondary={selApp.nameSecondary || ""} />
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

  // ── GEOKAART ───────────────────────────────────────────────────────────────
  function GeoKaart() {
    const vandaag = new Date().toLocaleDateString("nl-NL", { day:"numeric", month:"long", year:"numeric" });

    const REGIO_COORDS = {
      "EU-West":   { lon:  10, lat: 51, label: "EU West"             },
      "EU-Noord":  { lon:  15, lat: 56, label: "EU Noord"            },
      "EU-Oost":   { lon:  22, lat: 49, label: "EU Oost"             },
      "NL":        { lon:   5, lat: 52, label: "Nederland"           },
      "VK":        { lon:  -2, lat: 53, label: "Verenigd Koninkrijk" },
      "VS-Oost":   { lon: -75, lat: 40, label: "VS Oost"             },
      "VS-West":   { lon:-120, lat: 37, label: "VS West"             },
      "VS":        { lon: -95, lat: 38, label: "Verenigde Staten"    },
      "Azië":      { lon: 105, lat: 30, label: "Azië"                },
      "China":     { lon: 110, lat: 35, label: "China"               },
      "India":     { lon:  78, lat: 20, label: "India"               },
      "Japan":     { lon: 138, lat: 36, label: "Japan"               },
      "Australië": { lon: 134, lat:-25, label: "Australië"           },
      "Brazilië":  { lon: -52, lat:-14, label: "Brazilië"            },
      "Onbekend":  { lon:   0, lat:  0, label: "Onbekend"            },
    };

    function regioVanScore(a1, a3) {
      let jurisRegio = a1 <= 2 ? "EU-West" : a1 === 3 ? "VS" : a1 >= 4 ? "VS" : "Onbekend";
      let dataRegio  = a3 <= 2 ? "EU-West" : a3 === 3 ? "EU-West" : a3 >= 4 ? "VS" : "Onbekend";
      return { jurisRegio, dataRegio };
    }

    function risicoKleur(score) {
      if (!score) return "#9ca3af";
      if (score <= 2) return "#16a34a";
      if (score === 3) return "#ca8a04";
      if (score === 4) return "#ea580c";
      return "#dc2626";
    }

    const visibleAppsGeo = apps.filter(function(a) { return !geoHidden.has(a.id); });

    const scored = visibleAppsGeo.map(function(a) {
      const a1 = a.scores["A1"] || 0;
      const a3 = a.scores["A3"] || 0;
      const sc = calcScores(a.scores);
      const regio = regioVanScore(a1, a3);
      return Object.assign({}, a, { a1, a3, sc, jurisRegio: regio.jurisRegio, dataRegio: regio.dataRegio });
    }).filter(function(a) { return a.a1 > 0 || a.a3 > 0; });

    const incomplete = visibleAppsGeo.filter(function(a) { return !a.scores["A1"] && !a.scores["A3"]; });

    const jurisGroups = {};
    const dataGroups  = {};
    scored.forEach(function(a) {
      if (!jurisGroups[a.jurisRegio]) jurisGroups[a.jurisRegio] = [];
      jurisGroups[a.jurisRegio].push(Object.assign({}, a, { score: a.a1 }));
      if (!dataGroups[a.dataRegio]) dataGroups[a.dataRegio] = [];
      dataGroups[a.dataRegio].push(Object.assign({}, a, { score: a.a3 }));
    });

    const a1lbl = ["","EU/EER volledig","EU/EER beperkt","Adequaat + risico","SCCs, geen adequaat","Geen waarborgen"];
    const a3lbl = ["","EU/EER contractueel","EU/EER + adequaat","EU/EER, geen garantie","Deels buiten EU","Buiten EU"];

    const sortedApps = [...apps].sort(function(a, b) {
      return Math.max(b.scores["A1"]||0, b.scores["A3"]||0) - Math.max(a.scores["A1"]||0, a.scores["A3"]||0);
    });
    const visibleSortedApps = sortedApps.filter(function(a) { return !geoHidden.has(a.id); });

    return (
      <div className="h-full overflow-y-auto" style={{ background:"#EBF3FF" }}>
        <div className="p-5 max-w-5xl mx-auto">

          {/* Header */}
          <div className="rounded p-5 mb-5 text-white" style={{ background:"linear-gradient(135deg, #0C2340 0%, #1A56A0 100%)" }}>
            <div className="flex items-start justify-between">
              <div className="flex items-center gap-3">
                <div className="px-3 py-2 border-2 border-white" style={{ borderRadius:2 }}>
                  <span className="font-bold leading-none" style={{ fontSize:10, letterSpacing:1 }}>NHL<br/>STENDEN</span>
                </div>
                <div className="w-px self-stretch" style={{ background:"#26B5AE", margin:"2px 0" }}/>
                <div>
                  <h1 className="font-bold" style={{ fontSize:17 }}>Geopolitiek Landschap</h1>
                  <p style={{ fontSize:12, color:"#7DD3D0" }}>Jurisdictie en datalocatie van het applicatieportfolio · {vandaag}</p>
                </div>
              </div>
              <span className="text-xs px-2 py-1 rounded" style={{ background:"rgba(255,255,255,0.15)", color:"#7DD3D0" }}>
                {scored.length} van {apps.length} apps in kaart
              </span>
            </div>
          </div>

          {/* Uitleg */}
          <div className="rounded p-4 mb-4" style={{ background:"#fff", border:"1px solid #D0E4F7" }}>
            <h3 className="font-bold text-sm mb-3" style={{ color:"#0C2340" }}>Hoe lees je deze kaart?</h3>
            <div className="grid grid-cols-2 gap-4">
              <div>
                <p className="text-xs leading-relaxed mb-3" style={{ color:"#374151" }}>
                  De kaart toont twee geopolitieke posities per applicatie, gebaseerd op de DAAF-scores
                  A1 (jurisdictie leverancier) en A3 (hosting en datalocatie):
                </p>
                <div className="space-y-2">
                  <div className="flex gap-2 items-start">
                    <div className="w-4 h-4 rounded-full flex-shrink-0 mt-0.5"
                      style={{ background:"#1A56A0", border:"2px solid white", boxShadow:"0 0 0 2px #1A56A0" }}/>
                    <p className="text-xs" style={{ color:"#374151" }}>
                      <strong style={{ color:"#1A56A0" }}>Gevulde cirkel</strong> = Jurisdictie leverancier (A1): onder welk rechtssysteem valt de leverancier?
                    </p>
                  </div>
                  <div className="flex gap-2 items-start">
                    <div className="w-4 h-4 rounded flex-shrink-0 mt-0.5"
                      style={{ background:"white", border:"2.5px solid #1A56A0" }}/>
                    <p className="text-xs" style={{ color:"#374151" }}>
                      <strong style={{ color:"#1A56A0" }}>Omrand vierkant</strong> = Datalocatie (A3): waar staat de data fysiek opgeslagen?
                    </p>
                  </div>
                </div>
              </div>
              <div>
                <p className="text-xs font-semibold mb-2" style={{ color:"#374151" }}>Kleurschaal:</p>
                <div className="space-y-1.5">
                  {[
                    { score:"1-2", label:"EU / EER",          kleur:"#16a34a", bg:"#dcfce7", tekst:"Volledig Europees. Laagste risico." },
                    { score:"3",   label:"Adequaat + risico",  kleur:"#ca8a04", bg:"#fef9c3", tekst:"VS met adequaatheidsbesluit. CLOUD Act van toepassing." },
                    { score:"4",   label:"SCCs, geen adequaat",kleur:"#ea580c", bg:"#ffedd5", tekst:"Contractuele waarborgen. Verhoogd risico." },
                    { score:"5",   label:"Geen waarborgen",    kleur:"#dc2626", bg:"#fee2e2", tekst:"Buiten EU, geen waarborgen." },
                    { score:"0",   label:"Niet ingevuld",      kleur:"#9ca3af", bg:"#f3f4f6", tekst:"Score nog niet ingevuld." },
                  ].map(function(r) {
                    return (
                      <div key={r.score} className="flex items-center gap-2">
                        <span className="w-2 h-2 rounded-full flex-shrink-0" style={{ background:r.kleur }}/>
                        <span className="text-xs px-1.5 py-0.5 rounded font-semibold flex-shrink-0"
                          style={{ background:r.bg, color:r.kleur }}>{r.score}</span>
                        <span className="text-xs font-semibold flex-shrink-0" style={{ color:r.kleur }}>{r.label}</span>
                        <span className="text-xs" style={{ color:"#9ca3af" }}>{r.tekst}</span>
                      </div>
                    );
                  })}
                </div>
              </div>
            </div>
          </div>

          {/* Applicatie-filter */}
          {apps.length > 0 && (
            <div className="rounded p-3 mb-4" style={{ background:"#fff", border:"1px solid #D0E4F7" }}>
              <div className="flex items-center justify-between mb-2">
                <div className="flex items-center gap-2">
                  <p className="text-xs font-semibold" style={{ color:"#0C2340" }}>Kaart-selectie</p>
                  <span className="text-xs px-2 py-0.5 rounded font-semibold"
                    style={{ background: geoHidden.size === 0 ? "#f0f9f9" : "#fffbeb",
                             color: geoHidden.size === 0 ? "#0f766e" : "#92400e" }}>
                    {apps.length - geoHidden.size} van {apps.length} zichtbaar
                  </span>
                </div>
                <div className="flex gap-2">
                  {geoHidden.size > 0 && (
                    <button onClick={() => setGeoHidden(new Set())}
                      className="text-xs px-2.5 py-1 font-medium"
                      style={{ border:"1px solid #D0E4F7", borderRadius:3, color:"#1A56A0", background:"#EBF3FF" }}>
                      Alles tonen
                    </button>
                  )}
                </div>
              </div>
              <div className="flex flex-wrap gap-1.5">
                {[...apps].sort((a, b) => displayName(a).localeCompare(displayName(b), "nl", { sensitivity:"base", numeric:true })).map(a => {
                  const sc = calcScores(a.scores || {});
                  const col = scoreColor(sc.autonomyScore);
                  const hidden = geoHidden.has(a.id);
                  return (
                    <button key={a.id}
                      onClick={() => setGeoHidden(p => {
                        const n = new Set(p);
                        n.has(a.id) ? n.delete(a.id) : n.add(a.id);
                        return n;
                      })}
                      className="flex items-center gap-1.5 text-xs px-2.5 py-1.5 font-medium transition-all"
                      style={{
                        borderRadius:4,
                        border: hidden ? "2px solid #e5e7eb" : "2px solid " + col,
                        background: hidden ? "#f9fafb" : col + "18",
                        color: hidden ? "#9ca3af" : col,
                        textDecoration: hidden ? "line-through" : "none",
                      }}>
                      <span className="w-2 h-2 rounded-full flex-shrink-0"
                        style={{ background: hidden ? "#d1d5db" : col }}/>
                      {displayName(a).substring(0,22)}
                      <span style={{ fontSize:10, marginLeft:2, opacity:0.7 }}>{hidden ? "＋" : "✕"}</span>
                    </button>
                  );
                })}
              </div>
            </div>
          )}

          {/* Kaart */}
          {apps.length === 0 ? (
            <div className="rounded p-10 text-center" style={{ background:"#fff", border:"2px dashed #D0E4F7", color:"#9ca3af" }}>
              Nog geen applicaties. Voeg applicaties toe en vul A1 en A3 scores in.
            </div>
          ) : (
            <div className="rounded mb-4" style={{ background:"#fff", border:"1px solid #D0E4F7", overflow:"hidden" }}>
              <div className="px-4 pt-4 pb-2">
                <h3 className="font-bold text-sm" style={{ color:"#0C2340" }}>Wereldkaart — jurisdictie en datalocatie per applicatie</h3>
                <p className="text-xs mt-0.5" style={{ color:"#9ca3af" }}>
                  Gebaseerd op DAAF-scores A1 en A3 · scroll om in te zoomen · sleep om te verschuiven
                </p>
              </div>
              <WorldMapD3
                scored={scored}
                jurisGroups={jurisGroups}
                dataGroups={dataGroups}
                geoHoverId={geoHoverId}
                setGeoHoverId={setGeoHoverId}
                geoTooltip={geoTooltip}
                setGeoTooltip={setGeoTooltip}
                risicoKleur={risicoKleur}
                displayName={displayName}
                REGIO_COORDS={REGIO_COORDS}
              />
            </div>
          )}

          {/* Tabel */}
          {scored.length > 0 && (
            <div className="rounded p-4 mb-4" style={{ background:"#fff", border:"1px solid #D0E4F7" }}>
              <h3 className="font-bold text-sm mb-3" style={{ color:"#0C2340" }}>Overzicht per applicatie — jurisdictie en datalocatie</h3>
              <table style={{ width:"100%", borderCollapse:"collapse", fontSize:11 }}>
                <thead>
                  <tr style={{ background:"#0C2340", color:"white" }}>
                    {["Applicatie","Leverancier","A1 Jurisdictie","A3 Datalocatie","Oordeel"].map(function(h) {
                      return <th key={h} style={{ padding:"7px 10px", textAlign:"left", fontSize:10 }}>{h}</th>;
                    })}
                  </tr>
                </thead>
                <tbody>
                  {visibleSortedApps.map(function(a, i) {
                    const a1 = a.scores["A1"] || 0;
                    const a3 = a.scores["A3"] || 0;
                    const maxScore = Math.max(a1, a3);
                    const oordeel = maxScore === 0 ? { t:"Niet beoordeeld", fg:"#9ca3af", bg:"#f3f4f6" }
                      : maxScore <= 2 ? { t:"EU-conform",  fg:"#15803d", bg:"#dcfce7" }
                      : maxScore === 3 ? { t:"Acceptabel",  fg:"#a16207", bg:"#fef9c3" }
                      : maxScore === 4 ? { t:"Aandacht",    fg:"#c2410c", bg:"#ffedd5" }
                      : { t:"Kritiek", fg:"#b91c1c", bg:"#fee2e2" };
                    return (
                      <tr key={a.id} style={{ background: i%2===0 ? "#f8fafc" : "white", borderBottom:"1px solid #f1f5f9" }}>
                        <td style={{ padding:"8px 10px", fontWeight:600, color:"#0C2340" }}>{displayName(a)}</td>
                        <td style={{ padding:"8px 10px", color:"#6b7280" }}>{a.supplier || "–"}</td>
                        <td style={{ padding:"8px 10px" }}>
                          <div className="flex items-center gap-1.5">
                            <span style={{ width:10, height:10, borderRadius:"50%", background:risicoKleur(a1), display:"inline-block", flexShrink:0 }}/>
                            <span style={{ color:"#374151" }}>{a1lbl[a1] || "–"}</span>
                            {a1 > 0 && <span style={{ fontSize:10, fontWeight:700, color:risicoKleur(a1) }}>({a1}/5)</span>}
                          </div>
                        </td>
                        <td style={{ padding:"8px 10px" }}>
                          <div className="flex items-center gap-1.5">
                            <span style={{ width:10, height:10, borderRadius:2, background:"white", border:"2px solid " + risicoKleur(a3), display:"inline-block", flexShrink:0 }}/>
                            <span style={{ color:"#374151" }}>{a3lbl[a3] || "–"}</span>
                            {a3 > 0 && <span style={{ fontSize:10, fontWeight:700, color:risicoKleur(a3) }}>({a3}/5)</span>}
                          </div>
                        </td>
                        <td style={{ padding:"8px 10px" }}>
                          <span style={{ fontSize:10, fontWeight:600, padding:"2px 8px", borderRadius:3, background:oordeel.bg, color:oordeel.fg }}>{oordeel.t}</span>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          )}

          {/* Niet ingevuld */}
          {incomplete.length > 0 && (
            <div className="rounded p-3 mb-4 text-xs" style={{ background:"#fffbeb", border:"1px solid #fde68a", color:"#92400e" }}>
              <strong>{incomplete.length} applicatie{incomplete.length !== 1 ? "s" : ""} nog niet beoordeeld op locatie:</strong>
              {" "}{incomplete.map(function(a) { return displayName(a); }).join(", ")}. Vul scores A1 en A3 in via het assessment.
            </div>
          )}

          {/* Footer */}
          <div className="rounded p-3 text-center" style={{ background:"#fff", border:"1px solid #D0E4F7" }}>
            <p className="text-xs" style={{ color:"#9ca3af" }}>
              NHL Stenden Hogeschool · Portfolioanalyse Digitale Soevereiniteit · {VERSION} · {vandaag}
            </p>
          </div>

        </div>
      </div>
    );
  }

  function Compare() {
    const minCompare = 2;
    const visibleCompare = apps.filter(a => !compareHidden.has(a.id));

    function toggleCompare(id) {
      setCompareHidden(prev => {
        const next = new Set(prev);
        const currentVisible = apps.filter(a => !next.has(a.id)).length;
        if (next.has(id)) {
          if (currentVisible >= MAX_COMPARE) return prev; // max bewaken
          next.delete(id);
          return next;
        }
        if (currentVisible <= minCompare) return prev; // min bewaken
        next.add(id);
        return next;
      });
    }

    function selectLaatste5() {
      setCompareHidden(new Set(apps.slice(0, Math.max(0, apps.length - MAX_COMPARE)).map(a => a.id)));
    }

    if (visibleCompare.length < 2) return (
      <div className="h-full flex items-center justify-center text-gray-400">
        <div className="text-center">
          <p className="text-lg mb-1">Selecteer minimaal 2 applicaties</p>
          <p className="text-sm">via de filter hieronder om te vergelijken.</p>
        </div>
      </div>
    );

    const name14 = n => n.substring(0, 14);

    // Radar met gewogen dimensiescores
    const dimLettersC = [...new Set(DAAF.map(d => d.dim))];
    const dimLabelC   = letter => { const f = DAAF.find(d => d.dim===letter); return f ? f.dimName.substring(0,12) : letter; };
    const dimScoreC   = (a, letter) => {
      if (letter === "A") {
        const a1 = a.scores["A1"]||0, a3 = a.scores["A3"]||0;
        const pairs = [[a1,3],[a3,2]].filter(([v])=>v>0);
        if (!pairs.length) return 0;
        const tw = pairs.reduce((s,[,w])=>s+w,0);
        return pairs.reduce((s,[v,w])=>s+v*w,0)/tw;
      }
      const qs = DAAF.filter(d=>d.dim===letter);
      const vals = qs.map(q=>a.scores[q.key]||0).filter(v=>v>0);
      return vals.length ? vals.reduce((s,v)=>s+v,0)/vals.length : 0;
    };
    const radarData = dimLettersC.map(letter => {
      const entry = { dim: dimLabelC(letter) };
      visibleCompare.forEach(a => { entry[name14(displayName(a))] = +dimScoreC(a, letter).toFixed(2); });
      return entry;
    });

    const barData = visibleCompare.map(a => {
      const s = calcScores(a.scores);
      return {
        name: displayName(a).substring(0, 16),
        Autonomiescore: s.autonomyScore ? +s.autonomyScore.toFixed(1) : 0,
        "DICTU x2": s.dictuAvg ? +(s.dictuAvg * 2).toFixed(1) : 0
      };
    });

    return (
      <div className="h-full overflow-y-auto" style={{ background:"#EBF3FF" }}>
        <div className="p-5 max-w-5xl mx-auto">

          {/* Filter strip */}
          {apps.length > 2 && (
            <div className="rounded p-3 mb-4 flex items-center gap-3 flex-wrap"
              style={{ background:"#fff", border:"1px solid #D0E4F7" }}>
              <div className="flex items-center gap-2 flex-shrink-0 flex-wrap">
                <span style={{ fontSize:11, fontWeight:600, color:"#0C2340" }}>Grafiek-selectie</span>
                <span className="text-xs px-2 py-0.5 rounded font-semibold"
                  style={{ background: visibleCompare.length >= MAX_COMPARE ? "#fffbeb" : "#f0f9f9",
                           color: visibleCompare.length >= MAX_COMPARE ? "#92400e" : "#0f766e" }}>
                  {visibleCompare.length} / {MAX_COMPARE} geselecteerd
                </span>
                {apps.length > MAX_COMPARE && (
                  <span style={{ fontSize:10, color:"#9ca3af" }}>max {MAX_COMPARE} voor leesbaarheid</span>
                )}
              </div>
              <div className="flex gap-2 flex-wrap flex-1">
                {apps.map((a, i) => {
                  const hidden  = compareHidden.has(a.id);
                  const isLast2 = !hidden && visibleCompare.length <= minCompare;
                  const col     = ["#1e40af","#7c3aed","#065f46","#92400e","#991b1b","#0f766e"][i % 6];
                  return (
                    <button key={a.id} onClick={() => toggleCompare(a.id)}
                      disabled={(isLast2 && !hidden) || (!hidden && visibleCompare.length >= MAX_COMPARE && false)}
                      title={
                        isLast2 && !hidden ? "Minimaal 2 applicaties voor vergelijking" :
                        hidden && visibleCompare.length >= MAX_COMPARE ? `Maximum van ${MAX_COMPARE} geselecteerd` : ""
                      }
                      className="flex items-center gap-1.5 text-xs px-2.5 py-1.5 font-medium"
                      style={{
                        borderRadius:4, border:`2px solid ${hidden?"#e5e7eb":col}`,
                        background: hidden?"#f9fafb":`${col}18`, color:hidden?"#9ca3af":col,
                        opacity: (isLast2&&!hidden) || (hidden && visibleCompare.length >= MAX_COMPARE) ? 0.4 : 1,
                        cursor: (isLast2&&!hidden) || (hidden && visibleCompare.length >= MAX_COMPARE) ? "not-allowed" : "pointer",
                        textDecoration:hidden?"line-through":"none"
                      }}>
                      <span className="w-2 h-2 rounded-full flex-shrink-0" style={{ background:hidden?"#d1d5db":col }}/>
                      {displayName(a).substring(0,20)} {hidden?"＋":"✕"}
                    </button>
                  );
                })}
              </div>
              {apps.length > MAX_COMPARE && (
                <button onClick={selectLaatste5}
                  className="text-xs px-2.5 py-1.5 flex-shrink-0 font-medium"
                  style={{ borderRadius:4, background:"#EBF3FF", color:"#1A56A0", border:"1px solid #D0E4F7" }}>
                  Laatste {MAX_COMPARE}
                </button>
              )}
              {compareHidden.size > 0 && apps.length <= MAX_COMPARE && (
                <button onClick={() => setCompareHidden(new Set())}
                  className="text-xs px-2.5 py-1.5 flex-shrink-0 font-medium"
                  style={{ borderRadius:4, background:"#EBF3FF", color:"#1A56A0", border:"1px solid #D0E4F7" }}>
                  Alles tonen
                </button>
              )}
            </div>
          )}

          {/* Leeswijzer */}
          <div className="rounded p-4 mb-4" style={{ background:"#fff", border:"1px solid #D0E4F7" }}>
            <h2 className="font-bold text-sm mb-3" style={{ color:"#0C2340" }}>Leeswijzer vergelijking</h2>
            <div className="grid grid-cols-3 gap-3">
              {[
                { badge:"DAAF", bg:"#1A56A0", title:"Autonomiescore (1-10)",
                  body:"Toont hoe urgent het autonomieprobleem is. Berekend via Mitigatie / (Risico x Belang). Hoger = minder urgent. De kleur geeft de status aan.",
                  statuses:[{t:">=7 Goed",sb:"#dcfce7",sf:"#15803d"},{t:"5-7 Acceptabel",sb:"#fef9c3",sf:"#a16207"},{t:"3-5 Zorgwekkend",sb:"#ffedd5",sf:"#c2410c"},{t:"<3 Kritiek",sb:"#fee2e2",sf:"#b91c1c"}]
                },
                { badge:"DAAF", bg:"#1A56A0", title:"Risico / Mitigatie / Belang",
                  body:"Drie DAAF-niveauscores (1-5). Risico (rood): laag is beter. Mitigatie (teal): hoog is beter. Belang (oranje): laag = minder urgent. Zijn gewogen gemiddelden van de ingevulde dimensies." },
                { badge:"DICTU", bg:"#26B5AE", title:"DICTU-score (1-5)",
                  body:"Gemiddelde van 4 vragen: 2.1 Dataresidency, 2.2 Technische beveiliging, 2.3 Juridische bescherming, 4.1 EU-infrastructuur. Score 1 = volledig afhankelijk, 5 = maximaal soeverein." }
              ].map(c => (
                <div key={c.title} className="rounded p-3" style={{ background:"#f8fafc", border:"1px solid #e5e7eb" }}>
                  <div className="flex items-center gap-2 mb-2">
                    <span className="text-xs font-bold px-1.5 py-0.5" style={{ background:c.bg, color:"#fff", borderRadius:3 }}>{c.badge}</span>
                    <span className="text-xs font-semibold" style={{ color:"#0C2340" }}>{c.title}</span>
                  </div>
                  <p className="text-xs leading-relaxed mb-2" style={{ color:"#6b7280" }}>{c.body}</p>
                  {c.statuses && (
                    <div className="flex gap-1 flex-wrap">
                      {c.statuses.map(s=>(
                        <span key={s.t} style={{ fontSize:9, background:s.sb, color:s.sf, borderRadius:2, padding:"1px 5px", fontWeight:600 }}>{s.t}</span>
                      ))}
                    </div>
                  )}
                </div>
              ))}
            </div>
          </div>

          <div className="grid grid-cols-2 gap-4 mb-4">
            <div className="rounded p-4" style={{ background:"#fff", border:"1px solid #D0E4F7" }}>
              <div className="flex items-center gap-2 mb-3">
                <span className="text-xs font-bold px-2 py-0.5" style={{ background:"#1A56A0", color:"#fff", borderRadius:3 }}>DAAF + DICTU</span>
                <p className="text-sm font-semibold" style={{ color:"#0C2340" }}>Scores per applicatie</p>
              </div>
              <ResponsiveContainer width="100%" height={280}>
                <BarChart data={barData} margin={{ top:5, right:10, bottom:80, left:0 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#f3f4f6" />
                  <XAxis dataKey="name" tick={{ fontSize:8 }} angle={-40} textAnchor="end" interval={0} height={75} />
                  <YAxis domain={[0,10]} tick={{ fontSize:9 }} width={24} />
                  <Tooltip wrapperStyle={{ fontSize:11 }} />
                  <Bar dataKey="Autonomiescore" radius={[3,3,0,0]} name="Autonomiescore (1-10)">
                    {barData.map((d,i) => <Cell key={i} fill={scoreColor(d.Autonomiescore)} />)}
                  </Bar>
                  <Bar dataKey="DICTU x2" fill="#26B5AE" fillOpacity={0.7} radius={[3,3,0,0]} name="DICTU x2 (schaal 0-10)" />
                </BarChart>
              </ResponsiveContainer>
              {/* Legenda buiten de grafiek */}
              <div className="flex gap-4 mt-2 justify-center">
                <div className="flex items-center gap-1.5">
                  <div style={{ width:12, height:12, background:"#ca8a04", borderRadius:2 }}/>
                  <span style={{ fontSize:10, color:"#6b7280" }}>Autonomiescore (1–10)</span>
                </div>
                <div className="flex items-center gap-1.5">
                  <div style={{ width:12, height:12, background:"#26B5AE", borderRadius:2, opacity:0.7 }}/>
                  <span style={{ fontSize:10, color:"#6b7280" }}>DICTU-score (schaal 0–10)</span>
                </div>
              </div>
            </div>

            <div className="rounded p-4" style={{ background:"#fff", border:"1px solid #D0E4F7" }}>
              <div className="flex items-center gap-2 mb-1">
                <span className="text-xs font-bold px-2 py-0.5" style={{ background:"#1A56A0", color:"#fff", borderRadius:3 }}>DAAF</span>
                <p className="text-sm font-semibold" style={{ color:"#0C2340" }}>Dimensieprofiel — vergelijking</p>
              </div>
              <p className="text-xs mb-2" style={{ color:"#9ca3af" }}>Gewogen scores 1–5 per dimensie. Kleurovergang toont de richting: groen = gewenste kant. Hover voor details.</p>
              <DivergingChart apps={visibleCompare} compact={true} useSecondaryName={useSecondaryName} />
              <div className="grid grid-cols-1 gap-1 mt-2">
                {[
                  "Risico-assen (A, B): groen links — punt dicht bij 1 is goed.",
                  "Mitigatie-assen (C, D, E): groen rechts — punt dicht bij 5 is goed.",
                  "Belang-assen (F, G, H): hoge score = meer urgentie, maar vraagt ook sterkere mitigatie.",
                ].map((t,i) => <p key={i} style={{ fontSize:9, color:"#9ca3af" }}>• {t}</p>)}
              </div>
            </div>
          </div>

          {/* Vergelijkingstabel */}
          <div className="rounded p-4 overflow-x-auto" style={{ background:"#fff", border:"1px solid #D0E4F7" }}>
            <div className="flex items-center gap-2 mb-3">
              <h3 className="font-semibold text-sm" style={{ color:"#0C2340" }}>Vergelijkingstabel</h3>
              <span className="text-xs text-gray-400">klik op een rij om naar het assessment te gaan</span>
            </div>
            <table className="w-full" style={{ fontSize:12 }}>
              <thead>
                <tr style={{ borderBottom:"2px solid #1A56A0" }}>
                  {[
                    "Applicatie","Leverancier",
                    "Autonomiescore (1-10)","Risico (1-5) ↓","Mitigatie (1-5) ↑","Belang (1-5) ↓",
                    "DICTU (1-5)","Volledigheid"
                  ].map(h => (
                    <th key={h} className="text-left py-2 px-2 font-semibold" style={{ color:"#1A56A0", fontSize:11 }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {visibleCompare.map(a => {
                  const s = calcScores(a.scores);
                  const lbl = scoreLabel(s.autonomyScore);
                  return (
                    <tr key={a.id} style={{ borderBottom:"1px solid #EBF3FF", cursor:"pointer" }}
                      onClick={() => { setSelId(a.id); setStep(0); setAssessReadOnly(false); setView("assess"); }}>
                      <td className="py-2 px-2 font-medium" style={{ color:"#0C2340" }}>{displayName(a)}</td>
                      <td className="py-2 px-2 text-gray-500">{a.supplier}</td>
                      <td className="py-2 px-2">
                        <span className="px-2 py-0.5 font-semibold text-xs" style={{ borderRadius:3, background:lbl.bg, color:lbl.fg }}>
                          {s.autonomyScore ? s.autonomyScore.toFixed(1) : "--"}
                        </span>
                      </td>
                      <td className="py-2 px-2 font-semibold" style={{ color:"#dc2626" }}>{s.risico    ? s.risico.toFixed(2)    : "--"}</td>
                      <td className="py-2 px-2 font-semibold" style={{ color:"#26B5AE" }}>{s.mitigatie ? s.mitigatie.toFixed(2) : "--"}</td>
                      <td className="py-2 px-2 font-semibold" style={{ color:"#E87722" }}>{s.belang    ? s.belang.toFixed(2)    : "--"}</td>
                      <td className="py-2 px-2 text-gray-600">{s.dictuAvg  ? s.dictuAvg.toFixed(1)+"/5"  : "--"}</td>
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
            <div className="flex items-center gap-2">
              {apps.some(a => !a.nameSecondary) && (
                <button onClick={() => {
                  let counter = 1;
                  setApps(p => p.map(a => ({
                    ...a,
                    nameSecondary: a.nameSecondary || `Dummie ${counter++}`
                  })));
                }}
                className="text-xs px-3 py-1.5 font-medium"
                style={{ border:"1px solid #fde68a", color:"#92400e", background:"#fffbeb", borderRadius:4 }}
                title="Wijs 'Dummie 1, 2...' toe als secundaire naam voor apps zonder secundaire naam">
                🏷 Dummie namen toewijzen
                </button>
              )}
              {/* Verborgen file input voor import */}
              <input type="file" accept=".json" id="db-import-input"
                style={{ display:"none" }} onChange={handleImportFile}/>

              <button onClick={exportDatabase} disabled={apps.length === 0}
                className="text-xs px-3 py-1.5 font-medium"
                style={{ border:"1px solid #86efac", borderRadius:4, color:"#15803d", background:"#f0fdf4",
                         opacity: apps.length === 0 ? 0.5 : 1 }}>
                ⬇ Exporteer database
              </button>
              <button onClick={() => document.getElementById("db-import-input").click()}
                className="text-xs px-3 py-1.5 font-medium"
                style={{ border:"1px solid #D0E4F7", borderRadius:4, color:"#1A56A0", background:"#EBF3FF" }}>
                ⬆ Importeer database
              </button>

              <button onClick={() => { setAdminUnlocked(false); setView("dashboard"); }}
                className="text-xs px-3 py-1.5 font-medium"
                style={{ border:"1px solid #D0E4F7", borderRadius:4, color:"#6b7280", background:"#fff" }}>
                🔒 Vergrendelen
              </button>
            </div>
          </div>

          {apps.length === 0 ? (
            <div className="text-center py-16 bg-white rounded" style={{ border:"2px dashed #D0E4F7", color:"#9ca3af" }}>
              Nog geen applicaties in het systeem.
            </div>
          ) : (
            <div className="space-y-3">
              {[...apps].sort((a, b) => {
                const na = (useSecondaryName && a.nameSecondary ? a.nameSecondary : a.name).toLowerCase();
                const nb = (useSecondaryName && b.nameSecondary ? b.nameSecondary : b.name).toLowerCase();
                return na.localeCompare(nb, "nl", { sensitivity:"base", numeric:true });
              }).map(a => {
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
                          {a.nameSecondary && (
                            <span className="text-xs px-2 py-0.5 font-medium"
                              style={{ background:"#fffbeb", border:"1px solid #fde68a", color:"#92400e", borderRadius:3 }}>
                              🏷 {a.nameSecondary}
                            </span>
                          )}
                          {a.supplier && <span className="text-xs text-gray-400">· {a.supplier}</span>}
                          {a.cat && <span className="text-xs px-2 py-0.5 rounded" style={{ background:"#EBF3FF", color:"#1A56A0" }}>{a.cat}</span>}
                        </div>
                        <div className="flex items-center gap-3 mt-1 flex-wrap">
                          {a.owner && <span className="text-xs text-gray-400">👤 {a.owner}</span>}
                          <span className="text-xs px-2 py-0.5 font-medium" style={{ borderRadius:3, background:lbl.bg, color:lbl.fg }}>{lbl.text}</span>
                          <span className="text-xs text-gray-400">{sc.completeness}% ingevuld · {filled.length}/{allQ.length} vragen</span>
                          <span className="text-xs text-gray-400">Aangemaakt: {new Date(a.createdAt).toLocaleDateString("nl-NL")}</span>
                        </div>
                        {a.appNotes && <p className="text-xs text-gray-400 mt-1 italic">"{adaptNote(a.appNotes, a)}"</p>}
                      </div>
                      {/* Action buttons */}
                      <div className="flex gap-2 flex-shrink-0">
                        <button
                          onClick={() => { setSelId(a.id); setStep(0); setAssessReadOnly(false); setView("assess"); }}
                          className="text-white text-xs px-3 py-1.5 font-medium"
                          style={{ background:"#1A56A0", borderRadius:4 }}>
                          ✏️ Invullen
                        </button>
                        <button
                          onClick={() => {
                            setEditAppId(a.id);
                            setEditForm({ name:a.name, nameSecondary:a.nameSecondary||"", cat:a.cat, supplier:a.supplier, owner:a.owner, appNotes:a.appNotes });
                          }}
                          className="text-xs px-3 py-1.5 font-medium"
                          style={{ border:"1px solid #D0E4F7", borderRadius:4, color:"#1A56A0", background:"#fff" }}>
                          ⚙️ Gegevens
                        </button>
                        <button
                          onClick={() => {
                            if (confirm(`"${displayName(a)}" definitief verwijderen? Dit kan niet ongedaan worden gemaakt.`)) {
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
                      <div className="flex flex-wrap gap-1.5 mb-2">
                        {allQ.map(q => {
                          const s = a.scores[q.key] || 0;
                          const n = (a.notes || {})[q.key] || "";
                          return (
                            <div key={q.key} title={`${q.key}: ${q.name}\nScore: ${s || "niet ingevuld"}${n ? `\nMotivatie: ${n}` : ""}`}
                              className="flex items-center gap-1 px-2 py-1 text-xs"
                              style={{ borderRadius:3, background: s ? "#EBF3FF" : "#f9fafb", border:`1px solid ${n ? "#86efac" : "#D0E4F7"}`, color:"#0C2340" }}>
                              <span style={{ color:"#1A56A0", fontWeight:600 }}>{q.key}</span>
                              <span style={{ fontWeight:700, color: s ? scoreColor(s, 5) : "#d1d5db" }}>{s || "–"}</span>
                              {n && <span style={{ color:"#16a34a", fontSize:9 }}>💬</span>}
                            </div>
                          );
                        })}
                      </div>
                      {/* Motivaties tonen als die er zijn */}
                      {allQ.some(q => (a.notes || {})[q.key]) && (
                        <div className="mt-2 space-y-1.5">
                          <p className="text-xs font-semibold" style={{ color:"#16a34a" }}>💬 Motivaties</p>
                          {allQ.filter(q => (a.notes || {})[q.key]).map(q => (
                            <div key={q.key} className="rounded px-3 py-2"
                              style={{ background:"#f0fdf4", border:"1px solid #86efac" }}>
                              <span className="text-xs font-semibold" style={{ color:"#0C2340" }}>{q.key} — {q.name}: </span>
                              <span className="text-xs" style={{ color:"#374151" }}>{adaptNote((a.notes || {})[q.key], a)}</span>
                            </div>
                          ))}
                        </div>
                      )}
                    </div>
                  </div>
                );
              })}
            </div>
          )}

          {/* Danger zone verwijderd — data wissen niet beschikbaar via de interface */}

          {/* Bevestigingsmodal alles wissen */}
          {showDeleteAll && (
            <div className="fixed inset-0 flex items-center justify-center z-50"
              style={{ background:"rgba(12,35,64,0.7)" }}
              onClick={() => setShowDeleteAll(false)}>
              <div className="bg-white rounded-lg mx-4 overflow-hidden"
                style={{ width:"100%", maxWidth:420, boxShadow:"0 8px 32px rgba(220,38,38,0.3)", border:"2px solid #dc2626" }}
                onClick={e => e.stopPropagation()}>
                {/* Header */}
                <div style={{ background:"#dc2626", padding:"14px 20px" }}>
                  <div className="flex items-center gap-3">
                    <span style={{ fontSize:22 }}>⚠️</span>
                    <div>
                      <h2 className="font-bold text-white text-sm">Alle data permanent verwijderen</h2>
                      <p style={{ fontSize:10, color:"#fca5a5", marginTop:2 }}>
                        Deze actie kan niet ongedaan worden gemaakt
                      </p>
                    </div>
                  </div>
                </div>
                {/* Inhoud */}
                <div style={{ padding:"20px 24px" }}>
                  <p className="text-sm mb-4" style={{ color:"#374151", lineHeight:1.6 }}>
                    Dit wist <strong>alle {apps.length} applicaties</strong> inclusief alle ingevulde assessments,
                    scores en motivaties. Er is geen herstel mogelijk.
                  </p>
                  <div className="rounded p-3 mb-4" style={{ background:"#fff5f5", border:"1px solid #fecaca" }}>
                    <p className="text-xs mb-2 font-semibold" style={{ color:"#dc2626" }}>
                      Typ <span style={{ fontFamily:"monospace", background:"#fee2e2", padding:"1px 6px", borderRadius:3 }}>VERWIJDER ALLES</span> om te bevestigen:
                    </p>
                    <input
                      autoFocus
                      value={deleteConfirmText}
                      onChange={e => setDeleteConfirmText(e.target.value)}
                      onPaste={e => e.preventDefault()}
                      placeholder="VERWIJDER ALLES"
                      className="w-full px-3 py-2 text-sm font-mono"
                      style={{
                        border: `2px solid ${deleteConfirmText === "VERWIJDER ALLES" ? "#22c55e" : "#fecaca"}`,
                        borderRadius:4, outline:"none",
                        background: deleteConfirmText === "VERWIJDER ALLES" ? "#f0fdf4" : "white",
                        color:"#0C2340", letterSpacing:"0.05em"
                      }}
                      onFocus={e => e.target.style.borderColor = deleteConfirmText === "VERWIJDER ALLES" ? "#22c55e" : "#dc2626"}
                    />
                    {deleteConfirmText.length > 0 && deleteConfirmText !== "VERWIJDER ALLES" && (
                      <p style={{ fontSize:10, color:"#dc2626", marginTop:4 }}>
                        Tekst klopt niet — typ exact: VERWIJDER ALLES
                      </p>
                    )}
                    {deleteConfirmText === "VERWIJDER ALLES" && (
                      <p style={{ fontSize:10, color:"#16a34a", marginTop:4 }}>✓ Bevestigingstekst correct</p>
                    )}
                  </div>
                  {/* Knoppen */}
                  <div className="flex gap-3">
                    <button
                      onClick={() => { setShowDeleteAll(false); setDeleteConfirmText(""); }}
                      className="flex-1 py-2 text-sm font-medium"
                      style={{ border:"1px solid #D0E4F7", borderRadius:4, color:"#6b7280", background:"white" }}>
                      Annuleren
                    </button>
                    <button
                      disabled={deleteConfirmText !== "VERWIJDER ALLES"}
                      onClick={() => {
                        setApps([]);
                        setSelId(null);
                        setShowDeleteAll(false);
                        setDeleteConfirmText("");
                      }}
                      className="flex-1 py-2 text-sm font-bold text-white"
                      style={{
                        borderRadius:4,
                        background: deleteConfirmText === "VERWIJDER ALLES" ? "#dc2626" : "#fca5a5",
                        cursor: deleteConfirmText === "VERWIJDER ALLES" ? "pointer" : "not-allowed",
                        transition:"background 0.2s"
                      }}>
                      🗑 Definitief alles wissen
                    </button>
                  </div>
                </div>
              </div>
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
                    { k:"name",          l:"Primaire naam *",      hint:"De officiële applicatienaam (bijv. Microsoft Teams)" },
                    { k:"nameSecondary", l:"Secundaire naam",       hint:"Alternatieve weergavenaam (bijv. Samenwerkingsplatform)" },
                    { k:"supplier",      l:"Leverancier",           hint:"" },
                    { k:"cat",           l:"Categorie",             hint:"" },
                    { k:"owner",         l:"Applicatie-eigenaar",   hint:"" },
                  ].map(f => (
                    <div key={f.k}>
                      <label className="text-xs font-semibold block mb-0.5" style={{ color:"#0C2340" }}>{f.l}</label>
                      {f.hint && <p className="text-xs mb-1" style={{ color:"#9ca3af" }}>{f.hint}</p>}
                      <input
                        value={editForm[f.k] || ""}
                        onChange={e => setEditForm(p => ({ ...p, [f.k]: e.target.value }))}
                        className="w-full border px-3 py-2 text-sm focus:outline-none"
                        style={{ borderColor: f.k==="nameSecondary" ? "#D0E4F7" : "#D0E4F7", borderRadius:4,
                                 background: f.k==="nameSecondary" ? "#fffbeb" : "white" }}
                        onFocus={e => e.target.style.borderColor="#1A56A0"}
                        onBlur={e => e.target.style.borderColor="#D0E4F7"}
                      />
                    </div>
                  ))}
                  <div>
                    <label className="text-xs font-semibold block mb-1" style={{ color:"#0C2340" }}>Toelichting</label>
                    <textarea
                      value={editForm.appNotes || ""}
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
                        ? { ...a, name:editForm.name, nameSecondary:editForm.nameSecondary||"", cat:editForm.cat, supplier:editForm.supplier, owner:editForm.owner, appNotes:editForm.appNotes }
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


          {/* ── IMPORT MODAL ────────────────────────────────────────── */}
          {showImport && importData && (
            <div className="fixed inset-0 z-50 flex items-center justify-center"
              style={{ background:"rgba(12,35,64,0.6)" }}>
              <div className="bg-white rounded shadow-xl w-full max-w-2xl mx-4"
                style={{ border:"1px solid #D0E4F7", maxHeight:"85vh", display:"flex", flexDirection:"column" }}>

                {/* Modal header */}
                <div className="flex items-center justify-between px-5 py-4"
                  style={{ borderBottom:"1px solid #EBF3FF" }}>
                  <div>
                    <h3 className="font-bold text-sm" style={{ color:"#0C2340" }}>Database importeren</h3>
                    <p className="text-xs mt-0.5" style={{ color:"#9ca3af" }}>
                      Geëxporteerd op {new Date(importData.exportedAt).toLocaleString("nl-NL")} &nbsp;·&nbsp;
                      {importData.appCount} applicatie{importData.appCount !== 1 ? "s" : ""} in bestand &nbsp;·&nbsp;
                      versie {importData.version}
                    </p>
                  </div>
                  <button onClick={() => { setShowImport(false); setImportData(null); setImportSel(new Set()); }}
                    className="text-xs px-3 py-1.5" style={{ color:"#6b7280" }}>✕ Sluiten</button>
                </div>

                {/* Selectie-acties */}
                <div className="flex items-center gap-3 px-5 py-3"
                  style={{ borderBottom:"1px solid #EBF3FF", background:"#f8fafc" }}>
                  <span className="text-xs font-semibold" style={{ color:"#374151" }}>
                    {importSel.size} van {importData.apps.length} geselecteerd
                  </span>
                  <button onClick={() => setImportSel(new Set(importData.apps.map(a => a.id)))}
                    className="text-xs px-2.5 py-1 font-medium"
                    style={{ border:"1px solid #D0E4F7", borderRadius:3, color:"#1A56A0", background:"#EBF3FF" }}>
                    Alles selecteren
                  </button>
                  <button onClick={() => setImportSel(new Set())}
                    className="text-xs px-2.5 py-1 font-medium"
                    style={{ border:"1px solid #D0E4F7", borderRadius:3, color:"#6b7280", background:"#fff" }}>
                    Niets selecteren
                  </button>
                  <span className="text-xs" style={{ color:"#9ca3af", marginLeft:"auto" }}>
                    Bestaande applicaties met hetzelfde ID worden overschreven
                  </span>
                </div>

                {/* Applicatielijst */}
                <div className="overflow-y-auto flex-1 px-5 py-3 space-y-2">
                  {importData.apps.map(a => {
                    const sel = importSel.has(a.id);
                    const exists = apps.some(x => x.id === a.id);
                    const sc = calcScores(a.scores || {});
                    const lbl = scoreLabel(sc.autonomyScore);
                    return (
                      <div key={a.id}
                        onClick={() => setImportSel(p => {
                          const n = new Set(p);
                          n.has(a.id) ? n.delete(a.id) : n.add(a.id);
                          return n;
                        })}
                        className="flex items-center gap-3 rounded p-3 cursor-pointer transition-all"
                        style={{ border:"2px solid " + (sel ? "#1A56A0" : "#e5e7eb"),
                                 background: sel ? "#EBF3FF" : "#fff" }}>
                        <div className="w-5 h-5 rounded flex items-center justify-center flex-shrink-0"
                          style={{ background: sel ? "#1A56A0" : "#f3f4f6",
                                   border:"2px solid " + (sel ? "#1A56A0" : "#d1d5db") }}>
                          {sel && <span style={{ color:"white", fontSize:11, fontWeight:700 }}>✓</span>}
                        </div>
                        <div className="flex-1 min-w-0">
                          <div className="flex items-center gap-2 flex-wrap">
                            <span className="font-semibold text-sm" style={{ color:"#0C2340" }}>{a.name}</span>
                            {a.supplier && <span className="text-xs" style={{ color:"#9ca3af" }}>{a.supplier}</span>}
                            {sc.autonomyScore && (
                              <span className="text-xs px-2 py-0.5 font-semibold rounded"
                                style={{ background:lbl.bg, color:lbl.fg }}>{lbl.text}</span>
                            )}
                            {exists && (
                              <span className="text-xs px-2 py-0.5 rounded font-medium"
                                style={{ background:"#fffbeb", border:"1px solid #fde68a", color:"#92400e" }}>
                                overschrijft bestaande
                              </span>
                            )}
                            {!exists && (
                              <span className="text-xs px-2 py-0.5 rounded font-medium"
                                style={{ background:"#f0fdf4", border:"1px solid #86efac", color:"#15803d" }}>
                                nieuw
                              </span>
                            )}
                          </div>
                          <div className="flex gap-3 mt-1 flex-wrap">
                            {a.cat && <span className="text-xs" style={{ color:"#6b7280" }}>{a.cat}</span>}
                            {a.owner && <span className="text-xs" style={{ color:"#6b7280" }}>👤 {a.owner}</span>}
                            {sc.autonomyScore && (
                              <span className="text-xs font-semibold" style={{ color: scoreColor(sc.autonomyScore) }}>
                                Score: {sc.autonomyScore.toFixed(1)}/10
                              </span>
                            )}
                            <span className="text-xs" style={{ color:"#9ca3af" }}>{sc.completeness}% ingevuld</span>
                            <span className="text-xs" style={{ color:"#9ca3af" }}>
                              Aangemaakt: {new Date(a.createdAt).toLocaleDateString("nl-NL")}
                            </span>
                          </div>
                        </div>
                      </div>
                    );
                  })}
                </div>

                {/* Modal footer */}
                <div className="flex items-center justify-between px-5 py-4"
                  style={{ borderTop:"1px solid #EBF3FF" }}>
                  <p className="text-xs" style={{ color:"#9ca3af" }}>
                    Klik op een applicatie om te selecteren of deselecteren
                  </p>
                  <div className="flex gap-2">
                    <button onClick={() => { setShowImport(false); setImportData(null); setImportSel(new Set()); }}
                      className="text-xs px-4 py-2"
                      style={{ border:"1px solid #D0E4F7", borderRadius:4, color:"#6b7280" }}>
                      Annuleren
                    </button>
                    <button onClick={executeImport} disabled={importSel.size === 0}
                      className="text-xs px-4 py-2 font-semibold text-white"
                      style={{ background: importSel.size === 0 ? "#9ca3af" : "#15803d",
                               borderRadius:4, cursor: importSel.size === 0 ? "not-allowed" : "pointer" }}>
                      {importSel.size === 0 ? "Geen selectie"
                        : importSel.size + " applicatie" + (importSel.size !== 1 ? "s" : "") + " importeren"}
                    </button>
                  </div>
                </div>
              </div>
            </div>
          )}

      </div>
    );
  }

  // ── TRANSPARANTIE ─────────────────────────────────────────
  function Transparantie() {

    const STACK = [
      {
        naam: "Netlify",
        rol: "Hosting, serverless functies en opslag",
        type: "Hosting & opslag",
        jurisdictie: "Verenigde Staten",
        vestiging: "San Francisco, CA — VS",
        daafA1: 3,
        daafA1toe: "Adequaatheidsbesluit EU-VS Data Privacy Framework (2023) van kracht, maar CLOUD Act en FISA 702 zijn ongewijzigd. Buitenlandse overheidsinstanties houden potentieel toegang tot data op Amerikaanse servers.",
        daafA3: 3,
        daafA3toe: "Netlify Blobs (de opslag van deze applicatie) staat in één primaire regio. Netlify publiceert de exacte regio niet. Op basis van standaard Netlify-configuratie is dit vermoedelijk een Amerikaanse regio (AWS us-east-1 of vergelijkbaar). Geen EU-datalocatiegarantie.",
        beveiliging: "TLS 1.2+ en AES-256 voor data in transit en at rest. Gratis HTTPS via Let's Encrypt. GDPR-DPA beschikbaar. SOC 2 Type II gecertificeerd.",
        opmerking: "De assessment-data (scores, applicatienamen, motivaties) wordt opgeslagen in Netlify Blobs. Aanbeveling: exporteer regelmatig naar Excel als lokale back-up.",
        url: "https://www.netlify.com/security/",
      },
      {
        naam: "Anthropic / Claude API",
        rol: "AI-model voor tekstverwerking (niet actief in deze app)",
        type: "AI-dienst",
        jurisdictie: "Verenigde Staten",
        vestiging: "San Francisco, CA — VS",
        daafA1: 4,
        daafA1toe: "Anthropic is een Amerikaans bedrijf (PBC). Geen EU-adequaatheidsbesluit specifiek voor de API. Standaard contractuele clausules (SCCs) van toepassing. CLOUD Act is van toepassing op data verwerkt via de API.",
        daafA3: 4,
        daafA3toe: "API-verwerking vindt primair plaats in VS-datacenters. Anthropic heeft per juni 2026 een datacenter in Memphis (TN) operationeel. Geen EU-regiokeuze beschikbaar voor de standaard API.",
        beveiliging: "Data die via de API wordt verstuurd wordt versleuteld (TLS). Anthropic slaat geen API-prompts op voor modeltraining zonder expliciete toestemming.",
        opmerking: "De Claude API wordt in deze applicatie NIET gebruikt voor het verwerken van assessment-data. De app werkt volledig client-side. De Claude API is uitsluitend gebruikt bij de ontwikkeling (code genereren). Er is geen actieve API-verbinding tijdens gebruik.",
        url: "https://www.anthropic.com/privacy",
      },
      {
        naam: "GitHub",
        rol: "Versiebeheer en broncode-opslag",
        type: "Broncode",
        jurisdictie: "Verenigde Staten",
        vestiging: "San Francisco, CA — VS (dochter van Microsoft)",
        daafA1: 3,
        daafA1toe: "GitHub is een dochteronderneming van Microsoft (VS). EU-VS Data Privacy Framework van toepassing. CLOUD Act-risico aanwezig maar beperkt: de broncode bevat geen persoonsgegevens van gebruikers.",
        daafA3: 3,
        daafA3toe: "Broncode wordt opgeslagen op GitHub-servers in de VS. De opgeslagen data betreft uitsluitend applicatiecode, geen gebruikersdata of assessment-scores.",
        beveiliging: "HTTPS/TLS, tweefactorauthenticatie beschikbaar, repository kan privé worden ingesteld. Microsoft/GitHub SOC 2 gecertificeerd.",
        opmerking: "De repository supergorkum/stamboom (momenteel privé, te ovewegen: public als open source) bevat uitsluitend de broncode van de applicatie. Geen gebruikersdata.",
        url: "https://github.com/security",
      },
      {
        naam: "React / Vite (open source)",
        rol: "Frontend-framework en buildsysteem",
        type: "Open source software",
        jurisdictie: "Geen jurisdictie — open source licentie",
        vestiging: "React: Meta Platforms (VS) — open source onder MIT-licentie. Vite: community open source (MIT).",
        daafA1: 1,
        daafA1toe: "Open source software zonder leveranciersrelatie. Geen data-uitwisseling met de ontwikkelende organisatie. Code is publiek verifieerbaar.",
        daafA3: 1,
        daafA3toe: "Code wordt lokaal gebundeld tijdens de build en als statische bestanden geserveerd. Geen runtime-verbinding met externe servers van React of Vite.",
        beveiliging: "Publiek auditeerbare code. Actief onderhouden met regelmatige beveiligingsupdates. NPM-pakketbeheer met bekende kwetsbaarheidsscanners.",
        opmerking: "Geen afhankelijkheidsrisico richting een commercieel bedrijf. Vervanging door alternatief open source framework is technisch haalbaar.",
        url: "https://react.dev",
      },
    ];

    const daafKleur = (s) => {
      if (!s) return "#9ca3af";
      if (s <= 2) return "#16a34a";
      if (s <= 3) return "#ca8a04";
      if (s <= 4) return "#ea580c";
      return "#dc2626";
    };
    const daafLabel = (s) => {
      if (!s) return "–";
      if (s <= 1) return "1 — EU/EER volledig";
      if (s <= 2) return "2 — EU/EER beperkt";
      if (s === 3) return "3 — Adequaat + risico";
      if (s === 4) return "4 — SCCs, geen adequaat";
      return "5 — Geen waarborgen";
    };

    const typeKleur = { "Hosting & opslag":"#1A56A0", "AI-dienst":"#6d28d9", "Broncode":"#374151", "Open source software":"#15803d" };

    return (
      <div className="h-full overflow-y-auto" style={{ background:"#EBF3FF" }}>
        <div className="p-5 max-w-4xl mx-auto">

          {/* Header */}
          <div className="rounded p-5 mb-5 text-white" style={{ background:"linear-gradient(135deg, #0C2340 0%, #1A56A0 100%)" }}>
            <div className="flex items-center gap-3 mb-3">
              <div className="px-3 py-2 border-2 border-white" style={{ borderRadius:2 }}>
                <span className="font-bold leading-none" style={{ fontSize:10, letterSpacing:1 }}>NHL<br/>STENDEN</span>
              </div>
              <div className="w-px self-stretch" style={{ background:"#26B5AE", margin:"2px 0" }}/>
              <div>
                <h1 className="font-bold" style={{ fontSize:17 }}>Transparantie over deze applicatie</h1>
                <p style={{ fontSize:12, color:"#7DD3D0" }}>Digitale soevereiniteit van het instrument zelf</p>
              </div>
            </div>
            <p className="text-sm leading-relaxed" style={{ color:"rgba(255,255,255,0.85)" }}>
              Deze pagina beschrijft welke software en diensten ten grondslag liggen aan de Portfolioanalyse Digitale Soevereiniteit,
              waar die partijen zijn gevestigd, hoe de beveiliging is geregeld en waar de data wordt opgeslagen.
              Voor elk component zijn de DAAF-indicatoren A1 (jurisdictie) en A3 (datalocatie) beoordeeld
              op dezelfde schaal die we voor andere applicaties hanteren.
            </p>
          </div>

          {/* Disclaimer */}
          <div className="rounded p-3 mb-5 text-xs" style={{ background:"#fffbeb", border:"1px solid #fde68a", color:"#92400e" }}>
            <strong>Let op:</strong> De DAAF-scores op deze pagina zijn een eigen inschatting op basis van publiek beschikbare informatie.
            Voor een formele beoordeling zijn verwerkersovereenkomsten, DPA's en contractuele clausules nodig.
            De scores voor A1 en A3 zijn weergegeven; de overige DAAF-dimensies (mitigatie, belang) zijn niet beoordeeld
            omdat dit een interne ontwikkeltool is zonder persoonsgegevens van studenten of medewerkers.
          </div>

          {/* Stack-kaarten */}
          <div className="space-y-4 mb-5">
            {STACK.map((s, i) => (
              <div key={i} className="rounded" style={{ background:"#fff", border:"1px solid #D0E4F7", overflow:"hidden" }}>
                {/* Koptekst */}
                <div className="px-4 py-3 flex items-center justify-between" style={{ borderBottom:"1px solid #f1f5f9" }}>
                  <div className="flex items-center gap-3">
                    <div className="w-8 h-8 flex items-center justify-center text-white font-bold text-sm flex-shrink-0"
                      style={{ background:"#0C2340", borderRadius:4 }}>{i+1}</div>
                    <div>
                      <div className="flex items-center gap-2">
                        <h3 className="font-bold text-sm" style={{ color:"#0C2340" }}>{s.naam}</h3>
                        <span className="text-xs px-2 py-0.5 rounded font-medium"
                          style={{ background: typeKleur[s.type] + "18", color: typeKleur[s.type], border:`1px solid ${typeKleur[s.type]}44` }}>
                          {s.type}
                        </span>
                      </div>
                      <p className="text-xs" style={{ color:"#6b7280" }}>{s.rol}</p>
                    </div>
                  </div>
                  <a href={s.url} target="_blank" rel="noopener noreferrer"
                    className="text-xs font-medium" style={{ color:"#1A56A0", textDecoration:"none", flexShrink:0 }}>
                    ↗ Meer info
                  </a>
                </div>

                <div className="p-4">
                  <div className="grid grid-cols-3 gap-3">

                    {/* Vestiging & jurisdictie */}
                    <div className="rounded p-3" style={{ background:"#f8fafc", border:"1px solid #e5e7eb" }}>
                      <p className="text-xs font-bold mb-1 uppercase tracking-wide" style={{ color:"#9ca3af" }}>Vestiging & jurisdictie</p>
                      <p className="text-xs font-semibold mb-0.5" style={{ color:"#0C2340" }}>{s.jurisdictie}</p>
                      <p className="text-xs" style={{ color:"#6b7280" }}>{s.vestiging}</p>
                    </div>

                    {/* DAAF A1 */}
                    <div className="rounded p-3" style={{ background:"#EBF3FF", border:"1px solid #D0E4F7" }}>
                      <p className="text-xs font-bold mb-1 uppercase tracking-wide" style={{ color:"#9ca3af" }}>DAAF A1 — Jurisdictie leverancier</p>
                      <div className="flex items-center gap-2 mb-1">
                        <span className="text-sm font-bold" style={{ color: daafKleur(s.daafA1) }}>{s.daafA1}/5</span>
                        <span className="text-xs px-1.5 py-0.5 rounded font-medium" style={{ background: daafKleur(s.daafA1) + "18", color: daafKleur(s.daafA1) }}>
                          {daafLabel(s.daafA1)}
                        </span>
                      </div>
                      <p className="text-xs leading-relaxed" style={{ color:"#374151" }}>{s.daafA1toe}</p>
                    </div>

                    {/* DAAF A3 */}
                    <div className="rounded p-3" style={{ background:"#EBF3FF", border:"1px solid #D0E4F7" }}>
                      <p className="text-xs font-bold mb-1 uppercase tracking-wide" style={{ color:"#9ca3af" }}>DAAF A3 — Hosting & datalocatie</p>
                      <div className="flex items-center gap-2 mb-1">
                        <span className="text-sm font-bold" style={{ color: daafKleur(s.daafA3) }}>{s.daafA3}/5</span>
                        <span className="text-xs px-1.5 py-0.5 rounded font-medium" style={{ background: daafKleur(s.daafA3) + "18", color: daafKleur(s.daafA3) }}>
                          {daafLabel(s.daafA3)}
                        </span>
                      </div>
                      <p className="text-xs leading-relaxed" style={{ color:"#374151" }}>{s.daafA3toe}</p>
                    </div>
                  </div>

                  {/* Beveiliging + opmerking */}
                  <div className="grid grid-cols-2 gap-3 mt-3">
                    <div className="rounded p-2.5" style={{ background:"#f8fafc", border:"1px solid #e5e7eb" }}>
                      <p className="text-xs font-bold mb-1" style={{ color:"#0C2340" }}>🔒 Beveiliging</p>
                      <p className="text-xs leading-relaxed" style={{ color:"#374151" }}>{s.beveiliging}</p>
                    </div>
                    <div className="rounded p-2.5" style={{ background:"#fffbeb", border:"1px solid #fde68a" }}>
                      <p className="text-xs font-bold mb-1" style={{ color:"#92400e" }}>📌 Opmerking</p>
                      <p className="text-xs leading-relaxed" style={{ color:"#78350f" }}>{s.opmerking}</p>
                    </div>
                  </div>
                </div>
              </div>
            ))}
          </div>

          {/* Gegevensstroomoverzicht */}
          <div className="rounded p-4 mb-5" style={{ background:"#fff", border:"1px solid #D0E4F7" }}>
            <h3 className="font-bold text-sm mb-3" style={{ color:"#0C2340" }}>Gegevensstroom — wat gaat waarheen</h3>
            <div className="space-y-2">
              {[
                { van:"Gebruiker (browser)", naar:"Netlify Blobs (VS)", data:"Assessment-scores, applicatienamen, motivaties", actie:"Opslaan bij elke wijziging", kleur:"#ea580c" },
                { van:"Netlify Blobs (VS)", naar:"Gebruiker (browser)", data:"Dezelfde data terug bij laden van de pagina", actie:"Laden bij inloggen", kleur:"#ea580c" },
                { van:"Gebruiker (browser)", naar:"Lokale download", data:"Excel-exportbestand, PDF-rapport", actie:"Op verzoek van gebruiker", kleur:"#16a34a" },
                { van:"Broncode (GitHub, VS)", naar:"Netlify (VS)", data:"Applicatiecode — geen gebruikersdata", actie:"Bij elke git push (deploy)", kleur:"#ca8a04" },
                { van:"Claude API (Anthropic, VS)", naar:"–", data:"Niet actief tijdens gebruik", actie:"Alleen gebruikt bij ontwikkeling", kleur:"#9ca3af" },
              ].map((r, i) => (
                <div key={i} className="flex items-start gap-3 rounded p-2.5" style={{ background:"#f8fafc", border:"1px solid #e5e7eb" }}>
                  <div className="w-2 h-2 rounded-full flex-shrink-0 mt-1.5" style={{ background:r.kleur }}/>
                  <div className="flex-1 grid grid-cols-4 gap-2 text-xs">
                    <div><span className="font-semibold" style={{ color:"#0C2340" }}>Van:</span> <span style={{ color:"#374151" }}>{r.van}</span></div>
                    <div><span className="font-semibold" style={{ color:"#0C2340" }}>Naar:</span> <span style={{ color:"#374151" }}>{r.naar}</span></div>
                    <div><span className="font-semibold" style={{ color:"#0C2340" }}>Data:</span> <span style={{ color:"#374151" }}>{r.data}</span></div>
                    <div><span className="font-semibold" style={{ color:"#0C2340" }}>Wanneer:</span> <span style={{ color:"#374151" }}>{r.actie}</span></div>
                  </div>
                </div>
              ))}
            </div>
          </div>

          {/* Beveiliging */}
          <div className="rounded p-4 mb-4" style={{ background:"#fff", border:"1px solid #D0E4F7" }}>
            <h3 className="font-bold text-sm mb-3" style={{ color:"#0C2340" }}>Beveiliging van deze applicatie</h3>

            {/* Versleuteling */}
            <div className="mb-3">
              <p className="text-xs font-bold mb-2 uppercase tracking-wide" style={{ color:"#9ca3af" }}>Versleuteling</p>
              <div className="grid grid-cols-2 gap-3">
                {[
                  { icon:"🔒", titel:"Data in transit (TLS)", status:"Actief", kleur:"#15803d", bg:"#dcfce7",
                    tekst:"Alle communicatie tussen de browser en de Netlify-servers verloopt via HTTPS met TLS 1.2 of hoger. Data kan onderweg niet worden onderschept of gewijzigd." },
                  { icon:"💾", titel:"Data at rest (Netlify Blobs)", status:"Platformniveau", kleur:"#ca8a04", bg:"#fef9c3",
                    tekst:"Netlify versleutelt opgeslagen data op schijfniveau (AES-256) als onderdeel van hun infrastructuur. Er is geen applicatieniveau-encryptie: de data is leesbaar voor Netlify zelf en voor iedereen met toegang tot het Netlify-account van NHL Stenden." },
                ].map(k => (
                  <div key={k.titel} className="rounded p-3" style={{ background:k.bg, border:`1px solid ${k.kleur}44` }}>
                    <div className="flex items-center gap-2 mb-1.5">
                      <span>{k.icon}</span>
                      <span className="text-xs font-bold" style={{ color:k.kleur }}>{k.titel}</span>
                      <span className="text-xs px-1.5 py-0.5 rounded font-medium ml-auto" style={{ background:k.kleur+"22", color:k.kleur }}>{k.status}</span>
                    </div>
                    <p className="text-xs leading-relaxed" style={{ color:"#374151" }}>{k.tekst}</p>
                  </div>
                ))}
              </div>
            </div>

            {/* Toegangsbeveiliging */}
            <div className="mb-3">
              <p className="text-xs font-bold mb-2 uppercase tracking-wide" style={{ color:"#9ca3af" }}>Toegangsbeveiliging</p>
              <div className="grid grid-cols-3 gap-3">
                {[
                  { icon:"🖥️", titel:"Inlogscherm (UI)", status:"Wachtwoord", kleur:"#1A56A0", bg:"#EBF3FF",
                    tekst:"De applicatie toont een inlogscherm dat toegang blokkeert zonder het juiste wachtwoord. De sessie wordt opgeslagen in sessionStorage van de browser en vervalt automatisch wanneer het tabblad of venster wordt gesloten." },
                  { icon:"🔑", titel:"API-tokencheck (server)", status:"Actief v1.7+", kleur:"#15803d", bg:"#dcfce7",
                    tekst:"Vanaf versie 1.7 stuurt de browser bij elke API-aanroep een token mee als request-header (x-api-token). De Netlify Function valideert dit token server-side tegen een omgevingsvariabele (APP_API_TOKEN). Verzoeken zonder geldig token worden geweigerd met HTTP 401." },
                  { icon:"🔐", titel:"Beheerpincode", status:"Afzonderlijk", kleur:"#6d28d9", bg:"#faf5ff",
                    tekst:"Naast het inlogwachtwoord is er een aparte beheerpincode voor de beheerfuncties (aanpassen, verwijderen). Dit beperkt de impact als het algemene wachtwoord breed wordt gedeeld." },
                ].map(k => (
                  <div key={k.titel} className="rounded p-3" style={{ background:k.bg, border:`1px solid ${k.kleur}44` }}>
                    <div className="flex items-center gap-2 mb-1.5">
                      <span>{k.icon}</span>
                      <span className="text-xs font-bold" style={{ color:k.kleur }}>{k.titel}</span>
                    </div>
                    <span className="text-xs px-1.5 py-0.5 rounded font-medium" style={{ background:k.kleur+"22", color:k.kleur }}>{k.status}</span>
                    <p className="text-xs leading-relaxed mt-1.5" style={{ color:"#374151" }}>{k.tekst}</p>
                  </div>
                ))}
              </div>
            </div>

            {/* Bekende beperkingen */}
            <div className="rounded p-3 mb-3" style={{ background:"#fffbeb", border:"1px solid #fde68a" }}>
              <p className="text-xs font-bold mb-2" style={{ color:"#92400e" }}>⚠️ Bekende beperkingen — transparant benoemd</p>
              <div className="grid grid-cols-2 gap-2">
                {[
                  { tekst:"Het inlogwachtwoord staat als platte tekst in de JavaScript-broncode die naar de browser wordt gestuurd. Iedereen die de browser-devtools opent kan het wachtwoord inzien. De server-side tokencheck (v1.7+) beperkt de schade, maar vervangt geen echte authenticatie." },
                  { tekst:"Er is geen rate limiting of lockout na foutieve inlogpogingen. Geautomatiseerde aanvallen op het inlogscherm zijn technisch mogelijk." },
                  { tekst:"Er is geen logging van wie wanneer heeft ingelogd of data heeft gewijzigd. Bij een incident is er geen audit trail." },
                  { tekst:"De data staat zonder applicatieniveau-encryptie in Netlify Blobs. NHL Stenden heeft geen zicht op wie binnen Netlify toegang heeft tot de ruwe opgeslagen data." },
                ].map((b, i) => (
                  <div key={i} className="flex gap-2 text-xs">
                    <span className="flex-shrink-0 font-bold" style={{ color:"#ea580c" }}>!</span>
                    <span style={{ color:"#78350f", lineHeight:1.5 }}>{b.tekst}</span>
                  </div>
                ))}
              </div>
            </div>

            {/* Advies doorontwikkeling */}
            <div>
              <p className="text-xs font-bold mb-2 uppercase tracking-wide" style={{ color:"#9ca3af" }}>Advies bij doorontwikkeling</p>
              <div className="grid grid-cols-3 gap-3">
                {[
                  { n:"1", titel:"Niveau 1 — Gedaan (v1.7)", tekst:"Server-side API-tokencheck in Netlify Functions. Verzoeken zonder token worden geweigerd met HTTP 401. Minimale inspanning, directe verbetering.", kleur:"#15803d", bg:"#dcfce7" },
                  { n:"2", titel:"Niveau 2 — Aanbevolen", tekst:"Vervang het plaintext-wachtwoord door Netlify Identity of SURFconext SSO. Dan is er echte authenticatie met gebruikersaccounts, geen wachtwoord in de broncode, en sessie-beheer aan de serverkant.", kleur:"#1A56A0", bg:"#EBF3FF" },
                  { n:"3", titel:"Niveau 3 — Bij institutionele uitrol", tekst:"Migreer naar EU-hosting (Hetzner, Cloudflare EU, Nederlandse aanbieder), voeg applicatieniveau-encryptie toe aan de opgeslagen data, en integreer logging en audit trail.", kleur:"#6d28d9", bg:"#faf5ff" },
                ].map(k => (
                  <div key={k.n} className="rounded p-3" style={{ background:k.bg, border:`1px solid ${k.kleur}44` }}>
                    <div className="flex items-center gap-2 mb-1.5">
                      <div className="w-5 h-5 rounded-full flex items-center justify-center text-white font-bold flex-shrink-0"
                        style={{ background:k.kleur, fontSize:10 }}>{k.n}</div>
                      <span className="text-xs font-bold" style={{ color:k.kleur }}>{k.titel}</span>
                    </div>
                    <p className="text-xs leading-relaxed" style={{ color:"#374151" }}>{k.tekst}</p>
                  </div>
                ))}
              </div>
            </div>
          </div>

          {/* Conclusie soevereiniteit */}
          <div className="rounded p-4 mb-5" style={{ background:"#fff", border:"2px solid #E87722" }}>
            <h3 className="font-bold text-sm mb-2" style={{ color:"#0C2340" }}>Conclusie — soevereiniteitsrisico van dit instrument</h3>
            <p className="text-xs leading-relaxed mb-3" style={{ color:"#374151" }}>
              De applicatie draait volledig op Amerikaanse infrastructuur (Netlify, GitHub). De assessment-data van NHL Stenden
              wordt opgeslagen in Netlify Blobs zonder gegarandeerde EU-datalocatie.
              Dit is een bewuste pragmatische keuze voor dit prototype. De data bevat geen persoonsgegevens
              van studenten of medewerkers — het betreft uitsluitend scores en omschrijvingen van softwareapplicaties.
            </p>
            <div className="grid grid-cols-3 gap-3">
              {[
                { titel:"Wat dit betekent", tekst:"Data over het applicatielandschap van NHL Stenden staat op Amerikaanse servers. Bij een juridische vordering onder de CLOUD Act is toegang door Amerikaanse autoriteiten niet uit te sluiten.", kleur:"#ea580c", bg:"#ffedd5" },
                { titel:"Mitigerende factoren", tekst:"Geen persoonsgegevens. Data is niet bedrijfskritisch in de zin van DPIA. Regelmatige Excel-export biedt lokale back-up. De broncode is volledig inzichtelijk.", kleur:"#ca8a04", bg:"#fef9c3" },
                { titel:"Aanbeveling", tekst:"Overweeg bij doorontwikkeling migratie naar een EU-hostingpartij (bijv. Cloudflare Pages EU-regio, Hetzner of een Nederlandse aanbieder) en gebruik van een Europese blob-storage dienst.", kleur:"#15803d", bg:"#dcfce7" },
              ].map(k => (
                <div key={k.titel} className="rounded p-3" style={{ background:k.bg, border:`1px solid ${k.kleur}44` }}>
                  <p className="text-xs font-bold mb-1" style={{ color:k.kleur }}>{k.titel}</p>
                  <p className="text-xs leading-relaxed" style={{ color:"#374151" }}>{k.tekst}</p>
                </div>
              ))}
            </div>
          </div>

          {/* Footer */}
          <div className="rounded p-3 text-center" style={{ background:"#fff", border:"1px solid #D0E4F7" }}>
            <p className="text-xs" style={{ color:"#9ca3af" }}>
              NHL Stenden Hogeschool · Portfolioanalyse Digitale Soevereiniteit · {VERSION}
              <br/>Prototype v2.2 · In voorbereiding: migratie naar NVIDIA DGX Spark (eigen NHL Stenden-infrastructuur, Leeuwarden)
              <br/>Transparantiepagina samengesteld op basis van publiek beschikbare informatie · bronnen: netlify.com/security, anthropic.com/privacy, github.com/security
            </p>
          </div>

        </div>
      </div>
    );
  }

  // ── BESTUUR (Portfolio) ─────────────────────────────────────
  function Bestuur() {
    const scored = apps
      .map(a => ({ ...a, sc: calcScores(a.scores) }))
      .filter(a => a.sc.autonomyScore);

    const sorted     = [...scored].sort((a, b) => (a.sc.autonomyScore||10) - (b.sc.autonomyScore||10));
    const kritiek    = sorted.filter(a => a.sc.autonomyScore < 3);
    const zorg       = sorted.filter(a => a.sc.autonomyScore >= 3 && a.sc.autonomyScore < 5);
    const acceptabel = sorted.filter(a => a.sc.autonomyScore >= 5 && a.sc.autonomyScore < 7);
    const goed       = sorted.filter(a => a.sc.autonomyScore >= 7);
    const avg        = scored.length
      ? scored.reduce((s, a) => s + (a.sc.autonomyScore || 0), 0) / scored.length
      : null;
    const avgDictu   = scored.filter(a => a.sc.dictuAvg).length
      ? scored.filter(a => a.sc.dictuAvg).reduce((s, a) => s + a.sc.dictuAvg, 0) / scored.filter(a => a.sc.dictuAvg).length
      : null;

    // ── Snapshot opslaan (1 per dag) ─────────────────────────
    if (avg !== null && scored.length > 0) {
      const vandaagStr = new Date().toISOString().slice(0,10);
      const bestaatVandaag = snapshots.some(s => s.datum === vandaagStr);
      if (!bestaatVandaag) {
        const nieuweSnap = {
          datum: vandaagStr,
          avg: +avg.toFixed(2),
          avgDictu: avgDictu ? +avgDictu.toFixed(2) : null,
          totaal: scored.length,
          kritiek: kritiek.length,
          zorg: zorg.length,
          acceptabel: acceptabel.length,
          goed: goed.length,
        };
        const updated = [...snapshots, nieuweSnap].slice(-90);
        setSnapshots(updated);
        try { localStorage.setItem("nhl_sov_snapshots", JSON.stringify(updated)); } catch {}
      }
    }

    const oordeel = !scored.length
      ? { kleur:"#6b7280", bg:"#f3f4f6", border:"#e5e7eb", tekst:"Nog geen applicaties beoordeeld." }
      : kritiek.length >= 3 || (kritiek.length > 0 && kritiek.length / scored.length > 0.3)
      ? { kleur:"#b91c1c", bg:"#fee2e2", border:"#fca5a5",
          tekst:`Het portfolio bevat ${kritiek.length} kritieke ${kritiek.length === 1 ? "applicatie" : "applicaties"} met een hoog autonomierisico en onvoldoende weerbaarheid. Directe besluitvorming is noodzakelijk.` }
      : zorg.length > 0
      ? { kleur:"#c2410c", bg:"#ffedd5", border:"#fed7aa",
          tekst:`Het portfolio vraagt aandacht: ${zorg.length + kritiek.length} ${zorg.length + kritiek.length === 1 ? "applicatie" : "applicaties"} scoren onder de acceptabele grens. Gerichte maatregelen zijn gewenst.` }
      : { kleur:"#15803d", bg:"#dcfce7", border:"#86efac",
          tekst:"Het portfolio is grotendeels op orde. De meeste applicaties zijn acceptabel tot goed beoordeeld. Periodieke monitoring volstaat." };

    const top3 = [...scored]
      .filter(a => a.sc.autonomyScore < 7)
      .sort((a, b) => {
        const urgA = (10 - (a.sc.autonomyScore || 10)) + (a.sc.belang || 0);
        const urgB = (10 - (b.sc.autonomyScore || 10)) + (b.sc.belang || 0);
        return urgB - urgA;
      })
      .slice(0, 3);

    const grafiekData = sorted.slice(0, 12).map(a => ({
      naam:  displayName(a).substring(0, 22),
      score: a.sc.autonomyScore ? +a.sc.autonomyScore.toFixed(1) : 0,
      kleur: scoreColor(a.sc.autonomyScore),
    }));

    const vandaag = new Date().toLocaleDateString("nl-NL", { day:"numeric", month:"long", year:"numeric" });

    function adviesTekst(a) {
      const s = a.sc;
      if (!s.risico || !s.mitigatie) return "Vul het assessment verder in voor een specifiek advies.";
      if (s.risico > 3.5 && s.mitigatie < 2.5)
        return "Hoog risico en lage weerbaarheid. Prioriteit: migreer naar een Europese aanbieder of versterk contractuele exit-clausules.";
      if (s.risico > 3.5)
        return "Hoog risico bij een leverancier buiten de EU. Aanbevolen: bouw alternatieven en versterk de contractuele positie.";
      if (s.mitigatie < 2.5)
        return "Beperkte weerbaarheid. Aanbevolen: ontwikkel intern alternatief, leg kennis vast en voeg exit-clausules toe.";
      if (s.belang > 3.5)
        return "Hoog strategisch belang maakt het risico urgenter. Monitor actief bij contractverlenging en leverancierswijzigingen.";
      return "Aandacht gewenst. Bespreek in het team of actie of bewuste acceptatie de juiste keuze is.";
    }

    // ── Top 3 concrete acties per aandachtspunt ───────────────
    function top3Acties(a) {
      const s = a.sc;
      const acties = [];
      if (s.risico > 3.5)
        acties.push("Verken een Europese alternatieven voor deze leverancier (inventariseer markt, voer een orienterend gesprek).");
      if (s.mitigatie < 2.5 && s.dimScores?.C < 2)
        acties.push("Documenteer en test een noodprocedure voor het geval deze applicatie uitvalt of niet meer toegankelijk is.");
      if (s.mitigatie < 2.5 && s.dimScores?.D < 2)
        acties.push("Leg de kennis over deze applicatie vast bij minimaal twee medewerkers. Verminder de afhankelijkheid van individuele sleutelpersonen.");
      if (s.mitigatie < 2.5 && s.dimScores?.E < 2)
        acties.push("Voeg bij de volgende contractverlenging een exit-clausule en een dataportabiliteitsgarantie toe.");
      if (s.dictuAvg && s.dictuAvg < 3)
        acties.push("Vraag de leverancier schriftelijk naar de datalocatie en of zij zich actief verzetten tegen niet-EU dataverzoeken (DICTU 2.3).");
      if (s.belang > 3.5 && s.risico > 3)
        acties.push("Formeel vastleggen: is het risico van deze applicatie bewust aanvaard door NHL Stenden? Leg de afweging schriftelijk vast.");
      // Fallback als niets specifiek
      if (acties.length === 0) {
        acties.push("Herassessment plannen bij eerstvolgende contractverlenging.");
        acties.push("Controleer of de gekozen scores nog actueel zijn na evt. leverancierswijzigingen.");
        acties.push("Bespreek de positie van deze applicatie in het kwadrant met de applicatie-eigenaar.");
      }
      return acties.slice(0, 3);
    }

    // ── SVG grafiek helper ────────────────────────────────────
    function VoortgangsGrafiek({ snaps }) {
      if (snaps.length < 2) return (
        <div className="rounded p-6 text-center" style={{ background:"#f8fafc", border:"1px dashed #D0E4F7" }}>
          <p className="text-xs" style={{ color:"#9ca3af" }}>
            De grafiek verschijnt zodra er op meerdere dagen data is vastgelegd.
            Kom morgen terug voor het eerste verloop.
          </p>
          <p className="text-xs mt-1" style={{ color:"#D0E4F7" }}>
            {snaps.length === 1 ? `Eerste meting: ${snaps[0].datum}` : "Nog geen metingen"}
          </p>
        </div>
      );

      const W = 680, H = 200, PAD = { t:20, r:20, b:40, l:45 };
      const iW = W - PAD.l - PAD.r;
      const iH = H - PAD.t - PAD.b;

      const toX = i => PAD.l + (i / (snaps.length - 1)) * iW;
      const toY = v => PAD.t + iH - ((v - 0) / 10) * iH;

      // Lijn voor avg
      const avgPath = snaps.map((s, i) => `${i === 0 ? "M" : "L"}${toX(i).toFixed(1)},${toY(s.avg).toFixed(1)}`).join(" ");
      // Lijn voor dictu (schaal 0-5 omgezet naar 0-10)
      const dictuPath = snaps
        .filter(s => s.avgDictu)
        .map((s, i) => {
          const origIdx = snaps.indexOf(s);
          return `${i === 0 ? "M" : "L"}${toX(origIdx).toFixed(1)},${toY(s.avgDictu * 2).toFixed(1)}`;
        }).join(" ");

      // Referentielijnen
      const refLines = [
        { y: 7, label: "Goed (7)", color: "#16a34a" },
        { y: 5, label: "Acceptabel (5)", color: "#ca8a04" },
        { y: 3, label: "Zorg (3)", color: "#ea580c" },
      ];

      // X-as labels: toon max 6 labels
      const xLabels = snaps.length <= 6
        ? snaps.map((s, i) => ({ i, label: s.datum.slice(5) }))
        : [0, Math.floor(snaps.length*0.25), Math.floor(snaps.length*0.5), Math.floor(snaps.length*0.75), snaps.length-1]
            .filter((v,i,a) => a.indexOf(v) === i)
            .map(i => ({ i, label: snaps[i].datum.slice(5) }));

      return (
        <div>
          <svg viewBox={`0 0 ${W} ${H}`} style={{ width:"100%", height:"auto" }}>
            <defs>
              <linearGradient id="avgGrad" x1="0" y1="0" x2="0" y2="1">
                <stop offset="0%" stopColor="#1A56A0" stopOpacity="0.15"/>
                <stop offset="100%" stopColor="#1A56A0" stopOpacity="0"/>
              </linearGradient>
            </defs>

            {/* Referentielijnen */}
            {refLines.map(r => (
              <g key={r.y}>
                <line x1={PAD.l} y1={toY(r.y)} x2={W-PAD.r} y2={toY(r.y)}
                  stroke={r.color} strokeWidth="1" strokeDasharray="4,3" opacity="0.5"/>
                <text x={PAD.l - 4} y={toY(r.y)+3} textAnchor="end" fontSize="8" fill={r.color} opacity="0.8">{r.y}</text>
              </g>
            ))}

            {/* Y-as */}
            <line x1={PAD.l} y1={PAD.t} x2={PAD.l} y2={PAD.t+iH} stroke="#e5e7eb" strokeWidth="1"/>

            {/* Vlakke vulling onder avg-lijn */}
            <path d={`${avgPath} L${toX(snaps.length-1).toFixed(1)},${PAD.t+iH} L${PAD.l},${PAD.t+iH} Z`}
              fill="url(#avgGrad)"/>

            {/* DICTU lijn */}
            {dictuPath && <path d={dictuPath} fill="none" stroke="#26B5AE" strokeWidth="1.5" strokeDasharray="5,3" opacity="0.7"/>}

            {/* AVG lijn */}
            <path d={avgPath} fill="none" stroke="#1A56A0" strokeWidth="2.5"/>

            {/* Datapunten avg */}
            {snaps.map((s, i) => (
              <circle key={i} cx={toX(i)} cy={toY(s.avg)} r="3.5"
                fill="#fff" stroke="#1A56A0" strokeWidth="2"/>
            ))}

            {/* Laatste punt waarde label */}
            {snaps.length > 0 && (
              <text x={toX(snaps.length-1)+6} y={toY(snaps[snaps.length-1].avg)+4}
                fontSize="10" fontWeight="bold" fill="#1A56A0">
                {snaps[snaps.length-1].avg}
              </text>
            )}

            {/* X-as labels */}
            {xLabels.map(({ i, label }) => (
              <text key={i} x={toX(i)} y={H-8} textAnchor="middle" fontSize="8" fill="#9ca3af">{label}</text>
            ))}
          </svg>

          {/* Legenda */}
          <div className="flex gap-4 justify-center mt-1">
            <div className="flex items-center gap-1.5 text-xs" style={{ color:"#6b7280" }}>
              <svg width="20" height="8"><line x1="0" y1="4" x2="20" y2="4" stroke="#1A56A0" strokeWidth="2.5"/></svg>
              Gem. autonomiescore (DAAF)
            </div>
            <div className="flex items-center gap-1.5 text-xs" style={{ color:"#6b7280" }}>
              <svg width="20" height="8"><line x1="0" y1="4" x2="20" y2="4" stroke="#26B5AE" strokeWidth="1.5" strokeDasharray="5,3"/></svg>
              Gem. DICTU-score (x2 voor schaal)
            </div>
          </div>
        </div>
      );
    }

    return (
      <div className="h-full overflow-y-auto" style={{ background:"#EBF3FF" }}>
        <div className="p-5 max-w-4xl mx-auto">

          {/* Header */}
          <div className="rounded p-5 mb-5 text-white" style={{ background:"linear-gradient(135deg, #0C2340 0%, #1A56A0 100%)" }}>
            <div className="flex items-start justify-between">
              <div className="flex items-center gap-3">
                <div className="px-3 py-2 border-2 border-white" style={{ borderRadius:2 }}>
                  <span className="font-bold leading-none" style={{ fontSize:10, letterSpacing:1 }}>NHL<br/>STENDEN</span>
                </div>
                <div className="w-px self-stretch" style={{ background:"#26B5AE", margin:"2px 0" }}/>
                <div>
                  <h1 className="font-bold" style={{ fontSize:17 }}>Portfoliostatus Digitale Soevereiniteit</h1>
                  <p style={{ fontSize:12, color:"#7DD3D0" }}>Ambassadeurslijn Digitale Soevereiniteit · {vandaag}</p>
                </div>
              </div>
              <span className="text-xs px-2 py-1 rounded flex-shrink-0" style={{ background:"rgba(255,255,255,0.15)", color:"#7DD3D0" }}>
                {scored.length} van {apps.length} apps beoordeeld
              </span>
            </div>
          </div>

          {scored.length === 0 ? (
            <div className="rounded p-10 text-center" style={{ background:"#fff", border:"2px dashed #D0E4F7", color:"#9ca3af" }}>
              Nog geen applicaties volledig beoordeeld. Vul eerst assessments in via het tabblad Applicaties.
            </div>
          ) : (<>

          {/* ── Rij 1: Stand van zaken + Kerngetallen ── */}
          <div className="grid grid-cols-3 gap-4 mb-4">
            <div className="col-span-2 rounded p-4" style={{ background:"#fff", border:`2px solid ${oordeel.border}` }}>
              <p className="text-xs font-bold mb-2 uppercase tracking-wide" style={{ color:"#9ca3af" }}>Waar staan we</p>
              <div className="rounded p-3 mb-3" style={{ background:oordeel.bg, border:`1px solid ${oordeel.border}` }}>
                <p className="text-sm font-semibold leading-relaxed" style={{ color:oordeel.kleur }}>{oordeel.tekst}</p>
              </div>
              <p className="text-xs leading-relaxed" style={{ color:"#374151" }}>
                NHL Stenden heeft {apps.length} kernapp{apps.length !== 1 ? "licaties" : "licatie"} in scope genomen
                voor de portfolioanalyse digitale soevereiniteit.
                {scored.length < apps.length && ` Van ${apps.length - scored.length} ${apps.length - scored.length === 1 ? "applicatie" : "applicaties"} is het assessment nog niet volledig ingevuld.`}
                {" "}De beoordeling combineert het DAAF-framework (autonomiescore 1-10) en de DICTU soevereiniteitscheck.
              </p>
            </div>
            <div className="space-y-2">
              {[
                { lbl:"Gem. autonomiescore", val: avg ? avg.toFixed(1) : "–", sub:"/10", kleur: scoreColor(avg) },
                { lbl:"Gem. DICTU-score",    val: avgDictu ? avgDictu.toFixed(1) : "–", sub:"/5",  kleur: scoreColor(avgDictu, 5) },
              ].map(k => (
                <div key={k.lbl} className="rounded p-3 text-center" style={{ background:"#fff", border:"1px solid #D0E4F7" }}>
                  <p style={{ fontSize:9, color:"#9ca3af", marginBottom:4, textTransform:"uppercase", letterSpacing:"0.05em" }}>{k.lbl}</p>
                  <p className="font-bold" style={{ fontSize:26, color:k.kleur, lineHeight:1 }}>
                    {k.val}<span style={{ fontSize:13, color:"#9ca3af" }}>{k.sub}</span>
                  </p>
                </div>
              ))}
              <div className="rounded p-3" style={{ background:"#fff", border:"1px solid #D0E4F7" }}>
                <p style={{ fontSize:9, color:"#9ca3af", marginBottom:6, textTransform:"uppercase", letterSpacing:"0.05em" }}>Verdeling portfolio</p>
                {[
                  { lbl:"Kritiek",     n:kritiek.length,    bg:"#fee2e2", fg:"#b91c1c", dot:"#dc2626" },
                  { lbl:"Zorgwekkend", n:zorg.length,       bg:"#ffedd5", fg:"#c2410c", dot:"#ea580c" },
                  { lbl:"Acceptabel",  n:acceptabel.length, bg:"#fef9c3", fg:"#a16207", dot:"#ca8a04" },
                  { lbl:"Goed",        n:goed.length,       bg:"#dcfce7", fg:"#15803d", dot:"#16a34a" },
                ].map(r => (
                  <div key={r.lbl} className="flex items-center justify-between mb-1">
                    <div className="flex items-center gap-1.5">
                      <span style={{ width:8, height:8, borderRadius:"50%", background:r.dot, display:"inline-block" }}/>
                      <span style={{ fontSize:10, color:r.fg, fontWeight:600 }}>{r.lbl}</span>
                    </div>
                    <span className="px-2 py-0.5 rounded text-xs font-bold" style={{ background:r.bg, color:r.fg }}>{r.n}</span>
                  </div>
                ))}
              </div>
            </div>
          </div>

          {/* ── Rij 2: Risicografiek ── */}
          <div className="rounded p-4 mb-4" style={{ background:"#fff", border:"1px solid #D0E4F7" }}>
            <div className="flex items-center justify-between mb-3">
              <div>
                <h3 className="font-bold text-sm" style={{ color:"#0C2340" }}>Autonomiescore per applicatie</h3>
                <p className="text-xs" style={{ color:"#9ca3af" }}>Gesorteerd van laagste naar hoogste · Gele lijn = grens acceptabel (5) · Groene lijn = grens goed (7)</p>
              </div>
            </div>
            <div className="space-y-1.5">
              {grafiekData.map((d, i) => (
                <div key={i} className="flex items-center gap-2">
                  <div className="text-xs text-right flex-shrink-0" style={{ width:150, color:"#374151", overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}
                    title={d.naam}>{d.naam}</div>
                  <div className="flex-1 rounded" style={{ background:"#f1f5f9", height:22, position:"relative" }}>
                    <div style={{ position:"absolute", left:"50%", top:0, bottom:0, width:1, background:"#fbbf24", opacity:0.7 }}/>
                    <div style={{ position:"absolute", left:"70%", top:0, bottom:0, width:1, background:"#4ade80", opacity:0.7 }}/>
                    <div style={{ position:"absolute", left:0, top:2, bottom:2, width:`${(d.score/10)*100}%`, background:d.kleur, borderRadius:3, minWidth:d.score>0?4:0 }}/>
                  </div>
                  <div className="text-xs font-bold flex-shrink-0" style={{ width:30, color:d.kleur, textAlign:"right" }}>{d.score}</div>
                </div>
              ))}
            </div>
          </div>

          {/* ── Rij 3: Top 3 aandachtspunten met concrete acties ── */}
          <div className="rounded p-4 mb-4" style={{ background:"#fff", border:"1px solid #D0E4F7" }}>
            <h3 className="font-bold text-sm mb-1" style={{ color:"#0C2340" }}>Top 3 aandachtspunten — inclusief concrete acties</h3>
            <p className="text-xs mb-4" style={{ color:"#9ca3af" }}>
              Geselecteerd op combinatie van laagste autonomiescore en hoogste strategisch belang. Per applicatie drie concrete vervolgstappen.
            </p>
            {top3.length === 0 ? (
              <div className="rounded p-4 text-center text-xs" style={{ background:"#dcfce7", border:"1px solid #86efac", color:"#15803d" }}>
                Alle beoordeelde applicaties scoren acceptabel of goed. Geen acute aandachtspunten.
              </div>
            ) : (
              <div className="space-y-4">
                {top3.map((a, i) => {
                  const lbl = scoreLabel(a.sc.autonomyScore);
                  const medalColors = ["#b91c1c","#c2410c","#a16207"];
                  const acties = top3Acties(a);
                  return (
                    <div key={a.id} className="rounded" style={{ background:"#f8fafc", border:"1px solid #e5e7eb", borderLeft:`4px solid ${scoreColor(a.sc.autonomyScore)}`, overflow:"hidden" }}>
                      {/* Applicatie-header */}
                      <div className="px-4 py-3 flex items-center gap-3" style={{ borderBottom:"1px solid #f1f5f9" }}>
                        <div className="w-7 h-7 flex items-center justify-center flex-shrink-0 text-white font-bold rounded text-xs"
                          style={{ background:medalColors[i] }}>{i+1}</div>
                        <div className="flex-1 flex items-center gap-2 flex-wrap">
                          <h4 className="font-bold text-sm" style={{ color:"#0C2340" }}>{displayName(a)}</h4>
                          {a.supplier && <span className="text-xs" style={{ color:"#9ca3af" }}>{a.supplier}</span>}
                          <span className="text-xs px-2 py-0.5 font-semibold rounded" style={{ background:lbl.bg, color:lbl.fg }}>{lbl.text}</span>
                          <span className="text-xs font-bold" style={{ color:scoreColor(a.sc.autonomyScore) }}>Score {a.sc.autonomyScore?.toFixed(1)}/10</span>
                        </div>
                      </div>
                      <div className="px-4 py-3 grid grid-cols-2 gap-4">
                        {/* Scores */}
                        <div>
                          <p className="text-xs font-bold mb-2" style={{ color:"#9ca3af" }}>SCORES</p>
                          <div className="flex gap-2 flex-wrap">
                            {[
                              { lbl:"Risico",    val:a.sc.risico,    c:"#dc2626" },
                              { lbl:"Mitigatie", val:a.sc.mitigatie, c:"#26B5AE" },
                              { lbl:"Belang",    val:a.sc.belang,    c:"#E87722" },
                              { lbl:"DICTU",     val:a.sc.dictuAvg,  c:"#6d28d9" },
                            ].map(s => (
                              <div key={s.lbl} className="text-center rounded px-2 py-1" style={{ background:"#fff", border:"1px solid #e5e7eb", minWidth:54 }}>
                                <div style={{ fontSize:13, fontWeight:700, color:s.val ? s.c : "#d1d5db" }}>{s.val ? s.val.toFixed(1) : "–"}</div>
                                <div style={{ fontSize:9, color:"#9ca3af" }}>{s.lbl}</div>
                              </div>
                            ))}
                          </div>
                        </div>
                        {/* Acties */}
                        <div>
                          <p className="text-xs font-bold mb-2" style={{ color:"#9ca3af" }}>TOP 3 ACTIES</p>
                          <div className="space-y-1.5">
                            {acties.map((actie, ai) => (
                              <div key={ai} className="flex gap-2 text-xs">
                                <span className="flex-shrink-0 w-4 h-4 rounded-full flex items-center justify-center text-white font-bold"
                                  style={{ background:medalColors[ai] || "#6b7280", fontSize:9 }}>{ai+1}</span>
                                <span style={{ color:"#374151", lineHeight:1.4 }}>{actie}</span>
                              </div>
                            ))}
                          </div>
                        </div>
                      </div>
                    </div>
                  );
                })}
              </div>
            )}
          </div>

          {/* ── Voortgangsgrafiek (gimmick, onderaan) ── */}
          <div className="rounded p-4 mb-4" style={{ background:"#fff", border:"1px solid #D0E4F7" }}>
            <div className="flex items-center justify-between mb-1">
              <div>
                <h3 className="font-bold text-sm" style={{ color:"#0C2340" }}>Verloop gemiddelde scores in de tijd</h3>
                <p className="text-xs" style={{ color:"#9ca3af" }}>
                  Één meting per dag, lokaal opgeslagen. Laat zien hoe het portfolio zich ontwikkelt naarmate meer assessments worden ingevuld of bijgewerkt.
                </p>
              </div>
              <button onClick={() => {
                if (window.confirm("Wil je de meethistorie wissen?")) {
                  localStorage.removeItem("nhl_sov_snapshots");
                  setSnapshots([]);
                }
              }} className="text-xs px-2 py-1 rounded" style={{ background:"#f3f4f6", color:"#9ca3af", border:"1px solid #e5e7eb" }}>
                Wissen
              </button>
            </div>
            <div className="mt-3">
              <VoortgangsGrafiek snaps={snapshots} />
            </div>
          </div>

          {/* ── Geopolitiek overzicht compact ── */}
          <GeoKaartCompact apps={apps} useSecondaryName={useSecondaryName} calcScores={calcScores} appColor={appColor} d3={d3} />

          {/* Footer */}
          <div className="rounded p-3 text-center" style={{ background:"#fff", border:"1px solid #D0E4F7" }}>
            <p className="text-xs" style={{ color:"#9ca3af" }}>
              NHL Stenden Hogeschool · Portfolioanalyse Digitale Soevereiniteit · {VERSION} · {vandaag}
              
            </p>
          </div>

          </>)}
        </div>
      </div>
    );
  }

  function About() {

    const ABOUT_TABS = [
      { k:"over",    icon:"🏠", label:"Over de tool"      },
      { k:"starten", icon:"🚀", label:"Aan de slag"        },
      { k:"scores",  icon:"📊", label:"Scores & grafieken" },
      { k:"tips",    icon:"💡", label:"Tips & beheer"      },
    ];

    return (
      <div className="h-full overflow-y-auto" style={{ background:"#EBF3FF" }}>
        <div className="p-5 max-w-4xl mx-auto">

          {/* Hero — altijd zichtbaar */}
          <div className="rounded p-5 mb-4 text-white" style={{ background:"linear-gradient(135deg, #0C2340 0%, #1A56A0 100%)" }}>
            <div className="flex items-center gap-3 mb-3">
              <div className="px-3 py-2 border-2 border-white" style={{ borderRadius:2 }}>
                <span className="font-bold leading-none" style={{ fontSize:10, letterSpacing:1 }}>NHL<br/>STENDEN</span>
              </div>
              <div className="w-px self-stretch" style={{ background:"#26B5AE", margin:"2px 0" }}/>
              <div>
                <h1 className="font-bold" style={{ fontSize:17 }}>Portfolioanalyse Digitale Soevereiniteit</h1>
                <p style={{ fontSize:12, color:"#7DD3D0" }}>Ambassadeurslijn Digitale Soevereiniteit · Programma Digitale Samenhang</p>
              </div>
            </div>
            <p className="text-sm leading-relaxed" style={{ color:"rgba(255,255,255,0.85)" }}>
              Een levend instrument voor NHL Stenden Hogeschool om per applicatie te beoordelen hoe urgent het
              autonomieprobleem is. Combinatie van <strong>DAAF</strong> (Universiteit Utrecht) en <strong>DICTU</strong> (Rijksoverheid).
              Gebruik de tabbladen hieronder om meer te leren over de tool.
            </p>
          </div>

          {/* Tab-navigatie */}
          <div className="flex gap-1 mb-5 rounded p-1" style={{ background:"#fff", border:"1px solid #D0E4F7" }}>
            {ABOUT_TABS.map(t => (
              <button key={t.k} onClick={() => setAboutTab(t.k)}
                className="flex-1 flex items-center justify-center gap-1.5 py-2 text-xs font-semibold transition-all"
                style={{
                  borderRadius: 4,
                  background: aboutTab === t.k ? "#1A56A0" : "transparent",
                  color:       aboutTab === t.k ? "#fff"    : "#6b7280",
                }}>
                <span>{t.icon}</span>
                <span>{t.label}</span>
              </button>
            ))}
          </div>

          {/* ── Tab 1: Over de tool ── */}
          {aboutTab === "over" && <>
            <Section title="Wat doet deze tool?">
              <div className="rounded p-4 mb-3" style={{ background:"#fff", border:"1px solid #D0E4F7" }}>
                <p className="text-sm leading-relaxed mb-3" style={{ color:"#374151" }}>
                  De <strong>Portfolioanalyse Digitale Soevereiniteit</strong> helpt NHL Stenden om per applicatie
                  te beoordelen hoe afhankelijk de instelling is van externe leveranciers, en hoe urgent het is
                  om actie te ondernemen. Geen statisch rapport, maar een realtime tool die je samen invult en
                  die direct resultaat laat zien.
                </p>
                <div className="grid grid-cols-3 gap-3">
                  {[
                    { icon:"📋", titel:"Beoordelen",  txt:"Vul per applicatie een assessment in met vragen over jurisdictie, datalocatie, contracten en strategisch belang." },
                    { icon:"📊", titel:"Vergelijken", txt:"Het dashboard en de vergelijkingspagina tonen alle applicaties naast elkaar. Je ziet direct welke urgent zijn." },
                    { icon:"📥", titel:"Rapporteren", txt:"Exporteer het volledige overzicht als Excel of genereer een PDF per applicatie voor besluitvorming." },
                  ].map(k => (
                    <div key={k.titel} className="rounded p-3 text-center" style={{ background:"#EBF3FF", border:"1px solid #D0E4F7" }}>
                      <div style={{ fontSize:24, marginBottom:6 }}>{k.icon}</div>
                      <p className="text-xs font-bold mb-1" style={{ color:"#0C2340" }}>{k.titel}</p>
                      <p className="text-xs" style={{ color:"#6b7280", lineHeight:1.5 }}>{k.txt}</p>
                    </div>
                  ))}
                </div>
              </div>

              {/* Prototype-status */}
              <div className="rounded p-4 mb-3" style={{ background:"#fffbeb", border:"1px solid #fde68a" }}>
                <div className="flex items-start gap-3">
                  <span style={{ fontSize:20, flexShrink:0 }}>🧪</span>
                  <div>
                    <p className="text-sm font-bold mb-1" style={{ color:"#92400e" }}>Status: Prototype</p>
                    <p className="text-xs leading-relaxed" style={{ color:"#374151" }}>
                      Dit instrument is een <strong>werkend prototype</strong> — volledig functioneel en in gebruik, maar nog niet definitief ingericht voor productie.
                      Het is ontwikkeld in mei–juni 2026 als onderdeel van de Ambassadeurslijn Digitale Soevereiniteit van het Programma Digitale Samenhang.
                      De huidige fase is gericht op het vullen van de database met applicaties en het valideren van de methodiek.
                    </p>
                  </div>
                </div>
              </div>

              {/* Migratie naar eigen server */}
              <div className="rounded p-4 mb-3" style={{ background:"#f0f9f9", border:"1px solid #26B5AE44" }}>
                <div className="flex items-start gap-3">
                  <span style={{ fontSize:20, flexShrink:0 }}>🖥️</span>
                  <div>
                    <p className="text-sm font-bold mb-1" style={{ color:"#0C2340" }}>Van cloud naar eigen NHL Stenden-infrastructuur</p>
                    <p className="text-xs leading-relaxed mb-2" style={{ color:"#374151" }}>
                      Het instrument draait momenteel op Amerikaanse cloudinfrastructuur (Netlify, GitHub). Het team Infrastructuur van NHL Stenden
                      bereidt de migratie voor naar een <strong>NVIDIA DGX Spark</strong> — de eigen mini-supercomputer van NHL Stenden.
                      Na de migratie draait het instrument volledig op eigen NHL Stenden-hardware in Leeuwarden, onder Nederlands recht,
                      zonder afhankelijkheid van externe cloudpartijen.
                    </p>
                    <p className="text-xs leading-relaxed" style={{ color:"#6b7280" }}>
                      Dat maakt dit product ook intern consistent: een instrument dat digitale soevereiniteit meet,
                      draait straks zelf maximaal soeverein — score 1 op alle DAAF en DICTU-dimensies.
                    </p>
                  </div>
                </div>
              </div>

              {/* Inspiratiebron */}
              <div className="rounded p-4" style={{ background:"#EBF3FF", border:"1px solid #D0E4F7" }}>
                <div className="flex items-start gap-3">
                  <span style={{ fontSize:20, flexShrink:0 }}>💡</span>
                  <div>
                    <p className="text-sm font-bold mb-1" style={{ color:"#0C2340" }}>Een voorbeeld voor beleidsontwikkeling in de eigen organisatie</p>
                    <p className="text-xs leading-relaxed" style={{ color:"#374151" }}>
                      Dit instrument is meer dan een tool voor digitale soevereiniteit. Het laat zien hoe je beleid en de toepassing daarvan
                      kunt ontwerpen, verwerken en gebruiken in een hedendaagse manier die aansluit bij de eigen organisatie.
                      Ontwikkeld in nauwe samenwerking met een AI-assistent, in enkele weken gebouwd van nul tot werkend product,
                      en direct inzetbaar voor het team — zonder externe leverancier, zonder IT-ticket, zonder lang traject.
                    </p>
                    <p className="text-xs leading-relaxed mt-2" style={{ color:"#6b7280" }}>
                      Dit is de praktijk van digitale geletterdheid: technologie begrijpen, zelf regie voeren, en instrumenten bouwen
                      die passen bij wat de organisatie nodig heeft.
                    </p>
                  </div>
                </div>
              </div>
            </Section>

            <Section title="Gebruikte frameworks">
              <div className="grid grid-cols-2 gap-3 mb-2">
                <div className="rounded p-4" style={{ background:"#fff", border:"1px solid #D0E4F7" }}>
                  <div className="flex items-center justify-between gap-2 mb-2">
                    <div className="flex items-center gap-2">
                      <span className="text-xs font-bold px-2 py-0.5" style={{ background:"#1A56A0", color:"#fff", borderRadius:3 }}>DAAF</span>
                      <span className="font-semibold text-sm" style={{ color:"#0C2340" }}>Digital Autonomy Assessment Framework</span>
                    </div>
                    <a href="https://utrechtuniversity.github.io/digital-autonomy-assessment-tool/" target="_blank" rel="noopener noreferrer"
                      className="text-xs font-medium px-2 py-0.5"
                      style={{ color:"#1A56A0", border:"1px solid #D0E4F7", borderRadius:3, textDecoration:"none", whiteSpace:"nowrap" }}>
                      🔗 Open DAAF-tool ↗
                    </a>
                  </div>
                  <p className="text-xs leading-relaxed mb-2" style={{ color:"#6b7280" }}>
                    Ontwikkeld door de <strong>Universiteit Utrecht</strong>. Beoordeelt digitale autonomie ten opzichte van leveranciers
                    op 22 indicatoren verdeeld over 8 dimensies en 3 analyseniveaus. Het framework is gratis beschikbaar en
                    continu in ontwikkeling.
                  </p>
                  <div className="space-y-1 mb-3">
                    {[
                      { lv:"Niveau 1 · Risico (A, B)",       txt:"Jurisdictie leverancier, hosting & datalocatie, vendor concentratie.",   c:"#dc2626" },
                      { lv:"Niveau 2 · Mitigatie (C, D, E)", txt:"Alternatieven beschikbaar, kennis in huis, contractuele bescherming.",   c:"#26B5AE" },
                      { lv:"Niveau 3 · Belang (F, G, H)",    txt:"Operationeel, data-gevoeligheid en academisch belang van de applicatie.", c:"#E87722" },
                    ].map(r => (
                      <div key={r.lv} className="flex gap-2 text-xs">
                        <span className="w-2 h-2 rounded-full mt-1 flex-shrink-0" style={{ background:r.c }}/>
                        <div><strong style={{ color:r.c }}>{r.lv}:</strong> {r.txt}</div>
                      </div>
                    ))}
                  </div>
                  <div className="rounded p-2 mb-2" style={{ background:"#fffbeb", border:"1px solid #fde68a" }}>
                    <p className="text-xs font-semibold mb-0.5" style={{ color:"#92400e" }}>🧪 Quickscan vs. volledige scan</p>
                    <p className="text-xs leading-relaxed" style={{ color:"#374151" }}>
                      Dit prototype gebruikt de <strong>DAAF Quickscan</strong> — 8 kernindicatoren (één per dimensie) voor een snelle indicatie.
                      De volledige DAAF-scan beoordeelt <strong>22 indicatoren</strong> voor een diepgaandere analyse.
                      In een vervolg op dit prototype wordt de uitgebreide scan geïmplementeerd voor een completer en robuuster beeld.
                    </p>
                  </div>
                  <p className="text-xs italic" style={{ color:"#9ca3af" }}>Autonomiescore (1–10) = Mitigatie ÷ (Risico × Belang), logaritmisch genormaliseerd.</p>
                </div>
                <div className="rounded p-4" style={{ background:"#fff", border:"1px solid #26B5AE44" }}>
                  <div className="flex items-center justify-between gap-2 mb-2">
                    <div className="flex items-center gap-2">
                      <span className="text-xs font-bold px-2 py-0.5" style={{ background:"#26B5AE", color:"#fff", borderRadius:3 }}>DICTU</span>
                      <span className="font-semibold text-sm" style={{ color:"#0C2340" }}>Toetsingsinstrument Soevereiniteit Clouddiensten</span>
                    </div>
                    <a href="https://www.dictu.nl/sites/default/files/bestanden/website/DICTU%20Toetsingsinstrument%20Soevereiniteit%20Clouddiensten%20v1.0.1.pdf" target="_blank" rel="noopener noreferrer"
                      className="text-xs font-medium px-2 py-0.5"
                      style={{ color:"#26B5AE", border:"1px solid #26B5AE44", borderRadius:3, textDecoration:"none", whiteSpace:"nowrap" }}>
                      🔗 Open DICTU-framework ↗
                    </a>
                  </div>
                  <p className="text-xs leading-relaxed mb-2" style={{ color:"#6b7280" }}>
                    Gepubliceerd door <strong>DICTU (Dienst ICT Uitvoering)</strong> van het Ministerie van Economische Zaken en Klimaat.
                    Het volledige instrument beoordeelt clouddiensten op vijf dimensies: juridisch, data & AI, technologie, operationeel en mens.
                    NHL Stenden past de vier meest relevante indicatoren toe:
                  </p>
                  <div className="space-y-1 mb-2">
                    {[
                      { k:"2.1", lbl:"Dataresidency",         txt:"Staat alle data (incl. back-ups) uitsluitend in de EU?" },
                      { k:"2.2", lbl:"Technische beveiliging", txt:"Zijn er verifieerbare garanties dat niemand de data kan inzien?" },
                      { k:"2.3", lbl:"Juridische bescherming", txt:"Verzet de aanbieder zich actief tegen niet-EU dataverzoeken?" },
                      { k:"4.1", lbl:"EU-infrastructuur",      txt:"Bevindt ook de control plane zich volledig in de EU?" },
                    ].map(r => (
                      <div key={r.k} className="flex gap-2 text-xs">
                        <span className="font-bold flex-shrink-0" style={{ color:"#26B5AE", minWidth:28 }}>{r.k}</span>
                        <div><strong>{r.lbl}:</strong> {r.txt}</div>
                      </div>
                    ))}
                  </div>
                  <p className="text-xs italic" style={{ color:"#9ca3af" }}>Score 1–5 per vraag. Gemiddelde = soevereiniteitsscore op de kleurenbalk (rood → groen).</p>
                </div>
              </div>
            </Section>

            <Section title="Bronnen en team">
              <div className="rounded p-4 mb-3" style={{ background:"#fff", border:"1px solid #D0E4F7" }}>
                <p className="text-sm font-bold mb-2" style={{ color:"#0C2340" }}>📚 Wetenschappelijke en beleidsmatige bronnen</p>
                <p className="text-xs leading-relaxed mb-3" style={{ color:"#374151" }}>
                  Dit instrument is gebouwd op twee erkende en publiek beschikbare normenkaders. Ze vormen de inhoudelijke ruggengraat
                  van alles wat je in deze tool ziet — van de score-knoppen tot de PDF-rapporten. Zonder deze frameworks zou dit instrument
                  niet hebben bestaan.
                </p>
                <div className="space-y-3">
                  <div className="rounded p-3" style={{ background:"#EBF3FF", border:"1px solid #D0E4F7" }}>
                    <div className="flex items-start justify-between gap-2">
                      <div>
                        <p className="text-xs font-bold mb-1" style={{ color:"#1A56A0" }}>
                          Digital Autonomy Assessment Framework (DAAF) — Universiteit Utrecht
                        </p>
                        <p className="text-xs leading-relaxed" style={{ color:"#374151" }}>
                          Ontwikkeld door Tim van Neerbos (lead enterprise architect) en prof. Albert Meijer (Bestuurskunde) aan de Universiteit Utrecht.
                          Het framework meet 22 indicatoren verdeeld over 8 dimensies en 3 niveaus, en is gratis beschikbaar voor alle organisaties.
                          Dit prototype gebruikt de quickscan-variant — de volledige scan wordt in een vervolgfase geïmplementeerd.
                        </p>
                      </div>
                      <a href="https://www.uu.nl/organisatie/digitale-autonomie/tools" target="_blank" rel="noopener noreferrer"
                        style={{ fontSize:11, color:"#1A56A0", textDecoration:"none", border:"1px solid #D0E4F7",
                                 padding:"3px 8px", borderRadius:3, whiteSpace:"nowrap", flexShrink:0 }}>
                        🔗 uu.nl/daaf ↗
                      </a>
                    </div>
                  </div>
                  <div className="rounded p-3" style={{ background:"#f0f9f9", border:"1px solid #26B5AE44" }}>
                    <div className="flex items-start justify-between gap-2">
                      <div>
                        <p className="text-xs font-bold mb-1" style={{ color:"#26B5AE" }}>
                          Toetsingsinstrument Soevereiniteit Clouddiensten — DICTU (Ministerie van EZK)
                        </p>
                        <p className="text-xs leading-relaxed" style={{ color:"#374151" }}>
                          Gepubliceerd door de Dienst ICT Uitvoering (DICTU) van het Ministerie van Economische Zaken en Klimaat in 2026.
                          Het instrument beoordeelt clouddiensten op vijf dimensies: juridisch, data & AI, technologie, operationeel en mens —
                          elk met vijf oplopende soevereiniteitsniveaus. Combineert elementen uit het EuroStack-initiatief en het
                          Cloud Sovereignty Framework van de Europese Commissie.
                        </p>
                      </div>
                      <a href="https://www.dictu.nl/sites/default/files/bestanden/website/DICTU%20Toetsingsinstrument%20Soevereiniteit%20Clouddiensten%20v1.0.1.pdf"
                        target="_blank" rel="noopener noreferrer"
                        style={{ fontSize:11, color:"#26B5AE", textDecoration:"none", border:"1px solid #26B5AE44",
                                 padding:"3px 8px", borderRadius:3, whiteSpace:"nowrap", flexShrink:0 }}>
                        🔗 DICTU-rapport ↗
                      </a>
                    </div>
                  </div>
                </div>
              </div>

              <div className="rounded p-4" style={{ background:"#fff", border:"1px solid #D0E4F7" }}>
                <p className="text-sm font-bold mb-3" style={{ color:"#0C2340" }}>👥 Betrokken bij dit product</p>
                <div className="grid grid-cols-2 gap-3">
                  <div>
                    <p className="text-xs font-semibold mb-2" style={{ color:"#1A56A0" }}>Programma Digitale Samenhang</p>
                    <div className="space-y-1.5">
                      {[
                        { naam:"E. van Gorkum",  rol:"Kwartiermaker Digitale Samenhang — initiatiefnemer en ontwikkelaar" },
                        { naam:"J. Haije",        rol:"Ambassadeur Digitale Soevereiniteit — initiatiefnemer" },
                        { naam:"E. Rolf",         rol:"Ambassadeur Digitale Soevereiniteit — initiatiefnemer en ontwikkelaar" },
                        { naam:"J. Blom",         rol:"Ambassadeur Digitale Soevereiniteit — initiatiefnemer en ontwikkelaar" },
                      ].map(p => (
                        <div key={p.naam} className="flex gap-2 text-xs">
                          <span className="font-semibold flex-shrink-0" style={{ color:"#0C2340", minWidth:120 }}>{p.naam}</span>
                          <span style={{ color:"#6b7280" }}>{p.rol}</span>
                        </div>
                      ))}
                    </div>
                  </div>
                  <div>
                    <p className="text-xs font-semibold mb-2" style={{ color:"#26B5AE" }}>Team Infrastructuur (DLWO) — server & migratie</p>
                    <div className="space-y-1.5">
                      {[
                        { naam:"Jeffrey Klein",   rol:"Infrastructuur engineer — migratie naar DGX Spark" },
                        { naam:"Dolf Keimpema",   rol:"Infrastructuur engineer — migratie naar DGX Spark" },
                      ].map(p => (
                        <div key={p.naam} className="flex gap-2 text-xs">
                          <span className="font-semibold flex-shrink-0" style={{ color:"#0C2340", minWidth:120 }}>{p.naam}</span>
                          <span style={{ color:"#6b7280" }}>{p.rol}</span>
                        </div>
                      ))}
                    </div>
                    <p className="text-xs font-semibold mt-3 mb-1.5" style={{ color:"#1A56A0" }}>Expertisegroep Digitale Soevereiniteit</p>
                    <p className="text-xs mb-3" style={{ color:"#6b7280" }}>
                      Met dank aan de ambassadeurs voor hun scherpte, betrokkenheid en de inspiratie
                      die zij dagelijks meebrengen. Jullie vragen, inzichten en richting hebben dit
                      instrument mede gemaakt tot wat het is — en zijn de drijvende kracht achter
                      de stappen die NHL Stenden zet richting echte digitale soevereiniteit.
                    </p>
                    <p className="text-xs font-semibold mt-1 mb-2" style={{ color:"#6b7280" }}>Ontwikkeling</p>
                    <div className="flex gap-2 text-xs">
                      <span className="font-semibold flex-shrink-0" style={{ color:"#0C2340", minWidth:120 }}>Claude (Anthropic)</span>
                      <span style={{ color:"#6b7280" }}>AI-ontwikkelassistent — volledige applicatiebouw</span>
                    </div>
                  </div>
                </div>
              </div>
            </Section>
          </>}

          {/* ── Tab 2: Aan de slag ── */}
          {aboutTab === "starten" && <>
            <Section title="In vier stappen aan de slag">
              <div className="rounded p-3 mb-4" style={{ background:"#fff", border:"2px solid #1A56A0" }}>
                <p className="text-xs" style={{ color:"#374151" }}>
                  <strong>Nieuw hier?</strong> Volg de stappen hieronder. Je kunt op elk moment stoppen en later verdergaan.
                  Alle data wordt automatisch opgeslagen.
                </p>
              </div>
              <div className="space-y-2">
                {[
                  { n:"1", icon:"➕", title:"Applicatie toevoegen",
                    txt:'Klik op "+ Applicatie toevoegen" in het dashboard of het tabblad Applicaties. Vul naam, leverancier, categorie en eigenaar in. De applicatie wordt direct opgeslagen en verschijnt in het overzicht.',
                    tip:"Begin met de applicaties die je het meest kritisch acht — je kunt er altijd meer toevoegen." },
                  { n:"2", icon:"📋", title:"DAAF Quick Scan invullen",
                    txt:'Open de applicatie en ga naar Stap 1. Je ziet 8 vragen verdeeld over drie niveaus: Risico, Mitigatie en Belang. Klik op de score (1–5) die het best past. Per vraag staat een uitgebreide toelichting en omschrijving per score.',
                    tip:"Voeg een motivatie toe bij elke vraag — dat maakt de score veel waardevoller voor toekomstige reviews." },
                  { n:"3", icon:"🔍", title:"DICTU Soevereiniteitscheck invullen",
                    txt:'Ga naar Stap 2 van het assessment. Vier vragen over dataresidency, technische beveiliging, juridische bescherming en EU-infrastructuur. Scores worden direct zichtbaar in de kleurenbalk op het dashboard.',
                    tip:"Voeg per vraag een motivatietekst toe — die wordt meegenomen in de Excel en PDF export." },
                  { n:"4", icon:"📊", title:"Dashboard en vergelijking lezen",
                    txt:'Na het invullen van meerdere applicaties toont het dashboard het volledige portfolio. Het autonomiekwadrant laat zien welke applicaties urgente aandacht vragen. Het tabblad Vergelijking plaatst alle applicaties naast elkaar.',
                    tip:"Gebruik de filterknopjes op het Dashboard om applicaties te verbergen en zo de vergelijking scherper te maken." },
                ].map(s => (
                  <div key={s.n} className="rounded p-4" style={{ background:"#fff", border:"1px solid #D0E4F7" }}>
                    <div className="flex gap-3">
                      <div className="w-9 h-9 flex items-center justify-center flex-shrink-0 text-white font-bold"
                        style={{ background:"#1A56A0", borderRadius:4, fontSize:15 }}>{s.n}</div>
                      <div className="flex-1">
                        <p className="text-sm font-semibold mb-1" style={{ color:"#0C2340" }}>{s.icon} {s.title}</p>
                        <p className="text-xs leading-relaxed" style={{ color:"#6b7280" }}>{s.txt}</p>
                        <div className="mt-2 rounded px-2 py-1.5 text-xs" style={{ background:"#EBF3FF", color:"#1A56A0" }}>
                          💡 {s.tip}
                        </div>
                      </div>
                    </div>
                  </div>
                ))}
              </div>
            </Section>

            <Section title="Exporteren" accent="#26B5AE">
              <div className="grid grid-cols-2 gap-3">
                <div className="rounded p-4" style={{ background:"#fff", border:"1px solid #D0E4F7" }}>
                  <p className="text-sm font-semibold mb-2" style={{ color:"#0C2340" }}>📥 Excel exporteren</p>
                  <p className="text-xs leading-relaxed mb-2" style={{ color:"#6b7280" }}>
                    Klik op "Exporteer Excel" op het Dashboard. Het bestand bevat vijf tabbladen:
                  </p>
                  {["Overzicht — alle scores per applicatie", "DAAF scores — per indicator inclusief motivatie", "DICTU scores — per vraag inclusief motivatie", "Vragenlijst — alle vragen met normen", "Motivaties — volledig overzicht toelichtingen"].map(t => (
                    <div key={t} className="flex gap-2 text-xs mb-1">
                      <span style={{ color:"#26B5AE", fontWeight:700 }}>→</span>
                      <span style={{ color:"#374151" }}>{t}</span>
                    </div>
                  ))}
                </div>
                <div className="rounded p-4" style={{ background:"#fff", border:"1px solid #D0E4F7" }}>
                  <p className="text-sm font-semibold mb-2" style={{ color:"#0C2340" }}>📄 PDF exporteren</p>
                  <p className="text-xs leading-relaxed mb-2" style={{ color:"#6b7280" }}>
                    Klik op de PDF-knop op het Dashboard. Met de filterknopjes bepaal je of je een enkelvoudig rapport (één app) of een portfoliorapport (meerdere apps) wilt. De PDF bevat:
                  </p>
                  {["Voorblad met naam, datum en versie", "Inhoudsopgave", "Frameworkuitleg (DAAF + DICTU)", "Scores, kwadrantpositie en aanbevelingen per app"].map(t => (
                    <div key={t} className="flex gap-2 text-xs mb-1">
                      <span style={{ color:"#E87722", fontWeight:700 }}>→</span>
                      <span style={{ color:"#374151" }}>{t}</span>
                    </div>
                  ))}
                </div>
              </div>
            </Section>
          </>}

          {/* ── Tab 3: Scores & grafieken ── */}
          {aboutTab === "scores" && <>
            <Section title="Wat betekent de autonomiescore?" accent="#E87722">
              <div className="rounded p-4 mb-3" style={{ background:"#fff", border:"2px solid #E87722" }}>
                <p className="text-sm font-semibold mb-2" style={{ color:"#0C2340" }}>
                  De score meet <em>de urgentie van het autonomieprobleem</em>, niet hoe autonoom een applicatie is.
                </p>
                <p className="text-xs leading-relaxed mb-3" style={{ color:"#374151" }}>
                  De score combineert risico, mitigatie en belang. Twee applicaties met dezelfde eindscore kunnen
                  een heel verschillende situatie beschrijven. Kijk altijd ook naar de drie niveauscores apart.
                </p>
                <div className="grid grid-cols-2 gap-3">
                  <div className="rounded p-3" style={{ background:"#dcfce7", border:"1px solid #86efac" }}>
                    <p className="text-xs font-bold mb-1" style={{ color:"#15803d" }}>✓ Hoge score (7+) — geen acute actie</p>
                    <p className="text-xs" style={{ color:"#166534" }}>
                      Risico's zijn goed gemitigeerd (alternatieven beschikbaar, sterke contracten, kennis aanwezig),
                      óf de applicatie heeft weinig strategisch belang. Monitor periodiek.
                    </p>
                  </div>
                  <div className="rounded p-3" style={{ background:"#fee2e2", border:"1px solid #fca5a5" }}>
                    <p className="text-xs font-bold mb-1" style={{ color:"#b91c1c" }}>✗ Lage score (&lt;5) — actie vereist</p>
                    <p className="text-xs" style={{ color:"#7f1d1d" }}>
                      Combinatie van hoog risico, zwakke mitigatie en/of hoog belang.
                      Kies: <strong>migreren</strong>, <strong>mitigeren</strong> of <strong>bewust accepteren</strong> met besluitvorming.
                    </p>
                  </div>
                </div>
              </div>
              <Tip label="Waar zit de verbetermarge?" color="#1A56A0" bg="#EBF3FF">
                Strategisch belang (Niveau 3) ligt grotendeels vast. Verbeteringen zitten in Risico verlagen
                (bijv. migreren naar EU-aanbieder) en Mitigatie verhogen (alternatieven ontwikkelen, contracten
                versterken, kennis opbouwen). Gebruik de drie niveauscores om te bepalen waar de grootste winst zit.
              </Tip>
            </Section>

            <Section title="Hoe lees je het dimensieprofiel?" accent="#1A56A0">
              <div className="rounded p-4 mb-3" style={{ background:"#EBF3FF", border:"1px solid #D0E4F7" }}>
                <p className="text-xs leading-relaxed mb-3" style={{ color:"#374151" }}>
                  Het dimensieprofiel toont per as een gekleurde balk. Anders dan een spindiagram heeft elke balk
                  een eigen richting: "hoge score" betekent niet overal hetzelfde. Hover over een punt voor de exacte waarde.
                </p>
                <div className="grid grid-cols-3 gap-3 mb-3">
                  {[
                    { letter:"A, B", naam:"Risico-assen", kleur:"#dc2626", bg:"#fee2e2",
                      uitleg:"Geopolitiek risico (A) en Leveranciersafhankelijkheid (B). Groen = laag risico, rood = hoog risico. Links is beter." },
                    { letter:"C, D, E", naam:"Mitigatie-assen", kleur:"#16a34a", bg:"#dcfce7",
                      uitleg:"Technische (C), Organisatorische (D) en Contractuele (E) weerbaarheid. Rood = weinig weerbaarheid, groen = sterk. Rechts is beter." },
                    { letter:"F, G, H", naam:"Belang-assen", kleur:"#d97706", bg:"#ffedd5",
                      uitleg:"Operationeel (F), Data (G) en Academisch belang (H). Hoge scores zijn niet slecht, maar verhogen de urgentie van risico en mitigatie." },
                  ].map(g => (
                    <div key={g.letter} className="rounded p-3" style={{ background:g.bg, border:`1px solid ${g.kleur}44` }}>
                      <div className="flex items-center gap-1.5 mb-1.5">
                        <span className="text-xs font-bold px-1.5 py-0.5 rounded" style={{ background:g.kleur, color:"#fff", fontSize:9 }}>{g.letter}</span>
                        <span className="text-xs font-bold" style={{ color:g.kleur }}>{g.naam}</span>
                      </div>
                      <p style={{ fontSize:9.5, color:"#374151", lineHeight:1.5 }}>{g.uitleg}</p>
                    </div>
                  ))}
                </div>
                <div className="rounded p-3 mb-3" style={{ background:"#fff", border:"1px solid #D0E4F7" }}>
                  <p className="text-xs font-bold mb-2" style={{ color:"#0C2340" }}>Kleurovergang per type as</p>
                  <div className="grid grid-cols-2 gap-3">
                    {[
                      { titel:"Risico-assen (A, B)", gradient:"linear-gradient(to right,#dcfce7,#fef9c3,#fca5a5,#dc2626)", tekst:"Groen links = laag risico (goed). Rood rechts = hoog risico." },
                      { titel:"Mitigatie-assen (C, D, E)", gradient:"linear-gradient(to right,#dc2626,#fca5a5,#fde68a,#86efac,#16a34a)", tekst:"Rood links = weinig weerbaarheid. Groen rechts = sterk weerbaar." },
                    ].map(g => (
                      <div key={g.titel}>
                        <p style={{ fontSize:10, fontWeight:600, color:"#0C2340", marginBottom:4 }}>{g.titel}</p>
                        <div style={{ height:16, borderRadius:3, background:g.gradient, marginBottom:4 }} />
                        <p style={{ fontSize:9, color:"#6b7280" }}>{g.tekst}</p>
                      </div>
                    ))}
                  </div>
                </div>
                <div className="rounded p-3" style={{ background:"#fff", border:"1px solid #D0E4F7" }}>
                  <p className="text-xs font-bold mb-2" style={{ color:"#0C2340" }}>Patronen om op te letten</p>
                  <div className="grid grid-cols-2 gap-2">
                    {[
                      { titel:"Punten dicht bij elkaar", tekst:"Vergelijkbaar profiel op die dimensie — gemeenschappelijk risico voor meerdere applicaties." },
                      { titel:"Punten ver uit elkaar",   tekst:"Fundamenteel verschil tussen applicaties — prioriteer de applicatie met het slechtste punt." },
                      { titel:"Risico hoog + mitigatie laag", tekst:"Hoog risico gecombineerd met lage weerbaarheid — dit vraagt directe actie." },
                      { titel:"Belang hoog + risico zwak",    tekst:"Hoge belang-score maakt een slechte risico- of mitigatiescore urgenter." },
                    ].map(t => (
                      <div key={t.titel} className="rounded p-2" style={{ background:"#f8fafc", border:"1px solid #e5e7eb" }}>
                        <p style={{ fontSize:10, fontWeight:700, color:"#0C2340", marginBottom:3 }}>📌 {t.titel}</p>
                        <p style={{ fontSize:9, color:"#6b7280", lineHeight:1.5 }}>{t.tekst}</p>
                      </div>
                    ))}
                  </div>
                </div>
              </div>
            </Section>

            <Section title="De berekening in vier stappen" accent="#26B5AE">
              <div className="rounded p-4" style={{ background:"#fff", border:"1px solid #D0E4F7" }}>
                <div className="grid grid-cols-4 gap-3 mb-3">
                  {[
                    { n:"1", lbl:"Dimensiescore", txt:"Per dimensie (A t/m H): gemiddelde van de ingevulde indicatorscores (1–5).", c:"#1A56A0" },
                    { n:"2", lbl:"Niveauscore",   txt:"Per niveau (Risico, Mitigatie, Belang): gemiddelde van de dimensiescores.", c:"#26B5AE" },
                    { n:"3", lbl:"Ruwe score",    txt:"Mitigatie ÷ (Risico × Belang). Bereik: 0,04 (slechtst) tot 5,0 (best).", c:"#E87722" },
                    { n:"4", lbl:"Normalisatie",  txt:"Logaritmische schaal zet de ruwe score om naar 1–10, zodat scores goed vergelijkbaar zijn.", c:"#6d28d9" },
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
                  een score tussen 1 en 3 krijgen. De logaritmische schaal spreidt scores uit over het volledige bereik 1–10.
                </div>
              </div>
            </Section>
          </>}

          {/* ── Tab 4: Tips & beheer ── */}
          {aboutTab === "tips" && <>
            <Section title="Praktische tips voor het team">
              <div className="grid grid-cols-2 gap-3">
                <Tip label="💾 Data opslaan" color="#1A56A0" bg="#EBF3FF">
                  Alle data wordt automatisch opgeslagen op de server (Netlify Blobs). Iedereen die inlogt ziet
                  dezelfde data, ongeacht browser of apparaat. Je kunt het venster sluiten en later verdergaan.
                  Exporteer regelmatig een Excel-bestand als extra back-up.
                </Tip>
                <Tip label="👥 Meerdere beoordelaars" color="#26B5AE" bg="#E6F7F7">
                  Laat beoordelaars het assessment onafhankelijk invullen en exporteer afzonderlijk naar Excel.
                  Bespreek grote afwijkingen in het team — die leveren vaak de meest waardevolle inzichten op.
                </Tip>
                <Tip label="📋 Welke vragen gebruik je?" color="#E87722" bg="#fff8e1">
                  Deze tool gebruikt een selectie: de DAAF Quick Scan (8 kernindicatoren, één per dimensie) en
                  4 relevante DICTU-vragen. Voor een volledig DAAF assessment met wegingen raadpleeg de Utrecht University tool.
                </Tip>
                <Tip label="🎯 Wat doe je met de uitkomst?" color="#6d28d9" bg="#faf5ff">
                  Bespreek het kwadrant. Applicaties rechtsonder (KRITIEK) vragen urgente actie.
                  Kies per applicatie: migreren, mitigeren of bewust accepteren. Leg de keuze vast.
                  Herhaal het assessment na contractwijzigingen of leveranciersveranderingen.
                </Tip>
              </div>
            </Section>

            <Section title="Beheeromgeving" accent="#0C2340">
              <div className="rounded p-4 mb-3" style={{ background:"#fff", border:"1px solid #D0E4F7" }}>
                <p className="text-xs leading-relaxed mb-3" style={{ color:"#374151" }}>
                  Via het tabblad <strong>🔐 Beheer</strong> (pincode vereist) kun je:
                </p>
                <div className="grid grid-cols-2 gap-2 mb-3">
                  {[
                    { icon:"✏️", txt:"Applicatiegegevens aanpassen (naam, leverancier, categorie, eigenaar)" },
                    { icon:"🗑️", txt:"Applicaties verwijderen uit het portfolio" },
                    { icon:"👁️", txt:"Volledig overzicht per applicatie bekijken inclusief alle scores en motivaties" },
                    { icon:"🔒", txt:"Anonieme presentatienamen toewijzen voor de Discreet-modus (toggle 🔓 Open / 🔒 Discreet)" },
                  ].map(t => (
                    <div key={t.txt} className="flex gap-2 rounded p-2.5" style={{ background:"#f8fafc", border:"1px solid #e5e7eb" }}>
                      <span style={{ fontSize:14, flexShrink:0 }}>{t.icon}</span>
                      <p className="text-xs" style={{ color:"#374151", lineHeight:1.5 }}>{t.txt}</p>
                    </div>
                  ))}
                </div>
                <div className="rounded p-2.5 text-xs" style={{ background:"#fffbeb", border:"1px solid #fde68a", color:"#92400e" }}>
                  <strong>Helder / Discreet:</strong> gebruik de toggle in de header om te wisselen tussen de echte namen (🔓 Helder) en anonieme presentatienamen (🔒 Discreet). Handig bij presentaties aan externen. De toggle verschijnt alleen als er anonieme namen zijn ingesteld via de beheeromgeving.
                </div>
              </div>

              {/* Import/Export sectie */}
              <div className="rounded p-4" style={{ background:"#fff", border:"1px solid #D0E4F7" }}>
                <p className="text-xs font-bold mb-3" style={{ color:"#0C2340" }}>📦 Database export en import</p>
                <div className="grid grid-cols-2 gap-3 mb-3">
                  <div className="rounded p-3" style={{ background:"#f0fdf4", border:"1px solid #86efac" }}>
                    <p className="text-xs font-bold mb-1" style={{ color:"#15803d" }}>⬇ Database exporteren</p>
                    <p className="text-xs leading-relaxed mb-2" style={{ color:"#374151" }}>
                      Klik op <strong>"Exporteer database"</strong> in de beheeromgeving. Er wordt een
                      JSON-bestand gedownload met <em>alle</em> applicaties inclusief scores, motivaties,
                      aanmaakdatum en metadata. Het bestand heeft altijd een datum en tijd in de naam, 
                      bijvoorbeeld <code style={{ background:"#dcfce7", padding:"0 3px", borderRadius:2 }}>NHL_Sov_Database_20260603_1423.json</code>.
                    </p>
                    <p className="text-xs" style={{ color:"#16a34a", fontWeight:600 }}>
                      Sla dit bestand op een veilige locatie op als back-up.
                    </p>
                  </div>
                  <div className="rounded p-3" style={{ background:"#EBF3FF", border:"1px solid #D0E4F7" }}>
                    <p className="text-xs font-bold mb-1" style={{ color:"#1A56A0" }}>⬆ Database importeren</p>
                    <p className="text-xs leading-relaxed" style={{ color:"#374151" }}>
                      Klik op <strong>"Importeer database"</strong> en selecteer een eerder geëxporteerd
                      JSON-bestand. Er verschijnt een overzicht van alle applicaties in het bestand.
                      Per applicatie zie je of het een nieuwe toevoeging is of een bestaande die wordt overschreven.
                    </p>
                  </div>
                </div>
                <div className="rounded p-3 mb-2" style={{ background:"#f8fafc", border:"1px solid #e5e7eb" }}>
                  <p className="text-xs font-bold mb-2" style={{ color:"#0C2340" }}>Selectief importeren</p>
                  <div className="space-y-1.5">
                    {[
                      { stap:"1", txt:"Open het importvenster via de knop in de beheeromgeving." },
                      { stap:"2", txt:"Je ziet alle applicaties uit het exportbestand, elk met scores en status (nieuw of overschrijft bestaande)." },
                      { stap:"3", txt:'Gebruik "Alles selecteren" voor een volledige herstel, of klik individuele applicaties aan voor een selectieve import.' },
                      { stap:"4", txt:"Klik op de groene importknop. Nieuwe applicaties worden toegevoegd, bestaande worden overschreven met de importdata." },
                    ].map(s => (
                      <div key={s.stap} className="flex gap-2 text-xs">
                        <div className="w-4 h-4 rounded-full flex items-center justify-center text-white font-bold flex-shrink-0"
                          style={{ background:"#1A56A0", fontSize:9 }}>{s.stap}</div>
                        <span style={{ color:"#374151", lineHeight:1.5 }}>{s.txt}</span>
                      </div>
                    ))}
                  </div>
                </div>
                <div className="rounded p-2.5 text-xs" style={{ background:"#fffbeb", border:"1px solid #fde68a", color:"#92400e" }}>
                  <strong>Per ongeluk verwijderd?</strong> Importeer het meest recente exportbestand en selecteer
                  alleen de verwijderde applicatie. De rest van het portfolio blijft ongewijzigd.
                </div>
              </div>
            </Section>
          </>}

          {/* Footer — altijd zichtbaar */}
          <div className="rounded p-3 text-center mt-4" style={{ background:"#fff", border:"1px solid #D0E4F7" }}>
            <p className="text-xs" style={{ color:"#9ca3af" }}>
              NHL Stenden Hogeschool · Project Digitale Soevereiniteit · Ambassadeurslijn Digitale Soevereiniteit
               · {VERSION}
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
              Prototype · Ambassadeurslijn Digitale Soevereiniteit
            </p>

          </div>
        </div>
        <div className="flex items-center gap-3">
          {/* App-teller + laatste opslag */}
          <div className="flex items-center gap-1.5 text-xs px-2.5 py-1 rounded"
            style={{ background:"rgba(255,255,255,0.1)", color:"#7DD3D0", border:"1px solid rgba(255,255,255,0.1)" }}>
            <span style={{ fontSize:12 }}>🗂</span>
            <span style={{ fontWeight:700 }}>{apps.length}</span>
            <span style={{ opacity:0.75 }}>applicatie{apps.length !== 1 ? "s" : ""}</span>
            {lastSaved && (
              <span style={{ opacity:0.6, fontSize:10, borderLeft:"1px solid rgba(255,255,255,0.2)", paddingLeft:6, marginLeft:2 }}>
                {(() => {
                  const d = new Date(lastSaved);
                  const now = new Date();
                  const sameDay = d.toDateString() === now.toDateString();
                  return sameDay
                    ? `opgeslagen ${d.toLocaleTimeString("nl-NL", {hour:"2-digit",minute:"2-digit"})}`
                    : `opgeslagen ${d.toLocaleDateString("nl-NL", {day:"numeric",month:"short"})} ${d.toLocaleTimeString("nl-NL", {hour:"2-digit",minute:"2-digit"})}`;
                })()}
              </span>
            )}
          </div>
          {/* Opslaan status */}
          {saving && (
            <span className="text-xs flex items-center gap-1.5 px-2 py-1 rounded"
              style={{ background:"rgba(251,191,36,0.2)", color:"#fbbf24" }}>
              <span className="animate-pulse">●</span> Opslaan…
            </span>
          )}
          {saveError && (
            <span className="text-xs px-2 py-1 rounded" style={{ background:"rgba(220,38,38,0.3)", color:"#fca5a5" }}>
              ⚠️ Opslaan mislukt
            </span>
          )}
          {!saving && !saveError && ready && (
            <span className="text-xs flex items-center gap-1.5 px-2 py-1 rounded"
              style={{ background:"rgba(38,181,174,0.15)", color:"#7DD3D0" }}>
              <span style={{ color:"#4ade80", fontSize:9 }}>●</span> Gesynchroniseerd
            </span>
          )}
          {/* Naamweergave toggle switch */}
          {apps.some(a => a.nameSecondary) && (
            <button onClick={() => setUseSecondaryName(p => !p)}
              className="flex items-center gap-2 text-xs font-medium px-3 py-1.5"
              style={{
                borderRadius: 20,
                border: `1px solid ${useSecondaryName ? "rgba(232,119,34,0.6)" : "rgba(255,255,255,0.2)"}`,
                background: useSecondaryName ? "rgba(232,119,34,0.2)" : "rgba(255,255,255,0.08)",
                color: useSecondaryName ? "#f8b87a" : "#94a3b8",
                transition: "all 0.2s",
              }}
              title={useSecondaryName ? "Schakel naar open weergave" : "Schakel naar discrete weergave"}>
              <span style={{ fontSize:11 }}>{useSecondaryName ? "🔒" : "🔓"}</span>
              {/* Schuifje */}
              <div style={{
                width: 28, height: 16, borderRadius: 8, position: "relative",
                background: useSecondaryName ? "#E87722" : "rgba(255,255,255,0.2)",
                transition: "background 0.2s", flexShrink: 0
              }}>
                <div style={{
                  width: 12, height: 12, borderRadius: 6,
                  background: "white", position: "absolute",
                  top: 2, transition: "left 0.2s",
                  left: useSecondaryName ? 14 : 2,
                  boxShadow: "0 1px 3px rgba(0,0,0,0.3)"
                }}/>
              </div>
              <span style={{ fontSize:10, letterSpacing: "0.02em" }}>
                {useSecondaryName ? "Discreet" : "Open"}
              </span>
            </button>
          )}

        </div>
      </header>

      {/* Sub-header nav — lichte blauwe balk (NHS Stenden stijl) */}
      <div className="flex-shrink-0" style={{ background:"#1A56A0" }}>
        <nav className="px-5 flex gap-0">
          {[
            { k:"dashboard", label:"Dashboard" },
            { k:"apps",      label:"Applicaties" },
            ...(selApp ? [{ k:"assess", label:displayName(selApp).substring(0,20) }] : []),
            { k:"compare",   label:"Vergelijking" },
            { k:"bestuur",   label:"📊 Portfolio" },
            { k:"about",     label:"ℹ️ Over & uitleg" },
            { k:"transparantie", label:"🔍 Over dit product" },
            { k:"SPACER",    label:"" },
            { k:"CHANGELOG", label:"📋 Changelog" },
            { k:"admin",     label:"🔐 Beheer" },
          ].map(t => {
            if (t.k === "SPACER") return <div key="spacer" className="flex-1"/>;
            if (t.k === "CHANGELOG") return (
              <button key={t.k} onClick={() => setShowChangelog(true)}
                className="px-3 py-2.5 text-sm font-medium border-b-2 transition-colors whitespace-nowrap"
                style={{ borderColor:"transparent", color:"rgba(255,255,255,0.5)", fontSize:12 }}>
                {t.label}
              </button>
            );
            if (t.k === "admin") return (
              <button key={t.k} onClick={() => setView(t.k)}
                className="px-4 py-2.5 text-sm font-medium border-b-2 transition-colors whitespace-nowrap ml-1"
                style={view === t.k
                  ? { borderColor:"#26B5AE", color:"#fff", background:"rgba(255,255,255,0.18)", borderRadius:"4px 4px 0 0" }
                  : { borderColor:"rgba(255,255,255,0.2)", color:"rgba(255,255,255,0.6)",
                      background:"rgba(255,255,255,0.07)", borderRadius:"4px 4px 0 0" }}>
                {t.label}
              </button>
            );
            return (
              <button key={t.k} onClick={() => setView(t.k)}
                className="px-4 py-2.5 text-sm font-medium border-b-2 transition-colors whitespace-nowrap"
                style={view === t.k
                  ? { borderColor:"#26B5AE", color:"#fff", background:"rgba(255,255,255,0.12)" }
                  : { borderColor:"transparent", color:"rgba(255,255,255,0.75)" }}>
                {t.label}
              </button>
            );
          })}
        </nav>
      </div>

      {/* Content */}
      <main className="flex-1 overflow-hidden" style={{ background:"#EBF3FF" }}>
        {view === "dashboard" && Dashboard()}
        {view === "apps"      && AppsList()}
        {view === "assess"    && Assess()}
        {view === "compare"   && Compare()}
        {view === "bestuur"   && Bestuur()}
        {view === "about"     && About()}
        {view === "transparantie" && Transparantie()}
        {view === "admin"     && Admin()}
      </main>

      {/* ── Changelog modal — root level zodat hij altijd beschikbaar is ── */}
      {showChangelog && (
        <div className="fixed inset-0 flex items-center justify-center z-50"
          style={{ background:"rgba(12,35,64,0.6)" }}
          onClick={() => setShowChangelog(false)}>
          <div className="bg-white w-full max-w-lg mx-4 overflow-hidden"
            style={{ borderRadius:8, boxShadow:"0 8px 32px rgba(12,35,64,0.4)", maxHeight:"80vh" }}
            onClick={e => e.stopPropagation()}>
            <div className="px-5 py-3 flex items-center justify-between"
              style={{ background:"#0C2340", borderBottom:"3px solid #26B5AE" }}>
              <div className="flex items-center gap-2">
                <span style={{ fontSize:16 }}>📋</span>
                <h2 className="font-bold text-white text-sm">Changelog</h2>
                <span className="text-xs px-2 py-0.5 rounded" style={{ background:"#26B5AE", color:"#0C2340", fontWeight:700 }}>
                  {VERSION}
                </span>
                <span className="text-xs" style={{ color:"#7DD3D0" }}>Overzicht van alle versies en wijzigingen</span>
              </div>
              <button onClick={() => setShowChangelog(false)}
                className="text-white hover:text-gray-300" style={{ fontSize:20, lineHeight:1 }}>×</button>
            </div>
            <div className="overflow-y-auto" style={{ maxHeight:"calc(80vh - 52px)" }}>
              {CHANGELOG.map((v, vi) => (
                <div key={v.versie} style={{ borderBottom:"1px solid #D0E4F7" }}>
                  <div className="flex items-center gap-3 px-5 py-2.5"
                    style={{ background: vi === 0 ? "#EBF3FF" : "#f8fafc" }}>
                    <span className="text-sm font-bold px-2.5 py-0.5"
                      style={{ background: vi === 0 ? "#1A56A0" : "#e5e7eb",
                               color: vi === 0 ? "white" : "#374151", borderRadius:4 }}>
                      {v.versie}
                    </span>
                    <span className="text-xs" style={{ color:"#9ca3af" }}>{v.datum}</span>
                    {vi === 0 ? <span className="text-xs font-semibold" style={{ color:"#26B5AE" }}>Huidige versie</span> : ""}
                  </div>
                  <ul className="px-5 py-2 space-y-1">
                    {v.wijzigingen.map((w, wi) => (
                      <li key={wi} className="flex gap-2 text-xs" style={{ color:"#374151" }}>
                        <span style={{ color:"#26B5AE", fontWeight:700, flexShrink:0 }}>→</span>
                        <span>{w}</span>
                      </li>
                    ))}
                  </ul>
                </div>
              ))}
            </div>
          </div>
        </div>
      )}

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
                  { k:"name",          l:"Primaire naam *",       p:"bijv. Microsoft 365",         hint:"De officiële applicatienaam" },
                  { k:"nameSecondary", l:"Secundaire naam",        p:"bijv. Productiviteitsplatform", hint:"Alternatieve weergavenaam (optioneel)", bg:"#fffbeb" },
                  { k:"supplier",      l:"Leverancier",            p:"bijv. Microsoft",              hint:"" },
                  { k:"cat",           l:"Categorie",              p:"bijv. Productiviteit, ERP, SIS",hint:"" },
                  { k:"owner",         l:"Applicatie-eigenaar",    p:"bijv. Functioneel beheerder",  hint:"" },
                ].map(f => (
                  <div key={f.k}>
                    <label className="text-xs font-semibold block mb-0.5" style={{ color:"#0C2340" }}>{f.l}</label>
                    {f.hint && <p className="text-xs mb-1" style={{ color:"#9ca3af" }}>{f.hint}</p>}
                    <input value={form[f.k] || ""} onChange={e => setForm(p => ({ ...p, [f.k]: e.target.value }))}
                      placeholder={f.p}
                      className="w-full border px-3 py-2 text-sm focus:outline-none"
                      style={{ borderColor:"#D0E4F7", borderRadius:4, background: f.bg || "white" }}
                      onFocus={e => e.target.style.borderColor="#1A56A0"}
                      onBlur={e => e.target.style.borderColor="#D0E4F7"}
                    />
                  </div>
                ))}
                <div>
                  <label className="text-xs font-semibold block mb-1" style={{ color:"#0C2340" }}>Toelichting</label>
                  <textarea value={form.appNotes} onChange={e => setForm(p => ({ ...p, appNotes: e.target.value }))}
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

export default App;
