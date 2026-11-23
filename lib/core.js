/* Sopralluoghi – logica comune (parsing file, verifiche personale, xlsx). Funziona nel browser e in Node. */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.Core = factory();
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  /* ---------- voci del modulo ASPI (testo identico al modello Word) ---------- */
  const CHECK = [
    { s: "A", t: "Controllo Documentale", v: [
      ["A.1", "P.O.S./Verbali coordinamento presenti in cantiere"],
      ["A.2", "Personale in cantiere con autorizzazione a manovra in corso di validità"],
      ["A.3", "Personale in cantiere con autorizzazione all'apertura di cantieri (art.21)"],
      ["A.4", "Applicazione della routine “3 minuti per la Sicurezza”"],
      ["A.5", "Presenza moduli cartacei della routine “3 minuti per la Sicurezza”"]] },
    { s: "B", t: "Controllo Dispositivi Protezione (D.P.I.)", v: [
      ["B.1", "D.P.I. utilizzati in maniera conforme"],
      ["B.2", "D.P.I. a disposizione del personale adeguati alle attività svolte (e.g. AV in classe 3)"],
      ["B.3", "D.P.C. a disposizione del personale adeguati alle attività svolte"]] },
    { s: "C", t: "Controllo Corretta Installazione Segnaletica", v: [
      ["C.1", "Segnaletica e presegnalamento conformi agli schemi previsti dal disciplinare"],
      ["C.2", "Visibilità segnaletica / adeguata installazione (NC in caso di: es. cartelli ruotati, illuminazione non adeguata, limitazione visibilità causa usura)"],
      ["C.3", "Adeguata sorveglianza della segnaletica installata (segnaletica/coni abbattuti, lampade blitz Non funzionanti o mal posizionate)"],
      ["C.4", "Segnaletica stradale di cantiere adeguata alle esigenze/obblighi di percorrenza di cantiere, anche in caso di emergenza"],
      ["C.5", "Adeguata illuminazione dell’area di lavoro (con torre faro) per le attività notturne"]] },
    { s: "D", t: "Esecuzione delle Attività", v: [
      ["D.1", "Corretta gestione house-keeping delle aree di lavoro"],
      ["D.2", "Apparecchiature elettriche conformi (Cavi prolunga, prese e spine, impianti di illuminazione trasportabili, ecc.)"],
      ["D.3", "Corretta esecuzione delle attività, così come da indicazioni del RUP/DL/DEC/RTC"],
      ["D.4", "Corretta compilazione e applicazione check-list “Allegato 1” delle “linee guida per il supporto nella gestione delle bonifiche nelle cantierizzazioni con compresenza di più imprese - prot. ASPI/T5/2026/0000179/EU del 19/06/2026”"]] },
    { s: "E", t: "Osservanza delle Procedure di Sicurezza Generali", v: [
      ["E.1", "Lavorazioni svolte con la copertura di idonee protezioni anche costituite dall'autocarro attrezzato o macchina operatrice purché di massa adeguata a tale funzione (le protezioni suddette devono precedere di 50 metri la zona di lavoro)"],
      ["E.2", "Mezzi dotati di dispositivi luminosi e cicalini retromarcia funzionanti"],
      ["E.3", "Autoveicoli in sosta con ruote girate verso il bordo esterno della carreggiata"],
      ["E.4", "Personale a terra al di fuori del raggio d’azione dei mezzi"],
      ["E.5", "Personale a terra al di dall’area di movimentazione di carichi sospesi"],
      ["E.6", "Corretta esecuzione delle attività in quota"],
      ["E.7", "Adeguata gestione delle interferenze con altre ditte e con la viabilità di cantiere (delimitazione/segnalazione delle aree di lavoro)"],
      ["E.8", "Manovre dei mezzi d’opera facilitate da moviere a terra"],
      ["E.9", "Disponibilità di acqua e di zona d’ombra per fronteggiare le temperature elevate (lavori all’aperto, quando applicabile)"]] }
  ];
  const ITEMS = CHECK.flatMap(s => s.v);
  const AUTOSTRADE = ["A1", "A12", "D18", "D19", "EXT"];
  const CARREGGIATE = ["NORD", "SUD", "SVINCOLO"];
  const CORSIE = ["E", "ML", "MV", "S", "DEVIAZIONE", "CHIUSURA SVINCOLO"];
  const RIGHE_IMPRESA = [["NOME", "Nome impresa"], ["PREPOSTO", "Preposto"], ["EMERG", "Addetto emergenze"],
    ...[1, 2, 3, 4, 5, 6, 7].map(n => ["OP" + n, "Operatore " + n])];

  /* ---------- date ---------- */
  const pad = n => String(n).padStart(2, "0");
  const isoLocal = dt => `${dt.getFullYear()}-${pad(dt.getMonth() + 1)}-${pad(dt.getDate())}`;
  const today = () => isoLocal(new Date());
  const parseISO = s => { const [y, m, d] = String(s).split("-").map(Number); return new Date(y, m - 1, d); };
  const diffDays = (a, b) => Math.round((parseISO(a) - parseISO(b)) / 86400000);
  const fmtD = s => { if (!s) return ""; const [y, m, d] = s.split("-"); return `${d}/${m}/${y}`; };
  const serialToISO = v => { const dt = new Date(Math.round((Math.floor(v) - 25569) * 86400000)); return `${dt.getUTCFullYear()}-${pad(dt.getUTCMonth() + 1)}-${pad(dt.getUTCDate())}`; };
  const serialToHM = v => { const m = Math.round((v - Math.floor(v)) * 1440) % 1440; return `${pad(Math.floor(m / 60))}:${pad(m % 60)}`; };
  function dateInText(s) { // "INDETERMINATO - DIST 30/6/26" -> 2026-06-30
    const m = String(s).match(/(\d{1,2})\/(\d{1,2})\/(\d{2,4})/g); if (!m) return null;
    const last = m[m.length - 1].split("/"); let y = +last[2]; if (y < 100) y += 2000;
    return `${y}-${pad(+last[1])}-${pad(+last[0])}`;
  }
  const norm = s => String(s || "").toUpperCase().normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/[^A-Z0-9 ]/g, " ").replace(/\s+/g, " ").trim();
  const normImp = s => norm(s).replace(/\b(SRL|SRLS|SPA|S P A|S R L|SNC|SAS|SOC|COOP)\b/g, " ").replace(/\s+/g, " ").trim();

  /* ---------- file verifiche ITP ---------- */
  const ITP_COLS = {
    nome: /ADDETTO/, baw: /^AUT\.?\s*BAW/, impSub: /DISTACCATARIA/, distacco: /^DISTACCO$/, subappalto: /^SUBAPPALTO/,
    nolo: /^NOLO/, impApp: /DISTACCANTE/, gen: /FORMAZIONE GENERALE/, preposto: /PREPOSTO/, segn: /SEGNALETICA/,
    aai: /ANTINCENDIO/, ps: /PRIMO SOCCORSO/, ple: /\bPLE\b/, bybridge: /BY.?BRIDGE/, funi: /FUNI/, quota: /IN QUOTA/,
    dpi3: /DPI 3/, spazi: /CONFINATI/, unilav: /^UNILAV/, fineRapp: /FINE RAPP/, idSan: /IDONEIT/, consDpi: /CONS DPI/,
    tesserino: /TESSERINO/, rilascio: /RILASCIO AUT/, note: /^NOTE/, contratto: /^N\s*°?\s*CONTRATTO|^N\.\s*CONTRATTO/, scadAut: /SCADENZA AUT/
  };
  const DOC_LABELS = { gen: "Form. generale", preposto: "Preposto", segn: "Segnaletica", aai: "Antincendio", ps: "Primo soccorso",
    ple: "PLE", bybridge: "By-bridge", funi: "Funi", quota: "Lavori in quota", dpi3: "DPI 3ª cat.", spazi: "Spazi confinati",
    unilav: "UNILAV", fineRapp: "Rapporto di lavoro", idSan: "Idoneità sanitaria", consDpi: "Consegna DPI", tesserino: "Tesserino" };
  function cellVal(v) {
    if (v === null || v === undefined) return null;
    if (typeof v === "number") return v > 20000 && v < 80000 ? { d: serialToISO(v) } : String(v);
    if (v instanceof Date) return { d: isoLocal(v) };
    const s = String(v).replace(/\s+/g, " ").trim();
    return s === "" ? null : s;
  }
  function parseITP(XLSX, buf) {
    const wb = XLSX.read(buf, { type: "array" });
    const sn = wb.SheetNames.find(n => { const ws = wb.Sheets[n]; return ws && ws.A1 && /ADDETTO/i.test(String(ws.A1.v)); }) || wb.SheetNames[0];
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, raw: true, defval: null, blankrows: false });
    const hdr = (rows[0] || []).map(h => norm(h));
    const idx = {};
    for (const [k, re] of Object.entries(ITP_COLS)) {
      const i = hdr.findIndex((h, j) => re.test(h) && !Object.values(idx).includes(j));
      if (i >= 0) idx[k] = i;
    }
    if (idx.nome === undefined) throw new Error("Colonna 'ADDETTO DA AUTORIZZARE' non trovata nel primo foglio.");
    const operators = [];
    rows.slice(1).forEach((r, i) => {
      const nome = r[idx.nome] ? String(r[idx.nome]).replace(/\s+/g, " ").trim().toUpperCase() : "";
      if (!nome) return;
      const o = { id: "r" + (i + 2), riga: i + 2, nome };
      for (const k of Object.keys(ITP_COLS)) if (k !== "nome") o[k] = idx[k] !== undefined ? cellVal(r[idx[k]]) : null;
      for (const k of ["gen", "preposto", "segn", "aai", "ps", "ple", "bybridge", "funi", "quota", "dpi3", "spazi", "idSan", "scadAut", "fineRapp"]) {
        if (typeof o[k] === "string") { const d = dateInText(o[k]); if (d) o[k] = { d, raw: o[k] }; }
      }
      if (o.contratto && typeof o.contratto === "object") o.contratto = null;
      o.baw = o.baw && typeof o.baw === "object" ? null : o.baw;
      operators.push(o);
    });
    const impSet = new Set();
    operators.forEach(o => [o.impSub, o.impApp].forEach(x => { if (typeof x === "string" && !/^N\/?A$/i.test(x)) impSet.add(x); }));
    return { foglio: sn, operators, imprese: [...impSet].sort() };
  }
  function trovaOperatore(nome, operators) {
    const n = norm(nome); if (!n) return null;
    let hit = operators.filter(o => norm(o.nome) === n);
    if (!hit.length) { const rev = n.split(" ").reverse().join(" "); hit = operators.filter(o => norm(o.nome) === rev); }
    if (!hit.length) {
      const toks = n.split(" ").filter(t => t.length > 2);
      if (toks.length) hit = operators.filter(o => { const on = norm(o.nome).split(" "); return toks.every(t => on.includes(t)); });
    }
    return hit.length === 1 ? hit[0] : (hit.length > 1 ? hit.sort((a, b) => (b.scadAut?.d || "").localeCompare(a.scadAut?.d || ""))[0] : null);
  }

  /* ---------- verifica personale ---------- */
  // ruolo: "preposto" | "emergenze" | "operatore"
  function stato(val, D, { req = false, warnDays = 30 } = {}) {
    if (val && typeof val === "object" && val.d) {
      const n = diffDays(val.d, D);
      if (n < 0) return { lv: "bad", txt: "scad. " + fmtD(val.d), scad: val.d };
      if (n <= warnDays) return { lv: "warn", txt: "scade " + fmtD(val.d), scad: val.d };
      return { lv: "ok", txt: fmtD(val.d) };
    }
    if (val === null || val === undefined) return req ? { lv: "bad", txt: "mancante" } : { lv: "none", txt: "—" };
    const s = String(val).toUpperCase();
    if (/^N\/?A$/.test(s)) return { lv: "none", txt: "N/A" };
    if (/ASSENTE|MAI RICEVUT|DA INTEGRARE/.test(s)) return req ? { lv: "bad", txt: s.toLowerCase() } : { lv: "none", txt: s.toLowerCase() };
    return { lv: "ok", txt: String(val) };
  }
  function contrattiDi(o) { return String(o.contratto || "").split(/[\/,;\s]+/).map(x => x.replace(/\D/g, "")).filter(x => x.length >= 6); }
  function verificaPersona(o, ruolo, D, contrattiCantiere) {
    const items = []; const issues = [];
    if (!o) return { trovato: false, lv: "bad", items, issues: [{ voce: "A.2", text: "non presente nel file verifiche ITP (autorizzazione a manovra non verificata)" }] };
    const rid = String(o.rilascio || "").toUpperCase();
    if (rid !== "OK") issues.push({ voce: "A.2", text: `rilascio aut. a manovra ${rid ? rid.toLowerCase() : "non indicato"}` });
    let aut;
    if (o.scadAut && o.scadAut.d) {
      const n = diffDays(o.scadAut.d, D);
      aut = n < 0 ? { lv: "bad", txt: "scad. " + fmtD(o.scadAut.d) } : n <= 30 ? { lv: "warn", txt: "scade " + fmtD(o.scadAut.d) } : { lv: "ok", txt: fmtD(o.scadAut.d) };
      if (n < 0) issues.push({ voce: "A.2", text: `autorizzazione a manovra scaduta il ${fmtD(o.scadAut.d)}` });
    } else { aut = { lv: "bad", txt: o.scadAut ? String(o.scadAut).toLowerCase() : "assente" }; issues.push({ voce: "A.2", text: `autorizzazione a manovra ${aut.txt}` }); }
    if (rid !== "OK" && aut.lv !== "bad") aut = { lv: "bad", txt: aut.txt + " · " + (rid || "n.d.").toLowerCase() };
    const cc = (contrattiCantiere || []).map(x => String(x).replace(/\D/g, "")).filter(x => x.length >= 6);
    const oc = contrattiDi(o);
    if (cc.length && !oc.some(x => cc.includes(x))) {
      aut = { lv: "bad", txt: oc.length ? "≠ contratto cantiere" : "senza contratto" };
      issues.push({ voce: "A.2", text: `autorizzazione a manovra non correlata al contratto del cantiere (${cc.join(" o ")}): risulta ${oc.join("/") || "nessun contratto"}` });
    }
    items.push({ k: "aut", label: "Aut. a manovra" + (oc.length ? " " + oc.join("/") : ""), ...aut });
    const add = (k, opts, voce = "oss") => {
      const st = stato(o[k], D, opts); items.push({ k, label: DOC_LABELS[k], ...st });
      if (st.lv === "bad") issues.push({ voce, text: `${DOC_LABELS[k].toLowerCase()} ${st.scad ? "scaduta il " + fmtD(st.scad) : st.txt}` });
    };
    add("gen", { req: true });
    if (ruolo === "preposto") add("preposto", { req: true });
    add("segn", {});
    if (ruolo === "emergenze") { add("aai", { req: true }); add("ps", { req: true }); }
    for (const k of ["ple", "bybridge", "funi", "quota", "dpi3", "spazi"]) {
      const st = stato(o[k], D, {}); if (st.lv === "bad" || st.lv === "warn") { items.push({ k, label: DOC_LABELS[k], ...st }); if (st.lv === "bad") issues.push({ voce: "oss", text: `${DOC_LABELS[k].toLowerCase()} scaduta il ${fmtD(st.scad)}` }); }
    }
    // UNILAV e fine rapporto
    const un = String(o.unilav || "").toUpperCase();
    let fr = { lv: "ok", txt: o.fineRapp && o.fineRapp.d ? "fino al " + fmtD(o.fineRapp.d) : (o.fineRapp ? String(o.fineRapp).toLowerCase() : "") };
    const frDate = o.fineRapp && o.fineRapp.d ? o.fineRapp.d : (o.fineRapp ? dateInText(o.fineRapp) : null);
    if (frDate && diffDays(frDate, D) < 0) { fr = { lv: "bad", txt: "terminato " + fmtD(frDate) }; issues.push({ voce: "oss", text: `rapporto di lavoro/distacco terminato il ${fmtD(frDate)}` }); }
    else if (frDate && diffDays(frDate, D) <= 30) fr = { lv: "warn", txt: "termina " + fmtD(frDate) };
    if (!un || /ASSENTE/.test(un)) { fr = { lv: "bad", txt: "UNILAV assente" }; issues.push({ voce: "oss", text: "UNILAV/contratto assente" }); }
    else if (/SCADUT/.test(un)) { fr = { lv: "bad", txt: un.toLowerCase() }; issues.push({ voce: "oss", text: un.toLowerCase() }); }
    items.push({ k: "rapporto", label: "UNILAV", ...fr });
    add("idSan", { req: true });
    const lv = items.reduce((a, it) => (["ok", "none", "warn", "bad"].indexOf(it.lv) > ["ok", "none", "warn", "bad"].indexOf(a) ? it.lv : a), "ok");
    return { trovato: true, lv, items, issues, contratti: oc, baw: o.baw, impresa: o.impSub && !/^N\/?A$/i.test(o.impSub) ? o.impSub : o.impApp };
  }
  function persone(v) { // elenco persone del verbale con ruolo e impresa
    const out = [];
    v.imprese.forEach((imp, i) => {
      if (imp.PREPOSTO) out.push({ imp: i, impresa: imp.NOME, ruolo: "preposto", nome: imp.PREPOSTO });
      if (imp.EMERG && imp.EMERG.trim() !== "\"") imp.EMERG.split(/[,;]/).map(x => x.trim()).filter(x => x && !/^"+$/.test(x)).forEach(n => out.push({ imp: i, impresa: imp.NOME, ruolo: "emergenze", nome: n }));
      for (let k = 1; k <= 7; k++) if (imp["OP" + k]) out.push({ imp: i, impresa: imp.NOME, ruolo: "operatore", nome: imp["OP" + k] });
    });
    return out;
  }
  function rilievi(v, operators) {
    const D = v.data || today();
    const cc = String(v.contratti || "").split(/[\/,;\s]+/).filter(Boolean);
    const res = [];
    persone(v).forEach(p => {
      const o = operators ? trovaOperatore(p.nome, operators) : null;
      const ver = operators ? verificaPersona(o, p.ruolo, D, cc) : null;
      if (!ver) return;
      ver.issues.forEach(is => {
        const who = `${(o ? o.nome : p.nome).toUpperCase()}${p.impresa ? " (" + p.impresa + ")" : ""}`;
        const text = `${who}: ${is.text}.`;
        res.push({ k: is.voce + "|" + text, voce: is.voce, text, imp: p.imp });
      });
    });
    return res;
  }

  /* ---------- programma settimanale ---------- */
  const CORSIE_PROG = ["E", "ML", "MV", "S", "DEV", "DEV+B", "FLESSO"];
  function fmtKm(v) {
    if (v === null || v === undefined || v === "") return "";
    const n = typeof v === "number" ? v : parseFloat(String(v).replace(",", "."));
    if (isNaN(n)) return String(v);
    const km = Math.floor(n + 1e-9), m = Math.round((n - km) * 1000);
    return `${km}+${String(m).padStart(3, "0")}`;
  }
  function parseProgramma(XLSX, buf) {
    const wb = XLSX.read(buf, { type: "array" });
    const cell = v => (v === null || v === undefined ? "" : String(v).trim().replace(/^(NOME|COGNOME|TELEFONO|MEZZI)\b[\s\S]*/i, ""));
    const cantieri = []; let settimana = null;
    wb.SheetNames.forEach(sn => {
      const ws = wb.Sheets[sn]; if (!ws["!ref"]) return;
      if (!settimana && ws.G2 && typeof ws.G2.v === "number") settimana = serialToISO(ws.G2.v);
      const rng = XLSX.utils.decode_range(ws["!ref"]);
      for (let R = 4; R <= rng.e.r; R++) {
        const g = C => { const c = ws[XLSX.utils.encode_cell({ r: R, c: C })]; return c ? c.v : null; };
        const hw = cell(g(1)); if (!hw || (g(2) == null && typeof g(14) !== "number")) continue;
        const split = (cols, withCompany, impDefault) => {
          const parts = cols.map(C => cell(g(C)).split(/\r?\n/).map(x => x.trim()));
          const n = Math.max(0, ...parts.map(a => a.filter(Boolean).length)); const out = [];
          for (let i = 0; i < n; i++) {
            const e = withCompany ? { impresa: parts[0][i] || parts[0][0] || "", nome: parts[1][i] || "", cognome: parts[2][i] || "" } : { impresa: impDefault, nome: parts[0][i] || "", cognome: parts[1][i] || "" };
            if ((e.nome + e.cognome).trim()) out.push(e);
          }
          return out;
        };
        const impLavori = cell(g(22)), impSegn = cell(g(24));
        const prep = [...split([28, 29], false, impSegn), ...split([32, 33], false, impLavori), ...split([36, 37, 38], true), ...split([41, 42, 43], true), ...split([46, 47, 48], true)];
        cantieri.push({
          id: `${sn}-${R + 1}`, foglio: sn, autostrada: hw, kmIn: fmtKm(g(2)), kmFin: fmtKm(g(3)), carr: cell(g(4)),
          corsie: CORSIE_PROG.filter((_, k) => /x/i.test(cell(g(5 + k)))), schema: cell(g(12)),
          dal: typeof g(14) === "number" ? serialToISO(g(14)) : "", oreDal: typeof g(15) === "number" ? serialToHM(g(15)) : cell(g(15)),
          al: typeof g(16) === "number" ? serialToISO(g(16)) : "", oreAl: typeof g(17) === "number" ? serialToHM(g(17)) : cell(g(17)),
          lavori: cell(g(21)).replace(/\s*\n\s*/g, " – "), impLavori, impSegn, note: cell(g(27)), prep
        });
      }
    });
    cantieri.sort((a, b) => (a.dal + a.oreDal).localeCompare(b.dal + b.oreDal));
    return { settimana, cantieri };
  }
  function nuovoVerbale(c, esecutore) {
    const imprese = [0, 1, 2, 3].map(() => Object.fromEntries(RIGHE_IMPRESA.map(([k]) => [k, ""])));
    const v = { data: today(), turno: "NOTTURNO", oraInizio: "", oraFine: "", tipologia: "", autostrada: "", kmInizio: "", kmFine: "",
      carreggiata: "", carrAltro: "", corsie: [], contratti: "", imprese, risposte: {}, osservazioni: "", azioni: "", esecutore: esecutore || "", titoloFoto: "", cantiereId: null };
    if (c) {
      Object.assign(v, { cantiereId: c.id, data: c.dal || v.data, oraInizio: c.oreDal || "", oraFine: c.oreAl || "", tipologia: (c.lavori || "").toUpperCase(),
        kmInizio: c.kmIn, kmFine: c.kmFin });
      v.autostrada = AUTOSTRADE.includes(String(c.autostrada).toUpperCase()) ? String(c.autostrada).toUpperCase() : "EXT";
      const cr = String(c.carr || "").toUpperCase();
      if (CARREGGIATE.includes(cr)) v.carreggiata = cr; else v.carrAltro = cr;
      v.corsie = [...new Set(c.corsie.map(x => (x.startsWith("DEV") ? "DEVIAZIONE" : x)).filter(x => CORSIE.includes(x)))];
      const h = parseInt((c.oreDal || "12").slice(0, 2), 10); v.turno = h >= 19 || h < 6 ? "NOTTURNO" : "DIURNO";
      const names = []; const add = n => { if (n && !names.some(x => normImp(x) === normImp(n))) names.push(n); };
      add(c.impSegn); add(c.impLavori); c.prep.forEach(e => add(e.impresa));
      names.slice(0, 4).forEach((n, i) => {
        v.imprese[i].NOME = n;
        const p = c.prep.find(e => normImp(e.impresa) === normImp(n));
        if (p) v.imprese[i].PREPOSTO = `${p.cognome} ${p.nome}`.trim().toUpperCase();
      });
    }
    return v;
  }

  /* ---------- titoli e nomi file ---------- */
  const dataUnd = s => (s || today()).replace(/-/g, "_");
  function titoloDefault(v) {
    const dir = v.carreggiata || v.carrAltro;
    return `${dataUnd(v.data)} - ${v.turno || ""} – ${(v.tipologia || "SOPRALLUOGO").slice(0, 60)} – ${v.autostrada || ""} ${dir || ""}`.replace(/\s+/g, " ").trim();
  }
  function nomeFile(v) {
    const dir = v.carreggiata ? `dir ${v.carreggiata}` : (v.carrAltro || "");
    return `${dataUnd(v.data)} - ${(v.tipologia || "SOPRALLUOGO").slice(0, 50)} - ${v.autostrada || ""} ${dir}`.replace(/[\\/:*?"<>|]+/g, "-").replace(/\s+/g, " ").trim();
  }

  /* ---------- campi per il Word ---------- */
  function campiDocx(v) {
    const f = {
      DATA: v.data ? fmtD(v.data) : "", ORA_INIZIO: v.oraInizio || "", ORA_FINE: v.oraFine || "", TIPOLOGIA: v.tipologia || "",
      KM_INIZIO: v.kmInizio || "", KM_FINE: v.kmFine || "", CARR_ALTRO: v.carrAltro || "", OSSERVAZIONI: v.osservazioni || "",
      AZIONI: v.azioni || "", ESECUTORE: (v.esecutore || "").toUpperCase(), TITOLO_FOTO: v.titoloFoto || titoloDefault(v)
    };
    v.imprese.forEach((imp, i) => RIGHE_IMPRESA.forEach(([k]) => { f[`I${i + 1}_${k}`] = imp[k] || ""; }));
    const marks = new Set();
    if (v.autostrada) marks.add("AUT|" + v.autostrada);
    if (v.carreggiata) marks.add("CARR|" + v.carreggiata);
    v.corsie.forEach(c => marks.add("CORS|" + c));
    Object.entries(v.risposte).forEach(([c, r]) => { if (r) marks.add(`${c}|${r}`); });
    return { fields: f, marks };
  }

  /* ---------- xlsx ---------- */
  const ESITO = { C: "C", NC: "NC", NA: "N/A" };
  function xlsxVerbale(XLSX, v, operators) {
    const D = v.data || today();
    const cc = String(v.contratti || "").split(/[\/,;\s]+/).filter(Boolean);
    const a = [["VERBALE DI SOPRALLUOGO"], [],
      ["Data", fmtD(v.data)], ["Ora inizio", v.oraInizio], ["Ora fine", v.oraFine], ["Turno", v.turno], ["Tipologia attività", v.tipologia],
      ["Autostrada", v.autostrada], ["Km inizio cantiere", v.kmInizio], ["Km fine cantiere", v.kmFine],
      ["Carreggiata", [v.carreggiata, v.carrAltro].filter(Boolean).join(" – ")], ["Corsie escluse", v.corsie.join(", ")],
      ["Contratti cantiere", v.contratti], ["Esecutore sopralluogo", v.esecutore], [],
      ["ANAGRAFICA IMPRESE", "Impresa 1", "Impresa 2", "Impresa 3", "Impresa 4"],
      ...RIGHE_IMPRESA.map(([k, l]) => [l, ...v.imprese.map(i => i[k] || "")]), [],
      ["Voce", "Descrizione", "Esito"], ...ITEMS.map(([c, t]) => [c, t, ESITO[v.risposte[c]] || ""]), [],
      ["Comportamenti osservati / altre osservazioni e note"], [v.osservazioni || ""], [], ["Azioni intraprese"], [v.azioni || ""]];
    const ws1 = XLSX.utils.aoa_to_sheet(a);
    ws1["!cols"] = [{ wch: 34 }, { wch: 70 }, { wch: 22 }, { wch: 22 }, { wch: 22 }];
    const b = [["Impresa", "Ruolo", "Nominativo", "Nel file ITP", "Riga file", "Cod. BAW", "N° contratto", "Rilascio aut.", "Scadenza aut.", "Form. generale", "Preposto", "Segnaletica", "Antincendio", "Primo soccorso", "UNILAV", "Fine rapporto", "Idoneità sanitaria", "Esito", "Rilievi"]];
    persone(v).forEach(p => {
      const o = operators ? trovaOperatore(p.nome, operators) : null;
      const ver = operators ? verificaPersona(o, p.ruolo, D, cc) : null;
      const dv = x => (x && x.d ? fmtD(x.d) : (x || ""));
      b.push([p.impresa, p.ruolo, o ? o.nome : p.nome, o ? "SI" : (operators ? "NO" : "file non caricato"), o ? o.riga : "", o ? (o.baw || "") : "", o ? (o.contratto || "") : "",
        o ? (o.rilascio || "") : "", o ? dv(o.scadAut) : "", o ? dv(o.gen) : "", o ? dv(o.preposto) : "", o ? dv(o.segn) : "", o ? dv(o.aai) : "", o ? dv(o.ps) : "",
        o ? (o.unilav || "") : "", o ? dv(o.fineRapp) : "", o ? dv(o.idSan) : "",
        ver ? { ok: "CONFORME", none: "CONFORME", warn: "IN SCADENZA", bad: "NON CONFORME" }[ver.lv] : "", ver ? ver.issues.map(x => x.text).join("; ") : ""]);
    });
    const ws2 = XLSX.utils.aoa_to_sheet(b);
    ws2["!cols"] = b[0].map((h, i) => ({ wch: i === 18 ? 60 : i === 2 ? 28 : i === 0 ? 24 : 14 }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws1, "Verbale");
    XLSX.utils.book_append_sheet(wb, ws2, "Verifica personale");
    return XLSX.write(wb, { type: "array", bookType: "xlsx" });
  }
  function xlsxRegistro(XLSX, verbali, operators) {
    const h = ["Data", "Turno", "Ora inizio", "Ora fine", "Tipologia attività", "Autostrada", "Km inizio", "Km fine", "Carreggiata", "Corsie escluse", "Imprese", "Preposti", "N° NC", "Voci NC", "Rilievi personale", "Esecutore", "File", "Cartella Drive"];
    const rows = verbali.filter(x => x.stato === "completato").sort((a, b) => (a.dati.data + a.dati.oraInizio).localeCompare(b.dati.data + b.dati.oraInizio)).map(x => {
      const v = x.dati; const nc = Object.entries(v.risposte).filter(([, r]) => r === "NC").map(([c]) => c);
      return [fmtD(v.data), v.turno, v.oraInizio, v.oraFine, v.tipologia, v.autostrada, v.kmInizio, v.kmFine, [v.carreggiata, v.carrAltro].filter(Boolean).join(" – "), v.corsie.join(", "),
        v.imprese.map(i => i.NOME).filter(Boolean).join(" / "), v.imprese.map(i => i.PREPOSTO).filter(Boolean).join(" / "), nc.length, nc.join(", "),
        operators ? rilievi(v, operators).length : "", v.esecutore, nomeFile(v), x.sync && x.sync.folderUrl ? x.sync.folderUrl : ""];
    });
    const ws = XLSX.utils.aoa_to_sheet([h, ...rows]);
    ws["!cols"] = h.map((_, i) => ({ wch: [4, 10, 11].includes(i) ? 40 : i === 17 ? 50 : 14 }));
    const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, "Registro sopralluoghi");
    return XLSX.write(wb, { type: "array", bookType: "xlsx" });
  }

  return { CHECK, ITEMS, AUTOSTRADE, CARREGGIATE, CORSIE, RIGHE_IMPRESA, today, fmtD, diffDays, norm, normImp,
    parseITP, trovaOperatore, verificaPersona, persone, rilievi, parseProgramma, nuovoVerbale, titoloDefault, nomeFile, campiDocx,
    xlsxVerbale, xlsxRegistro, fmtKm };
});
