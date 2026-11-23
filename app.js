/* Sopralluoghi – app offline per compilare la check-list ASPI e salvarla su Google Drive */
"use strict";
const VERSION = "1.1.1";
const CARTELLA_LAVORO = "FILE DI LAVORO";
const GIORNI_AVVISO = 7;
const CARTELLA_DRIVE = "VERBALE SOPRALLUOGHI ASPI";
const CLIENT_ID_PREDEFINITO = "648194852883-b7rlulsqt3vbmg37r6p1bpbhaa744e2m.apps.googleusercontent.com";
const XMIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const DMIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const esc = s => String(s ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const uid = p => p + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
const GG = ["dom", "lun", "mar", "mer", "gio", "ven", "sab"];
const fmtG = s => { if (!s) return ""; const [y, m, d] = s.split("-").map(Number); const dt = new Date(y, m - 1, d); return `${GG[dt.getDay()]} ${d}/${m}`; };
const fmtTs = s => { if (!s) return ""; const d = new Date(s); return `${String(d.getDate()).padStart(2, "0")}/${String(d.getMonth() + 1).padStart(2, "0")} ${String(d.getHours()).padStart(2, "0")}:${String(d.getMinutes()).padStart(2, "0")}`; };

/* ---------- archivio locale (IndexedDB) ---------- */
const DB = (() => {
  let p;
  const open = () => p || (p = new Promise((res, rej) => {
    const r = indexedDB.open("sopralluoghi", 1);
    r.onupgradeneeded = () => { const d = r.result; ["kv", "foto"].forEach(n => { if (!d.objectStoreNames.contains(n)) d.createObjectStore(n); }); if (!d.objectStoreNames.contains("verbali")) d.createObjectStore("verbali", { keyPath: "id" }); };
    r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error);
  }));
  const run = async (store, mode, fn) => { const d = await open(); return new Promise((res, rej) => { const t = d.transaction(store, mode); const r = fn(t.objectStore(store)); t.oncomplete = () => res(r ? r.result : undefined); t.onerror = t.onabort = () => rej(t.error); }); };
  return {
    get: k => run("kv", "readonly", s => s.get(k)), set: (k, v) => run("kv", "readwrite", s => s.put(v, k)),
    all: () => run("verbali", "readonly", s => s.getAll()), put: v => run("verbali", "readwrite", s => s.put(v)), remove: id => run("verbali", "readwrite", s => s.delete(id)),
    getFoto: id => run("foto", "readonly", s => s.get(id)), putFoto: (id, b) => run("foto", "readwrite", s => s.put(b, id)), delFoto: id => run("foto", "readwrite", s => s.delete(id))
  };
})();

/* ---------- stato ---------- */
const S = { settings: {}, tpl: null, itp: null, prog: null, verbali: [], curId: null, view: "elenco",
  ui: { newSheet: false, delAsk: null, errors: [], q: "", busy: "" } };
let installPrompt = null;
const cur = () => S.verbali.find(x => x.id === S.curId) || null;
const pending = () => S.verbali.filter(x => x.stato === "completato" && (!x.sync || x.sync.stato !== "caricato"));
async function saveSettings() { await DB.set("settings", S.settings); }
const saveT = {};
function saveV(x, now) {
  x.updated = new Date().toISOString();
  if (x.sync && x.sync.stato === "caricato") x.sync.stato = "modificato";
  clearTimeout(saveT[x.id]);
  const go = () => DB.put(x).catch(e => toast("Salvataggio non riuscito: " + e.message));
  if (now) return go(); saveT[x.id] = setTimeout(go, 350);
}

/* ---------- utilità interfaccia ---------- */
function toast(msg) { const t = document.getElementById("toast"); t.textContent = msg; t.hidden = false; clearTimeout(toast.t); toast.t = setTimeout(() => (t.hidden = true), 3800); }
function getPath(o, p) { return p.split(".").reduce((a, k) => (a == null ? a : a[k]), o); }
function setPath(o, p, v) { const ks = p.split("."); const last = ks.pop(); const t = ks.reduce((a, k) => a[k], o); t[last] = v; }
function download(blob, name) { const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = name; document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 5000); }
const lvTxt = { ok: "Conforme", none: "Conforme", warn: "In scadenza", bad: "Non conforme" };

/* ---------- documenti ---------- */
async function generaDocx(x) {
  const t = await DB.get("template");
  if (!t) throw new Error("Carica il modello Word in Impostazioni.");
  const { fields, marks } = Core.campiDocx(x.dati);
  const photos = [];
  for (const [i, f] of x.foto.entries()) {
    const b = await DB.getFoto(f.id); if (!b) continue;
    photos.push({ bytes: new Uint8Array(await b.arrayBuffer()), w: f.w, h: f.h, caption: `Foto ${i + 1}${f.voce ? " – " + f.voce : ""}${f.nota ? ": " + f.nota : ""}` });
  }
  return DocxFill.fill(JSZip, t.bytes, fields, marks, photos, "blob");
}
const generaXlsx = x => new Blob([Core.xlsxVerbale(XLSX, x.dati, S.itp && S.itp.operators)], { type: XMIME });
async function compress(file) {
  let bmp;
  try { bmp = await createImageBitmap(file, { imageOrientation: "from-image" }); } catch (e) { bmp = await createImageBitmap(file); }
  const k = Math.min(1, 1600 / Math.max(bmp.width, bmp.height));
  const w = Math.round(bmp.width * k), h = Math.round(bmp.height * k);
  const c = document.createElement("canvas"); c.width = w; c.height = h; c.getContext("2d").drawImage(bmp, 0, 0, w, h);
  const blob = await new Promise(r => c.toBlob(r, "image/jpeg", 0.82));
  return { blob, w, h };
}

/* ---------- cartelle e file di lavoro su Drive ---------- */
async function cartelle() {
  if (!(await Drive.exists(S.settings.rootId))) { const r = await Drive.folder(CARTELLA_DRIVE, null, true); S.settings.rootId = r.id; S.settings.rootUrl = r.webViewLink; }
  if (!(await Drive.exists(S.settings.lavoroId))) { const l = await Drive.folder(CARTELLA_LAVORO, S.settings.rootId); S.settings.lavoroId = l.id; S.settings.lavoroUrl = l.webViewLink; }
  await saveSettings();
}
function tipoFile(nome) {
  const n = nome.toUpperCase();
  if (/ITP/.test(n)) return "loadItp";
  if (/PROGRAMMA/.test(n)) return "loadProg";
  if (/BAW/.test(n)) return "baw";
  return null;
}
async function aggiornaDaDrive() {
  const files = (await Drive.list(S.settings.lavoroId)).filter(f => /\.xls[xm]?$/i.test(f.name) || f.mimeType === "application/vnd.google-apps.spreadsheet");
  const best = {};
  files.forEach(f => { const t = tipoFile(f.name); if (t && (!best[t] || f.modifiedTime > best[t].modifiedTime)) best[t] = f; });
  const fatto = [], note = [];
  for (const [t, f] of Object.entries(best)) {
    if (t === "baw") { note.push("scarico BAW trovato (integrazione in arrivo)"); continue; }
    const cur = t === "loadItp" ? S.itp : S.prog;
    if (cur && cur.src && cur.src.id === f.id && cur.src.mod === f.modifiedTime) continue;
    try { fatto.push(await loadBuf(t, f.name, await Drive.download(f), { origine: "drive", id: f.id, mod: f.modifiedTime }, true)); }
    catch (e) { note.push(`${f.name}: ${e.message}`); }
  }
  return { fatto, note, vuota: !Object.keys(best).length };
}
function etaFile(x) { if (!x) return null; const d = new Date((x.src && x.src.mod) || x.at); return Math.floor((Date.now() - d) / 86400000); }

/* ---------- sincronizzazione Drive ---------- */
async function sincronizza() {
  if (S.ui.busy) return;
  if (!navigator.onLine) { toast("Sei offline: i sopralluoghi restano sul telefono, li carichi quando torna la rete."); return; }
  if (!S.settings.clientId) { toast("Prima collega Google Drive in Impostazioni."); go("impostazioni"); return; }
  const list = pending();
  S.ui.busy = "Accesso a Google…"; renderBar();
  try {
    await Drive.auth(S.settings.clientId);
    await cartelle();
    S.ui.busy = "Controllo i file di lavoro…"; renderBar();
    let agg = { fatto: [], note: [] };
    try { agg = await aggiornaDaDrive(); } catch (e) { agg.note.push("file di lavoro non letti: " + e.message); }
    let n = 0, err = 0;
    for (const x of list) {
      S.ui.busy = `Carico ${++n}/${list.length}…`; renderBar();
      try { await syncOne(x, S.settings.rootId); }
      catch (e) { x.sync = { ...(x.sync || {}), stato: "errore", err: e.message }; await DB.put(x); err++; if (e.status === 401) throw e; }
    }
    S.ui.busy = "Aggiorno il registro…"; renderBar();
    const reg = new Blob([Core.xlsxRegistro(XLSX, S.verbali, S.itp && S.itp.operators)], { type: XMIME });
    const f = await Drive.upload("Registro_sopralluoghi.xlsx", XMIME, reg, S.settings.rootId, S.settings.registroId || null);
    S.settings.registroId = f.id; await saveSettings();
    const parti = [err ? `${n - err} caricati, ${err} con errore` : (n ? `${n} sopralluoghi caricati` : "sopralluoghi già su Drive")];
    if (agg.fatto.length) parti.push("aggiornati: " + agg.fatto.join(", "));
    if (agg.note.length) parti.push(agg.note.join("; "));
    toast(parti.join(" · ") + ".");
  } catch (e) { toast(e.message); }
  finally { S.ui.busy = ""; render(); }
}
async function syncOne(x, rootId) {
  const v = x.dati; const s = x.sync || {}; const base = Core.nomeFile(v);
  let folderId = s.folderId, folderUrl = s.folderUrl;
  if (!(await Drive.exists(folderId))) {
    const month = await Drive.folder((v.data || Core.today()).slice(0, 7), rootId);
    const fo = await Drive.folder(base, month.id); folderId = fo.id; folderUrl = fo.webViewLink;
  }
  const fd = await Drive.upload(base + ".docx", DMIME, await generaDocx(x), folderId, s.files && s.files.docx);
  const fx = await Drive.upload(base + ".xlsx", XMIME, generaXlsx(x), folderId, s.files && s.files.xlsx);
  const foto = { ...((s.files && s.files.foto) || {}) };
  for (const [i, f] of x.foto.entries()) {
    if (foto[f.id]) continue;
    const b = await DB.getFoto(f.id); if (!b) continue;
    const r = await Drive.upload(`foto_${String(i + 1).padStart(2, "0")}${f.voce ? "_" + f.voce : ""}.jpg`, "image/jpeg", b, folderId);
    foto[f.id] = r.id;
  }
  x.sync = { stato: "caricato", at: new Date().toISOString(), folderId, folderUrl, docxUrl: fd.webViewLink, files: { docx: fd.id, xlsx: fx.id, foto } };
  await DB.put(x);
}

/* ---------- viste ---------- */
const VIEWS = [["elenco", "≡", "Sopralluoghi"], ["compila", "✎", "Compila"], ["personale", "◉", "Personale"], ["impostazioni", "⚙", "Impostazioni"]];
function go(v) { S.view = v; if (location.hash !== "#" + v) history.replaceState(null, "", "#" + v); render(); window.scrollTo(0, 0); }
function renderBar() {
  const n = pending().length;
  document.getElementById("netState").textContent = (navigator.onLine ? "online" : "offline") + (n ? ` · ${n} da caricare` : "");
  const b = document.getElementById("syncBtn"); b.textContent = S.ui.busy || (n ? `Sincronizza (${n})` : "Sincronizza"); b.disabled = !!S.ui.busy;
}
function render() {
  renderBar();
  document.getElementById("tabs").innerHTML = VIEWS.map(([k, ic, l]) => `<button class="tab" type="button" data-act="go" data-v="${k}" ${S.view === k ? 'aria-current="page"' : ""}><b aria-hidden="true">${ic}</b>${l}</button>`).join("");
  const ae = document.activeElement && document.activeElement.id;
  document.getElementById("app").innerHTML = ({ elenco: vElenco, compila: vCompila, personale: vPersonale, impostazioni: vImpostazioni }[S.view] || vElenco)();
  if (ae) { const el = document.getElementById(ae); if (el && el.focus) el.focus({ preventScroll: true }); }
  hydrateThumbs();
  if (S.view === "impostazioni" && navigator.storage && navigator.storage.estimate) navigator.storage.estimate().then(m => {
    const el = document.getElementById("storeInfo"); if (el) el.textContent = `Memoria usata dall'app: ${(m.usage / 1048576).toFixed(1)} MB (disponibili ${(m.quota / 1073741824).toFixed(1)} GB).`;
  }).catch(() => {});
}
const URLS = new Map();
async function hydrateThumbs() {
  for (const img of document.querySelectorAll("img[data-foto]")) {
    const id = img.dataset.foto;
    if (!URLS.has(id)) { const b = await DB.getFoto(id); if (b) URLS.set(id, URL.createObjectURL(b)); }
    if (URLS.has(id)) img.src = URLS.get(id);
  }
}

/* Elenco */
function srcTxt(x) { const g = etaFile(x); const da = x.src && x.src.origine === "drive" ? "da Drive" : "dal telefono"; return `${da}, ${g === 0 ? "oggi" : g === 1 ? "ieri" : g + " giorni fa"}`; }
function avvisoFile() {
  const w = [];
  if (S.itp && etaFile(S.itp) > GIORNI_AVVISO) w.push(`il file ITP è di ${etaFile(S.itp)} giorni fa`);
  if (S.prog && S.prog.settimana && Core.diffDays(Core.today(), S.prog.settimana) > 6) w.push("il programma cantieri è della settimana scorsa");
  return w.length ? `<section class="card lv-warn"><p><b>Attenzione:</b> ${w.join(" e ")}. Premi Sincronizza con la rete per prendere le versioni nuove da Drive.</p></section>` : "";
}
function vElenco() {
  const miss = [];
  if (!S.tpl) miss.push("Carica il <b>modello Word ASPI</b> preparato per l'app");
  if (!S.itp) miss.push("Carica il <b>file verifiche ITP</b> (operatori e contratti)");
  if (!S.settings.clientId) miss.push("Collega <b>Google Drive</b> (Client ID)");
  const ns = S.ui.newSheet;
  return `
  ${miss.length ? `<section class="card lv-warn"><h3>Per iniziare</h3><ul class="setup">${miss.map(m => `<li>${m}</li>`).join("")}</ul><button class="btn" type="button" data-act="go" data-v="impostazioni">Apri Impostazioni</button></section>` : ""}
  ${installPrompt ? `<button class="btn wide" type="button" data-act="install">Installa l'app sul telefono</button>` : ""}
  ${avvisoFile()}
  <div class="row"><button class="btn pri wide" type="button" data-act="newOpen">+ Nuovo sopralluogo</button></div>
  ${ns ? vNuovo() : ""}
  <h2>Sopralluoghi (${S.verbali.length})</h2>
  <div class="list">${S.verbali.map(cardVerbale).join("") || `<p class="muted">Nessun sopralluogo. Creane uno con il pulsante qui sopra.</p>`}</div>`;
}
function vNuovo() {
  const cs = (S.prog && S.prog.cantieri) || [];
  return `<section class="card"><div class="row between"><h3>Da quale cantiere parti?</h3><button class="btn sm" type="button" data-act="newClose">Annulla</button></div>
    <button class="btn" type="button" data-act="newFrom" data-id="">Sopralluogo senza programma</button>
    ${cs.length ? `<p class="muted">Programma settimana dal ${Core.fmtD(S.prog.settimana)} (${esc(S.prog.name)})</p>
    <div class="sheet">${cs.map(c => `<button class="cant" type="button" data-act="newFrom" data-id="${esc(c.id)}"><b>${esc(c.autostrada)} ${esc(c.carr)} · km ${esc(c.kmIn)} → ${esc(c.kmFin)}</b><span class="muted">${esc(fmtG(c.dal))} ${esc(c.oreDal)}–${esc(c.oreAl)} · ${esc((c.lavori || "").slice(0, 70))}</span><span class="muted">${esc([c.impSegn, c.impLavori].filter(Boolean).join(" · "))}</span></button>`).join("")}</div>`
      : `<p class="muted">Nessun programma caricato: puoi importarlo in Impostazioni per precompilare dati e imprese.</p>`}</section>`;
}
function cardVerbale(x) {
  const v = x.dati; const nc = Object.values(v.risposte).filter(r => r === "NC").length;
  const ss = x.stato !== "completato" ? `<span class="pill none">Bozza</span>` :
    !x.sync || x.sync.stato === "da_caricare" ? `<span class="pill warn">Da caricare</span>` :
      x.sync.stato === "caricato" ? `<span class="pill ok">Su Drive ${esc(fmtTs(x.sync.at))}</span>` :
        x.sync.stato === "modificato" ? `<span class="pill warn">Modificato, da ricaricare</span>` : `<span class="pill bad">Errore caricamento</span>`;
  const del = S.ui.delAsk === x.id;
  return `<article class="card vrb ${nc ? "lv-bad" : x.stato === "completato" ? "lv-ok" : ""}">
    <div class="t">${esc(v.titoloFoto || Core.titoloDefault(v))}</div>
    <div class="row">${ss}${nc ? `<span class="pill bad">${nc} NC</span>` : ""}<span class="muted">${x.foto.length} foto</span></div>
    ${x.sync && x.sync.stato === "errore" ? `<p class="err">${esc(x.sync.err)}</p>` : ""}
    <div class="row">
      <button class="btn sm pri" type="button" data-act="open" data-id="${x.id}">Apri</button>
      <button class="btn sm" type="button" data-act="docx" data-id="${x.id}">Word</button>
      <button class="btn sm" type="button" data-act="xlsx" data-id="${x.id}">Excel</button>
      <button class="btn sm" type="button" data-act="share" data-id="${x.id}">Condividi</button>
      ${x.sync && x.sync.folderUrl ? `<a class="btn sm" href="${esc(x.sync.folderUrl)}" target="_blank" rel="noopener">Cartella Drive</a>` : ""}
      ${del ? `<button class="btn sm danger" type="button" data-act="delDo" data-id="${x.id}">Conferma eliminazione</button><button class="btn sm" type="button" data-act="delNo">No</button>`
        : `<button class="btn sm danger" type="button" data-act="delAsk" data-id="${x.id}">Elimina</button>`}
    </div></article>`;
}

/* Compila */
function contratti(v) { return String(v.contratti || "").split(/[\/,;\s]+/).filter(Boolean); }
function personBlock(v, nome, ruolo) {
  if (!nome || !nome.trim()) return "";
  if (!S.itp) return `<div class="docs"><span class="d none">File ITP non caricato: verifica non disponibile</span></div>`;
  const o = Core.trovaOperatore(nome, S.itp.operators);
  if (!o) return `<div class="docs"><span class="d bad">Non trovato nel file ITP</span></div>`;
  const ver = Core.verificaPersona(o, ruolo, v.data || Core.today(), contratti(v));
  const items = ver.items.filter(it => it.lv === "bad" || it.lv === "warn" || it.k === "aut");
  return `<div class="docs"><span class="pill ${ver.lv === "none" ? "ok" : ver.lv}">${lvTxt[ver.lv]}</span>${Core.norm(o.nome) !== Core.norm(nome) ? `<span class="d none">= ${esc(o.nome)}</span>` : ""}
    ${items.map(it => `<span class="d ${it.lv}">${esc(it.label)}: ${esc(it.txt)}</span>`).join("")}</div>`;
}
function inputP(id, label, path, v, ruolo) {
  const val = getPath(v, path) || "";
  return `<div class="person"><label class="f" for="${id}">${esc(label)}<input id="${id}" list="dl-persone" autocomplete="off" autocapitalize="characters" value="${esc(val)}" data-f="${path}"></label>${ruolo ? personBlock(v, val, ruolo) : ""}</div>`;
}
function vCompila() {
  const x = cur();
  if (!x) return `<section class="card"><h3>Nessun sopralluogo aperto</h3><p class="muted">Crea un nuovo sopralluogo o aprine uno dall'elenco.</p><button class="btn pri" type="button" data-act="newOpen2">+ Nuovo sopralluogo</button></section>`;
  const v = x.dati;
  const done = Core.ITEMS.filter(([c]) => v.risposte[c]).length;
  const ril = S.itp ? Core.rilievi(v, S.itp.operators) : [];
  const applied = v.applicati || [];
  const chip = (f, val, multi) => { const on = multi ? (v[f] || []).includes(val) : v[f] === val; return `<button type="button" class="chip" aria-pressed="${on}" data-act="${multi ? "toggle" : "pick"}" data-f="${f}" data-v="${esc(val)}">${esc(val)}</button>`; };
  return `
  <div class="row between"><h2>${x.stato === "completato" ? "Sopralluogo completato" : "Sopralluogo in corso"}</h2><span class="muted">${done}/${Core.ITEMS.length} voci</span></div>
  <div class="prog" aria-hidden="true"><i style="width:${Math.round(done / Core.ITEMS.length * 100)}%"></i></div>
  ${S.itp ? `<p class="muted">Verifiche sul file ITP ${esc(S.itp.name)} (${srcTxt(S.itp)})</p>` : ""}
  ${avvisoFile()}
  <section class="card"><h3>Dati generali cantiere</h3>
    <div class="grid2">
      <label class="f" for="f-data">Data<input id="f-data" type="date" value="${esc(v.data)}" data-f="data"></label>
      <div class="f"><span class="f" style="font-size:.82rem;color:var(--muted);font-weight:600">Turno</span><div class="chips">${chip("turno", "NOTTURNO")}${chip("turno", "DIURNO")}</div></div>
      <label class="f" for="f-oi">Ora inizio<input id="f-oi" type="time" value="${esc(v.oraInizio)}" data-f="oraInizio"></label>
      <label class="f" for="f-of">Ora fine<input id="f-of" type="time" value="${esc(v.oraFine)}" data-f="oraFine"></label>
    </div>
    <label class="f" for="f-tip">Tipologia attività<input id="f-tip" value="${esc(v.tipologia)}" data-f="tipologia" autocapitalize="characters"></label>
    <div class="f"><span>Autostrada</span><div class="chips">${Core.AUTOSTRADE.map(a => chip("autostrada", a)).join("")}</div></div>
    <div class="grid2">
      <label class="f" for="f-ki">Km inizio cantiere<input id="f-ki" value="${esc(v.kmInizio)}" data-f="kmInizio" placeholder="20+000" inputmode="text"></label>
      <label class="f" for="f-kf">Km fine cantiere<input id="f-kf" value="${esc(v.kmFine)}" data-f="kmFine" placeholder="21+600" inputmode="text"></label>
    </div>
    <div class="f"><span>Carreggiata</span><div class="chips">${Core.CARREGGIATE.map(a => chip("carreggiata", a)).join("")}</div></div>
    <label class="f" for="f-ca">Dettaglio carreggiata (es. TRONCHETTO DIR RM EST)<input id="f-ca" value="${esc(v.carrAltro)}" data-f="carrAltro" autocapitalize="characters"></label>
    <div class="f"><span>Corsie escluse</span><div class="chips">${Core.CORSIE.map(a => chip("corsie", a, true)).join("")}</div></div>
    <label class="f" for="f-con">Contratti del cantiere (C.A./AQ) per verificare le autorizzazioni<input id="f-con" value="${esc(v.contratti)}" data-f="contratti" inputmode="numeric" placeholder="es. 81032064, 81026625"></label>
  </section>
  <section class="card"><h3>Anagrafica imprese e verifica personale</h3>
    ${v.imprese.map((imp, i) => {
      const nOps = Math.min(7, Math.max(1, [1, 2, 3, 4, 5, 6, 7].filter(k => imp["OP" + k]).length + 1));
      const filled = Object.values(imp).some(Boolean);
      return `<details class="imp" ${filled || i === 0 ? "open" : ""}><summary>Impresa ${i + 1}${imp.NOME ? " · " + esc(imp.NOME) : ""}</summary><div class="in">
        <label class="f" for="i${i}-n">Nome impresa<input id="i${i}-n" list="dl-imprese" autocomplete="off" autocapitalize="characters" value="${esc(imp.NOME)}" data-f="imprese.${i}.NOME"></label>
        ${inputP(`i${i}-p`, "Preposto", `imprese.${i}.PREPOSTO`, v, "preposto")}
        ${inputP(`i${i}-e`, "Addetto emergenze", `imprese.${i}.EMERG`, v, "emergenze")}
        ${[1, 2, 3, 4, 5, 6, 7].slice(0, nOps).map(k => inputP(`i${i}-o${k}`, `Operatore ${k}`, `imprese.${i}.OP${k}`, v, "operatore")).join("")}
      </div></details>`;
    }).join("")}
  </section>
  ${ril.length ? `<section class="card lv-warn"><div class="row between"><h3>Rilievi dal file ITP (${ril.length})</h3><button class="btn sm" type="button" data-act="applicaTutti">Applica tutti</button></div>
    <p class="muted">“Applica” mette il rilievo in NC sulla voce indicata (A.2) o nelle osservazioni.</p>
    <div class="sugg">${ril.map(r => { const d = applied.includes(r.k); return `<div class="s ${d ? "done" : ""}"><span><b>${r.voce === "oss" ? "Oss." : esc(r.voce)}</b> ${esc(r.text)}</span>${d ? `<span class="pill none">Applicato</span>` : `<button class="btn sm" type="button" data-act="applica" data-k="${esc(r.k)}">Applica</button>`}</div>`; }).join("")}</div></section>` : ""}
  <div class="row"><button class="btn sm" type="button" data-act="restC">Voci vuote = C</button><button class="btn sm" type="button" data-act="restNA">Voci vuote = N/A</button></div>
  ${Core.CHECK.map(sec => `<section class="sec"><h3>${esc(sec.s)} – ${esc(sec.t)}</h3>${sec.v.map(([c, t]) => {
    const r = v.risposte[c] || "";
    const b = (k, l) => `<button type="button" class="${k}" aria-pressed="${r === k}" data-act="ans" data-c="${c}" data-v="${k}" aria-label="${c} ${l}">${l}</button>`;
    return `<div class="item ${r}"><div class="t"><b>${esc(c)}</b>${esc(t)}</div><div class="seg" role="group" aria-label="Esito ${esc(c)}">${b("C", "C")}${b("NC", "NC")}${b("NA", "N/A")}</div></div>`;
  }).join("")}</section>`).join("")}
  <section class="card"><h3>Foto (${x.foto.length})</h3>
    <div class="row">
      <label class="btn pri file" for="camIn">Scatta foto<input id="camIn" type="file" accept="image/*" capture="environment" data-act="foto"></label>
      <label class="btn file" for="galIn">Dalla galleria<input id="galIn" type="file" accept="image/*" multiple data-act="foto"></label>
    </div>
    <div class="thumbs">${x.foto.map((f, i) => `<div class="th"><img data-foto="${f.id}" alt="Foto ${i + 1}">
      <select aria-label="Voce foto ${i + 1}" data-act="fotoVoce" data-id="${f.id}"><option value="">Foto ${i + 1} · nessuna voce</option>${Core.ITEMS.map(([c]) => `<option ${f.voce === c ? "selected" : ""}>${c}</option>`).join("")}</select>
      <input class="inp" aria-label="Didascalia foto ${i + 1}" placeholder="Didascalia" value="${esc(f.nota)}" data-fnota="${f.id}">
      <button class="btn sm danger" type="button" data-act="fotoDel" data-id="${f.id}">Rimuovi</button></div>`).join("")}</div>
  </section>
  <section class="card"><h3>Osservazioni e azioni</h3>
    <label class="f" for="f-oss">Comportamenti osservati / altre osservazioni e note<textarea id="f-oss" rows="8" data-f="osservazioni">${esc(v.osservazioni)}</textarea></label>
    <label class="f" for="f-az">Azioni intraprese<textarea id="f-az" rows="4" data-f="azioni">${esc(v.azioni)}</textarea></label>
    <label class="f" for="f-es">Esecutore sopralluogo<input id="f-es" value="${esc(v.esecutore)}" data-f="esecutore" autocapitalize="characters"></label>
    <label class="f" for="f-tit">Titolo pagina foto<input id="f-tit" value="${esc(v.titoloFoto)}" data-f="titoloFoto" placeholder="${esc(Core.titoloDefault(v))}"></label>
  </section>
  ${S.ui.errors.length ? `<div class="card lv-bad" role="alert">${S.ui.errors.map(e => `<p class="err">${esc(e)}</p>`).join("")}</div>` : ""}
  <div class="sticky"><button class="btn" type="button" data-act="docx" data-id="${x.id}">Anteprima Word</button><button class="btn pri" style="flex:1" type="button" data-act="concludi">${x.stato === "completato" ? "Salva modifiche" : "Concludi"}</button></div>`;
}

/* Personale */
function vPersonale() {
  if (!S.itp) return `<section class="card"><h3>File verifiche ITP non caricato</h3><p class="muted">Caricalo in Impostazioni per cercare operatori e controllare le scadenze anche offline.</p><button class="btn" type="button" data-act="go" data-v="impostazioni">Apri Impostazioni</button></section>`;
  const q = Core.norm(S.ui.q);
  const res = q.length < 2 ? [] : S.itp.operators.filter(o => Core.norm(o.nome).includes(q) || Core.norm(o.impApp).includes(q) || Core.norm(o.impSub).includes(q) || String(o.baw || "").includes(q) || String(o.contratto || "").includes(q)).slice(0, 40);
  const D = Core.today();
  return `<h2>Personale verificato</h2>
  <p class="muted">${S.itp.operators.length} operatori da ${esc(S.itp.name)} (caricato ${esc(fmtTs(S.itp.at))}). Scadenze calcolate a oggi.</p>
  <label class="f" for="q">Cerca per nome, impresa, codice BAW o contratto<input id="q" type="search" value="${esc(S.ui.q)}" data-act="q" autocomplete="off" autocapitalize="characters"></label>
  <div class="list">${res.map(o => {
    const merged = new Map();
    ["operatore", "preposto", "emergenze"].forEach(r => Core.verificaPersona(o, r, D, []).items.forEach(it => { if (!merged.has(it.k)) merged.set(it.k, it); }));
    const ver = Core.verificaPersona(o, "operatore", D, []);
    return `<article class="card lv-${ver.lv === "none" ? "ok" : ver.lv}"><div class="row between"><b>${esc(o.nome)}</b><span class="pill ${ver.lv === "none" ? "ok" : ver.lv}">${lvTxt[ver.lv]}</span></div>
      <dl class="kv"><dt>Impresa</dt><dd>${esc([o.impApp, o.impSub].filter(x => x && !/^N\/?A$/i.test(x)).join(" → "))}</dd><dt>Contratto</dt><dd>${esc(o.contratto || "—")}</dd><dt>Cod. BAW</dt><dd>${esc(o.baw || "—")}</dd><dt>Rilascio aut.</dt><dd>${esc(o.rilascio || "—")}</dd>${o.note ? `<dt>Note</dt><dd>${esc(o.note)}</dd>` : ""}</dl>
      <div class="docs">${[...merged.values()].map(it => `<span class="d ${it.lv}">${esc(it.label)}: ${esc(it.txt)}</span>`).join("")}</div></article>`;
  }).join("") || (q.length >= 2 ? `<p class="muted">Nessun risultato.</p>` : "")}</div>`;
}

/* Impostazioni */
function vImpostazioni() {
  const st = S.settings;
  return `<h2>Impostazioni</h2>
  <section class="card"><h3>Esecutore</h3>
    <label class="f" for="s-es">Nome sul verbale<input id="s-es" value="${esc(st.esecutore)}" data-s="esecutore" autocapitalize="characters"></label>
  </section>
  <section class="card"><h3>File di lavoro (restano sul telefono)</h3>
    <div class="row between"><div><b>Modello Word ASPI</b><p class="muted">${S.tpl ? `${esc(S.tpl.name)} · ${esc(fmtTs(S.tpl.at))}` : "Non caricato"}</p></div><label class="btn sm file" for="tplIn">${S.tpl ? "Sostituisci" : "Carica"}<input id="tplIn" type="file" accept=".docx" data-act="loadTpl"></label></div>
    <div class="row between"><div><b>Verifiche ITP (operatori e contratti)</b><p class="muted">${S.itp ? `${esc(S.itp.name)} · ${S.itp.operators.length} operatori · ${srcTxt(S.itp)}` : "Non caricato"}</p></div><label class="btn sm file" for="itpIn">${S.itp ? "Aggiorna" : "Carica"}<input id="itpIn" type="file" accept=".xlsx,.xlsm" data-act="loadItp"></label></div>
    <div class="row between"><div><b>Programma cantieri ASPI</b><p class="muted">${S.prog ? `${esc(S.prog.name)} · ${S.prog.cantieri.length} cantieri · settimana dal ${Core.fmtD(S.prog.settimana)} · ${srcTxt(S.prog)}` : "Facoltativo: precompila dati e imprese"}</p></div><label class="btn sm file" for="progIn">${S.prog ? "Aggiorna" : "Carica"}<input id="progIn" type="file" accept=".xlsx,.xlsm" data-act="loadProg"></label></div>
    <div class="row between"><div><b>Scarico BAW</b><p class="muted">In arrivo: si collega appena ricevi il file.</p></div><button class="btn sm" type="button" disabled>Carica</button></div>
    <p class="muted">Aggiornamento automatico: metti i nuovi xlsx (nome con “ITP”, “PROGRAMMA” o “BAW”) nella cartella Drive “${esc(CARTELLA_DRIVE)}/${esc(CARTELLA_LAVORO)}”. A ogni Sincronizza l'app prende la versione più recente.</p>
    <div class="row"><button class="btn" type="button" data-act="pullDrive">Aggiorna ora da Drive</button>${S.settings.lavoroUrl ? `<a class="btn" href="${esc(S.settings.lavoroUrl)}" target="_blank" rel="noopener">Apri ${esc(CARTELLA_LAVORO)}</a>` : ""}</div>
  </section>
  <section class="card"><h3>Google Drive</h3>
    <label class="f" for="s-cid">Client ID OAuth (Google Cloud)<input id="s-cid" value="${esc(st.clientId)}" data-s="clientId" placeholder="xxxxxxxx.apps.googleusercontent.com" autocapitalize="off" spellcheck="false"></label>
    <div class="row"><button class="btn pri" type="button" data-act="connect">Collega e verifica</button>${st.rootUrl ? `<a class="btn" href="${esc(st.rootUrl)}" target="_blank" rel="noopener">Apri cartella</a>` : ""}</div>
    <p class="muted">I sopralluoghi vanno in “${esc(CARTELLA_DRIVE)}/anno-mese/nome verbale” con Word, Excel e foto, più il file Registro_sopralluoghi.xlsx. L'app scrive solo nelle cartelle che crea lei e legge il Drive per prendere i file di lavoro.</p>
  </section>
  <section class="card"><h3>Telefono</h3>
    <p class="muted" id="storeInfo">Memoria: calcolo in corso…</p>
    ${installPrompt ? `<button class="btn" type="button" data-act="install">Installa l'app</button>` : ""}
    <p class="muted">Versione ${VERSION}</p>
  </section>`;
}

/* ---------- azioni ---------- */
async function nuovo(cantId) {
  const c = cantId && S.prog ? S.prog.cantieri.find(k => k.id === cantId) : null;
  const x = { id: uid("v"), created: new Date().toISOString(), updated: new Date().toISOString(), stato: "bozza", sync: null, foto: [], dati: Core.nuovoVerbale(c, S.settings.esecutore) };
  x.dati.applicati = [];
  S.verbali.unshift(x); S.curId = x.id; S.ui.newSheet = false; S.ui.errors = [];
  await DB.put(x); await DB.set("curId", x.id); go("compila");
}
function applica(x, r) {
  const v = x.dati; v.applicati = v.applicati || [];
  if (v.applicati.includes(r.k)) return;
  if (r.voce === "oss") v.osservazioni = (v.osservazioni ? v.osservazioni.replace(/\s+$/, "") + "\n" : "") + "• " + r.text;
  else v.risposte[r.voce] = "NC";
  if (r.voce !== "oss") v.osservazioni = (v.osservazioni ? v.osservazioni.replace(/\s+$/, "") + "\n" : "") + `• ${r.voce} – ${r.text}`;
  v.applicati.push(r.k);
}
async function concludi() {
  const x = cur(); const v = x.dati; const e = [];
  if (!v.data) e.push("Indica la data del sopralluogo.");
  if (!v.autostrada) e.push("Seleziona l'autostrada.");
  const vuote = Core.ITEMS.filter(([c]) => !v.risposte[c]).map(([c]) => c);
  if (vuote.length) e.push(`Mancano ${vuote.length} voci: ${vuote.join(", ")}.`);
  if (!v.imprese.some(i => i.NOME)) e.push("Indica almeno un'impresa.");
  S.ui.errors = e;
  if (e.length) { render(); document.querySelector("[role=alert]")?.scrollIntoView({ block: "center" }); return; }
  x.stato = "completato";
  if (!x.sync || x.sync.stato !== "caricato") x.sync = { ...(x.sync || {}), stato: x.sync && x.sync.folderId ? "modificato" : "da_caricare" };
  await saveV(x, true);
  S.curId = null; await DB.set("curId", null);
  go("elenco");
  toast(navigator.onLine ? "Sopralluogo completato. Premi Sincronizza per caricarlo su Drive." : "Sopralluogo completato e salvato sul telefono. Lo carichi quando torna la rete.");
}
async function loadFile(kind, file) { return loadBuf(kind, file.name, await file.arrayBuffer(), { origine: "telefono", mod: new Date().toISOString() }); }
async function loadBuf(kind, name, buf, src, quiet) {
  const file = { name };
  if (kind === "loadTpl") {
    const z = await JSZip.loadAsync(buf); const xml = await z.file("word/document.xml").async("string");
    if (!xml.includes("{{DATA}}") || !xml.includes("{{FOTO}}")) throw new Error("Questo Word non è il modello preparato per l'app: usa Modello_sopralluogo_ASPI_app.docx.");
    await DB.set("template", { name: file.name, bytes: buf, at: new Date().toISOString() }); S.tpl = { name: file.name, at: new Date().toISOString() };
    toast("Modello Word caricato.");
  } else if (kind === "loadItp") {
    const r = Core.parseITP(XLSX, buf);
    S.itp = { name: file.name, at: new Date().toISOString(), src, foglio: r.foglio, operators: r.operators, imprese: r.imprese }; await DB.set("itp", S.itp);
    fillDatalists(); if (!quiet) toast(`File ITP caricato: ${r.operators.length} operatori.`);
    return `ITP ${r.operators.length} operatori`;
  } else if (kind === "loadProg") {
    const r = Core.parseProgramma(XLSX, buf);
    if (!r.cantieri.length) throw new Error("Nessun cantiere trovato: serve il programma settimanale (fogli per autostrada, dati dalla riga 5).");
    S.prog = { name: file.name, at: new Date().toISOString(), src, settimana: r.settimana, cantieri: r.cantieri }; await DB.set("programma", S.prog);
    fillDatalists(); if (!quiet) toast(`Programma caricato: ${r.cantieri.length} cantieri.`);
    return `programma ${r.cantieri.length} cantieri`;
  }
}
function fillDatalists() {
  const imp = new Set((S.itp && S.itp.imprese) || []);
  ((S.prog && S.prog.cantieri) || []).forEach(c => [c.impSegn, c.impLavori, ...c.prep.map(p => p.impresa)].forEach(n => n && imp.add(n)));
  document.getElementById("dl-imprese").innerHTML = [...imp].sort().map(n => `<option value="${esc(n)}">`).join("");
  document.getElementById("dl-persone").innerHTML = ((S.itp && S.itp.operators) || []).map(o => `<option value="${esc(o.nome)}">${esc([o.impApp, o.contratto].filter(Boolean).join(" · "))}</option>`).join("");
}
async function share(x) {
  const name = Core.nomeFile(x.dati);
  const files = [new File([await generaDocx(x)], name + ".docx", { type: DMIME }), new File([generaXlsx(x)], name + ".xlsx", { type: XMIME })];
  if (navigator.canShare && navigator.canShare({ files })) { try { await navigator.share({ files, title: name }); } catch (e) { if (e.name !== "AbortError") throw e; } }
  else { files.forEach(f => download(f, f.name)); toast("Condivisione non disponibile: file scaricati."); }
}

document.addEventListener("click", async e => {
  const el = e.target.closest("[data-act]"); if (!el) return;
  const a = el.dataset.act; if (el.tagName === "INPUT" || el.tagName === "SELECT") return;
  const x = cur();
  try {
    switch (a) {
      case "go": go(el.dataset.v); break;
      case "sync": await sincronizza(); break;
      case "newOpen": S.ui.newSheet = true; render(); break;
      case "newOpen2": S.ui.newSheet = true; go("elenco"); break;
      case "newClose": S.ui.newSheet = false; render(); break;
      case "newFrom": await nuovo(el.dataset.id); break;
      case "open": S.curId = el.dataset.id; S.ui.errors = []; await DB.set("curId", S.curId); go("compila"); break;
      case "pick": { const f = el.dataset.f, v = el.dataset.v; x.dati[f] = x.dati[f] === v ? "" : v; saveV(x); render(); break; }
      case "toggle": { const f = el.dataset.f, v = el.dataset.v; const arr = x.dati[f] || []; x.dati[f] = arr.includes(v) ? arr.filter(k => k !== v) : [...arr, v]; saveV(x); render(); break; }
      case "ans": { const c = el.dataset.c, v = el.dataset.v; x.dati.risposte[c] = x.dati.risposte[c] === v ? "" : v; saveV(x); render(); break; }
      case "restC": case "restNA": Core.ITEMS.forEach(([c]) => { if (!x.dati.risposte[c]) x.dati.risposte[c] = a === "restC" ? "C" : "NA"; }); saveV(x); render(); break;
      case "applica": { const r = Core.rilievi(x.dati, S.itp.operators).find(k => k.k === el.dataset.k); if (r) applica(x, r); saveV(x); render(); break; }
      case "applicaTutti": Core.rilievi(x.dati, S.itp.operators).forEach(r => applica(x, r)); saveV(x); render(); toast("Rilievi applicati: controlla A.2 e le osservazioni."); break;
      case "fotoDel": { const id = el.dataset.id; x.foto = x.foto.filter(f => f.id !== id); await DB.delFoto(id); saveV(x, true); render(); break; }
      case "concludi": await concludi(); break;
      case "docx": { const t = S.verbali.find(k => k.id === el.dataset.id); el.disabled = true; download(await generaDocx(t), Core.nomeFile(t.dati) + ".docx"); el.disabled = false; break; }
      case "xlsx": { const t = S.verbali.find(k => k.id === el.dataset.id); download(generaXlsx(t), Core.nomeFile(t.dati) + ".xlsx"); break; }
      case "share": await share(S.verbali.find(k => k.id === el.dataset.id)); break;
      case "delAsk": S.ui.delAsk = el.dataset.id; render(); break;
      case "delNo": S.ui.delAsk = null; render(); break;
      case "delDo": { const t = S.verbali.find(k => k.id === el.dataset.id); for (const f of t.foto) await DB.delFoto(f.id); await DB.remove(t.id); S.verbali = S.verbali.filter(k => k !== t); if (S.curId === t.id) { S.curId = null; await DB.set("curId", null); } S.ui.delAsk = null; render(); toast("Sopralluogo eliminato dal telefono (su Drive resta)."); break; }
      case "connect": case "pullDrive": {
        if (!navigator.onLine) { toast("Serve la rete per collegarsi a Drive."); break; }
        el.disabled = true; await Drive.auth(S.settings.clientId); await cartelle();
        const r = await aggiornaDaDrive(); render();
        toast(r.fatto.length ? "Aggiornati da Drive: " + r.fatto.join(", ") + "." : r.vuota ? `Drive collegato. Metti i file in “${CARTELLA_LAVORO}” per gli aggiornamenti automatici.` : "Drive collegato: i file di lavoro sono già aggiornati." + (r.note.length ? " " + r.note.join("; ") : ""));
        break; }
      case "install": if (installPrompt) { installPrompt.prompt(); await installPrompt.userChoice; installPrompt = null; render(); } break;
    }
  } catch (err) { toast(err.message || String(err)); if (el) el.disabled = false; }
});
document.addEventListener("input", e => {
  const el = e.target; const x = cur();
  if (el.dataset.f && x) { setPath(x.dati, el.dataset.f, el.value); saveV(x); return; }
  if (el.dataset.fnota && x) { const f = x.foto.find(k => k.id === el.dataset.fnota); if (f) { f.nota = el.value; saveV(x); } return; }
  if (el.dataset.s) { S.settings[el.dataset.s] = el.value.trim(); clearTimeout(saveSettings.t); saveSettings.t = setTimeout(saveSettings, 300); return; }
  if (el.dataset.act === "q") { S.ui.q = el.value; render(); }
});
document.addEventListener("change", async e => {
  const el = e.target; const a = el.dataset.act; const x = cur();
  try {
    if (a === "foto" && el.files.length) {
      toast("Elaboro le foto…");
      for (const file of el.files) { const { blob, w, h } = await compress(file); const id = uid("f"); await DB.putFoto(id, blob); x.foto.push({ id, w, h, voce: "", nota: "" }); }
      el.value = ""; await saveV(x, true); render(); toast("Foto aggiunte."); return;
    }
    if (a === "fotoVoce") { const f = x.foto.find(k => k.id === el.dataset.id); f.voce = el.value; saveV(x); return; }
    if (["loadTpl", "loadItp", "loadProg"].includes(a) && el.files[0]) { await loadFile(a, el.files[0]); el.value = ""; render(); return; }
    if (el.dataset.f && x && (/^imprese\./.test(el.dataset.f) || ["contratti", "data"].includes(el.dataset.f))) { setPath(x.dati, el.dataset.f, el.dataset.f.startsWith("imprese.") ? el.value.toUpperCase().trim() : el.value); saveV(x); render(); }
  } catch (err) { toast(err.message || String(err)); el.value = ""; }
});
window.addEventListener("hashchange", () => { const h = location.hash.slice(1); if (VIEWS.some(v => v[0] === h) && h !== S.view) { S.view = h; render(); } });
window.addEventListener("online", renderBar); window.addEventListener("offline", renderBar);
window.addEventListener("beforeinstallprompt", e => { e.preventDefault(); installPrompt = e; render(); });

/* ---------- avvio ---------- */
(async function boot() {
  try {
    S.settings = Object.assign({ esecutore: "DANILO CORAZZINI CANTISANO", clientId: CLIENT_ID_PREDEFINITO, rootId: "", rootUrl: "", registroId: "" }, (await DB.get("settings")) || {});
    if (!S.settings.clientId) S.settings.clientId = CLIENT_ID_PREDEFINITO;
    const t = await DB.get("template"); S.tpl = t ? { name: t.name, at: t.at } : null;
    S.itp = (await DB.get("itp")) || null; S.prog = (await DB.get("programma")) || null;
    S.verbali = ((await DB.all()) || []).sort((a, b) => b.updated.localeCompare(a.updated));
    S.curId = (await DB.get("curId")) || null;
  } catch (e) { toast("Archivio locale non disponibile: " + e.message); }
  const h = location.hash.slice(1); if (VIEWS.some(v => v[0] === h)) S.view = h;
  fillDatalists(); render();
  if (navigator.storage && navigator.storage.persist) navigator.storage.persist().catch(() => {});
  if ("serviceWorker" in navigator && location.protocol === "https:") navigator.serviceWorker.register("sw.js").catch(() => {});
})();
