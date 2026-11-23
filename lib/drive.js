/* Google Drive: accesso con Google Identity Services (solo ambito drive.file: l'app vede solo i file che crea). */
const Drive = (() => {
  "use strict";
  // drive.file: scrive solo nelle cartelle create dall'app; drive.readonly: legge i file di lavoro messi a mano
  const SCOPE = "https://www.googleapis.com/auth/drive.file https://www.googleapis.com/auth/drive.readonly";
  const API = "https://www.googleapis.com/drive/v3/files";
  const UP = "https://www.googleapis.com/upload/drive/v3/files";
  let token = null, exp = 0, gis = null;
  function loadGis() {
    return gis || (gis = new Promise((res, rej) => {
      const s = document.createElement("script");
      s.src = "https://accounts.google.com/gsi/client"; s.async = true;
      s.onload = () => res(); s.onerror = () => { gis = null; rej(new Error("Servizio di accesso Google non raggiungibile: controlla la connessione.")); };
      document.head.appendChild(s);
    }));
  }
  async function auth(clientId) {
    if (token && Date.now() < exp) return token;
    if (!clientId) throw new Error("Manca il Client ID Google: inseriscilo in Impostazioni.");
    await loadGis();
    return new Promise((res, rej) => {
      const tc = google.accounts.oauth2.initTokenClient({
        client_id: clientId, scope: SCOPE,
        callback: r => {
          if (r.error) return rej(new Error("Accesso Google non riuscito: " + r.error));
          if (google.accounts.oauth2.hasGrantedAllScopes && !google.accounts.oauth2.hasGrantedAllScopes(r, ...SCOPE.split(" "))) return rej(new Error("Per funzionare l'app ha bisogno di entrambi i permessi richiesti: riprova e spunta tutte le caselle."));
          token = r.access_token; exp = Date.now() + (Number(r.expires_in || 3600) - 120) * 1000; res(token); },
        error_callback: e => rej(new Error(e && e.type === "popup_closed" ? "Hai chiuso la finestra di accesso Google." : "Accesso Google non riuscito" + (e && e.type ? " (" + e.type + ")" : "") + "."))
      });
      tc.requestAccessToken({ prompt: "" });
    });
  }
  async function api(url, opts = {}) {
    const r = await fetch(url, { ...opts, headers: { Authorization: "Bearer " + token, ...(opts.headers || {}) } });
    if (r.status === 401) { token = null; throw Object.assign(new Error("Sessione Google scaduta: premi di nuovo Sincronizza."), { status: 401 }); }
    if (!r.ok) { const t = await r.text(); throw Object.assign(new Error(`Errore Drive ${r.status}: ${t.slice(0, 160)}`), { status: r.status }); }
    return r.status === 204 ? null : r.json();
  }
  async function exists(id) {
    if (!id) return false;
    try { const f = await api(`${API}/${id}?fields=id,trashed`); return !!f && !f.trashed; }
    catch (e) { if (e.status === 404) return false; throw e; }
  }
  // root=true: cerca solo la cartella principale creata dall'app (marcata), non omonime create a mano
  async function folder(name, parent, root) {
    const qn = name.replace(/\\/g, "\\\\").replace(/'/g, "\\'");
    let q = `name='${qn}' and mimeType='application/vnd.google-apps.folder' and trashed=false` + (parent ? ` and '${parent}' in parents` : "");
    if (root) q += " and appProperties has { key='sopralluoghiApp' and value='root' }";
    const l = await api(`${API}?q=${encodeURIComponent(q)}&fields=files(id,webViewLink)&spaces=drive`);
    if (l.files && l.files.length) return l.files[0];
    return api(`${API}?fields=id,webViewLink`, { method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ name, mimeType: "application/vnd.google-apps.folder", ...(parent ? { parents: [parent] } : {}), ...(root ? { appProperties: { sopralluoghiApp: "root" } } : {}) }) });
  }
  async function list(parent) {
    const q = `'${parent}' in parents and trashed=false and mimeType != 'application/vnd.google-apps.folder'`;
    const l = await api(`${API}?q=${encodeURIComponent(q)}&fields=files(id,name,mimeType,modifiedTime,size)&orderBy=modifiedTime desc&pageSize=100&supportsAllDrives=true&includeItemsFromAllDrives=true`);
    return l.files || [];
  }
  async function download(f) {
    const url = f.mimeType === "application/vnd.google-apps.spreadsheet"
      ? `${API}/${f.id}/export?mimeType=${encodeURIComponent("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")}`
      : `${API}/${f.id}?alt=media&supportsAllDrives=true`;
    const r = await fetch(url, { headers: { Authorization: "Bearer " + token } });
    if (!r.ok) throw new Error(`download non riuscito (${r.status})`);
    return r.arrayBuffer();
  }
  // crea (fileId assente) o aggiorna un file
  async function upload(name, mime, blob, parent, fileId) {
    const meta = fileId ? { name } : { name, parents: [parent] };
    const b = "sopr" + Math.random().toString(36).slice(2);
    const body = new Blob([`--${b}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n${JSON.stringify(meta)}\r\n--${b}\r\nContent-Type: ${mime}\r\n\r\n`, blob, `\r\n--${b}--`]);
    const url = fileId ? `${UP}/${fileId}?uploadType=multipart&fields=id,webViewLink` : `${UP}?uploadType=multipart&fields=id,webViewLink`;
    try { return await api(url, { method: fileId ? "PATCH" : "POST", headers: { "Content-Type": `multipart/related; boundary=${b}` }, body }); }
    catch (e) { if (fileId && e.status === 404) return upload(name, mime, blob, parent, null); throw e; }
  }
  return { auth, folder, list, download, upload, exists, connected: () => !!token && Date.now() < exp };
})();
