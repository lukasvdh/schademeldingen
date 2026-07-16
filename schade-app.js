/* ============================================================================
 *  SCHADEMELDINGEN — zelfstandige app voor Cloudflare Pages
 *  MSAL (delegated) + Microsoft Graph -> eigen SharePoint-lijsten + fotobibliotheek
 * ==========================================================================
 *
 *  Zie STAPPENPLAN.md voor de volledige installatie.
 *  Extern geladen: verpa-ui.css, schade-app.css, lucide icons
 * ============================================================================ */
const CONFIG = {
  clientId:    "b2bbb9ba-b6f1-4046-91fe-cc8fa5df7a7b",
  tenantId:    "e65dbe4b-d1e2-4283-b0f5-aa7717e81077",
  redirectUri: "https://verpa-schade.pages.dev",
  siteHostname:"verpabenelux.sharepoint.com",
  sitePath:    "/sites/OfficeData",
  listName:    "Schademeldingen",
  attachFolder:"Schadefotos",
  settingsList:"SchadeInstellingen",
  commentsList:"SchadeReacties",
  adminRole:   "Admin",
  scopes:      ["User.Read", "Sites.ReadWrite.All", "Files.ReadWrite.All"],
  mailEnabled: false,
};

const GRAPH = "https://graph.microsoft.com/v1.0";

/* ------------------------------------------------------------------ *
 *  Fasen + velden (single source of truth)
 * ------------------------------------------------------------------ */
const COMMON_HEAD = [
  { name: "warehouse", label: "Magazijn", type: "select", options: ["LAAKDAL", "STORA"], required: true },
  { name: "artikelnummer", label: "Artikelnummer", type: "text", placeholder: "bv. 100234", required: true, mono: true },
  { name: "omschrijving", label: "Omschrijving", type: "text", placeholder: "Artikelomschrijving", full: true },
  { name: "aantal", label: "Aantal beschadigd", type: "number", placeholder: "0", required: true },
];
const COMMON_TAIL = [
  { name: "prioriteit", label: "Prioriteit", type: "select", options: ["Laag", "Normaal", "Hoog"], required: true },
  { name: "melder", label: "Gemeld door", type: "text", placeholder: "Naam" },
  { name: "opmerkingen", label: "Opmerkingen", type: "textarea", placeholder: "Extra details over de schade…", full: true },
];
const STAGES = [
  { key: "inkomend", stap: "Ontvangst", label: "Inkomende goederen", sub: "Schade vastgesteld bij ontvangst", icon: "arrow-down-to-line",
    fields: [
      { name: "leverancier", label: "Leverancier", type: "text", placeholder: "Naam leverancier" },
      { name: "inkooporder", label: "Inkooporder-nr.", type: "text", placeholder: "BC inkooporder", mono: true },
      { name: "ontvangstdatum", label: "Ontvangstdatum", type: "date" },
      { name: "vervoerder", label: "Vervoerder", type: "text", placeholder: "Transporteur" },
      { name: "typeSchade", label: "Soort schade", type: "select", options: ["Transportschade", "Verpakkingsschade", "Productdefect", "Nat / vochtig", "Ontbrekend"] },
      { name: "actie", label: "Actie bij ontvangst", type: "select", options: ["Geweigerd", "Onder voorbehoud aangenomen", "Volledig aangenomen", "Retour naar leverancier"] },
    ] },
  { key: "voorraad", stap: "Voorraad", label: "Schade aan voorraad", sub: "Tijdens opslag of orderverwerking", icon: "warehouse",
    fields: [
      { name: "locatie", label: "Locatie / bin", type: "text", placeholder: "bv. A-12-03", mono: true },
      { name: "oorzaak", label: "Oorzaak", type: "select", options: ["Handling (val)", "Heftruck / transpallet", "THT verlopen", "Waterschade", "Mispick-schade", "Onbekend"] },
      { name: "ontdektBij", label: "Ontdekt bij", type: "select", options: ["Orderpicking", "Cyclustelling", "Routine-inspectie", "Bijvullen"] },
      { name: "gekoppeldeOrder", label: "Gekoppelde order (optioneel)", type: "text", placeholder: "Verkooporder tijdens picking", mono: true },
      { name: "afhandeling", label: "Afhandeling", type: "select", options: ["Afgeschreven", "Afgeprijsd", "Hersteld", "In quarantaine"] },
    ] },
  { key: "uitgaand", stap: "Verzending", label: "Uitgaande goederen", sub: "Schade vóór of bij verzending", icon: "package-check",
    fields: [
      { name: "klant", label: "Klant", type: "text", placeholder: "Naam klant" },
      { name: "verkooporder", label: "Verkooporder-nr.", type: "text", placeholder: "BC verkooporder", mono: true },
      { name: "ontdektTijdens", label: "Ontdekt tijdens", type: "select", options: ["Verpakken", "Laden", "Eindcontrole"] },
      { name: "vervoerder", label: "Vervoerder", type: "text", placeholder: "Transporteur" },
      { name: "actie", label: "Actie", type: "select", options: ["Opnieuw picken", "Vervangen", "Order gesplitst", "Verzending uitgesteld"] },
    ] },
  { key: "levering", stap: "Levering", label: "Schade bij levering", sub: "Na aflevering gemeld", icon: "truck",
    fields: [
      { name: "klant", label: "Klant", type: "text", placeholder: "Naam klant" },
      { name: "verkooporder", label: "Verkooporder-nr.", type: "text", placeholder: "BC verkooporder", mono: true },
      { name: "leveringsdatum", label: "Leveringsdatum", type: "date" },
      { name: "vervoerder", label: "Vervoerder", type: "text", placeholder: "Transporteur" },
      { name: "gemeldDoor", label: "Gemeld door", type: "select", options: ["Klant", "Chauffeur", "Vertegenwoordiger"] },
      { name: "claimVervoerder", label: "Claim bij vervoerder", type: "select", options: ["Ja", "Nee", "In behandeling"] },
      { name: "oplossing", label: "Oplossing", type: "select", options: ["Creditnota", "Herlevering", "Retour", "Geen actie"] },
    ] },
];
const stageByKey = (k) => STAGES.find((s) => s.key === k);
const KEY_TO_FASE = { inkomend: "Ontvangst", voorraad: "Voorraad", uitgaand: "Verzending", levering: "Levering" };
const FASE_TO_KEY = Object.fromEntries(Object.entries(KEY_TO_FASE).map(([k, v]) => [v, k]));

const STATUS_FLOW = ["Nieuw", "In behandeling", "Afgehandeld"];
const STATUS_CSS = {
  "Nieuw":          "status-nieuw",
  "In behandeling": "status-in-behandeling",
  "Afgehandeld":    "status-afgehandeld",
};
const STATUS_ICON = {
  "Nieuw":          "circle",
  "In behandeling": "clock",
  "Afgehandeld":    "check-circle-2",
};

const PRIO_CSS = { Hoog: "prio-hoog", Normaal: "prio-normaal", Laag: "prio-laag" };

const COLS = {
  prioriteit: { col: "Prioriteit" }, warehouse: { col: "Magazijn" }, artikelnummer: { col: "Artikelnummer" },
  omschrijving: { col: "Omschrijving" }, aantal: { col: "Aantal", kind: "number" }, melder: { col: "Melder" },
  opmerkingen: { col: "Opmerkingen" }, aangemaakt: { col: "Melddatum", kind: "date" }, fotos: { col: "Fotos", kind: "json" },
  leverancier: { col: "Leverancier" }, inkooporder: { col: "Inkooporder" }, ontvangstdatum: { col: "Ontvangstdatum", kind: "date" },
  vervoerder: { col: "Vervoerder" }, typeSchade: { col: "SoortSchade" }, actie: { col: "Actie" },
  locatie: { col: "Locatie" }, oorzaak: { col: "Oorzaak" }, ontdektBij: { col: "OntdektBij" },
  gekoppeldeOrder: { col: "GekoppeldeOrder" }, afhandeling: { col: "Afhandeling" }, klant: { col: "Klant" },
  verkooporder: { col: "Verkooporder" }, ontdektTijdens: { col: "OntdektTijdens" }, leveringsdatum: { col: "Leveringsdatum", kind: "date" },
  gemeldDoor: { col: "GemeldDoor" }, claimVervoerder: { col: "ClaimVervoerder" }, oplossing: { col: "Oplossing" },
};
const COLUMN_DEFS = [
  { name: "Fase", choice: { choices: ["Ontvangst", "Voorraad", "Verzending", "Levering"], displayAs: "dropDownMenu" } },
  { name: "Status", choice: { choices: STATUS_FLOW } },
  { name: "Prioriteit", choice: { choices: ["Laag", "Normaal", "Hoog"] } },
  { name: "Magazijn", choice: { choices: ["LAAKDAL", "STORA"] } },
  { name: "Artikelnummer", text: {} }, { name: "Omschrijving", text: {} }, { name: "Aantal", number: {} },
  { name: "Melder", text: {} }, { name: "Opmerkingen", text: { allowMultipleLines: true } },
  { name: "Melddatum", dateTime: { format: "dateOnly", displayAs: "standard" } },
  { name: "Fotos", text: { allowMultipleLines: true } },
  { name: "Leverancier", text: {} }, { name: "Inkooporder", text: {} },
  { name: "Ontvangstdatum", dateTime: { format: "dateOnly", displayAs: "standard" } },
  { name: "Vervoerder", text: {} },
  { name: "SoortSchade", choice: { choices: ["Transportschade", "Verpakkingsschade", "Productdefect", "Nat / vochtig", "Ontbrekend"] } },
  { name: "Actie", text: {} }, { name: "Locatie", text: {} },
  { name: "Oorzaak", choice: { choices: ["Handling (val)", "Heftruck / transpallet", "THT verlopen", "Waterschade", "Mispick-schade", "Onbekend"] } },
  { name: "OntdektBij", choice: { choices: ["Orderpicking", "Cyclustelling", "Routine-inspectie", "Bijvullen"] } },
  { name: "GekoppeldeOrder", text: {} },
  { name: "Afhandeling", choice: { choices: ["Afgeschreven", "Afgeprijsd", "Hersteld", "In quarantaine"] } },
  { name: "Klant", text: {} }, { name: "Verkooporder", text: {} },
  { name: "OntdektTijdens", choice: { choices: ["Verpakken", "Laden", "Eindcontrole"] } },
  { name: "Leveringsdatum", dateTime: { format: "dateOnly", displayAs: "standard" } },
  { name: "GemeldDoor", choice: { choices: ["Klant", "Chauffeur", "Vertegenwoordiger"] } },
  { name: "ClaimVervoerder", choice: { choices: ["Ja", "Nee", "In behandeling"] } },
  { name: "Oplossing", choice: { choices: ["Creditnota", "Herlevering", "Retour", "Geen actie"] } },
];

/* ------------------------------------------------------------------ *
 *  Helpers
 * ------------------------------------------------------------------ */
const $ = (sel, root = document) => root.querySelector(sel);
const todayISO = () => new Date().toISOString().slice(0, 10);
const esc = (v) => String(v == null ? "" : v).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const fmtDate = (iso) => { if (!iso) return "—"; try { return new Date(iso).toLocaleDateString("nl-BE", { day: "2-digit", month: "short", year: "numeric" }); } catch { return iso; } };
const fmtDateTime = (iso) => { if (!iso) return ""; try { return new Date(iso).toLocaleString("nl-BE", { day: "2-digit", month: "short", hour: "2-digit", minute: "2-digit" }); } catch { return iso; } };
const icons = () => window.lucide && window.lucide.createIcons();

/* ------------------------------------------------------------------ *
 *  MSAL — laden + auth
 * ------------------------------------------------------------------ */
function loadScript(src) {
  return new Promise((resolve, reject) => {
    const s = document.createElement("script");
    s.src = src; s.async = true;
    s.onload = () => resolve(src);
    s.onerror = () => reject(new Error("kon niet laden: " + src));
    document.head.appendChild(s);
  });
}
async function loadMsal() {
  if (window.msal) return;
  const sources = [
    "https://cdn.jsdelivr.net/npm/@azure/msal-browser@2.38.4/lib/msal-browser.min.js",
    "https://unpkg.com/@azure/msal-browser@2.38.4/lib/msal-browser.min.js",
    "https://alcdn.msauth.net/browser/2.38.4/js/msal-browser.min.js",
  ];
  for (const src of sources) {
    try { await loadScript(src); if (window.msal) return; } catch (e) { console.warn("MSAL-bron faalde:", src); }
  }
  throw new Error("MSAL kon uit geen enkele bron geladen worden.");
}

let pca = null;
async function initAuth() {
  await loadMsal();
  pca = new msal.PublicClientApplication({
    auth: { clientId: CONFIG.clientId, authority: `https://login.microsoftonline.com/${CONFIG.tenantId}`, redirectUri: CONFIG.redirectUri },
    cache: { cacheLocation: "localStorage", storeAuthStateInCookie: false },
  });
  const resp = await pca.handleRedirectPromise();
  if (resp && resp.account) pca.setActiveAccount(resp.account);
  const acc = pca.getActiveAccount() || pca.getAllAccounts()[0];
  if (acc) pca.setActiveAccount(acc);
  return acc || null;
}
function login() { pca.loginRedirect({ scopes: CONFIG.scopes }); }
function logout() { pca.logoutRedirect(); }
async function getToken() {
  const account = pca.getActiveAccount();
  const r = await pca.acquireTokenSilent({ scopes: CONFIG.scopes, account }).catch(async () => {
    await pca.acquireTokenRedirect({ scopes: CONFIG.scopes, account }); return null;
  });
  return r && r.accessToken;
}
function isAdmin() {
  const a = pca && pca.getActiveAccount();
  const roles = (a && a.idTokenClaims && a.idTokenClaims.roles) || [];
  return roles.includes(CONFIG.adminRole);
}

/* ------------------------------------------------------------------ *
 *  Graph datalaag
 * ------------------------------------------------------------------ */
async function graph(pathOrUrl, { method = "GET", body, headers } = {}) {
  const token = await getToken();
  const url = pathOrUrl.startsWith("http") ? pathOrUrl : GRAPH + pathOrUrl;
  const res = await fetch(url, {
    method,
    headers: { Authorization: `Bearer ${token}`, ...(body ? { "Content-Type": "application/json" } : {}), ...(headers || {}) },
    body: body ? JSON.stringify(body) : undefined,
  });
  if (!res.ok) throw new Error(`Graph ${method} ${res.status}: ${await res.text().catch(() => "")}`);
  return res.status === 204 ? null : res.json();
}

let _ctx = null;
let INTERNAL = null;
function internal(name) { return (INTERNAL && INTERNAL[name]) || name; }
async function getSiteId() {
  const site = await graph(`/sites/${CONFIG.siteHostname}:${CONFIG.sitePath}`);
  return site.id;
}
async function ensureList() {
  if (_ctx) return _ctx;
  const siteId = await getSiteId();
  const found = await graph(`/sites/${siteId}/lists?$filter=displayName eq '${encodeURIComponent(CONFIG.listName)}'`);
  if (!found.value || !found.value.length) throw new Error(`SharePoint-lijst "${CONFIG.listName}" niet gevonden.`);
  const listId = found.value[0].id;
  _ctx = { siteId, listId };
  const cols = await graph(`/sites/${siteId}/lists/${listId}/columns?$select=name,displayName`);
  INTERNAL = {};
  (cols.value || []).forEach((c) => { if (c.displayName) INTERNAL[c.displayName] = c.name; INTERNAL[c.name] = c.name; });
  return _ctx;
}
function encode(def, value) {
  if (def.kind === "number") return Number(value) || 0;
  if (def.kind === "date") return `${String(value).slice(0, 10)}T00:00:00Z`;
  if (def.kind === "json") return JSON.stringify(value || []);
  return value;
}
function toFields(rec) {
  const f = {};
  if (rec.id != null) f[internal("Title")] = rec.id;
  if (rec.type) f[internal("Fase")] = KEY_TO_FASE[rec.type] || rec.type;
  if (rec.status) f[internal("Status")] = rec.status;
  for (const [k, def] of Object.entries(COLS)) {
    const v = rec[k];
    if (v === undefined || v === null || v === "") continue;
    f[internal(def.col)] = encode(def, v);
  }
  return f;
}
function fromItem(item) {
  const f = item.fields || {};
  const gv = (n) => f[internal(n)];
  const rec = { spId: item.id, id: gv("Title"), status: gv("Status") || "Nieuw", type: FASE_TO_KEY[gv("Fase")] || gv("Fase") };
  for (const [k, def] of Object.entries(COLS)) {
    const raw = gv(def.col);
    if (raw === undefined || raw === null || raw === "") continue;
    if (def.kind === "date") rec[k] = String(raw).slice(0, 10);
    else if (def.kind === "number") rec[k] = Number(raw);
    else if (def.kind === "json") { try { rec[k] = JSON.parse(raw); } catch { rec[k] = []; } }
    else rec[k] = raw;
  }
  return rec;
}
async function fetchReports() {
  const { siteId, listId } = await ensureList();
  const out = [];
  let url = `/sites/${siteId}/lists/${listId}/items?expand=fields&$top=200`;
  while (url) { const d = await graph(url); out.push(...(d.value || []).map(fromItem)); url = d["@odata.nextLink"] || null; }
  return out.sort((a, b) => (String(a.aangemaakt) < String(b.aangemaakt) ? 1 : -1));
}
async function createReport(rec) {
  const { siteId, listId } = await ensureList();
  const created = await graph(`/sites/${siteId}/lists/${listId}/items`, { method: "POST", body: { fields: toFields(rec) } });
  return fromItem(created);
}
async function updateReport(spId, patch) {
  const { siteId, listId } = await ensureList();
  await graph(`/sites/${siteId}/lists/${listId}/items/${spId}/fields`, { method: "PATCH", body: toFields(patch) });
}
async function deleteReport(spId) {
  const { siteId, listId } = await ensureList();
  await graph(`/sites/${siteId}/lists/${listId}/items/${spId}`, { method: "DELETE" });
}

/* ------------------------------------------------------------------ *
 *  Instellingen
 * ------------------------------------------------------------------ */
let _settingsCtx = null;
async function ensureSettingsList() {
  if (_settingsCtx) return _settingsCtx;
  const siteId = await getSiteId();
  const found = await graph(`/sites/${siteId}/lists?$filter=displayName eq '${encodeURIComponent(CONFIG.settingsList)}'`);
  if (!found.value || !found.value.length) throw new Error(`SharePoint-lijst "${CONFIG.settingsList}" niet gevonden.`);
  _settingsCtx = { siteId, listId: found.value[0].id };
  return _settingsCtx;
}
async function loadSettings() {
  try {
    const { siteId, listId } = await ensureSettingsList();
    const d = await graph(`/sites/${siteId}/lists/${listId}/items?expand=fields&$top=10`);
    const items = d.value || [];
    const item = items.find((it) => (it.fields || {}).Title === "config") || items[0];
    if (!item) return { spId: null, enabled: true, recipients: [] };
    const f = item.fields || {};
    let recipients = []; try { recipients = JSON.parse(f.Ontvangers || "[]"); } catch { recipients = []; }
    return { spId: item.id, enabled: f.MeldingenAan !== false, recipients };
  } catch (e) { console.warn("Instellingen laden mislukt:", e); return { spId: null, enabled: true, recipients: [] }; }
}
async function saveSettings(settings) {
  const { siteId, listId } = await ensureSettingsList();
  const fields = { Title: "config", Ontvangers: JSON.stringify(settings.recipients || []), MeldingenAan: !!settings.enabled };
  if (settings.spId) await graph(`/sites/${siteId}/lists/${listId}/items/${settings.spId}/fields`, { method: "PATCH", body: fields });
  else { const created = await graph(`/sites/${siteId}/lists/${listId}/items`, { method: "POST", body: { fields } }); settings.spId = created.id; }
  return settings;
}

/* ------------------------------------------------------------------ *
 *  E-mailmelding
 * ------------------------------------------------------------------ */
function appCaseUrl(id) { return `${location.origin}${location.pathname}?case=${encodeURIComponent(id)}`; }
async function sendMail(recipients, subject, html) {
  const to = recipients.map((r) => ({ emailAddress: { address: typeof r === "string" ? r : r.email } }));
  await graph(`/me/sendMail`, {
    method: "POST",
    body: { message: { subject, body: { contentType: "HTML", content: html }, toRecipients: to }, saveToSentItems: true },
  });
}
async function notifyNewCase(rec) {
  if (!CONFIG.mailEnabled) return;
  const s = S.settings || {};
  if (!s.enabled) return;
  const actief = (s.recipients || []).filter((r) => {
    if (!r || !r.email) return false;
    const fasen = Array.isArray(r.fasen) && r.fasen.length ? r.fasen : ["inkomend","voorraad","uitgaand","levering"];
    return fasen.includes(rec.type);
  });
  if (!actief.length) return;
  await sendMail(actief, `Nieuwe schademelding ${rec.id} \u2014 ${stageByKey(rec.type).stap}`, buildEmailHtml(rec));
}
function buildEmailHtml(rec) {
  const stage = stageByKey(rec.type);
  const HEX = {
    inkomend: ["#2563EB", "#DBEAFE", "#1E40AF"], voorraad: ["#D97706", "#FEF3C7", "#92400E"],
    uitgaand: ["#7C3AED", "#EDE9FE", "#5B21B6"], levering: ["#E11D48", "#FFE4E6", "#9F1239"],
  };
  const [main, soft, dark] = HEX[rec.type] || HEX.inkomend;
  const url = appCaseUrl(rec.id);
  const row = (label, val) => val ? `<tr><td style="padding:6px 0;color:#64748b;font-size:13px;width:150px;vertical-align:top;">${label}</td><td style="padding:6px 0;color:#0f172a;font-size:13px;font-weight:600;">${esc(val)}</td></tr>` : "";
  const extra = rec.klant ? row("Klant", rec.klant) : (rec.leverancier ? row("Leverancier", rec.leverancier) : "");
  const prio = rec.prioriteit === "Hoog"
    ? `<span style="display:inline-block;background:#FEE2E2;color:#991B1B;font-size:12px;font-weight:700;padding:2px 8px;border-radius:9999px;">Hoog</span>`
    : `<span style="color:#0f172a;font-size:13px;font-weight:600;">${esc(rec.prioriteit || "Normaal")}</span>`;
  return `<!DOCTYPE html><html><body style="margin:0;padding:0;background:#f1f5f9;">
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#f1f5f9;padding:24px 12px;"><tr><td align="center">
      <table role="presentation" width="560" cellpadding="0" cellspacing="0" style="max-width:560px;width:100%;background:#ffffff;border-radius:16px;overflow:hidden;border:1px solid #e2e8f0;font-family:'Segoe UI',Roboto,Helvetica,Arial,sans-serif;">
        <tr><td style="background:${main};padding:20px 24px;">
          <div style="color:#ffffff;font-size:12px;font-weight:600;letter-spacing:.4px;text-transform:uppercase;opacity:.9;">Nieuwe schademelding</div>
          <div style="color:#ffffff;font-size:20px;font-weight:700;margin-top:2px;">${esc(stage.stap)} · ${esc(rec.id)}</div>
        </td></tr>
        <tr><td style="padding:22px 24px 6px;">
          <span style="display:inline-block;background:${soft};color:${dark};font-size:12px;font-weight:700;padding:4px 10px;border-radius:9999px;">${esc(stage.label)}</span>
          <p style="color:#334155;font-size:14px;line-height:1.5;margin:14px 0 8px;">Er is een nieuw schadegeval geregistreerd${rec.melder ? ` door ${esc(rec.melder)}` : ""}.</p>
          <table role="presentation" width="100%" cellpadding="0" cellspacing="0">
            ${row("Magazijn", rec.warehouse)}
            ${row("Artikel", [rec.artikelnummer, rec.omschrijving].filter(Boolean).join(" — "))}
            ${row("Aantal", rec.aantal != null ? rec.aantal + " stuks" : "")}
            ${extra}
            <tr><td style="padding:6px 0;color:#64748b;font-size:13px;">Prioriteit</td><td style="padding:6px 0;">${prio}</td></tr>
            ${row("Datum", fmtDate(rec.aangemaakt))}
          </table>
        </td></tr>
        <tr><td style="padding:10px 24px 26px;">
          <table role="presentation" cellpadding="0" cellspacing="0"><tr><td style="border-radius:10px;background:${main};">
            <a href="${url}" style="display:inline-block;padding:12px 22px;color:#ffffff;font-size:14px;font-weight:600;text-decoration:none;border-radius:10px;">Bekijk in de app →</a>
          </td></tr></table>
        </td></tr>
        <tr><td style="padding:14px 24px;background:#f8fafc;border-top:1px solid #e2e8f0;color:#94a3b8;font-size:12px;">Verpa Benelux · Schademeldingen</td></tr>
      </table>
    </td></tr></table>
  </body></html>`;
}

/* ------------------------------------------------------------------ *
 *  Gesprek (reacties)
 * ------------------------------------------------------------------ */
let _commentsCtx = null;
async function ensureCommentsList() {
  if (_commentsCtx) return _commentsCtx;
  const siteId = await getSiteId();
  const found = await graph(`/sites/${siteId}/lists?$filter=displayName eq '${encodeURIComponent(CONFIG.commentsList)}'`);
  if (!found.value || !found.value.length) throw new Error(`SharePoint-lijst "${CONFIG.commentsList}" niet gevonden.`);
  _commentsCtx = { siteId, listId: found.value[0].id };
  return _commentsCtx;
}
async function fetchComments(meldingId) {
  const { siteId, listId } = await ensureCommentsList();
  const q = String(meldingId).replace(/'/g, "''");
  const d = await graph(
    `/sites/${siteId}/lists/${listId}/items?expand=fields&$top=200&$filter=fields/Title eq '${q}'`,
    { headers: { Prefer: "HonorNonIndexedQueriesWarningMayFailRandomly=true" } }
  );
  return (d.value || [])
    .map((it) => ({ id: it.id, auteur: (it.fields || {}).Auteur, bericht: (it.fields || {}).Bericht, datum: it.createdDateTime }))
    .sort((a, b) => (String(a.datum) < String(b.datum) ? -1 : 1));
}
async function addComment(meldingId, auteur, bericht) {
  const { siteId, listId } = await ensureCommentsList();
  const created = await graph(`/sites/${siteId}/lists/${listId}/items`, {
    method: "POST", body: { fields: { Title: String(meldingId), Auteur: auteur, Bericht: bericht } },
  });
  return { id: created.id, auteur, bericht, datum: created.createdDateTime };
}

/* Foto's */
const CHUNK = 5 * 1024 * 1024;
const encPath = (p) => p.split("/").map(encodeURIComponent).join("/");
async function uploadFotos(meldingId, files) {
  const { siteId } = await ensureList();
  const refs = [];
  for (const file of files) {
    const path = `${CONFIG.attachFolder}/${meldingId}/${file.name}`;
    const session = await graph(`/sites/${siteId}/drive/root:/${encPath(path)}:/createUploadSession`,
      { method: "POST", body: { item: { "@microsoft.graph.conflictBehavior": "rename" } } });
    const uploadUrl = session.uploadUrl;
    const total = file.size; let start = 0, item = null;
    do {
      const end = Math.min(start + CHUNK, total);
      const res = await fetch(uploadUrl, { method: "PUT", headers: { "Content-Range": `bytes ${start}-${end - 1}/${total}` }, body: file.slice(start, end) });
      if (res.status === 200 || res.status === 201) item = await res.json();
      else if (res.status !== 202) throw new Error(`Foto-upload ${res.status}`);
      start = end;
    } while (start < total);
    refs.push({ name: file.name, driveId: item.parentReference.driveId, itemId: item.id });
  }
  return refs;
}
const fotoCache = new Map();
async function getFotoBlobUrl(ref) {
  if (ref.url) return ref.url;
  if (ref.itemId && fotoCache.has(ref.itemId)) return fotoCache.get(ref.itemId);
  const token = await getToken();
  const res = await fetch(`${GRAPH}/drives/${ref.driveId}/items/${ref.itemId}/content`, { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) throw new Error("Foto laden mislukt");
  const url = URL.createObjectURL(await res.blob());
  if (ref.itemId) fotoCache.set(ref.itemId, url);
  return url;
}

/* ==================================================================== *
 *  STATE + VIEWS
 * ==================================================================== */
const S = {
  account: null, isAdmin: false, settings: null, reports: [], loading: true, error: null,
  view: { name: "dashboard" },
  filters: { q: "", status: "Alle", stage: "Alle", warehouse: "Alle" },
};
let formPhotos = [];

function go(view) { S.view = view; render(); window.scrollTo(0, 0); }

function nextId() {
  const year = new Date().getFullYear();
  const nums = S.reports.map((r) => r.id).filter((id) => id && id.startsWith(`SCH-${year}-`)).map((id) => parseInt(id.split("-")[2], 10));
  return `SCH-${year}-${String((nums.length ? Math.max(...nums) : 0) + 1).padStart(4, "0")}`;
}

/* ── Shell (topbar) ── */
function shell(content) {
  const name = S.account ? esc(S.account.name || S.account.username || "") : "";
  return `
  <div class="v-topbar">
    <button data-nav='dashboard' class="v-topbar-brand">
      <div class="v-logo"><i data-lucide="alert-triangle" style="width:18px;height:18px;"></i></div>
      <div>
        <div class="v-topbar-title">Schademeldingen</div>
        <div class="v-topbar-sub">Verpa Benelux · Magazijn</div>
      </div>
    </button>
    <div class="v-topbar-right">
      ${S.isAdmin ? `<span class="v-role-badge">Beheerder</span>` : ""}
      <span class="v-user-name" style="max-width:9rem;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">${name}</span>
      ${S.isAdmin && CONFIG.mailEnabled ? `<button data-nav='settings' title="E-mailinstellingen" class="btn-icon"><i data-lucide="settings" style="width:16px;height:16px;"></i></button>` : ""}
      <button data-action='logout' title="Afmelden" class="btn-icon"><i data-lucide="log-out" style="width:16px;height:16px;"></i></button>
    </div>
  </div>
  <main class="app-main">${content}</main>`;
}

/* ── Dashboard ── */
function viewDashboard() {
  const openCounts = {}; STAGES.forEach((s) => (openCounts[s.key] = 0));
  S.reports.forEach((r) => { if (r.status !== "Afgehandeld") openCounts[r.type] = (openCounts[r.type] || 0) + 1; });
  const open = S.reports.filter((r) => r.status !== "Afgehandeld").length;
  const hoog = S.reports.filter((r) => r.prioriteit === "Hoog" && r.status !== "Afgehandeld").length;
  const done = S.reports.filter((r) => r.status === "Afgehandeld").length;

  const cards = STAGES.map((s) => {
    const cnt = openCounts[s.key];
    const badge = `<span class="open-badge ${cnt > 0 ? "has-open" : ""}">${cnt} open</span>`;
    return `<button data-newstage='${s.key}' class="stage-card stage-${s.key}">
      <div class="stage-card-head">
        <div class="stage-icon-box"><i data-lucide="${s.icon}" style="width:18px;height:18px;"></i></div>
        ${badge}
      </div>
      <div class="stage-card-body">
        <div class="stap-name">${s.stap}</div>
        <div class="stap-sub">${s.sub}</div>
      </div>
      <div class="stage-card-cta"><i data-lucide="plus" style="width:13px;height:13px;"></i> Melding maken</div>
    </button>`;
  }).join("");

  const rows = filteredReports().map((r) => {
    const stage = stageByKey(r.type);
    const statusCss = STATUS_CSS[r.status] || "status-nieuw";
    const statusIcon = STATUS_ICON[r.status] || "circle";
    return `
    <button data-open='${esc(r.id)}' class="melding-row">
      <div class="melding-dot melding-dot-${r.type}"></div>
      <div class="melding-body">
        <div class="melding-top">
          <span class="melding-id">${esc(r.id)}</span>
          <span class="v-pill stage-badge-${r.type}">${stage.stap}</span>
        </div>
        <div class="melding-title">
          <span style="font-family:monospace;color:var(--muted);">${esc(r.artikelnummer || "")}</span>${r.omschrijving ? ` · ${esc(r.omschrijving)}` : ""}
        </div>
        <div class="melding-meta">
          <span class="melding-meta-item"><i data-lucide="map-pin" style="width:11px;height:11px;"></i>${esc(r.warehouse || "")}</span>
          <span class="melding-meta-item"><i data-lucide="boxes" style="width:11px;height:11px;"></i>${esc(r.aantal || 0)} st.</span>
          ${(r.klant || r.leverancier) ? `<span class="melding-meta-item">${esc(r.klant || r.leverancier)}</span>` : ""}
          <span>${fmtDate(r.aangemaakt)}</span>
        </div>
      </div>
      <div class="melding-right">
        <span class="v-badge ${statusCss}"><i data-lucide="${statusIcon}" style="width:11px;height:11px;"></i> ${r.status}</span>
        <span class="prio-badge ${PRIO_CSS[r.prioriteit] || PRIO_CSS.Normaal}">${esc(r.prioriteit || "Normaal")}</span>
      </div>
    </button>`;
  }).join("");

  const filterSel = (id, value, options, labelFn) =>
    `<select data-filter='${id}' class="filter-select">
      ${options.map((o) => `<option value='${esc(o)}' ${o === value ? "selected" : ""}>${esc(labelFn(o))}</option>`).join("")}
    </select>`;

  const listBlock = filteredReports().length
    ? `<div class="melding-list-wrap">${rows}</div>`
    : `<div class="melding-empty">
         <i data-lucide="clipboard-list" style="width:32px;height:32px;"></i>
         <p>Geen meldingen gevonden</p>
         <small>Pas de filters aan of maak een nieuwe melding via de fasen hierboven.</small>
       </div>`;

  return shell(`
    <section style="margin-bottom:28px;">
      <div class="section-head">
        <h1>Waar is de schade ontstaan?</h1>
        <p>Kies een fase in de goederenstroom om een melding te maken.</p>
      </div>
      <div class="stage-grid">${cards}</div>
    </section>

    <section>
      <div class="section-between">
        <div class="section-left">
          <span class="section-title">Meldingen</span>
          <div class="stats-row">
            <span><span class="stat-dot dot-blue"></span>${open} open</span>
            <span><span class="stat-dot dot-red"></span>${hoog} hoog</span>
            <span><span class="stat-dot dot-green"></span>${done} afgehandeld</span>
          </div>
        </div>
        <button data-action='export' class="btn btn-sm"><i data-lucide="download" style="width:13px;height:13px;"></i> Export CSV</button>
      </div>
      <div class="filter-bar">
        <div class="filter-search-wrap">
          <span class="filter-search-icon"><i data-lucide="search" style="width:15px;height:15px;"></i></span>
          <input data-filter='q' value='${esc(S.filters.q)}' placeholder="Zoek op nr., artikel, klant of leverancier…" class="filter-search" />
        </div>
        <div class="filter-selects">
          ${filterSel("stage", S.filters.stage, ["Alle", ...STAGES.map((s) => s.key)], (v) => v === "Alle" ? "Alle fasen" : stageByKey(v).stap)}
          ${filterSel("status", S.filters.status, ["Alle", ...STATUS_FLOW], (v) => v === "Alle" ? "Alle statussen" : v)}
          ${filterSel("warehouse", S.filters.warehouse, ["Alle", "LAAKDAL", "STORA"], (v) => v === "Alle" ? "Alle magazijnen" : v)}
        </div>
      </div>
      ${listBlock}
    </section>`);
}

function filteredReports() {
  const f = S.filters; const needle = f.q.trim().toLowerCase();
  return S.reports
    .filter((r) => f.status === "Alle" || r.status === f.status)
    .filter((r) => f.stage === "Alle" || r.type === f.stage)
    .filter((r) => f.warehouse === "Alle" || r.warehouse === f.warehouse)
    .filter((r) => !needle || [r.id, r.artikelnummer, r.omschrijving, r.klant, r.leverancier, r.verkooporder, r.inkooporder].filter(Boolean).join(" ").toLowerCase().includes(needle));
}

/* ── Formulier ── */
function fieldHtml(f, value) {
  const req = f.required ? `<span class="v-req"> *</span>` : "";
  let control;
  if (f.type === "select") {
    control = `<select data-field='${f.name}' class="v-fg"><select><option value="">— Kies —</option>${f.options.map((o) => `<option ${o === value ? "selected" : ""}>${esc(o)}</option>`).join("")}</select>`;
    // Use proper select
    control = `<select data-field='${f.name}'><option value="">— Kies —</option>${f.options.map((o) => `<option ${o === value ? "selected" : ""}>${esc(o)}</option>`).join("")}</select>`;
  } else if (f.type === "textarea") {
    control = `<textarea data-field='${f.name}' rows="3" placeholder="${esc(f.placeholder || "")}">${esc(value || "")}</textarea>`;
  } else {
    const t = f.type === "number" ? "number" : f.type === "date" ? "date" : "text";
    control = `<input data-field='${f.name}' type="${t}" value='${esc(value || "")}' placeholder="${esc(f.placeholder || "")}" ${f.mono ? 'style="font-family:monospace;"' : ""} />`;
  }
  return `<div class="v-fg ${f.full ? "form-full" : ""}"><label>${esc(f.label)}${req}</label>${control}</div>`;
}

function viewForm(stageKey) {
  const stage = stageByKey(stageKey);
  const fields = [...COMMON_HEAD, ...stage.fields, ...COMMON_TAIL];
  const id = nextId();
  const defaults = {}; fields.forEach((f) => { if (f.type === "date") defaults[f.name] = todayISO(); if (f.name === "prioriteit") defaults[f.name] = "Normaal"; });
  return shell(`
    <button data-nav='dashboard' class="back-btn"><i data-lucide="chevron-left" style="width:15px;height:15px;"></i> Terug</button>
    <div class="form-card">
      <div class="form-card-head">
        <div class="stage-icon-box stage-${stageKey}" style="flex-shrink:0;"><i data-lucide="${stage.icon}" style="width:18px;height:18px;"></i></div>
        <div>
          <div class="form-card-label">Nieuwe schademelding</div>
          <div class="form-card-title">${stage.label}</div>
        </div>
        <span class="form-id">${id}</span>
      </div>
      <div id="formfields" data-stage='${stageKey}' data-id='${id}' class="form-card-body">
        ${fields.map((f) => fieldHtml(f, defaults[f.name])).join("")}
        <div class="v-fg form-full">
          <label>Foto's</label>
          <div class="foto-btns">
            <label class="foto-btn-label">
              <i data-lucide="camera" style="width:15px;height:15px;"></i> Foto maken
              <input type="file" accept="image/*" capture="environment" class="hidden" data-addfiles />
            </label>
            <label class="foto-btn-label">
              <i data-lucide="image-plus" style="width:15px;height:15px;"></i> Uit galerij
              <input type="file" accept="image/*" multiple class="hidden" data-addfiles />
            </label>
          </div>
          <div id="formthumbs" class="thumb-grid" style="margin-top:10px;"></div>
        </div>
      </div>
      <div class="form-card-foot">
        <button data-nav='dashboard' class="btn btn-secondary">Annuleren</button>
        <button data-action='save' class="btn btn-primary">Melding opslaan</button>
      </div>
    </div>`);
}

function renderFormThumbs() {
  const box = $("#formthumbs"); if (!box) return;
  box.innerHTML = thumbsHtml(formPhotos, true);
  icons();
  hydrateThumbs(box, formPhotos);
}

/* ── Detail ── */
function viewDetail(id) {
  const r = S.reports.find((x) => x.id === id);
  if (!r) return shell(`<p style="color:var(--muted)">Melding niet gevonden.</p>`);
  const stage = stageByKey(r.type);
  const statusCss = STATUS_CSS[r.status] || "status-nieuw";
  const statusIcon = STATUS_ICON[r.status] || "circle";
  const detailFields = [...COMMON_HEAD, ...stage.fields, ...COMMON_TAIL];
  const defs = detailFields.map((f) => {
    const v = r[f.name]; if (v === undefined || v === "" || v === null) return "";
    const val = f.name === "aantal" ? `${v} stuks` : esc(v);
    return `<div class="${f.full ? "full" : ""}"><dt>${esc(f.label)}</dt><dd ${f.mono ? 'class="mono"' : ""}>${val}</dd></div>`;
  }).join("");
  const statusBtns = STATUS_FLOW.map((s) => {
    const active = r.status === s;
    return `<button data-status='${s}' class="status-btn ${active ? "active" : ""}">${s}</button>`;
  }).join("");
  const fotoBlock = (r.fotos && r.fotos.length)
    ? `<div class="full"><dt>Foto's</dt><dd id="detailthumbs" class="thumb-grid"></dd></div>` : "";

  return shell(`
    <button data-nav='dashboard' class="back-btn"><i data-lucide="chevron-left" style="width:15px;height:15px;"></i> Terug naar overzicht</button>
    <div class="detail-card">
      <div class="detail-head">
        <div class="stage-icon-box stage-${r.type}"><i data-lucide="${stage.icon}" style="width:18px;height:18px;"></i></div>
        <div style="min-width:0;">
          <div style="display:flex;align-items:center;gap:6px;margin-bottom:3px;">
            <span style="font-family:monospace;font-size:11px;color:var(--muted);">${esc(r.id)}</span>
            <span class="v-pill stage-badge-${r.type}">${stage.stap}</span>
          </div>
          <div style="font-size:15px;font-weight:700;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">${esc(r.omschrijving || r.artikelnummer || "")}</div>
        </div>
        <span class="v-badge ${statusCss}" style="margin-left:auto;flex-shrink:0;"><i data-lucide="${statusIcon}" style="width:12px;height:12px;"></i> ${r.status}</span>
      </div>
      <div class="detail-status-bar">
        <span style="font-size:12px;font-weight:600;color:var(--muted);">Status:</span>
        ${statusBtns}
      </div>
      <dl class="detail-dl">
        ${defs}
        <div><dt>Aangemaakt</dt><dd>${fmtDate(r.aangemaakt)}</dd></div>
        ${fotoBlock}
      </dl>
      <div class="detail-foot">
        ${S.isAdmin
          ? `<button data-action='delete' data-id='${esc(r.id)}' class="btn btn-ghost btn-sm" style="color:var(--red);"><i data-lucide="trash-2" style="width:14px;height:14px;"></i> Verwijderen</button>`
          : `<span style="font-size:12px;color:var(--muted);display:flex;align-items:center;gap:5px;"><i data-lucide="lock" style="width:13px;height:13px;"></i> Alleen beheerders kunnen verwijderen</span>`}
      </div>
    </div>

    <div class="gesprek-card">
      <div class="gesprek-head">
        <i data-lucide="messages-square" style="width:16px;height:16px;color:var(--muted);"></i>
        Gesprek
      </div>
      <div id="gesprek" class="gesprek-body"></div>
      <div class="gesprek-foot">
        <textarea id="comment-input" rows="2" placeholder="Schrijf een bericht… (Ctrl/⌘ + Enter om te versturen)" class="comment-input"></textarea>
        <div class="comment-actions">
          <button data-comment-send class="btn btn-primary btn-sm"><i data-lucide="send" style="width:13px;height:13px;"></i> Verstuur</button>
        </div>
      </div>
    </div>`);
}

/* ── Gesprek ── */
function commentListHtml(comments) {
  if (!comments.length) return `<p style="color:var(--muted);font-size:13px;padding:8px 0;">Nog geen berichten. Start het gesprek hieronder.</p>`;
  const me = (S.account && (S.account.name || S.account.username)) || "";
  return comments.map((c) => {
    const mine = c.auteur === me;
    const initials = (c.auteur || "?").trim().split(/\s+/).map((w) => w[0]).slice(0, 2).join("").toUpperCase();
    return `<div class="comment-bubble">
      <span class="comment-avatar ${mine ? "mine" : "other"}">${esc(initials || "?")}</span>
      <div style="min-width:0;flex:1;">
        <div class="comment-meta">
          <span class="comment-auteur">${esc(c.auteur || "Onbekend")}</span>
          <span class="comment-datum">${fmtDateTime(c.datum)}</span>
        </div>
        <div class="comment-text">${esc(c.bericht || "")}</div>
      </div>
    </div>`;
  }).join("");
}
async function loadAndRenderComments(meldingId) {
  const box = $("#gesprek"); if (!box) return;
  box.innerHTML = `<div class="v-loading"><div class="v-spinner"></div></div>`;
  try {
    const comments = await fetchComments(meldingId);
    box.innerHTML = commentListHtml(comments); icons();
    box.scrollTop = box.scrollHeight;
  } catch (e) { console.error(e); box.innerHTML = `<p style="color:var(--muted);font-size:13px;">Berichten laden mislukt.</p>`; }
}
async function postComment(meldingId) {
  const ta = $("#comment-input"); if (!ta) return;
  const text = ta.value.trim(); if (!text) return;
  const btn = $("[data-comment-send]");
  const auteur = (S.account && (S.account.name || S.account.username)) || "Onbekend";
  if (btn) btn.disabled = true;
  ta.disabled = true;
  try { await addComment(meldingId, auteur, text); ta.value = ""; await loadAndRenderComments(meldingId); }
  catch (e) { console.error(e); toast("Bericht versturen mislukt", "err"); }
  finally { if (btn) btn.disabled = false; ta.disabled = false; ta.focus(); }
}

/* ── Instellingen ── */
function viewSettings() {
  if (!S.isAdmin) return shell(`<p style="color:var(--muted);">Geen toegang. Deze pagina is enkel voor beheerders.</p>`);
  const s = S.settings || { enabled: true, recipients: [] };
  const FASEN = [
    { key:"inkomend",  label:"Ontvangst" },
    { key:"voorraad",  label:"Voorraad" },
    { key:"uitgaand",  label:"Verzending" },
    { key:"levering",  label:"Levering" },
  ];
  const ALL_KEYS = FASEN.map((f) => f.key);
  const rows = (s.recipients || []).map((r, i) => {
    const fasen = Array.isArray(r.fasen) && r.fasen.length ? r.fasen : ALL_KEYS;
    const faseBoxes = FASEN.map((f) => {
      const chk = fasen.includes(f.key) ? "checked" : "";
      return `<label class="recip-fasen" style="display:inline-flex;align-items:center;gap:5px;font-size:12px;color:var(--muted);cursor:pointer;">
        <input type="checkbox" data-recip-fase='${i}' data-fase='${f.key}' ${chk} style="width:13px;height:13px;" />
        <span class="v-pill stage-badge-${f.key}">${f.label}</span>
      </label>`;
    }).join("");
    return `
    <div class="recip-row">
      <div class="recip-row-inputs">
        <input data-recip-email='${i}' value='${esc(r.email || "")}' placeholder="naam@verpa.be" type="email" style="flex:1;min-width:0;" class="v-fg" />
        <input data-recip-naam='${i}' value='${esc(r.naam || "")}' placeholder="Naam (optioneel)" style="width:140px;" class="v-fg" />
        <button data-recip-remove='${i}' title="Verwijderen" class="btn-icon" style="color:var(--red);">
          <i data-lucide="trash-2" style="width:15px;height:15px;"></i></button>
      </div>
      <div class="recip-fasen" style="display:flex;flex-wrap:wrap;gap:8px 12px;align-items:center;">
        <span style="font-size:12px;font-weight:600;color:var(--muted);">Ontvangt:</span>
        ${faseBoxes}
      </div>
    </div>`;
  }).join("");

  return shell(`
    <button data-nav='dashboard' class="back-btn"><i data-lucide="chevron-left" style="width:15px;height:15px;"></i> Terug</button>
    <div class="settings-card">
      <div class="settings-head">
        <div style="font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:.05em;color:var(--muted);">Beheer</div>
        <div style="font-size:15px;font-weight:700;margin:2px 0;">E-mailinstellingen</div>
        <p style="font-size:13px;color:var(--muted);margin-top:4px;">Stel per ontvanger in voor welke fasen ze een e-mail ontvangen bij een nieuwe melding.</p>
      </div>
      <div class="settings-body">
        <label class="toggle-label">
          <input type="checkbox" data-notify-toggle ${s.enabled ? "checked" : ""} style="width:15px;height:15px;" />
          E-mailmeldingen inschakelen
        </label>
        <div>
          <div style="font-size:13px;font-weight:700;color:var(--text);margin-bottom:10px;">Ontvangers</div>
          <div id="recip-list" style="display:flex;flex-direction:column;gap:8px;">
            ${rows || `<p style="font-size:13px;color:var(--muted);">Nog geen ontvangers toegevoegd.</p>`}
          </div>
          <button data-recip-add class="btn btn-sm" style="margin-top:12px;">
            <i data-lucide="plus" style="width:14px;height:14px;"></i> Ontvanger toevoegen</button>
        </div>
      </div>
      <div class="settings-foot">
        <button data-action='settings-save' class="btn btn-primary">Opslaan</button>
      </div>
    </div>`);
}
function syncRecipientsFromDom() {
  if (!S.settings) S.settings = { enabled: true, recipients: [] };
  const list = [];
  document.querySelectorAll("[data-recip-email]").forEach((el) => {
    const i = +el.getAttribute("data-recip-email");
    const naam = document.querySelector(`[data-recip-naam='${i}']`);
    const fasen = Array.from(document.querySelectorAll(`[data-recip-fase='${i}']:checked`)).map((cb) => cb.getAttribute("data-fase"));
    list.push({ email: el.value.trim(), naam: naam ? naam.value.trim() : "", fasen, actief: true });
  });
  S.settings.recipients = list;
  const t = document.querySelector("[data-notify-toggle]");
  if (t) S.settings.enabled = t.checked;
}
function wireSettings() {
  const ALL_KEYS = ["inkomend","voorraad","uitgaand","levering"];
  const add = document.querySelector("[data-recip-add]");
  if (add) add.onclick = () => { syncRecipientsFromDom(); S.settings.recipients.push({ email: "", naam: "", fasen: ALL_KEYS, actief: true }); render(); };
  document.querySelectorAll("[data-recip-remove]").forEach((b) => b.onclick = () => {
    syncRecipientsFromDom(); S.settings.recipients.splice(+b.getAttribute("data-recip-remove"), 1); render();
  });
}

/* ── Thumbnails + viewer ── */
function thumbsHtml(photos, removable) {
  if (!photos || !photos.length) return "";
  return photos.map((p, i) => `
    <div class="thumb-item">
      <button data-thumb='${i}' class="thumb-open"><i data-lucide="loader-2" style="width:18px;height:18px;" class="animate-spin"></i></button>
      <div class="thumb-hover-overlay"><i data-lucide="zoom-in" style="width:11px;height:11px;"></i> Openen</div>
      ${removable ? `<button data-removefoto='${i}' class="thumb-remove"><i data-lucide="x" style="width:11px;height:11px;"></i></button>` : ""}
    </div>`).join("");
}
async function hydrateThumbs(container, photos) {
  const btns = container.querySelectorAll("[data-thumb]");
  btns.forEach((b) => b.onclick = () => openThumb(b));
  container.querySelectorAll("[data-removefoto]").forEach((b) => b.onclick = () => { formPhotos.splice(+b.getAttribute("data-removefoto"), 1); renderFormThumbs(); });
  photos.forEach(async (p, i) => {
    try {
      const url = await getFotoBlobUrl(p);
      const btn = btns[i]; if (!btn) return;
      btn.innerHTML = `<img src="${url}" alt="${esc(p.name)}" style="width:100%;height:100%;object-fit:cover;" />`;
    } catch {
      const btn = btns[i]; if (btn) { btn.innerHTML = `<i data-lucide="images" style="width:20px;height:20px;"></i>`; icons(); }
    }
  });
}
function openThumb(btn) {
  const idx = +btn.getAttribute("data-thumb");
  const photos = S.view.name === "form" ? formPhotos : (S.reports.find((r) => r.id === S.view.id)?.fotos || []);
  if (photos.length) openViewer(photos, idx);
}
function openViewer(photos, startIndex) {
  let index = startIndex;
  const root = $("#overlay-root");
  root.innerHTML = `
    <div id="viewer" class="foto-viewer">
      <div class="viewer-topbar">
        <div>
          <div id="v-name" class="viewer-name"></div>
          ${photos.length > 1 ? `<div id="v-count" class="viewer-count"></div>` : ""}
        </div>
        <div style="display:flex;align-items:center;gap:4px;">
          <a id="v-download" href="#" download class="viewer-btn-icon"><i data-lucide="download" style="width:18px;height:18px;"></i></a>
          <button data-vclose class="viewer-btn-icon"><i data-lucide="x" style="width:18px;height:18px;"></i></button>
        </div>
      </div>
      <div id="v-stage" class="viewer-stage">
        <i data-lucide="loader-2" style="width:28px;height:28px;color:rgba(255,255,255,.6);" class="animate-spin"></i>
        ${photos.length > 1 ? `
          <button data-vprev class="viewer-nav viewer-prev"><i data-lucide="chevron-left" style="width:22px;height:22px;"></i></button>
          <button data-vnext class="viewer-nav viewer-next"><i data-lucide="chevron-right" style="width:22px;height:22px;"></i></button>` : ""}
      </div>
    </div>`;
  icons();
  const stage = $("#v-stage"), nameEl = $("#v-name"), countEl = $("#v-count"), dl = $("#v-download");
  async function show() {
    nameEl.textContent = photos[index].name || "";
    if (countEl) countEl.textContent = `${index + 1} / ${photos.length}`;
    [...stage.querySelectorAll("img")].forEach((n) => n.remove());
    try {
      const url = await getFotoBlobUrl(photos[index]);
      const img = document.createElement("img");
      img.src = url; img.alt = photos[index].name || "";
      img.className = "viewer-img"; img.onclick = (e) => e.stopPropagation();
      stage.insertBefore(img, stage.firstChild);
      dl.href = url; dl.setAttribute("download", photos[index].name || "foto");
    } catch { /* laat spinner staan */ }
  }
  const goPh = (d) => { index = (index + d + photos.length) % photos.length; show(); };
  function close() { root.innerHTML = ""; document.removeEventListener("keydown", onKey); }
  function onKey(e) { if (e.key === "Escape") close(); else if (e.key === "ArrowRight" && photos.length > 1) goPh(1); else if (e.key === "ArrowLeft" && photos.length > 1) goPh(-1); }
  document.addEventListener("keydown", onKey);
  const viewer = $("#viewer");
  viewer.addEventListener("click", (e) => { if (e.target.closest("#v-stage") && !e.target.closest("button") && !e.target.closest("a") && e.target.tagName !== "IMG") close(); });
  $("[data-vclose]").onclick = close;
  const prev = $("[data-vprev]"), next = $("[data-vnext]");
  if (prev) prev.onclick = (e) => { e.stopPropagation(); goPh(-1); };
  if (next) next.onclick = (e) => { e.stopPropagation(); goPh(1); };
  let startX = null;
  viewer.addEventListener("touchstart", (e) => (startX = e.touches[0].clientX), { passive: true });
  viewer.addEventListener("touchend", (e) => { if (startX == null) return; const dx = e.changedTouches[0].clientX - startX; if (Math.abs(dx) > 50 && photos.length > 1) goPh(dx < 0 ? 1 : -1); startX = null; });
  show();
}

/* ── CSV export ── */
function exportCSV() {
  const cols = ["id", "type", "status", "prioriteit", "aangemaakt", "warehouse", "artikelnummer", "omschrijving", "aantal", "melder"];
  const q = (v) => `"${String(v == null ? "" : v).replace(/"/g, '""')}"`;
  const rows = filteredReports().map((r) => {
    const stage = stageByKey(r.type);
    const details = stage.fields.map((f) => (r[f.name] ? `${f.label}: ${r[f.name]}` : "")).filter(Boolean).join(" | ");
    return [...cols.map((c) => q(c === "type" ? stage.label : r[c])), q(details)].join(",");
  });
  const csv = "\uFEFF" + [...cols, "details"].join(",") + "\n" + rows.join("\n");
  const url = URL.createObjectURL(new Blob([csv], { type: "text/csv;charset=utf-8" }));
  const a = document.createElement("a"); a.href = url; a.download = `schademeldingen_${todayISO()}.csv`; a.click();
  URL.revokeObjectURL(url);
}

/* ── Toast ── */
function toast(msg, type = "ok") {
  let el = document.getElementById("v-toast");
  if (!el) { el = document.createElement("div"); el.id = "v-toast"; document.body.appendChild(el); }
  el.textContent = msg;
  el.className = `show ${type}`;
  clearTimeout(el._t);
  el._t = setTimeout(() => { el.className = el.className.replace("show", "").trim(); }, 2600);
}

/* ── Schermen ── */
function loadingScreen() {
  return `<div class="screen-center"><div class="v-loading"><div class="v-spinner"></div><span>Laden…</span></div></div>`;
}
function loginScreen() {
  return `<div class="screen-center">
    <div class="screen-card">
      <div class="screen-icon screen-icon-dark"><i data-lucide="alert-triangle" style="width:24px;height:24px;"></i></div>
      <h1>Schademeldingen</h1>
      <p>Verpa Benelux · Magazijn<br>Meld aan met je Microsoft-account om verder te gaan.</p>
      <button data-action='login' class="btn btn-primary btn-lg" style="width:100%;">Aanmelden met Microsoft</button>
    </div>
  </div>`;
}
function errorScreen(msg) {
  return `<div class="screen-center">
    <div class="screen-card">
      <div class="screen-icon screen-icon-err"><i data-lucide="alert-triangle" style="width:24px;height:24px;"></i></div>
      <h1>Er ging iets mis</h1>
      <p>${esc(msg)}</p>
      <button data-action='reload' class="btn btn-primary">Opnieuw proberen</button>
    </div>
  </div>`;
}
function setupScreen() {
  return `<div class="screen-center">
    <div class="screen-card" style="text-align:left;">
      <h1>Nog instellen</h1>
      <p>Vul bovenaan in <code>schade-app.js</code> de <code>CONFIG</code> in: clientId, tenantId, siteHostname en sitePath uit je Azure-app en SharePoint-site.</p>
    </div>
  </div>`;
}

/* ── Render + events ── */
function render() {
  const app = $("#app");
  if (S.loading) { app.innerHTML = loadingScreen(); icons(); return; }
  if (S.error) { app.innerHTML = errorScreen(S.error); icons(); wireEvents(); return; }
  if (S.view.name === "form")     app.innerHTML = viewForm(S.view.stageKey);
  else if (S.view.name === "detail")   app.innerHTML = viewDetail(S.view.id);
  else if (S.view.name === "settings") app.innerHTML = viewSettings();
  else app.innerHTML = viewDashboard();
  icons();
  wireEvents();
  if (S.view.name === "settings") wireSettings();
  if (S.view.name === "form") { formPhotos = formPhotos || []; renderFormThumbs(); }
  if (S.view.name === "detail") {
    const r = S.reports.find((x) => x.id === S.view.id);
    const box = $("#detailthumbs");
    if (r && r.fotos && r.fotos.length && box) { box.innerHTML = thumbsHtml(r.fotos, false); icons(); hydrateThumbs(box, r.fotos); }
    if (r) {
      loadAndRenderComments(r.id);
      const sendBtn = $("[data-comment-send]"); if (sendBtn) sendBtn.onclick = () => postComment(r.id);
      const ta = $("#comment-input"); if (ta) ta.onkeydown = (e) => { if ((e.ctrlKey || e.metaKey) && e.key === "Enter") { e.preventDefault(); postComment(r.id); } };
    }
  }
}

function refreshList() {
  if (S.view.name !== "dashboard") return;
  const hadFocus = document.activeElement && document.activeElement.getAttribute && document.activeElement.getAttribute("data-filter") === "q";
  const caret = hadFocus ? document.activeElement.selectionStart : null;
  render();
  if (hadFocus) { const q = $("[data-filter='q']"); if (q) { q.focus(); if (caret != null) q.setSelectionRange(caret, caret); } }
}

function wireEvents() {
  document.querySelectorAll("[data-nav]").forEach((b) => b.onclick = () => go({ name: b.getAttribute("data-nav") }));
  document.querySelectorAll("[data-newstage]").forEach((b) => b.onclick = () => { formPhotos = []; go({ name: "form", stageKey: b.getAttribute("data-newstage") }); });
  document.querySelectorAll("[data-open]").forEach((b) => b.onclick = () => go({ name: "detail", id: b.getAttribute("data-open") }));
  document.querySelectorAll("[data-status]").forEach((b) => b.onclick = () => changeStatus(S.view.id, b.getAttribute("data-status")));
  document.querySelectorAll("[data-thumb]").forEach((b) => b.onclick = () => openThumb(b));
  document.querySelectorAll("[data-removefoto]").forEach((b) => b.onclick = () => { formPhotos.splice(+b.getAttribute("data-removefoto"), 1); renderFormThumbs(); });
  document.querySelectorAll("[data-addfiles]").forEach((inp) => inp.onchange = (e) => { addFiles(e.target.files); e.target.value = ""; });
  document.querySelectorAll("[data-filter]").forEach((el) => {
    const key = el.getAttribute("data-filter");
    if (key === "q") el.oninput = () => { S.filters.q = el.value; refreshList(); };
    else el.onchange = () => { S.filters[key] = el.value; refreshList(); };
  });
  const actions = {
    login, logout, reload: () => location.reload(),
    export: exportCSV,
    save: saveReport,
    delete: (b) => confirmDelete(b.getAttribute("data-id")),
    "settings-save": saveSettingsFromUI,
  };
  document.querySelectorAll("[data-action]").forEach((b) => b.onclick = () => actions[b.getAttribute("data-action")] && actions[b.getAttribute("data-action")](b));
}

async function saveSettingsFromUI() {
  syncRecipientsFromDom();
  S.settings.recipients = (S.settings.recipients || []).filter((r) => r.email);
  const btn = document.querySelector("[data-action='settings-save']");
  if (btn) { btn.disabled = true; btn.textContent = "Opslaan…"; }
  try { await saveSettings(S.settings); toast("Instellingen opgeslagen"); go({ name: "dashboard" }); }
  catch (e) { console.error(e); if (btn) { btn.disabled = false; btn.textContent = "Opslaan"; } toast("Opslaan mislukt", "err"); }
}

function addFiles(fileList) {
  const items = Array.from(fileList || []).map((file) => ({ name: file.name, url: URL.createObjectURL(file), file }));
  formPhotos.push(...items);
  renderFormThumbs();
}

async function saveReport() {
  const box = $("#formfields"); if (!box) return;
  const stageKey = box.getAttribute("data-stage"); const id = box.getAttribute("data-id");
  const stage = stageByKey(stageKey);
  const fields = [...COMMON_HEAD, ...stage.fields, ...COMMON_TAIL];
  const data = {}; box.querySelectorAll("[data-field]").forEach((el) => { data[el.getAttribute("data-field")] = el.value; });
  // validatie
  let bad = null;
  box.querySelectorAll("[data-field]").forEach((el) => el.classList.remove("v-err"));
  fields.forEach((f) => { if (f.required && !String(data[f.name] || "").trim()) { const el = box.querySelector(`[data-field='${f.name}']`); if (el) el.classList.add("v-err"); if (!bad) bad = el; } });
  if (bad) { bad.scrollIntoView({ behavior: "smooth", block: "center" }); return; }
  const btn = $("[data-action='save']");
  btn.disabled = true; btn.textContent = formPhotos.length ? "Foto's uploaden…" : "Opslaan…";
  try {
    let fotoRefs = [];
    if (formPhotos.length) fotoRefs = await uploadFotos(id, formPhotos.map((p) => p.file));
    const rec = await createReport({ ...data, id, type: stageKey, status: "Nieuw", aangemaakt: todayISO(), aantal: Number(data.aantal) || 0, fotos: fotoRefs });
    S.reports.unshift(rec);
    formPhotos = [];
    notifyNewCase(rec).catch((e) => console.warn("E-mailmelding niet verstuurd:", e));
    toast(`Melding ${rec.id} opgeslagen`);
    go({ name: "detail", id: rec.id });
  } catch (e) {
    btn.disabled = false; btn.textContent = "Melding opslaan";
    toast("Opslaan mislukt — probeer opnieuw", "err");
    console.error(e);
  }
}

async function changeStatus(id, status) {
  const r = S.reports.find((x) => x.id === id); if (!r) return;
  const prev = r.status; r.status = status; render();
  try { await updateReport(r.spId, { status }); } catch (e) { r.status = prev; render(); toast("Status bijwerken mislukt", "err"); }
}

function confirmDelete(id) {
  if (!S.isAdmin) { toast("Alleen beheerders kunnen verwijderen", "err"); return; }
  const root = $("#overlay-root");
  root.innerHTML = `<div class="v-overlay">
    <div class="v-modal v-modal-sm">
      <div class="v-modal-head"><h3>Melding verwijderen?</h3></div>
      <div class="v-modal-body">
        <p style="font-size:13px;color:var(--muted);">${esc(id)} wordt definitief uit SharePoint verwijderd.</p>
      </div>
      <div class="v-modal-foot">
        <button data-cancel class="btn btn-secondary">Annuleren</button>
        <button data-confirm class="btn btn-danger">Verwijderen</button>
      </div>
    </div>
  </div>`;
  $("[data-cancel]").onclick = () => (root.innerHTML = "");
  $("[data-confirm]").onclick = async () => {
    root.innerHTML = "";
    const r = S.reports.find((x) => x.id === id); if (!r) return;
    try { await deleteReport(r.spId); S.reports = S.reports.filter((x) => x.id !== id); toast("Melding verwijderd"); go({ name: "dashboard" }); }
    catch { toast("Verwijderen mislukt", "err"); }
  };
}

/* ==================================================================== *
 *  INIT
 * ==================================================================== */
(async function init() {
  render();
  if (!CONFIG.clientId || !CONFIG.tenantId) { S.loading = false; $("#app").innerHTML = setupScreen(); icons(); return; }
  try {
    S.account = await initAuth();
    if (!S.account) { S.loading = false; $("#app").innerHTML = loginScreen(); icons(); wireEvents(); return; }
    S.isAdmin = isAdmin();
    S.reports = await fetchReports();
    S.settings = await loadSettings().catch(() => ({ spId: null, enabled: false, recipients: [] }));
    const caseId = new URLSearchParams(location.search).get("case");
    if (caseId && S.reports.find((r) => r.id === caseId)) S.view = { name: "detail", id: caseId };
    S.loading = false;
    render();
  } catch (e) {
    console.error(e);
    S.loading = false; S.error = e.message || String(e); render();
  }
})();
