/* ============================================================================
 *  SCHADEMELDINGEN — zelfstandige app voor Cloudflare Pages
 *  MSAL (delegated) + Microsoft Graph -> eigen SharePoint-lijsten + fotobibliotheek
 *  Refactored: externe JS met Verpa-UI
 * ==========================================================================
 *
 *  Deze app staat volledig op zichzelf (eigen Azure-app-registratie, eigen
 *  lijsten). Zie het meegeleverde STAPPENPLAN.md voor de volledige installatie.
 *  Kort:
 *   1) Maak een eigen Azure-app-registratie "Schademeldingen" (SPA) en vul
 *      clientId + tenantId hieronder in.
 *   2) Zet redirectUri gelijk aan je Cloudflare Pages-adres en voeg datzelfde
 *      adres toe als SPA-redirect in die app-registratie.
 *   3) Sleep dit bestand als index.html naar een Cloudflare Pages-project.
 *  De lijsten (Schademeldingen, SchadeInstellingen, SchadeReacties) en de
 *  fotomap worden automatisch aangemaakt bij het eerste gebruik.
 *
 *  Nog nodig in de app-registratie:
 *   - Delegated Graph-permissies: User.Read, Sites.ReadWrite.All, Files.ReadWrite.All
 *   - Een App Role "Admin" definiëren en toewijzen aan de beheerders.
 * ============================================================================ */

/* ─────────────────────────────────────────────────────────────────────────── *
 *  CONFIG
 * ─────────────────────────────────────────────────────────────────────────── */
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

/* ─────────────────────────────────────────────────────────────────────────── *
 *  FASEN + VELDEN (single source of truth)
 * ─────────────────────────────────────────────────────────────────────────── */
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
    ]
  },
  { key: "voorraad", stap: "Voorraad", label: "Schade aan voorraad", sub: "Tijdens opslag of orderverwerking", icon: "warehouse",
    fields: [
      { name: "locatie", label: "Locatie / bin", type: "text", placeholder: "bv. A-12-03", mono: true },
      { name: "oorzaak", label: "Oorzaak", type: "select", options: ["Handling (val)", "Heftruck / transpallet", "THT verlopen", "Waterschade", "Mispick-schade", "Onbekend"] },
      { name: "ontdektBij", label: "Ontdekt bij", type: "select", options: ["Orderpicking", "Cyclustelling", "Routine-inspectie", "Bijvullen"] },
      { name: "gekoppeldeOrder", label: "Gekoppelde order (optioneel)", type: "text", placeholder: "Verkooporder tijdens picking", mono: true },
      { name: "afhandeling", label: "Afhandeling", type: "select", options: ["Afgeschreven", "Afgeprijsd", "Hersteld", "In quarantaine"] },
    ]
  },
  { key: "uitgaand", stap: "Verzending", label: "Uitgaande goederen", sub: "Schade vóór of bij verzending", icon: "package-check",
    fields: [
      { name: "klant", label: "Klant", type: "text", placeholder: "Naam klant" },
      { name: "verkooporder", label: "Verkooporder-nr.", type: "text", placeholder: "BC verkooporder", mono: true },
      { name: "ontdektTijdens", label: "Ontdekt tijdens", type: "select", options: ["Verpakken", "Laden", "Eindcontrole"] },
      { name: "vervoerder", label: "Vervoerder", type: "text", placeholder: "Transporteur" },
      { name: "actie", label: "Actie", type: "select", options: ["Opnieuw picken", "Vervangen", "Order gesplitst", "Verzending uitgesteld"] },
    ]
  },
  { key: "levering", stap: "Levering", label: "Schage bij levering", sub: "Na aflevering gemeld", icon: "truck",
    fields: [
      { name: "klant", label: "Klant", type: "text", placeholder: "Naam klant" },
      { name: "verkooporder", label: "Verkooporder-nr.", type: "text", placeholder: "BC verkooporder", mono: true },
      { name: "leveringsdatum", label: "Leveringsdatum", type: "date" },
      { name: "vervoerder", label: "Vervoerder", type: "text", placeholder: "Transporteur" },
      { name: "gemeldDoor", label: "Gemeld door", type: "select", options: ["Klant", "Chauffeur", "Vertegenwoordiger"] },
      { name: "claimVervoerder", label: "Claim bij vervoerder", type: "select", options: ["Ja", "Nee", "In behandeling"] },
      { name: "oplossing", label: "Oplossing", type: "select", options: ["Creditnota", "Herlevering", "Retour", "Geen actie"] },
    ]
  },
];

const stageByKey = (k) => STAGES.find((s) => s.key === k);
const KEY_TO_FASE = { inkomend: "Ontvangst", voorraad: "Voorraad", uitgaand: "Verzending", levering: "Levering" };
const FASE_TO_KEY = Object.fromEntries(Object.entries(KEY_TO_FASE).map(([k, v]) => [v, k]));

const STATUS_FLOW = ["Nieuw", "In behandeling", "Afgehandeld"];
const STATUS_META = {
  "Nieuw":          { badgeClass: "v-badge-info", icon: "circle" },
  "In behandeling": { badgeClass: "v-badge-pending", icon: "clock" },
  "Afgehandeld":    { badgeClass: "v-badge-approved", icon: "check-circle-2" },
};

/* ─────────────────────────────────────────────────────────────────────────── *
 *  STATE
 * ─────────────────────────────────────────────────────────────────────────── */
const S = {
  loading:  true,
  error:    null,
  account:  null,
  isAdmin:  false,
  reports:  [],
  settings: { spId: null, enabled: false, recipients: [] },
  view:     { name: "dashboard" },
  filter:   { q: "", fase: "", status: "" },
};

let formPhotos = [];
let photoViewer = null;

/* ─────────────────────────────────────────────────────────────────────────── *
 *  HELPERS
 * ─────────────────────────────────────────────────────────────────────────── */
const $ = (sel) => document.querySelector(sel);
const $$ = (sel) => document.querySelectorAll(sel);
const esc = (s) => String(s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]);
const todayISO = () => new Date().toISOString().split("T")[0];

function toast(msg, type = "info") {
  let toast = $("#v-toast");
  if (!toast) {
    toast = document.createElement("div");
    toast.id = "v-toast";
    document.body.appendChild(toast);
  }
  toast.textContent = msg;
  toast.className = `show ${type}`;
  setTimeout(() => toast.classList.remove("show"), 3000);
}

function icons() {
  $$("[data-lucide]").forEach((el) => {
    const name = el.getAttribute("data-lucide");
    el.innerHTML = lucide.icons[name]?.toSvg() || "";
  });
}

/* ─────────────────────────────────────────────────────────────────────────── *
 *  NAVIGATION
 * ─────────────────────────────────────────────────────────────────────────── */
function go(view) {
  S.view = view;
  render();
}

/* ─────────────────────────────────────────────────────────────────────────── *
 *  RENDER — MAIN
 * ─────────────────────────────────────────────────────────────────────────── */
function render() {
  const app = $("#app");
  if (!app) return;

  if (S.loading) {
    app.innerHTML = `<div class="loading-screen"><div class="loading-box"><div class="loading-spinner"></div><p>Laden…</p></div></div>`;
    return;
  }

  if (S.error) {
    app.innerHTML = `<div class="v-alert v-alert-err" style="margin: 20px;"><span>Fout: ${esc(S.error)}</span></div>`;
    return;
  }

  if (!S.account) {
    app.innerHTML = loginScreen();
    wireEvents();
    icons();
    return;
  }

  // Topbar + content
  app.innerHTML = `
    <div class="app-shell">
      ${renderTopbar()}
      <div class="app-content">
        ${renderSidebar()}
        <div class="app-main">${renderView()}</div>
      </div>
    </div>
  `;
  wireEvents();
  icons();
}

/* ─────────────────────────────────────────────────────────────────────────── *
 *  TOPBAR
 * ─────────────────────────────────────────────────────────────────────────── */
function renderTopbar() {
  const user = S.account?.name || "Gebruiker";
  const role = S.isAdmin ? "Beheerder" : "Gebruiker";
  return `
    <div class="v-topbar">
      <div class="v-logo">V</div>
      <div class="v-topbar-title">Schademeldingen</div>
      <div class="v-topbar-nav">
        <button class="v-nav-btn ${S.view.name === "dashboard" ? "active" : ""}" data-nav="dashboard">Dashboard</button>
        <button class="v-nav-btn ${S.view.name === "create" ? "active" : ""}" data-nav="create">Nieuwe melding</button>
        ${S.isAdmin ? `<button class="v-nav-btn ${S.view.name === "settings" ? "active" : ""}" data-nav="settings">Instellingen</button>` : ""}
      </div>
      <div class="v-topbar-right">
        <div class="v-user-name">${esc(user)}</div>
        <span class="v-role-badge">${role}</span>
        <button class="btn-icon" data-action="logout" title="Afmelden">
          <i data-lucide="log-out"></i>
        </button>
      </div>
    </div>
  `;
}

/* ─────────────────────────────────────────────────────────────────────────── *
 *  SIDEBAR
 * ─────────────────────────────────────────────────────────────────────────── */
function renderSidebar() {
  return `
    <div class="app-sidebar">
      <div class="nav-group">
        <div class="nav-group-title">Menu</div>
        <div class="nav-item ${S.view.name === "dashboard" ? "active" : ""}" data-nav="dashboard">
          <i data-lucide="layout-grid"></i>
          <span>Dashboard</span>
        </div>
        <div class="nav-item ${S.view.name === "create" ? "active" : ""}" data-nav="create">
          <i data-lucide="plus-circle"></i>
          <span>Nieuwe melding</span>
        </div>
      </div>
      ${S.isAdmin ? `
      <div class="nav-group">
        <div class="nav-group-title">Beheer</div>
        <div class="nav-item ${S.view.name === "settings" ? "active" : ""}" data-nav="settings">
          <i data-lucide="settings"></i>
          <span>Instellingen</span>
        </div>
      </div>
      ` : ""}
    </div>
  `;
}

/* ─────────────────────────────────────────────────────────────────────────── *
 *  VIEW ROUTER
 * ─────────────────────────────────────────────────────────────────────────── */
function renderView() {
  const { name, id } = S.view;
  if (name === "dashboard") return renderDashboard();
  if (name === "create") return renderForm();
  if (name === "detail") return renderDetail(id);
  if (name === "settings") return renderSettings();
  return `<div style="padding: 28px;">View niet gevonden</div>`;
}

/* ─────────────────────────────────────────────────────────────────────────── *
 *  DASHBOARD
 * ─────────────────────────────────────────────────────────────────────────── */
function renderDashboard() {
  const fasen = [
    { key: "", label: "Alle fasen" },
    { key: "inkomend", label: "Ontvangst" },
    { key: "voorraad", label: "Voorraad" },
    { key: "uitgaand", label: "Verzending" },
    { key: "levering", label: "Levering" },
  ];

  const statussen = [
    { key: "", label: "Alle statussen" },
    { key: "Nieuw", label: "Nieuw" },
    { key: "In behandeling", label: "In behandeling" },
    { key: "Afgehandeld", label: "Afgehandeld" },
  ];

  const filtered = S.reports.filter((r) => {
    if (S.filter.q && !r.id.includes(S.filter.q.toUpperCase())) return false;
    if (S.filter.fase && r.type !== S.filter.fase) return false;
    if (S.filter.status && r.status !== S.filter.status) return false;
    return true;
  });

  return `
    <div style="padding: 28px 32px; max-width: 1100px; margin: 0 auto;">
      <h1 style="font-size: 20px; font-weight: 700; margin-bottom: 20px;">Dashboard</h1>

      <div class="dashboard-filters">
        <input type="text" class="filter-input" placeholder="Zoek melding-ID…" data-filter="q" value="${esc(S.filter.q)}" />
        <select class="filter-select" data-filter="fase">
          ${fasen.map((f) => `<option value="${f.key}" ${S.filter.fase === f.key ? "selected" : ""}>${esc(f.label)}</option>`).join("")}
        </select>
        <select class="filter-select" data-filter="status">
          ${statussen.map((s) => `<option value="${s.key}" ${S.filter.status === s.key ? "selected" : ""}>${esc(s.label)}</option>`).join("")}
        </select>
      </div>

      ${filtered.length === 0 ? `
        <div class="empty-state">
          <div class="empty-state-icon">📋</div>
          <div class="empty-state-title">Geen meldingen gevonden</div>
          <div class="empty-state-desc">Maak een nieuwe melding aan of pas je filters aan</div>
        </div>
      ` : `
        <div class="v-list">
          ${filtered.map((r) => renderCardRow(r)).join("")}
        </div>
      `}
    </div>
  `;
}

function renderCardRow(r) {
  const stage = stageByKey(r.type);
  const meta = STATUS_META[r.status];
  const faseLabel = KEY_TO_FASE[r.type] || r.type;

  return `
    <div class="v-card-row" data-click="detail" data-id="${r.id}">
      <div class="v-card-body">
        <div class="v-card-title">${esc(r.id)}</div>
        <div class="v-card-meta">
          <span class="v-pill v-pill-blue">${esc(stage?.label || "?")}</span>
          <span class="v-meta-item">
            <i data-lucide="box" style="width: 14px; height: 14px;"></i>
            ${esc(r.artikelnummer || "?")}
          </span>
          <span class="v-meta-item">
            <i data-lucide="calendar" style="width: 14px; height: 14px;"></i>
            ${r.aangemaakt || "?"}
          </span>
        </div>
      </div>
      <div class="v-card-right">
        <span class="v-badge ${meta?.badgeClass || "v-badge-info"}">
          <span class="v-badge-dot"></span>
          ${esc(r.status || "?")}
        </span>
        ${r.aantal ? `<span style="font-size: 12px; color: var(--muted);">× ${r.aantal}</span>` : ""}
      </div>
    </div>
  `;
}

/* ─────────────────────────────────────────────────────────────────────────── *
 *  FORMULIER (NIEUW/BEWERK)
 * ─────────────────────────────────────────────────────────────────────────── */
function renderForm(reportId) {
  const isNew = !reportId;
  const report = reportId ? S.reports.find((r) => r.id === reportId) : null;
  const defaultStage = report ? STAGES.find((s) => s.key === report.type) : STAGES[0];

  return `
    <div style="padding: 28px 32px; max-width: 900px; margin: 0 auto;">
      <h1 style="font-size: 20px; font-weight: 700; margin-bottom: 20px;">${isNew ? "Nieuwe melding" : `Bewerk ${esc(reportId)}`}</h1>

      <div class="stage-tabs">
        ${STAGES.map((s) => `
          <button class="stage-tab ${s === defaultStage ? "active" : ""}" data-stage="${s.key}">
            <i data-lucide="${s.icon}" style="width: 14px; height: 14px; margin-right: 6px;"></i>
            ${esc(s.stap)}
          </button>
        `).join("")}
      </div>

      <div class="v-form-wrap">
        <div id="formfields" data-stage="${defaultStage.key}" data-id="${reportId || ""}">
          ${renderFormFields(defaultStage, report)}
        </div>
        <div class="v-form-actions" style="margin-top: 20px;">
          <button class="btn btn-primary" data-action="save">Opslaan</button>
          <button class="btn btn-ghost" data-action="cancel">Annuleren</button>
        </div>
      </div>

      <div style="margin-top: 24px;">
        <h3 style="font-size: 14px; font-weight: 700; margin-bottom: 12px;">Foto's</h3>
        <div class="upload-zone" data-drop="photos">
          <div class="upload-zone-icon">📷</div>
          <div class="upload-zone-text">Sleep foto's hier of klik om te selecteren</div>
          <input type="file" multiple accept="image/*" style="display: none;" data-upload="photos" />
        </div>
        <div class="photo-gallery" id="photo-gallery"></div>
      </div>
    </div>
  `;
}

function renderFormFields(stage, report) {
  const fields = [...COMMON_HEAD, ...stage.fields, ...COMMON_TAIL];
  const data = report || {};

  const rows = [];
  let row = [];
  for (const f of fields) {
    if (row.length === 2 || f.full) {
      if (row.length) rows.push(row);
      row = [];
    }
    row.push(f);
    if (f.full) {
      rows.push(row);
      row = [];
    }
  }
  if (row.length) rows.push(row);

  return `
    <div class="v-form-grid">
      ${rows.map((r) => r.map((f) => `
        <div ${f.full ? 'class="v-form-full"' : ""}>
          <div class="v-fg">
            <label>
              ${esc(f.label)}
              ${f.required ? '<span class="v-req">*</span>' : ""}
            </label>
            ${f.type === "select" ? `
              <select data-field="${f.name}">
                <option value="">— Selecteer —</option>
                ${f.options.map((o) => `<option value="${o}" ${data[f.name] === o ? "selected" : ""}>${esc(o)}</option>`).join("")}
              </select>
            ` : f.type === "textarea" ? `
              <textarea data-field="${f.name}" placeholder="${esc(f.placeholder || "")}">${esc(data[f.name] || "")}</textarea>
            ` : `
              <input type="${f.type}" data-field="${f.name}" placeholder="${esc(f.placeholder || "")}" value="${esc(data[f.name] || "")}" />
            `}
          </div>
        </div>
      `).join("")).join("")}
    </div>
  `;
}

/* ─────────────────────────────────────────────────────────────────────────── *
 *  DETAIL
 * ─────────────────────────────────────────────────────────────────────────── */
function renderDetail(reportId) {
  const report = S.reports.find((r) => r.id === reportId);
  if (!report) return `<div style="padding: 28px;">Melding niet gevonden</div>`;

  const stage = stageByKey(report.type);
  const meta = STATUS_META[report.status];

  return `
    <div style="padding: 28px 32px; max-width: 900px; margin: 0 auto;">
      <div style="display: flex; gap: 12px; margin-bottom: 20px;">
        <button class="btn btn-ghost" data-action="back">← Terug</button>
      </div>

      <div class="detail-header">
        <div class="detail-header-left">
          <div class="detail-header-id">#${esc(reportId)}</div>
          <div class="detail-header-title">${esc(report.omschrijving || report.artikelnummer || "?")}</div>
          <div class="detail-header-meta">
            <div class="detail-meta-item">
              <strong>Fase:</strong> ${esc(stage?.label || "?")}
            </div>
            <div class="detail-meta-item">
              <strong>Status:</strong> ${esc(report.status || "?")}
            </div>
            <div class="detail-meta-item">
              <strong>Aantal:</strong> ${report.aantal || 0}
            </div>
          </div>
        </div>
        <div>
          <span class="v-badge ${meta?.badgeClass || "v-badge-info"}">
            <span class="v-badge-dot"></span>
            ${esc(report.status || "?")}
          </span>
        </div>
      </div>

      <div style="background: var(--surface); border: 1px solid var(--border); border-radius: 10px; padding: 20px; margin-bottom: 20px;">
        <h3 style="font-size: 14px; font-weight: 700; margin-bottom: 12px;">Basisgegevens</h3>
        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 16px;">
          ${[
            { label: "Artikel", value: report.artikelnummer },
            { label: "Magazijn", value: report.warehouse },
            { label: "Prioriteit", value: report.prioriteit },
            { label: "Melder", value: report.melder },
            { label: "Gemeld op", value: report.aangemaakt },
          ].map((item) => `
            <div>
              <div style="font-size: 11px; font-weight: 700; color: var(--light); text-transform: uppercase;">
                ${esc(item.label)}
              </div>
              <div style="font-size: 13px; color: var(--text); margin-top: 4px;">
                ${esc(item.value || "—")}
              </div>
            </div>
          `).join("")}
        </div>
      </div>

      <div style="background: var(--surface); border: 1px solid var(--border); border-radius: 10px; padding: 20px; margin-bottom: 20px;">
        <h3 style="font-size: 14px; font-weight: 700; margin-bottom: 12px;">Fase-specifieke gegevens</h3>
        ${stage ? `<pre style="font-size: 12px; overflow: auto; color: var(--muted);">${esc(JSON.stringify({...report}, null, 2))}</pre>` : "Geen data"}
      </div>

      <div class="status-buttons">
        ${STATUS_FLOW.map((s) => `
          <button class="status-btn ${report.status === s ? "active" : ""}" data-status="${s}">
            ${esc(s)}
          </button>
        `).join("")}
      </div>

      ${S.isAdmin ? `
        <div style="margin-top: 20px;">
          <button class="btn btn-danger btn-sm" data-action="delete" data-id="${reportId}">Verwijderen</button>
        </div>
      ` : ""}
    </div>
  `;
}

/* ─────────────────────────────────────────────────────────────────────────── *
 *  INSTELLINGEN (ADMIN)
 * ─────────────────────────────────────────────────────────────────────────── */
function renderSettings() {
  if (!S.isAdmin) return `<div class="v-alert v-alert-err" style="margin: 20px;">Alleen beheerders kunnen instellingen wijzigen</div>`;

  return `
    <div style="padding: 28px 32px; max-width: 600px;">
      <h1 style="font-size: 20px; font-weight: 700; margin-bottom: 20px;">Instellingen</h1>

      <div class="v-form-wrap">
        <h3 style="font-size: 14px; font-weight: 700; margin-bottom: 12px;">E-mailmeldingen</h3>
        <div class="v-fg" style="margin-bottom: 14px;">
          <label>
            <input type="checkbox" ${S.settings.enabled ? "checked" : ""} data-setting="enabled" />
            E-mailmeldingen inschakelen
          </label>
        </div>

        <h3 style="font-size: 14px; font-weight: 700; margin: 20px 0 12px;">Ontvangers</h3>
        <div id="recipients-list"></div>

        <button class="btn btn-sm btn-ghost" style="margin-top: 10px;" data-action="add-recipient">+ E-mailadres toevoegen</button>

        <div style="margin-top: 20px; display: flex; gap: 8px;">
          <button class="btn btn-primary" data-action="settings-save">Opslaan</button>
          <button class="btn btn-ghost" data-action="cancel">Annuleren</button>
        </div>
      </div>
    </div>
  `;
}

/* ─────────────────────────────────────────────────────────────────────────── *
 *  LOGIN
 * ─────────────────────────────────────────────────────────────────────────── */
function loginScreen() {
  return `
    <div class="login-screen">
      <div class="login-card">
        <div class="login-logo">
          <div class="login-logo-icon">V</div>
          <span style="font-size: 16px; font-weight: 700; color: var(--navy);">Verpa</span>
        </div>
        <h1 class="login-title">Schademeldingen</h1>
        <p class="login-subtitle">Meld schade aan goederen in het magazijn of bij levering.</p>
        <button class="btn btn-primary btn-lg" style="width: 100%;" data-action="login">
          Inloggen met Microsoft
        </button>
      </div>
    </div>
  `;
}

/* ─────────────────────────────────────────────────────────────────────────── *
 *  STUBS: Authentication + SharePoint (MSAL + Graph)
 *  Vervang deze met echte implementatie
 * ─────────────────────────────────────────────────────────────────────────── */
async function initAuth() {
  // TODO: implementeer MSAL
  return null; // of echte account
}

function isAdmin() {
  return false; // TODO: check app roles
}

async function fetchReports() {
  // TODO: haal meldingen uit SharePoint
  return [
    {
      id: "SCH-001",
      type: "inkomend",
      warehouse: "LAAKDAL",
      artikelnummer: "100234",
      omschrijving: "Stoelkussen",
      aantal: 5,
      prioriteit: "Hoog",
      melder: "Jan Pieterse",
      status: "Nieuw",
      aangemaakt: "2025-01-15",
      fotos: [],
      leverancier: "DHL Supply",
      inkooporder: "PO-45782",
    },
  ];
}

async function createReport(data) {
  // TODO: opslaan naar SharePoint
  return { ...data, spId: "temp-id", fotos: [] };
}

async function updateReport(spId, data) {
  // TODO: bijwerken in SharePoint
}

async function deleteReport(spId) {
  // TODO: verwijderen uit SharePoint
}

async function loadSettings() {
  // TODO: laad instellingen
  return { spId: null, enabled: false, recipients: [] };
}

async function saveSettings(data) {
  // TODO: sla instellingen op
}

async function uploadFotos(reportId, files) {
  // TODO: upload naar SharePoint
  return [];
}

async function notifyNewCase(report) {
  // TODO: stuur e-mail als enabled
}

/* ─────────────────────────────────────────────────────────────────────────── *
 *  EVENT HANDLERS
 * ─────────────────────────────────────────────────────────────────────────── */
function wireEvents() {
  // Navigation
  $$("[data-nav]").forEach((el) => {
    el.onclick = () => go({ name: el.getAttribute("data-nav") });
  });

  // Filters
  $$("[data-filter]").forEach((el) => {
    el.onchange = el.oninput = () => {
      S.filter[el.getAttribute("data-filter")] = el.value;
      render();
    };
  });

  // Card clicks
  $$("[data-click='detail']").forEach((el) => {
    el.onclick = () => go({ name: "detail", id: el.getAttribute("data-id") });
  });

  // Status buttons
  $$("[data-status]").forEach((el) => {
    el.onclick = () => changeStatus(S.view.id, el.getAttribute("data-status"));
  });

  // Actions
  const actions = {
    login: () => initAuth().then(() => { S.account = { name: "Test User" }; render(); }),
    logout: () => { S.account = null; S.view = { name: "dashboard" }; render(); },
    cancel: () => go({ name: "dashboard" }),
    back: () => go({ name: "dashboard" }),
    save: saveReport,
    delete: (b) => confirmDelete(b.getAttribute("data-id")),
    "add-recipient": addRecipientField,
    "settings-save": saveSettingsFromUI,
  };

  $$("[data-action]").forEach((b) => {
    b.onclick = () => {
      const action = actions[b.getAttribute("data-action")];
      if (action) action(b);
    };
  });

  // Stage tabs
  $$(".stage-tab").forEach((el) => {
    el.onclick = () => {
      const stage = stageByKey(el.getAttribute("data-stage"));
      if (stage) {
        $$(".stage-tab").forEach((t) => t.classList.remove("active"));
        el.classList.add("active");
        const formBox = $("#formfields");
        if (formBox) {
          formBox.setAttribute("data-stage", stage.key);
          formBox.innerHTML = renderFormFields(stage, {});
        }
      }
    };
  });

  // Foto's
  const uploadZone = $("[data-drop='photos']");
  const uploadInput = $('[data-upload="photos"]');
  if (uploadZone && uploadInput) {
    uploadZone.onclick = () => uploadInput.click();
    uploadZone.ondrop = (e) => {
      e.preventDefault();
      uploadZone.classList.remove("drag-over");
      if (e.dataTransfer.files) addFiles(e.dataTransfer.files);
    };
    uploadZone.ondragover = (e) => {
      e.preventDefault();
      uploadZone.classList.add("drag-over");
    };
    uploadZone.ondragleave = () => uploadZone.classList.remove("drag-over");
    uploadInput.onchange = (e) => addFiles(e.target.files);
  }

  renderFormThumbs();
}

function renderFormThumbs() {
  const gallery = $("#photo-gallery");
  if (!gallery) return;
  gallery.innerHTML = formPhotos.map((p, i) => `
    <div class="photo-thumb">
      <img src="${p.url}" />
      <button class="photo-thumb-del" data-del="${i}">✕</button>
    </div>
  `).join("");

  $$("[data-del]").forEach((btn) => {
    btn.onclick = (e) => {
      e.stopPropagation();
      formPhotos.splice(+btn.getAttribute("data-del"), 1);
      renderFormThumbs();
    };
  });
}

function addFiles(fileList) {
  const items = Array.from(fileList || []).map((file) => ({ name: file.name, url: URL.createObjectURL(file), file }));
  formPhotos.push(...items);
  renderFormThumbs();
}

async function saveReport() {
  const box = $("#formfields");
  if (!box) return;

  const stageKey = box.getAttribute("data-stage");
  const id = box.getAttribute("data-id");
  const stage = stageByKey(stageKey);
  const fields = [...COMMON_HEAD, ...stage.fields, ...COMMON_TAIL];
  const data = {};

  box.querySelectorAll("[data-field]").forEach((el) => {
    data[el.getAttribute("data-field")] = el.value;
  });

  // Validatie
  let bad = null;
  box.querySelectorAll("[data-field]").forEach((el) => el.classList.remove("v-err"));
  fields.forEach((f) => {
    if (f.required && !String(data[f.name] || "").trim()) {
      const el = box.querySelector(`[data-field='${f.name}']`);
      if (el) {
        el.classList.add("v-err");
        if (!bad) bad = el;
      }
    }
  });
  if (bad) {
    bad.scrollIntoView({ behavior: "smooth", block: "center" });
    return;
  }

  const btn = $("[data-action='save']");
  btn.disabled = true;
  btn.textContent = formPhotos.length ? "Foto's uploaden…" : "Opslaan…";

  try {
    let fotoRefs = [];
    if (formPhotos.length) fotoRefs = await uploadFotos(id, formPhotos.map((p) => p.file));
    const rec = await createReport({
      ...data,
      id,
      type: stageKey,
      status: "Nieuw",
      aangemaakt: todayISO(),
      aantal: Number(data.aantal) || 0,
      fotos: fotoRefs,
    });
    S.reports.unshift(rec);
    formPhotos = [];
    notifyNewCase(rec).catch((e) => console.warn("E-mailmelding niet verstuurd:", e));
    toast(`Melding ${rec.id} opgeslagen`);
    go({ name: "detail", id: rec.id });
  } catch (e) {
    btn.disabled = false;
    btn.textContent = "Melding opslaan";
    toast("Opslaan mislukt — probeer opnieuw", "err");
    console.error(e);
  }
}

async function changeStatus(id, status) {
  const r = S.reports.find((x) => x.id === id);
  if (!r) return;
  const prev = r.status;
  r.status = status;
  render();
  try {
    await updateReport(r.spId, { status });
  } catch (e) {
    r.status = prev;
    render();
    toast("Status bijwerken mislukt", "err");
  }
}

function confirmDelete(id) {
  if (!S.isAdmin) {
    toast("Alleen beheerders kunnen verwijderen", "err");
    return;
  }
  const root = $("#overlay-root");
  root.innerHTML = `
    <div class="modal-overlay">
      <div class="modal-content">
        <h3 style="font-size: 16px; font-weight: 700;">Melding verwijderen?</h3>
        <p style="font-size: 13px; color: var(--muted); margin: 8px 0 20px;">
          ${esc(id)} wordt definitief verwijderd.
        </p>
        <div style="display: flex; gap: 8px; justify-content: flex-end;">
          <button class="btn btn-ghost" data-cancel>Annuleren</button>
          <button class="btn btn-danger" data-confirm>Verwijderen</button>
        </div>
      </div>
    </div>
  `;

  $("[data-cancel]").onclick = () => (root.innerHTML = "");
  $("[data-confirm]").onclick = async () => {
    root.innerHTML = "";
    const r = S.reports.find((x) => x.id === id);
    if (!r) return;
    try {
      await deleteReport(r.spId);
      S.reports = S.reports.filter((x) => x.id !== id);
      toast("Melding verwijderd");
      go({ name: "dashboard" });
    } catch (e) {
      toast("Verwijderen mislukt", "err");
    }
  };
}

function addRecipientField() {
  S.settings.recipients.push({ email: "" });
  render();
}

async function saveSettingsFromUI() {
  S.settings.enabled = !!$('input[data-setting="enabled"]')?.checked;
  const btn = $("[data-action='settings-save']");
  if (btn) {
    btn.disabled = true;
    btn.textContent = "Opslaan…";
  }
  try {
    await saveSettings(S.settings);
    toast("Instellingen opgeslagen");
    go({ name: "dashboard" });
  } catch (e) {
    console.error(e);
    if (btn) {
      btn.disabled = false;
      btn.textContent = "Opslaan";
    }
    toast("Opslaan mislukt", "err");
  }
}

/* ─────────────────────────────────────────────────────────────────────────── *
 *  INIT
 * ─────────────────────────────────────────────────────────────────────────── */
(async function init() {
  render(); // Show loading
  if (!CONFIG.clientId || !CONFIG.tenantId) {
    S.loading = false;
    $("#app").innerHTML = `
      <div class="v-alert v-alert-err" style="margin: 20px;">
        <strong>Configuratie incompleet</strong><br>
        Vul alstublieft CONFIG.clientId en CONFIG.tenantId in de app.js in.
      </div>
    `;
    return;
  }

  try {
    S.account = await initAuth();
    if (!S.account) {
      S.loading = false;
      $("#app").innerHTML = loginScreen();
      wireEvents();
      icons();
      return;
    }
    S.isAdmin = isAdmin();
    S.reports = await fetchReports();
    S.settings = await loadSettings().catch(() => ({ spId: null, enabled: false, recipients: [] }));
    
    const caseId = new URLSearchParams(location.search).get("case");
    if (caseId && S.reports.find((r) => r.id === caseId)) {
      S.view = { name: "detail", id: caseId };
    }
    
    S.loading = false;
    render();
  } catch (e) {
    console.error(e);
    S.loading = false;
    S.error = e.message || String(e);
    render();
  }
})();
