/* ============================================================================
 *  SCHADEMELDINGEN — zelfstandige app voor Cloudflare Pages
 *  MSAL + Microsoft Graph → SharePoint + Foto's + Chat/Reacties
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

/* ─────────────────────────────────────────────────────────────────────────── */
const COMMON_HEAD = [
  { name: "warehouse", label: "Magazijn", type: "select", options: ["LAAKDAL", "STORA"], required: true },
  { name: "artikelnummer", label: "Artikelnummer", type: "text", placeholder: "bv. 100234", required: true, mono: true },
  { name: "omschrijving", label: "Omschrijving", type: "text", placeholder: "Artikelomschrijving", full: true },
  { name: "aantal", label: "Aantal beschadigd", type: "number", placeholder: "0", required: true },
];
const COMMON_TAIL = [
  { name: "prioriteit", label: "Prioriteit", type: "select", options: ["Laag", "Normaal", "Hoog"], required: true },
  { name: "melder", label: "Gemeld door", type: "text", placeholder: "Naam" },
  { name: "opmerkingen", label: "Opmerkingen", type: "textarea", placeholder: "Extra details…", full: true },
];

const STAGES = [
  { key: "inkomend", stap: "Ontvangst", label: "Inkomende goederen", sub: "Schade bij ontvangst", icon: "arrow-down-to-line",
    fields: [
      { name: "leverancier", label: "Leverancier", type: "text", placeholder: "Naam leverancier" },
      { name: "inkooporder", label: "Inkooporder-nr.", type: "text", placeholder: "BC inkooporder", mono: true },
      { name: "ontvangstdatum", label: "Ontvangstdatum", type: "date" },
      { name: "vervoerder", label: "Vervoerder", type: "text", placeholder: "Transporteur" },
      { name: "typeSchade", label: "Soort schade", type: "select", options: ["Transportschade", "Verpakkingsschade", "Productdefect", "Nat / vochtig", "Ontbrekend"] },
      { name: "actie", label: "Actie", type: "select", options: ["Geweigerd", "Onder voorbehoud", "Volledig", "Retour"] },
    ]
  },
  { key: "voorraad", stap: "Voorraad", label: "Schade aan voorraad", sub: "Tijdens opslag", icon: "warehouse",
    fields: [
      { name: "locatie", label: "Locatie / bin", type: "text", placeholder: "A-12-03", mono: true },
      { name: "oorzaak", label: "Oorzaak", type: "select", options: ["Handling (val)", "Heftruck", "THT verlopen", "Waterschade", "Mispick", "Onbekend"] },
      { name: "ontdektBij", label: "Ontdekt bij", type: "select", options: ["Orderpicking", "Cyclustelling", "Inspectie", "Bijvullen"] },
      { name: "gekoppeldeOrder", label: "Gekoppelde order", type: "text", placeholder: "VO-...", mono: true },
      { name: "afhandeling", label: "Afhandeling", type: "select", options: ["Afgeschreven", "Afgeprijsd", "Hersteld", "Quarantaine"] },
    ]
  },
  { key: "uitgaand", stap: "Verzending", label: "Uitgaande goederen", sub: "Bij verzending", icon: "package-check",
    fields: [
      { name: "klant", label: "Klant", type: "text", placeholder: "Naam klant" },
      { name: "verkooporder", label: "Verkooporder-nr.", type: "text", placeholder: "BC VO", mono: true },
      { name: "ontdektTijdens", label: "Ontdekt", type: "select", options: ["Verpakken", "Laden", "Eindcontrole"] },
      { name: "vervoerder", label: "Vervoerder", type: "text", placeholder: "Transporteur" },
      { name: "actie", label: "Actie", type: "select", options: ["Opnieuw picken", "Vervangen", "Gesplitst", "Uitgesteld"] },
    ]
  },
  { key: "levering", stap: "Levering", label: "Schade bij levering", sub: "Na aflevering", icon: "truck",
    fields: [
      { name: "klant", label: "Klant", type: "text", placeholder: "Naam klant" },
      { name: "verkooporder", label: "Verkooporder-nr.", type: "text", placeholder: "BC VO", mono: true },
      { name: "leveringsdatum", label: "Leveringsdatum", type: "date" },
      { name: "vervoerder", label: "Vervoerder", type: "text", placeholder: "Transporteur" },
      { name: "gemeldDoor", label: "Gemeld door", type: "select", options: ["Klant", "Chauffeur", "Vertegenwoordiger"] },
      { name: "claimVervoerder", label: "Claim vervoerder", type: "select", options: ["Ja", "Nee", "In behandeling"] },
      { name: "oplossing", label: "Oplossing", type: "select", options: ["Creditnota", "Herlevering", "Retour", "Geen"] },
    ]
  },
];

const stageByKey = (k) => STAGES.find((s) => s.key === k);
const KEY_TO_FASE = { inkomend: "Ontvangst", voorraad: "Voorraad", uitgaand: "Verzending", levering: "Levering" };
const STATUS_FLOW = ["Nieuw", "In behandeling", "Afgehandeld"];

/* ─────────────────────────────────────────────────────────────────────────── */
const S = {
  loading:  true,
  error:    null,
  account:  null,
  isAdmin:  false,
  reports:  [],
  settings: { spId: null, enabled: false, recipients: [] },
  comments: {},
  view:     { name: "dashboard" },
  filter:   { q: "", fase: "", status: "" },
};

let formPhotos = [];

const $ = (sel) => document.querySelector(sel);
const $$ = (sel) => document.querySelectorAll(sel);
const esc = (s) => String(s).replace(/[&<>"']/g, c => ({
  "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"
})[c]);
const todayISO = () => new Date().toISOString().split("T")[0];
const formatDate = (d) => d ? new Date(d).toLocaleDateString("nl-NL") : "-";

function toast(msg, type = "info") {
  let t = $("#v-toast");
  if (!t) {
    t = document.createElement("div");
    t.id = "v-toast";
    document.body.appendChild(t);
  }
  t.textContent = msg;
  t.className = `show ${type}`;
  setTimeout(() => t.classList.remove("show"), 3000);
}

function icons() {
  $$("[data-lucide]").forEach((el) => {
    const name = el.getAttribute("data-lucide");
    if (lucide.icons[name]) {
      el.innerHTML = lucide.icons[name].toSvg();
      el.firstChild?.setAttribute("stroke-width", "1.5");
    }
  });
}

function go(view) {
  S.view = view;
  render();
  window.scrollTo({ top: 0, behavior: "smooth" });
}

/* ─────────────────────────────────────────────────────────────────────────── */
function render() {
  const app = $("#app");
  if (!app) return;

  if (S.loading) {
    app.innerHTML = `<div class="loading-screen"><div class="loading-box"><div class="loading-spinner"></div><p>Laden…</p></div></div>`;
    return;
  }

  if (S.error) {
    app.innerHTML = `<div class="v-alert v-alert-err" style="margin: 20px;"><strong>Fout:</strong> ${esc(S.error)}</div>`;
    return;
  }

  if (!S.account) {
    app.innerHTML = loginScreen();
    wireEvents();
    icons();
    return;
  }

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

/* ─────────────────────────────────────────────────────────────────────────── */
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

function renderView() {
  const { name, id } = S.view;
  if (name === "dashboard") return renderDashboard();
  if (name === "create") return renderForm();
  if (name === "detail") return renderDetail(id);
  if (name === "settings") return renderSettings();
  return `<div style="padding: 28px;">View niet gevonden</div>`;
}

/* ─────────────────────────────────────────────────────────────────────────── */
function renderDashboard() {
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
          <option value="">Alle fasen</option>
          <option value="inkomend">Ontvangst</option>
          <option value="voorraad">Voorraad</option>
          <option value="uitgaand">Verzending</option>
          <option value="levering">Levering</option>
        </select>
        <select class="filter-select" data-filter="status">
          <option value="">Alle statussen</option>
          <option value="Nieuw">Nieuw</option>
          <option value="In behandeling">In behandeling</option>
          <option value="Afgehandeld">Afgehandeld</option>
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
  const statusClass = r.status === "Nieuw" ? "v-badge-info" : r.status === "In behandeling" ? "v-badge-pending" : "v-badge-approved";
  return `
    <div class="v-card-row" data-click="detail" data-id="${r.id}">
      <div class="v-card-body">
        <div class="v-card-title">${esc(r.id)}</div>
        <div class="v-card-meta">
          <span class="v-pill v-pill-blue">${esc(stage?.label || "?")}</span>
          <span class="v-meta-item"><i data-lucide="package" style="width: 14px;"></i>${esc(r.artikelnummer || "?")}</span>
          <span class="v-meta-item"><i data-lucide="calendar" style="width: 14px;"></i>${formatDate(r.aangemaakt)}</span>
        </div>
      </div>
      <div class="v-card-right">
        <span class="v-badge ${statusClass}"><span class="v-badge-dot"></span>${esc(r.status || "?")}</span>
        ${r.aantal ? `<span style="font-size: 12px; color: var(--muted);">× ${r.aantal}</span>` : ""}
      </div>
    </div>
  `;
}

/* ─────────────────────────────────────────────────────────────────────────── */
function renderForm() {
  const defaultStage = STAGES[0];

  return `
    <div style="padding: 28px 32px; max-width: 900px; margin: 0 auto;">
      <h1 style="font-size: 20px; font-weight: 700; margin-bottom: 20px;">Nieuwe melding</h1>

      <div class="stage-tabs">
        ${STAGES.map((s) => `
          <button class="stage-tab ${s === defaultStage ? "active" : ""}" data-stage="${s.key}">
            <i data-lucide="${s.icon}" style="width: 14px; margin-right: 6px;"></i> ${esc(s.stap)}
          </button>
        `).join("")}
      </div>

      <div class="v-form-wrap">
        <div id="formfields" data-stage="${defaultStage.key}">
          ${renderFormFields(defaultStage, {})}
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

function renderFormFields(stage, data = {}) {
  const fields = [...COMMON_HEAD, ...stage.fields, ...COMMON_TAIL];
  const rows = [];
  let row = [];
  
  fields.forEach((f) => {
    if (row.length === 2 || f.full) {
      if (row.length) rows.push(row);
      row = [];
    }
    row.push(f);
    if (f.full) {
      rows.push(row);
      row = [];
    }
  });
  if (row.length) rows.push(row);

  return `
    <div class="v-form-grid">
      ${rows.map((r) => r.map((f) => `
        <div ${f.full ? 'class="v-form-full"' : ""}>
          <div class="v-fg">
            <label>${esc(f.label)} ${f.required ? '<span class="v-req">*</span>' : ""}</label>
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

/* ─────────────────────────────────────────────────────────────────────────── */
function renderDetail(reportId) {
  const report = S.reports.find((r) => r.id === reportId);
  if (!report) return `<div style="padding: 28px;">Melding niet gevonden</div>`;

  const stage = stageByKey(report.type);
  const statusClass = report.status === "Nieuw" ? "v-badge-info" : report.status === "In behandeling" ? "v-badge-pending" : "v-badge-approved";
  const comments = S.comments[reportId] || [];

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
            <div class="detail-meta-item"><strong>Fase:</strong> ${esc(stage?.label || "?")}</div>
            <div class="detail-meta-item"><strong>Aantal:</strong> ${report.aantal || 0}</div>
            <div class="detail-meta-item"><strong>Gemeld:</strong> ${formatDate(report.aangemaakt)}</div>
          </div>
        </div>
        <div>
          <span class="v-badge ${statusClass}"><span class="v-badge-dot"></span>${esc(report.status || "?")}</span>
        </div>
      </div>

      <div style="background: var(--surface); border: 1px solid var(--border); border-radius: 10px; padding: 20px; margin-bottom: 20px;">
        <h3 style="font-size: 14px; font-weight: 700; margin-bottom: 12px;">Basisgegevens</h3>
        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 16px;">
          <div><div style="font-size: 11px; font-weight: 700; color: var(--light); text-transform: uppercase;">Artikel</div><div style="font-size: 13px; margin-top: 4px;">${esc(report.artikelnummer || "—")}</div></div>
          <div><div style="font-size: 11px; font-weight: 700; color: var(--light); text-transform: uppercase;">Magazijn</div><div style="font-size: 13px; margin-top: 4px;">${esc(report.warehouse || "—")}</div></div>
          <div><div style="font-size: 11px; font-weight: 700; color: var(--light); text-transform: uppercase;">Prioriteit</div><div style="font-size: 13px; margin-top: 4px;">${esc(report.prioriteit || "—")}</div></div>
          <div><div style="font-size: 11px; font-weight: 700; color: var(--light); text-transform: uppercase;">Melder</div><div style="font-size: 13px; margin-top: 4px;">${esc(report.melder || "—")}</div></div>
        </div>
      </div>

      <div class="status-buttons">
        ${STATUS_FLOW.map((s) => `
          <button class="status-btn ${report.status === s ? "active" : ""}" data-status="${s}">
            ${esc(s)}
          </button>
        `).join("")}
      </div>

      <div style="margin-top: 24px; padding: 20px; background: var(--gray-bg); border-radius: 10px;">
        <h3 style="font-size: 14px; font-weight: 700; margin-bottom: 12px;">Chat / Reacties</h3>
        <div style="max-height: 300px; overflow-y: auto; margin-bottom: 12px; padding: 12px; background: var(--surface); border-radius: 8px;">
          ${comments.length === 0 ? `<p style="color: var(--muted); font-size: 12px;">Geen reacties nog</p>` : `
            ${comments.map((c) => `
              <div style="margin-bottom: 12px; padding: 10px; background: var(--navy-bg); border-radius: 6px;">
                <div style="font-weight: 600; font-size: 12px; color: var(--navy);">${esc(c.author)}</div>
                <div style="font-size: 12px; color: var(--light); margin-top: 2px;">${esc(c.date)}</div>
                <div style="font-size: 13px; margin-top: 6px;">${esc(c.text)}</div>
              </div>
            `).join("")}
          `}
        </div>
        <div style="display: flex; gap: 8px;">
          <input type="text" class="filter-input" placeholder="Typ een reactie…" style="flex: 1; margin: 0;" data-comment-input="${reportId}" />
          <button class="btn btn-primary btn-sm" data-action="add-comment" data-id="${reportId}">Verzenden</button>
        </div>
      </div>

      ${S.isAdmin ? `
        <div style="margin-top: 20px;">
          <button class="btn btn-danger btn-sm" data-action="delete" data-id="${reportId}">Verwijderen</button>
        </div>
      ` : ""}
    </div>
  `;
}

/* ─────────────────────────────────────────────────────────────────────────── */
function renderSettings() {
  if (!S.isAdmin) return `<div class="v-alert v-alert-err" style="margin: 20px;">Alleen beheerders kunnen instellingen wijzigen</div>`;

  return `
    <div style="padding: 28px 32px; max-width: 600px;">
      <h1 style="font-size: 20px; font-weight: 700; margin-bottom: 20px;">Instellingen</h1>

      <div class="v-form-wrap">
        <h3 style="font-size: 14px; font-weight: 700; margin-bottom: 12px;">E-mailmeldingen</h3>
        <div class="v-fg" style="margin-bottom: 14px;">
          <label style="display: flex; gap: 8px; align-items: center; font-weight: 400;">
            <input type="checkbox" ${S.settings.enabled ? "checked" : ""} data-setting="enabled" />
            E-mailmeldingen inschakelen
          </label>
        </div>

        <div style="margin-top: 20px; display: flex; gap: 8px;">
          <button class="btn btn-primary" data-action="settings-save">Opslaan</button>
          <button class="btn btn-ghost" data-action="cancel">Annuleren</button>
        </div>
      </div>
    </div>
  `;
}

/* ─────────────────────────────────────────────────────────────────────────── */
function loginScreen() {
  return `
    <div class="v-login-screen">
      <div class="v-login-card">
        <div class="v-login-logo">
          <div style="width: 40px; height: 40px; background: var(--navy); border-radius: 8px; display: flex; align-items: center; justify-content: center; color: #fff; font-weight: 800; font-size: 18px;">V</div>
          <span style="color: var(--navy); font-weight: 700;">Verpa</span>
        </div>
        <h1>Schademeldingen</h1>
        <p>Meld en volg schade aan goederen in het magazijn of bij levering.</p>
        <button class="btn btn-primary btn-lg" style="width: 100%;" data-action="login">
          <i data-lucide="microsoft"></i> Inloggen met Microsoft
        </button>
      </div>
    </div>
  `;
}

/* ─────────────────────────────────────────────────────────────────────────── */
function wireEvents() {
  $$("[data-nav]").forEach((el) => {
    el.onclick = () => go({ name: el.getAttribute("data-nav") });
  });

  $$("[data-filter]").forEach((el) => {
    el.onchange = el.oninput = () => {
      S.filter[el.getAttribute("data-filter")] = el.value;
      render();
    };
  });

  $$("[data-click='detail']").forEach((el) => {
    el.onclick = () => go({ name: "detail", id: el.getAttribute("data-id") });
  });

  $$("[data-status]").forEach((el) => {
    el.onclick = () => changeStatus(S.view.id, el.getAttribute("data-status"));
  });

  $$(".stage-tab").forEach((el) => {
    el.onclick = () => {
      const stage = stageByKey(el.getAttribute("data-stage"));
      if (stage) {
        $$(".stage-tab").forEach((t) => t.classList.remove("active"));
        el.classList.add("active");
        const box = $("#formfields");
        if (box) {
          box.setAttribute("data-stage", stage.key);
          box.innerHTML = renderFormFields(stage, {});
        }
      }
    };
  });

  const uploadZone = $("[data-drop='photos']");
  const uploadInput = $('[data-upload="photos"]');
  if (uploadZone && uploadInput) {
    uploadZone.onclick = () => uploadInput.click();
    uploadZone.ondrop = (e) => { e.preventDefault(); uploadZone.classList.remove("drag-over"); if (e.dataTransfer.files) addFiles(e.dataTransfer.files); };
    uploadZone.ondragover = (e) => { e.preventDefault(); uploadZone.classList.add("drag-over"); };
    uploadZone.ondragleave = () => uploadZone.classList.remove("drag-over");
    uploadInput.onchange = (e) => addFiles(e.target.files);
  }

  const actions = {
    login: () => { S.account = { name: "Test User", id: "test@verpa.be" }; S.isAdmin = true; render(); },
    logout: () => { S.account = null; S.view = { name: "dashboard" }; render(); },
    cancel: () => go({ name: "dashboard" }),
    back: () => go({ name: "dashboard" }),
    save: saveReport,
    delete: (b) => confirmDelete(b.getAttribute("data-id")),
    "add-comment": addComment,
    "settings-save": saveSettingsFromUI,
  };

  $$("[data-action]").forEach((b) => {
    b.onclick = () => {
      const action = actions[b.getAttribute("data-action")];
      if (action) action(b);
    };
  });

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
    btn.onclick = (e) => { e.stopPropagation(); formPhotos.splice(+btn.getAttribute("data-del"), 1); renderFormThumbs(); };
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
  const stage = stageByKey(stageKey);
  const fields = [...COMMON_HEAD, ...stage.fields, ...COMMON_TAIL];
  const data = {};

  box.querySelectorAll("[data-field]").forEach((el) => {
    data[el.getAttribute("data-field")] = el.value;
  });

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
  btn.textContent = "Opslaan…";

  try {
    const id = "SCH-" + String(S.reports.length + 1).padStart(3, "0");
    const rec = {
      ...data,
      id,
      type: stageKey,
      status: "Nieuw",
      aangemaakt: todayISO(),
      aantal: Number(data.aantal) || 0,
      fotos: formPhotos.length,
      spId: "sp-" + Date.now(),
    };
    S.reports.unshift(rec);
    S.comments[id] = [];
    formPhotos = [];
    toast(`Melding ${id} opgeslagen`, "ok");
    go({ name: "detail", id });
  } catch (e) {
    btn.disabled = false;
    btn.textContent = "Opslaan";
    toast("Opslaan mislukt", "err");
    console.error(e);
  }
}

async function changeStatus(id, status) {
  const r = S.reports.find((x) => x.id === id);
  if (!r) return;
  r.status = status;
  render();
  toast(`Status gewijzigd naar: ${status}`, "ok");
}

function confirmDelete(id) {
  if (!S.isAdmin) { toast("Alleen beheerders kunnen verwijderen", "err"); return; }
  const root = $("#overlay-root");
  root.innerHTML = `
    <div class="v-overlay">
      <div class="v-modal v-modal-sm">
        <div class="v-modal-head">
          <h3>Melding verwijderen?</h3>
        </div>
        <div class="v-modal-body">
          <p>${esc(id)} wordt permanent verwijderd.</p>
        </div>
        <div class="v-modal-foot">
          <button class="btn btn-ghost" data-cancel>Annuleren</button>
          <button class="btn btn-danger" data-confirm>Verwijderen</button>
        </div>
      </div>
    </div>
  `;

  $("[data-cancel]").onclick = () => (root.innerHTML = "");
  $("[data-confirm]").onclick = () => {
    root.innerHTML = "";
    S.reports = S.reports.filter((x) => x.id !== id);
    delete S.comments[id];
    toast("Melding verwijderd", "ok");
    go({ name: "dashboard" });
  };
}

function addComment(btn) {
  const id = btn.getAttribute("data-id");
  const input = $(`[data-comment-input="${id}"]`);
  const text = input?.value?.trim();
  if (!text) return;

  if (!S.comments[id]) S.comments[id] = [];
  S.comments[id].push({
    author: S.account?.name || "Anoniem",
    date: new Date().toLocaleString("nl-NL"),
    text,
  });
  input.value = "";
  render();
  toast("Reactie toegevoegd", "ok");
}

function saveSettingsFromUI() {
  S.settings.enabled = !!$('input[data-setting="enabled"]')?.checked;
  toast("Instellingen opgeslagen", "ok");
  go({ name: "dashboard" });
}

/* ─────────────────────────────────────────────────────────────────────────── */
(async function init() {
  render();
  await new Promise(r => setTimeout(r, 500));
  
  // DEMO MODE - vervang met MSAL later
  S.account = { name: "Jan Pieterse", id: "jan@verpa.be" };
  S.isAdmin = true;
  S.reports = [
    {
      id: "SCH-001",
      type: "inkomend",
      warehouse: "LAAKDAL",
      artikelnummer: "100234",
      omschrijving: "Stoelkussens",
      aantal: 5,
      prioriteit: "Hoog",
      melder: "Jan Pieterse",
      status: "Nieuw",
      aangemaakt: todayISO(),
      fotos: 2,
      spId: "sp-1",
    },
    {
      id: "SCH-002",
      type: "voorraad",
      warehouse: "STORA",
      artikelnummer: "200567",
      omschrijving: "Printpapier A4",
      aantal: 1,
      prioriteit: "Laag",
      melder: "Maria Santos",
      status: "In behandeling",
      aangemaakt: "2025-01-14",
      fotos: 0,
      spId: "sp-2",
    },
  ];
  S.comments["SCH-001"] = [
    { author: "Beheerder", date: "14-01-2025 10:30", text: "Foto's ontvangen, wacht op vervoerder reactie" },
  ];
  S.comments["SCH-002"] = [];
  
  S.loading = false;
  render();
})();
