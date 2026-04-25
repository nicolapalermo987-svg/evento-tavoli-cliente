/**
 * App organizzazione tavoli — stato in memoria + localStorage
 * Max 9 ospiti per tavolo; marcatori sulla piantina sincronizzati con i tavoli.
 */

const MAX_GUESTS = 9;
const STORAGE_KEY_BASE = "evento-tavoli-v1";

const state = {
  eventName: "",
  floorPlanDataUrl: "",
  tables: [],
  /** Legacy singola piantina (mantenuta solo per migrazione dati) */
  markerPositions: {},
  /** [{ id, name, imageDataUrl, markerPositions: { [tableId]: {x,y} } }] */
  floorPlans: [],
};

const acceptanceUiState = {
  query: "",
  selectedTableId: "",
};

const APP_ROLES = ["cliente", "struttura", "staff"];
const DISTRIBUTION_ROLE_RAW = String(window.APP_DISTRIBUTION_ROLE || "struttura").trim().toLowerCase();
const appRole = APP_ROLES.includes(DISTRIBUTION_ROLE_RAW) ? DISTRIBUTION_ROLE_RAW : "struttura";

function getStateStorageKey() {
  return `${STORAGE_KEY_BASE}-${appRole}`;
}

function canAccessArea(view) {
  if (appRole === "cliente") return view === "tables";
  return view === "tables" || view === "map" || view === "acceptance" || view === "segnatavolo" || view === "menu";
}

function canEditTablesArea() {
  return appRole === "cliente" || appRole === "struttura";
}

function canEditMapArea() {
  return appRole === "struttura";
}

function canEditStudioAreas() {
  return appRole === "struttura";
}

function canEditEventName() {
  return appRole === "cliente" || appRole === "struttura";
}

/** Durante `renderFloorPlans`, riapre il menu tavoli sullo stesso riquadro (es. dopo una checkbox). */
let preservePlanVisibilityMenuPlanId = null;

const els = {
  eventName: document.getElementById("eventName"),
  stickyToolbar: document.getElementById("stickyToolbar"),
  addTableBtn: document.getElementById("addTableBtn"),
  showTablesAreaBtn: document.getElementById("showTablesAreaBtn"),
  showMapAreaBtn: document.getElementById("showMapAreaBtn"),
  showAcceptanceAreaBtn: document.getElementById("showAcceptanceAreaBtn"),
  openSegnatavoloAreaBtn: document.getElementById("openSegnatavoloAreaBtn"),
  openMenuBookletAreaBtn: document.getElementById("openMenuBookletAreaBtn"),
  tablesAreaPanel: document.getElementById("tablesAreaPanel"),
  mapAreaPanel: document.getElementById("mapAreaPanel"),
  acceptanceAreaPanel: document.getElementById("acceptanceAreaPanel"),
  segnatavoloAreaPanel: document.getElementById("segnatavoloAreaPanel"),
  menuBookletAreaPanel: document.getElementById("menuBookletAreaPanel"),
  acceptanceSearchInput: document.getElementById("acceptanceSearchInput"),
  acceptanceResults: document.getElementById("acceptanceResults"),
  acceptanceTablePreview: document.getElementById("acceptanceTablePreview"),
  sortTablesBtn: document.getElementById("sortTablesBtn"),
  renumberTablesBtn: document.getElementById("renumberTablesBtn"),
  tablesList: document.getElementById("tablesList"),
  tableCountLabel: document.getElementById("tableCountLabel"),
  globalOccupancyLine: document.getElementById("globalOccupancyLine"),
  globalSummaryLine: document.getElementById("globalSummaryLine"),
  tableActionsDropdown: document.getElementById("tableActionsDropdown"),
  tableActionsMenuBtn: document.getElementById("tableActionsMenuBtn"),
  tableActionsMenu: document.getElementById("tableActionsMenu"),
  topActionsDropdown: document.getElementById("topActionsDropdown"),
  topActionsMenuBtn: document.getElementById("topActionsMenuBtn"),
  topActionsMenu: document.getElementById("topActionsMenu"),
  dropdownBackdrop: document.getElementById("dropdownBackdrop"),
  importClientBtn: document.getElementById("importClientBtn"),
  importFullBtn: document.getElementById("importFullBtn"),
  addPlanPanelBtn: document.getElementById("addPlanPanelBtn"),
  floorPlansContainer: document.getElementById("floorPlansContainer"),
  exportGuestsBtn: document.getElementById("exportGuestsBtn"),
  exportKitchenBtn: document.getElementById("exportKitchenBtn"),
  exportClientBtn: document.getElementById("exportClientBtn"),
  exportFullBtn: document.getElementById("exportFullBtn"),
  importInput: document.getElementById("importInput"),
  tableCardTpl: document.getElementById("tableCardTpl"),
  guestRowTpl: document.getElementById("guestRowTpl"),
  segnatavoloExportConfirm: document.getElementById("segnatavoloExportConfirm"),
  segnatavoloPreviewAllList: document.getElementById("segnatavoloPreviewAllList"),
  segnatavoloControlSampleCard: document.getElementById("segnatavoloControlSampleCard"),
  segnatavoloControlSampleWrap: document.getElementById("segnatavoloControlSampleWrap"),
  segnatavoloControlSampleDeco: document.getElementById("segnatavoloControlSampleDeco"),
  segnatavoloControlSampleTitle: document.getElementById("segnatavoloControlSampleTitle"),
  segnatavoloControlSampleSubtitle: document.getElementById("segnatavoloControlSampleSubtitle"),
  segnatavoloControlSampleNames: document.getElementById("segnatavoloControlSampleNames"),
  segnatavoloControlSampleEmpty: document.getElementById("segnatavoloControlSampleEmpty"),
  segnatavoloControlSampleFoot: document.getElementById("segnatavoloControlSampleFoot"),
  segnatavoloPreview: document.getElementById("segnatavoloPreview"),
  segnatavoloPreviewDeco: document.getElementById("segnatavoloPreviewDeco"),
  segnatavoloPreviewTitle: document.getElementById("segnatavoloPreviewTitle"),
  segnatavoloPreviewSubtitle: document.getElementById("segnatavoloPreviewSubtitle"),
  segnatavoloPreviewNames: document.getElementById("segnatavoloPreviewNames"),
  segnatavoloPreviewEmpty: document.getElementById("segnatavoloPreviewEmpty"),
  segnatavoloPreviewFoot: document.getElementById("segnatavoloPreviewFoot"),
  segnatavoloOccasionChips: document.getElementById("segnatavoloOccasionChips"),
  segnatavoloThemeChips: document.getElementById("segnatavoloThemeChips"),
  segnatavoloPaletteChips: document.getElementById("segnatavoloPaletteChips"),
  segnatavoloGraphicChips: document.getElementById("segnatavoloGraphicChips"),
  segnatavoloCustomColors: document.getElementById("segnatavoloCustomColors"),
  segnatavoloColorTitle: document.getElementById("segnatavoloColorTitle"),
  segnatavoloColorBody: document.getElementById("segnatavoloColorBody"),
  segnatavoloColorDeco: document.getElementById("segnatavoloColorDeco"),
  segnatavoloPaperFormat: document.getElementById("segnatavoloPaperFormat"),
  segnatavoloOrientation: document.getElementById("segnatavoloOrientation"),
  segnatavoloHeaderMode: document.getElementById("segnatavoloHeaderMode"),
  segnatavoloShowGuests: document.getElementById("segnatavoloShowGuests"),
  menuBookletLogoInput: document.getElementById("menuBookletLogoInput"),
  menuBookletDishSelect: document.getElementById("menuBookletDishSelect"),
  menuBookletEditorPanel: document.getElementById("menuBookletEditorPanel"),
  menuBookletEditorTitle: document.getElementById("menuBookletEditorTitle"),
  menuBookletEditorInput: document.getElementById("menuBookletEditorInput"),
  menuBookletEditorConfirmBtn: document.getElementById("menuBookletEditorConfirmBtn"),
  menuBookletEditorCancelBtn: document.getElementById("menuBookletEditorCancelBtn"),
  menuBookletList: document.getElementById("menuBookletList"),
  menuBookletExportBtn: document.getElementById("menuBookletExportBtn"),
  menuBookletPreviewStack: document.getElementById("menuBookletPreviewStack"),
  menuBookletPreviewLogo: document.getElementById("menuBookletPreviewLogo"),
  menuBookletPreviewEvent: document.getElementById("menuBookletPreviewEvent"),
  menuBookletPreviewList: document.getElementById("menuBookletPreviewList"),
};

function createEmptyGuest() {
  return { cognome: "", nome: "", menu: "adulto", note: "" };
}

function ensureTableGuestSlots(table) {
  if (table.customName == null) table.customName = "";
  if (table.tableNote == null) table.tableNote = "";
  const current = Array.isArray(table.guests) ? table.guests : [];
  const normalized = current.slice(0, MAX_GUESTS).map((g) => ({
    cognome: String((g && g.cognome) || ""),
    nome: String((g && g.nome) || ""),
    menu: g && g.menu === "bambino" ? "bambino" : "adulto",
    note: String(
      (g && g.note) ||
        [String((g && g.variazione) || "").trim(), String((g && g.allergia) || "").trim()]
          .filter(Boolean)
          .join(" - ")
    ),
  }));

  while (normalized.length < MAX_GUESTS) {
    normalized.push(createEmptyGuest());
  }
  table.guests = normalized;
}

function uid() {
  return "t_" + Math.random().toString(36).slice(2, 11);
}

function createFloorPlan(name = "") {
  return {
    id: uid(),
    name: String(name || "").trim(),
    imageDataUrl: "",
    markerPositions: {},
    /** Tavoli nascosti su questo riquadro (id). Default: nessuno = tutti visibili */
    hiddenTableIds: [],
  };
}

function sanitizeFloorPlanName(name) {
  const raw = String(name || "").trim();
  if (!raw) return "";
  if (/^piantina della sala\b/i.test(raw)) return "";
  return raw;
}

function ensureFloorPlansState() {
  if (!Array.isArray(state.floorPlans)) state.floorPlans = [];
  state.floorPlans = state.floorPlans
    .filter(Boolean)
    .map((plan, idx) => ({
      id: plan.id || uid(),
      name: sanitizeFloorPlanName(plan.name),
      imageDataUrl: String(plan.imageDataUrl || ""),
      markerPositions:
        plan.markerPositions && typeof plan.markerPositions === "object" ? { ...plan.markerPositions } : {},
      hiddenTableIds: Array.isArray(plan.hiddenTableIds)
        ? plan.hiddenTableIds.filter((id) => typeof id === "string" && id)
        : [],
    }));

  if (!state.floorPlans.length) {
    const migrated = createFloorPlan("");
    migrated.imageDataUrl = String(state.floorPlanDataUrl || "");
    migrated.markerPositions = state.markerPositions && typeof state.markerPositions === "object"
      ? { ...state.markerPositions }
      : {};
    state.floorPlans.push(migrated);
  }

  const validTableIds = new Set(state.tables.map((t) => t.id));
  for (const plan of state.floorPlans) {
    plan.hiddenTableIds = (plan.hiddenTableIds || []).filter((id) => validTableIds.has(id));
  }
}

function loadState() {
  try {
    const raw = localStorage.getItem(getStateStorageKey());
    if (!raw) return;
    const data = JSON.parse(raw);
    if (data.eventName != null) state.eventName = String(data.eventName);
    if (data.floorPlanDataUrl != null) state.floorPlanDataUrl = String(data.floorPlanDataUrl);
    if (Array.isArray(data.tables)) state.tables = data.tables;
    state.tables.forEach(ensureTableGuestSlots);
    if (data.markerPositions && typeof data.markerPositions === "object") {
      state.markerPositions = data.markerPositions;
    }
    if (Array.isArray(data.floorPlans)) {
      state.floorPlans = data.floorPlans;
    }
    ensureFloorPlansState();
  } catch (_) {
    /* ignore */
  }
}

function saveState() {
  try {
    localStorage.setItem(
      getStateStorageKey(),
      JSON.stringify({
        eventName: state.eventName,
        floorPlanDataUrl: state.floorPlanDataUrl,
        tables: state.tables,
        markerPositions: state.markerPositions,
        floorPlans: state.floorPlans,
      })
    );
  } catch (_) {
    /* quota piena o navigazione privata */
  }
}

function nextTableNumber() {
  const used = new Set();
  for (const t of state.tables) {
    const n = Number(t.number);
    if (!Number.isNaN(n) && n > 0) used.add(n);
  }

  let candidate = 1;
  while (used.has(candidate)) {
    candidate += 1;
  }
  return candidate;
}

function addTable() {
  if (!canEditTablesArea()) return;
  ensureFloorPlansState();
  const id = uid();
  state.tables.push({
    id,
    number: nextTableNumber(),
    customName: "",
    tableNote: "",
    guests: Array.from({ length: MAX_GUESTS }, createEmptyGuest),
  });
  for (const plan of state.floorPlans) {
    if (!plan.markerPositions[id]) {
      plan.markerPositions[id] = { x: 50, y: 50 };
    }
    if (Array.isArray(plan.hiddenTableIds) && plan.hiddenTableIds.includes(id)) {
      plan.hiddenTableIds = plan.hiddenTableIds.filter((x) => x !== id);
    }
  }
  saveState();
  renderTables();
  renderFloorMarkers();
}

function removeTable(tableId) {
  if (!canEditTablesArea()) return;
  ensureFloorPlansState();
  state.tables = state.tables.filter((t) => t.id !== tableId);
  for (const plan of state.floorPlans) {
    if (plan.markerPositions && plan.markerPositions[tableId]) {
      delete plan.markerPositions[tableId];
    }
    if (Array.isArray(plan.hiddenTableIds)) {
      plan.hiddenTableIds = plan.hiddenTableIds.filter((id) => id !== tableId);
    }
  }
  saveState();
  renderTables();
  renderFloorMarkers();
}

function sortTablesAscending() {
  if (!canEditTablesArea()) return;
  state.tables.sort((a, b) => {
    const aNum = Number(a.number);
    const bNum = Number(b.number);
    const aValid = Number.isFinite(aNum);
    const bValid = Number.isFinite(bNum);
    if (aValid && bValid) return aNum - bNum;
    if (aValid) return -1;
    if (bValid) return 1;
    return String(a.number).localeCompare(String(b.number), "it", { numeric: true });
  });
  saveState();
  renderTables();
  renderFloorMarkers();
}

function renumberTablesSmart() {
  if (!canEditTablesArea()) return;
  ensureFloorPlansState();
  const emptyTables = state.tables.filter((table) => guestCount(table) === 0);
  const willDelete = emptyTables.length;
  const beforeCount = state.tables.length;
  const afterCount = beforeCount - willDelete;
  const messageLines = [
    "Confermi la rinumerazione intelligente dei tavoli?",
    "",
    `- Tavoli attuali: ${beforeCount}`,
    `- Tavoli vuoti da eliminare: ${willDelete}`,
    `- Tavoli che resteranno: ${afterCount}`,
    "",
    "I tavoli rimanenti saranno rinumerati in ordine attuale (1, 2, 3, ...).",
    "Nomi/note tavolo resteranno invariati e i marker manterranno la posizione.",
  ];
  if (!window.confirm(messageLines.join("\n"))) return;

  if (willDelete > 0) {
    const emptyIds = new Set(emptyTables.map((t) => t.id));
    state.tables = state.tables.filter((t) => !emptyIds.has(t.id));
    for (const tableId of emptyIds) {
      for (const plan of state.floorPlans) {
        if (plan.markerPositions && plan.markerPositions[tableId]) {
          delete plan.markerPositions[tableId];
        }
        if (Array.isArray(plan.hiddenTableIds)) {
          plan.hiddenTableIds = plan.hiddenTableIds.filter((id) => id !== tableId);
        }
      }
    }
  }

  state.tables.forEach((table, index) => {
    table.number = index + 1;
  });

  saveState();
  renderTables();
  renderFloorMarkers();
}

function guestCount(table) {
  return table.guests.filter((g) => {
    const cognome = (g.cognome || "").trim();
    const nome = (g.nome || "").trim();
    return cognome || nome;
  }).length;
}

function addGuest(tableId) {
  if (!canEditTablesArea()) return;
  const table = state.tables.find((t) => t.id === tableId);
  if (!table) return;
  ensureTableGuestSlots(table);
  saveState();
  renderTables();
  renderFloorMarkers();
}

function removeGuest(tableId, index) {
  if (!canEditTablesArea()) return;
  const table = state.tables.find((t) => t.id === tableId);
  if (!table || !table.guests[index]) return;
  table.guests[index] = createEmptyGuest();
  saveState();
  renderTables();
  renderFloorMarkers();
}

function updateGuest(tableId, index, field, value) {
  if (!canEditTablesArea()) return;
  const table = state.tables.find((t) => t.id === tableId);
  if (!table || !table.guests[index]) return;
  table.guests[index][field] = value;
  saveState();
  refreshTableLiveInfoById(tableId);
  renderFloorMarkers();
  renderAcceptanceArea();
}

function updateTableMeta(tableId, field, value) {
  if (!canEditTablesArea()) return;
  const table = state.tables.find((t) => t.id === tableId);
  if (!table) return;
  table[field] = value;
  saveState();
  renderFloorMarkers();
  renderAcceptanceArea();
}

function guestDisplayName(guest, fallbackIndex) {
  const full = `${(guest.cognome || "").trim()} ${(guest.nome || "").trim()}`.trim();
  return full || `Ospite ${fallbackIndex + 1}`;
}

function getActiveGuests(table) {
  return table.guests.filter((g) => {
    const cognome = (g.cognome || "").trim();
    const nome = (g.nome || "").trim();
    return cognome || nome;
  });
}

function setGuestRowVisualState(rowEl, hasIdentity) {
  rowEl.classList.toggle("guest-row--filled", hasIdentity);
  rowEl.classList.toggle("guest-row--empty", !hasIdentity);
}

/**
 * Pallini solo se l’ospite ha cognome o nome:
 * - nessun pallino: adulto e senza allergie
 * - giallo: menù bambino
 * - rosso: testo allergie/intolleranze
 * - giallo + rosso affiancati: entrambi
 */
function updateGuestRowIndicatorDots(rowEl, hasIdentity, menu, note) {
  const grid = rowEl.querySelector(".guest-row__grid");
  const wrap = rowEl.querySelector(".guest-row__indicators");
  const dB = rowEl.querySelector(".guest-dot--bambino");
  const dA = rowEl.querySelector(".guest-dot--allergy");
  if (!dB || !dA || !grid || !wrap) return;

  const nt = String(note || "").trim();
  const m = String(menu || "adulto").toLowerCase() === "bambino" ? "bambino" : "adulto";

  const showYellow = Boolean(hasIdentity) && m === "bambino";
  const showRed = Boolean(hasIdentity) && nt.length > 0;
  const showStrip = showYellow || showRed;

  dB.toggleAttribute("hidden", !showYellow);
  dA.toggleAttribute("hidden", !showRed);

  dB.title = showYellow ? "Menù bambino" : "";
  dA.title = showRed ? `Allergie: ${nt.length > 72 ? `${nt.slice(0, 72)}…` : nt}` : "";

  grid.classList.toggle("guest-row__grid--with-markers", showStrip);
  wrap.hidden = !showStrip;
  wrap.setAttribute("aria-hidden", showStrip ? "false" : "true");
}

function syncGuestMenuPickFromValue(sheet, menuValue) {
  const v = menuValue === "bambino" ? "bambino" : "adulto";
  sheet.querySelectorAll(".guest-row__menu-opt").forEach((b) => {
    b.classList.toggle("is-selected", b.dataset.value === v);
  });
}

function applyGuestSheetIdentityUi(sheet, hasIdentity) {
  const hint = sheet.querySelector("[data-guest-sheet-hint]");
  const opts = sheet.querySelectorAll(".guest-row__menu-opt");
  const note = sheet.querySelector(".guest-note");
  if (hint) hint.hidden = hasIdentity;
  opts.forEach((o) => {
    o.disabled = !hasIdentity;
  });
  if (note) note.disabled = !hasIdentity;
}

function refreshBodyScrollForGuestSheets() {
  const anyOpen = Boolean(document.querySelector(".guest-row__sheet:not([hidden])"));
  document.body.style.overflow = anyOpen ? "hidden" : "";
}

function closeAllGuestOptionSheets() {
  document.querySelectorAll(".guest-row__sheet:not([hidden])").forEach((sheet) => {
    sheet.hidden = true;
    const row = sheet.closest(".guest-row");
    const btn = row?.querySelector(".guest-row__more");
    if (btn) btn.setAttribute("aria-expanded", "false");
  });
  refreshBodyScrollForGuestSheets();
}

function openGuestOptionSheet(sheetEl) {
  closeAllGuestOptionSheets();
  sheetEl.hidden = false;
  const row = sheetEl.closest(".guest-row");
  const btn = row?.querySelector(".guest-row__more");
  if (btn) btn.setAttribute("aria-expanded", "true");
  refreshBodyScrollForGuestSheets();
  requestAnimationFrame(() => {
    const ta = sheetEl.querySelector(".guest-note");
    if (ta && !ta.disabled) ta.focus();
    else {
      const firstOpt = sheetEl.querySelector(".guest-row__menu-opt:not(:disabled)");
      firstOpt?.focus();
    }
  });
}

function focusGuestField(card, rowIndex, colIndex) {
  const candidate = card.querySelector(
    `.guest-row[data-row-index="${rowIndex}"] [data-grid-col="${colIndex}"]`
  );
  if (!candidate) return false;
  if (candidate.disabled) return false;
  candidate.focus();
  if (candidate.tagName === "INPUT") {
    candidate.select();
  }
  return true;
}

function moveGuestGridFocus(card, fromRow, fromCol, key) {
  const vectors = {
    ArrowUp: [-1, 0],
    ArrowDown: [1, 0],
    ArrowLeft: [0, -1],
    ArrowRight: [0, 1],
  };
  const vector = vectors[key];
  if (!vector) return;

  const [dRow, dCol] = vector;
  let row = fromRow;
  let col = fromCol;

  for (let step = 0; step < MAX_GUESTS * 4; step += 1) {
    row += dRow;
    col += dCol;
    if (row < 0 || row >= MAX_GUESTS || col < 0 || col > 1) break;
    if (focusGuestField(card, row, col)) return;
  }
}

function moveGuestGridNextField(card, fromRow, fromCol) {
  const maxCols = 2;
  const startIndex = fromRow * maxCols + fromCol;
  for (let index = startIndex + 1; index < MAX_GUESTS * maxCols; index += 1) {
    const row = Math.floor(index / maxCols);
    const col = index % maxCols;
    if (focusGuestField(card, row, col)) return;
  }
}

function moveGuestGridPrevField(card, fromRow, fromCol) {
  const maxCols = 2;
  const startIndex = fromRow * maxCols + fromCol;
  for (let index = startIndex - 1; index >= 0; index -= 1) {
    const row = Math.floor(index / maxCols);
    const col = index % maxCols;
    if (focusGuestField(card, row, col)) return;
  }
}

function getTableMenuCounts(table) {
  const activeGuests = getActiveGuests(table);
  let adulti = 0;
  let bambini = 0;
  for (const guest of activeGuests) {
    if (guest.menu === "bambino") {
      bambini += 1;
    } else {
      adulti += 1;
    }
  }
  return { adulti, bambini };
}

/**
 * Elenco stringhe tipo "1A no lattosio", "2B no glutine" (raggruppate per testo nota + menù).
 */
function getTableAllergenParts(table) {
  const map = new Map();
  for (const guest of getActiveGuests(table)) {
    const note = (guest.note || "").trim();
    if (!note) continue;
    const menu = guest.menu === "bambino" ? "bambino" : "adulto";
    const key = `${menu}\0${note}`;
    map.set(key, (map.get(key) || 0) + 1);
  }
  const items = [];
  for (const [key, count] of map) {
    const [menu, text] = key.split("\0");
    const letter = menu === "bambino" ? "B" : "A";
    items.push({ menu, text, letter, count });
  }
  items.sort((a, b) => {
    if (a.menu !== b.menu) return a.menu === "adulto" ? -1 : 1;
    return a.text.localeCompare(b.text, "it");
  });
  return items.map((p) => `${p.count}${p.letter} ${p.text}`);
}

function getTableMenuCountString(table) {
  const c = getTableMenuCounts(table);
  return `${c.adulti}A+${c.bambini}B`;
}

/** Resoconto completo sotto al tavolo: 5A+2B (1A … | 1B …) — eventuali note tavolo */
function getTableSummaryLine(table) {
  const counts = getTableMenuCountString(table);
  const allergenParts = getTableAllergenParts(table);
  let line = counts;
  if (allergenParts.length) {
    line += ` (${allergenParts.join(" | ")})`;
  }
  const tableNote = (table.tableNote || "").trim();
  if (tableNote) {
    line += ` — Note tavolo: ${tableNote}`;
  }
  return line;
}

function getCurrentVisibleTable() {
  if (!els.tablesList) return null;
  const cards = [...els.tablesList.querySelectorAll(".table-card")];
  if (!cards.length) return null;
  const rootRect = els.tablesList.getBoundingClientRect();
  const anchorY = rootRect.top + rootRect.height * 0.32;
  let best = cards[0];
  let bestDist = Infinity;
  for (const card of cards) {
    const r = card.getBoundingClientRect();
    const center = r.top + r.height / 2;
    const d = Math.abs(center - anchorY);
    if (d < bestDist) {
      bestDist = d;
      best = card;
    }
  }
  const id = best.dataset.tableId;
  return state.tables.find((t) => t.id === id) || null;
}

/** Barra fissa in basso: sempre allineata al tavolo visibile nell’area scroll. */
function updateCurrentTableContextUi() {
  if (!state.tables.length) {
    if (els.globalOccupancyLine) {
      els.globalOccupancyLine.textContent = "Posti occupati: —";
      els.globalOccupancyLine.classList.remove("is-full");
    }
    if (els.globalSummaryLine) els.globalSummaryLine.textContent = "Resoconto piantina: —";
    return;
  }
  const table = getCurrentVisibleTable();
  if (!table) {
    if (els.globalOccupancyLine) {
      els.globalOccupancyLine.textContent = "Posti occupati: —";
      els.globalOccupancyLine.classList.remove("is-full");
    }
    if (els.globalSummaryLine) els.globalSummaryLine.textContent = "Resoconto piantina: —";
    return;
  }
  const name = (table.customName || "").trim();
  const count = guestCount(table);
  const capText =
    count >= MAX_GUESTS
      ? `Posti occupati: ${count}/${MAX_GUESTS} (massimo raggiunto)`
      : `Posti occupati: ${count}/${MAX_GUESTS}`;
  if (els.globalOccupancyLine) {
    els.globalOccupancyLine.textContent = `Tavolo ${table.number}${name ? ` — ${name}` : ""} · ${capText}`;
    els.globalOccupancyLine.classList.toggle("is-full", count >= MAX_GUESTS);
  }
  if (els.globalSummaryLine) {
    els.globalSummaryLine.textContent = `Resoconto piantina: ${getTableSummaryLine(table)}`;
  }
}

function closeAllDropdowns() {
  if (els.tableActionsMenu) {
    els.tableActionsMenu.hidden = true;
  }
  if (els.topActionsMenu) {
    els.topActionsMenu.hidden = true;
  }
  if (els.tableActionsMenuBtn) {
    els.tableActionsMenuBtn.setAttribute("aria-expanded", "false");
  }
  if (els.topActionsMenuBtn) {
    els.topActionsMenuBtn.setAttribute("aria-expanded", "false");
  }
  document.querySelectorAll(".top-actions-submenu").forEach((sub) => {
    sub.hidden = true;
  });
  document.querySelectorAll("[data-submenu-toggle]").forEach((btn) => {
    btn.setAttribute("aria-expanded", "false");
  });
  document.querySelectorAll(".plan-visibility-panel").forEach((panel) => {
    panel.hidden = true;
  });
  document.querySelectorAll("[data-plan-visibility-toggle]").forEach((btn) => {
    btn.setAttribute("aria-expanded", "false");
  });
}

function toggleDropdown(panel, button) {
  const willOpen = panel.hidden;
  closeAllDropdowns();
  panel.hidden = !willOpen;
  button.setAttribute("aria-expanded", willOpen ? "true" : "false");
  if (willOpen) {
    panel.scrollTop = 0;
  }
}

function setMainAreaView(view) {
  let targetView =
    view === "map" || view === "acceptance" || view === "segnatavolo" || view === "menu"
      ? view
      : "tables";
  if (!canAccessArea(targetView)) {
    targetView = canAccessArea("tables") ? "tables" : "acceptance";
  }
  const showTables = targetView === "tables";
  closeAllDropdowns();
  if (els.tablesAreaPanel) {
    els.tablesAreaPanel.hidden = !showTables;
  }
  if (els.mapAreaPanel) {
    els.mapAreaPanel.hidden = targetView !== "map";
  }
  if (els.acceptanceAreaPanel) {
    els.acceptanceAreaPanel.hidden = targetView !== "acceptance";
  }
  if (els.segnatavoloAreaPanel) {
    els.segnatavoloAreaPanel.hidden = targetView !== "segnatavolo";
  }
  if (els.menuBookletAreaPanel) {
    els.menuBookletAreaPanel.hidden = targetView !== "menu";
  }
  if (els.stickyToolbar) {
    els.stickyToolbar.hidden = !showTables;
  }
  if (els.showTablesAreaBtn) {
    const isActive = showTables;
    els.showTablesAreaBtn.classList.toggle("is-active", isActive);
    els.showTablesAreaBtn.setAttribute("aria-pressed", isActive ? "true" : "false");
  }
  if (els.showMapAreaBtn) {
    const isActive = targetView === "map";
    els.showMapAreaBtn.classList.toggle("is-active", isActive);
    els.showMapAreaBtn.setAttribute("aria-pressed", isActive ? "true" : "false");
  }
  if (els.showAcceptanceAreaBtn) {
    const isActive = targetView === "acceptance";
    els.showAcceptanceAreaBtn.classList.toggle("is-active", isActive);
    els.showAcceptanceAreaBtn.setAttribute("aria-pressed", isActive ? "true" : "false");
  }
}

function getCurrentMainAreaView() {
  if (els.mapAreaPanel && !els.mapAreaPanel.hidden) return "map";
  if (els.acceptanceAreaPanel && !els.acceptanceAreaPanel.hidden) return "acceptance";
  if (els.segnatavoloAreaPanel && !els.segnatavoloAreaPanel.hidden) return "segnatavolo";
  if (els.menuBookletAreaPanel && !els.menuBookletAreaPanel.hidden) return "menu";
  return "tables";
}

function isSegnatavoloOrMenuStudioView() {
  const v = getCurrentMainAreaView();
  return v === "segnatavolo" || v === "menu";
}

function setPanelEditable(rootEl, editable) {
  if (!rootEl) return;
  rootEl.querySelectorAll("input, textarea, select, button").forEach((el) => {
    if (el.dataset.roleAlwaysEnabled === "true") return;
    el.disabled = !editable;
  });
}

function applyRoleUi() {
  if (els.showMapAreaBtn) els.showMapAreaBtn.hidden = appRole === "cliente";
  if (els.showAcceptanceAreaBtn) els.showAcceptanceAreaBtn.hidden = appRole === "cliente";
  if (els.openSegnatavoloAreaBtn) els.openSegnatavoloAreaBtn.hidden = appRole === "cliente";
  if (els.openMenuBookletAreaBtn) els.openMenuBookletAreaBtn.hidden = appRole === "cliente";

  if (els.exportGuestsBtn) els.exportGuestsBtn.hidden = appRole !== "struttura";
  if (els.exportKitchenBtn) els.exportKitchenBtn.hidden = appRole !== "struttura";
  if (els.exportClientBtn) els.exportClientBtn.hidden = appRole === "staff";
  if (els.exportFullBtn) els.exportFullBtn.hidden = appRole !== "struttura";
  if (els.importClientBtn) els.importClientBtn.hidden = appRole === "staff";
  if (els.importFullBtn) els.importFullBtn.hidden = appRole === "cliente";

  if (els.eventName) els.eventName.disabled = !canEditEventName();
  setPanelEditable(els.tablesAreaPanel, canEditTablesArea());
  setPanelEditable(els.mapAreaPanel, canEditMapArea());
  setPanelEditable(els.segnatavoloAreaPanel, canEditStudioAreas());
  setPanelEditable(els.menuBookletAreaPanel, canEditStudioAreas());
  if (els.acceptanceSearchInput) {
    els.acceptanceSearchInput.dataset.roleAlwaysEnabled = "true";
    els.acceptanceSearchInput.disabled = false;
  }

  const current = getCurrentMainAreaView();
  if (!canAccessArea(current)) {
    setMainAreaView("tables");
  }
}

function waitForNextFrame() {
  return new Promise((resolve) => {
    requestAnimationFrame(() => requestAnimationFrame(resolve));
  });
}

function waitForFloorImageReady(timeoutMs = 1200, root = document) {
  return new Promise((resolve) => {
    const imgs = Array.from(root.querySelectorAll(".floor-img")).filter(
      (img) => !img.hidden && img.getAttribute("src")
    );
    if (!imgs.length) {
      resolve();
      return;
    }
    if (imgs.every((img) => img.complete)) {
      resolve();
      return;
    }
    let pending = imgs.length;
    let done = false;
    const finishOne = () => {
      if (done) return;
      pending -= 1;
      if (pending <= 0) {
        done = true;
        clearTimeout(timer);
        resolve();
      }
    };
    const timer = setTimeout(() => {
      if (done) return;
      done = true;
      resolve();
    }, timeoutMs);
    imgs.forEach((img) => {
      img.addEventListener("load", finishOne, { once: true });
      img.addEventListener("error", finishOne, { once: true });
    });
  });
}

/** Testo sotto il marcatore sulla piantina: solo allergie/note tavolo, senza ripetere 5A+2B */
function getTableMarkerNotesText(table) {
  const parts = [];
  const allergenParts = getTableAllergenParts(table);
  if (allergenParts.length) {
    parts.push(`(${allergenParts.join(" | ")})`);
  }
  const tableNote = (table.tableNote || "").trim();
  if (tableNote) {
    parts.push(`Note tavolo: ${tableNote}`);
  }
  return parts.join(" ");
}

function refreshTableLiveInfoById(tableId) {
  const table = state.tables.find((t) => t.id === tableId);
  if (!table) return;
  const card = els.tablesList.querySelector(`[data-table-id="${CSS.escape(tableId)}"]`);
  if (!card) return;
  const capEl = card.querySelector("[data-capacity]");
  if (capEl) {
    const count = guestCount(table);
    capEl.textContent =
      count >= MAX_GUESTS
        ? `Posti occupati: ${count}/${MAX_GUESTS} (massimo raggiunto)`
        : `Posti occupati: ${count}/${MAX_GUESTS}`;
    capEl.classList.toggle("is-full", count >= MAX_GUESTS);
  }
  const summaryEl = card.querySelector("[data-summary]");
  if (summaryEl) {
    summaryEl.textContent = `Resoconto piantina: ${getTableSummaryLine(table)}`;
  }
  updateCurrentTableContextUi();
}

function getSortedParticipants() {
  const rows = [];
  state.tables.forEach((table) => {
    const tableName = (table.customName || "").trim();
    table.guests.forEach((g) => {
      const cognome = (g.cognome || "").trim();
      const nome = (g.nome || "").trim();
      if (!cognome && !nome) return;
      rows.push({
        cognome,
        nome,
        tavoloNumero: table.number,
        tavoloNome: tableName,
      });
    });
  });
  rows.sort((a, b) => {
    const k1 = `${a.cognome} ${a.nome}`.trim().toLowerCase();
    const k2 = `${b.cognome} ${b.nome}`.trim().toLowerCase();
    return k1.localeCompare(k2, "it");
  });
  return rows;
}

function normalizeSearchText(value) {
  return String(value || "").trim().toLocaleLowerCase("it");
}

function renderAcceptanceArea() {
  if (!els.acceptanceResults || !els.acceptanceTablePreview) return;
  const query = normalizeSearchText(acceptanceUiState.query);
  const rows = [];
  state.tables.forEach((table) => {
    const tableName = (table.customName || "").trim();
    const tableLabel = tableName ? `Tavolo ${table.number} — ${tableName}` : `Tavolo ${table.number}`;
    table.guests.forEach((g) => {
      const cognome = String(g.cognome || "").trim();
      const nome = String(g.nome || "").trim();
      if (!cognome && !nome) return;
      const fullName = `${cognome} ${nome}`.trim();
      rows.push({
        fullName,
        cognome,
        nome,
        tableId: table.id,
        tableLabel,
      });
    });
  });
  rows.sort((a, b) => a.fullName.localeCompare(b.fullName, "it", { sensitivity: "base" }));

  const filtered = query
    ? rows.filter((row) => normalizeSearchText(`${row.fullName} ${row.tableLabel}`).includes(query))
    : rows;

  els.acceptanceResults.innerHTML = "";
  if (!filtered.length) {
    const empty = document.createElement("p");
    empty.className = "acceptance-empty";
    empty.textContent = "Nessun risultato per la ricerca.";
    els.acceptanceResults.appendChild(empty);
  } else {
    filtered.forEach((row) => {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "acceptance-result-item";
      btn.dataset.tableId = row.tableId;
      const isActive = acceptanceUiState.selectedTableId === row.tableId;
      if (isActive) btn.classList.add("is-active");
      btn.innerHTML = `<strong>${row.fullName}</strong><span>${row.tableLabel}</span>`;
      btn.addEventListener("click", () => {
        acceptanceUiState.selectedTableId = row.tableId;
        renderAcceptanceArea();
      });
      els.acceptanceResults.appendChild(btn);
    });
  }

  const validSelected = state.tables.some((t) => t.id === acceptanceUiState.selectedTableId);
  if (!validSelected) {
    acceptanceUiState.selectedTableId = filtered[0] ? filtered[0].tableId : "";
  }
  const selectedTable = state.tables.find((t) => t.id === acceptanceUiState.selectedTableId);
  els.acceptanceTablePreview.innerHTML = "";
  if (!selectedTable) {
    const empty = document.createElement("p");
    empty.className = "acceptance-empty";
    empty.textContent = "Seleziona un partecipante per vedere il tavolo completo.";
    els.acceptanceTablePreview.appendChild(empty);
    return;
  }

  const title = document.createElement("h3");
  const customName = String(selectedTable.customName || "").trim();
  title.className = "acceptance-preview-title";
  title.textContent = customName
    ? `Tavolo ${selectedTable.number} — ${customName}`
    : `Tavolo ${selectedTable.number}`;

  const summary = document.createElement("p");
  summary.className = "acceptance-preview-summary";
  summary.textContent = getTableSummaryLine(selectedTable);

  const list = document.createElement("ol");
  list.className = "acceptance-preview-list";
  getActiveGuests(selectedTable).forEach((guest) => {
    const li = document.createElement("li");
    const fullName = `${String(guest.cognome || "").trim()} ${String(guest.nome || "").trim()}`.trim();
    const menuLabel = guest.menu === "bambino" ? "Bambino" : "Adulto";
    const note = String(guest.note || "").trim();
    li.innerHTML = `<span>${fullName}</span><small>${menuLabel}${note ? ` • ${note}` : ""}</small>`;
    list.appendChild(li);
  });
  if (!list.children.length) {
    const li = document.createElement("li");
    li.className = "acceptance-empty";
    li.textContent = "Nessun partecipante inserito in questo tavolo.";
    list.appendChild(li);
  }

  els.acceptanceTablePreview.appendChild(title);
  els.acceptanceTablePreview.appendChild(summary);
  els.acceptanceTablePreview.appendChild(list);
}

function renderTables() {
  els.tablesList.innerHTML = "";
  const n = state.tables.length;
  els.tableCountLabel.textContent = n === 1 ? "1 tavolo" : `${n} tavoli`;

  for (const table of state.tables) {
    ensureTableGuestSlots(table);
    const node = els.tableCardTpl.content.cloneNode(true);
    const card = node.querySelector(".table-card");
    card.dataset.tableId = table.id;
    node.querySelector(".table-title").textContent = `Tavolo ${table.number}`;
    const summaryEl = node.querySelector("[data-summary]");
    const customNameInput = node.querySelector(".table-custom-name");
    const tableNoteInput = node.querySelector(".table-note");
    customNameInput.value = table.customName || "";
    tableNoteInput.value = table.tableNote || "";
    customNameInput.addEventListener("input", () => {
      updateTableMeta(table.id, "customName", customNameInput.value);
      summaryEl.textContent = `Resoconto piantina: ${getTableSummaryLine(table)}`;
      const ctx = card.querySelector("[data-guests-context]");
      if (ctx) {
        const cn = (table.customName || "").trim();
        ctx.textContent = cn ? `Tavolo ${table.number} — ${cn}` : `Tavolo ${table.number}`;
      }
      updateCurrentTableContextUi();
    });
    tableNoteInput.addEventListener("input", () => {
      updateTableMeta(table.id, "tableNote", tableNoteInput.value);
      summaryEl.textContent = `Resoconto piantina: ${getTableSummaryLine(table)}`;
      updateCurrentTableContextUi();
    });

    const guestsEl = node.querySelector("[data-guests]");
    const guestsContext = node.querySelector("[data-guests-context]");
    const cap = node.querySelector("[data-capacity]");
    if (guestsContext) {
      const cn = (table.customName || "").trim();
      guestsContext.textContent = cn ? `Tavolo ${table.number} — ${cn}` : `Tavolo ${table.number}`;
    }

    table.guests.forEach((g, idx) => {
      const row = els.guestRowTpl.content.cloneNode(true);
      const rowEl = row.querySelector(".guest-row");
      rowEl.dataset.rowIndex = String(idx);
      const sheet = rowEl.querySelector(".guest-row__sheet");
      const moreBtn = rowEl.querySelector(".guest-row__more");
      const cognome = rowEl.querySelector(".guest-cognome");
      const nome = rowEl.querySelector(".guest-nome");
      const menu = rowEl.querySelector(".guest-menu");
      const note = rowEl.querySelector(".guest-note");
      cognome.dataset.gridCol = "0";
      nome.dataset.gridCol = "1";
      cognome.value = g.cognome;
      nome.value = g.nome;
      menu.value = g.menu === "bambino" ? "bambino" : "adulto";
      note.value = g.note || "";
      syncGuestMenuPickFromValue(sheet, menu.value);

      const applyIdentityConstraints = () => {
        const hasIdentity = Boolean(cognome.value.trim() || nome.value.trim());
        menu.disabled = !hasIdentity;
        if (moreBtn) moreBtn.disabled = !hasIdentity;
        setGuestRowVisualState(rowEl, hasIdentity);
        applyGuestSheetIdentityUi(sheet, hasIdentity);
        if (!hasIdentity) {
          if (menu.value !== "adulto") {
            menu.value = "adulto";
            updateGuest(table.id, idx, "menu", "adulto");
          }
          if (note.value.trim()) {
            note.value = "";
            updateGuest(table.id, idx, "note", "");
          }
          syncGuestMenuPickFromValue(sheet, "adulto");
        }
        updateGuestRowIndicatorDots(rowEl, hasIdentity, menu.value, note.value);
      };

      applyIdentityConstraints();

      cognome.addEventListener("input", () => {
        updateGuest(table.id, idx, "cognome", cognome.value);
        applyIdentityConstraints();
      });
      nome.addEventListener("input", () => {
        updateGuest(table.id, idx, "nome", nome.value);
        applyIdentityConstraints();
      });

      menu.addEventListener("change", () => {
        updateGuest(table.id, idx, "menu", menu.value);
        syncGuestMenuPickFromValue(sheet, menu.value);
        updateGuestRowIndicatorDots(rowEl, Boolean(cognome.value.trim() || nome.value.trim()), menu.value, note.value);
      });

      note.addEventListener("input", () => {
        updateGuest(table.id, idx, "note", note.value);
        updateGuestRowIndicatorDots(rowEl, Boolean(cognome.value.trim() || nome.value.trim()), menu.value, note.value);
      });

      sheet.querySelectorAll(".guest-row__menu-opt").forEach((optBtn) => {
        optBtn.addEventListener("click", () => {
          if (optBtn.disabled) return;
          const val = optBtn.dataset.value === "bambino" ? "bambino" : "adulto";
          menu.value = val;
          menu.dispatchEvent(new Event("change", { bubbles: true }));
        });
      });

      if (moreBtn) {
        moreBtn.addEventListener("click", () => {
          if (moreBtn.disabled) return;
          if (!sheet.hidden) {
            closeAllGuestOptionSheets();
            return;
          }
          const hasIdentity = Boolean(cognome.value.trim() || nome.value.trim());
          applyGuestSheetIdentityUi(sheet, hasIdentity);
          syncGuestMenuPickFromValue(sheet, menu.value);
          openGuestOptionSheet(sheet);
        });
      }

      const scrim = sheet.querySelector(".guest-row__sheet-scrim");
      const doneBtn = sheet.querySelector(".guest-row__sheet-done");
      scrim?.addEventListener("click", () => closeAllGuestOptionSheets());
      doneBtn?.addEventListener("click", () => closeAllGuestOptionSheets());

      row.querySelector(".remove-guest").addEventListener("click", () => removeGuest(table.id, idx));

      [cognome, nome].forEach((fieldEl) => {
        fieldEl.addEventListener("keydown", (e) => {
          if (e.key === "Enter") {
            const currentRow = Number(fieldEl.closest(".guest-row")?.dataset.rowIndex);
            const currentCol = Number(fieldEl.dataset.gridCol);
            if (Number.isNaN(currentRow) || Number.isNaN(currentCol)) return;
            e.preventDefault();
            if (e.shiftKey) {
              moveGuestGridPrevField(card, currentRow, currentCol);
            } else {
              moveGuestGridNextField(card, currentRow, currentCol);
            }
            return;
          }
          if (
            e.key !== "ArrowUp" &&
            e.key !== "ArrowDown" &&
            e.key !== "ArrowLeft" &&
            e.key !== "ArrowRight"
          ) {
            return;
          }
          const currentRow = Number(fieldEl.closest(".guest-row")?.dataset.rowIndex);
          const currentCol = Number(fieldEl.dataset.gridCol);
          if (Number.isNaN(currentRow) || Number.isNaN(currentCol)) return;
          e.preventDefault();
          moveGuestGridFocus(card, currentRow, currentCol, e.key);
        });
      });

      guestsEl.appendChild(row);
    });

    const count = guestCount(table);
    cap.textContent =
      count >= MAX_GUESTS
        ? `Posti occupati: ${count}/${MAX_GUESTS} (massimo raggiunto)`
        : `Posti occupati: ${count}/${MAX_GUESTS}`;
    cap.classList.toggle("is-full", count >= MAX_GUESTS);
    summaryEl.textContent = `Resoconto piantina: ${getTableSummaryLine(table)}`;

    node.querySelector(".remove-table").addEventListener("click", () => removeTable(table.id));

    els.tablesList.appendChild(node);
  }

  updateCurrentTableContextUi();
  renderAcceptanceArea();
  requestAnimationFrame(() => {
    updateCurrentTableContextUi();
  });
  applyRoleUi();
}

function applyFloorPlanFromState() {
  renderFloorPlans();
}

function renderFloorPlans() {
  ensureFloorPlansState();
  if (!els.floorPlansContainer) return;
  els.floorPlansContainer.innerHTML = "";

  const openMenuForPlanId = preservePlanVisibilityMenuPlanId;
  preservePlanVisibilityMenuPlanId = null;

  state.floorPlans.forEach((plan, idx) => {
    const card = document.createElement("article");
    card.className = "floor-plan-card";
    card.dataset.planId = plan.id;

    const head = document.createElement("div");
    head.className = "floor-plan-card__head";
    const actions = document.createElement("div");
    actions.className = "floor-plan-card__actions";
    actions.innerHTML = `
      <label class="btn btn-secondary btn-sm file-label">
        Carica immagine
        <input type="file" class="plan-upload-input" accept="image/*" hidden />
      </label>
      <button type="button" class="btn btn-ghost btn-sm" data-plan-clear-btn="true">Rimuovi immagine</button>
      <button type="button" class="btn btn-danger btn-sm" data-plan-remove-btn="true">Elimina riquadro</button>
      <button type="button" class="btn btn-secondary btn-sm" data-plan-export-pdf="true">Esporta piantina PDF</button>
    `;

    const hiddenSet = new Set(plan.hiddenTableIds || []);
    const totalTables = state.tables.length;
    const shownTables = state.tables.reduce((acc, table) => (hiddenSet.has(table.id) ? acc : acc + 1), 0);
    const showingAllTables = shownTables === totalTables;

    const visWrap = document.createElement("div");
    visWrap.className = "plan-visibility";
    const visToggleBtn = document.createElement("button");
    visToggleBtn.type = "button";
    visToggleBtn.className = `btn btn-sm ${showingAllTables ? "btn-plan-visibility-all" : "btn-plan-visibility-partial"}`;
    visToggleBtn.dataset.planVisibilityToggle = "true";
    visToggleBtn.setAttribute("aria-expanded", "false");
    visToggleBtn.setAttribute("aria-haspopup", "true");
    visToggleBtn.textContent = showingAllTables
      ? `Tutti i ${totalTables} tavoli sono mostrati`
      : `Solo ${shownTables} tavoli di ${totalTables} tavoli sono mostrati`;

    const visPanel = document.createElement("div");
    visPanel.className = "plan-visibility-panel";
    visPanel.hidden = true;
    visPanel.setAttribute("role", "region");
    visPanel.setAttribute("aria-label", "Selezione tavoli su questo riquadro");

    const visTitle = document.createElement("p");
    visTitle.className = "plan-visibility-panel__title";
    visTitle.textContent = "Mostra o nascondi i tavoli su questa piantina";

    const visList = document.createElement("div");
    visList.className = "plan-visibility-panel__list";

    for (const table of state.tables) {
      const cn = String(table.customName || "").trim();
      const rowLabel = cn ? `Tavolo ${table.number} — ${cn}` : `Tavolo ${table.number}`;
      const label = document.createElement("label");
      label.className = "plan-visibility-row";
      const cb = document.createElement("input");
      cb.type = "checkbox";
      cb.className = "plan-visibility-checkbox";
      cb.dataset.tableId = table.id;
      cb.checked = !hiddenSet.has(table.id);
      const span = document.createElement("span");
      span.textContent = rowLabel;
      span.title = rowLabel;
      label.appendChild(cb);
      label.appendChild(span);
      visList.appendChild(label);
    }

    const visFooter = document.createElement("div");
    visFooter.className = "plan-visibility-actions";
    const showAllBtn = document.createElement("button");
    showAllBtn.type = "button";
    showAllBtn.className = "btn btn-ghost btn-sm plan-visibility-show-all";
    showAllBtn.dataset.planVisibilityShowAll = "true";
    showAllBtn.textContent = "Mostra tutti";
    visFooter.appendChild(showAllBtn);

    visPanel.appendChild(visTitle);
    visPanel.appendChild(visList);
    visPanel.appendChild(visFooter);
    visWrap.appendChild(visToggleBtn);
    visWrap.appendChild(visPanel);
    actions.prepend(visWrap);

    if (openMenuForPlanId === plan.id) {
      visPanel.hidden = false;
      visToggleBtn.setAttribute("aria-expanded", "true");
    }

    head.appendChild(actions);
    card.appendChild(head);

    const wrap = document.createElement("div");
    wrap.className = "floor-wrap";
    const stage = document.createElement("div");
    stage.className = "floor-stage";
    stage.dataset.planStage = plan.id;

    const img = document.createElement("img");
    img.className = "floor-img";
    img.alt = "Immagine piantina";
    if (plan.imageDataUrl) {
      img.src = plan.imageDataUrl;
      img.hidden = false;
    } else {
      img.hidden = true;
    }

    const hint = document.createElement("div");
    hint.className = "floor-hint";
    hint.hidden = Boolean(plan.imageDataUrl);
    hint.innerHTML = `
      <p>Nessuna immagine caricata.</p>
      <p class="small">Carica la piantina di questo riquadro e trascina i tavoli dove vuoi.</p>
    `;

    const markers = document.createElement("div");
    markers.className = "floor-markers";
    markers.dataset.planMarkers = plan.id;

    for (const table of state.tables) {
      if (hiddenSet.has(table.id)) continue;
      const pos = plan.markerPositions[table.id] || { x: 50, y: 50 };
      const marker = document.createElement("div");
      marker.className = "table-marker";
      marker.style.left = `${pos.x}%`;
      marker.style.top = `${pos.y}%`;
      marker.dataset.tableId = table.id;

      const circle = document.createElement("button");
      circle.type = "button";
      circle.className = "table-marker__circle";
      circle.textContent = String(table.number);
      if (canEditMapArea()) {
        circle.setAttribute("aria-label", `Tavolo ${table.number}, trascina sulla piantina`);
      } else {
        circle.setAttribute("aria-label", `Tavolo ${table.number}, sola visualizzazione`);
        circle.style.pointerEvents = "none";
        circle.style.cursor = "default";
      }

      const menus = document.createElement("div");
      menus.className = "table-marker__menus";
      const counts = getTableMenuCounts(table);
      menus.textContent = `${counts.adulti}A+${counts.bambini}B`;

      const notes = document.createElement("div");
      notes.className = "table-marker__notes";
      const markerNotes = getTableMarkerNotesText(table);
      notes.textContent = markerNotes;
      notes.hidden = !markerNotes;
      notes.title = getTableSummaryLine(table);

      marker.appendChild(circle);
      marker.appendChild(menus);
      marker.appendChild(notes);
      attachMarkerDrag(marker, circle, table.id, plan, markers);
      markers.appendChild(marker);
    }

    stage.appendChild(img);
    stage.appendChild(hint);
    stage.appendChild(markers);
    wrap.appendChild(stage);
    card.appendChild(wrap);
    els.floorPlansContainer.appendChild(card);
    layoutMarkerNotes(markers);
  });
  applyRoleUi();
}

function renderFloorMarkers() {
  renderFloorPlans();
}

function rectsOverlap(a, b) {
  return a.left < b.right && a.right > b.left && a.top < b.bottom && a.bottom > b.top;
}

function expandRect(rect, padding) {
  return {
    left: rect.left - padding,
    right: rect.right + padding,
    top: rect.top - padding,
    bottom: rect.bottom + padding,
  };
}

function markerHeaderRect(marker) {
  const circle = marker.querySelector(".table-marker__circle");
  const menus = marker.querySelector(".table-marker__menus");
  if (!circle) return null;

  const circleRect = circle.getBoundingClientRect();
  if (!menus) return circleRect;

  const menusRect = menus.getBoundingClientRect();
  return {
    left: Math.min(circleRect.left, menusRect.left),
    right: Math.max(circleRect.right, menusRect.right),
    top: Math.min(circleRect.top, menusRect.top),
    bottom: Math.max(circleRect.bottom, menusRect.bottom),
  };
}

function layoutMarkerNotes(markersRoot) {
  if (!markersRoot) return;
  const markers = Array.from(markersRoot.querySelectorAll(".table-marker"));
  if (!markers.length) return;

  // Reset defaults before recomputing.
  for (const marker of markers) {
    const note = marker.querySelector(".table-marker__notes");
    if (note) {
      note.style.setProperty("--note-left", "50%");
      note.style.setProperty("--note-top", "calc(100% + 4px)");
      note.style.setProperty("--note-translate-x", "-50%");
      note.style.setProperty("--note-translate-y", "0px");
      note.style.setProperty("--note-width", "180px");
    }
  }

  // Sort top-to-bottom, left-to-right for deterministic layout.
  markers.sort((a, b) => {
    const ta = parseFloat(a.style.top) || 0;
    const tb = parseFloat(b.style.top) || 0;
    if (ta !== tb) return ta - tb;
    const la = parseFloat(a.style.left) || 0;
    const lb = parseFloat(b.style.left) || 0;
    return la - lb;
  });

  const placedNotes = [];
  const blockers = markers
    .map((marker) => {
      const rect = markerHeaderRect(marker);
      if (!rect) return null;
      return { tableId: marker.dataset.tableId, rect };
    })
    .filter(Boolean);
  const gap = 8;
  const stageRect = markersRoot.getBoundingClientRect();
  const widthOptions = [200, 170, 145, 125, 110];
  const placements = [
    { key: "bottom", left: "50%", top: "calc(100% + 4px)", tx: "-50%", ty: "0px" },
    { key: "top", left: "50%", top: "-4px", tx: "-50%", ty: "-100%" },
    { key: "right", left: "calc(100% + 6px)", top: "50%", tx: "0%", ty: "-50%" },
    { key: "left", left: "-6px", top: "50%", tx: "-100%", ty: "-50%" },
  ];

  function scoreRect(rect, markerId) {
    let overlaps = 0;
    let edgePenalty = 0;

    for (const blocker of blockers) {
      if (blocker.tableId === markerId) continue;
      if (rectsOverlap(expandRect(rect, gap), blocker.rect)) overlaps += 3;
    }
    for (const p of placedNotes) {
      if (rectsOverlap(expandRect(rect, gap), p)) overlaps += 4;
    }

    if (rect.left < stageRect.left + 2) edgePenalty += stageRect.left + 2 - rect.left;
    if (rect.right > stageRect.right - 2) edgePenalty += rect.right - (stageRect.right - 2);
    if (rect.top < stageRect.top + 2) edgePenalty += stageRect.top + 2 - rect.top;
    if (rect.bottom > stageRect.bottom - 2) edgePenalty += rect.bottom - (stageRect.bottom - 2);

    return overlaps * 10000 + edgePenalty;
  }

  for (const marker of markers) {
    const markerId = marker.dataset.tableId;
    const note = marker.querySelector(".table-marker__notes");
    if (!note || note.hidden) continue;
    let best = null;

    for (const width of widthOptions) {
      note.style.setProperty("--note-width", `${width}px`);
      for (const place of placements) {
        note.style.setProperty("--note-left", place.left);
        note.style.setProperty("--note-top", place.top);
        note.style.setProperty("--note-translate-x", place.tx);
        note.style.setProperty("--note-translate-y", place.ty);
        const rect = note.getBoundingClientRect();
        const score = scoreRect(rect, markerId);
        if (!best || score < best.score) {
          best = { score, width, place };
        }
      }
    }

    if (best) {
      note.style.setProperty("--note-width", `${best.width}px`);
      note.style.setProperty("--note-left", best.place.left);
      note.style.setProperty("--note-top", best.place.top);
      note.style.setProperty("--note-translate-x", best.place.tx);
      note.style.setProperty("--note-translate-y", best.place.ty);
    }

    const finalRect = note.getBoundingClientRect();
    placedNotes.push(finalRect);
  }
}

function attachMarkerDrag(containerEl, handleEl, tableId, plan, markersRoot) {
  let startX, startY, origLeft, origTop, rect;

  function onPointerMove(e) {
    if (!canEditMapArea()) return;
    if (!rect) return;
    const dx = e.clientX - startX;
    const dy = e.clientY - startY;
    let x = origLeft + (dx / rect.width) * 100;
    let y = origTop + (dy / rect.height) * 100;
    x = Math.max(2, Math.min(98, x));
    y = Math.max(2, Math.min(98, y));
    containerEl.style.left = `${x}%`;
    containerEl.style.top = `${y}%`;
    layoutMarkerNotes(markersRoot);
  }

  function onPointerUp(e) {
    if (!canEditMapArea()) return;
    handleEl.releasePointerCapture(e.pointerId);
    handleEl.removeEventListener("pointermove", onPointerMove);
    handleEl.removeEventListener("pointerup", onPointerUp);
    handleEl.removeEventListener("pointercancel", onPointerUp);
    const left = parseFloat(containerEl.style.left);
    const top = parseFloat(containerEl.style.top);
    plan.markerPositions[tableId] = { x: left, y: top };
    saveState();
    rect = null;
    layoutMarkerNotes(markersRoot);
  }

  handleEl.addEventListener("pointerdown", (e) => {
    if (!canEditMapArea()) return;
    if (e.button !== 0) return;
    rect = markersRoot.getBoundingClientRect();
    startX = e.clientX;
    startY = e.clientY;
    origLeft = parseFloat(containerEl.style.left) || 50;
    origTop = parseFloat(containerEl.style.top) || 50;
    handleEl.setPointerCapture(e.pointerId);
    handleEl.addEventListener("pointermove", onPointerMove);
    handleEl.addEventListener("pointerup", onPointerUp);
    handleEl.addEventListener("pointercancel", onPointerUp);
    e.preventDefault();
  });
}

function readFileAsDataUrl(file) {
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve(r.result);
    r.onerror = () => reject(r.error);
    r.readAsDataURL(file);
  });
}

function createPdfDocument(orientation = "portrait") {
  if (!window.jspdf || !window.jspdf.jsPDF) return null;
  return new window.jspdf.jsPDF({ orientation, unit: "mm", format: "a4" });
}

const SEGNATAVOLO_STYLE_KEY = "evento-segnatavolo-style-v1";
const MENU_BOOKLET_KEY = "evento-menu-booklet-v1";
const SEGNATAVOLO_PAPER_FORMATS = {
  a5: [148, 210],
  a6: [105, 148],
  "10x15": [100, 150],
  "15x20": [150, 200],
};

/** 10 combinazioni tipografiche (jsPDF: helvetica | times | courier). */
const SEGNATAVOLO_THEMES = [
  { id: "t1", label: "Classico", titleFamily: "times", titleStyle: "bold", titleSize: 20, subFamily: "times", subStyle: "italic", subSize: 10, bodyFamily: "times", bodyStyle: "normal", bodySize: 11.5, footFamily: "times", footStyle: "normal", footSize: 7.5, previewTitle: "Georgia, 'Times New Roman', serif", previewBody: "Georgia, serif", titlePx: 1.35, bodyPx: 0.82 },
  { id: "t2", label: "Moderno", titleFamily: "helvetica", titleStyle: "bold", titleSize: 19, subFamily: "helvetica", subStyle: "normal", subSize: 9.5, bodyFamily: "helvetica", bodyStyle: "normal", bodySize: 11, footFamily: "helvetica", footStyle: "normal", footSize: 7, previewTitle: '"Source Sans 3", system-ui, sans-serif', previewBody: '"Source Sans 3", sans-serif', titlePx: 1.28, bodyPx: 0.8 },
  { id: "t3", label: "Elegante corsivo", titleFamily: "times", titleStyle: "italic", titleSize: 21, subFamily: "times", subStyle: "normal", subSize: 10, bodyFamily: "times", bodyStyle: "normal", bodySize: 11.5, footFamily: "times", footStyle: "italic", footSize: 7, previewTitle: "Georgia, serif", previewBody: "Georgia, serif", titlePx: 1.38, bodyPx: 0.82 },
  { id: "t4", label: "Macchina da scrivere", titleFamily: "courier", titleStyle: "bold", titleSize: 17, subFamily: "courier", subStyle: "normal", subSize: 9, bodyFamily: "courier", bodyStyle: "normal", bodySize: 10.5, footFamily: "courier", footStyle: "normal", footSize: 7, previewTitle: '"Courier New", Courier, monospace', previewBody: '"Courier New", monospace', titlePx: 1.15, bodyPx: 0.76 },
  { id: "t5", label: "Titolo serif + elenco sans", titleFamily: "times", titleStyle: "bold", titleSize: 20, subFamily: "times", subStyle: "italic", subSize: 10, bodyFamily: "helvetica", bodyStyle: "normal", bodySize: 11, footFamily: "helvetica", footStyle: "normal", footSize: 7, previewTitle: "Georgia, serif", previewBody: '"Source Sans 3", sans-serif', titlePx: 1.32, bodyPx: 0.78 },
  { id: "t6", label: "Sobrio compatto", titleFamily: "helvetica", titleStyle: "bold", titleSize: 16, subFamily: "helvetica", subStyle: "normal", subSize: 9, bodyFamily: "helvetica", bodyStyle: "normal", bodySize: 10.5, footFamily: "helvetica", footStyle: "normal", footSize: 6.8, previewTitle: '"Source Sans 3", sans-serif', previewBody: '"Source Sans 3", sans-serif', titlePx: 1.1, bodyPx: 0.74 },
  { id: "t7", label: "Grande titolo", titleFamily: "times", titleStyle: "bold", titleSize: 24, subFamily: "times", subStyle: "normal", subSize: 10.5, bodyFamily: "times", bodyStyle: "normal", bodySize: 12, footFamily: "times", footStyle: "normal", footSize: 7.5, previewTitle: "Georgia, serif", previewBody: "Georgia, serif", titlePx: 1.55, bodyPx: 0.85 },
  { id: "t8", label: "Contrasto", titleFamily: "helvetica", titleStyle: "bold", titleSize: 22, subFamily: "helvetica", subStyle: "normal", subSize: 9.5, bodyFamily: "times", bodyStyle: "normal", bodySize: 11.5, footFamily: "helvetica", footStyle: "normal", footSize: 7, previewTitle: '"Source Sans 3", sans-serif', previewBody: "Georgia, serif", titlePx: 1.42, bodyPx: 0.8 },
  { id: "t9", label: "Invito formale", titleFamily: "times", titleStyle: "bold", titleSize: 18, subFamily: "times", subStyle: "italic", subSize: 10, bodyFamily: "times", bodyStyle: "normal", bodySize: 11, footFamily: "times", footStyle: "italic", footSize: 7, previewTitle: "Georgia, serif", previewBody: "Georgia, serif", titlePx: 1.2, bodyPx: 0.8 },
  { id: "t10", label: "Leggero aria", titleFamily: "helvetica", titleStyle: "normal", titleSize: 18, subFamily: "helvetica", subStyle: "bold", subSize: 9, bodyFamily: "helvetica", bodyStyle: "normal", bodySize: 11, footFamily: "helvetica", footStyle: "normal", footSize: 7, previewTitle: '"Source Sans 3", sans-serif', previewBody: '"Source Sans 3", sans-serif', titlePx: 1.18, bodyPx: 0.8 },
];

const SEGNATAVOLO_PALETTES = [
  { id: "p1", label: "Antracite caldo", title: [44, 38, 32], subtitle: [105, 95, 85], body: [50, 44, 38], foot: [130, 120, 110], deco: [195, 175, 145] },
  { id: "p2", label: "Blu notte", title: [28, 42, 68], subtitle: [70, 85, 115], body: [38, 48, 72], foot: [100, 110, 130], deco: [140, 160, 195] },
  { id: "p3", label: "Borgogna", title: [92, 28, 38], subtitle: [130, 70, 80], body: [75, 35, 45], foot: [130, 110, 115], deco: [200, 160, 165] },
  { id: "p4", label: "Verde bosco", title: [32, 58, 42], subtitle: [70, 100, 78], body: [40, 62, 48], foot: [105, 125, 110], deco: [160, 190, 165] },
  { id: "p5", label: "Oro su carta", title: [75, 52, 18], subtitle: [120, 95, 55], body: [48, 42, 32], foot: [125, 115, 95], deco: [190, 155, 95] },
  { id: "p6", label: "Lavanda", title: [72, 58, 95], subtitle: [110, 95, 130], body: [65, 55, 85], foot: [120, 110, 135], deco: [175, 165, 205] },
  { id: "p7", label: "Ardesia", title: [48, 55, 62], subtitle: [95, 100, 108], body: [55, 60, 68], foot: [115, 118, 125], deco: [155, 165, 175] },
  { id: "p8", label: "Cioccolato", title: [78, 48, 32], subtitle: [120, 85, 65], body: [70, 50, 38], foot: [130, 115, 105], deco: [200, 170, 140] },
  { id: "p9", label: "Nero elegante", title: [22, 20, 18], subtitle: [85, 80, 75], body: [38, 36, 34], foot: [115, 110, 105], deco: [75, 72, 68] },
  { id: "p10", label: "Colori personalizzati", custom: true, title: [60, 42, 20], subtitle: [100, 85, 60], body: [44, 38, 32], foot: [120, 110, 100], deco: [196, 163, 90] },
];

const SEGNATAVOLO_GRAPHICS = [
  { id: 0, label: "Nessuna" },
  { id: 1, label: "Doppia cornice" },
  { id: 2, label: "Angoli a L" },
  { id: 3, label: "Fasce alto/basso" },
  { id: 4, label: "Cerchietti angolo" },
  { id: 5, label: "Diagonali angolo" },
  { id: 6, label: "Trama punti (margini)" },
  { id: 7, label: "Colonne laterali" },
  { id: 8, label: "Losanga in alto" },
  { id: 9, label: "Tratteggio fascia" },
];

/**
 * Preset per occasioni: applicano insieme tema tipografico, palette e grafica marginale.
 * Puoi cambiare ogni voce dopo; il chip resta evidenziato solo se la tripla coincide ancora.
 */
const SEGNATAVOLO_OCCASIONS = [
  { id: "none", label: "Generica", themeId: "t1", paletteId: "p1", graphicIndex: 0 },
  { id: "battesimo", label: "Battesimo", themeId: "t2", paletteId: "p1", graphicIndex: 1 },
  { id: "comunione", label: "Comunione", themeId: "t9", paletteId: "p5", graphicIndex: 1 },
  { id: "matrimonio", label: "Matrimonio", themeId: "t1", paletteId: "p3", graphicIndex: 1 },
  { id: "diciottesimo", label: "Diciottesimo", themeId: "t8", paletteId: "p3", graphicIndex: 9 },
  { id: "compleanno", label: "Compleanno", themeId: "t10", paletteId: "p7", graphicIndex: 3 },
  { id: "laurea", label: "Laurea", themeId: "t7", paletteId: "p9", graphicIndex: 8 },
];

const DEFAULT_SEGNATAVOLO_SETTINGS = {
  themeId: "t1",
  paletteId: "p1",
  graphicIndex: 0,
  /** Illustrazioni tematiche ai margini (null = solo grafica geometrica sotto). */
  festiveMarginId: null,
  customTitleHex: "#3c2614",
  customBodyHex: "#2c2620",
  customDecoHex: "#c4a35a",
  paperFormat: "a5",
  paperOrientation: "portrait",
  headerMode: "both",
  showGuests: true,
};
let segnatavoloSettings = { ...DEFAULT_SEGNATAVOLO_SETTINGS };

const MENU_BOOKLET_DISHES = [
  { category: "Show cooking", text: "SHOW COOKING\nFRIGGITORE\nSALTAPASTA\nBRACE" },
  { category: "Crudi di mare", text: "CRUDI DI MARE\nQUANTITA\nTIPOLOGIA DI SERVIZIO" },
  { category: "Aperitivi", text: "BOLLICINE E SUCCHI NATURALI\nDEGUSTAZIONE FINGER FOOD" },
  { category: "Aperitivi", text: "SPRITZ, BOLLICINE E SUCCHI NATURALI\nDEGUSTAZIONE FINGER FOOD" },
  { category: "Antipasti freddi", text: "POLPO SU FRESELLINA CREMOSA, POMODORO E POLVERE DI OLIVE DISIDRATATE\nSALMONE AL SALE PROFUMATO ALLE ERBE FINI\nJULIENNE DI SEPPIE CON VERDURINE IN AGRODOLCE\nGAMBERI MARINATI AGLI AGRUMI CON MAYO DELICATA AL LIME" },
  { category: "Antipasti freddi", text: "JULIENNE DI SEPPIA CON PESTO GENTILE DI BASILICO, PEPERONI E MANDORLE TOSTATE\nPESCE SPADA AL SIDRO DI MELE\nINSALATINA DI MAZZANCOLLE\nSALMONE AL SALE CON AGRUMI E CANNELLA" },
  { category: "Antipasti caldi", text: "TOTANO RIPIENO ALLA NAPOLETANA CON DRIPPING DI POMODORO, OLIO AL PREZZEMOLO E BIANCO DI BUFALA" },
  { category: "Antipasti caldi", text: "POLPO ALLA BRACE CON SIFONATA DI SEDANO RAPA, TARALLO E GOCCE DI PATATA VIOLA" },
  { category: "Antipasti caldi", text: "FILETTO DI BACCALA IN CROSTA DI CHIPS ALLA PAPRIKA, CREMOSO DI PATATA VIOLA E PEPERONI ALLA PARTENOPEA" },
  { category: "Primi di mare", text: "MEZZA MANICA ALLA NERANO CON PESCATRICE, VONGOLE E POLVERE DI POMODORO" },
  { category: "Primi di mare", text: "DITALONE CON CALAMARO, COZZE E PECORINO AL BASILICO" },
  { category: "Primi di mare", text: "VESUVIANA CON VONGOLE, EMULSIONE DI DATTERINO GIALLO E TARTARE DI TONNO ROSSO" },
  { category: "Primi di mare", text: "TORTELLO DEL PRETE\nRICOTTA SALATA, PATATA DOLCE E FIORE DI ZUCCA\nSCAMPI IN DOPPIA CONSISTENZA\nSALSA DI PROVOLA AFFUMICATA" },
  { category: "Primi di mare", text: "MESCAFANCRESCA DI PATATE E MARE" },
  { category: "Primi di mare", text: "CALAMARATA SBAGLIATA CON CANNELLINI, COZZE, SALSA AL RICCIO E SALICORNIA IN DOPPIA CONSISTENZA" },
  { category: "Primi di mare", text: "MEZZO PACCHERO CON BISQUE DI CROSTACEI, MAZZANCOLLE, PESCATRICE E DATTERINO" },
  { category: "Primi di mare", text: "RISOTTO AGLI AGRUMI CON GAMBERI E STRACCIATA DI BUFALA" },
  { category: "Primi di mare", text: "RISOTTO IN ASSOLUTO DI MARE" },
  { category: "Primi di terra", text: "RISOTTO MANTECATO CON PROVOLA, SPECK E TERRA DI TARALLO AL CAFFE" },
  { category: "Primi di terra", text: "CARNAROLI CON PISTILLI DI ZAFFERANO, SPUMA DI MOZZARELLA E BATTUTA DI MANZO" },
  { category: "Primi di terra", text: "RISOTTO COMM'E NA CAS E OV\nPISELLI IN TRE COTTURE\nUOVO MARINATO\nGUANCIALE CROCCANTE\nSALSA AL PECORINO" },
  { category: "Primi di terra", text: "CALAMARATA SBAGLIATA COMM'E NA CAS E OV\nPISELLI IN TRE COTTURE\nUOVO MARINATO\nGUANCIALE CROCCANTE\nSALSA AL PECORINO" },
  { category: "Secondi di mare", text: "SOVRAPPOSTO DI ORATA RIPIENO DI GAMBERI E PROVOLA\nCHIPS DI SCAPECE\nCREMOSO DI ZUCCHINE" },
  { category: "Secondi di mare", text: "RICCIOLA NELL'ORTO\nBIETA RIPASSATA\nPOMODORINI MULTICOLOR\nCREMOSO FRESCO DI PATATE ALL'INSALATA" },
  { category: "Secondi di mare", text: "OMBRINA SCOTTATA CON FINOCCHIO SPADELLATO\nCREMOSO DI CAVOLFIORE AL CARAMELLO SALATO\nSFERA DI FIORE DI ZUCCA RIPIENA DI ACCIUGHE E RICOTTA DI BUFALA" },
  { category: "Secondi di terra", text: "VITELLO IN CBT (24H 68°)\nFONDO DI COTTURA\nVERDURE AL BURRO" },
  { category: "Secondi di terra", text: "COPPA DI MAIALE IN CBT\nCAPONATA DI VERDURE\nSALSA DI PEPERONI PICCANTI\nCREMOSO DI BUFALA" },
  { category: "Frutta, dolci e torta", text: "SFERA DI CHEESECAKE\nGELEE DI FRUTTI ROSSI\nLEMON CURD\nCRUMBLE ALLE MANDORLE\nCUBETTATA DI ANANAS E FRUTTI ROSSI" },
  { category: "Frutta, dolci e torta", text: "BUFFET DI FRUTTA E DOLCI\nTORTA" },
  { category: "Frutta, dolci e torta", text: "BUFFET DI DOLCI\nTORTA" },
];

const menuBookletState = {
  logoDataUrl: "",
  items: [],
};

const menuBookletUiState = {
  mode: null, // "add" | "edit"
  targetId: null,
  draftType: "dish",
};

function hexToRgb(hex) {
  const m = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex.trim());
  if (!m) return [44, 38, 32];
  return [parseInt(m[1], 16), parseInt(m[2], 16), parseInt(m[3], 16)];
}

function rgbToCss(rgb) {
  return `rgb(${rgb[0]}, ${rgb[1]}, ${rgb[2]})`;
}

function getDetectedSegnatavoloOccasionId() {
  const { themeId, paletteId, graphicIndex } = segnatavoloSettings;
  for (const o of SEGNATAVOLO_OCCASIONS) {
    if (o.id === "none") continue;
    if (o.themeId === themeId && o.paletteId === paletteId && o.graphicIndex === graphicIndex) return o.id;
  }
  return "none";
}

function getSegnatavoloOccasionChipId() {
  if (segnatavoloSettings.festiveMarginId) return segnatavoloSettings.festiveMarginId;
  return getDetectedSegnatavoloOccasionId();
}

function getSegnatavoloPaperSizeMm() {
  const fmt = String(segnatavoloSettings.paperFormat || "a5").toLowerCase();
  const base = SEGNATAVOLO_PAPER_FORMATS[fmt] || SEGNATAVOLO_PAPER_FORMATS.a5;
  const orientation = segnatavoloSettings.paperOrientation === "landscape" ? "landscape" : "portrait";
  return orientation === "landscape" ? [base[1], base[0]] : [base[0], base[1]];
}

function getSegnatavoloPdfFormatAndKey() {
  const key = SEGNATAVOLO_PAPER_FORMATS[segnatavoloSettings.paperFormat]
    ? segnatavoloSettings.paperFormat
    : "a5";
  return { key, format: SEGNATAVOLO_PAPER_FORMATS[key] };
}

/** Come in CSS/stampa: 72 pt = 1 pollice = 96 px. Allinea l’anteprima ai punti usati da jsPDF nel PDF. */
function segnatavoloPdfPtToCssPx(pt) {
  return (Number(pt) || 0) * (96 / 72);
}

/** Conversione geometria foglio: millimetri -> CSS px (96px = 1in). */
function segnatavoloMmToCssPx(mm) {
  return (Number(mm) || 0) * (96 / 25.4);
}

/** Punti tipografici effettivi (stesso `computeSegnatavoloBatchLayout` del PDF). */
function getSegnatavoloPreviewTypePts(theme, batchLayout) {
  const bl = batchLayout;
  const body = bl ? bl.bodySize : theme.bodySize;
  return {
    titlePt: bl ? bl.titleSize : theme.titleSize,
    subPt: bl ? bl.subSize : theme.subSize,
    bodyPt: body,
    emptyBodyPt: Math.max(5.5, body * 0.88),
    footPt: theme.footSize,
  };
}

function getSegnatavoloHeaderParts(table, settingsOverride) {
  const S = settingsOverride || segnatavoloSettings;
  const custom = String(table.customName || "").trim();
  const mode = S.headerMode || "both";
  if (mode === "name") {
    if (custom) return { mainTitle: custom, subtitle: "" };
    return { mainTitle: `Tavolo ${table.number}`, subtitle: "" };
  }
  if (mode === "number") {
    return { mainTitle: String(table.number), subtitle: "" };
  }
  if (custom) {
    return { mainTitle: custom, subtitle: `Tavolo ${table.number}` };
  }
  return { mainTitle: `Tavolo ${table.number}`, subtitle: "" };
}

/** Margine di sicurezza stampa: 10% su ogni lato (stesso valore per anteprima e PDF). */
const SEGNATAVOLO_SAFE_MARGIN_FRAC = 0.1;

function segnatavoloTitleLineHeightMm(fontSizePt) {
  return Math.max(5.2, fontSizePt * 0.42);
}

function measureSegnatavoloListHeaderMm(pdf, theme, innerW, mainTitle, subtitle, titleSizePt, subSizePt) {
  pdf.setFont(theme.titleFamily, theme.titleStyle);
  pdf.setFontSize(titleSizePt);
  const titleLines = pdf.splitTextToSize(String(mainTitle || ""), innerW);
  const lhT = segnatavoloTitleLineHeightMm(titleSizePt);
  let h = titleLines.length * lhT;
  if (subtitle) {
    pdf.setFont(theme.subFamily, theme.subStyle);
    pdf.setFontSize(subSizePt);
    const subLines = pdf.splitTextToSize(String(subtitle), innerW);
    const lhS = segnatavoloTitleLineHeightMm(subSizePt);
    h += 0.8 + subLines.length * lhS;
  } else {
    h += 2.2;
  }
  h += 3.2;
  h += 1.2;
  h += 5.5;
  return h;
}

function measureSegnatavoloHeroHeaderMm(pdf, theme, innerW, mainTitle, subtitle, titleSizePt, subSizePt) {
  pdf.setFont(theme.titleFamily, theme.titleStyle);
  pdf.setFontSize(titleSizePt);
  const titleLines = pdf.splitTextToSize(String(mainTitle || ""), innerW);
  const lhT = segnatavoloTitleLineHeightMm(titleSizePt);
  let h = titleLines.length * lhT;
  if (subtitle) {
    pdf.setFont(theme.subFamily, theme.subStyle);
    pdf.setFontSize(subSizePt);
    const subLines = pdf.splitTextToSize(String(subtitle), innerW);
    const lhS = segnatavoloTitleLineHeightMm(subSizePt);
    h += 1.2 + subLines.length * lhS;
  }
  return h;
}

function measureSegnatavoloNamesBlockMm(pdf, theme, innerW, names, bodySizePt) {
  if (!names.length) {
    pdf.setFont(theme.bodyFamily, "italic");
    const sz = Math.max(5.5, bodySizePt * 0.88);
    pdf.setFontSize(sz);
    const lh = Math.max(4.2, sz * 0.4);
    const lines = pdf.splitTextToSize("Nessun ospite in elenco.", innerW);
    return lines.length * lh;
  }
  pdf.setFont(theme.bodyFamily, theme.bodyStyle);
  pdf.setFontSize(bodySizePt);
  const lh = Math.max(1.9, bodySizePt * 0.34);
  let total = 0;
  for (const n of names) {
    const lines = pdf.splitTextToSize(String(n), innerW);
    total += Math.max(1, lines.length) * lh;
  }
  return total;
}

/**
 * Calcolo unico per tutto il set: stessa scala tipografica su ogni segnatavolo (anteprima = PDF).
 * @returns {{ marginX: number, marginY: number, innerW: number, innerTop: number, contentBottom: number, contentH: number, footY: number, scale: number, titleSize: number, subSize: number, bodySize: number, heroMode: boolean, showGuests: boolean }}
 */
function computeSegnatavoloBatchLayout(pdf, tables, settings, theme, W, H) {
  const safe = SEGNATAVOLO_SAFE_MARGIN_FRAC;
  const marginX = W * safe;
  const marginY = H * safe;
  const innerW = Math.max(36, W - 2 * marginX);
  const innerTop = marginY;
  const innerBottom = H - marginY;
  const footBandMm = 8.5;
  const footY = innerBottom - 2.5;
  const contentBottom = footY - footBandMm;
  const contentH = Math.max(28, contentBottom - innerTop);

  const showGuests = Boolean(settings.showGuests);
  const list = Array.isArray(tables) ? tables : [];

  function fitsListScale(scale) {
    const titleSize = theme.titleSize * scale;
    const subSize = theme.subSize * scale;
    const bodySize = theme.bodySize * scale;
    for (const table of list) {
      const hp = getSegnatavoloHeaderParts(table, settings);
      const hh = measureSegnatavoloListHeaderMm(pdf, theme, innerW, hp.mainTitle, hp.subtitle, titleSize, subSize);
      const names = showGuests ? getTableGuestNamesOrdered(table) : [];
      const nh = showGuests ? measureSegnatavoloNamesBlockMm(pdf, theme, innerW, names, bodySize) : 0;
      if (hh + nh > contentH - 0.35) return false;
    }
    return true;
  }

  function fitsHeroScale(scale) {
    const titleSize = theme.titleSize * scale;
    const subSize = theme.subSize * scale;
    for (const table of list) {
      const hp = getSegnatavoloHeaderParts(table, settings);
      const hh = measureSegnatavoloHeroHeaderMm(pdf, theme, innerW, hp.mainTitle, hp.subtitle, titleSize, subSize);
      if (hh > contentH - 0.35) return false;
    }
    return true;
  }

  function maxScale(low, high, fits) {
    let best = low;
    for (let s = high; s >= low - 1e-6; s -= 0.025) {
      if (fits(s)) {
        best = s;
        break;
      }
    }
    return Math.max(low, best);
  }

  const heroMode = !showGuests;
  const hi = 2.6;
  const lo = 0.12;
  const scale = heroMode ? maxScale(lo, hi, fitsHeroScale) : maxScale(lo, hi, fitsListScale);

  return {
    marginX,
    marginY,
    innerW,
    innerTop,
    innerBottom,
    contentBottom,
    contentH,
    footY,
    scale,
    titleSize: Math.max(4, theme.titleSize * scale),
    subSize: Math.max(3.5, theme.subSize * scale),
    bodySize: Math.max(1.8, theme.bodySize * scale),
    heroMode,
    showGuests,
  };
}

/** Stesso calcolo del PDF su tutti i tavoli ordinati (per anteprima WYSIWYG). */
function getSegnatavoloLayoutForCurrentState(tablesForLayout) {
  const theme = getSegnatavoloTheme();
  const ordered =
    tablesForLayout || [...state.tables].sort((a, b) => Number(a.number) - Number(b.number));
  if (!ordered.length || !window.jspdf || !window.jspdf.jsPDF) return null;
  const orientation = segnatavoloSettings.paperOrientation === "landscape" ? "landscape" : "portrait";
  const { format } = getSegnatavoloPdfFormatAndKey();
  const pdf = new window.jspdf.jsPDF({ orientation, unit: "mm", format });
  const W = pdf.internal.pageSize.getWidth();
  const H = pdf.internal.pageSize.getHeight();
  return computeSegnatavoloBatchLayout(
    pdf,
    ordered,
    { headerMode: segnatavoloSettings.headerMode, showGuests: Boolean(segnatavoloSettings.showGuests) },
    theme,
    W,
    H
  );
}

function loadSegnatavoloSettings() {
  try {
    const raw = localStorage.getItem(SEGNATAVOLO_STYLE_KEY);
    if (raw) {
      const o = JSON.parse(raw);
      if (o.themeId) segnatavoloSettings.themeId = o.themeId;
      if (o.paletteId) segnatavoloSettings.paletteId = o.paletteId;
      if (typeof o.graphicIndex === "number") segnatavoloSettings.graphicIndex = o.graphicIndex;
      if (o.customTitleHex) segnatavoloSettings.customTitleHex = o.customTitleHex;
      if (o.customBodyHex) segnatavoloSettings.customBodyHex = o.customBodyHex;
      if (o.customDecoHex) segnatavoloSettings.customDecoHex = o.customDecoHex;
      if (Object.prototype.hasOwnProperty.call(o, "festiveMarginId")) {
        const fm = o.festiveMarginId;
        segnatavoloSettings.festiveMarginId =
          fm === null || fm === "" || fm === "none" ? null : String(fm);
      }
      if (o.paperFormat && SEGNATAVOLO_PAPER_FORMATS[String(o.paperFormat).toLowerCase()]) {
        segnatavoloSettings.paperFormat = String(o.paperFormat).toLowerCase();
      }
      if (o.paperOrientation === "portrait" || o.paperOrientation === "landscape") {
        segnatavoloSettings.paperOrientation = o.paperOrientation;
      }
      if (o.headerMode === "both" || o.headerMode === "name" || o.headerMode === "number") {
        segnatavoloSettings.headerMode = o.headerMode;
      }
      if (typeof o.showGuests === "boolean") {
        segnatavoloSettings.showGuests = o.showGuests;
      }
      if (
        segnatavoloSettings.festiveMarginId &&
        !SEGNATAVOLO_OCCASIONS.some((oc) => oc.id === segnatavoloSettings.festiveMarginId)
      ) {
        segnatavoloSettings.festiveMarginId = null;
      }
    }
  } catch (_) {}
  if (els.segnatavoloColorTitle) {
    els.segnatavoloColorTitle.value = segnatavoloSettings.customTitleHex;
    els.segnatavoloColorBody.value = segnatavoloSettings.customBodyHex;
    els.segnatavoloColorDeco.value = segnatavoloSettings.customDecoHex;
  }
  if (els.segnatavoloPaperFormat) els.segnatavoloPaperFormat.value = segnatavoloSettings.paperFormat;
  if (els.segnatavoloOrientation) els.segnatavoloOrientation.value = segnatavoloSettings.paperOrientation;
  if (els.segnatavoloHeaderMode) els.segnatavoloHeaderMode.value = segnatavoloSettings.headerMode;
  if (els.segnatavoloShowGuests) els.segnatavoloShowGuests.checked = Boolean(segnatavoloSettings.showGuests);
}

function saveSegnatavoloSettings() {
  if (els.segnatavoloColorTitle) {
    segnatavoloSettings.customTitleHex = els.segnatavoloColorTitle.value;
    segnatavoloSettings.customBodyHex = els.segnatavoloColorBody.value;
    segnatavoloSettings.customDecoHex = els.segnatavoloColorDeco.value;
  }
  if (els.segnatavoloPaperFormat) {
    const fmt = String(els.segnatavoloPaperFormat.value || "a5").toLowerCase();
    segnatavoloSettings.paperFormat = SEGNATAVOLO_PAPER_FORMATS[fmt] ? fmt : "a5";
  }
  if (els.segnatavoloOrientation) {
    segnatavoloSettings.paperOrientation =
      els.segnatavoloOrientation.value === "landscape" ? "landscape" : "portrait";
  }
  if (els.segnatavoloHeaderMode) {
    const mode = els.segnatavoloHeaderMode.value;
    segnatavoloSettings.headerMode = mode === "name" || mode === "number" ? mode : "both";
  }
  if (els.segnatavoloShowGuests) {
    segnatavoloSettings.showGuests = Boolean(els.segnatavoloShowGuests.checked);
  }
  try {
    localStorage.setItem(SEGNATAVOLO_STYLE_KEY, JSON.stringify(segnatavoloSettings));
  } catch (_) {}
}

function normalizeMenuBookletItems(items) {
  if (!Array.isArray(items)) return [];
  return items
    .filter(Boolean)
    .map((it) => ({
      id: it.id || uid(),
      type: it.type === "separator_short" ? "separator_short" : it.type === "separator_long" || it.type === "separator" ? "separator_long" : "dish",
      text: String(it.text || "").trim(),
    }))
    .filter((it) => it.text);
}

function loadMenuBookletState() {
  try {
    const raw = localStorage.getItem(MENU_BOOKLET_KEY);
    if (!raw) return;
    const o = JSON.parse(raw);
    if (o.logoDataUrl) menuBookletState.logoDataUrl = String(o.logoDataUrl);
    menuBookletState.items = normalizeMenuBookletItems(o.items);
  } catch (_) {
    menuBookletState.items = [];
  }
}

function saveMenuBookletState() {
  try {
    localStorage.setItem(
      MENU_BOOKLET_KEY,
      JSON.stringify({
        logoDataUrl: menuBookletState.logoDataUrl,
        items: menuBookletState.items,
      })
    );
  } catch (_) {}
}

function populateMenuBookletDishSelect() {
  if (!els.menuBookletDishSelect) return;
  els.menuBookletDishSelect.innerHTML = "";
  let currentCategory = "";
  let currentGroup = null;
  MENU_BOOKLET_DISHES.forEach((dish, idx) => {
    if (dish.category !== currentCategory) {
      currentCategory = dish.category;
      const grp = document.createElement("optgroup");
      grp.label = currentCategory;
      els.menuBookletDishSelect.appendChild(grp);
      currentGroup = grp;
    }
    const target = currentGroup || els.menuBookletDishSelect;
    const opt = document.createElement("option");
    opt.value = `dish:${idx}`;
    const preview = dish.text.split("\n")[0];
    opt.textContent = preview.length > 70 ? `${preview.slice(0, 70)}...` : preview;
    target.appendChild(opt);
  });
  const toolsGroup = document.createElement("optgroup");
  toolsGroup.label = "Elementi rapidi";
  [
    { value: "custom", label: "Pietanza personalizzata", text: "NUOVA PIETANZA" },
    { value: "separator_long", label: "Separatore linea lunga", text: "SEPARATORE" },
    { value: "separator_short", label: "Separatore linea corta", text: "SEPARATORE BREVE" },
  ].forEach((entry) => {
    const opt = document.createElement("option");
    opt.value = `tool:${entry.value}`;
    opt.textContent = entry.label;
    opt.dataset.defaultText = entry.text;
    toolsGroup.appendChild(opt);
  });
  els.menuBookletDishSelect.appendChild(toolsGroup);
}

function getMenuBookletSelectionDraft() {
  const raw = String(els.menuBookletDishSelect ? els.menuBookletDishSelect.value : "");
  if (raw.startsWith("dish:")) {
    const idx = Number(raw.split(":")[1]);
    const dish = MENU_BOOKLET_DISHES[idx];
    if (dish) return { type: "dish", text: dish.text, title: "Pietanza dal catalogo" };
  }
  if (raw === "tool:separator_long") {
    return { type: "separator_long", text: "", title: "Separatore linea lunga" };
  }
  if (raw === "tool:separator_short") {
    return { type: "separator_short", text: "", title: "Separatore linea corta" };
  }
  return { type: "dish", text: "NUOVA PIETANZA", title: "Pietanza personalizzata" };
}

function openMenuBookletEditor(draftText, mode, type, targetId) {
  if (!els.menuBookletEditorPanel || !els.menuBookletEditorInput || !els.menuBookletEditorTitle || !els.menuBookletEditorConfirmBtn) return;
  menuBookletUiState.mode = mode;
  menuBookletUiState.targetId = targetId || null;
  menuBookletUiState.draftType = type || "dish";
  els.menuBookletEditorPanel.hidden = false;
  els.menuBookletEditorInput.value = String(draftText || "");
  const byMode =
    mode === "edit"
      ? "Modifica elemento selezionato"
      : "Aggiungi nuovo elemento";
  els.menuBookletEditorTitle.textContent = byMode;
  els.menuBookletEditorConfirmBtn.textContent =
    mode === "edit" ? "Accetta modifica" : "Accetta e aggiungi";
  els.menuBookletEditorInput.focus();
  els.menuBookletEditorInput.select();
}

function closeMenuBookletEditor() {
  if (els.menuBookletEditorPanel) els.menuBookletEditorPanel.hidden = true;
  menuBookletUiState.mode = null;
  menuBookletUiState.targetId = null;
  menuBookletUiState.draftType = "dish";
}

function commitMenuBookletEditor() {
  if (!canEditStudioAreas()) return;
  if (!els.menuBookletEditorInput) return;
  const text = String(els.menuBookletEditorInput.value || "").trim();
  const type = menuBookletUiState.draftType || "dish";
  const isSeparator = type === "separator_long" || type === "separator_short";
  if (!text && !isSeparator) {
    alert("Inserisci un testo prima di confermare.");
    return;
  }
  const mode = menuBookletUiState.mode;
  const targetId = menuBookletUiState.targetId;
  const finalText = isSeparator ? "" : text;
  if (mode === "edit" && targetId) {
    const item = menuBookletState.items.find((x) => x.id === targetId);
    if (item) {
      item.text = finalText;
      item.type = type;
    }
  } else {
    menuBookletState.items.push({ id: uid(), type, text: finalText });
  }
  saveMenuBookletState();
  closeMenuBookletEditor();
  renderMenuBookletList();
}

function renderMenuBookletList() {
  if (!els.menuBookletList) return;
  els.menuBookletList.innerHTML = "";
  if (!menuBookletState.items.length) {
    const empty = document.createElement("div");
    empty.className = "menu-booklet-empty";
    empty.textContent = "Nessuna voce inserita. Scegli dalla tendina a sinistra e conferma l'aggiunta.";
    els.menuBookletList.appendChild(empty);
    renderMenuBookletPreview();
    return;
  }

  menuBookletState.items.forEach((item, idx) => {
    const row = document.createElement("div");
    row.className = "menu-booklet-row";

    const head = document.createElement("div");
    head.className = "menu-booklet-row__head";

    const kind = document.createElement("div");
    kind.className = "menu-booklet-row__kind";
    kind.textContent =
      item.type === "separator_short"
        ? "Separatore corto"
        : item.type === "separator_long"
          ? "Separatore lungo"
          : "Pietanza";

    const controlsWrap = document.createElement("div");
    controlsWrap.className = "menu-booklet-row__controls";
    controlsWrap.innerHTML = `
      <button type="button" class="btn btn-ghost btn-sm" data-menu-move-up="${item.id}" ${idx === 0 ? "disabled" : ""} aria-label="Sposta su">↑</button>
      <button type="button" class="btn btn-ghost btn-sm" data-menu-move-down="${item.id}" ${idx === menuBookletState.items.length - 1 ? "disabled" : ""} aria-label="Sposta giù">↓</button>
    `;
    const controlsMenu = document.createElement("details");
    controlsMenu.className = "menu-booklet-row__menu";
    controlsMenu.innerHTML = `
      <summary aria-label="Azioni elemento">...</summary>
      <div class="menu-booklet-row__menu-panel">
        <button type="button" data-menu-action="edit" data-menu-id="${item.id}">Modifica</button>
        <button type="button" data-menu-action="delete" data-menu-id="${item.id}">Elimina</button>
      </div>
    `;
    controlsWrap.appendChild(controlsMenu);

    head.appendChild(kind);
    head.appendChild(controlsWrap);
    const text = document.createElement("div");
    text.className = "menu-booklet-row__text";
    if (item.type === "separator_long" || item.type === "separator_short") {
      text.innerHTML = `<span class="menu-booklet-row__separator-line ${item.type === "separator_short" ? "menu-booklet-row__separator-line--short" : ""}"></span>`;
    } else {
      const firstLine = String(item.text || "").split("\n")[0];
      text.textContent = firstLine.length > 180 ? `${firstLine.slice(0, 180)}...` : item.text;
    }

    row.appendChild(head);
    row.appendChild(text);
    els.menuBookletList.appendChild(row);
  });
  renderMenuBookletPreview();
}

function renderMenuBookletPreview() {
  if (!els.menuBookletPreviewStack) return;
  const theme = getSegnatavoloTheme();
  const palette = getSegnatavoloPaletteResolved();
  els.menuBookletPreviewStack.style.setProperty("--seg-deco", rgbToCss(palette.deco));
  els.menuBookletPreviewStack.querySelectorAll(".menu-booklet-preview-page").forEach((pageEl) => {
    let deco = pageEl.querySelector(".segnatavolo-preview__deco");
    if (!deco) {
      deco = document.createElement("div");
      deco.className = "segnatavolo-preview__deco";
      pageEl.prepend(deco);
    }
    applySegnatavoloPreviewDeco(deco);
  });

  if (els.menuBookletPreviewEvent) {
    els.menuBookletPreviewEvent.textContent = (state.eventName || "Evento").trim() || "Evento";
    els.menuBookletPreviewEvent.style.fontFamily = theme.previewTitle;
    els.menuBookletPreviewEvent.style.color = rgbToCss(palette.title);
  }

  if (els.menuBookletPreviewLogo) {
    if (menuBookletState.logoDataUrl) {
      els.menuBookletPreviewLogo.src = menuBookletState.logoDataUrl;
      els.menuBookletPreviewLogo.hidden = false;
    } else {
      els.menuBookletPreviewLogo.hidden = true;
      els.menuBookletPreviewLogo.removeAttribute("src");
    }
  }

  if (!els.menuBookletPreviewList) return;
  const menuTitleEl = els.menuBookletPreviewStack.querySelector(".menu-booklet-preview-menu-title");
  let menuScale = 1;
  if (window.jspdf && window.jspdf.jsPDF) {
    const scratch = new window.jspdf.jsPDF({ orientation: "landscape", unit: "mm", format: "a4" });
    const Wmm = scratch.internal.pageSize.getWidth();
    const Hmm = scratch.internal.pageSize.getHeight();
    const layoutMb = computeMenuBookletMenuLayout(scratch, theme, Wmm / 2, 0, Wmm / 2, Hmm);
    const refBody = Math.max(8.2, theme.bodySize - 1.2);
    menuScale = layoutMb.bodySize / refBody;
    if (menuTitleEl) {
      menuTitleEl.style.fontSize = `${Math.max(0.62, 0.72 * (layoutMb.titleSz / theme.titleSize))}rem`;
    }
  } else if (menuTitleEl) {
    menuTitleEl.style.fontSize = "";
  }
  if (menuTitleEl) {
    menuTitleEl.style.fontFamily = theme.previewTitle;
    menuTitleEl.style.color = rgbToCss(palette.title);
  }

  els.menuBookletPreviewList.style.fontSize = `${Math.max(0.42, 0.62 * menuScale)}rem`;
  els.menuBookletPreviewList.innerHTML = "";
  if (!menuBookletState.items.length) {
    const row = document.createElement("div");
    row.className = "menu-booklet-preview-menu-item";
    row.style.color = rgbToCss(palette.subtitle);
    row.style.fontStyle = "italic";
    row.textContent = "Nessuna pietanza selezionata.";
    els.menuBookletPreviewList.appendChild(row);
    scheduleFitStudioMenuBookletScale();
    return;
  }

  menuBookletState.items.forEach((item) => {
    const row = document.createElement("div");
    row.className =
      item.type === "separator_short" || item.type === "separator_long"
        ? `menu-booklet-preview-menu-item menu-booklet-preview-menu-separator ${item.type === "separator_short" ? "menu-booklet-preview-menu-separator--short" : ""}`
        : "menu-booklet-preview-menu-item";
    row.textContent = item.type === "separator_short" || item.type === "separator_long" ? "" : item.text;
    row.style.fontFamily = theme.previewBody;
    row.style.color = item.type === "separator_short" || item.type === "separator_long" ? rgbToCss(palette.subtitle) : rgbToCss(palette.body);
    els.menuBookletPreviewList.appendChild(row);
  });
  scheduleFitStudioMenuBookletScale();
}

function getSegnatavoloTheme() {
  return SEGNATAVOLO_THEMES.find((t) => t.id === segnatavoloSettings.themeId) || SEGNATAVOLO_THEMES[0];
}

function getSegnatavoloPaletteResolved() {
  const p = SEGNATAVOLO_PALETTES.find((x) => x.id === segnatavoloSettings.paletteId) || SEGNATAVOLO_PALETTES[0];
  if (p.custom && els.segnatavoloColorTitle && els.segnatavoloColorBody && els.segnatavoloColorDeco) {
    const title = hexToRgb(segnatavoloSettings.customTitleHex || els.segnatavoloColorTitle.value);
    const body = hexToRgb(segnatavoloSettings.customBodyHex || els.segnatavoloColorBody.value);
    const deco = hexToRgb(segnatavoloSettings.customDecoHex || els.segnatavoloColorDeco.value);
    const subtitle = [
      Math.round((title[0] + body[0]) / 2),
      Math.round((title[1] + body[1]) / 2),
      Math.round((title[2] + body[2]) / 2),
    ];
    const foot = [
      Math.round((subtitle[0] + body[0]) / 2),
      Math.round((subtitle[1] + body[1]) / 2),
      Math.round((subtitle[2] + body[2]) / 2),
    ];
    return { title, subtitle, body, foot, deco };
  }
  return {
    title: p.title,
    subtitle: p.subtitle,
    body: p.body,
    foot: p.foot,
    deco: p.deco,
  };
}

/** Decorazioni solo in fascia marginale (non invadono la colonna centrale ~20–128 mm). */
function drawSegnatavoloGraphic(pdf, W, H, decoRgb, idx) {
  const [dr, dg, db] = decoRgb;
  pdf.setDrawColor(dr, dg, db);
  pdf.setLineWidth(0.35);
  const m = 4;
  switch (idx) {
    case 0:
      return;
    case 1:
      pdf.rect(m, m, W - 2 * m, H - 2 * m);
      pdf.rect(m + 3, m + 3, W - 2 * m - 6, H - 2 * m - 6);
      break;
    case 2: {
      const len = 11;
      pdf.line(m, m, m + len, m);
      pdf.line(m, m, m, m + len);
      pdf.line(W - m, m, W - m - len, m);
      pdf.line(W - m, m, W - m, m + len);
      pdf.line(m, H - m, m + len, H - m);
      pdf.line(m, H - m, m, H - m - len);
      pdf.line(W - m, H - m, W - m - len, H - m);
      pdf.line(W - m, H - m, W - m, H - m - len);
      break;
    }
    case 3:
      pdf.setLineWidth(0.45);
      pdf.line(12, 5, W - 12, 5);
      pdf.line(12, H - 5, W - 12, H - 5);
      break;
    case 4: {
      const r = 3.5;
      const inset = m + r;
      pdf.circle(inset, inset, r, "S");
      pdf.circle(W - inset, inset, r, "S");
      pdf.circle(inset, H - inset, r, "S");
      pdf.circle(W - inset, H - inset, r, "S");
      break;
    }
    case 5: {
      const d = 12;
      pdf.line(m, m, m + d, m + d);
      pdf.line(W - m, m, W - m - d, m + d);
      pdf.line(m, H - m, m + d, H - m - d);
      pdf.line(W - m, H - m, W - m - d, H - m - d);
      break;
    }
    case 6: {
      pdf.setFillColor(dr, dg, db);
      const spots = [
        [8, 8],
        [11, 10],
        [8, 12],
        [W - 8, 8],
        [W - 11, 10],
        [W - 8, 12],
        [8, H - 8],
        [11, H - 10],
        [8, H - 12],
        [W - 8, H - 8],
        [W - 11, H - 10],
        [W - 8, H - 12],
      ];
      for (const [x, y] of spots) {
        pdf.circle(x, y, 0.45, "F");
      }
      break;
    }
    case 7:
      pdf.line(8, 20, 8, H - 20);
      pdf.line(W - 8, 20, W - 8, H - 20);
      break;
    case 8: {
      const cx = W / 2;
      const cy = 7;
      const s = 3.5;
      pdf.line(cx, cy - s, cx + s, cy);
      pdf.line(cx + s, cy, cx, cy + s);
      pdf.line(cx, cy + s, cx - s, cy);
      pdf.line(cx - s, cy, cx, cy - s);
      break;
    }
    case 9: {
      pdf.setLineWidth(0.2);
      for (let x = m; x < W - m - 1; x += 2.2) {
        pdf.line(x, m, x + 1.1, m + 0.9);
        pdf.line(x, H - m, x + 1.1, H - m - 0.9);
      }
      break;
    }
    default:
      break;
  }
  pdf.setLineWidth(0.35);
  pdf.setDrawColor(0, 0, 0);
}

function festiveMix(a, b, t) {
  return [
    Math.round(a[0] * (1 - t) + b[0] * t),
    Math.round(a[1] * (1 - t) + b[1] * t),
    Math.round(a[2] * (1 - t) + b[2] * t),
  ];
}

/**
 * Illustrazioni riconoscibili solo nei margini (fuori dalla colonna centrale ~20–128 mm).
 * Colori vivaci mixati alla palette scelta.
 */
function drawSegnatavoloFestiveMarginArt(pdf, W, H, palette, festiveId) {
  const [tr, tg, tb] = palette.title;
  const [dr, dg, db] = palette.deco;
  const tCol = [tr, tg, tb];
  const dCol = [dr, dg, db];
  const gold = festiveMix(dCol, [240, 200, 60], 0.55);
  const red = festiveMix(tCol, [210, 45, 50], 0.35);
  const green = festiveMix(dCol, [30, 130, 70], 0.45);
  const orange = festiveMix(dCol, [235, 120, 35], 0.5);
  const purple = festiveMix(tCol, [120, 60, 160], 0.4);
  const sky = festiveMix(dCol, [70, 150, 220], 0.5);
  const pink = festiveMix(tCol, [240, 150, 185], 0.45);
  const white = [252, 250, 248];
  const brown = festiveMix(dCol, [120, 75, 45], 0.5);

  pdf.setLineWidth(0.25);
  pdf.setDrawColor(dr, dg, db);

  switch (festiveId) {
    case "natale": {
      /* Albero + stella (basso sinistra) */
      pdf.setFillColor(...brown);
      pdf.rect(5.5, H - 7, 2.2, 5, "F");
      pdf.setFillColor(...green);
      pdf.rect(3, H - 14, 7, 4, "F");
      pdf.rect(3.8, H - 19, 5.5, 4, "F");
      pdf.rect(4.8, H - 24, 3.5, 4, "F");
      pdf.setDrawColor(...gold);
      pdf.setLineWidth(0.35);
      pdf.line(6.2, H - 25.5, 6.2, H - 28.5);
      pdf.line(4.8, H - 27, 7.6, H - 27);
      pdf.line(5.4, H - 26.2, 7, H - 28);
      pdf.line(7, H - 26.2, 5.4, H - 28);
      pdf.line(4.8, H - 26.2, 7.6, H - 28.2);
      pdf.line(7.6, H - 26.2, 4.8, H - 28.2);
      pdf.setFillColor(...red);
      pdf.circle(5, H - 16, 0.7, "F");
      pdf.circle(7.5, H - 18, 0.65, "F");
      pdf.setFillColor(...gold);
      pdf.circle(6.5, H - 21, 0.55, "F");
      /* Presepe mini (basso destra) */
      pdf.setFillColor(...brown);
      pdf.lines([[0, 0], [5, 0], [2.5, -4]], W - 16, H - 7, [1, 1], "F", true);
      pdf.setFillColor(...white);
      pdf.circle(W - 13.5, H - 8.2, 0.9, "F");
      pdf.setFillColor(255, 235, 160);
      pdf.rect(W - 15.2, H - 7.1, 3.4, 0.7, "F");
      pdf.setFillColor(180, 140, 90);
      pdf.circle(W - 15.5, H - 6.2, 0.45, "F");
      pdf.setFillColor(90, 90, 120);
      pdf.circle(W - 13.8, H - 6.1, 0.45, "F");
      pdf.setFillColor(200, 60, 60);
      pdf.circle(W - 12.2, H - 6.2, 0.45, "F");
      pdf.setDrawColor(...gold);
      pdf.setLineWidth(0.25);
      pdf.line(W - 15.5, H - 12, W - 11.5, H - 12);
      pdf.line(W - 13.5, H - 13.5, W - 13.5, H - 10.5);
      pdf.line(W - 14.8, H - 11.2, W - 12.2, H - 13.2);
      pdf.line(W - 12.2, H - 11.2, W - 14.8, H - 13.2);
      /* Regali */
      pdf.setFillColor(...red);
      pdf.rect(W - 9, H - 11, 3.5, 3.2, "F");
      pdf.setFillColor(...green);
      pdf.rect(W - 12.5, H - 9.5, 3, 2.8, "F");
      pdf.setDrawColor(...gold);
      pdf.setLineWidth(0.3);
      pdf.line(W - 7.25, H - 11, W - 7.25, H - 7.8);
      pdf.line(W - 9, H - 9.4, W - 5.5, H - 9.4);
      pdf.line(W - 11, H - 9.5, W - 11, H - 6.7);
      pdf.line(W - 12.5, H - 8.4, W - 9.5, H - 8.4);
      /* Babbo Natale — cappello (alto destra) */
      pdf.setFillColor(...red);
      pdf.lines([[0, 0], [6, 2], [0, 5]], W - 9, 6, [1, 1], "F", true);
      pdf.setFillColor(...white);
      pdf.rect(W - 10.5, 9.5, 7, 1.2, "F");
      pdf.circle(W - 9.8, 11.8, 1.1, "F");
      pdf.setFillColor(240, 220, 200);
      pdf.circle(W - 7.5, 12.5, 1.4, "F");
      pdf.setFillColor(...white);
      pdf.circle(W - 6.2, 11.2, 0.45, "F");
      /* Fiocchi di neve (alto sinistra) */
      pdf.setDrawColor(180, 200, 230);
      pdf.setLineWidth(0.2);
      for (const [sx, sy] of [
        [8, 10],
        [12, 7],
      ]) {
        pdf.line(sx - 2, sy, sx + 2, sy);
        pdf.line(sx, sy - 2, sx, sy + 2);
        pdf.line(sx - 1.4, sy - 1.4, sx + 1.4, sy + 1.4);
        pdf.line(sx - 1.4, sy + 1.4, sx + 1.4, sy - 1.4);
      }
      break;
    }
    case "pasqua": {
      pdf.setFillColor(...pink);
      pdf.ellipse(9, 12, 3.2, 4, "F");
      pdf.setDrawColor(...purple);
      pdf.setLineWidth(0.2);
      pdf.ellipse(9, 12, 3.2, 4, "S");
      pdf.setDrawColor(...white);
      for (let i = -2; i <= 2; i++) {
        pdf.line(6.5 + i * 1.2, 10, 7 + i * 1.2, 14);
      }
      pdf.setFillColor(255, 248, 200);
      pdf.ellipse(W - 9, 11, 2.8, 3.6, "F");
      pdf.setFillColor(...orange);
      pdf.circle(W - 9, 11, 0.5, "F");
      pdf.circle(W - 7.5, 12, 0.45, "F");
      pdf.circle(W - 10.2, 12.5, 0.4, "F");
      /* Orecchie coniglio */
      pdf.setFillColor(255, 235, 240);
      pdf.ellipse(6, H - 12, 1.6, 3.8, "F");
      pdf.ellipse(11, H - 12, 1.6, 3.8, "F");
      pdf.setFillColor(...pink);
      pdf.circle(6, H - 14, 0.5, "F");
      pdf.circle(11, H - 14, 0.5, "F");
      /* Uovo e tulipano */
      pdf.setFillColor(...sky);
      pdf.ellipse(W - 10, H - 10, 2.5, 3.2, "F");
      pdf.setFillColor(...green);
      pdf.line(W - 10, H - 13, W - 10, H - 7);
      pdf.lines([[0, 0], [-2, -3], [2, -3]], W - 10, H - 15, [1, 1], "F", true);
      pdf.setFillColor(...red);
      pdf.circle(W - 10, H - 16.5, 1, "F");
      break;
    }
    case "halloween": {
      pdf.setFillColor(...orange);
      pdf.circle(10, 12, 4.5, "F");
      pdf.setFillColor(20, 18, 22);
      pdf.lines([[0, 0], [1.2, -2], [2.4, 0]], 8, 13.5, [1, 1], "F", true);
      pdf.lines([[0, 0], [1.2, -2], [2.4, 0]], 10.8, 13.5, [1, 1], "F", true);
      pdf.lines([[0, 0], [1.5, -1], [3, 0]], 8.8, 16.5, [1, 1], "F", true);
      pdf.setFillColor(...purple);
      pdf.lines([[0, 0], [-3, 2], [-1.5, 0], [0, -1.5], [1.5, 0], [3, 2]], W - 12, 8, [1, 1], "F", true);
      pdf.setFillColor(35, 32, 38);
      pdf.circle(W - 12, 8.5, 0.5, "F");
      /* Ragno */
      pdf.line(W - 8, H - 6, W - 8, H - 12);
      pdf.line(W - 8, H - 9, W - 11, H - 11);
      pdf.line(W - 8, H - 9, W - 5, H - 11);
      pdf.line(W - 8, H - 8, W - 10.5, H - 6.5);
      pdf.line(W - 8, H - 8, W - 5.5, H - 6.5);
      pdf.circle(W - 8, H - 5.8, 0.7, "F");
      /* Zucca piccola BR */
      pdf.setFillColor(200, 95, 20);
      pdf.circle(W - 11, H - 10, 3, "F");
      pdf.setFillColor(15, 12, 10);
      pdf.rect(W - 11.3, H - 13.5, 0.6, 1.2, "F");
      break;
    }
    case "ferragosto": {
      pdf.setFillColor(255, 220, 80);
      pdf.circle(W - 11, 11, 5, "F");
      pdf.setDrawColor(255, 180, 40);
      pdf.setLineWidth(0.35);
      for (let a = 0; a < 360; a += 30) {
        const rad = (a * Math.PI) / 180;
        pdf.line(W - 11, 11, W - 11 + Math.cos(rad) * 7.5, 11 + Math.sin(rad) * 7.5);
      }
      pdf.setFillColor(...sky);
      for (let i = 0; i < 5; i++) {
        pdf.lines(
          [[0, 0], [3, -0.8], [6, 0], [6, 1.5], [3, 0.7], [0, 1.5]],
          4 + i * 0.3,
          H - 6 - (i % 2) * 0.4,
          [1, 1],
          "F",
          true
        );
      }
      pdf.setDrawColor(100, 160, 210);
      pdf.setLineWidth(0.3);
      for (let x = 3; x < 18; x += 2.5) {
        pdf.line(x, H - 8, x + 1.2, H - 7);
      }
      /* Gelato */
      pdf.setFillColor(255, 248, 240);
      pdf.triangle(6, 18, 8, 23, 10, 18, "F");
      pdf.setFillColor(...pink);
      pdf.circle(8, 16.5, 2.2, "F");
      pdf.setFillColor(...orange);
      pdf.circle(8, 14.2, 1.9, "F");
      break;
    }
    case "capodanno": {
      const bursts = [
        [12, 14],
        [W - 12, 12],
        [10, H - 10],
        [W - 10, H - 11],
      ];
      for (const [bx, by] of bursts) {
        pdf.setDrawColor(...gold);
        pdf.setLineWidth(0.25);
        for (let a = 0; a < 360; a += 45) {
          const rad = (a * Math.PI) / 180;
          pdf.line(bx, by, bx + Math.cos(rad) * 4, by + Math.sin(rad) * 4);
        }
        pdf.setFillColor(255, 230, 120);
        pdf.circle(bx, by, 0.5, "F");
      }
      pdf.setDrawColor(200, 200, 210);
      pdf.setLineWidth(0.2);
      for (let i = 0; i < 6; i++) {
        const sx = 6 + i * 2.2;
        const sy = 8 + (i % 2);
        pdf.line(sx, sy, sx + 0.8, sy - 1.2);
        pdf.line(sx + 0.8, sy - 1.2, sx + 1.6, sy);
      }
      /* Bottiglia e calice */
      pdf.setFillColor(40, 110, 70);
      pdf.rect(W - 9, H - 14, 3, 7, "F");
      pdf.setFillColor(...gold);
      pdf.rect(W - 8.7, H - 15.2, 2.4, 1.2, "F");
      pdf.setFillColor(230, 235, 250);
      pdf.lines([[0, 0], [1.8, 0], [2.2, 4], [-2.2, 4], [-1.8, 0]], W - 15, H - 8, [1, 1], "F", true);
      pdf.setDrawColor(180, 190, 210);
      pdf.lines([[0, 0], [1.8, 0], [2.2, 4], [-2.2, 4], [-1.8, 0]], W - 15, H - 8, [1, 1], "S", true);
      pdf.setDrawColor(dr, dg, db);
      break;
    }
    case "comunione": {
      pdf.setDrawColor(...gold);
      pdf.setLineWidth(0.45);
      pdf.line(8, 5, 8, 16);
      pdf.line(4, 9.5, 12, 9.5);
      pdf.setFillColor(...white);
      pdf.circle(8, 20, 3.8, "F");
      pdf.setDrawColor(...gold);
      pdf.circle(8, 20, 3.8, "S");
      pdf.setDrawColor(220, 200, 160);
      pdf.setLineWidth(0.2);
      for (let a = 0; a < 360; a += 40) {
        const rad = (a * Math.PI) / 180;
        pdf.line(8, 20, 8 + Math.cos(rad) * 5.5, 20 + Math.sin(rad) * 5.5);
      }
      // Basso destra: calice con ostia.
      pdf.setFillColor(248, 241, 224);
      pdf.ellipse(W - 9.6, H - 11.2, 2.1, 1.05, "F");
      pdf.setFillColor(236, 224, 195);
      pdf.rect(W - 10.05, H - 10.2, 0.9, 2.2, "F");
      pdf.ellipse(W - 9.6, H - 7.6, 1.45, 0.45, "F");
      pdf.setDrawColor(...gold);
      pdf.setLineWidth(0.22);
      pdf.ellipse(W - 9.6, H - 11.2, 2.1, 1.05, "S");
      pdf.rect(W - 10.05, H - 10.2, 0.9, 2.2, "S");
      pdf.ellipse(W - 9.6, H - 7.6, 1.45, 0.45, "S");
      pdf.setFillColor(255, 252, 244);
      pdf.circle(W - 9.6, H - 13.2, 1.0, "F");
      pdf.setDrawColor(230, 208, 150);
      pdf.circle(W - 9.6, H - 13.2, 1.0, "S");
      pdf.setDrawColor(...gold);
      pdf.setLineWidth(0.18);
      pdf.line(W - 9.6, H - 14, W - 9.6, H - 12.4);
      pdf.line(W - 10.3, H - 13.2, W - 8.9, H - 13.2);
      break;
    }
    case "cresima": {
      pdf.setFillColor(240, 240, 245);
      pdf.ellipse(10, 11, 3.5, 2.2, "F");
      pdf.setDrawColor(160, 165, 175);
      pdf.ellipse(10, 11, 3.5, 2.2, "S");
      pdf.lines([[0, 0], [-2, -1], [-3, -4]], 12, 12, [1, 1], "S", false);
      pdf.setFillColor(255, 140, 60);
      pdf.lines([[0, 0], [1, -4], [2, 0], [1, 1]], W - 10, 16, [1, 1], "F", true);
      pdf.setFillColor(255, 220, 100);
      pdf.lines([[0, 0], [0.8, -3], [1.6, 0]], W - 9.2, 14.5, [1, 1], "F", true);
      pdf.setDrawColor(...green);
      pdf.setLineWidth(0.4);
      pdf.line(W - 8, H - 6, W - 8, H - 14);
      pdf.line(W - 11, H - 10, W - 5, H - 10);
      break;
    }
    case "laurea": {
      pdf.setFillColor(30, 30, 35);
      pdf.rect(5, 6, 8, 1.2, "F");
      pdf.rect(7.5, 7.2, 3, 4, "F");
      pdf.setDrawColor(...gold);
      pdf.setLineWidth(0.25);
      pdf.line(9, 11.2, 9, 14);
      pdf.circle(9, 14.8, 0.7, "F");
      // In basso sinistra: pergamena.
      pdf.setFillColor(255, 251, 242);
      pdf.rect(4.8, H - 11.4, 6.8, 3.4, "F");
      pdf.setDrawColor(200, 182, 130);
      pdf.setLineWidth(0.18);
      pdf.rect(4.8, H - 11.4, 6.8, 3.4, "S");
      pdf.setFillColor(...gold);
      pdf.circle(4.8, H - 9.7, 0.7, "F");
      pdf.circle(11.6, H - 9.7, 0.7, "F");
      pdf.setDrawColor(145, 128, 92);
      pdf.setLineWidth(0.15);
      pdf.line(6, H - 10.2, 10.3, H - 10.2);
      pdf.line(6, H - 9.2, 10.3, H - 9.2);

      // In basso destra: medaglia accademica con nastro.
      pdf.setFillColor(...gold);
      pdf.circle(W - 8.4, H - 9.6, 1.65, "F");
      pdf.setFillColor(255, 242, 195);
      pdf.circle(W - 8.4, H - 9.6, 1.05, "F");
      pdf.setDrawColor(165, 130, 55);
      pdf.setLineWidth(0.2);
      pdf.circle(W - 8.4, H - 9.6, 1.65, "S");
      pdf.setFillColor(45, 95, 165);
      pdf.triangle(W - 9.25, H - 11.1, W - 8.45, H - 13.4, W - 7.9, H - 11.1, "F");
      pdf.setFillColor(62, 120, 188);
      pdf.triangle(W - 7.55, H - 11.1, W - 6.75, H - 13.3, W - 6.2, H - 11.1, "F");
      pdf.setFillColor(120, 88, 35);
      pdf.circle(W - 8.4, H - 9.6, 0.32, "F");
      pdf.setDrawColor(...gold);
      pdf.setLineWidth(0.35);
      pdf.line(W - 14, 8, W - 8, 12);
      pdf.line(W - 8, 12, W - 6, 9);
      break;
    }
    case "promessa":
    case "matrimonio": {
      pdf.setDrawColor(...gold);
      pdf.setLineWidth(0.35);
      pdf.circle(7.5, 11, 2.8, "S");
      pdf.circle(10.5, 11, 2.8, "S");
      if (festiveId === "matrimonio") {
        // Angoli matrimonio: doppia linea elegante + micro-perle.
        const m = 4.8;
        const m2 = 6.6;
        const len = 7.2;
        pdf.setDrawColor(...gold);
        pdf.setLineWidth(0.28);
        // Alto sinistra
        pdf.line(m, m + 1, m + len, m + 1);
        pdf.line(m + 1, m, m + 1, m + len);
        pdf.setLineWidth(0.16);
        pdf.line(m2, m2 + 0.7, m2 + len - 2.2, m2 + 0.7);
        pdf.line(m2 + 0.7, m2, m2 + 0.7, m2 + len - 2.2);
        // Alto destra
        pdf.setLineWidth(0.28);
        pdf.line(W - m, m + 1, W - m - len, m + 1);
        pdf.line(W - m - 1, m, W - m - 1, m + len);
        pdf.setLineWidth(0.16);
        pdf.line(W - m2, m2 + 0.7, W - m2 - len + 2.2, m2 + 0.7);
        pdf.line(W - m2 - 0.7, m2, W - m2 - 0.7, m2 + len - 2.2);
        // Basso sinistra
        pdf.setLineWidth(0.28);
        pdf.line(m, H - m - 1, m + len, H - m - 1);
        pdf.line(m + 1, H - m, m + 1, H - m - len);
        pdf.setLineWidth(0.16);
        pdf.line(m2, H - m2 - 0.7, m2 + len - 2.2, H - m2 - 0.7);
        pdf.line(m2 + 0.7, H - m2, m2 + 0.7, H - m2 - len + 2.2);
        // Basso destra
        pdf.setLineWidth(0.28);
        pdf.line(W - m, H - m - 1, W - m - len, H - m - 1);
        pdf.line(W - m - 1, H - m, W - m - 1, H - m - len);
        pdf.setLineWidth(0.16);
        pdf.line(W - m2, H - m2 - 0.7, W - m2 - len + 2.2, H - m2 - 0.7);
        pdf.line(W - m2 - 0.7, H - m2, W - m2 - 0.7, H - m2 - len + 2.2);
        // Micro-perle
        pdf.setFillColor(...white);
        pdf.circle(m + len - 0.2, m + 1, 0.28, "F");
        pdf.circle(m + 1, m + len - 0.2, 0.28, "F");
        pdf.circle(W - m - len + 0.2, m + 1, 0.28, "F");
        pdf.circle(W - m - 1, m + len - 0.2, 0.28, "F");
        pdf.circle(m + len - 0.2, H - m - 1, 0.28, "F");
        pdf.circle(m + 1, H - m - len + 0.2, 0.28, "F");
        pdf.circle(W - m - len + 0.2, H - m - 1, 0.28, "F");
        pdf.circle(W - m - 1, H - m - len + 0.2, 0.28, "F");
      } else {
        // Promessa: mantiene una variante più semplice.
        pdf.setFillColor(...pink);
        pdf.lines([[0, 0], [1.2, -1], [2.4, 0], [1.2, 1.8]], W - 12, H - 10, [1, 1], "F", true);
        pdf.lines([[0, 0], [1, -0.8], [2, 0], [1, 1.5]], W - 9, H - 12, [1, 1], "F", true);
      }
      break;
    }
    case "battesimo": {
      // SOLO 2 simboli: colomba oro + calice (stile comunione).
      const doveGold = festiveMix(gold, [230, 180, 70], 0.35);
      const doveOutline = festiveMix(doveGold, [140, 105, 45], 0.45);
      pdf.setFillColor(...doveGold);
      pdf.setDrawColor(...doveOutline);
      pdf.setLineWidth(0.22);
      // Colomba oro (alto sinistra) - silhouette più pulita.
      pdf.ellipse(8.9, 9.9, 2.45, 1.05, "FD"); // corpo
      pdf.ellipse(7.05, 8.55, 1.65, 0.68, "FD"); // ala sinistra
      pdf.ellipse(10.45, 8.55, 1.65, 0.68, "FD"); // ala destra
      pdf.ellipse(10.95, 9.55, 0.52, 0.45, "FD"); // testa
      // coda
      pdf.lines([[0, 0], [-1.35, 0.42], [-2.55, -0.12], [-1.7, 0.95], [-0.45, 0.72], [0.35, 0.2]], 6.9, 10.35, [1, 1], "FD", true);
      // dettaglio piuma centrale
      pdf.setLineWidth(0.14);
      pdf.line(8.0, 9.15, 9.95, 9.15);
      pdf.setFillColor(255, 248, 228);
      pdf.circle(11.02, 9.5, 0.07, "F");
      pdf.setFillColor(...doveGold);
      pdf.triangle(11.35, 9.62, 12.0, 9.42, 11.38, 9.9, "F");

      // Calice con ostia (basso destra)
      pdf.setFillColor(248, 241, 224);
      pdf.ellipse(W - 9.6, H - 11.2, 2.1, 1.05, "F");
      pdf.setFillColor(236, 224, 195);
      pdf.rect(W - 10.05, H - 10.2, 0.9, 2.2, "F");
      pdf.ellipse(W - 9.6, H - 7.6, 1.45, 0.45, "F");
      pdf.setDrawColor(...gold);
      pdf.setLineWidth(0.22);
      pdf.ellipse(W - 9.6, H - 11.2, 2.1, 1.05, "S");
      pdf.rect(W - 10.05, H - 10.2, 0.9, 2.2, "S");
      pdf.ellipse(W - 9.6, H - 7.6, 1.45, 0.45, "S");
      pdf.setFillColor(255, 252, 244);
      pdf.circle(W - 9.6, H - 13.2, 1.0, "F");
      pdf.setDrawColor(230, 208, 150);
      pdf.circle(W - 9.6, H - 13.2, 1.0, "S");
      pdf.setDrawColor(...gold);
      pdf.setLineWidth(0.18);
      pdf.line(W - 9.6, H - 14, W - 9.6, H - 12.4);
      pdf.line(W - 10.3, H - 13.2, W - 8.9, H - 13.2);
      break;
    }
    case "compleanno": {
      pdf.setFillColor(...pink);
      pdf.rect(4, H - 8, 8, 2.2, "F");
      pdf.setFillColor(...cream);
      pdf.rect(4.5, H - 10.5, 7, 2.2, "F");
      pdf.setFillColor(...white);
      pdf.rect(5, H - 13, 6, 2.2, "F");
      pdf.setFillColor(...gold);
      for (const cx of [6.5, 8, 9.5]) {
        pdf.rect(cx - 0.15, H - 15, 0.35, 2, "F");
        pdf.circle(cx, H - 15.8, 0.35, "F");
      }
      pdf.setFillColor(...red);
      pdf.circle(6, 8, 2.2, "F");
      pdf.line(6, 10.2, 6, 13);
      pdf.triangle(5.2, 10.2, 6.8, 10.2, 6, 11.2, "F");
      pdf.setFillColor(...sky);
      pdf.circle(W - 7, 10, 1.8, "F");
      pdf.line(W - 7, 11.8, W - 7, 14);
      pdf.triangle(W - 7.8, 11.8, W - 6.2, 11.8, W - 7, 12.8, "F");
      pdf.setFillColor(...orange);
      for (let i = 0; i < 8; i++) {
        pdf.circle(W - 12 + (i % 4) * 1.1, H - 5 - Math.floor(i / 4) * 1.2, 0.25, "F");
      }
      break;
    }
    case "diciottesimo": {
      pdf.setFillColor(...purple);
      pdf.rect(4, H - 7, 8, 2, "F");
      pdf.setFillColor(...pink);
      pdf.rect(4.5, H - 9.2, 7, 2, "F");
      pdf.setFillColor(...gold);
      pdf.rect(5, H - 11.4, 6, 2, "F");
      pdf.setFont("helvetica", "bold");
      pdf.setFontSize(11);
      pdf.setTextColor(...gold);
      pdf.text("18", 6, 8);
      pdf.setTextColor(0, 0, 0);
      for (let a = 0; a < 360; a += 72) {
        const rad = (a * Math.PI) / 180;
        pdf.setDrawColor(...gold);
        pdf.line(W - 11, 12, W - 11 + Math.cos(rad) * 3.5, 12 + Math.sin(rad) * 3.5);
      }
      pdf.setFillColor(255, 240, 200);
      pdf.circle(W - 11, 12, 0.6, "F");
      pdf.setFillColor(230, 230, 250);
      pdf.lines([[0, 0], [1.2, 0], [1.5, 3.5], [-1.5, 3.5], [-1.2, 0]], W - 8, H - 10, [1, 1], "F", true);
      pdf.setDrawColor(...tCol);
      pdf.lines([[0, 0], [1.2, 0], [1.5, 3.5], [-1.5, 3.5], [-1.2, 0]], W - 8, H - 10, [1, 1], "S", true);
      break;
    }
    case "san_valentino": {
      pdf.setFillColor(...red);
      pdf.lines([[0, 0], [1.8, -1.5], [3.6, 0], [1.8, 2.8]], 5, H - 10, [1, 1], "F", true);
      pdf.lines([[0, 0], [1.5, -1.2], [3, 0], [1.5, 2.4]], W - 11, H - 9, [1, 1], "F", true);
      pdf.setDrawColor(...gold);
      pdf.setLineWidth(0.3);
      pdf.line(4, 10, W - 4, H - 12);
      pdf.triangle(3, 9.5, 4.5, 9.5, 4, 11, "F");
      pdf.setFillColor(...red);
      pdf.circle(W - 6, 8, 2, "F");
      pdf.setFillColor(40, 120, 40);
      pdf.line(W - 6, 10, W - 6, 14);
      pdf.setFillColor(...green);
      pdf.ellipse(W - 6, 7.5, 1.2, 0.8, "F");
      break;
    }
    case "carnevale": {
      pdf.setFillColor(...white);
      pdf.ellipse(10, 11, 5, 3.2, "F");
      pdf.setFillColor(25, 22, 28);
      pdf.circle(8, 10.5, 0.65, "F");
      pdf.circle(12, 10.5, 0.65, "F");
      pdf.setDrawColor(...red);
      pdf.setLineWidth(0.35);
      pdf.line(8.5, 12.5, 11.5, 12.5);
      pdf.setFillColor(...purple);
      pdf.lines([[0, 0], [-1, -3], [1, -3]], W - 9, 8, [1, 1], "F", true);
      pdf.setFillColor(...orange);
      pdf.lines([[0, 0], [1, -2.5], [-1, -2.5]], W - 12, 7.5, [1, 1], "F", true);
      pdf.setFillColor(...pink);
      for (let i = 0; i < 12; i++) {
        pdf.circle(4 + (i % 4) * 2.5, H - 5 - Math.floor(i / 4) * 1.5, 0.35, "F");
      }
      break;
    }
    case "anniversario": {
      pdf.setDrawColor(...gold);
      pdf.setLineWidth(0.35);
      pdf.circle(7.5, 11, 2.5, "S");
      pdf.circle(10.2, 11, 2.5, "S");
      pdf.setFillColor(...cream);
      pdf.rect(W - 12, H - 8, 5, 1.8, "F");
      pdf.setFillColor(...pink);
      pdf.rect(W - 11.5, H - 9.8, 4, 1.8, "F");
      pdf.setFillColor(...gold);
      pdf.circle(W - 9.5, H - 11.5, 0.5, "F");
      pdf.setFillColor(240, 248, 255);
      pdf.lines([[0, 0], [1, 0], [1.2, 3], [-1.2, 3], [-1, 0]], W - 7, H - 11, [1, 1], "F", true);
      pdf.lines([[0, 0], [1, 0], [1.2, 3], [-1.2, 3], [-1, 0]], W - 4.5, H - 11, [1, 1], "F", true);
      break;
    }
    case "pensionamento": {
      pdf.setFillColor(255, 220, 100);
      pdf.circle(W - 11, 11, 4.5, "F");
      pdf.setDrawColor(255, 160, 50);
      for (let a = 0; a < 360; a += 45) {
        const rad = (a * Math.PI) / 180;
        pdf.line(W - 11, 11, W - 11 + Math.cos(rad) * 6.5, 11 + Math.sin(rad) * 6.5);
      }
      pdf.setFillColor(...brown);
      pdf.rect(5, H - 9, 6, 0.8, "F");
      pdf.rect(5, H - 9, 0.8, 5, "F");
      pdf.rect(5, H - 5.2, 6, 0.8, "F");
      pdf.setFillColor(180, 200, 220);
      pdf.rect(5.8, H - 8.2, 4.4, 3.2, "F");
      pdf.setDrawColor(120, 130, 150);
      pdf.line(8, H - 12, 6, H - 14);
      pdf.line(8, H - 12, 10, H - 14);
      break;
    }
    case "diploma": {
      pdf.setFillColor(255, 252, 245);
      pdf.rect(5, H - 12, 9, 6, "F");
      pdf.setDrawColor(...gold);
      pdf.setLineWidth(0.4);
      pdf.rect(5, H - 12, 9, 6, "S");
      pdf.circle(6.5, H - 9, 1.1, "S");
      pdf.setFillColor(...gold);
      pdf.circle(6.5, H - 9, 0.85, "F");
      pdf.setFillColor(200, 60, 50);
      pdf.circle(6.5, H - 9, 0.45, "F");
      pdf.setDrawColor(...tCol);
      pdf.setLineWidth(0.25);
      pdf.line(8.5, H - 10, 12.5, H - 10);
      pdf.line(8.5, H - 8.5, 12.5, H - 8.5);
      pdf.setFillColor(30, 30, 35);
      pdf.rect(7.5, 6, 4, 0.9, "F");
      pdf.line(9, 6.9, 9, 9.5);
      pdf.circle(9, 10.2, 0.55, "F");
      break;
    }
    case "festa_donna": {
      const mim = [255, 215, 80];
      pdf.setFillColor(...mim);
      for (const [fx, fy] of [
        [8, 10],
        [9.2, 9.2],
        [7, 9.5],
        [8.5, 8.3],
        [10, 10.5],
      ]) {
        pdf.circle(fx, fy, 0.9, "F");
      }
      pdf.setDrawColor(...green);
      pdf.setLineWidth(0.35);
      pdf.line(8, 12, 8, 16);
      pdf.line(W - 10, 12, W - 9, 15);
      pdf.line(W - 10, 12, W - 11, 15);
      pdf.setFillColor(...green);
      pdf.lines([[0, 0], [-1.2, -2], [1.2, -2]], W - 10, 12, [1, 1], "F", true);
      pdf.setFillColor(255, 100, 140);
      pdf.lines([[0, 0], [1, -0.8], [2, 0], [1, 1.4]], W - 7, H - 10, [1, 1], "F", true);
      break;
    }
    case "nascita": {
      pdf.setFillColor(...sky);
      pdf.circle(9, H - 8, 2.2, "F");
      pdf.setFillColor(...white);
      pdf.rect(10.5, H - 8.5, 2.8, 0.9, "F");
      pdf.circle(13.8, H - 8, 0.55, "F");
      pdf.setDrawColor(...pink);
      pdf.setLineWidth(0.25);
      pdf.circle(9, H - 8, 2.2, "S");
      pdf.setFillColor(...gold);
      for (let i = 0; i < 5; i++) {
        const ang = (i / 5) * Math.PI * 2;
        pdf.circle(6 + Math.cos(ang) * 2.5, 9 + Math.sin(ang) * 2.5, 0.25, "F");
      }
      pdf.setFillColor(...pink);
      pdf.lines([[0, 0], [0.8, -0.6], [1.6, 0], [0.8, 1]], W - 9, H - 9, [1, 1], "F", true);
      break;
    }
    case "addio_festa": {
      pdf.setFillColor(255, 240, 250);
      pdf.lines([[0, 0], [1, 0], [1.2, 3.2], [-1.2, 3.2], [-1, 0]], W - 12, H - 11, [1, 1], "F", true);
      pdf.lines([[0, 0], [1, 0], [1.2, 3.2], [-1.2, 3.2], [-1, 0]], W - 8.5, H - 11, [1, 1], "F", true);
      pdf.setFillColor(...pink);
      for (let i = 0; i < 10; i++) {
        pdf.circle(4 + (i % 5) * 2, 7 + Math.floor(i / 5) * 2, 0.3, "F");
      }
      pdf.setFillColor(...gold);
      for (let a = 0; a < 360; a += 60) {
        const rad = (a * Math.PI) / 180;
        pdf.line(10, 12, 10 + Math.cos(rad) * 2.8, 12 + Math.sin(rad) * 2.8);
      }
      break;
    }
    case "inaugurazione": {
      pdf.setDrawColor(...gold);
      pdf.setLineWidth(0.35);
      pdf.line(W - 12, 8, W - 8, 12);
      pdf.line(W - 8, 12, W - 6, 9);
      pdf.setFillColor(...red);
      pdf.lines([[0, 0], [1.5, 0.5], [1.2, 2], [-1.2, 2], [-1.5, 0.5]], 7, 8, [1, 1], "F", true);
      pdf.setFillColor(...gold);
      pdf.circle(7, 9.5, 0.45, "F");
      pdf.setDrawColor(...tCol);
      pdf.setLineWidth(0.25);
      for (let a = 0; a < 360; a += 90) {
        const rad = (a * Math.PI) / 180;
        pdf.line(W - 10, H - 9, W - 10 + Math.cos(rad) * 2.5, H - 9 + Math.sin(rad) * 2.5);
      }
      break;
    }
    default:
      break;
  }
  pdf.setDrawColor(0, 0, 0);
  pdf.setLineWidth(0.35);
  pdf.setTextColor(0, 0, 0);
}

/** Anteprima HTML/SVG allineata ai disegni PDF (margini soltanto). */
function buildSegnatavoloFestivePreviewInnerHTML(palette, festiveId) {
  if (!festiveId) return "";
  const tCol = palette.title;
  const dCol = palette.deco;
  const gold = festiveMix(dCol, [240, 200, 60], 0.55);
  const red = festiveMix(tCol, [210, 45, 50], 0.35);
  const green = festiveMix(dCol, [30, 130, 70], 0.45);
  const orange = festiveMix(dCol, [235, 120, 35], 0.5);
  const purple = festiveMix(tCol, [120, 60, 160], 0.4);
  const sky = festiveMix(dCol, [70, 150, 220], 0.5);
  const pink = festiveMix(tCol, [240, 150, 185], 0.45);
  const white = [252, 250, 248];
  const brown = festiveMix(dCol, [120, 75, 45], 0.5);
  const cream = [255, 248, 200];
  const c = (rgb) => rgbToCss(rgb);
  const W = 148;
  const H = 210;
  const svgStart = `<svg class="segnatavolo-festive-layer" viewBox="0 0 ${W} ${H}" preserveAspectRatio="none" aria-hidden="true">`;

  switch (festiveId) {
    case "natale":
      return `${svgStart}
        <rect x="5.5" y="${H - 12}" width="2.2" height="5" fill="${c(brown)}"/>
        <rect x="3" y="${H - 19}" width="7" height="4" fill="${c(green)}"/>
        <rect x="3.8" y="${H - 24}" width="5.5" height="4" fill="${c(green)}"/>
        <rect x="4.8" y="${H - 29}" width="3.5" height="4" fill="${c(green)}"/>
        <g stroke="${c(gold)}" fill="none" stroke-width="0.45">
          <path d="M6.2 ${H - 25.5} L6.2 ${H - 28.5} M4.8 ${H - 27} L7.6 ${H - 27} M4.8 ${H - 26.2} L7.6 ${H - 28.2} M7.6 ${H - 26.2} L4.8 ${H - 28.2}"/>
        </g>
        <circle cx="5" cy="${H - 16}" r="0.75" fill="${c(red)}"/>
        <circle cx="7.5" cy="${H - 18}" r="0.65" fill="${c(red)}"/>
        <circle cx="6.5" cy="${H - 21}" r="0.55" fill="${c(gold)}"/>
        <polygon points="${W - 16},${H - 7} ${W - 11},${H - 7} ${W - 13.5},${H - 11}" fill="${c(brown)}"/>
        <circle cx="${W - 13.5}" cy="${H - 8.2}" r="0.95" fill="${c(white)}"/>
        <rect x="${W - 15.2}" y="${H - 7.1}" width="3.4" height="0.7" fill="rgb(255,235,160)"/>
        <circle cx="${W - 15.5}" cy="${H - 6.2}" r="0.45" fill="rgb(180,140,90)"/>
        <circle cx="${W - 13.8}" cy="${H - 6.1}" r="0.45" fill="rgb(90,90,120)"/>
        <circle cx="${W - 12.2}" cy="${H - 6.2}" r="0.45" fill="rgb(200,60,60)"/>
        <g stroke="${c(gold)}" fill="none" stroke-width="0.3">
          <path d="M${W - 15.5} ${H - 12} L${W - 11.5} ${H - 12} M${W - 13.5} ${H - 13.5} L${W - 13.5} ${H - 10.5}"/>
        </g>
        <rect x="${W - 9}" y="${H - 11}" width="3.5" height="3.2" fill="${c(red)}"/>
        <rect x="${W - 12.5}" y="${H - 9.5}" width="3" height="2.8" fill="${c(green)}"/>
        <g stroke="${c(gold)}" stroke-width="0.25">
          <path d="M${W - 7.25} ${H - 11} L${W - 7.25} ${H - 7.8} M${W - 9} ${H - 9.4} L${W - 5.5} ${H - 9.4}"/>
        </g>
        <polygon points="${W - 9},6 ${W - 3},8 ${W - 9},11" fill="${c(red)}"/>
        <rect x="${W - 10.5}" y="9.5" width="7" height="1.2" fill="${c(white)}"/>
        <circle cx="${W - 9.8}" cy="11.8" r="1" fill="${c(white)}"/>
        <circle cx="${W - 7.5}" cy="12.5" r="1.35" fill="rgb(240,220,200)"/>
        <circle cx="${W - 6.2}" cy="11.2" r="0.45" fill="${c(white)}"/>
        <g stroke="rgb(180,200,230)" fill="none" stroke-width="0.2">
          <path d="M6 10 L10 10 M8 8 L8 12 M6.6 8.6 L9.4 11.4 M9.4 8.6 L6.6 11.4"/>
          <path d="M10 7 L14 7 M12 5 L12 9 M10.6 5.6 L13.4 8.4 M13.4 5.6 L10.6 8.4"/>
        </g>
      </svg>`;
    case "pasqua":
      return `${svgStart}
        <ellipse cx="9" cy="12" rx="3.2" ry="4" fill="${c(pink)}" stroke="${c(purple)}" stroke-width="0.2"/>
        <ellipse cx="${W - 9}" cy="11" rx="2.8" ry="3.6" fill="${c(cream)}"/>
        <circle cx="${W - 9}" cy="11" r="0.5" fill="${c(orange)}"/>
        <circle cx="${W - 7.5}" cy="12" r="0.45" fill="${c(orange)}"/>
        <ellipse cx="6" cy="${H - 12}" rx="1.6" ry="3.8" fill="rgb(255,235,240)"/>
        <ellipse cx="11" cy="${H - 12}" rx="1.6" ry="3.8" fill="rgb(255,235,240)"/>
        <circle cx="6" cy="${H - 14}" r="0.5" fill="${c(pink)}"/>
        <circle cx="11" cy="${H - 14}" r="0.5" fill="${c(pink)}"/>
        <ellipse cx="${W - 10}" cy="${H - 10}" rx="2.5" ry="3.2" fill="${c(sky)}"/>
        <line x1="${W - 10}" y1="${H - 13}" x2="${W - 10}" y2="${H - 7}" stroke="${c(green)}" stroke-width="0.35"/>
        <polygon points="${W - 10},${H - 15} ${W - 12},${H - 18} ${W - 8},${H - 18}" fill="${c(green)}"/>
        <circle cx="${W - 10}" cy="${H - 16.5}" r="1" fill="${c(red)}"/>
      </svg>`;
    case "halloween":
      return `${svgStart}
        <circle cx="10" cy="12" r="4.5" fill="${c(orange)}"/>
        <polygon points="8,13.5 9.2,11.5 10.4,13.5" fill="rgb(20,18,22)"/>
        <polygon points="10.8,13.5 12,11.5 13.2,13.5" fill="rgb(20,18,22)"/>
        <polygon points="8.8,16.5 10.3,15.5 11.8,16.5" fill="rgb(20,18,22)"/>
        <polygon points="${W - 12},8 ${W - 15},10 ${W - 13.5},10 ${W - 12},8.5 ${W - 10.5},10 ${W - 9},10" fill="${c(purple)}"/>
        <circle cx="${W - 8}" cy="${H - 5.8}" r="0.7" fill="rgb(35,32,38)"/>
        <path d="M${W - 8} ${H - 6} L${W - 8} ${H - 12} M${W - 8} ${H - 9} L${W - 11} ${H - 11} M${W - 8} ${H - 9} L${W - 5} ${H - 11}" stroke="rgb(35,32,38)" fill="none" stroke-width="0.25"/>
        <circle cx="${W - 11}" cy="${H - 10}" r="3" fill="rgb(200,95,20)"/>
        <rect x="${W - 11.3}" y="${H - 13.5}" width="0.6" height="1.2" fill="rgb(15,12,10)"/>
      </svg>`;
    case "ferragosto":
      return `${svgStart}
        <circle cx="${W - 11}" cy="11" r="5" fill="rgb(255,220,80)"/>
        <g stroke="rgb(255,180,40)" fill="none" stroke-width="0.35">
          ${[0, 30, 60, 90, 120, 150, 180, 210, 240, 270, 300, 330]
            .map((deg) => {
              const r = (deg * Math.PI) / 180;
              const x2 = W - 11 + Math.cos(r) * 7.5;
              const y2 = 11 + Math.sin(r) * 7.5;
              return `<line x1="${W - 11}" y1="11" x2="${x2.toFixed(2)}" y2="${y2.toFixed(2)}"/>`;
            })
            .join("")}
        </g>
        <path d="M3 ${H - 7} Q8 ${H - 9} 14 ${H - 7}" fill="none" stroke="${c(sky)}" stroke-width="0.8"/>
        <polygon points="6,18 8,23 10,18" fill="rgb(255,248,240)"/>
        <circle cx="8" cy="16.5" r="2.2" fill="${c(pink)}"/>
        <circle cx="8" cy="14.2" r="1.9" fill="${c(orange)}"/>
      </svg>`;
    case "capodanno":
      return `${svgStart}
        ${[
          [12, 14],
          [W - 12, 12],
          [10, H - 10],
          [W - 10, H - 11],
        ]
          .map(([bx, by]) => {
            let g = `<g>`;
            for (let a = 0; a < 360; a += 45) {
              const rad = (a * Math.PI) / 180;
              g += `<line x1="${bx}" y1="${by}" x2="${bx + Math.cos(rad) * 4}" y2="${by + Math.sin(rad) * 4}" stroke="${c(gold)}" stroke-width="0.25"/>`;
            }
            g += `<circle cx="${bx}" cy="${by}" r="0.5" fill="rgb(255,230,120)"/></g>`;
            return g;
          })
          .join("")}
        <rect x="${W - 9}" y="${H - 14}" width="3" height="7" fill="rgb(40,110,70)"/>
        <rect x="${W - 8.7}" y="${H - 15.2}" width="2.4" height="1.2" fill="${c(gold)}"/>
        <polygon points="${W - 15},${H - 8} ${W - 13.2},${H - 8} ${W - 12.8},${H - 4} ${W - 17.2},${H - 4}" fill="rgb(230,235,250)" stroke="rgb(180,190,210)" stroke-width="0.15"/>
      </svg>`;
    case "comunione":
      return `${svgStart}
        <path d="M8 5 L8 16 M4 9.5 L12 9.5" stroke="${c(gold)}" fill="none" stroke-width="0.5"/>
        <circle cx="8" cy="20" r="3.8" fill="${c(white)}" stroke="${c(gold)}" stroke-width="0.35"/>
        <g stroke="rgb(220,200,160)" fill="none" stroke-width="0.2">
          ${[0, 40, 80, 120, 160, 200, 240, 280, 320]
            .map((deg) => {
              const r = (deg * Math.PI) / 180;
              return `<line x1="8" y1="20" x2="${8 + Math.cos(r) * 5.5}" y2="${20 + Math.sin(r) * 5.5}"/>`;
            })
            .join("")}
        </g>
        <ellipse cx="${W - 9.6}" cy="${H - 11.2}" rx="2.1" ry="1.05" fill="rgb(248,241,224)" stroke="${c(gold)}" stroke-width="0.22"/>
        <rect x="${W - 10.05}" y="${H - 10.2}" width="0.9" height="2.2" fill="rgb(236,224,195)" stroke="${c(gold)}" stroke-width="0.22"/>
        <ellipse cx="${W - 9.6}" cy="${H - 7.6}" rx="1.45" ry="0.45" fill="rgb(236,224,195)" stroke="${c(gold)}" stroke-width="0.22"/>
        <circle cx="${W - 9.6}" cy="${H - 13.2}" r="1.0" fill="rgb(255,252,244)" stroke="rgb(230,208,150)" stroke-width="0.2"/>
        <line x1="${W - 9.6}" y1="${H - 14}" x2="${W - 9.6}" y2="${H - 12.4}" stroke="${c(gold)}" stroke-width="0.18"/>
        <line x1="${W - 10.3}" y1="${H - 13.2}" x2="${W - 8.9}" y2="${H - 13.2}" stroke="${c(gold)}" stroke-width="0.18"/>
      </svg>`;
    case "cresima":
      return `${svgStart}
        <ellipse cx="10" cy="11" rx="3.5" ry="2.2" fill="rgb(240,240,245)" stroke="rgb(160,165,175)" stroke-width="0.2"/>
        <path d="M12 12 L10 11 L7 8" fill="none" stroke="rgb(160,165,175)" stroke-width="0.25"/>
        <polygon points="${W - 10},16 ${W - 8},16 ${W - 9},12" fill="rgb(255,140,60)"/>
        <polygon points="${W - 9.2},14.5 ${W - 8.4},14.5 ${W - 9},11.5" fill="rgb(255,220,100)"/>
        <path d="M${W - 8} ${H - 6} L${W - 8} ${H - 14} M${W - 11} ${H - 10} L${W - 5} ${H - 10}" stroke="${c(green)}" fill="none" stroke-width="0.4"/>
      </svg>`;
    case "laurea":
      return `${svgStart}
        <rect x="5" y="6" width="8" height="1.2" fill="rgb(30,30,35)"/>
        <rect x="7.5" y="7.2" width="3" height="4" fill="rgb(30,30,35)"/>
        <line x1="9" y1="11.2" x2="9" y2="14" stroke="${c(gold)}" stroke-width="0.25"/>
        <circle cx="9" cy="14.8" r="0.7" fill="${c(gold)}"/>
        <rect x="4.8" y="${H - 11.4}" width="6.8" height="3.4" fill="rgb(255,251,242)" stroke="rgb(200,182,130)" stroke-width="0.18"/>
        <circle cx="4.8" cy="${H - 9.7}" r="0.7" fill="${c(gold)}"/>
        <circle cx="11.6" cy="${H - 9.7}" r="0.7" fill="${c(gold)}"/>
        <line x1="6" y1="${H - 10.2}" x2="10.3" y2="${H - 10.2}" stroke="rgb(145,128,92)" stroke-width="0.15"/>
        <line x1="6" y1="${H - 9.2}" x2="10.3" y2="${H - 9.2}" stroke="rgb(145,128,92)" stroke-width="0.15"/>
        <circle cx="${W - 8.4}" cy="${H - 9.6}" r="1.65" fill="${c(gold)}" stroke="rgb(165,130,55)" stroke-width="0.2"/>
        <circle cx="${W - 8.4}" cy="${H - 9.6}" r="1.05" fill="rgb(255,242,195)"/>
        <polygon points="${W - 9.25},${H - 11.1} ${W - 8.45},${H - 13.4} ${W - 7.9},${H - 11.1}" fill="rgb(45,95,165)"/>
        <polygon points="${W - 7.55},${H - 11.1} ${W - 6.75},${H - 13.3} ${W - 6.2},${H - 11.1}" fill="rgb(62,120,188)"/>
        <circle cx="${W - 8.4}" cy="${H - 9.6}" r="0.32" fill="rgb(120,88,35)"/>
        <path d="M${W - 14} 8 L${W - 8} 12 L${W - 6} 9" fill="none" stroke="${c(gold)}" stroke-width="0.35"/>
      </svg>`;
    case "promessa":
    case "matrimonio":
      return `${svgStart}
        <circle cx="7.5" cy="11" r="2.8" fill="none" stroke="${c(gold)}" stroke-width="0.35"/>
        <circle cx="10.5" cy="11" r="2.8" fill="none" stroke="${c(gold)}" stroke-width="0.35"/>
        ${
          festiveId === "matrimonio"
            ? `
              <g stroke="${c(gold)}" fill="none">
                <g stroke-width="0.28">
                  <line x1="4.8" y1="5.8" x2="12" y2="5.8"/><line x1="5.8" y1="4.8" x2="5.8" y2="12"/>
                  <line x1="${W - 4.8}" y1="5.8" x2="${W - 12}" y2="5.8"/><line x1="${W - 5.8}" y1="4.8" x2="${W - 5.8}" y2="12"/>
                  <line x1="4.8" y1="${H - 5.8}" x2="12" y2="${H - 5.8}"/><line x1="5.8" y1="${H - 4.8}" x2="5.8" y2="${H - 12}"/>
                  <line x1="${W - 4.8}" y1="${H - 5.8}" x2="${W - 12}" y2="${H - 5.8}"/><line x1="${W - 5.8}" y1="${H - 4.8}" x2="${W - 5.8}" y2="${H - 12}"/>
                </g>
                <g stroke-width="0.16">
                  <line x1="6.6" y1="7.3" x2="11.6" y2="7.3"/><line x1="7.3" y1="6.6" x2="7.3" y2="11.6"/>
                  <line x1="${W - 6.6}" y1="7.3" x2="${W - 11.6}" y2="7.3"/><line x1="${W - 7.3}" y1="6.6" x2="${W - 7.3}" y2="11.6"/>
                  <line x1="6.6" y1="${H - 7.3}" x2="11.6" y2="${H - 7.3}"/><line x1="7.3" y1="${H - 6.6}" x2="7.3" y2="${H - 11.6}"/>
                  <line x1="${W - 6.6}" y1="${H - 7.3}" x2="${W - 11.6}" y2="${H - 7.3}"/><line x1="${W - 7.3}" y1="${H - 6.6}" x2="${W - 7.3}" y2="${H - 11.6}"/>
                </g>
              </g>
              <circle cx="11.8" cy="5.8" r="0.28" fill="${c(white)}"/>
              <circle cx="5.8" cy="11.8" r="0.28" fill="${c(white)}"/>
              <circle cx="${W - 11.8}" cy="5.8" r="0.28" fill="${c(white)}"/>
              <circle cx="${W - 5.8}" cy="11.8" r="0.28" fill="${c(white)}"/>
              <circle cx="11.8" cy="${H - 5.8}" r="0.28" fill="${c(white)}"/>
              <circle cx="5.8" cy="${H - 11.8}" r="0.28" fill="${c(white)}"/>
              <circle cx="${W - 11.8}" cy="${H - 5.8}" r="0.28" fill="${c(white)}"/>
              <circle cx="${W - 5.8}" cy="${H - 11.8}" r="0.28" fill="${c(white)}"/>
            `
            : `
              <polygon points="${W - 12},${H - 10} ${W - 10.8},${H - 11} ${W - 9.6},${H - 10} ${W - 10.8},${H - 8.2}" fill="${c(pink)}"/>
              <polygon points="${W - 9},${H - 12} ${W - 8},${H - 12.8} ${W - 7},${H - 12} ${W - 8},${H - 10.5}" fill="${c(pink)}"/>
            `
        }
      </svg>`;
    case "battesimo":
      return `${svgStart}
        <ellipse cx="8.9" cy="9.9" rx="2.45" ry="1.05" fill="${c(festiveMix(gold, [230,180,70], 0.35))}" stroke="${c(festiveMix(festiveMix(gold, [230,180,70], 0.35), [140,105,45], 0.45))}" stroke-width="0.22"/>
        <ellipse cx="7.05" cy="8.55" rx="1.65" ry="0.68" fill="${c(festiveMix(gold, [230,180,70], 0.35))}" stroke="${c(festiveMix(festiveMix(gold, [230,180,70], 0.35), [140,105,45], 0.45))}" stroke-width="0.22"/>
        <ellipse cx="10.45" cy="8.55" rx="1.65" ry="0.68" fill="${c(festiveMix(gold, [230,180,70], 0.35))}" stroke="${c(festiveMix(festiveMix(gold, [230,180,70], 0.35), [140,105,45], 0.45))}" stroke-width="0.22"/>
        <ellipse cx="10.95" cy="9.55" rx="0.52" ry="0.45" fill="${c(festiveMix(gold, [230,180,70], 0.35))}" stroke="${c(festiveMix(festiveMix(gold, [230,180,70], 0.35), [140,105,45], 0.45))}" stroke-width="0.22"/>
        <path d="M6.9 10.35 C6.0 10.7,4.9 10.45,4.35 10.9 C5.25 11.4,6.35 11.3,7.25 10.9 Z" fill="${c(festiveMix(gold, [230,180,70], 0.35))}" stroke="${c(festiveMix(festiveMix(gold, [230,180,70], 0.35), [140,105,45], 0.45))}" stroke-width="0.22"/>
        <line x1="8.0" y1="9.15" x2="9.95" y2="9.15" stroke="${c(festiveMix(festiveMix(gold, [230,180,70], 0.35), [140,105,45], 0.45))}" stroke-width="0.14"/>
        <circle cx="11.02" cy="9.5" r="0.07" fill="rgb(255,248,228)"/>
        <polygon points="11.35,9.62 12.0,9.42 11.38,9.9" fill="${c(festiveMix(gold, [230,180,70], 0.35))}"/>

        <ellipse cx="${W - 9.6}" cy="${H - 11.2}" rx="2.1" ry="1.05" fill="rgb(248,241,224)" stroke="${c(gold)}" stroke-width="0.22"/>
        <rect x="${W - 10.05}" y="${H - 10.2}" width="0.9" height="2.2" fill="rgb(236,224,195)" stroke="${c(gold)}" stroke-width="0.22"/>
        <ellipse cx="${W - 9.6}" cy="${H - 7.6}" rx="1.45" ry="0.45" fill="rgb(236,224,195)" stroke="${c(gold)}" stroke-width="0.22"/>
        <circle cx="${W - 9.6}" cy="${H - 13.2}" r="1.0" fill="rgb(255,252,244)" stroke="rgb(230,208,150)" stroke-width="0.2"/>
        <line x1="${W - 9.6}" y1="${H - 14}" x2="${W - 9.6}" y2="${H - 12.4}" stroke="${c(gold)}" stroke-width="0.18"/>
        <line x1="${W - 10.3}" y1="${H - 13.2}" x2="${W - 8.9}" y2="${H - 13.2}" stroke="${c(gold)}" stroke-width="0.18"/>
      </svg>`;
    case "compleanno":
      return `${svgStart}
        <rect x="4" y="${H - 8}" width="8" height="2.2" fill="${c(pink)}"/>
        <rect x="4.5" y="${H - 10.5}" width="7" height="2.2" fill="${c(cream)}"/>
        <rect x="5" y="${H - 13}" width="6" height="2.2" fill="${c(white)}"/>
        ${[6.5, 8, 9.5].map((cx) => `<rect x="${cx - 0.15}" y="${H - 15}" width="0.35" height="2" fill="${c(gold)}"/><circle cx="${cx}" cy="${H - 15.8}" r="0.35" fill="${c(gold)}"/>`).join("")}
        <circle cx="6" cy="8" r="2.2" fill="${c(red)}"/>
        <line x1="6" y1="10.2" x2="6" y2="13" stroke="${c(red)}" stroke-width="0.25"/>
        <polygon points="5.2,10.2 6.8,10.2 6,11.2" fill="${c(red)}"/>
        <circle cx="${W - 7}" cy="10" r="1.8" fill="${c(sky)}"/>
        <line x1="${W - 7}" y1="11.8" x2="${W - 7}" y2="14" stroke="${c(sky)}" stroke-width="0.25"/>
        <polygon points="${W - 7.8},11.8 ${W - 6.2},11.8 ${W - 7},12.8" fill="${c(sky)}"/>
        ${[0, 1, 2, 3, 4, 5, 6, 7]
          .map((i) => `<circle cx="${W - 12 + (i % 4) * 1.1}" cy="${H - 5 - Math.floor(i / 4) * 1.2}" r="0.25" fill="${c(orange)}"/>`)
          .join("")}
      </svg>`;
    case "diciottesimo":
      return `${svgStart}
        <rect x="4" y="${H - 7}" width="8" height="2" fill="${c(purple)}"/>
        <rect x="4.5" y="${H - 9.2}" width="7" height="2" fill="${c(pink)}"/>
        <rect x="5" y="${H - 11.4}" width="6" height="2" fill="${c(gold)}"/>
        <text x="6" y="9" font-size="8" font-weight="700" fill="${c(gold)}" font-family="Arial,sans-serif">18</text>
        <g stroke="${c(gold)}" fill="none" stroke-width="0.25">${[0, 72, 144, 216, 288]
          .map((deg) => {
            const r = (deg * Math.PI) / 180;
            return `<line x1="${W - 11}" y1="12" x2="${W - 11 + Math.cos(r) * 3.5}" y2="${12 + Math.sin(r) * 3.5}"/>`;
          })
          .join("")}</g>
        <circle cx="${W - 11}" cy="12" r="0.6" fill="rgb(255,240,200)"/>
        <polygon points="${W - 8},${H - 10} ${W - 6.8},${H - 10} ${W - 6.5},${H - 6.5} ${W - 9.5},${H - 6.5}" fill="rgb(230,230,250)" stroke="${c(tCol)}" stroke-width="0.15"/>
      </svg>`;
    case "san_valentino":
      return `${svgStart}
        <polygon points="5,${H - 10} 6.8,${H - 11.5} 8.6,${H - 10} 6.8,${H - 7.2}" fill="${c(red)}"/>
        <polygon points="${W - 11},${H - 9} ${W - 9.2},${H - 10.2} ${W - 7.4},${H - 9} ${W - 9.2},${H - 6.6}" fill="${c(red)}"/>
        <line x1="4" y1="10" x2="${W - 4}" y2="${H - 12}" stroke="${c(gold)}" stroke-width="0.3"/>
        <polygon points="3,9.5 4.5,9.5 4,11" fill="${c(gold)}"/>
        <circle cx="${W - 6}" cy="8" r="2" fill="${c(red)}"/>
        <line x1="${W - 6}" y1="10" x2="${W - 6}" y2="14" stroke="rgb(40,120,40)" stroke-width="0.35"/>
        <ellipse cx="${W - 6}" cy="7.5" rx="1.2" ry="0.8" fill="${c(green)}"/>
      </svg>`;
    case "carnevale":
      return `${svgStart}
        <ellipse cx="10" cy="11" rx="5" ry="3.2" fill="${c(white)}"/>
        <circle cx="8" cy="10.5" r="0.65" fill="rgb(25,22,28)"/>
        <circle cx="12" cy="10.5" r="0.65" fill="rgb(25,22,28)"/>
        <line x1="8.5" y1="12.5" x2="11.5" y2="12.5" stroke="${c(red)}" stroke-width="0.35"/>
        <polygon points="${W - 9},8 ${W - 10},5 ${W - 8},5" fill="${c(purple)}"/>
        <polygon points="${W - 12},7.5 ${W - 11},5 ${W - 13},5" fill="${c(orange)}"/>
        ${[0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11]
          .map((i) => `<circle cx="${4 + (i % 4) * 2.5}" cy="${H - 5 - Math.floor(i / 4) * 1.5}" r="0.35" fill="${c(pink)}"/>`)
          .join("")}
      </svg>`;
    case "anniversario":
      return `${svgStart}
        <circle cx="7.5" cy="11" r="2.5" fill="none" stroke="${c(gold)}" stroke-width="0.35"/>
        <circle cx="10.2" cy="11" r="2.5" fill="none" stroke="${c(gold)}" stroke-width="0.35"/>
        <rect x="${W - 12}" y="${H - 8}" width="5" height="1.8" fill="${c(cream)}"/>
        <rect x="${W - 11.5}" y="${H - 9.8}" width="4" height="1.8" fill="${c(pink)}"/>
        <circle cx="${W - 9.5}" cy="${H - 11.5}" r="0.5" fill="${c(gold)}"/>
        <polygon points="${W - 7},${H - 11} ${W - 6},${H - 11} ${W - 5.8},${H - 8} ${W - 8.2},${H - 8}" fill="rgb(240,248,255)"/>
        <polygon points="${W - 4.5},${H - 11} ${W - 3.5},${H - 11} ${W - 3.3},${H - 8} ${W - 5.7},${H - 8}" fill="rgb(240,248,255)"/>
      </svg>`;
    case "pensionamento":
      return `${svgStart}
        <circle cx="${W - 11}" cy="11" r="4.5" fill="rgb(255,220,100)"/>
        <g stroke="rgb(255,160,50)" fill="none" stroke-width="0.35">${[0, 45, 90, 135, 180, 225, 270, 315]
          .map((deg) => {
            const r = (deg * Math.PI) / 180;
            return `<line x1="${W - 11}" y1="11" x2="${W - 11 + Math.cos(r) * 6.5}" y2="${11 + Math.sin(r) * 6.5}"/>`;
          })
          .join("")}</g>
        <rect x="5" y="${H - 9}" width="6" height="0.8" fill="${c(brown)}"/>
        <rect x="5" y="${H - 9}" width="0.8" height="5" fill="${c(brown)}"/>
        <rect x="5" y="${H - 5.2}" width="6" height="0.8" fill="${c(brown)}"/>
        <rect x="5.8" y="${H - 8.2}" width="4.4" height="3.2" fill="rgb(180,200,220)"/>
        <path d="M8 ${H - 12} L6 ${H - 14} M8 ${H - 12} L10 ${H - 14}" stroke="rgb(120,130,150)" fill="none" stroke-width="0.25"/>
      </svg>`;
    case "diploma":
      return `${svgStart}
        <rect x="5" y="${H - 12}" width="9" height="6" fill="rgb(255,252,245)" stroke="${c(gold)}" stroke-width="0.4"/>
        <circle cx="6.5" cy="${H - 9}" r="1.1" fill="none" stroke="${c(gold)}" stroke-width="0.35"/>
        <circle cx="6.5" cy="${H - 9}" r="0.85" fill="${c(gold)}"/>
        <circle cx="6.5" cy="${H - 9}" r="0.45" fill="rgb(200,60,50)"/>
        <line x1="8.5" y1="${H - 10}" x2="12.5" y2="${H - 10}" stroke="${c(tCol)}" stroke-width="0.2"/>
        <line x1="8.5" y1="${H - 8.5}" x2="12.5" y2="${H - 8.5}" stroke="${c(tCol)}" stroke-width="0.2"/>
        <rect x="7.5" y="6" width="4" height="0.9" fill="rgb(30,30,35)"/>
        <line x1="9" y1="6.9" x2="9" y2="9.5" stroke="rgb(30,30,35)" stroke-width="0.35"/>
        <circle cx="9" cy="10.2" r="0.55" fill="${c(gold)}"/>
      </svg>`;
    case "festa_donna":
      return `${svgStart}
        ${[
          [8, 10],
          [9.2, 9.2],
          [7, 9.5],
          [8.5, 8.3],
          [10, 10.5],
        ]
          .map(([fx, fy]) => `<circle cx="${fx}" cy="${fy}" r="0.9" fill="rgb(255,215,80)"/>`)
          .join("")}
        <line x1="8" y1="12" x2="8" y2="16" stroke="${c(green)}" stroke-width="0.35"/>
        <line x1="${W - 10}" y1="12" x2="${W - 9}" y2="15" stroke="${c(green)}" stroke-width="0.35"/>
        <line x1="${W - 10}" y1="12" x2="${W - 11}" y2="15" stroke="${c(green)}" stroke-width="0.35"/>
        <polygon points="${W - 10},12 ${W - 11.2},10 ${W - 8.8},10" fill="${c(green)}"/>
        <polygon points="${W - 7},${H - 10} ${W - 6},${H - 10.8} ${W - 5},${H - 10} ${W - 6},${H - 8.6}" fill="rgb(255,100,140)"/>
      </svg>`;
    case "nascita":
      return `${svgStart}
        <circle cx="9" cy="${H - 8}" r="2.2" fill="${c(sky)}"/>
        <rect x="10.5" y="${H - 8.5}" width="2.8" height="0.9" fill="${c(white)}"/>
        <circle cx="13.8" cy="${H - 8}" r="0.55" fill="${c(white)}"/>
        <circle cx="9" cy="${H - 8}" r="2.2" fill="none" stroke="${c(pink)}" stroke-width="0.25"/>
        ${[0, 1, 2, 3, 4]
          .map((i) => {
            const ang = (i / 5) * Math.PI * 2;
            return `<circle cx="${6 + Math.cos(ang) * 2.5}" cy="${9 + Math.sin(ang) * 2.5}" r="0.25" fill="${c(gold)}"/>`;
          })
          .join("")}
        <polygon points="${W - 9},${H - 9} ${W - 8.2},${H - 9.6} ${W - 7.4},${H - 9} ${W - 8.2},${H - 8}" fill="${c(pink)}"/>
      </svg>`;
    case "addio_festa":
      return `${svgStart}
        <polygon points="${W - 12},${H - 11} ${W - 11},${H - 11} ${W - 10.8},${H - 7.8} ${W - 13.2},${H - 7.8}" fill="rgb(255,240,250)"/>
        <polygon points="${W - 8.5},${H - 11} ${W - 7.5},${H - 11} ${W - 7.3},${H - 7.8} ${W - 9.7},${H - 7.8}" fill="rgb(255,240,250)"/>
        ${[0, 1, 2, 3, 4, 5, 6, 7, 8, 9]
          .map((i) => `<circle cx="${4 + (i % 5) * 2}" cy="${7 + Math.floor(i / 5) * 2}" r="0.3" fill="${c(pink)}"/>`)
          .join("")}
        <g stroke="${c(gold)}" fill="none" stroke-width="0.25">${[0, 60, 120, 180, 240, 300]
          .map((deg) => {
            const r = (deg * Math.PI) / 180;
            return `<line x1="10" y1="12" x2="${10 + Math.cos(r) * 2.8}" y2="${12 + Math.sin(r) * 2.8}"/>`;
          })
          .join("")}</g>
      </svg>`;
    case "inaugurazione":
      return `${svgStart}
        <path d="M${W - 12} 8 L${W - 8} 12 L${W - 6} 9" fill="none" stroke="${c(gold)}" stroke-width="0.35"/>
        <polygon points="7,8 8.5,8.5 8.2,10 5.8,10 5.5,8.5" fill="${c(red)}"/>
        <circle cx="7" cy="9.5" r="0.45" fill="${c(gold)}"/>
        <g stroke="${c(tCol)}" fill="none" stroke-width="0.25">${[0, 90, 180, 270]
          .map((deg) => {
            const r = (deg * Math.PI) / 180;
            return `<line x1="${W - 10}" y1="${H - 9}" x2="${W - 10 + Math.cos(r) * 2.5}" y2="${H - 9 + Math.sin(r) * 2.5}"/>`;
          })
          .join("")}</g>
      </svg>`;
    default:
      return "";
  }
}

/** Nomi completi ospiti con almeno cognome o nome (ordine come in griglia). */
function getTableGuestNamesOrdered(table) {
  const out = [];
  for (const g of table.guests) {
    const cognome = String(g.cognome || "").trim();
    const nome = String(g.nome || "").trim();
    if (!cognome && !nome) continue;
    out.push(`${cognome} ${nome}`.trim());
  }
  return out;
}

function getSegnatavoloPreviewTable() {
  if (state.tables.length) return state.tables[0];
  return {
    number: 1,
    customName: "Esempio",
    guests: [
      { cognome: "Rossi", nome: "Marco", menu: "adulto", note: "" },
      { cognome: "Bianchi", nome: "Anna", menu: "adulto", note: "" },
      { cognome: "Verdi", nome: "Luca", menu: "bambino", note: "" },
    ],
  };
}

function drawSegnatavoloPage(pdf, table, settings) {
  const theme = SEGNATAVOLO_THEMES.find((t) => t.id === settings.themeId) || SEGNATAVOLO_THEMES[0];
  const palette = settings._palette || getSegnatavoloPaletteResolved();
  const gi = settings.graphicIndex;
  const targetRect = settings._targetRect || null;
  const pageW = pdf.internal.pageSize.getWidth();
  const pageH = pdf.internal.pageSize.getHeight();
  const W = targetRect ? targetRect.w : pageW;
  const H = targetRect ? targetRect.h : pageH;
  const offsetX = targetRect ? targetRect.x : 0;
  const offsetY = targetRect ? targetRect.y : 0;
  const cx = W / 2;
  const hdrSettings = { headerMode: settings.headerMode ?? segnatavoloSettings.headerMode };
  const layout =
    settings._layout ||
    computeSegnatavoloBatchLayout(
      pdf,
      [table],
      { headerMode: hdrSettings.headerMode, showGuests: Boolean(settings.showGuests) },
      theme,
      W,
      H
    );
  const { marginX, innerW, innerTop, footY, titleSize, subSize, bodySize, heroMode, showGuests } = layout;
  const names = showGuests ? getTableGuestNamesOrdered(table) : [];
  const headers = getSegnatavoloHeaderParts(table, hdrSettings);
  const mainTitle = headers.mainTitle;
  const subtitle = headers.subtitle;
  const foot = (state.eventName || "").trim() || "Evento";

  const titleLineH = (sz) => segnatavoloTitleLineHeightMm(sz);

  function drawFooter() {
    pdf.setFont(theme.footFamily, theme.footStyle);
    pdf.setFontSize(theme.footSize);
    pdf.setTextColor(...palette.foot);
    pdf.text(foot, cx, footY, { align: "center" });
    pdf.setTextColor(0, 0, 0);
  }

  if (targetRect) {
    pdf.internal.write(`q 1 0 0 1 ${offsetX} ${offsetY} cm`);
  }
  try {
  const fid = settings.festiveMarginId || null;
  if (fid) drawSegnatavoloFestiveMarginArt(pdf, W, H, palette, fid);
  else drawSegnatavoloGraphic(pdf, W, H, palette.deco, gi);

  if (heroMode) {
    const heroH = measureSegnatavoloHeroHeaderMm(pdf, theme, innerW, mainTitle, subtitle, titleSize, subSize);
    let y = innerTop + Math.max(0, (layout.contentH - heroH) / 2);
    pdf.setFont(theme.titleFamily, theme.titleStyle);
    pdf.setFontSize(titleSize);
    pdf.setTextColor(...palette.title);
    const titleLines = pdf.splitTextToSize(mainTitle, innerW);
    for (const line of titleLines) {
      pdf.text(line, cx, y, { align: "center" });
      y += titleLineH(titleSize);
    }
    if (subtitle) {
      pdf.setFont(theme.subFamily, theme.subStyle);
      pdf.setFontSize(subSize);
      pdf.setTextColor(...palette.subtitle);
      const subLines = pdf.splitTextToSize(subtitle, innerW);
      for (const line of subLines) {
        pdf.text(line, cx, y + 0.4, { align: "center" });
        y += titleLineH(subSize);
      }
    }
    drawFooter();
    pdf.setTextColor(0, 0, 0);
    return;
  }

  let y = innerTop + 3.5;
  pdf.setFont(theme.titleFamily, theme.titleStyle);
  pdf.setFontSize(titleSize);
  pdf.setTextColor(...palette.title);
  const titleLines = pdf.splitTextToSize(mainTitle, innerW);
  for (const line of titleLines) {
    pdf.text(line, cx, y, { align: "center" });
    y += titleLineH(titleSize);
  }
  if (subtitle) {
    pdf.setFont(theme.subFamily, theme.subStyle);
    pdf.setFontSize(subSize);
    pdf.setTextColor(...palette.subtitle);
    const subLines = pdf.splitTextToSize(subtitle, innerW);
    for (const line of subLines) {
      pdf.text(line, cx, y + 0.5, { align: "center" });
      y += titleLineH(subSize);
    }
  } else {
    y += 2;
  }
  y += 3;
  pdf.setDrawColor(...palette.deco);
  pdf.setLineWidth(0.35);
  pdf.line(marginX, y, W - marginX, y);
  y += 6.5;

  if (!names.length) {
    pdf.setFont(theme.bodyFamily, "italic");
    pdf.setFontSize(Math.max(5.5, bodySize * 0.88));
    pdf.setTextColor(...palette.subtitle);
    pdf.text("Nessun ospite in elenco.", cx, y + 4, { align: "center" });
    drawFooter();
    pdf.setTextColor(0, 0, 0);
    return;
  }

  pdf.setFont(theme.bodyFamily, theme.bodyStyle);
  pdf.setFontSize(bodySize);
  pdf.setTextColor(...palette.body);
  const lineH = Math.max(1.9, bodySize * 0.34);
  for (const n of names) {
    const lines = pdf.splitTextToSize(n, innerW);
    for (const line of lines) {
      pdf.text(line, cx, y, { align: "center" });
      y += lineH;
    }
  }
  drawFooter();
  pdf.setTextColor(0, 0, 0);
  } finally {
    if (targetRect) {
      pdf.internal.write("Q");
    }
  }
}

function writePdfLines(pdf, lines, options = {}) {
  const margin = options.margin ?? 10;
  const lineHeight = options.lineHeight ?? 5;
  const pageW = pdf.internal.pageSize.getWidth();
  const pageH = pdf.internal.pageSize.getHeight();
  const usableW = pageW - margin * 2;
  let y = options.startY ?? 12;

  for (const rawLine of lines) {
    const wrapped = pdf.splitTextToSize(rawLine, usableW);
    for (const line of wrapped) {
      if (y > pageH - margin) {
        pdf.addPage();
        y = margin;
      }
      pdf.text(line, margin, y);
      y += lineHeight;
    }
  }
}

function getGroupedDietNotesByMenu() {
  const mapAdulti = new Map();
  const mapBambini = new Map();
  for (const table of state.tables) {
    for (const g of table.guests) {
      const cognome = String(g.cognome || "").trim();
      const nome = String(g.nome || "").trim();
      if (!cognome && !nome) continue;
      const note = String(g.note || "").trim();
      if (!note) continue;
      const targetMap = g.menu === "bambino" ? mapBambini : mapAdulti;
      const key = note.toLocaleLowerCase("it");
      const curr = targetMap.get(key);
      if (curr) {
        curr.count += 1;
      } else {
        targetMap.set(key, { label: note, count: 1 });
      }
    }
  }
  const byCountThenLabel = (a, b) => b.count - a.count || a.label.localeCompare(b.label, "it");
  return {
    adulti: [...mapAdulti.values()].sort(byCountThenLabel),
    bambini: [...mapBambini.values()].sort(byCountThenLabel),
  };
}

els.eventName.value = state.eventName;
els.eventName.addEventListener("input", () => {
  if (!canEditEventName()) return;
  state.eventName = els.eventName.value;
  saveState();
  if (isSegnatavoloOrMenuStudioView()) {
    updateSegnatavoloPreview();
  }
});

els.addTableBtn.addEventListener("click", () => {
  if (!canEditTablesArea()) return;
  addTable();
});

if (els.showTablesAreaBtn) {
  els.showTablesAreaBtn.addEventListener("click", () => {
    setMainAreaView("tables");
  });
}

if (els.showMapAreaBtn) {
  els.showMapAreaBtn.addEventListener("click", () => {
    if (!canAccessArea("map")) return;
    setMainAreaView("map");
  });
}

if (els.showAcceptanceAreaBtn) {
  els.showAcceptanceAreaBtn.addEventListener("click", () => {
    if (!canAccessArea("acceptance")) return;
    setMainAreaView("acceptance");
    if (els.acceptanceSearchInput) {
      els.acceptanceSearchInput.focus();
    }
  });
}

if (els.openSegnatavoloAreaBtn) {
  els.openSegnatavoloAreaBtn.addEventListener("click", () => {
    if (!canAccessArea("segnatavolo")) return;
    closeAllDropdowns();
    setMainAreaView("segnatavolo");
    updateSegnatavoloPreview();
  });
}

if (els.openMenuBookletAreaBtn) {
  els.openMenuBookletAreaBtn.addEventListener("click", () => {
    if (!canAccessArea("menu")) return;
    closeAllDropdowns();
    setMainAreaView("menu");
    renderMenuBookletPreview();
  });
}

if (els.acceptanceSearchInput) {
  els.acceptanceSearchInput.addEventListener("input", () => {
    acceptanceUiState.query = els.acceptanceSearchInput.value;
    renderAcceptanceArea();
  });
}

els.sortTablesBtn.addEventListener("click", () => {
  if (!canEditTablesArea()) return;
  closeAllDropdowns();
  sortTablesAscending();
});

els.renumberTablesBtn.addEventListener("click", () => {
  if (!canEditTablesArea()) return;
  closeAllDropdowns();
  renumberTablesSmart();
});

if (els.tableActionsMenuBtn && els.tableActionsMenu) {
  els.tableActionsMenuBtn.addEventListener("click", (e) => {
    e.stopPropagation();
    toggleDropdown(els.tableActionsMenu, els.tableActionsMenuBtn);
  });
}

if (els.topActionsMenuBtn && els.topActionsMenu) {
  els.topActionsMenuBtn.addEventListener("click", (e) => {
    e.stopPropagation();
    toggleDropdown(els.topActionsMenu, els.topActionsMenuBtn);
  });
}

if (els.topActionsMenu) {
  els.topActionsMenu.addEventListener("click", (e) => {
    const toggleBtn = e.target.closest("[data-submenu-toggle]");
    if (!toggleBtn) return;
    const key = toggleBtn.dataset.submenuToggle || "";
    const target = els.topActionsMenu.querySelector(`[data-submenu="${key}"]`);
    if (!target) return;
    const willOpen = target.hidden;
    els.topActionsMenu.querySelectorAll(".top-actions-submenu").forEach((el) => {
      el.hidden = true;
    });
    els.topActionsMenu.querySelectorAll("[data-submenu-toggle]").forEach((btn) => {
      btn.setAttribute("aria-expanded", "false");
    });
    target.hidden = !willOpen;
    toggleBtn.setAttribute("aria-expanded", willOpen ? "true" : "false");
  });
}

let pendingImportKind = "";
const CLIENT_FILE_EXT = ".evtcliente";
const FULL_FILE_EXT = ".evtstruttura";

function formatExportTimestampForFile() {
  const d = new Date();
  const pad = (n) => String(n).padStart(2, "0");
  const yyyy = d.getFullYear();
  const mm = pad(d.getMonth() + 1);
  const dd = pad(d.getDate());
  const hh = pad(d.getHours());
  const mi = pad(d.getMinutes());
  return `${yyyy}${mm}${dd}-${hh}${mi}`;
}

function getImportAcceptByKind(kind) {
  return kind === "cliente"
    ? `${CLIENT_FILE_EXT},application/json,.json`
    : `${FULL_FILE_EXT},application/json,.json`;
}

function fileLooksLikeExpectedKind(fileName, kind) {
  const name = String(fileName || "").toLowerCase();
  if (kind === "cliente") {
    return name.endsWith(CLIENT_FILE_EXT) || name.endsWith(".json");
  }
  return name.endsWith(FULL_FILE_EXT) || name.endsWith(".json");
}

function promptImport(kind) {
  if (!els.importInput) return;
  pendingImportKind = kind;
  closeAllDropdowns();
  els.importInput.setAttribute("accept", getImportAcceptByKind(kind));
  els.importInput.click();
}

if (els.importClientBtn && els.importInput) {
  els.importClientBtn.addEventListener("click", () => {
    if (appRole === "staff") return;
    promptImport("cliente");
  });
}
if (els.importFullBtn && els.importInput) {
  els.importFullBtn.addEventListener("click", () => {
    if (appRole === "cliente") return;
    promptImport("completo");
  });
}

document.addEventListener("click", (e) => {
  const t = e.target;
  if (t && t.closest && t.closest("[data-close-dropdown='true']")) {
    closeAllDropdowns();
    return;
  }
  if (els.tableActionsMenu && els.tableActionsMenu.contains(t)) return;
  if (els.topActionsMenu && els.topActionsMenu.contains(t)) return;
  if (els.tableActionsDropdown && els.tableActionsDropdown.contains(t)) return;
  if (els.topActionsDropdown && els.topActionsDropdown.contains(t)) return;
  if (t && t.closest && t.closest(".plan-visibility")) return;
  closeAllDropdowns();
});

if (els.addPlanPanelBtn) {
  els.addPlanPanelBtn.addEventListener("click", () => {
    if (!canEditMapArea()) return;
    ensureFloorPlansState();
    const plan = createFloorPlan("");
    for (const table of state.tables) {
      plan.markerPositions[table.id] = { x: 50, y: 50 };
    }
    state.floorPlans.push(plan);
    renderFloorPlans();
    saveState();
  });
}

if (els.floorPlansContainer) {
  els.floorPlansContainer.addEventListener("change", async (e) => {
    if (!canEditMapArea()) return;
    const visCb = e.target.closest(".plan-visibility-checkbox");
    if (visCb) {
      const card = visCb.closest(".floor-plan-card");
      const pid = card ? card.dataset.planId : "";
      const tableId = visCb.dataset.tableId || "";
      const p = state.floorPlans.find((pl) => pl.id === pid);
      if (p && tableId) {
        const set = new Set(p.hiddenTableIds || []);
        if (visCb.checked) set.delete(tableId);
        else set.add(tableId);
        p.hiddenTableIds = [...set];
        saveState();
        preservePlanVisibilityMenuPlanId = pid;
        renderFloorPlans();
      }
      return;
    }

    const input = e.target.closest(".plan-upload-input");
    if (!input) return;
    const card = input.closest(".floor-plan-card");
    const planId = card ? card.dataset.planId : "";
    const file = input.files && input.files[0];
    if (!planId || !file) return;
    const plan = state.floorPlans.find((p) => p.id === planId);
    if (!plan) return;
    try {
      plan.imageDataUrl = await readFileAsDataUrl(file);
      renderFloorPlans();
      saveState();
    } catch (_) {
      alert("Impossibile leggere il file. Prova con PNG o JPG.");
    }
    input.value = "";
  });

  els.floorPlansContainer.addEventListener("click", (e) => {
    if (!canEditMapArea()) return;
    const card = e.target.closest(".floor-plan-card");
    if (!card) return;
    const planId = card.dataset.planId || "";
    const plan = state.floorPlans.find((p) => p.id === planId);
    if (!plan) return;

    const visToggle = e.target.closest("[data-plan-visibility-toggle]");
    if (visToggle) {
      const panel = card.querySelector(".plan-visibility-panel");
      if (!panel) return;
      const willOpen = panel.hidden;
      closeAllDropdowns();
      if (willOpen) {
        panel.hidden = false;
        visToggle.setAttribute("aria-expanded", "true");
      }
      return;
    }

    if (e.target.closest("[data-plan-visibility-show-all]")) {
      plan.hiddenTableIds = [];
      saveState();
      preservePlanVisibilityMenuPlanId = planId;
      renderFloorPlans();
      return;
    }

    if (e.target.closest("[data-plan-export-pdf='true']")) {
      exportPlanPdf(planId);
      return;
    }

    if (e.target.closest("[data-plan-clear-btn='true']")) {
      plan.imageDataUrl = "";
      renderFloorPlans();
      saveState();
      return;
    }
    if (e.target.closest("[data-plan-remove-btn='true']")) {
      if (state.floorPlans.length <= 1) {
        alert("Deve rimanere almeno un riquadro piantina.");
        return;
      }
      state.floorPlans = state.floorPlans.filter((p) => p.id !== planId);
      renderFloorPlans();
      saveState();
    }
  });

  document.addEventListener("mousedown", (e) => {
    if (e.target.closest(".plan-visibility")) return;
    document.querySelectorAll(".plan-visibility-panel").forEach((panel) => {
      panel.hidden = true;
    });
    document.querySelectorAll("[data-plan-visibility-toggle]").forEach((btn) => {
      btn.setAttribute("aria-expanded", "false");
    });
  });
}

function downloadJsonFile(filename, payload) {
  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  a.click();
  URL.revokeObjectURL(a.href);
}

function buildClientExportPayload() {
  return {
    app: "evento-tavoli",
    fileType: "cliente",
    version: 1,
    exportedAt: new Date().toISOString(),
    data: {
      eventName: state.eventName,
      tables: state.tables,
    },
  };
}

function buildFullExportPayload() {
  return {
    app: "evento-tavoli",
    fileType: "struttura",
    version: 1,
    exportedAt: new Date().toISOString(),
    data: {
      eventName: state.eventName,
      floorPlanDataUrl: state.floorPlanDataUrl,
      tables: state.tables,
      markerPositions: state.markerPositions,
      floorPlans: state.floorPlans,
      segnatavoloSettings,
      menuBookletState,
    },
  };
}

if (els.exportClientBtn) {
  els.exportClientBtn.addEventListener("click", () => {
    if (appRole === "staff") return;
    closeAllDropdowns();
    const stamp = formatExportTimestampForFile();
    downloadJsonFile(`evento-tavoli-cliente-${stamp}${CLIENT_FILE_EXT}`, buildClientExportPayload());
  });
}

if (els.exportFullBtn) {
  els.exportFullBtn.addEventListener("click", () => {
    if (appRole !== "struttura") return;
    closeAllDropdowns();
    const stamp = formatExportTimestampForFile();
    downloadJsonFile(`evento-tavoli-struttura-${stamp}${FULL_FILE_EXT}`, buildFullExportPayload());
  });
}

els.exportGuestsBtn.addEventListener("click", () => {
  closeAllDropdowns();
  const rows = getSortedParticipants();
  if (!rows.length) {
    alert("Non ci sono partecipanti da esportare.");
    return;
  }
  const pdf = createPdfDocument("portrait");
  if (!pdf) {
    alert("Libreria PDF non disponibile.");
    return;
  }
  pdf.setFontSize(14);
  pdf.text(`Lista partecipanti A-Z - ${state.eventName || "Evento"}`, 10, 12);
  pdf.setFontSize(10);
  const lines = [];
  rows.forEach((row, idx) => {
    const fullName = `${row.cognome} ${row.nome}`.trim();
    const tavoloNome = row.tavoloNome ? ` - ${row.tavoloNome}` : "";
    lines.push(`${idx + 1}. ${fullName} - Tavolo ${row.tavoloNumero}${tavoloNome}`);
  });
  writePdfLines(pdf, lines, { startY: 20, margin: 10, lineHeight: 5 });
  pdf.save("lista-partecipanti-alfabetica.pdf");
});

els.exportKitchenBtn.addEventListener("click", () => {
  closeAllDropdowns();
  const pdf = createPdfDocument("portrait");
  if (!pdf) {
    alert("Libreria PDF non disponibile.");
    return;
  }
  let totAdulti = 0;
  let totBambini = 0;
  const lines = [];
  const orderedTables = [...state.tables].sort((a, b) => Number(a.number) - Number(b.number));
  for (const table of orderedTables) {
    const counts = getTableMenuCounts(table);
    totAdulti += counts.adulti;
    totBambini += counts.bambini;
  }

  lines.push(`Evento: ${state.eventName || "Senza nome"}`);
  lines.push(`Data esportazione: ${new Date().toLocaleString("it-IT")}`);
  lines.push("");
  lines.push("TOTALE MENÙ");
  lines.push(`- Menù adulti: ${totAdulti}`);
  lines.push(`- Menù bambini: ${totBambini}`);
  const grouped = getGroupedDietNotesByMenu();
  lines.push("");
  lines.push("INTOLLERANZE/VARIAZIONI RAGGRUPPATE");
  lines.push("- Adulti:");
  if (!grouped.adulti.length) {
    lines.push("- Nessuna specifica inserita");
  } else {
    grouped.adulti.forEach((g) => lines.push(`  - ${g.label}: ${g.count}`));
  }
  lines.push("- Bambini:");
  if (!grouped.bambini.length) {
    lines.push("- Nessuna specifica inserita");
  } else {
    grouped.bambini.forEach((g) => lines.push(`  - ${g.label}: ${g.count}`));
  }

  pdf.setFontSize(14);
  pdf.text(`Riepilogo cucina - ${state.eventName || "Evento"}`, 10, 12);
  pdf.setFontSize(10);
  writePdfLines(pdf, lines, { startY: 20, margin: 10, lineHeight: 5 });
  pdf.save("riepilogo-cucina.pdf");
});

function applySegnatavoloPreviewDeco(targetDecoEl) {
  if (!targetDecoEl) return;
  const fid = segnatavoloSettings.festiveMarginId || null;
  const base = "segnatavolo-preview__deco ";
  if (fid) {
    const pal = getSegnatavoloPaletteResolved();
    const inner = buildSegnatavoloFestivePreviewInnerHTML(pal, fid);
    if (inner) {
      targetDecoEl.className = base + "segnatavolo-festive-wrap";
      targetDecoEl.innerHTML = inner;
      return;
    }
  }
  const gi = segnatavoloSettings.graphicIndex;
  if (gi === 2) {
    targetDecoEl.className = base + "segnatavolo-deco--2";
    targetDecoEl.innerHTML =
      '<span class="segnatavolo-deco__bl"></span><span class="segnatavolo-deco__br"></span>';
  } else if (gi === 4) {
    targetDecoEl.className = base + "segnatavolo-deco--4";
    targetDecoEl.innerHTML =
      '<span class="segnatavolo-deco__c3"></span><span class="segnatavolo-deco__c4"></span>';
  } else if (gi === 5) {
    targetDecoEl.className = base + "segnatavolo-deco--5";
    targetDecoEl.innerHTML =
      '<span class="segnatavolo-deco__d3"></span><span class="segnatavolo-deco__d4"></span>';
  } else {
    targetDecoEl.innerHTML = "";
    targetDecoEl.className = base + (gi === 0 ? "" : `segnatavolo-deco--${gi}`);
  }
}

function setSegnatavoloPreviewDecoDom() {
  applySegnatavoloPreviewDeco(els.segnatavoloPreviewDeco);
}

function refreshSegnatavoloChipSelection() {
  if (els.segnatavoloThemeChips) {
    els.segnatavoloThemeChips.querySelectorAll(".segnatavolo-chip").forEach((b) => {
      b.classList.toggle("is-selected", b.dataset.themeId === segnatavoloSettings.themeId);
    });
  }
  if (els.segnatavoloPaletteChips) {
    els.segnatavoloPaletteChips.querySelectorAll(".segnatavolo-chip").forEach((b) => {
      b.classList.toggle("is-selected", b.dataset.paletteId === segnatavoloSettings.paletteId);
    });
  }
  if (els.segnatavoloGraphicChips) {
    els.segnatavoloGraphicChips.querySelectorAll(".segnatavolo-chip").forEach((b) => {
      b.classList.toggle("is-selected", Number(b.dataset.graphicIndex) === segnatavoloSettings.graphicIndex);
    });
  }
  if (els.segnatavoloOccasionChips) {
    const occ = getSegnatavoloOccasionChipId();
    els.segnatavoloOccasionChips.querySelectorAll(".segnatavolo-chip").forEach((b) => {
      b.classList.toggle("is-selected", b.dataset.occasionId === occ);
    });
  }
}

function updateSegnatavoloCustomColorsVisibility() {
  if (els.segnatavoloCustomColors) {
    els.segnatavoloCustomColors.hidden = segnatavoloSettings.paletteId !== "p10";
  }
}

function fitSegnatavoloPreviewNames() {
  fitSegnatavoloPreviewNamesIn(els.segnatavoloPreviewNames);
}

function fitSegnatavoloPreviewNamesIn(list) {
  if (!list) return;
  const items = Array.from(list.querySelectorAll("li"));
  if (!items.length) return;
  list.style.height = "";
  list.style.flex = "1";
  list.style.transform = "none";
  list.style.transformOrigin = "top center";
  list.style.width = "100%";
  list.style.overflow = "hidden";

  list.style.columnCount = "1";
  list.style.columnGap = "0.7rem";
  list.classList.remove("is-two-columns");

  const basePx = parseFloat(items[0].style.fontSize || "14") || 14;
  const minPx = 1;

  // Anteprima fedele al PDF: sempre una sola colonna, riduce solo il font.
  for (let size = basePx; size >= minPx; size -= 0.25) {
    items.forEach((li) => {
      li.style.fontSize = `${size}px`;
      li.style.lineHeight = "1.0";
    });
    if (list.scrollHeight <= list.clientHeight + 1) {
      return;
    }
  }

  // Fallback estremo: non perdere mai nomi in anteprima.
  const scale = Math.max(0.2, (list.clientHeight - 1) / Math.max(1, list.scrollHeight));
  list.style.transform = `scale(${scale})`;
  list.style.width = `${100 / scale}%`;
}

function scheduleFitSegnatavoloStudioSnapPreviews() {
  requestAnimationFrame(() => {
    requestAnimationFrame(() => {
      fitSegnatavoloStudioSnapPreviews();
    });
  });
}

function scheduleFitStudioMenuBookletScale() {
  requestAnimationFrame(() => {
    requestAnimationFrame(() => {
      fitStudioMenuBookletScale();
    });
  });
}

function scheduleFitSegnatavoloControlSample() {
  requestAnimationFrame(() => {
    requestAnimationFrame(() => {
      fitSegnatavoloControlSample();
    });
  });
}

/** Adatta ogni anteprima tavolo al riquadro visibile (una schermata per tavolo, scroll solo tra tavoli). */
function fitSegnatavoloStudioSnapPreviews() {
  if (!els.segnatavoloAreaPanel || els.segnatavoloAreaPanel.hidden || !els.segnatavoloPreviewAllList) return;
  els.segnatavoloPreviewAllList.querySelectorAll(".segnatavolo-preview-scale-wrap").forEach((wrap) => {
    const card = wrap.querySelector(".segnatavolo-preview");
    if (!card) return;
    card.style.transform = "";
    const pad = 6;
    const ww = Math.max(0, wrap.clientWidth - pad);
    const wh = Math.max(0, wrap.clientHeight - pad);
    const rw = card.offsetWidth;
    const rh = card.offsetHeight;
    if (rw < 8 || rh < 8) return;
    const s = Math.min(1, ww / rw, wh / rh) * 0.995;
    if (s < 0.998) {
      card.style.transformOrigin = "center center";
      card.style.transform = `scale(${s})`;
    }
  });
}

/** Riduce l’anteprima libricino se non entra nel canvas. */
function fitStudioMenuBookletScale() {
  if (!els.menuBookletAreaPanel || els.menuBookletAreaPanel.hidden) return;
  const scroll = document.querySelector(".studio-canvas__scroll--menu-pages");
  const stack = els.menuBookletPreviewStack;
  if (!scroll || !stack) return;
  stack.style.transform = "";
}

/** Mostra sempre interamente il foglio mockup nel riquadro, con proporzioni esatte. */
function fitSegnatavoloControlSample() {
  if (!els.segnatavoloAreaPanel || els.segnatavoloAreaPanel.hidden) return;
  const wrap = els.segnatavoloControlSampleWrap;
  const card = els.segnatavoloControlSampleCard;
  if (!wrap || !card) return;
  card.style.transform = "";
  const pad = 6;
  const ww = Math.max(0, wrap.clientWidth - pad);
  const wh = Math.max(0, wrap.clientHeight - pad);
  const rw = card.offsetWidth;
  const rh = card.offsetHeight;
  if (ww < 8 || wh < 8 || rw < 8 || rh < 8) return;
  const s = Math.min(1, ww / rw, wh / rh) * 0.995;
  if (s < 0.998) {
    card.style.transformOrigin = "center center";
    card.style.transform = `scale(${s})`;
  }
}

function renderAllSegnatavoloPreviews(theme, pal, previewRatio, batchLayoutOpt) {
  if (!els.segnatavoloPreviewAllList) return;
  els.segnatavoloPreviewAllList.innerHTML = "";
  els.segnatavoloPreviewAllList.className = "segnatavolo-preview-all-list segnatavolo-preview-all-list--snap";
  const [paperMmW, paperMmH] = getSegnatavoloPaperSizeMm();
  const ordered = [...state.tables].sort((a, b) => Number(a.number) - Number(b.number));
  if (!ordered.length) {
    const empty = document.createElement("div");
    empty.className = "segnatavolo-preview-all-item segnatavolo-preview-all-item--empty menu-booklet-empty";
    empty.textContent = "Nessun tavolo disponibile per l'anteprima.";
    els.segnatavoloPreviewAllList.appendChild(empty);
    scheduleFitSegnatavoloStudioSnapPreviews();
    return;
  }

  const batchLayout = batchLayoutOpt != null ? batchLayoutOpt : getSegnatavoloLayoutForCurrentState(ordered);
  const heroMode = batchLayout ? batchLayout.heroMode : !segnatavoloSettings.showGuests;
  const pts = getSegnatavoloPreviewTypePts(theme, batchLayout);
  const titlePx = segnatavoloPdfPtToCssPx(pts.titlePt);
  const subPx = segnatavoloPdfPtToCssPx(pts.subPt);
  const bodyPx = segnatavoloPdfPtToCssPx(pts.bodyPt);
  const emptyPx = segnatavoloPdfPtToCssPx(pts.emptyBodyPt);
  const footPx = segnatavoloPdfPtToCssPx(pts.footPt);

  ordered.forEach((table) => {
    const item = document.createElement("div");
    item.className = "segnatavolo-preview-all-item";

    const cap = document.createElement("div");
    cap.className = "segnatavolo-preview-all-item__title";
    cap.textContent = `Anteprima Tavolo ${table.number}`;
    item.appendChild(cap);

    const card = document.createElement("div");
    card.className = "segnatavolo-preview";
    card.style.setProperty("--seg-deco", rgbToCss(pal.deco));
    card.style.setProperty("--seg-preview-ratio", String(previewRatio));
    card.style.setProperty("--seg-paper-w-mm", String(paperMmW));
    card.style.setProperty("--seg-paper-h-mm", String(paperMmH));
    card.style.width = `${segnatavoloMmToCssPx(paperMmW)}px`;
    card.style.height = `${segnatavoloMmToCssPx(paperMmH)}px`;

    const deco = document.createElement("div");
    deco.className = "segnatavolo-preview__deco";
    card.appendChild(deco);
    applySegnatavoloPreviewDeco(deco);

    const content = document.createElement("div");
    content.className = "segnatavolo-preview__content";
    if (heroMode) content.classList.add("segnatavolo-preview__content--hero");

    const headers = getSegnatavoloHeaderParts(table);
    const title = document.createElement("div");
    title.className = "segnatavolo-preview__title";
    title.style.fontFamily = theme.previewTitle;
    title.style.fontSize = `${titlePx}px`;
    title.style.color = rgbToCss(pal.title);
    title.textContent = headers.mainTitle;

    const subtitle = document.createElement("div");
    subtitle.className = "segnatavolo-preview__subtitle";
    subtitle.style.fontFamily = theme.previewBody;
    subtitle.style.fontSize = `${subPx}px`;
    subtitle.style.color = rgbToCss(pal.subtitle);
    subtitle.textContent = headers.subtitle || "";
    subtitle.hidden = !headers.subtitle;

    if (heroMode) {
      const cluster = document.createElement("div");
      cluster.className = "segnatavolo-preview__hero-cluster";
      cluster.appendChild(title);
      cluster.appendChild(subtitle);
      content.appendChild(cluster);
    } else {
      content.appendChild(title);
      content.appendChild(subtitle);
      const rule = document.createElement("div");
      rule.className = "segnatavolo-preview__rule";
      rule.style.background = rgbToCss(pal.deco);
      content.appendChild(rule);
    }

    const names = document.createElement("ul");
    names.className = "segnatavolo-preview__names";
    if (!heroMode && segnatavoloSettings.showGuests) {
      getTableGuestNamesOrdered(table).forEach((n) => {
        const li = document.createElement("li");
        li.textContent = n;
        li.style.fontFamily = theme.previewBody;
        li.style.fontSize = `${bodyPx}px`;
        li.style.color = rgbToCss(pal.body);
        names.appendChild(li);
      });
    }
    if (!heroMode) content.appendChild(names);

    const empty = document.createElement("div");
    empty.className = "segnatavolo-preview__empty";
    empty.style.color = rgbToCss(pal.subtitle);
    empty.style.fontSize = `${emptyPx}px`;
    empty.style.fontFamily = theme.previewBody;
    empty.textContent = "Nessun ospite in elenco.";
    empty.hidden = heroMode || !segnatavoloSettings.showGuests || names.children.length > 0;
    content.appendChild(empty);

    const foot = document.createElement("div");
    foot.className = "segnatavolo-preview__foot";
    foot.style.color = rgbToCss(pal.foot);
    foot.style.fontFamily = theme.previewBody;
    foot.style.fontSize = `${footPx}px`;
    foot.textContent = (state.eventName || "").trim() || "Evento";
    content.appendChild(foot);

    card.appendChild(content);
    const scaleWrap = document.createElement("div");
    scaleWrap.className = "segnatavolo-preview-scale-wrap";
    scaleWrap.appendChild(card);
    item.appendChild(scaleWrap);
    els.segnatavoloPreviewAllList.appendChild(item);
  });
  scheduleFitSegnatavoloStudioSnapPreviews();
}

function renderSegnatavoloControlSample(theme, pal, batchLayoutOpt) {
  if (!els.segnatavoloControlSampleCard) return;
  const ratio = 148 / 210; // Anteprima standard fissa: mostra solo resa grafica, non impaginazione carta.
  const ordered = [...state.tables].sort((a, b) => Number(a.number) - Number(b.number));
  const table = ordered[0] || { number: 1, customName: "", guests: [] };
  const pts = getSegnatavoloPreviewTypePts(theme, null);
  const titlePx = segnatavoloPdfPtToCssPx(pts.titlePt);
  const subPx = segnatavoloPdfPtToCssPx(pts.subPt);
  const bodyPx = segnatavoloPdfPtToCssPx(pts.bodyPt);
  const emptyPx = segnatavoloPdfPtToCssPx(pts.emptyBodyPt);
  const footPx = segnatavoloPdfPtToCssPx(pts.footPt);
  const headers = getSegnatavoloHeaderParts(table);

  els.segnatavoloControlSampleCard.style.setProperty("--sample-ratio", String(ratio));
  els.segnatavoloControlSampleCard.style.setProperty("--seg-deco", rgbToCss(pal.deco));
  els.segnatavoloControlSampleCard.style.width = "";
  els.segnatavoloControlSampleCard.style.height = "";
  if (els.segnatavoloControlSampleDeco) applySegnatavoloPreviewDeco(els.segnatavoloControlSampleDeco);

  if (els.segnatavoloControlSampleTitle) {
    els.segnatavoloControlSampleTitle.textContent = headers.mainTitle || "Tavolo 1";
    els.segnatavoloControlSampleTitle.style.fontFamily = theme.previewTitle;
    els.segnatavoloControlSampleTitle.style.fontSize = `${titlePx}px`;
    els.segnatavoloControlSampleTitle.style.color = rgbToCss(pal.title);
  }
  if (els.segnatavoloControlSampleSubtitle) {
    els.segnatavoloControlSampleSubtitle.hidden = !headers.subtitle;
    els.segnatavoloControlSampleSubtitle.textContent = headers.subtitle || "";
    els.segnatavoloControlSampleSubtitle.style.fontFamily = theme.previewBody;
    els.segnatavoloControlSampleSubtitle.style.fontSize = `${subPx}px`;
    els.segnatavoloControlSampleSubtitle.style.color = rgbToCss(pal.subtitle);
  }

  if (els.segnatavoloControlSampleNames) {
    els.segnatavoloControlSampleNames.innerHTML = "";
    els.segnatavoloControlSampleNames.hidden = !segnatavoloSettings.showGuests;
    if (segnatavoloSettings.showGuests) {
      const names = getTableGuestNamesOrdered(table);
      const list = names.length ? names.slice(0, 6) : ["Mario Rossi", "Anna Bianchi", "Luca Verdi"];
      list.forEach((n) => {
        const li = document.createElement("li");
        li.textContent = n;
        li.style.fontFamily = theme.previewBody;
        li.style.fontSize = `${bodyPx}px`;
        li.style.color = rgbToCss(pal.body);
        els.segnatavoloControlSampleNames.appendChild(li);
      });
    }
  }
  if (els.segnatavoloControlSampleEmpty) {
    const showEmpty = !segnatavoloSettings.showGuests || (getTableGuestNamesOrdered(table).length === 0);
    els.segnatavoloControlSampleEmpty.hidden = !showEmpty;
    els.segnatavoloControlSampleEmpty.style.fontFamily = theme.previewBody;
    els.segnatavoloControlSampleEmpty.style.fontSize = `${emptyPx}px`;
    els.segnatavoloControlSampleEmpty.style.color = rgbToCss(pal.subtitle);
  }
  if (els.segnatavoloControlSampleFoot) {
    els.segnatavoloControlSampleFoot.textContent = (state.eventName || "").trim() || "Evento";
    els.segnatavoloControlSampleFoot.style.fontFamily = theme.previewBody;
    els.segnatavoloControlSampleFoot.style.fontSize = `${footPx}px`;
    els.segnatavoloControlSampleFoot.style.color = rgbToCss(pal.foot);
  }
  scheduleFitSegnatavoloControlSample();
}

function updateSegnatavoloPreview() {
  const theme = getSegnatavoloTheme();
  const pal = getSegnatavoloPaletteResolved();
  const [paperW, paperH] = getSegnatavoloPaperSizeMm();
  const previewRatio = paperH > 0 ? paperW / paperH : 148 / 210;
  const batchLayout = getSegnatavoloLayoutForCurrentState();
  const pts = getSegnatavoloPreviewTypePts(theme, batchLayout);
  const titlePx = segnatavoloPdfPtToCssPx(pts.titlePt);
  const subPx = segnatavoloPdfPtToCssPx(pts.subPt);
  const bodyPx = segnatavoloPdfPtToCssPx(pts.bodyPt);
  const emptyPx = segnatavoloPdfPtToCssPx(pts.emptyBodyPt);
  const footPx = segnatavoloPdfPtToCssPx(pts.footPt);
  if (els.segnatavoloPreview && els.segnatavoloPreviewTitle) {
    const table = getSegnatavoloPreviewTable();
    els.segnatavoloPreview.style.setProperty("--seg-deco", rgbToCss(pal.deco));
    els.segnatavoloPreview.style.setProperty("--seg-preview-ratio", String(previewRatio));
    els.segnatavoloPreview.style.setProperty("--seg-paper-w-mm", String(paperW));
    els.segnatavoloPreview.style.setProperty("--seg-paper-h-mm", String(paperH));
    els.segnatavoloPreview.style.width = `${segnatavoloMmToCssPx(paperW)}px`;
    els.segnatavoloPreview.style.height = `${segnatavoloMmToCssPx(paperH)}px`;

    const headers = getSegnatavoloHeaderParts(table);
    els.segnatavoloPreviewTitle.style.fontFamily = theme.previewTitle;
    els.segnatavoloPreviewTitle.style.fontSize = `${titlePx}px`;
    els.segnatavoloPreviewTitle.style.color = rgbToCss(pal.title);
    els.segnatavoloPreviewTitle.textContent = headers.mainTitle;

    if (headers.subtitle) {
      els.segnatavoloPreviewSubtitle.hidden = false;
      els.segnatavoloPreviewSubtitle.style.fontFamily = theme.previewBody;
      els.segnatavoloPreviewSubtitle.style.fontSize = `${subPx}px`;
      els.segnatavoloPreviewSubtitle.style.color = rgbToCss(pal.subtitle);
      els.segnatavoloPreviewSubtitle.textContent = headers.subtitle;
    } else {
      els.segnatavoloPreviewSubtitle.hidden = true;
    }

    if (els.segnatavoloPreviewRule) {
      els.segnatavoloPreviewRule.style.background = rgbToCss(pal.deco);
    }

    const names = getTableGuestNamesOrdered(table);
    els.segnatavoloPreviewNames.innerHTML = "";
    if (segnatavoloSettings.showGuests) {
      names.forEach((n) => {
        const li = document.createElement("li");
        li.textContent = n;
        li.style.fontFamily = theme.previewBody;
        li.style.fontSize = `${bodyPx}px`;
        li.style.color = rgbToCss(pal.body);
        els.segnatavoloPreviewNames.appendChild(li);
      });
    }
    if (els.segnatavoloPreviewEmpty) {
      els.segnatavoloPreviewEmpty.hidden = !segnatavoloSettings.showGuests || names.length > 0;
      els.segnatavoloPreviewEmpty.style.color = rgbToCss(pal.subtitle);
      els.segnatavoloPreviewEmpty.style.fontSize = `${emptyPx}px`;
      els.segnatavoloPreviewEmpty.style.fontFamily = theme.previewBody;
    }
    if (els.segnatavoloPreviewFoot) {
      els.segnatavoloPreviewFoot.textContent = (state.eventName || "").trim() || "Evento";
      els.segnatavoloPreviewFoot.style.color = rgbToCss(pal.foot);
      els.segnatavoloPreviewFoot.style.fontFamily = theme.previewBody;
      els.segnatavoloPreviewFoot.style.fontSize = `${footPx}px`;
    }
  }
  renderSegnatavoloControlSample(theme, pal, batchLayout);
  refreshSegnatavoloChipSelection();
  updateSegnatavoloCustomColorsVisibility();
  renderMenuBookletPreview();
}

function buildSegnatavoloDesignerUi() {
  if (els.segnatavoloOccasionChips) {
    els.segnatavoloOccasionChips.innerHTML = "";
    for (const o of SEGNATAVOLO_OCCASIONS) {
      const b = document.createElement("button");
      b.type = "button";
      b.className = "segnatavolo-chip";
      b.textContent = o.label;
      b.dataset.occasionId = o.id;
      b.title =
        "Applica caratteri, colori e illustrazioni tematiche ai margini (presepe, zucche, fuochi d’artificio…). Puoi rifinire tutto sotto.";
      b.addEventListener("click", () => {
        segnatavoloSettings.themeId = o.themeId;
        segnatavoloSettings.paletteId = o.paletteId;
        segnatavoloSettings.graphicIndex = o.graphicIndex;
        segnatavoloSettings.festiveMarginId = o.id === "none" ? null : o.id;
        saveSegnatavoloSettings();
        updateSegnatavoloPreview();
      });
      els.segnatavoloOccasionChips.appendChild(b);
    }
  }
  if (els.segnatavoloThemeChips) {
    els.segnatavoloThemeChips.innerHTML = "";
    for (const t of SEGNATAVOLO_THEMES) {
      const b = document.createElement("button");
      b.type = "button";
      b.className = "segnatavolo-chip";
      b.textContent = t.label;
      b.dataset.themeId = t.id;
      b.addEventListener("click", () => {
        segnatavoloSettings.themeId = t.id;
        saveSegnatavoloSettings();
        updateSegnatavoloPreview();
      });
      els.segnatavoloThemeChips.appendChild(b);
    }
  }

  if (els.segnatavoloPaletteChips) {
    els.segnatavoloPaletteChips.innerHTML = "";
    for (const p of SEGNATAVOLO_PALETTES) {
      const b = document.createElement("button");
      b.type = "button";
      b.className = "segnatavolo-chip";
      b.textContent = p.label;
      b.dataset.paletteId = p.id;
      b.addEventListener("click", () => {
        segnatavoloSettings.paletteId = p.id;
        saveSegnatavoloSettings();
        updateSegnatavoloPreview();
      });
      els.segnatavoloPaletteChips.appendChild(b);
    }
  }

  if (els.segnatavoloGraphicChips) {
    els.segnatavoloGraphicChips.innerHTML = "";
    for (const g of SEGNATAVOLO_GRAPHICS) {
      const b = document.createElement("button");
      b.type = "button";
      b.className = "segnatavolo-chip";
      b.textContent = g.label;
      b.dataset.graphicIndex = String(g.id);
      b.addEventListener("click", () => {
        segnatavoloSettings.graphicIndex = g.id;
        segnatavoloSettings.festiveMarginId = null;
        saveSegnatavoloSettings();
        updateSegnatavoloPreview();
      });
      els.segnatavoloGraphicChips.appendChild(b);
    }
  }

  ["segnatavoloColorTitle", "segnatavoloColorBody", "segnatavoloColorDeco"].forEach((id) => {
    const el = document.getElementById(id);
    if (el) {
      el.addEventListener("input", () => {
        if (els.segnatavoloColorTitle) {
          segnatavoloSettings.customTitleHex = els.segnatavoloColorTitle.value;
          segnatavoloSettings.customBodyHex = els.segnatavoloColorBody.value;
          segnatavoloSettings.customDecoHex = els.segnatavoloColorDeco.value;
        }
        saveSegnatavoloSettings();
        updateSegnatavoloPreview();
      });
    }
  });

  [["segnatavoloPaperFormat", "change"], ["segnatavoloOrientation", "change"], ["segnatavoloHeaderMode", "change"]]
    .forEach(([id, eventName]) => {
      const el = document.getElementById(id);
      if (!el) return;
      el.addEventListener(eventName, () => {
        saveSegnatavoloSettings();
        updateSegnatavoloPreview();
      });
    });

  if (els.segnatavoloShowGuests) {
    els.segnatavoloShowGuests.addEventListener("change", () => {
      saveSegnatavoloSettings();
      updateSegnatavoloPreview();
    });
  }
}

function runSegnatavoliPdfExport() {
  if (!state.tables.length) {
    alert("Non ci sono tavoli da esportare.");
    return;
  }
  const orientation = segnatavoloSettings.paperOrientation === "landscape" ? "landscape" : "portrait";
  const { key: formatKey, format } = getSegnatavoloPdfFormatAndKey();
  const pdf = window.jspdf && window.jspdf.jsPDF
    ? new window.jspdf.jsPDF({ orientation, unit: "mm", format })
    : null;
  if (!pdf) {
    alert("Libreria PDF non disponibile.");
    return;
  }
  const palette = getSegnatavoloPaletteResolved();
  const theme = getSegnatavoloTheme();
  const W = pdf.internal.pageSize.getWidth();
  const H = pdf.internal.pageSize.getHeight();
  const ordered = [...state.tables].sort((a, b) => Number(a.number) - Number(b.number));
  const layoutSettings = {
    headerMode: segnatavoloSettings.headerMode,
    showGuests: Boolean(segnatavoloSettings.showGuests),
  };
  const batchLayout = computeSegnatavoloBatchLayout(pdf, ordered, layoutSettings, theme, W, H);
  const opts = {
    themeId: segnatavoloSettings.themeId,
    paletteId: segnatavoloSettings.paletteId,
    graphicIndex: segnatavoloSettings.graphicIndex,
    festiveMarginId: segnatavoloSettings.festiveMarginId || null,
    headerMode: segnatavoloSettings.headerMode,
    showGuests: layoutSettings.showGuests,
    _palette: palette,
    _layout: batchLayout,
  };
  ordered.forEach((table, idx) => {
    if (idx > 0) pdf.addPage();
    drawSegnatavoloPage(pdf, table, opts);
  });
  pdf.save(`segnatavoli-${formatKey}-${orientation === "landscape" ? "orizzontale" : "verticale"}.pdf`);
}

function drawMenuBookletPanelFrame(pdf, x, y, w, h, palette) {
  const fid = segnatavoloSettings.festiveMarginId || null;
  const gi = Number(segnatavoloSettings.graphicIndex || 0);
  pdf.internal.write(`q 1 0 0 1 ${x} ${y} cm`);
  try {
    if (fid) {
      drawSegnatavoloFestiveMarginArt(pdf, w, h, palette, fid);
    } else {
      drawSegnatavoloGraphic(pdf, w, h, palette.deco, gi);
    }
  } finally {
    pdf.internal.write("Q");
  }
}

function drawMenuBookletCoverPage(pdf, x, y, w, h, theme, palette) {
  drawMenuBookletPanelFrame(pdf, x, y, w, h, palette);
  const title = (state.eventName || "Evento").trim() || "Evento";
  let cursorY = y + 18;
  if (menuBookletState.logoDataUrl) {
    try {
      const imgW = Math.min(w * 0.5, 52);
      const imgH = Math.min(h * 0.26, 42);
      const ix = x + (w - imgW) / 2;
      pdf.addImage(menuBookletState.logoDataUrl, "PNG", ix, cursorY, imgW, imgH);
      cursorY += imgH + 10;
    } catch (_) {
      // logo non leggibile: prosegue con titolo evento.
    }
  } else {
    cursorY += 22;
  }
  pdf.setFont(theme.titleFamily, theme.titleStyle);
  pdf.setFontSize(Math.max(14, theme.titleSize));
  pdf.setTextColor(...palette.title);
  const lines = pdf.splitTextToSize(title, w - 26);
  lines.forEach((line) => {
    pdf.text(line, x + w / 2, cursorY, { align: "center" });
    cursorY += Math.max(8, theme.titleSize * 0.42);
  });
}

function computeMenuBookletMenuLayout(pdf, theme, x, y, w, h) {
  const safe = SEGNATAVOLO_SAFE_MARGIN_FRAC;
  const mx = w * safe;
  const my = h * safe;
  const left = x + mx;
  const right = x + w - mx;
  const innerW = Math.max(36, right - left);
  const center = x + w / 2;
  const top = y + my;
  const maxY = y + h - my;
  const items = menuBookletState.items;

  function simHeight(bodySize, titleSz, subSz) {
    let cursorY = top + 3;
    pdf.setFont(theme.titleFamily, theme.titleStyle);
    pdf.setFontSize(titleSz);
    const titleLines = pdf.splitTextToSize("MENU", innerW);
    const tLh = Math.max(5.8, titleSz * 0.33);
    cursorY += titleLines.length * tLh + 2;
    cursorY += 6;
    const lineH = Math.max(3.6, bodySize * 0.35);
    if (!items.length) {
      cursorY += lineH * 1.8;
      return cursorY;
    }
    for (const item of items) {
      if (item.type === "separator_long" || item.type === "separator_short") {
        cursorY += 8;
      } else {
        const text = String(item.text || "").trim();
        if (!text) continue;
        pdf.setFont(theme.bodyFamily, theme.bodyStyle);
        pdf.setFontSize(bodySize);
        const wrapped = pdf.splitTextToSize(text, innerW);
        cursorY += wrapped.length * lineH + 3.2;
      }
    }
    return cursorY;
  }

  let bodySize = theme.bodySize + 1.2;
  let titleSz = theme.titleSize;
  let subSz = theme.subSize;
  const minB = 4.2;
  while (bodySize >= minB) {
    if (simHeight(bodySize, titleSz, subSz) <= maxY + 0.5) {
      return {
        left,
        right,
        innerW,
        center,
        top,
        maxY,
        bodySize,
        titleSz,
        subSz,
        lineH: Math.max(3.6, bodySize * 0.35),
      };
    }
    bodySize -= 0.3;
    titleSz = Math.max(9, titleSz - 0.12);
    subSz = Math.max(5.5, subSz - 0.12);
  }
  return {
    left,
    right,
    innerW,
    center,
    top,
    maxY,
    bodySize: minB,
    titleSz: 9,
    subSz: 5.5,
    lineH: Math.max(3.6, minB * 0.35),
  };
}

function drawMenuBookletMenuPage(pdf, x, y, w, h, theme, palette, layoutIn) {
  drawMenuBookletPanelFrame(pdf, x, y, w, h, palette);
  const L = layoutIn || computeMenuBookletMenuLayout(pdf, theme, x, y, w, h);
  const { left, right, innerW, center, top, maxY, bodySize, titleSz, subSz, lineH } = L;
  let cursorY = top + 3;

  pdf.setFont(theme.titleFamily, theme.titleStyle);
  pdf.setFontSize(titleSz);
  pdf.setTextColor(...palette.title);
  const titleLines = pdf.splitTextToSize("MENU", innerW);
  const tLh = Math.max(5.8, titleSz * 0.33);
  titleLines.forEach((line) => {
    pdf.text(line, center, cursorY, { align: "center" });
    cursorY += tLh;
  });
  cursorY += 2;
  pdf.setDrawColor(...palette.deco);
  pdf.setLineWidth(0.33);
  pdf.line(left + innerW * 0.06, cursorY, right - innerW * 0.06, cursorY);
  cursorY += 6;

  if (!menuBookletState.items.length) {
    pdf.setFont(theme.bodyFamily, "italic");
    pdf.setFontSize(Math.max(6, bodySize * 0.95));
    pdf.setTextColor(...palette.subtitle);
    pdf.text("Nessuna pietanza selezionata.", center, cursorY + 6, { align: "center" });
    return;
  }

  for (const item of menuBookletState.items) {
    if (cursorY >= maxY) break;
    if (item.type === "separator_long" || item.type === "separator_short") {
      cursorY += 2;
      pdf.setDrawColor(...palette.deco);
      pdf.setLineWidth(0.25);
      const inset = item.type === "separator_short" ? innerW * 0.3 : innerW * 0.12;
      pdf.line(left + inset, cursorY, right - inset, cursorY);
      cursorY += 4;
      pdf.setFont(theme.bodyFamily, theme.bodyStyle);
      pdf.setFontSize(bodySize);
      pdf.setTextColor(...palette.body);
      continue;
    }
    const text = String(item.text || "").trim();
    if (!text) continue;

    pdf.setFont(theme.bodyFamily, theme.bodyStyle);
    pdf.setFontSize(bodySize);
    pdf.setTextColor(...palette.body);
    const wrapped = pdf.splitTextToSize(text, innerW);
    for (const line of wrapped) {
      if (cursorY + lineH > maxY) break;
      pdf.text(line, center, cursorY, { align: "center" });
      cursorY += lineH;
    }
    cursorY += 3.2;
  }
}

function runMenuBookletPdfExport() {
  if (!window.jspdf || !window.jspdf.jsPDF) {
    alert("Libreria PDF non disponibile.");
    return;
  }
  const theme = getSegnatavoloTheme();
  const palette = getSegnatavoloPaletteResolved();
  const pdf = new window.jspdf.jsPDF({ orientation: "landscape", unit: "mm", format: "a4" });
  const W = pdf.internal.pageSize.getWidth();
  const H = pdf.internal.pageSize.getHeight();
  const half = W / 2;
  const menuLayout = computeMenuBookletMenuLayout(pdf, theme, half, 0, half, H);

  // Foglio lato esterno: sinistra pagina 4 (vuota), destra pagina 1 (copertina)
  drawMenuBookletCoverPage(pdf, half, 0, half, H, theme, palette);

  // Foglio lato interno: sinistra pagina 2 (vuota), destra pagina 3 (menu)
  pdf.addPage();
  drawMenuBookletMenuPage(pdf, half, 0, half, H, theme, palette, menuLayout);

  pdf.save("libricino-menu-a4-piegato.pdf");
}

function initSegnatavoloModal() {
  if (els.segnatavoloAreaPanel && !els.segnatavoloAreaPanel.dataset.studioChipClose) {
    els.segnatavoloAreaPanel.dataset.studioChipClose = "1";
    els.segnatavoloAreaPanel.addEventListener("click", (e) => {
      const chip = e.target.closest(".segnatavolo-chip");
      if (!chip || !chip.closest("details.studio-dd")) return;
      const dd = chip.closest("details.studio-dd");
      requestAnimationFrame(() => dd?.removeAttribute("open"));
    });
  }
  buildSegnatavoloDesignerUi();
  loadSegnatavoloSettings();
  loadMenuBookletState();
  populateMenuBookletDishSelect();
  renderMenuBookletList();
  if (els.segnatavoloExportConfirm) {
    els.segnatavoloExportConfirm.addEventListener("click", () => {
      saveSegnatavoloSettings();
      runSegnatavoliPdfExport();
    });
  }
  if (els.menuBookletLogoInput) {
    els.menuBookletLogoInput.addEventListener("change", async () => {
      if (!canEditStudioAreas()) return;
      const file = els.menuBookletLogoInput.files && els.menuBookletLogoInput.files[0];
      if (!file) return;
      try {
        menuBookletState.logoDataUrl = await readFileAsDataUrl(file);
        saveMenuBookletState();
        renderMenuBookletPreview();
      } catch (_) {
        alert("Impossibile leggere il logo selezionato.");
      }
      els.menuBookletLogoInput.value = "";
    });
  }
  if (els.menuBookletDishSelect) {
    els.menuBookletDishSelect.addEventListener("change", () => {
      if (!canEditStudioAreas()) return;
      const draft = getMenuBookletSelectionDraft();
      openMenuBookletEditor(draft.text, "add", draft.type, null);
    });
  }
  if (els.menuBookletEditorConfirmBtn) {
    els.menuBookletEditorConfirmBtn.addEventListener("click", () => {
      if (!canEditStudioAreas()) return;
      commitMenuBookletEditor();
    });
  }
  if (els.menuBookletEditorCancelBtn) {
    els.menuBookletEditorCancelBtn.addEventListener("click", () => {
      closeMenuBookletEditor();
    });
  }
  if (els.menuBookletList) {
    els.menuBookletList.addEventListener("click", (e) => {
      const up = e.target.closest("[data-menu-move-up]");
      if (up) {
        const idx = menuBookletState.items.findIndex((x) => x.id === up.dataset.menuMoveUp);
        if (idx > 0) {
          const tmp = menuBookletState.items[idx - 1];
          menuBookletState.items[idx - 1] = menuBookletState.items[idx];
          menuBookletState.items[idx] = tmp;
          saveMenuBookletState();
          renderMenuBookletList();
        }
        return;
      }
      const down = e.target.closest("[data-menu-move-down]");
      if (down) {
        const idx = menuBookletState.items.findIndex((x) => x.id === down.dataset.menuMoveDown);
        if (idx >= 0 && idx < menuBookletState.items.length - 1) {
          const tmp = menuBookletState.items[idx + 1];
          menuBookletState.items[idx + 1] = menuBookletState.items[idx];
          menuBookletState.items[idx] = tmp;
          saveMenuBookletState();
          renderMenuBookletList();
        }
        return;
      }
      const actionBtn = e.target.closest("[data-menu-action][data-menu-id]");
      if (!actionBtn) return;
      const action = actionBtn.dataset.menuAction;
      const id = actionBtn.dataset.menuId;
      if (!action || !id) return;
      const menuWrap = actionBtn.closest(".menu-booklet-row__menu");
      if (menuWrap) menuWrap.open = false;
      if (action === "delete") {
        menuBookletState.items = menuBookletState.items.filter((x) => x.id !== id);
        saveMenuBookletState();
        renderMenuBookletList();
        return;
      }
      const item = menuBookletState.items.find((x) => x.id === id);
      if (!item) return;
      if (action === "edit") {
        openMenuBookletEditor(item.text, "edit", item.type, id);
        return;
      }
    });
  }
  if (els.menuBookletExportBtn) {
    els.menuBookletExportBtn.addEventListener("click", () => {
      saveMenuBookletState();
      runMenuBookletPdfExport();
    });
  }
}


async function exportPlanPdf(planId) {
  closeAllDropdowns();
  if (!planId) return;
  if (!window.html2canvas || !window.jspdf || !window.jspdf.jsPDF) {
    alert("Librerie PDF non disponibili.");
    return;
  }
  const plan = state.floorPlans.find((p) => p.id === planId);
  if (!plan) {
    alert("Riquadro piantina non trovato.");
    return;
  }
  const previousView = getCurrentMainAreaView();
  try {
    if (previousView !== "map") {
      setMainAreaView("map");
      await waitForNextFrame();
    }
    const esc = typeof CSS !== "undefined" && CSS.escape ? CSS.escape(planId) : planId.replace(/"/g, '\\"');
    const card = document.querySelector(`.floor-plan-card[data-plan-id="${esc}"]`);
    const targetEl = card && card.querySelector(".floor-wrap");
    if (!targetEl) {
      alert("Impossibile esportare: riquadro non disponibile al momento.");
      return;
    }
    await waitForFloorImageReady(1200, targetEl);
    const rect = targetEl.getBoundingClientRect();
    if (rect.width < 10 || rect.height < 10) {
      alert("Impossibile esportare: area piantina troppo piccola.");
      return;
    }
    const canvas = await window.html2canvas(targetEl, {
      backgroundColor: "#ffffff",
      scale: 2,
      useCORS: true,
    });
    const imgData = canvas.toDataURL("image/png");
    const { jsPDF } = window.jspdf;
    const pdf = new jsPDF({ orientation: "landscape", unit: "mm", format: "a4" });
    const pageW = pdf.internal.pageSize.getWidth();
    const pageH = pdf.internal.pageSize.getHeight();
    const margin = 8;
    const planIndex = Math.max(1, state.floorPlans.findIndex((p) => p.id === planId) + 1);
    const multi = state.floorPlans.length > 1;
    const imgTop = multi ? 13 : 10;
    const usableW = pageW - margin * 2;
    const usableH = pageH - imgTop - margin;
    const ratio = canvas.width / canvas.height;
    let drawW = usableW;
    let drawH = drawW / ratio;
    if (drawH > usableH) {
      drawH = usableH;
      drawW = drawH * ratio;
    }
    pdf.setFontSize(11);
    pdf.text(state.eventName || "Piantina tavoli", margin, 6);
    if (multi) {
      pdf.setFontSize(9);
      pdf.text(`Riquadro ${planIndex}`, margin, 10);
    }
    pdf.addImage(imgData, "PNG", margin, imgTop, drawW, drawH);
    const safeName =
      String(state.eventName || "piantina")
        .replace(/[/\\?%*:|"<>]/g, "")
        .trim()
        .replace(/\s+/g, "-")
        .slice(0, 40) || "piantina";
    pdf.save(`${safeName}-${planIndex}.pdf`);
  } catch (_) {
    alert("Errore durante l'esportazione PDF della piantina.");
  } finally {
    if (getCurrentMainAreaView() !== previousView) {
      setMainAreaView(previousView);
    }
  }
}

function applyClientProjectData(clientData, asNewProject) {
  const payload = clientData && typeof clientData === "object" ? clientData : {};
  if (payload.eventName != null) state.eventName = String(payload.eventName);
  if (Array.isArray(payload.tables)) state.tables = payload.tables;
  state.tables.forEach(ensureTableGuestSlots);
  if (asNewProject) {
    state.floorPlanDataUrl = "";
    state.markerPositions = {};
    state.floorPlans = [];
    menuBookletState.logoDataUrl = "";
    menuBookletState.items = [];
    segnatavoloSettings = { ...DEFAULT_SEGNATAVOLO_SETTINGS };
    saveMenuBookletState();
    saveSegnatavoloSettings();
    renderMenuBookletList();
    updateSegnatavoloPreview();
  }
  ensureFloorPlansState();
  els.eventName.value = state.eventName;
  renderTables();
  renderFloorPlans();
  saveState();
}

function applyFullProjectData(fullData) {
  const payload = fullData && typeof fullData === "object" ? fullData : {};
  if (payload.eventName != null) state.eventName = String(payload.eventName);
  if (payload.floorPlanDataUrl != null) state.floorPlanDataUrl = String(payload.floorPlanDataUrl);
  if (Array.isArray(payload.tables)) state.tables = payload.tables;
  state.tables.forEach(ensureTableGuestSlots);
  if (payload.markerPositions && typeof payload.markerPositions === "object") {
    state.markerPositions = payload.markerPositions;
  }
  if (Array.isArray(payload.floorPlans)) {
    state.floorPlans = payload.floorPlans;
  }
  if (payload.segnatavoloSettings && typeof payload.segnatavoloSettings === "object") {
    segnatavoloSettings = { ...DEFAULT_SEGNATAVOLO_SETTINGS, ...payload.segnatavoloSettings };
    saveSegnatavoloSettings();
  }
  if (payload.menuBookletState && typeof payload.menuBookletState === "object") {
    menuBookletState.logoDataUrl = String(payload.menuBookletState.logoDataUrl || "");
    menuBookletState.items = normalizeMenuBookletItems(payload.menuBookletState.items);
    saveMenuBookletState();
    renderMenuBookletList();
  }
  ensureFloorPlansState();
  els.eventName.value = state.eventName;
  renderTables();
  renderFloorPlans();
  updateSegnatavoloPreview();
  saveState();
}

els.importInput.addEventListener("change", (e) => {
  const file = e.target.files && e.target.files[0];
  if (!file) return;
  const kind = pendingImportKind || "";
  if (kind && !fileLooksLikeExpectedKind(file.name, kind)) {
    alert(
      kind === "cliente"
        ? `Formato file non coerente. Seleziona un file ${CLIENT_FILE_EXT} (oppure un .json valido).`
        : `Formato file non coerente. Seleziona un file ${FULL_FILE_EXT} (oppure un .json valido).`
    );
    pendingImportKind = "";
    e.target.value = "";
    return;
  }
  const reader = new FileReader();
  reader.onload = () => {
    try {
      const parsed = JSON.parse(reader.result);
      const fileType = parsed && parsed.fileType;
      const data = parsed && parsed.data ? parsed.data : parsed;
      if (kind === "cliente") {
        if (fileType && fileType !== "cliente") {
          alert("Questo non è un file Cliente.");
          return;
        }
        if (appRole === "struttura") {
          if (!window.confirm("Import Cliente come nuovo progetto? I dati progetto attuali verranno sostituiti.")) return;
          applyClientProjectData(data, true);
        } else {
          applyClientProjectData(data, false);
        }
      } else if (kind === "completo") {
        if (fileType && fileType !== "struttura") {
          alert("Questo non è un file Struttura completo.");
          return;
        }
        applyFullProjectData(data);
      } else {
        alert("Seleziona prima il tipo di importazione dal menu.");
      }
    } catch (_) {
      alert("File JSON non valido.");
    }
  };
  reader.readAsText(file);
  pendingImportKind = "";
  e.target.value = "";
});

window.addEventListener("resize", () => {
  document.querySelectorAll(".floor-markers").forEach((root) => layoutMarkerNotes(root));
  updateCurrentTableContextUi();
  if (els.segnatavoloAreaPanel && !els.segnatavoloAreaPanel.hidden) {
    scheduleFitSegnatavoloStudioSnapPreviews();
    scheduleFitSegnatavoloControlSample();
  }
  if (els.menuBookletAreaPanel && !els.menuBookletAreaPanel.hidden) {
    scheduleFitStudioMenuBookletScale();
  }
});

if (els.tablesList) {
  els.tablesList.addEventListener("scroll", updateCurrentTableContextUi, { passive: true });
}

document.addEventListener("keydown", (e) => {
  if (e.key === "Escape") {
    if (document.querySelector(".guest-row__sheet:not([hidden])")) {
      closeAllGuestOptionSheets();
      return;
    }
    document.querySelectorAll("#segnatavoloAreaPanel details[open], #menuBookletAreaPanel details[open]").forEach((d) => {
      d.removeAttribute("open");
    });
    closeAllDropdowns();
  }
});

initSegnatavoloModal();
loadState();
setMainAreaView("tables");
els.eventName.value = state.eventName;
applyFloorPlanFromState();
renderTables();
renderFloorMarkers();
applyRoleUi();
