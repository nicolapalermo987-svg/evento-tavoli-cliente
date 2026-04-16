/**
 * App organizzazione tavoli — stato in memoria + localStorage
 * Max 9 ospiti per tavolo; marcatori sulla piantina sincronizzati con i tavoli.
 */

const MAX_GUESTS = 9;
const STORAGE_KEY = "evento-tavoli-v1";

const state = {
  eventName: "",
  floorPlanDataUrl: "",
  tables: [],
  /** { [tableId]: { x: 0-100, y: 0-100 } } percentuali rispetto al contenitore */
  markerPositions: {},
};

const acceptanceUiState = {
  query: "",
  selectedTableId: "",
};

const gridPlanState = {
  cols: 7,
  rows: 8,
  roomWidth: 7,
  roomHeight: 8,
  roomOriginX: 0,
  roomOriginY: 0,
  tileSize: 0.6,
  tileSizeX: 0.6,
  tileSizeY: 0.6,
  cells: new Map(),
  segments: [],
  pendingNodeAction: "",
  firstSelectedNode: null,
  activeCell: null,
  a4Orientation: "portrait",
  zoom: 1,
  pendingAddTilesAnchor: null,
};

const sketchPlanState = {
  fabricCanvas: null,
  selectedZone: null,
  baseSnapshot: [],
  globalScale: 1,
  nextZoneId: 1,
};

const els = {
  eventName: document.getElementById("eventName"),
  stickyToolbar: document.getElementById("stickyToolbar"),
  addTableBtn: document.getElementById("addTableBtn"),
  showTablesAreaBtn: document.getElementById("showTablesAreaBtn"),
  showMapAreaBtn: document.getElementById("showMapAreaBtn"),
  showAcceptanceAreaBtn: document.getElementById("showAcceptanceAreaBtn"),
  tablesAreaPanel: document.getElementById("tablesAreaPanel"),
  mapAreaPanel: document.getElementById("mapAreaPanel"),
  acceptanceAreaPanel: document.getElementById("acceptanceAreaPanel"),
  acceptanceSearchInput: document.getElementById("acceptanceSearchInput"),
  acceptanceResults: document.getElementById("acceptanceResults"),
  acceptanceTablePreview: document.getElementById("acceptanceTablePreview"),
  sortTablesBtn: document.getElementById("sortTablesBtn"),
  renumberTablesBtn: document.getElementById("renumberTablesBtn"),
  tablesList: document.getElementById("tablesList"),
  tableCountLabel: document.getElementById("tableCountLabel"),
  currentTableBadge: document.getElementById("currentTableBadge"),
  globalOccupancyLine: document.getElementById("globalOccupancyLine"),
  globalSummaryLine: document.getElementById("globalSummaryLine"),
  tableActionsDropdown: document.getElementById("tableActionsDropdown"),
  tableActionsMenuBtn: document.getElementById("tableActionsMenuBtn"),
  tableActionsMenu: document.getElementById("tableActionsMenu"),
  exportDropdown: document.getElementById("exportDropdown"),
  exportMenuBtn: document.getElementById("exportMenuBtn"),
  exportMenu: document.getElementById("exportMenu"),
  importDropdown: document.getElementById("importDropdown"),
  importMenuBtn: document.getElementById("importMenuBtn"),
  importMenu: document.getElementById("importMenu"),
  dropdownBackdrop: document.getElementById("dropdownBackdrop"),
  importJsonBtn: document.getElementById("importJsonBtn"),
  floorPlanInput: document.getElementById("floorPlanInput"),
  createFloorPlanBtn: document.getElementById("createFloorPlanBtn"),
  clearPlanBtn: document.getElementById("clearPlanBtn"),
  floorWrap: document.getElementById("floorWrap"),
  floorHint: document.getElementById("floorHint"),
  floorStage: document.getElementById("floorStage"),
  floorImg: document.getElementById("floorImg"),
  floorMarkers: document.getElementById("floorMarkers"),
  tileSizeInput: document.getElementById("tileSizeInput"),
  tileSizeXInput: document.getElementById("tileSizeXInput"),
  tileSizeYInput: document.getElementById("tileSizeYInput"),
  roomTilesXInput: document.getElementById("roomTilesXInput"),
  roomTilesYInput: document.getElementById("roomTilesYInput"),
  planGridSummary: document.getElementById("planGridSummary"),
  planGrid: document.getElementById("planGrid"),
  planOverlay: document.getElementById("planOverlay"),
  planCellMenu: document.getElementById("planCellMenu"),
  planNodePicker: document.getElementById("planNodePicker"),
  planNodePickerTitle: document.getElementById("planNodePickerTitle"),
  planNodePickerGrid: document.getElementById("planNodePickerGrid"),
  planNodePickerCloseBtn: document.getElementById("planNodePickerCloseBtn"),
  planAddTilesDialog: document.getElementById("planAddTilesDialog"),
  planAddTileSizeXInput: document.getElementById("planAddTileSizeXInput"),
  planAddTileSizeYInput: document.getElementById("planAddTileSizeYInput"),
  planAddTilesWidthInput: document.getElementById("planAddTilesWidthInput"),
  planAddTilesHeightInput: document.getElementById("planAddTilesHeightInput"),
  planAddTilesDirectionSelect: document.getElementById("planAddTilesDirectionSelect"),
  planAddTilesCancelBtn: document.getElementById("planAddTilesCancelBtn"),
  planAddTilesApplyBtn: document.getElementById("planAddTilesApplyBtn"),
  planEditorModal: document.getElementById("planEditorModal"),
  planEditorBackdrop: document.getElementById("planEditorBackdrop"),
  planEditorCloseBtn: document.getElementById("planEditorCloseBtn"),
  planA4Stage: document.getElementById("planA4Stage"),
  planA4Sheet: document.getElementById("planA4Sheet"),
  planEditor: document.getElementById("planEditor"),
  planSetupPanel: document.getElementById("planSetupPanel"),
  planSetupStepTileSize: document.getElementById("planSetupStepTileSize"),
  planSetupStepTilesCount: document.getElementById("planSetupStepTilesCount"),
  planSetupGenerateBtn: document.getElementById("planSetupGenerateBtn"),
  planSetupDrawGridLandscapeBtn: document.getElementById("planSetupDrawGridLandscapeBtn"),
  planSetupDrawGridPortraitBtn: document.getElementById("planSetupDrawGridPortraitBtn"),
  planZoomOutBtn: document.getElementById("planZoomOutBtn"),
  planZoomResetBtn: document.getElementById("planZoomResetBtn"),
  planZoomInBtn: document.getElementById("planZoomInBtn"),
  planEditorBody: document.getElementById("planEditorBody"),
  planEditorStatus: document.getElementById("planEditorStatus"),
  exportPlanPdfBtn: document.getElementById("exportPlanPdfBtn"),
  exportGuestsBtn: document.getElementById("exportGuestsBtn"),
  exportKitchenBtn: document.getElementById("exportKitchenBtn"),
  exportSegnatavoliBtn: document.getElementById("exportSegnatavoliBtn"),
  exportBtn: document.getElementById("exportBtn"),
  importInput: document.getElementById("importInput"),
  tableCardTpl: document.getElementById("tableCardTpl"),
  guestRowTpl: document.getElementById("guestRowTpl"),
  segnatavoloModal: document.getElementById("segnatavoloModal"),
  segnatavoloModalBackdrop: document.getElementById("segnatavoloModalBackdrop"),
  segnatavoloModalClose: document.getElementById("segnatavoloModalClose"),
  segnatavoloModalCancel: document.getElementById("segnatavoloModalCancel"),
  segnatavoloExportConfirm: document.getElementById("segnatavoloExportConfirm"),
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

function loadState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return;
    const data = JSON.parse(raw);
    if (data.eventName != null) state.eventName = String(data.eventName);
    if (data.floorPlanDataUrl != null) state.floorPlanDataUrl = String(data.floorPlanDataUrl);
    if (Array.isArray(data.tables)) state.tables = data.tables;
    state.tables.forEach(ensureTableGuestSlots);
    if (data.markerPositions && typeof data.markerPositions === "object") {
      state.markerPositions = data.markerPositions;
    }
  } catch (_) {
    /* ignore */
  }
}

function saveState() {
  try {
    localStorage.setItem(
      STORAGE_KEY,
      JSON.stringify({
        eventName: state.eventName,
        floorPlanDataUrl: state.floorPlanDataUrl,
        tables: state.tables,
        markerPositions: state.markerPositions,
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
  const id = uid();
  state.tables.push({
    id,
    number: nextTableNumber(),
    customName: "",
    tableNote: "",
    guests: Array.from({ length: MAX_GUESTS }, createEmptyGuest),
  });
  if (!state.markerPositions[id]) {
    state.markerPositions[id] = { x: 50, y: 50 };
  }
  saveState();
  renderTables();
  renderFloorMarkers();
}

function removeTable(tableId) {
  state.tables = state.tables.filter((t) => t.id !== tableId);
  delete state.markerPositions[tableId];
  saveState();
  renderTables();
  renderFloorMarkers();
}

function sortTablesAscending() {
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
      delete state.markerPositions[tableId];
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
  const table = state.tables.find((t) => t.id === tableId);
  if (!table) return;
  ensureTableGuestSlots(table);
  saveState();
  renderTables();
  renderFloorMarkers();
}

function removeGuest(tableId, index) {
  const table = state.tables.find((t) => t.id === tableId);
  if (!table || !table.guests[index]) return;
  table.guests[index] = createEmptyGuest();
  saveState();
  renderTables();
  renderFloorMarkers();
}

function updateGuest(tableId, index, field, value) {
  const table = state.tables.find((t) => t.id === tableId);
  if (!table || !table.guests[index]) return;
  table.guests[index][field] = value;
  saveState();
  refreshTableLiveInfoById(tableId);
  renderFloorMarkers();
  renderAcceptanceArea();
}

function updateTableMeta(tableId, field, value) {
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

/** Badge in alto + barra fissa in basso: sempre allineati al tavolo visibile nell’area scroll. */
function updateCurrentTableContextUi() {
  if (!els.currentTableBadge) return;
  if (!state.tables.length) {
    els.currentTableBadge.textContent = "Nessun tavolo";
    if (els.globalOccupancyLine) {
      els.globalOccupancyLine.textContent = "Posti occupati: —";
      els.globalOccupancyLine.classList.remove("is-full");
    }
    if (els.globalSummaryLine) els.globalSummaryLine.textContent = "Resoconto piantina: —";
    return;
  }
  const table = getCurrentVisibleTable();
  if (!table) {
    els.currentTableBadge.textContent = "Tavolo —";
    if (els.globalOccupancyLine) {
      els.globalOccupancyLine.textContent = "Posti occupati: —";
      els.globalOccupancyLine.classList.remove("is-full");
    }
    if (els.globalSummaryLine) els.globalSummaryLine.textContent = "Resoconto piantina: —";
    return;
  }
  const name = (table.customName || "").trim();
  els.currentTableBadge.textContent = name
    ? `Tavolo ${table.number} — ${name}`
    : `Tavolo ${table.number}`;
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
  if (els.exportMenu) {
    els.exportMenu.hidden = true;
  }
  if (els.importMenu) {
    els.importMenu.hidden = true;
  }
  if (els.tableActionsMenuBtn) {
    els.tableActionsMenuBtn.setAttribute("aria-expanded", "false");
  }
  if (els.exportMenuBtn) {
    els.exportMenuBtn.setAttribute("aria-expanded", "false");
  }
  if (els.importMenuBtn) {
    els.importMenuBtn.setAttribute("aria-expanded", "false");
  }
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

function ensureExportPlanPdfMenuItem() {
  if (!els.exportMenu) return;
  let btn = els.exportMenu.querySelector("#exportPlanPdfBtn");
  if (!btn) {
    btn = document.createElement("button");
    btn.type = "button";
    btn.className = "dropdown__item";
    btn.id = "exportPlanPdfBtn";
    btn.setAttribute("role", "menuitem");
    btn.textContent = "Esporta piantina PDF";
    els.exportMenu.prepend(btn);
  }
}

function setMainAreaView(view) {
  const targetView = view === "map" || view === "acceptance" ? view : "tables";
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

function planCellKey(x, y) {
  return `${x},${y}`;
}

function roomBounds() {
  return {
    x0: gridPlanState.roomOriginX,
    y0: gridPlanState.roomOriginY,
    x1: gridPlanState.roomOriginX + gridPlanState.roomWidth - 1,
    y1: gridPlanState.roomOriginY + gridPlanState.roomHeight - 1,
  };
}

function isInsideRoom(x, y) {
  return gridPlanState.cells.has(planCellKey(x, y));
}

function getPlanCellType(x, y) {
  return gridPlanState.cells.get(planCellKey(x, y)) || null;
}

function createPlanTileData(tileSizeX = gridPlanState.tileSizeX, tileSizeY = gridPlanState.tileSizeY) {
  return {
    type: "room",
    tileSizeX,
    tileSizeY,
  };
}

function rebuildPlanBounds() {
  const coords = Array.from(gridPlanState.cells.keys()).map((key) => key.split(",").map(Number));
  if (!coords.length) {
    gridPlanState.cols = 0;
    gridPlanState.rows = 0;
    gridPlanState.roomWidth = 0;
    gridPlanState.roomHeight = 0;
    return;
  }
  const xs = coords.map(([x]) => x);
  const ys = coords.map(([, y]) => y);
  const minX = Math.min(...xs);
  const maxX = Math.max(...xs);
  const minY = Math.min(...ys);
  const maxY = Math.max(...ys);

  if (minX < 0 || minY < 0) {
    const shiftX = minX < 0 ? -minX : 0;
    const shiftY = minY < 0 ? -minY : 0;
    const movedCells = new Map();
    for (const [key, value] of gridPlanState.cells.entries()) {
      const [x, y] = key.split(",").map(Number);
      movedCells.set(planCellKey(x + shiftX, y + shiftY), value);
    }
    gridPlanState.cells = movedCells;
    gridPlanState.segments = gridPlanState.segments.map((segment) => ({
      ...segment,
      a: { ...segment.a, x: segment.a.x + shiftX, y: segment.a.y + shiftY },
      b: { ...segment.b, x: segment.b.x + shiftX, y: segment.b.y + shiftY },
    }));
    if (gridPlanState.firstSelectedNode) {
      gridPlanState.firstSelectedNode = {
        ...gridPlanState.firstSelectedNode,
        x: gridPlanState.firstSelectedNode.x + shiftX,
        y: gridPlanState.firstSelectedNode.y + shiftY,
      };
    }
    if (gridPlanState.activeCell) {
      gridPlanState.activeCell = {
        x: gridPlanState.activeCell.x + shiftX,
        y: gridPlanState.activeCell.y + shiftY,
      };
    }
    return rebuildPlanBounds();
  }

  gridPlanState.roomOriginX = 0;
  gridPlanState.roomOriginY = 0;
  gridPlanState.cols = maxX + 1;
  gridPlanState.rows = maxY + 1;
  gridPlanState.roomWidth = gridPlanState.cols;
  gridPlanState.roomHeight = gridPlanState.rows;
}

function computePlanSummary() {
  let tiles = 0;
  let totalArea = 0;
  for (const tile of gridPlanState.cells.values()) {
    tiles += 1;
    totalArea += Number(tile.tileSizeX || gridPlanState.tileSizeX) * Number(tile.tileSizeY || gridPlanState.tileSizeY);
  }
  return { tiles, totalArea };
}

function updatePlanSummaryUi() {
  if (!els.planGridSummary) return;
  const s = computePlanSummary();
  els.planGridSummary.textContent = `Griglia: ${gridPlanState.cols} x ${gridPlanState.rows} • Mattonelle: ${s.tiles} • Area totale: ${s.totalArea.toFixed(2)} m²`;
}

function updatePlanToolButtonsUi() {
  return;
}

function setPlanEditorScreen(screen) {
  if (els.planSetupPanel) els.planSetupPanel.hidden = screen === "editor";
  if (els.planEditorBody) els.planEditorBody.hidden = screen !== "editor";
  if (screen === "editor") {
    requestAnimationFrame(() => {
      fitPlanSheetIntoStage();
      fitPlanGridIntoA4();
      renderPlanGrid();
    });
  }
}

function setPlanSetupStep(step) {
  if (els.planSetupStepTileSize) els.planSetupStepTileSize.hidden = step !== "tile-size";
  if (els.planSetupStepTilesCount) els.planSetupStepTilesCount.hidden = step !== "tiles-count";
}

function renderPlanGrid() {
  if (!els.planGrid) return;
  els.planGrid.innerHTML = "";
  els.planGrid.style.gridTemplateColumns = `repeat(${Math.max(1, gridPlanState.cols)}, var(--plan-cell-width, 26px))`;
  const frag = document.createDocumentFragment();
  for (let y = 0; y < gridPlanState.rows; y += 1) {
    for (let x = 0; x < gridPlanState.cols; x += 1) {
      const cell = getPlanCellType(x, y);
      const tile = document.createElement("button");
      tile.type = "button";
      tile.className = "plan-cell";
      tile.dataset.x = String(x);
      tile.dataset.y = String(y);
      tile.title = `Mattonella ${x + 1}, ${y + 1}`;
      if (!cell) {
        tile.classList.add("plan-cell--outside");
      }
      if (gridPlanState.firstSelectedNode && gridPlanState.firstSelectedNode.x === x && gridPlanState.firstSelectedNode.y === y) {
        tile.classList.add("plan-cell--node-source");
      }
      frag.appendChild(tile);
    }
  }
  els.planGrid.appendChild(frag);
  fitPlanGridIntoA4();
  renderPlanOverlay();
  updatePlanSummaryUi();
}

function renderPlanOverlay() {
  if (!els.planOverlay) return;
  const styles = getComputedStyle(els.planGrid);
  const cellW = Number.parseFloat(styles.getPropertyValue("--plan-cell-width")) || 26;
  const cellH = Number.parseFloat(styles.getPropertyValue("--plan-cell-height")) || 26;
  const gap = 0;
  const width = gridPlanState.cols * cellW + Math.max(0, gridPlanState.cols - 1) * gap;
  const height = gridPlanState.rows * cellH + Math.max(0, gridPlanState.rows - 1) * gap;
  els.planOverlay.setAttribute("viewBox", `0 0 ${width} ${height}`);
  els.planOverlay.innerHTML = "";

  gridPlanState.segments.forEach((seg) => {
    const line = document.createElementNS("http://www.w3.org/2000/svg", "line");
    line.setAttribute("x1", String(nodeToPixel(seg.a.x, seg.a.nx, cellW, gap)));
    line.setAttribute("y1", String(nodeToPixel(seg.a.y, seg.a.ny, cellH, gap)));
    line.setAttribute("x2", String(nodeToPixel(seg.b.x, seg.b.nx, cellW, gap)));
    line.setAttribute("y2", String(nodeToPixel(seg.b.y, seg.b.ny, cellH, gap)));
    line.setAttribute("class", "plan-segment");
    line.dataset.segmentId = seg.id;
    els.planOverlay.appendChild(line);
  });

  if (gridPlanState.firstSelectedNode) {
    const marker = document.createElementNS("http://www.w3.org/2000/svg", "circle");
    marker.setAttribute("cx", String(nodeToPixel(gridPlanState.firstSelectedNode.x, gridPlanState.firstSelectedNode.nx, cellW, gap)));
    marker.setAttribute("cy", String(nodeToPixel(gridPlanState.firstSelectedNode.y, gridPlanState.firstSelectedNode.ny, cellH, gap)));
    marker.setAttribute("r", "6");
    marker.setAttribute("class", "plan-node-highlight");
    els.planOverlay.appendChild(marker);
  }
}

function nodeToPixel(cellIndex, nodeIndex, cellSize, gap) {
  return cellIndex * (cellSize + gap) + (nodeIndex / 10) * cellSize;
}

function hidePlanMenus() {
  if (els.planCellMenu) els.planCellMenu.hidden = true;
  if (els.planNodePicker) els.planNodePicker.hidden = true;
  if (els.planAddTilesDialog) els.planAddTilesDialog.hidden = true;
}

function openPlanCellMenu(x, y, clientX, clientY) {
  gridPlanState.activeCell = { x, y };
  if (!els.planCellMenu) return;
  hidePlanMenus();
  els.planCellMenu.hidden = false;
  els.planCellMenu.style.left = `${clientX}px`;
  els.planCellMenu.style.top = `${clientY}px`;
}

function openNodePicker(action, clientX, clientY) {
  if (!els.planNodePicker || !els.planNodePickerGrid || !gridPlanState.activeCell) return;
  hidePlanMenus();
  gridPlanState.pendingNodeAction = action;
  const activeTile = getPlanCellType(gridPlanState.activeCell.x, gridPlanState.activeCell.y);
  const ratioX = Math.max(0.05, Number(activeTile?.tileSizeX || gridPlanState.tileSizeX || 1));
  const ratioY = Math.max(0.05, Number(activeTile?.tileSizeY || gridPlanState.tileSizeY || 1));
  const base = 18;
  const scale = Math.min(2.4, Math.max(0.45, 1 / Math.max(ratioX, ratioY)));
  const nodeW = Math.max(8, Math.round(base * ratioX * scale));
  const nodeH = Math.max(8, Math.round(base * ratioY * scale));
  els.planNodePickerGrid.style.setProperty("--plan-node-cell-w", `${nodeW}px`);
  els.planNodePickerGrid.style.setProperty("--plan-node-cell-h", `${nodeH}px`);
  els.planNodePickerTitle.textContent =
    action === "add-tiles"
      ? "Scegli il nodo di ancoraggio"
      : action === "label"
        ? "Scegli il nodo dove ancorare il testo"
        : "Scegli un nodo 11x11";
  els.planNodePickerGrid.innerHTML = "";
  for (let ny = 0; ny <= 10; ny += 1) {
    for (let nx = 0; nx <= 10; nx += 1) {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "plan-node-picker__node";
      btn.dataset.nx = String(nx);
      btn.dataset.ny = String(ny);
      els.planNodePickerGrid.appendChild(btn);
    }
  }
  els.planNodePicker.hidden = false;
  els.planNodePicker.style.left = `${Math.max(8, clientX - 140)}px`;
  els.planNodePicker.style.top = `${Math.max(8, clientY - 40)}px`;
}

function commitNodeSelection(nx, ny) {
  const cell = gridPlanState.activeCell;
  if (!cell) return;
  const tile = getPlanCellType(cell.x, cell.y);
  if (!tile) {
    hidePlanMenus();
    setPlanEditorStatus("Seleziona una mattonella attiva.");
    return;
  }
  if (gridPlanState.pendingNodeAction === "add-tiles") {
    openAddTilesDialog(cell, nx, ny);
    return;
  }
  const node = { x: cell.x, y: cell.y, nx, ny };
  if (!gridPlanState.firstSelectedNode) {
    gridPlanState.firstSelectedNode = node;
    hidePlanMenus();
    renderPlanOverlay();
    setPlanEditorStatus("Primo nodo selezionato. Scegli ora il secondo nodo.");
  } else {
    if (
      gridPlanState.firstSelectedNode.x === node.x &&
      gridPlanState.firstSelectedNode.y === node.y &&
      gridPlanState.firstSelectedNode.nx === node.nx &&
      gridPlanState.firstSelectedNode.ny === node.ny
    ) {
      gridPlanState.firstSelectedNode = null;
      hidePlanMenus();
      renderPlanOverlay();
      setPlanEditorStatus("Secondo nodo identico al primo: nessun segmento disegnato.");
      return;
    }
    gridPlanState.segments.push({
      id: `seg_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`,
      a: gridPlanState.firstSelectedNode,
      b: node,
    });
    gridPlanState.firstSelectedNode = null;
    hidePlanMenus();
    renderPlanOverlay();
    setPlanEditorStatus("Segmento creato.");
  }
}

function openAddTilesDialog(cell, nx, ny) {
  if (!els.planAddTilesDialog) return;
  gridPlanState.pendingAddTilesAnchor = { cell, nx, ny };
  if (els.planAddTileSizeXInput) els.planAddTileSizeXInput.value = String(gridPlanState.tileSizeX);
  if (els.planAddTileSizeYInput) els.planAddTileSizeYInput.value = String(gridPlanState.tileSizeY);
  if (els.planAddTilesWidthInput) els.planAddTilesWidthInput.value = "3";
  if (els.planAddTilesHeightInput) els.planAddTilesHeightInput.value = "5";
  if (els.planAddTilesDirectionSelect) els.planAddTilesDirectionSelect.value = "destra";
  els.planAddTilesDialog.hidden = false;
  els.planAddTilesDialog.style.left = `${Math.max(8, window.innerWidth / 2 - 200)}px`;
  els.planAddTilesDialog.style.top = `${Math.max(8, window.innerHeight / 2 - 140)}px`;
}

function applyAddTilesFromDialog() {
  const anchor = gridPlanState.pendingAddTilesAnchor;
  if (!anchor) return;
  const tileSizeX = Math.max(0.1, Number(els.planAddTileSizeXInput?.value || gridPlanState.tileSizeX));
  const tileSizeY = Math.max(0.1, Number(els.planAddTileSizeYInput?.value || gridPlanState.tileSizeY));
  const width = Math.max(1, Math.round(Number(els.planAddTilesWidthInput?.value || 1)));
  const height = Math.max(1, Math.round(Number(els.planAddTilesHeightInput?.value || 1)));
  const direction = String(els.planAddTilesDirectionSelect?.value || "destra");
  const directionMap = {
    alto: { xMode: "align", yMode: "before" },
    basso: { xMode: "align", yMode: "after" },
    sinistra: { xMode: "before", yMode: "align" },
    destra: { xMode: "after", yMode: "align" },
    "alto/sinistra": { xMode: "before", yMode: "before" },
    "alto/destra": { xMode: "after", yMode: "before" },
    "basso/sinistra": { xMode: "before", yMode: "after" },
    "basso/destra": { xMode: "after", yMode: "after" },
  };
  const map = directionMap[direction];
  if (!map) {
    setPlanEditorStatus("Direzione non valida.", true);
    return;
  }

  const { cell, nx, ny } = anchor;
  const startX =
    map.xMode === "before" ? cell.x - width : map.xMode === "after" ? cell.x + 1 : cell.x;
  const startY =
    map.yMode === "before" ? cell.y - height : map.yMode === "after" ? cell.y + 1 : cell.y;

  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      const px = startX + x;
      const py = startY + y;
      gridPlanState.cells.set(planCellKey(px, py), createPlanTileData(tileSizeX, tileSizeY));
    }
  }

  gridPlanState.tileSize = tileSizeX;
  gridPlanState.tileSizeX = tileSizeX;
  gridPlanState.tileSizeY = tileSizeY;
  rebuildPlanBounds();
  hidePlanMenus();
  renderPlanGrid();
  setPlanEditorStatus(
    `Aggiunte ${width}x${height} mattonelle (${tileSizeX.toFixed(2)}x${tileSizeY.toFixed(2)} m) verso ${direction}. Nodo: ${nx}/10, ${ny}/10.`
  );
  gridPlanState.pendingAddTilesAnchor = null;
}

function getGridProjectData() {
  return {
    tileSize: gridPlanState.tileSize,
    tileSizeX: gridPlanState.tileSizeX,
    tileSizeY: gridPlanState.tileSizeY,
    cols: gridPlanState.cols,
    rows: gridPlanState.rows,
    roomWidth: gridPlanState.roomWidth,
    roomHeight: gridPlanState.roomHeight,
    roomOriginX: gridPlanState.roomOriginX,
    roomOriginY: gridPlanState.roomOriginY,
    cells: Array.from(gridPlanState.cells.entries()),
    segments: gridPlanState.segments,
  };
}

function loadGridProjectData(data) {
  gridPlanState.tileSize = Number(data.tileSize || 0.6);
  gridPlanState.tileSizeX = Number(data.tileSizeX || 0.6);
  gridPlanState.tileSizeY = Number(data.tileSizeY || 0.6);
  gridPlanState.cols = Number(data.cols || 7);
  gridPlanState.rows = Number(data.rows || 8);
  gridPlanState.roomWidth = Number(data.roomWidth || 7);
  gridPlanState.roomHeight = Number(data.roomHeight || 8);
  gridPlanState.roomOriginX = Number(data.roomOriginX || 0);
  gridPlanState.roomOriginY = Number(data.roomOriginY || 0);
  gridPlanState.cells = new Map(Array.isArray(data.cells) ? data.cells : []);
  gridPlanState.segments = Array.isArray(data.segments) ? data.segments : [];
  gridPlanState.firstSelectedNode = null;
  if (els.tileSizeInput) els.tileSizeInput.value = String(gridPlanState.tileSize);
  if (els.tileSizeXInput) els.tileSizeXInput.value = String(gridPlanState.tileSizeX);
  if (els.tileSizeYInput) els.tileSizeYInput.value = String(gridPlanState.tileSizeY);
  if (els.roomTilesXInput) els.roomTilesXInput.value = String(gridPlanState.roomWidth);
  if (els.roomTilesYInput) els.roomTilesYInput.value = String(gridPlanState.roomHeight);
  renderPlanGrid();
}

function fitPlanGridIntoA4() {
  if (!els.planGrid || !els.planA4Sheet) return;
  fitPlanSheetIntoStage();
  const sheetRect = els.planA4Sheet.getBoundingClientRect();
  const availableW = Math.max(40, sheetRect.width - 36);
  const availableH = Math.max(40, sheetRect.height - 36);
  const gap = 0;
  const ratioX = Math.max(0.05, Number(gridPlanState.tileSizeX) || 1);
  const ratioY = Math.max(0.05, Number(gridPlanState.tileSizeY) || 1);
  const rawScale = Math.min(
    (availableW - gap * Math.max(0, gridPlanState.cols - 1)) / Math.max(1, gridPlanState.cols * ratioX),
    (availableH - gap * Math.max(0, gridPlanState.rows - 1)) / Math.max(1, gridPlanState.rows * ratioY)
  );
  const scale = Math.max(1, rawScale);
  const cellW = Math.max(1, Math.floor(scale * ratioX));
  const cellH = Math.max(1, Math.floor(scale * ratioY));
  els.planGrid.style.setProperty("--plan-cell-width", `${cellW}px`);
  els.planGrid.style.setProperty("--plan-cell-height", `${cellH}px`);
}

function fitPlanSheetIntoStage() {
  if (!els.planA4Stage || !els.planA4Sheet) return;
  const stageRect = els.planA4Stage.getBoundingClientRect();
  const availableW = Math.max(200, stageRect.width - 24);
  const availableH = Math.max(260, stageRect.height - 24);
  const a4Ratio = gridPlanState.a4Orientation === "landscape" ? 297 / 210 : 210 / 297;
  let sheetW = availableW;
  let sheetH = sheetW / a4Ratio;
  if (sheetH > availableH) {
    sheetH = availableH;
    sheetW = sheetH * a4Ratio;
  }
  els.planA4Sheet.style.width = `${Math.floor(sheetW)}px`;
  els.planA4Sheet.style.height = `${Math.floor(sheetH)}px`;
  els.planA4Sheet.style.transform = `scale(${gridPlanState.zoom})`;
}

function setPlanZoom(nextZoom) {
  gridPlanState.zoom = Math.max(0.4, Math.min(2.5, Number(nextZoom) || 1));
  if (els.planZoomResetBtn) {
    els.planZoomResetBtn.textContent = `${Math.round(gridPlanState.zoom * 100)}%`;
  }
  fitPlanSheetIntoStage();
}

function applyRoomSizeFromInputs() {
  const rw = Math.max(1, Math.min(120, Number(els.roomTilesXInput?.value || 7)));
  const rh = Math.max(1, Math.min(120, Number(els.roomTilesYInput?.value || 8)));
  const base = Math.max(0.1, Math.min(10, Number(els.tileSizeInput?.value || 0.6)));
  const tx = Math.max(0.1, Math.min(10, Number(els.tileSizeXInput?.value || base)));
  const ty = Math.max(0.1, Math.min(10, Number(els.tileSizeYInput?.value || base)));
  gridPlanState.tileSize = base;
  gridPlanState.roomWidth = Math.round(rw);
  gridPlanState.roomHeight = Math.round(rh);
  gridPlanState.tileSizeX = tx;
  gridPlanState.tileSizeY = ty;
  gridPlanState.cols = gridPlanState.roomWidth;
  gridPlanState.rows = gridPlanState.roomHeight;
  gridPlanState.roomOriginX = 0;
  gridPlanState.roomOriginY = 0;
  gridPlanState.cells = new Map();
  for (let y = 0; y < gridPlanState.rows; y += 1) {
    for (let x = 0; x < gridPlanState.cols; x += 1) {
      gridPlanState.cells.set(planCellKey(x, y), createPlanTileData(tx, ty));
    }
  }
  gridPlanState.segments = [];
  gridPlanState.firstSelectedNode = null;
  renderPlanGrid();
}

function initGridPlanEditor() {
  if (!els.planGrid) return;
  if (els.planSetupGenerateBtn) {
    els.planSetupGenerateBtn.addEventListener("click", () => {
      const tx = Math.max(0.1, Math.min(10, Number(els.tileSizeXInput?.value || 0.6)));
      const ty = Math.max(0.1, Math.min(10, Number(els.tileSizeYInput?.value || 0.6)));
      gridPlanState.tileSize = tx;
      gridPlanState.tileSizeX = tx;
      gridPlanState.tileSizeY = ty;
      setPlanSetupStep("tiles-count");
      setPlanEditorStatus(`Dimensioni mattonella salvate (${tx.toFixed(2)}m x ${ty.toFixed(2)}m). Ora inserisci quante mattonelle per lato.`);
    });
  }

  if (els.planSetupDrawGridLandscapeBtn) {
    els.planSetupDrawGridLandscapeBtn.addEventListener("click", () => {
      gridPlanState.a4Orientation = "landscape";
      setPlanZoom(1);
      applyRoomSizeFromInputs();
      setPlanEditorScreen("editor");
      setPlanEditorStatus("Griglia A4 orizzontale generata.");
    });
  }

  if (els.planSetupDrawGridPortraitBtn) {
    els.planSetupDrawGridPortraitBtn.addEventListener("click", () => {
      gridPlanState.a4Orientation = "portrait";
      setPlanZoom(1);
      applyRoomSizeFromInputs();
      setPlanEditorScreen("editor");
      setPlanEditorStatus("Griglia A4 verticale generata.");
    });
  }

  if (els.planZoomInBtn) {
    els.planZoomInBtn.addEventListener("click", () => {
      setPlanZoom(gridPlanState.zoom + 0.1);
    });
  }
  if (els.planZoomOutBtn) {
    els.planZoomOutBtn.addEventListener("click", () => {
      setPlanZoom(gridPlanState.zoom - 0.1);
    });
  }
  if (els.planZoomResetBtn) {
    els.planZoomResetBtn.addEventListener("click", () => {
      setPlanZoom(1);
    });
  }

  els.planGrid.addEventListener("dblclick", (e) => {
    const target = e.target.closest(".plan-cell");
    if (!target) return;
    const x = Number(target.dataset.x);
    const y = Number(target.dataset.y);
    gridPlanState.activeCell = { x, y };
    if (gridPlanState.firstSelectedNode) {
      openNodePicker("segment", e.clientX, e.clientY);
      setPlanEditorStatus("Seleziona il Punto B sulla griglia 10x10.");
      e.preventDefault();
      return;
    }
    openPlanCellMenu(x, y, e.clientX, e.clientY);
    e.preventDefault();
  });

  if (els.planCellMenu) {
    els.planCellMenu.addEventListener("click", (e) => {
      const action = e.target.closest("[data-plan-action]")?.dataset.planAction;
      if (!action || !gridPlanState.activeCell) return;
      if (action === "select-nodes") {
        openNodePicker("segment", e.clientX || 180, e.clientY || 180);
        setPlanEditorStatus("Seleziona il Punto A sulla griglia 10x10.");
        return;
      }
      if (action === "add-tiles") {
        openNodePicker("add-tiles", e.clientX || 180, e.clientY || 180);
      }
    });
  }

  if (els.planNodePickerGrid) {
    els.planNodePickerGrid.addEventListener("click", (e) => {
      const node = e.target.closest(".plan-node-picker__node");
      if (!node) return;
      commitNodeSelection(Number(node.dataset.nx), Number(node.dataset.ny));
    });
  }

  if (els.planNodePickerCloseBtn) {
    els.planNodePickerCloseBtn.addEventListener("click", () => {
      hidePlanMenus();
    });
  }

  if (els.planAddTilesCancelBtn) {
    els.planAddTilesCancelBtn.addEventListener("click", () => {
      gridPlanState.pendingAddTilesAnchor = null;
      hidePlanMenus();
    });
  }
  if (els.planAddTilesApplyBtn) {
    els.planAddTilesApplyBtn.addEventListener("click", () => {
      applyAddTilesFromDialog();
    });
  }

  document.addEventListener("click", (e) => {
    const t = e.target;
    if (els.planCellMenu && els.planCellMenu.contains(t)) return;
    if (els.planNodePicker && els.planNodePicker.contains(t)) return;
    if (els.planAddTilesDialog && els.planAddTilesDialog.contains(t)) return;
    if (t && t.closest && t.closest(".plan-cell")) return;
    hidePlanMenus();
  });

  applyRoomSizeFromInputs();
  setPlanZoom(1);
  setPlanEditorScreen("setup");
  setPlanSetupStep("tile-size");
  updatePlanToolButtonsUi();

  window.addEventListener("resize", () => {
    if (els.planEditorModal && !els.planEditorModal.hidden && els.planEditorBody && !els.planEditorBody.hidden) {
      fitPlanSheetIntoStage();
      fitPlanGridIntoA4();
      renderPlanGrid();
    }
  });
}

function openPlanEditorModal() {
  if (els.planEditorModal) els.planEditorModal.hidden = false;
  document.body.style.overflow = "hidden";
  setPlanEditorScreen("setup");
  setPlanSetupStep("tile-size");
  setPlanEditorStatus("Inserisci solo le dimensioni orizzontale e verticale.");
  requestAnimationFrame(() => {
    fitPlanGridIntoA4();
  });
}

function closePlanEditorModal() {
  if (els.planEditorModal) els.planEditorModal.hidden = true;
  document.body.style.overflow = "";
}

function getCurrentMainAreaView() {
  if (els.mapAreaPanel && !els.mapAreaPanel.hidden) return "map";
  if (els.acceptanceAreaPanel && !els.acceptanceAreaPanel.hidden) return "acceptance";
  return "tables";
}

function waitForNextFrame() {
  return new Promise((resolve) => {
    requestAnimationFrame(() => requestAnimationFrame(resolve));
  });
}

function waitForFloorImageReady(timeoutMs = 1200) {
  return new Promise((resolve) => {
    const img = els.floorImg;
    if (!img || img.hidden || !img.getAttribute("src")) {
      resolve();
      return;
    }
    if (img.complete) {
      resolve();
      return;
    }
    let done = false;
    const finish = () => {
      if (done) return;
      done = true;
      clearTimeout(timer);
      img.removeEventListener("load", finish);
      img.removeEventListener("error", finish);
      resolve();
    };
    const timer = setTimeout(finish, timeoutMs);
    img.addEventListener("load", finish, { once: true });
    img.addEventListener("error", finish, { once: true });
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
}

function applyFloorPlanFromState() {
  if (state.floorPlanDataUrl) {
    els.floorImg.src = state.floorPlanDataUrl;
    els.floorImg.hidden = false;
    els.floorHint.hidden = true;
  } else {
    els.floorImg.removeAttribute("src");
    els.floorImg.hidden = true;
    els.floorHint.hidden = false;
  }
}

function renderFloorMarkers() {
  els.floorMarkers.innerHTML = "";
  for (const table of state.tables) {
    const pos = state.markerPositions[table.id] || { x: 50, y: 50 };
    const marker = document.createElement("div");
    marker.className = "table-marker";
    marker.style.left = `${pos.x}%`;
    marker.style.top = `${pos.y}%`;
    marker.dataset.tableId = table.id;

    const circle = document.createElement("button");
    circle.type = "button";
    circle.className = "table-marker__circle";
    circle.textContent = String(table.number);
    circle.setAttribute("aria-label", `Tavolo ${table.number}, trascina sulla piantina`);

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
    attachMarkerDrag(marker, circle, table.id);
    els.floorMarkers.appendChild(marker);
  }
  layoutMarkerNotes();
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

function layoutMarkerNotes() {
  const markers = Array.from(els.floorMarkers.querySelectorAll(".table-marker"));
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
  const stageRect = els.floorMarkers.getBoundingClientRect();
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

function attachMarkerDrag(containerEl, handleEl, tableId) {
  let startX, startY, origLeft, origTop, rect;

  function onPointerMove(e) {
    if (!rect) return;
    const dx = e.clientX - startX;
    const dy = e.clientY - startY;
    let x = origLeft + (dx / rect.width) * 100;
    let y = origTop + (dy / rect.height) * 100;
    x = Math.max(2, Math.min(98, x));
    y = Math.max(2, Math.min(98, y));
    containerEl.style.left = `${x}%`;
    containerEl.style.top = `${y}%`;
    layoutMarkerNotes();
  }

  function onPointerUp(e) {
    handleEl.releasePointerCapture(e.pointerId);
    handleEl.removeEventListener("pointermove", onPointerMove);
    handleEl.removeEventListener("pointerup", onPointerUp);
    handleEl.removeEventListener("pointercancel", onPointerUp);
    const left = parseFloat(containerEl.style.left);
    const top = parseFloat(containerEl.style.top);
    state.markerPositions[tableId] = { x: left, y: top };
    saveState();
    rect = null;
    layoutMarkerNotes();
  }

  handleEl.addEventListener("pointerdown", (e) => {
    if (e.button !== 0) return;
    rect = els.floorMarkers.getBoundingClientRect();
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

function setPlanEditorStatus(message, isError = false) {
  if (!els.planEditorStatus) return;
  els.planEditorStatus.textContent = message;
  els.planEditorStatus.style.color = isError ? "#9a3b3b" : "#5f564b";
}

function ensurePlanCanvas() {
  if (sketchPlanState.fabricCanvas) return sketchPlanState.fabricCanvas;
  if (!window.fabric || !els.planCanvas) return null;
  const canvas = new window.fabric.Canvas("planCanvas", {
    selection: true,
    preserveObjectStacking: true,
  });
  sketchPlanState.fabricCanvas = canvas;

  canvas.on("selection:created", (e) => handlePlanSelection(e.selected && e.selected[0]));
  canvas.on("selection:updated", (e) => handlePlanSelection(e.selected && e.selected[0]));
  canvas.on("selection:cleared", () => handlePlanSelection(null));
  canvas.on("object:moving", (e) => syncZoneLabelForObject(e.target));
  canvas.on("object:scaling", (e) => syncZoneLabelForObject(e.target));
  canvas.on("object:rotating", (e) => syncZoneLabelForObject(e.target));
  return canvas;
}

function resizePlanCanvasToWrap(widthPx, heightPx) {
  const canvas = ensurePlanCanvas();
  if (!canvas) return;
  const wrapW = els.planCanvasWrap ? Math.max(320, els.planCanvasWrap.clientWidth - 2) : widthPx;
  const ratio = widthPx / Math.max(1, heightPx);
  const outW = Math.min(wrapW, widthPx);
  const outH = Math.round(outW / Math.max(0.1, ratio));
  canvas.setWidth(outW);
  canvas.setHeight(outH);
  canvas.renderAll();
}

function waitForCvReady(timeoutMs = 8000) {
  return new Promise((resolve, reject) => {
    const start = Date.now();
    const tick = () => {
      if (window.cv && typeof window.cv.imread === "function" && typeof window.cv.Mat === "function") {
        resolve(window.cv);
        return;
      }
      if (Date.now() - start > timeoutMs) {
        reject(new Error("OpenCV.js non disponibile."));
        return;
      }
      setTimeout(tick, 120);
    };
    tick();
  });
}

function loadImageElementFromFile(file) {
  return new Promise((resolve, reject) => {
    const url = URL.createObjectURL(file);
    const img = new Image();
    img.onload = () => {
      URL.revokeObjectURL(url);
      resolve(img);
    };
    img.onerror = () => {
      URL.revokeObjectURL(url);
      reject(new Error("Immagine non valida."));
    };
    img.src = url;
  });
}

function loadImageElementFromDataUrl(dataUrl) {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve(img);
    img.onerror = () => reject(new Error("Immagine non valida."));
    img.src = dataUrl;
  });
}

function handlePlanSelection(obj) {
  sketchPlanState.selectedZone = obj && obj.customType === "zone" ? obj : null;
  if (els.planZoneLabelInput) {
    els.planZoneLabelInput.value = sketchPlanState.selectedZone ? String(sketchPlanState.selectedZone.zoneLabel || "") : "";
  }
}

function syncZoneLabelForObject(obj) {
  if (!obj || obj.customType !== "zone" || !obj.linkedLabel) return;
  const label = obj.linkedLabel;
  const center = obj.getCenterPoint();
  label.set({
    left: center.x,
    top: center.y,
    angle: obj.angle || 0,
  });
  label.setCoords();
  const canvas = ensurePlanCanvas();
  canvas?.renderAll();
}

function snapshotPlanBaseLayout() {
  const canvas = ensurePlanCanvas();
  if (!canvas) return;
  sketchPlanState.baseSnapshot = canvas
    .getObjects()
    .filter((o) => o.customType !== "background")
    .map((o) => ({
      obj: o,
      left: o.left,
      top: o.top,
      scaleX: o.scaleX || 1,
      scaleY: o.scaleY || 1,
    }));
  sketchPlanState.globalScale = 1;
  if (els.planGlobalScale) els.planGlobalScale.value = "100";
}

function applyGlobalPlanScale(factor) {
  const canvas = ensurePlanCanvas();
  if (!canvas) return;
  const cx = canvas.getWidth() / 2;
  const cy = canvas.getHeight() / 2;
  sketchPlanState.baseSnapshot.forEach((entry) => {
    const obj = entry.obj;
    obj.set({
      left: cx + (entry.left - cx) * factor,
      top: cy + (entry.top - cy) * factor,
      scaleX: entry.scaleX * factor,
      scaleY: entry.scaleY * factor,
    });
    obj.setCoords();
    syncZoneLabelForObject(obj);
  });
  sketchPlanState.globalScale = factor;
  canvas.renderAll();
}

async function analyzeSketchToInteractivePlan(source) {
  const canvas = ensurePlanCanvas();
  if (!canvas) {
    setPlanEditorStatus("Fabric.js non disponibile.", true);
    return;
  }
  const cv = await waitForCvReady();
  const img =
    typeof source === "string"
      ? await loadImageElementFromDataUrl(source)
      : await loadImageElementFromFile(source);
  const maxSide = 1200;
  const scale = Math.min(1, maxSide / Math.max(img.width, img.height));
  const workW = Math.max(640, Math.round(img.width * scale));
  const workH = Math.max(420, Math.round(img.height * scale));
  resizePlanCanvasToWrap(workW, workH);

  const tmp = document.createElement("canvas");
  tmp.width = workW;
  tmp.height = workH;
  const tctx = tmp.getContext("2d");
  tctx.drawImage(img, 0, 0, workW, workH);

  const src = cv.imread(tmp);
  const gray = new cv.Mat();
  const binary = new cv.Mat();
  const lines = new cv.Mat();
  const contours = new cv.MatVector();
  const hierarchy = new cv.Mat();

  cv.cvtColor(src, gray, cv.COLOR_RGBA2GRAY, 0);
  cv.GaussianBlur(gray, gray, new cv.Size(3, 3), 0, 0, cv.BORDER_DEFAULT);
  cv.threshold(gray, binary, 0, 255, cv.THRESH_BINARY_INV + cv.THRESH_OTSU);
  cv.HoughLinesP(binary, lines, 1, Math.PI / 180, 55, 40, 10);
  cv.findContours(binary, contours, hierarchy, cv.RETR_EXTERNAL, cv.CHAIN_APPROX_SIMPLE);

  canvas.clear();
  const bg = new window.fabric.Image(tmp, {
    selectable: false,
    evented: false,
    opacity: 0.23,
    left: 0,
    top: 0,
    customType: "background",
  });
  canvas.add(bg);

  let lineCount = 0;
  for (let i = 0; i < lines.rows; i += 1) {
    const p = lines.data32S;
    const x1 = p[i * 4 + 0];
    const y1 = p[i * 4 + 1];
    const x2 = p[i * 4 + 2];
    const y2 = p[i * 4 + 3];
    const wall = new window.fabric.Line([x1, y1, x2, y2], {
      stroke: "#4b4136",
      strokeWidth: 2,
      selectable: true,
      hasControls: true,
      customType: "wall",
    });
    canvas.add(wall);
    lineCount += 1;
  }

  let zoneCount = 0;
  for (let i = 0; i < contours.size(); i += 1) {
    const cnt = contours.get(i);
    const area = cv.contourArea(cnt);
    if (area < 1800 || area > workW * workH * 0.8) continue;
    const peri = cv.arcLength(cnt, true);
    const approx = new cv.Mat();
    cv.approxPolyDP(cnt, approx, 0.02 * peri, true);
    if (approx.rows < 4) {
      approx.delete();
      continue;
    }
    const points = [];
    for (let j = 0; j < approx.rows; j += 1) {
      points.push({ x: approx.data32S[j * 2], y: approx.data32S[j * 2 + 1] });
    }
    approx.delete();
    const zone = new window.fabric.Polygon(points, {
      fill: "rgba(140,105,30,0.14)",
      stroke: "#8b6914",
      strokeWidth: 2,
      objectCaching: false,
      customType: "zone",
      zoneId: `z_${sketchPlanState.nextZoneId++}`,
      zoneLabel: "",
    });
    canvas.add(zone);
    zoneCount += 1;
  }

  src.delete();
  gray.delete();
  binary.delete();
  lines.delete();
  contours.delete();
  hierarchy.delete();

  if (els.planEditor) els.planEditor.hidden = false;
  canvas.renderAll();
  snapshotPlanBaseLayout();
  setPlanEditorStatus(`Rilevamento completato: ${lineCount} linee e ${zoneCount} aree selezionabili.`);
}

function createPdfDocument(orientation = "portrait") {
  if (!window.jspdf || !window.jspdf.jsPDF) return null;
  return new window.jspdf.jsPDF({ orientation, unit: "mm", format: "a4" });
}

function createPdfDocumentA5Portrait() {
  if (!window.jspdf || !window.jspdf.jsPDF) return null;
  return new window.jspdf.jsPDF({ orientation: "portrait", unit: "mm", format: "a5" });
}

const SEGNATAVOLO_STYLE_KEY = "evento-segnatavolo-style-v1";

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
  { id: "cresima", label: "Cresima", themeId: "t1", paletteId: "p4", graphicIndex: 2 },
  { id: "laurea", label: "Laurea", themeId: "t7", paletteId: "p9", graphicIndex: 8 },
  { id: "diploma", label: "Diploma / maturità", themeId: "t7", paletteId: "p4", graphicIndex: 9 },
  { id: "promessa", label: "Promessa di nozze", themeId: "t3", paletteId: "p6", graphicIndex: 2 },
  { id: "matrimonio", label: "Matrimonio", themeId: "t1", paletteId: "p3", graphicIndex: 1 },
  { id: "anniversario", label: "Anniversario", themeId: "t5", paletteId: "p5", graphicIndex: 7 },
  { id: "san_valentino", label: "San Valentino", themeId: "t3", paletteId: "p8", graphicIndex: 4 },
  { id: "nascita", label: "Nascita", themeId: "t10", paletteId: "p6", graphicIndex: 4 },
  { id: "compleanno", label: "Compleanno", themeId: "t10", paletteId: "p7", graphicIndex: 3 },
  { id: "diciottesimo", label: "Diciottesimo", themeId: "t8", paletteId: "p3", graphicIndex: 9 },
  { id: "addio_festa", label: "Addio al nubilato/celibato", themeId: "t9", paletteId: "p2", graphicIndex: 8 },
  { id: "pensionamento", label: "Pensionamento", themeId: "t4", paletteId: "p9", graphicIndex: 0 },
  { id: "inaugurazione", label: "Inaugurazione", themeId: "t6", paletteId: "p8", graphicIndex: 3 },
  { id: "festa_donna", label: "Festa della donna", themeId: "t3", paletteId: "p7", graphicIndex: 2 },
  { id: "carnevale", label: "Carnevale", themeId: "t6", paletteId: "p6", graphicIndex: 6 },
  { id: "pasqua", label: "Pasqua", themeId: "t2", paletteId: "p6", graphicIndex: 4 },
  { id: "ferragosto", label: "Ferragosto", themeId: "t2", paletteId: "p2", graphicIndex: 7 },
  { id: "halloween", label: "Halloween", themeId: "t8", paletteId: "p8", graphicIndex: 5 },
  { id: "natale", label: "Natale", themeId: "t1", paletteId: "p3", graphicIndex: 4 },
  { id: "capodanno", label: "Capodanno", themeId: "t7", paletteId: "p2", graphicIndex: 6 },
];

let segnatavoloSettings = {
  themeId: "t1",
  paletteId: "p1",
  graphicIndex: 0,
  /** Illustrazioni tematiche ai margini (null = solo grafica geometrica sotto). */
  festiveMarginId: null,
  customTitleHex: "#3c2614",
  customBodyHex: "#2c2620",
  customDecoHex: "#c4a35a",
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
    }
  } catch (_) {}
  if (els.segnatavoloColorTitle) {
    els.segnatavoloColorTitle.value = segnatavoloSettings.customTitleHex;
    els.segnatavoloColorBody.value = segnatavoloSettings.customBodyHex;
    els.segnatavoloColorDeco.value = segnatavoloSettings.customDecoHex;
  }
}

function saveSegnatavoloSettings() {
  if (els.segnatavoloColorTitle) {
    segnatavoloSettings.customTitleHex = els.segnatavoloColorTitle.value;
    segnatavoloSettings.customBodyHex = els.segnatavoloColorBody.value;
    segnatavoloSettings.customDecoHex = els.segnatavoloColorDeco.value;
  }
  try {
    localStorage.setItem(SEGNATAVOLO_STYLE_KEY, JSON.stringify(segnatavoloSettings));
  } catch (_) {}
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
      pdf.setFillColor(240, 235, 255);
      pdf.lines([[0, 0], [1.5, 0], [1.8, 5], [0.9, 6.5], [0, 5]], W - 10, H - 12, [1, 1], "F", true);
      pdf.setDrawColor(...tCol);
      pdf.lines([[0, 0], [1.5, 0], [1.8, 5], [0.9, 6.5], [0, 5]], W - 10, H - 12, [1, 1], "S", true);
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
      for (let i = 0; i < 7; i++) {
        const ang = Math.PI * 0.15 + i * 0.35;
        const cx = 6 + Math.cos(ang) * 5.5;
        const cy = H - 8 + Math.sin(ang) * 4;
        pdf.setFillColor(...green);
        pdf.circle(cx, cy, 0.55, "F");
      }
      for (let i = 0; i < 7; i++) {
        const ang = Math.PI * 0.5 + i * 0.35;
        const cx = W - 6 - Math.cos(ang) * 5.5;
        const cy = H - 8 + Math.sin(ang) * 4;
        pdf.circle(cx, cy, 0.55, "F");
      }
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
      pdf.setFillColor(...pink);
      pdf.lines([[0, 0], [1.2, -1], [2.4, 0], [1.2, 1.8]], W - 12, H - 10, [1, 1], "F", true);
      pdf.lines([[0, 0], [1, -0.8], [2, 0], [1, 1.5]], W - 9, H - 12, [1, 1], "F", true);
      pdf.setFillColor(...white);
      for (let i = 0; i < 5; i++) {
        const ang = (i / 5) * Math.PI * 2;
        pdf.circle(12 + Math.cos(ang) * 1.8, 7 + Math.sin(ang) * 1.8, 0.35, "F");
      }
      pdf.setDrawColor(...tCol);
      pdf.setLineWidth(0.2);
      pdf.line(W - 14, 7, W - 10, 10);
      pdf.line(W - 10, 10, W - 6, 7);
      if (festiveId === "matrimonio") {
        pdf.setFillColor(...gold);
        pdf.lines([[0, 0], [1.2, 0], [1.4, 2.5], [-1.4, 2.5], [-1.2, 0]], 6, H - 12, [1, 1], "F", true);
        pdf.lines([[0, 0], [1.2, 0], [1.4, 2.5], [-1.4, 2.5], [-1.2, 0]], 9.5, H - 12, [1, 1], "F", true);
      }
      break;
    }
    case "battesimo": {
      const shell = festiveMix(sky, white, 0.4);
      pdf.setFillColor(...shell);
      pdf.lines([[0, 0], [1.2, -0.8], [2.2, -2], [2.8, -3.5], [2.5, -5], [1.5, -5.8], [0, -5.5], [-1.5, -4.5], [-2, -2.5], [-1.2, -0.8]], 9, H - 8, [1, 1], "F", true);
      pdf.setFillColor(...sky);
      pdf.circle(5, H - 6, 0.55, "F");
      pdf.circle(6.8, H - 5.2, 0.45, "F");
      pdf.circle(4.2, H - 7.5, 0.4, "F");
      pdf.setDrawColor(...gold);
      pdf.setLineWidth(0.35);
      pdf.line(W - 10, H - 12, W - 10, H - 7);
      pdf.line(W - 12.5, H - 9.5, W - 7.5, H - 9.5);
      pdf.setFillColor(255, 250, 200);
      for (let i = 0; i < 4; i++) {
        pdf.circle(7 + i * 2.2, 9, 0.35, "F");
      }
      pdf.setFillColor(...pink);
      pdf.circle(W - 7, 11, 1.2, "F");
      pdf.circle(W - 5.5, 9.5, 0.5, "F");
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
        <polygon points="${W - 10},${H - 12} ${W - 8.5},${H - 12} ${W - 8.2},${H - 7} ${W - 10.9},${H - 7}" fill="rgb(240,235,255)" stroke="${c(tCol)}" stroke-width="0.2"/>
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
        ${[0, 1, 2, 3, 4, 5, 6]
          .map((i) => {
            const ang = Math.PI * 0.15 + i * 0.35;
            const cx = 6 + Math.cos(ang) * 5.5;
            const cy = H - 8 + Math.sin(ang) * 4;
            return `<circle cx="${cx.toFixed(2)}" cy="${cy.toFixed(2)}" r="0.55" fill="${c(green)}"/>`;
          })
          .join("")}
        ${[0, 1, 2, 3, 4, 5, 6]
          .map((i) => {
            const ang = Math.PI * 0.5 + i * 0.35;
            const cx = W - 6 - Math.cos(ang) * 5.5;
            const cy = H - 8 + Math.sin(ang) * 4;
            return `<circle cx="${cx.toFixed(2)}" cy="${cy.toFixed(2)}" r="0.55" fill="${c(green)}"/>`;
          })
          .join("")}
        <path d="M${W - 14} 8 L${W - 8} 12 L${W - 6} 9" fill="none" stroke="${c(gold)}" stroke-width="0.35"/>
      </svg>`;
    case "promessa":
    case "matrimonio":
      return `${svgStart}
        <circle cx="7.5" cy="11" r="2.8" fill="none" stroke="${c(gold)}" stroke-width="0.35"/>
        <circle cx="10.5" cy="11" r="2.8" fill="none" stroke="${c(gold)}" stroke-width="0.35"/>
        <polygon points="${W - 12},${H - 10} ${W - 10.8},${H - 11} ${W - 9.6},${H - 10} ${W - 10.8},${H - 8.2}" fill="${c(pink)}"/>
        <polygon points="${W - 9},${H - 12} ${W - 8},${H - 12.8} ${W - 7},${H - 12} ${W - 8},${H - 10.5}" fill="${c(pink)}"/>
        <g fill="${c(white)}">
          ${[0, 1, 2, 3, 4]
            .map((i) => {
              const ang = (i / 5) * Math.PI * 2;
              return `<circle cx="${12 + Math.cos(ang) * 1.8}" cy="${7 + Math.sin(ang) * 1.8}" r="0.35"/>`;
            })
            .join("")}
        </g>
        <path d="M${W - 14} 7 L${W - 10} 10 L${W - 6} 7" fill="none" stroke="${c(tCol)}" stroke-width="0.2"/>
        ${
          festiveId === "matrimonio"
            ? `<polygon points="6,${H - 12} 7.2,${H - 12} 7.4,${H - 9.5} 5.6,${H - 9.5}" fill="${c(gold)}"/>
               <polygon points="9.5,${H - 12} 10.7,${H - 12} 10.9,${H - 9.5} 9.1,${H - 9.5}" fill="${c(gold)}"/>`
            : ""
        }
      </svg>`;
    case "battesimo":
      return `${svgStart}
        <path d="M9 ${H - 8} L10.2 ${H - 8.8} L11 ${H - 10} L11.8 ${H - 11.5} L11.5 ${H - 13} L10.5 ${H - 13.8} L9 ${H - 13.5} L7.5 ${H - 12.5} L7 ${H - 10.5} L7.8 ${H - 8.8} Z" fill="${c(festiveMix(sky, white, 0.4))}"/>
        <circle cx="5" cy="${H - 6}" r="0.55" fill="${c(sky)}"/>
        <circle cx="6.8" cy="${H - 5.2}" r="0.45" fill="${c(sky)}"/>
        <path d="M${W - 10} ${H - 12} L${W - 10} ${H - 7} M${W - 12.5} ${H - 9.5} L${W - 7.5} ${H - 9.5}" stroke="${c(gold)}" fill="none" stroke-width="0.35"/>
        <g fill="rgb(255,250,200)">${[0, 1, 2, 3].map((i) => `<circle cx="${7 + i * 2.2}" cy="9" r="0.35"/>`).join("")}</g>
        <circle cx="${W - 7}" cy="11" r="1.2" fill="${c(pink)}"/>
        <circle cx="${W - 5.5}" cy="9.5" r="0.5" fill="${c(pink)}"/>
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
  const W = pdf.internal.pageSize.getWidth();
  const H = pdf.internal.pageSize.getHeight();
  const innerW = W - 40;
  const cx = W / 2;
  const names = getTableGuestNamesOrdered(table);
  const custom = (table.customName || "").trim();
  const mainTitle = custom || `Tavolo ${table.number}`;
  const foot = (state.eventName || "").trim() || "Evento";

  const titleLineH = (sz) => Math.max(5.5, sz * 0.4);
  const bodyLineH = (sz) => Math.max(5, sz * 0.42);

  function drawFooter() {
    pdf.setFont(theme.footFamily, theme.footStyle);
    pdf.setFontSize(theme.footSize);
    pdf.setTextColor(...palette.foot);
    pdf.text(foot, cx, H - 7, { align: "center" });
    pdf.setTextColor(0, 0, 0);
  }

  function drawPageHeader(startY, isContinuation) {
    const fid = settings.festiveMarginId || null;
    if (fid) drawSegnatavoloFestiveMarginArt(pdf, W, H, palette, fid);
    else drawSegnatavoloGraphic(pdf, W, H, palette.deco, gi);
    let y = startY;
    pdf.setFont(theme.titleFamily, theme.titleStyle);
    pdf.setFontSize(theme.titleSize);
    pdf.setTextColor(...palette.title);
    const titleLines = pdf.splitTextToSize(mainTitle, innerW);
    for (const line of titleLines) {
      pdf.text(line, cx, y, { align: "center" });
      y += titleLineH(theme.titleSize);
    }
    if (custom) {
      pdf.setFont(theme.subFamily, theme.subStyle);
      pdf.setFontSize(theme.subSize);
      pdf.setTextColor(...palette.subtitle);
      pdf.text(`Tavolo ${table.number}`, cx, y + 0.5, { align: "center" });
      y += titleLineH(theme.subSize) + 2;
    } else {
      y += 3;
    }
    y += 3;
    pdf.setDrawColor(...palette.deco);
    pdf.setLineWidth(0.35);
    pdf.line(20, y, W - 20, y);
    y += 7;
    if (isContinuation) {
      pdf.setFont(theme.bodyFamily, "italic");
      pdf.setFontSize(theme.subSize);
      pdf.setTextColor(...palette.subtitle);
      pdf.text("(segue)", cx, y, { align: "center" });
      y += 6;
    }
    pdf.setFont(theme.bodyFamily, theme.bodyStyle);
    pdf.setFontSize(theme.bodySize);
    pdf.setTextColor(...palette.body);
    return y;
  }

  const maxY = H - 11;
  const lh = bodyLineH(theme.bodySize);
  let nameIdx = 0;
  let isFirstPage = true;

  for (;;) {
    if (!isFirstPage) pdf.addPage();
    isFirstPage = false;
    const continuation = nameIdx > 0;
    let y = drawPageHeader(continuation ? 16 : 18, continuation);

    if (names.length === 0) {
      pdf.setFont(theme.bodyFamily, "italic");
      pdf.setFontSize(theme.subSize);
      pdf.setTextColor(...palette.subtitle);
      pdf.text("Nessun ospite in elenco.", cx, y, { align: "center" });
      drawFooter();
      break;
    }

    let wroteOnPage = false;
    while (nameIdx < names.length) {
      const wrapped = pdf.splitTextToSize(names[nameIdx], innerW);
      const blockH = wrapped.length * lh;
      if (wroteOnPage && y + blockH > maxY) break;
      for (const line of wrapped) {
        pdf.text(line, cx, y, { align: "center" });
        y += lh;
      }
      nameIdx += 1;
      wroteOnPage = true;
    }
    drawFooter();
    if (nameIdx >= names.length) break;
  }
  pdf.setTextColor(0, 0, 0);
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
  state.eventName = els.eventName.value;
  saveState();
});

els.addTableBtn.addEventListener("click", addTable);

if (els.showTablesAreaBtn) {
  els.showTablesAreaBtn.addEventListener("click", () => {
    setMainAreaView("tables");
  });
}

if (els.showMapAreaBtn) {
  els.showMapAreaBtn.addEventListener("click", () => {
    setMainAreaView("map");
  });
}

if (els.showAcceptanceAreaBtn) {
  els.showAcceptanceAreaBtn.addEventListener("click", () => {
    setMainAreaView("acceptance");
    if (els.acceptanceSearchInput) {
      els.acceptanceSearchInput.focus();
    }
  });
}

if (els.acceptanceSearchInput) {
  els.acceptanceSearchInput.addEventListener("input", () => {
    acceptanceUiState.query = els.acceptanceSearchInput.value;
    renderAcceptanceArea();
  });
}

els.sortTablesBtn.addEventListener("click", () => {
  closeAllDropdowns();
  sortTablesAscending();
});

els.renumberTablesBtn.addEventListener("click", () => {
  closeAllDropdowns();
  renumberTablesSmart();
});

if (els.tableActionsMenuBtn && els.tableActionsMenu) {
  els.tableActionsMenuBtn.addEventListener("click", (e) => {
    e.stopPropagation();
    toggleDropdown(els.tableActionsMenu, els.tableActionsMenuBtn);
  });
}

if (els.exportMenuBtn && els.exportMenu) {
  els.exportMenuBtn.addEventListener("click", (e) => {
    e.stopPropagation();
    toggleDropdown(els.exportMenu, els.exportMenuBtn);
  });
}

if (els.exportMenu) {
  els.exportMenu.addEventListener("click", (e) => {
    const planBtn = e.target.closest("#exportPlanPdfBtn");
    if (!planBtn) return;
    e.preventDefault();
    exportPlanPdf();
  });
}

if (els.importMenuBtn && els.importMenu) {
  els.importMenuBtn.addEventListener("click", (e) => {
    e.stopPropagation();
    toggleDropdown(els.importMenu, els.importMenuBtn);
  });
}

if (els.importJsonBtn && els.importInput) {
  els.importJsonBtn.addEventListener("click", () => {
    closeAllDropdowns();
    els.importInput.click();
  });
}

document.addEventListener("click", (e) => {
  const t = e.target;
  if (t && t.closest && t.closest("[data-close-dropdown='true']")) {
    closeAllDropdowns();
    return;
  }
  if (els.tableActionsMenu && els.tableActionsMenu.contains(t)) return;
  if (els.exportMenu && els.exportMenu.contains(t)) return;
  if (els.importMenu && els.importMenu.contains(t)) return;
  if (els.tableActionsDropdown && els.tableActionsDropdown.contains(t)) return;
  if (els.exportDropdown && els.exportDropdown.contains(t)) return;
  if (els.importDropdown && els.importDropdown.contains(t)) return;
  closeAllDropdowns();
});

async function applyFloorPlanFromFileInput(inputEl) {
  if (!inputEl) return;
  const file = inputEl.files && inputEl.files[0];
  if (!file) return;
  try {
    state.floorPlanDataUrl = await readFileAsDataUrl(file);
    applyFloorPlanFromState();
    renderFloorMarkers();
    saveState();
  } catch (_) {
    alert("Impossibile leggere il file. Prova con PNG o JPG.");
  }
  inputEl.value = "";
}

function applyFloorPlanDataUrl(dataUrl) {
  state.floorPlanDataUrl = dataUrl;
  applyFloorPlanFromState();
  renderFloorMarkers();
  saveState();
}

async function openFloorCameraCapture() {
  if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) {
    if (els.floorSketchInput) els.floorSketchInput.click();
    return;
  }
  const overlay = document.createElement("div");
  overlay.className = "camera-capture-overlay";
  overlay.innerHTML = `
    <div class="camera-capture-panel">
      <div class="camera-capture-topbar">
        <label class="camera-capture-device-label">
          Fotocamera
          <select class="camera-capture-device-select" data-role="camera-select"></select>
        </label>
      </div>
      <video class="camera-capture-video" autoplay playsinline></video>
      <div class="camera-capture-actions">
        <button type="button" class="btn btn-secondary" data-action="cancel">Annulla</button>
        <button type="button" class="btn btn-primary" data-action="shot">Scatta foto</button>
      </div>
    </div>
  `;
  document.body.appendChild(overlay);
  const video = overlay.querySelector("video");
  const cameraSelect = overlay.querySelector("[data-role='camera-select']");
  let stream = null;

  const stopCurrentStream = () => {
    if (!stream) return;
    stream.getTracks().forEach((track) => track.stop());
    stream = null;
  };

  const startStream = async (deviceId = "") => {
    stopCurrentStream();
    const constraints = deviceId
      ? { video: { deviceId: { exact: deviceId } }, audio: false }
      : { video: { facingMode: { ideal: "environment" } }, audio: false };
    stream = await navigator.mediaDevices.getUserMedia(constraints);
    video.srcObject = stream;
  };

  const rearCameraScore = (label) => {
    const txt = String(label || "").toLowerCase();
    let score = 0;
    if (/rear|back|environment|world|posteriore/.test(txt)) score += 4;
    if (/front|user|anteriore|facetime/.test(txt)) score -= 4;
    return score;
  };

  const loadDevices = async () => {
    const devices = await navigator.mediaDevices.enumerateDevices();
    const cams = devices
      .filter((d) => d.kind === "videoinput")
      .sort((a, b) => rearCameraScore(b.label) - rearCameraScore(a.label));
    cameraSelect.innerHTML = "";
    cams.forEach((cam, idx) => {
      const opt = document.createElement("option");
      opt.value = cam.deviceId;
      opt.textContent = cam.label || `Fotocamera ${idx + 1}`;
      cameraSelect.appendChild(opt);
    });
    cameraSelect.hidden = cams.length <= 1;
    const activeId = stream?.getVideoTracks()[0]?.getSettings()?.deviceId || "";
    if (activeId) cameraSelect.value = activeId;
    const preferred = cams.find((cam) => rearCameraScore(cam.label) > 0);
    return { cams, preferredId: preferred ? preferred.deviceId : "" };
  };

  try {
    await startStream();
    const { cams, preferredId } = await loadDevices();
    const activeId = stream?.getVideoTracks()[0]?.getSettings()?.deviceId || "";
    if (preferredId && preferredId !== activeId && cams.length > 1) {
      await startStream(preferredId);
      await loadDevices();
    }
  } catch (_) {
    overlay.remove();
    if (els.floorSketchInput) els.floorSketchInput.click();
    return;
  }

  const stopAndClose = () => {
    stopCurrentStream();
    overlay.remove();
  };

  cameraSelect?.addEventListener("change", async () => {
    const nextId = cameraSelect.value;
    if (!nextId) return;
    try {
      await startStream(nextId);
      await loadDevices();
    } catch (_) {
      setPlanEditorStatus("Impossibile attivare la fotocamera selezionata.", true);
    }
  });

  overlay.addEventListener("click", async (e) => {
    const action = e.target && e.target.dataset ? e.target.dataset.action : "";
    if (action === "cancel") {
      stopAndClose();
      return;
    }
    if (action !== "shot") return;

    const canvas = document.createElement("canvas");
    const vw = video.videoWidth || 1280;
    const vh = video.videoHeight || 720;
    canvas.width = vw;
    canvas.height = vh;
    const ctx = canvas.getContext("2d");
    if (!ctx) {
      stopAndClose();
      return;
    }
    ctx.drawImage(video, 0, 0, vw, vh);
    const dataUrl = canvas.toDataURL("image/jpeg", 0.88);
    applyFloorPlanDataUrl(dataUrl);
    try {
      setPlanEditorStatus("Analisi bozza in corso...");
      await analyzeSketchToInteractivePlan(dataUrl);
    } catch (_) {
      setPlanEditorStatus("Piantina caricata, ma analisi interattiva non disponibile sul dispositivo.", true);
    }
    stopAndClose();
  });
}

if (els.floorPlanInput) {
  els.floorPlanInput.addEventListener("change", () => {
    applyFloorPlanFromFileInput(els.floorPlanInput);
  });
}

if (els.createFloorPlanBtn) {
  els.createFloorPlanBtn.addEventListener("click", () => {
    openPlanEditorModal();
    setPlanEditorStatus("Editor griglia attivo: scegli modalita e disegna la sala con click o trascinamento.");
  });
}

if (els.planEditorBackdrop) {
  els.planEditorBackdrop.addEventListener("click", () => {
    closePlanEditorModal();
  });
}

if (els.planEditorCloseBtn) {
  els.planEditorCloseBtn.addEventListener("click", () => {
    closePlanEditorModal();
  });
}

els.clearPlanBtn.addEventListener("click", () => {
  state.floorPlanDataUrl = "";
  applyFloorPlanFromState();
  saveState();
});

els.exportBtn.addEventListener("click", () => {
  closeAllDropdowns();
  const blob = new Blob(
    [
      JSON.stringify(
        {
          eventName: state.eventName,
          floorPlanDataUrl: state.floorPlanDataUrl,
          tables: state.tables,
          markerPositions: state.markerPositions,
          exportedAt: new Date().toISOString(),
        },
        null,
        2
      ),
    ],
    { type: "application/json" }
  );
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "evento-tavoli.json";
  a.click();
  URL.revokeObjectURL(a.href);
});

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

function setSegnatavoloPreviewDecoDom() {
  if (!els.segnatavoloPreviewDeco) return;
  const fid = segnatavoloSettings.festiveMarginId || null;
  const base = "segnatavolo-preview__deco ";
  if (fid) {
    const pal = getSegnatavoloPaletteResolved();
    const inner = buildSegnatavoloFestivePreviewInnerHTML(pal, fid);
    if (inner) {
      els.segnatavoloPreviewDeco.className = base + "segnatavolo-festive-wrap";
      els.segnatavoloPreviewDeco.innerHTML = inner;
      return;
    }
  }
  const gi = segnatavoloSettings.graphicIndex;
  if (gi === 2) {
    els.segnatavoloPreviewDeco.className = base + "segnatavolo-deco--2";
    els.segnatavoloPreviewDeco.innerHTML =
      '<span class="segnatavolo-deco__bl"></span><span class="segnatavolo-deco__br"></span>';
  } else if (gi === 4) {
    els.segnatavoloPreviewDeco.className = base + "segnatavolo-deco--4";
    els.segnatavoloPreviewDeco.innerHTML =
      '<span class="segnatavolo-deco__c3"></span><span class="segnatavolo-deco__c4"></span>';
  } else if (gi === 5) {
    els.segnatavoloPreviewDeco.className = base + "segnatavolo-deco--5";
    els.segnatavoloPreviewDeco.innerHTML =
      '<span class="segnatavolo-deco__d3"></span><span class="segnatavolo-deco__d4"></span>';
  } else {
    els.segnatavoloPreviewDeco.innerHTML = "";
    els.segnatavoloPreviewDeco.className = base + (gi === 0 ? "" : `segnatavolo-deco--${gi}`);
  }
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

function updateSegnatavoloPreview() {
  if (!els.segnatavoloPreview || !els.segnatavoloPreviewTitle) return;
  const table = getSegnatavoloPreviewTable();
  const theme = getSegnatavoloTheme();
  const pal = getSegnatavoloPaletteResolved();
  els.segnatavoloPreview.style.setProperty("--seg-deco", rgbToCss(pal.deco));

  const custom = (table.customName || "").trim();
  const mainTitle = custom || `Tavolo ${table.number}`;
  els.segnatavoloPreviewTitle.style.fontFamily = theme.previewTitle;
  els.segnatavoloPreviewTitle.style.fontSize = `${0.9 * theme.titlePx}rem`;
  els.segnatavoloPreviewTitle.style.color = rgbToCss(pal.title);
  els.segnatavoloPreviewTitle.textContent = mainTitle;

  if (custom) {
    els.segnatavoloPreviewSubtitle.hidden = false;
    els.segnatavoloPreviewSubtitle.style.fontFamily = theme.previewBody;
    els.segnatavoloPreviewSubtitle.style.fontSize = `${0.85 * theme.bodyPx}rem`;
    els.segnatavoloPreviewSubtitle.style.color = rgbToCss(pal.subtitle);
    els.segnatavoloPreviewSubtitle.textContent = `Tavolo ${table.number}`;
  } else {
    els.segnatavoloPreviewSubtitle.hidden = true;
  }

  if (els.segnatavoloPreviewRule) {
    els.segnatavoloPreviewRule.style.background = rgbToCss(pal.deco);
  }

  const names = getTableGuestNamesOrdered(table);
  els.segnatavoloPreviewNames.innerHTML = "";
  names.forEach((n) => {
    const li = document.createElement("li");
    li.textContent = n;
    li.style.fontFamily = theme.previewBody;
    li.style.fontSize = `${0.9 * theme.bodyPx}rem`;
    li.style.color = rgbToCss(pal.body);
    els.segnatavoloPreviewNames.appendChild(li);
  });
  if (els.segnatavoloPreviewEmpty) {
    els.segnatavoloPreviewEmpty.hidden = names.length > 0;
    els.segnatavoloPreviewEmpty.style.color = rgbToCss(pal.subtitle);
  }
  if (els.segnatavoloPreviewFoot) {
    els.segnatavoloPreviewFoot.textContent = (state.eventName || "").trim() || "Evento";
    els.segnatavoloPreviewFoot.style.color = rgbToCss(pal.foot);
  }
  setSegnatavoloPreviewDecoDom();
  refreshSegnatavoloChipSelection();
  updateSegnatavoloCustomColorsVisibility();
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
  if (!els.segnatavoloThemeChips) return;
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
}

function openSegnatavoloModal() {
  if (!els.segnatavoloModal) return;
  loadSegnatavoloSettings();
  updateSegnatavoloPreview();
  els.segnatavoloModal.hidden = false;
  els.segnatavoloModal.setAttribute("aria-hidden", "false");
  document.body.style.overflow = "hidden";
}

function closeSegnatavoloModal() {
  if (!els.segnatavoloModal) return;
  els.segnatavoloModal.hidden = true;
  els.segnatavoloModal.setAttribute("aria-hidden", "true");
  document.body.style.overflow = "";
}

function runSegnatavoliPdfExport() {
  if (!state.tables.length) {
    alert("Non ci sono tavoli da esportare.");
    return;
  }
  const pdf = createPdfDocumentA5Portrait();
  if (!pdf) {
    alert("Libreria PDF non disponibile.");
    return;
  }
  const palette = getSegnatavoloPaletteResolved();
  const opts = {
    themeId: segnatavoloSettings.themeId,
    paletteId: segnatavoloSettings.paletteId,
    graphicIndex: segnatavoloSettings.graphicIndex,
    festiveMarginId: segnatavoloSettings.festiveMarginId || null,
    _palette: palette,
  };
  const ordered = [...state.tables].sort((a, b) => Number(a.number) - Number(b.number));
  ordered.forEach((table, idx) => {
    if (idx > 0) pdf.addPage();
    drawSegnatavoloPage(pdf, table, opts);
  });
  pdf.save("segnatavoli-a5.pdf");
}

function initSegnatavoloModal() {
  buildSegnatavoloDesignerUi();
  loadSegnatavoloSettings();
  if (els.exportSegnatavoliBtn) {
    els.exportSegnatavoliBtn.addEventListener("click", () => {
      closeAllDropdowns();
      if (!state.tables.length) {
        alert("Non ci sono tavoli da esportare.");
        return;
      }
      openSegnatavoloModal();
    });
  }
  if (els.segnatavoloModalClose) els.segnatavoloModalClose.addEventListener("click", closeSegnatavoloModal);
  if (els.segnatavoloModalCancel) els.segnatavoloModalCancel.addEventListener("click", closeSegnatavoloModal);
  if (els.segnatavoloModalBackdrop) els.segnatavoloModalBackdrop.addEventListener("click", closeSegnatavoloModal);
  if (els.segnatavoloExportConfirm) {
    els.segnatavoloExportConfirm.addEventListener("click", () => {
      saveSegnatavoloSettings();
      closeSegnatavoloModal();
      runSegnatavoliPdfExport();
    });
  }
  loadSegnatavoloSettings();
}


async function exportPlanPdf() {
  closeAllDropdowns();
  if (!window.html2canvas || !window.jspdf || !window.jspdf.jsPDF) {
    alert("Librerie PDF non disponibili.");
    return;
  }
  const previousView = getCurrentMainAreaView();
  try {
    if (previousView !== "map") {
      setMainAreaView("map");
      await waitForNextFrame();
    }
    await waitForFloorImageReady();
    const rect = els.floorStage.getBoundingClientRect();
    if (rect.width < 10 || rect.height < 10) {
      alert("Impossibile esportare: area piantina non disponibile al momento.");
      return;
    }
    const canvas = await window.html2canvas(els.floorStage, {
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
    const usableW = pageW - margin * 2;
    const usableH = pageH - margin * 2 - 10;
    const ratio = canvas.width / canvas.height;
    let drawW = usableW;
    let drawH = drawW / ratio;
    if (drawH > usableH) {
      drawH = usableH;
      drawW = drawH * ratio;
    }
    pdf.setFontSize(11);
    pdf.text(state.eventName || "Piantina tavoli", margin, 6);
    pdf.addImage(imgData, "PNG", margin, 10, drawW, drawH);
    pdf.save("piantina-tavoli.pdf");
  } catch (_) {
    alert("Errore durante l'esportazione PDF della piantina.");
  } finally {
    if (getCurrentMainAreaView() !== previousView) {
      setMainAreaView(previousView);
    }
  }
}

els.importInput.addEventListener("change", (e) => {
  const file = e.target.files && e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = () => {
    try {
      const data = JSON.parse(reader.result);
      if (data.eventName != null) state.eventName = String(data.eventName);
      if (data.floorPlanDataUrl != null) state.floorPlanDataUrl = String(data.floorPlanDataUrl);
      if (Array.isArray(data.tables)) state.tables = data.tables;
      state.tables.forEach(ensureTableGuestSlots);
      if (data.markerPositions && typeof data.markerPositions === "object") {
        state.markerPositions = data.markerPositions;
      }
      els.eventName.value = state.eventName;
      applyFloorPlanFromState();
      renderTables();
      renderFloorMarkers();
      saveState();
    } catch (_) {
      alert("File JSON non valido.");
    }
  };
  reader.readAsText(file);
  e.target.value = "";
});

window.addEventListener("resize", () => {
  layoutMarkerNotes();
  updateCurrentTableContextUi();
  fitPlanGridIntoA4();
});

if (els.tablesList) {
  els.tablesList.addEventListener("scroll", updateCurrentTableContextUi, { passive: true });
}

document.addEventListener("keydown", (e) => {
  if (e.key === "Escape") {
    if (els.planEditorModal && !els.planEditorModal.hidden) {
      closePlanEditorModal();
      return;
    }
    if (document.querySelector(".guest-row__sheet:not([hidden])")) {
      closeAllGuestOptionSheets();
      return;
    }
    closeAllDropdowns();
    if (els.segnatavoloModal && !els.segnatavoloModal.hidden) {
      closeSegnatavoloModal();
    }
  }
});

initSegnatavoloModal();
loadState();
ensureExportPlanPdfMenuItem();
initGridPlanEditor();
setMainAreaView("tables");
els.eventName.value = state.eventName;
applyFloorPlanFromState();
renderTables();
renderFloorMarkers();
