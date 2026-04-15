<<<<<<< HEAD
/**
 * Versione cliente: solo tavoli e ospiti (nessuna piantina).
 * Compatibile con lo stesso JSON della versione completa.
 */

const MAX_GUESTS = 9;
const STORAGE_KEY = "evento-tavoli-cliente-v1";

const state = {
  eventName: "",
  floorPlanDataUrl: "",
  tables: [],
  markerPositions: {},
};

const els = {
  eventName: document.getElementById("eventName"),
  addTableBtn: document.getElementById("addTableBtn"),
  sortTablesBtn: document.getElementById("sortTablesBtn"),
  renumberTablesBtn: document.getElementById("renumberTablesBtn"),
  tablesList: document.getElementById("tablesList"),
  tableCountLabel: document.getElementById("tableCountLabel"),
  exportGuestsBtn: document.getElementById("exportGuestsBtn"),
  exportKitchenBtn: document.getElementById("exportKitchenBtn"),
  exportBtn: document.getElementById("exportBtn"),
  importInput: document.getElementById("importInput"),
  tableCardTpl: document.getElementById("tableCardTpl"),
  guestRowTpl: document.getElementById("guestRowTpl"),
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
  while (normalized.length < MAX_GUESTS) normalized.push(createEmptyGuest());
  table.guests = normalized;
}

function uid() {
  return "t_" + Math.random().toString(36).slice(2, 11);
}

function saveState() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

function loadState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return;
    const data = JSON.parse(raw);
    state.eventName = String(data.eventName || "");
    state.floorPlanDataUrl = String(data.floorPlanDataUrl || "");
    state.tables = Array.isArray(data.tables) ? data.tables : [];
    state.markerPositions = data.markerPositions && typeof data.markerPositions === "object" ? data.markerPositions : {};
    state.tables.forEach(ensureTableGuestSlots);
  } catch (_) {}
}

function nextTableNumber() {
  const used = new Set();
  for (const t of state.tables) {
    const n = Number(t.number);
    if (!Number.isNaN(n) && n > 0) used.add(n);
  }
  let candidate = 1;
  while (used.has(candidate)) candidate += 1;
  return candidate;
}

function guestCount(table) {
  return table.guests.filter((g) => (g.cognome || "").trim() || (g.nome || "").trim()).length;
}

function getActiveGuests(table) {
  return table.guests.filter((g) => (g.cognome || "").trim() || (g.nome || "").trim());
}

function getTableMenuCounts(table) {
  const activeGuests = getActiveGuests(table);
  let adulti = 0;
  let bambini = 0;
  for (const guest of activeGuests) {
    if (guest.menu === "bambino") bambini += 1;
    else adulti += 1;
  }
  return { adulti, bambini };
}

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
    items.push({ menu, text, count, letter: menu === "bambino" ? "B" : "A" });
  }
  items.sort((a, b) => {
    if (a.menu !== b.menu) return a.menu === "adulto" ? -1 : 1;
    return a.text.localeCompare(b.text, "it");
  });
  return items.map((i) => `${i.count}${i.letter} ${i.text}`);
}

function getTableSummaryLine(table) {
  const c = getTableMenuCounts(table);
  const parts = [`${c.adulti}A+${c.bambini}B`];
  const allergies = getTableAllergenParts(table);
  if (allergies.length) parts.push(`(${allergies.join(" | ")})`);
  const note = (table.tableNote || "").trim();
  if (note) parts.push(`Note tavolo: ${note}`);
  return parts.join(" — ");
}

function setGuestRowVisualState(rowEl, hasIdentity) {
  rowEl.classList.toggle("guest-row--filled", hasIdentity);
  rowEl.classList.toggle("guest-row--empty", !hasIdentity);
}

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
  const noteEl = sheet.querySelector(".guest-note");
  if (hint) hint.hidden = hasIdentity;
  opts.forEach((o) => {
    o.disabled = !hasIdentity;
  });
  if (noteEl) noteEl.disabled = !hasIdentity;
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
    else sheetEl.querySelector(".guest-row__menu-opt:not(:disabled)")?.focus();
  });
}

function focusGuestField(card, rowIndex, colIndex) {
  const candidate = card.querySelector(
    `.guest-row[data-row-index="${rowIndex}"] [data-grid-col="${colIndex}"]`
  );
  if (!candidate || candidate.disabled) return false;
  candidate.focus();
  if (candidate.tagName === "INPUT") candidate.select();
  return true;
}

function moveGuestGridFocus(card, fromRow, fromCol, key) {
  const vectors = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
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
  const start = fromRow * maxCols + fromCol;
  for (let i = start + 1; i < MAX_GUESTS * maxCols; i += 1) {
    const row = Math.floor(i / maxCols);
    const col = i % maxCols;
    if (focusGuestField(card, row, col)) return;
  }
}

function moveGuestGridPrevField(card, fromRow, fromCol) {
  const maxCols = 2;
  const start = fromRow * maxCols + fromCol;
  for (let i = start - 1; i >= 0; i -= 1) {
    const row = Math.floor(i / maxCols);
    const col = i % maxCols;
    if (focusGuestField(card, row, col)) return;
  }
}

function refreshTableLiveInfoById(tableId) {
  const table = state.tables.find((t) => t.id === tableId);
  if (!table) return;
  const card = els.tablesList.querySelector(`[data-table-id="${CSS.escape(tableId)}"]`);
  if (!card) return;
  const capEl = card.querySelector("[data-capacity]");
  const count = guestCount(table);
  capEl.textContent =
    count >= MAX_GUESTS
      ? `Posti occupati: ${count}/${MAX_GUESTS} (massimo raggiunto)`
      : `Posti occupati: ${count}/${MAX_GUESTS}`;
  capEl.classList.toggle("is-full", count >= MAX_GUESTS);
  card.querySelector("[data-summary]").textContent = `Resoconto: ${getTableSummaryLine(table)}`;
}

function renderTables() {
  els.tablesList.innerHTML = "";
  els.tableCountLabel.textContent = state.tables.length === 1 ? "1 tavolo" : `${state.tables.length} tavoli`;

  for (const table of state.tables) {
    ensureTableGuestSlots(table);
    const node = els.tableCardTpl.content.cloneNode(true);
    const card = node.querySelector(".table-card");
    card.dataset.tableId = table.id;
    node.querySelector(".table-title").textContent = `Tavolo ${table.number}`;

    const summaryEl = node.querySelector("[data-summary]");
    const customNameInput = node.querySelector(".table-custom-name");
    const tableNoteInput = node.querySelector(".table-note");
    const guestsEl = node.querySelector("[data-guests]");
    const cap = node.querySelector("[data-capacity]");

    customNameInput.value = table.customName || "";
    tableNoteInput.value = table.tableNote || "";
    customNameInput.addEventListener("input", () => {
      table.customName = customNameInput.value;
      saveState();
      summaryEl.textContent = `Resoconto: ${getTableSummaryLine(table)}`;
    });
    tableNoteInput.addEventListener("input", () => {
      table.tableNote = tableNoteInput.value;
      saveState();
      summaryEl.textContent = `Resoconto: ${getTableSummaryLine(table)}`;
    });

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

      const saveGuestAndRefresh = () => {
        g.cognome = cognome.value;
        g.nome = nome.value;
        g.menu = menu.value;
        g.note = note.value;
        saveState();
        refreshTableLiveInfoById(table.id);
      };

      const applyIdentityConstraints = () => {
        const hasIdentity = Boolean(cognome.value.trim() || nome.value.trim());
        menu.disabled = !hasIdentity;
        if (moreBtn) moreBtn.disabled = !hasIdentity;
        setGuestRowVisualState(rowEl, hasIdentity);
        applyGuestSheetIdentityUi(sheet, hasIdentity);
        let dirty = false;
        if (!hasIdentity) {
          if (menu.value !== "adulto") {
            menu.value = "adulto";
            g.menu = "adulto";
            dirty = true;
          }
          if (note.value.trim()) {
            note.value = "";
            g.note = "";
            dirty = true;
          }
          syncGuestMenuPickFromValue(sheet, "adulto");
        }
        updateGuestRowIndicatorDots(rowEl, hasIdentity, menu.value, note.value);
        if (dirty) {
          saveState();
          refreshTableLiveInfoById(table.id);
        }
      };

      applyIdentityConstraints();

      cognome.addEventListener("input", () => {
        applyIdentityConstraints();
        saveGuestAndRefresh();
      });
      nome.addEventListener("input", () => {
        applyIdentityConstraints();
        saveGuestAndRefresh();
      });

      menu.addEventListener("change", () => {
        syncGuestMenuPickFromValue(sheet, menu.value);
        updateGuestRowIndicatorDots(rowEl, Boolean(cognome.value.trim() || nome.value.trim()), menu.value, note.value);
        saveGuestAndRefresh();
      });

      note.addEventListener("input", () => {
        updateGuestRowIndicatorDots(rowEl, Boolean(cognome.value.trim() || nome.value.trim()), menu.value, note.value);
        saveGuestAndRefresh();
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

      sheet.querySelector(".guest-row__sheet-scrim")?.addEventListener("click", () => closeAllGuestOptionSheets());
      sheet.querySelector(".guest-row__sheet-done")?.addEventListener("click", () => closeAllGuestOptionSheets());

      row.querySelector(".remove-guest").addEventListener("click", () => {
        table.guests[idx] = createEmptyGuest();
        saveState();
        renderTables();
      });

      [cognome, nome].forEach((fieldEl) => {
        fieldEl.addEventListener("keydown", (e) => {
          const rowIndex = Number(fieldEl.closest(".guest-row")?.dataset.rowIndex);
          const colIndex = Number(fieldEl.dataset.gridCol);
          if (Number.isNaN(rowIndex) || Number.isNaN(colIndex)) return;

          if (e.key === "Enter") {
            e.preventDefault();
            if (e.shiftKey) moveGuestGridPrevField(card, rowIndex, colIndex);
            else moveGuestGridNextField(card, rowIndex, colIndex);
            return;
          }
          if (e.key === "ArrowUp" || e.key === "ArrowDown" || e.key === "ArrowLeft" || e.key === "ArrowRight") {
            e.preventDefault();
            moveGuestGridFocus(card, rowIndex, colIndex, e.key);
          }
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
    summaryEl.textContent = `Resoconto: ${getTableSummaryLine(table)}`;

    node.querySelector(".remove-table").addEventListener("click", () => {
      state.tables = state.tables.filter((t) => t.id !== table.id);
      delete state.markerPositions[table.id];
      saveState();
      renderTables();
    });
    els.tablesList.appendChild(node);
  }
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
  if (!state.markerPositions[id]) state.markerPositions[id] = { x: 50, y: 50 };
  saveState();
  renderTables();
}

function sortTablesAscending() {
  state.tables.sort((a, b) => Number(a.number) - Number(b.number));
  saveState();
  renderTables();
}

function renumberTablesSmart() {
  const emptyTables = state.tables.filter((t) => guestCount(t) === 0);
  const messageLines = [
    "Confermi la rinumerazione intelligente dei tavoli?",
    "",
    `- Tavoli attuali: ${state.tables.length}`,
    `- Tavoli vuoti da eliminare: ${emptyTables.length}`,
    `- Tavoli che resteranno: ${state.tables.length - emptyTables.length}`,
    "",
    "I tavoli rimanenti saranno rinumerati in ordine attuale (1, 2, 3, ...).",
  ];
  if (!window.confirm(messageLines.join("\n"))) return;

  if (emptyTables.length) {
    const emptyIds = new Set(emptyTables.map((t) => t.id));
    state.tables = state.tables.filter((t) => !emptyIds.has(t.id));
    for (const id of emptyIds) delete state.markerPositions[id];
  }
  state.tables.forEach((t, i) => {
    t.number = i + 1;
  });
  saveState();
  renderTables();
}

function createPdfDocument(orientation = "portrait") {
  if (!window.jspdf || !window.jspdf.jsPDF) return null;
  return new window.jspdf.jsPDF({ orientation, unit: "mm", format: "a4" });
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

function getSortedParticipants() {
  const rows = [];
  state.tables.forEach((table) => {
    table.guests.forEach((g) => {
      const cognome = (g.cognome || "").trim();
      const nome = (g.nome || "").trim();
      if (!cognome && !nome) return;
      rows.push({
        cognome,
        nome,
        tavoloNumero: table.number,
        tavoloNome: (table.customName || "").trim(),
      });
    });
  });
  rows.sort((a, b) => `${a.cognome} ${a.nome}`.toLowerCase().localeCompare(`${b.cognome} ${b.nome}`.toLowerCase(), "it"));
  return rows;
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
      if (curr) curr.count += 1;
      else targetMap.set(key, { label: note, count: 1 });
    }
  }
  const byCountThenLabel = (a, b) => b.count - a.count || a.label.localeCompare(b.label, "it");
  return {
    adulti: [...mapAdulti.values()].sort(byCountThenLabel),
    bambini: [...mapBambini.values()].sort(byCountThenLabel),
  };
}

els.eventName.addEventListener("input", () => {
  state.eventName = els.eventName.value;
  saveState();
});
els.addTableBtn.addEventListener("click", addTable);
els.sortTablesBtn.addEventListener("click", sortTablesAscending);
els.renumberTablesBtn.addEventListener("click", renumberTablesSmart);

els.exportBtn.addEventListener("click", () => {
  const blob = new Blob(
    [JSON.stringify({ ...state, exportedAt: new Date().toISOString() }, null, 2)],
    { type: "application/json" }
  );
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "evento-tavoli.json";
  a.click();
  URL.revokeObjectURL(a.href);
});

els.importInput.addEventListener("change", (e) => {
  const file = e.target.files && e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = () => {
    try {
      const data = JSON.parse(reader.result);
      state.eventName = String(data.eventName || "");
      state.floorPlanDataUrl = String(data.floorPlanDataUrl || "");
      state.tables = Array.isArray(data.tables) ? data.tables : [];
      state.tables.forEach(ensureTableGuestSlots);
      state.markerPositions =
        data.markerPositions && typeof data.markerPositions === "object" ? data.markerPositions : {};
      els.eventName.value = state.eventName;
      saveState();
      renderTables();
    } catch (_) {
      alert("File JSON non valido.");
    }
  };
  reader.readAsText(file);
  e.target.value = "";
});

els.exportGuestsBtn.addEventListener("click", () => {
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
  const lines = rows.map((row, idx) => {
    const fullName = `${row.cognome} ${row.nome}`.trim();
    const tavoloNome = row.tavoloNome ? ` - ${row.tavoloNome}` : "";
    return `${idx + 1}. ${fullName} - Tavolo ${row.tavoloNumero}${tavoloNome}`;
  });
  writePdfLines(pdf, lines, { startY: 20, margin: 10, lineHeight: 5 });
  pdf.save("lista-partecipanti-alfabetica.pdf");
});

els.exportKitchenBtn.addEventListener("click", () => {
  const pdf = createPdfDocument("portrait");
  if (!pdf) {
    alert("Libreria PDF non disponibile.");
    return;
  }
  let totAdulti = 0;
  let totBambini = 0;
  for (const table of state.tables) {
    const counts = getTableMenuCounts(table);
    totAdulti += counts.adulti;
    totBambini += counts.bambini;
  }
  const grouped = getGroupedDietNotesByMenu();
  const lines = [
    `Evento: ${state.eventName || "Senza nome"}`,
    `Data esportazione: ${new Date().toLocaleString("it-IT")}`,
    "",
    "TOTALE MENU",
    `- Menu adulti: ${totAdulti}`,
    `- Menu bambini: ${totBambini}`,
    "",
    "INTOLLERANZE/VARIAZIONI RAGGRUPPATE",
  ];
  lines.push("- Adulti:");
  if (!grouped.adulti.length) lines.push("- Nessuna specifica inserita");
  else grouped.adulti.forEach((g) => lines.push(`  - ${g.label}: ${g.count}`));
  lines.push("- Bambini:");
  if (!grouped.bambini.length) lines.push("- Nessuna specifica inserita");
  else grouped.bambini.forEach((g) => lines.push(`  - ${g.label}: ${g.count}`));
  pdf.setFontSize(14);
  pdf.text(`Riepilogo cucina - ${state.eventName || "Evento"}`, 10, 12);
  pdf.setFontSize(10);
  writePdfLines(pdf, lines, { startY: 20, margin: 10, lineHeight: 5 });
  pdf.save("riepilogo-cucina.pdf");
});

document.addEventListener("keydown", (e) => {
  if (e.key === "Escape" && document.querySelector(".guest-row__sheet:not([hidden])")) {
    closeAllGuestOptionSheets();
  }
});

loadState();
els.eventName.value = state.eventName;
renderTables();
=======
/**
 * Versione cliente: solo tavoli e ospiti (nessuna piantina).
 * Compatibile con lo stesso JSON della versione completa.
 */

const MAX_GUESTS = 9;
const STORAGE_KEY = "evento-tavoli-cliente-v1";

const state = {
  eventName: "",
  floorPlanDataUrl: "",
  tables: [],
  markerPositions: {},
};

const els = {
  eventName: document.getElementById("eventName"),
  addTableBtn: document.getElementById("addTableBtn"),
  sortTablesBtn: document.getElementById("sortTablesBtn"),
  renumberTablesBtn: document.getElementById("renumberTablesBtn"),
  tablesList: document.getElementById("tablesList"),
  tableCountLabel: document.getElementById("tableCountLabel"),
  exportGuestsBtn: document.getElementById("exportGuestsBtn"),
  exportKitchenBtn: document.getElementById("exportKitchenBtn"),
  exportBtn: document.getElementById("exportBtn"),
  importInput: document.getElementById("importInput"),
  tableCardTpl: document.getElementById("tableCardTpl"),
  guestRowTpl: document.getElementById("guestRowTpl"),
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
  while (normalized.length < MAX_GUESTS) normalized.push(createEmptyGuest());
  table.guests = normalized;
}

function uid() {
  return "t_" + Math.random().toString(36).slice(2, 11);
}

function saveState() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

function loadState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return;
    const data = JSON.parse(raw);
    state.eventName = String(data.eventName || "");
    state.floorPlanDataUrl = String(data.floorPlanDataUrl || "");
    state.tables = Array.isArray(data.tables) ? data.tables : [];
    state.markerPositions = data.markerPositions && typeof data.markerPositions === "object" ? data.markerPositions : {};
    state.tables.forEach(ensureTableGuestSlots);
  } catch (_) {}
}

function nextTableNumber() {
  const used = new Set();
  for (const t of state.tables) {
    const n = Number(t.number);
    if (!Number.isNaN(n) && n > 0) used.add(n);
  }
  let candidate = 1;
  while (used.has(candidate)) candidate += 1;
  return candidate;
}

function guestCount(table) {
  return table.guests.filter((g) => (g.cognome || "").trim() || (g.nome || "").trim()).length;
}

function getActiveGuests(table) {
  return table.guests.filter((g) => (g.cognome || "").trim() || (g.nome || "").trim());
}

function getTableMenuCounts(table) {
  const activeGuests = getActiveGuests(table);
  let adulti = 0;
  let bambini = 0;
  for (const guest of activeGuests) {
    if (guest.menu === "bambino") bambini += 1;
    else adulti += 1;
  }
  return { adulti, bambini };
}

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
    items.push({ menu, text, count, letter: menu === "bambino" ? "B" : "A" });
  }
  items.sort((a, b) => {
    if (a.menu !== b.menu) return a.menu === "adulto" ? -1 : 1;
    return a.text.localeCompare(b.text, "it");
  });
  return items.map((i) => `${i.count}${i.letter} ${i.text}`);
}

function getTableSummaryLine(table) {
  const c = getTableMenuCounts(table);
  const parts = [`${c.adulti}A+${c.bambini}B`];
  const allergies = getTableAllergenParts(table);
  if (allergies.length) parts.push(`(${allergies.join(" | ")})`);
  const note = (table.tableNote || "").trim();
  if (note) parts.push(`Note tavolo: ${note}`);
  return parts.join(" — ");
}

function setGuestRowVisualState(rowEl, hasIdentity) {
  rowEl.classList.toggle("guest-row--filled", hasIdentity);
  rowEl.classList.toggle("guest-row--empty", !hasIdentity);
}

function focusGuestField(card, rowIndex, colIndex) {
  const candidate = card.querySelector(
    `.guest-row[data-row-index="${rowIndex}"] [data-grid-col="${colIndex}"]`
  );
  if (!candidate || candidate.disabled) return false;
  candidate.focus();
  if (candidate.tagName === "INPUT") candidate.select();
  return true;
}

function moveGuestGridFocus(card, fromRow, fromCol, key) {
  const vectors = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
  const vector = vectors[key];
  if (!vector) return;
  const [dRow, dCol] = vector;
  let row = fromRow;
  let col = fromCol;
  for (let step = 0; step < MAX_GUESTS * 4; step += 1) {
    row += dRow;
    col += dCol;
    if (row < 0 || row >= MAX_GUESTS || col < 0 || col > 3) break;
    if (focusGuestField(card, row, col)) return;
  }
}

function moveGuestGridNextField(card, fromRow, fromCol) {
  const maxCols = 4;
  const start = fromRow * maxCols + fromCol;
  for (let i = start + 1; i < MAX_GUESTS * maxCols; i += 1) {
    const row = Math.floor(i / maxCols);
    const col = i % maxCols;
    if (focusGuestField(card, row, col)) return;
  }
}

function moveGuestGridPrevField(card, fromRow, fromCol) {
  const maxCols = 4;
  const start = fromRow * maxCols + fromCol;
  for (let i = start - 1; i >= 0; i -= 1) {
    const row = Math.floor(i / maxCols);
    const col = i % maxCols;
    if (focusGuestField(card, row, col)) return;
  }
}

function refreshTableLiveInfoById(tableId) {
  const table = state.tables.find((t) => t.id === tableId);
  if (!table) return;
  const card = els.tablesList.querySelector(`[data-table-id="${CSS.escape(tableId)}"]`);
  if (!card) return;
  const capEl = card.querySelector("[data-capacity]");
  const count = guestCount(table);
  capEl.textContent =
    count >= MAX_GUESTS
      ? `Posti occupati: ${count}/${MAX_GUESTS} (massimo raggiunto)`
      : `Posti occupati: ${count}/${MAX_GUESTS}`;
  capEl.classList.toggle("is-full", count >= MAX_GUESTS);
  card.querySelector("[data-summary]").textContent = `Resoconto: ${getTableSummaryLine(table)}`;
}

function renderTables() {
  els.tablesList.innerHTML = "";
  els.tableCountLabel.textContent = state.tables.length === 1 ? "1 tavolo" : `${state.tables.length} tavoli`;

  for (const table of state.tables) {
    ensureTableGuestSlots(table);
    const node = els.tableCardTpl.content.cloneNode(true);
    const card = node.querySelector(".table-card");
    card.dataset.tableId = table.id;
    node.querySelector(".table-title").textContent = `Tavolo ${table.number}`;

    const summaryEl = node.querySelector("[data-summary]");
    const customNameInput = node.querySelector(".table-custom-name");
    const tableNoteInput = node.querySelector(".table-note");
    const guestsEl = node.querySelector("[data-guests]");
    const cap = node.querySelector("[data-capacity]");

    customNameInput.value = table.customName || "";
    tableNoteInput.value = table.tableNote || "";
    customNameInput.addEventListener("input", () => {
      table.customName = customNameInput.value;
      saveState();
      summaryEl.textContent = `Resoconto: ${getTableSummaryLine(table)}`;
    });
    tableNoteInput.addEventListener("input", () => {
      table.tableNote = tableNoteInput.value;
      saveState();
      summaryEl.textContent = `Resoconto: ${getTableSummaryLine(table)}`;
    });

    table.guests.forEach((g, idx) => {
      const row = els.guestRowTpl.content.cloneNode(true);
      const rowEl = row.querySelector(".guest-row");
      rowEl.dataset.rowIndex = String(idx);
      const cognome = row.querySelector(".guest-cognome");
      const nome = row.querySelector(".guest-nome");
      const menu = row.querySelector(".guest-menu");
      const note = row.querySelector(".guest-note");
      cognome.dataset.gridCol = "0";
      nome.dataset.gridCol = "1";
      menu.dataset.gridCol = "2";
      note.dataset.gridCol = "3";
      cognome.value = g.cognome;
      nome.value = g.nome;
      menu.value = g.menu === "bambino" ? "bambino" : "adulto";
      note.value = g.note || "";

      const applyIdentityConstraints = () => {
        const hasIdentity = Boolean(cognome.value.trim() || nome.value.trim());
        menu.disabled = !hasIdentity;
        note.disabled = !hasIdentity;
        setGuestRowVisualState(rowEl, hasIdentity);
        if (!hasIdentity) {
          if (menu.value !== "adulto") menu.value = "adulto";
          if (note.value.trim()) note.value = "";
          g.menu = "adulto";
          g.note = "";
        }
      };
      applyIdentityConstraints();

      const saveGuestAndRefresh = () => {
        g.cognome = cognome.value;
        g.nome = nome.value;
        g.menu = menu.value;
        g.note = note.value;
        saveState();
        refreshTableLiveInfoById(table.id);
      };

      cognome.addEventListener("input", () => {
        applyIdentityConstraints();
        saveGuestAndRefresh();
      });
      nome.addEventListener("input", () => {
        applyIdentityConstraints();
        saveGuestAndRefresh();
      });
      menu.addEventListener("change", saveGuestAndRefresh);
      note.addEventListener("input", saveGuestAndRefresh);
      row.querySelector(".remove-guest").addEventListener("click", () => {
        table.guests[idx] = createEmptyGuest();
        saveState();
        renderTables();
      });

      [cognome, nome, menu, note].forEach((fieldEl) => {
        fieldEl.addEventListener("keydown", (e) => {
          const rowIndex = Number(fieldEl.closest(".guest-row")?.dataset.rowIndex);
          const colIndex = Number(fieldEl.dataset.gridCol);
          if (Number.isNaN(rowIndex) || Number.isNaN(colIndex)) return;

          if (e.key === "Enter") {
            e.preventDefault();
            if (e.shiftKey) moveGuestGridPrevField(card, rowIndex, colIndex);
            else moveGuestGridNextField(card, rowIndex, colIndex);
            return;
          }
          if (e.key === "ArrowUp" || e.key === "ArrowDown" || e.key === "ArrowLeft" || e.key === "ArrowRight") {
            e.preventDefault();
            moveGuestGridFocus(card, rowIndex, colIndex, e.key);
          }
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
    summaryEl.textContent = `Resoconto: ${getTableSummaryLine(table)}`;

    node.querySelector(".remove-table").addEventListener("click", () => {
      state.tables = state.tables.filter((t) => t.id !== table.id);
      delete state.markerPositions[table.id];
      saveState();
      renderTables();
    });
    els.tablesList.appendChild(node);
  }
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
  if (!state.markerPositions[id]) state.markerPositions[id] = { x: 50, y: 50 };
  saveState();
  renderTables();
}

function sortTablesAscending() {
  state.tables.sort((a, b) => Number(a.number) - Number(b.number));
  saveState();
  renderTables();
}

function renumberTablesSmart() {
  const emptyTables = state.tables.filter((t) => guestCount(t) === 0);
  const messageLines = [
    "Confermi la rinumerazione intelligente dei tavoli?",
    "",
    `- Tavoli attuali: ${state.tables.length}`,
    `- Tavoli vuoti da eliminare: ${emptyTables.length}`,
    `- Tavoli che resteranno: ${state.tables.length - emptyTables.length}`,
    "",
    "I tavoli rimanenti saranno rinumerati in ordine attuale (1, 2, 3, ...).",
  ];
  if (!window.confirm(messageLines.join("\n"))) return;

  if (emptyTables.length) {
    const emptyIds = new Set(emptyTables.map((t) => t.id));
    state.tables = state.tables.filter((t) => !emptyIds.has(t.id));
    for (const id of emptyIds) delete state.markerPositions[id];
  }
  state.tables.forEach((t, i) => {
    t.number = i + 1;
  });
  saveState();
  renderTables();
}

function createPdfDocument(orientation = "portrait") {
  if (!window.jspdf || !window.jspdf.jsPDF) return null;
  return new window.jspdf.jsPDF({ orientation, unit: "mm", format: "a4" });
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

function getSortedParticipants() {
  const rows = [];
  state.tables.forEach((table) => {
    table.guests.forEach((g) => {
      const cognome = (g.cognome || "").trim();
      const nome = (g.nome || "").trim();
      if (!cognome && !nome) return;
      rows.push({
        cognome,
        nome,
        tavoloNumero: table.number,
        tavoloNome: (table.customName || "").trim(),
      });
    });
  });
  rows.sort((a, b) => `${a.cognome} ${a.nome}`.toLowerCase().localeCompare(`${b.cognome} ${b.nome}`.toLowerCase(), "it"));
  return rows;
}

function getGroupedDietNotes() {
  const map = new Map();
  for (const table of state.tables) {
    for (const g of table.guests) {
      const note = String(g.note || "").trim();
      if (!note) continue;
      const key = note.toLocaleLowerCase("it");
      const curr = map.get(key);
      if (curr) curr.count += 1;
      else map.set(key, { label: note, count: 1 });
    }
  }
  return [...map.values()].sort((a, b) => b.count - a.count || a.label.localeCompare(b.label, "it"));
}

els.eventName.addEventListener("input", () => {
  state.eventName = els.eventName.value;
  saveState();
});
els.addTableBtn.addEventListener("click", addTable);
els.sortTablesBtn.addEventListener("click", sortTablesAscending);
els.renumberTablesBtn.addEventListener("click", renumberTablesSmart);

els.exportBtn.addEventListener("click", () => {
  const blob = new Blob(
    [JSON.stringify({ ...state, exportedAt: new Date().toISOString() }, null, 2)],
    { type: "application/json" }
  );
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "evento-tavoli.json";
  a.click();
  URL.revokeObjectURL(a.href);
});

els.importInput.addEventListener("change", (e) => {
  const file = e.target.files && e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = () => {
    try {
      const data = JSON.parse(reader.result);
      state.eventName = String(data.eventName || "");
      state.floorPlanDataUrl = String(data.floorPlanDataUrl || "");
      state.tables = Array.isArray(data.tables) ? data.tables : [];
      state.tables.forEach(ensureTableGuestSlots);
      state.markerPositions =
        data.markerPositions && typeof data.markerPositions === "object" ? data.markerPositions : {};
      els.eventName.value = state.eventName;
      saveState();
      renderTables();
    } catch (_) {
      alert("File JSON non valido.");
    }
  };
  reader.readAsText(file);
  e.target.value = "";
});

els.exportGuestsBtn.addEventListener("click", () => {
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
  const lines = rows.map((row, idx) => {
    const fullName = `${row.cognome} ${row.nome}`.trim();
    const tavoloNome = row.tavoloNome ? ` - ${row.tavoloNome}` : "";
    return `${idx + 1}. ${fullName} - Tavolo ${row.tavoloNumero}${tavoloNome}`;
  });
  writePdfLines(pdf, lines, { startY: 20, margin: 10, lineHeight: 5 });
  pdf.save("lista-partecipanti-alfabetica.pdf");
});

els.exportKitchenBtn.addEventListener("click", () => {
  const pdf = createPdfDocument("portrait");
  if (!pdf) {
    alert("Libreria PDF non disponibile.");
    return;
  }
  let totAdulti = 0;
  let totBambini = 0;
  for (const table of state.tables) {
    const counts = getTableMenuCounts(table);
    totAdulti += counts.adulti;
    totBambini += counts.bambini;
  }
  const grouped = getGroupedDietNotes();
  const lines = [
    `Evento: ${state.eventName || "Senza nome"}`,
    `Data esportazione: ${new Date().toLocaleString("it-IT")}`,
    "",
    "TOTALE MENU",
    `- Menu adulti: ${totAdulti}`,
    `- Menu bambini: ${totBambini}`,
    "",
    "INTOLLERANZE/VARIAZIONI RAGGRUPPATE",
  ];
  if (!grouped.length) lines.push("- Nessuna specifica inserita");
  else grouped.forEach((g) => lines.push(`- ${g.label}: ${g.count}`));
  pdf.setFontSize(14);
  pdf.text(`Riepilogo cucina - ${state.eventName || "Evento"}`, 10, 12);
  pdf.setFontSize(10);
  writePdfLines(pdf, lines, { startY: 20, margin: 10, lineHeight: 5 });
  pdf.save("riepilogo-cucina.pdf");
});

loadState();
els.eventName.value = state.eventName;
renderTables();
>>>>>>> 5f9a0327b864451f430a2105a0c1caf3fb1316bc
