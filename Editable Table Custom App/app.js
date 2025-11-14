/* MCC Lead Source Creator Drop Down Selections Admin
 * Full CRUD with persistent column + row order
 * Proper AppDB updates (no duplicates)
 * Add/Rename Column Modal + Confirm Delete Modal (iframe-safe)
 * Auto-seeds default columns and row when collection is empty
 * Excel-style multi-cell paste
 * Manual save only
 * Exports + Dark Mode + Toasts
 */

document.addEventListener("DOMContentLoaded", () => {
  // ============================================================
  // CONFIG
  // ============================================================
  const COLLECTION_NAME = "app_table_1";
  const endpoint = (path = "") =>
    `/domo/datastores/v1/collections/${encodeURIComponent(COLLECTION_NAME)}${path}`;

  // ============================================================
  // STATE
  // ============================================================
  const state = {
    data: [],
    columns: [],
    rowOrder: [],
    metaDocId: null,
    darkMode: localStorage.getItem("darkMode") === "true",
    currentPage: 1,
    rowsPerPage: 100,
    searchTerm: "",
    deleteRowIndex: null,
  };

  // ============================================================
  // ELEMENTS
  // ============================================================
  const elements = {
    tableBody: document.getElementById("table-body"),
    columnHeaders: document.getElementById("column-headers"),
    searchInput: document.getElementById("search-input"),
    addRowBtn: document.getElementById("add-row"),
    saveAllBtn: document.getElementById("save-all"),
    exportBtn: document.getElementById("export-btn"),
    dropdownMenu: document.querySelector(".dropdown-menu"),
    prevPageBtn: document.getElementById("prev-page"),
    nextPageBtn: document.getElementById("next-page"),
    currentPageEl: document.getElementById("current-page"),
    totalPagesEl: document.getElementById("total-pages"),
    themeToggle: document.querySelector(".theme-toggle"),
    toastContainer: document.getElementById("toast-container"),
    colMenu: document.getElementById("context-menu"),
    rowMenu: document.getElementById("row-context-menu"),
    // Column modal
    columnModal: document.getElementById("column-modal"),
    columnTitle: document.getElementById("column-modal-title"),
    columnInput: document.getElementById("column-name-input"),
    confirmBtn: document.getElementById("confirm-column-modal"),
    cancelBtn: document.getElementById("cancel-column-modal"),
    closeBtn: document.getElementById("close-column-modal"),
    // Delete modal
    deleteModal: document.getElementById("delete-popup"),
    deleteConfirm: document.getElementById("popup-confirm"),
    deleteCancel: document.getElementById("popup-cancel"),
  };

  // ============================================================
  // API HELPERS
  // ============================================================
  const listDocuments = async () => await domo.get(`${endpoint("/documents")}?limit=1000`);
  const createDocument = async (doc) => await domo.post(endpoint("/documents"), doc);
  const updateDocument = async (id, doc) =>
    await domo.put(endpoint(`/documents/${encodeURIComponent(id)}`), doc);
  const deleteDocumentApi = async (id) =>
    await domo.delete(endpoint(`/documents/${encodeURIComponent(id)}`));

  // ============================================================
  // LOAD DATA (auto seed if empty)
  // ============================================================
  async function loadData() {
    try {
      const docs = await listDocuments();

      if (!docs || docs.length === 0) {
        console.log("Collection empty — seeding starter data...");
        state.columns = ["a", "b", "c", "d"];
        const seedRow = { a: "test", b: "", c: "", d: "", _row_id: "row-1" };
        const createdRow = await createDocument({ content: seedRow });
        seedRow.id = createdRow.id;
        state.data = [seedRow];
        state.rowOrder = ["row-1"];

        const meta = await createDocument({
          content: { _meta_order: { columns: state.columns, rows: state.rowOrder } },
        });
        state.metaDocId = meta.id;

        renderColumnHeaders();
        renderTable();
        showToast("Initialized collection with sample row and columns.", "info");
        return;
      }

      const metaDoc = docs.find((d) => d.content && d.content._meta_order);
      const dataDocs = docs.filter((d) => !d.content._meta_order);

      if (metaDoc) {
        state.metaDocId = metaDoc.id;
        state.columns = metaDoc.content._meta_order.columns || [];
        state.rowOrder = metaDoc.content._meta_order.rows || [];
      } else {
        const keys = new Set();
        dataDocs.forEach((d) => Object.keys(d.content).forEach((k) => keys.add(k)));
        state.columns = [...keys].filter((k) => k !== "id");
        state.rowOrder = dataDocs.map((_, i) => `row-${i + 1}`);
      }

      const rows = state.rowOrder.map((id) => {
        const match = dataDocs.find((d) => d.content._row_id === id);
        if (match) return { id: match.id, ...match.content };
        return { _row_id: id, ...Object.fromEntries(state.columns.map((c) => [c, ""])) };
      });

      state.data = rows;
      renderColumnHeaders();
      renderTable();
      showToast(`Loaded ${rows.length} rows`, "success");
    } catch (error) {
      console.error("Error loading collection:", error);
      showToast("Failed to load data", "error");
    }
  }

  // ============================================================
  // SAVE FUNCTIONS
  // ============================================================
  async function saveRow(row) {
    try {
      if (row.id) await updateDocument(row.id, { content: row });
      else {
        const created = await createDocument({ content: row });
        row.id = created.id;
      }
    } catch (error) {
      console.error("Row save failed:", error);
    }
  }

  async function saveMeta() {
    try {
      const metaContent = {
        _meta_order: {
          columns: state.columns,
          rows: state.data.map((r) => r._row_id),
        },
      };
      if (state.metaDocId) await updateDocument(state.metaDocId, { content: metaContent });
      else {
        const created = await createDocument({ content: metaContent });
        state.metaDocId = created.id;
      }
    } catch (error) {
      console.error("Meta save failed:", error);
    }
  }

  async function saveAllRows() {
    if (!state.data.length) return showToast("No data to save", "info");
    showToast("Saving...", "info");
    state.data.forEach((r, i) => (!r._row_id ? (r._row_id = `row-${i + 1}`) : null));
    for (const row of state.data) await saveRow(row);
    await saveMeta();
    showToast("All data and order updated successfully.", "success");
  }

  // ============================================================
  // TABLE RENDERING (includes Excel-style paste)
  // ============================================================
  function renderColumnHeaders() {
    const header = elements.columnHeaders;
    while (header.children.length > 1) header.removeChild(header.lastChild);
    state.columns.forEach((col) => {
      const th = document.createElement("th");
      th.textContent = col;
      th.dataset.column = col;
      header.appendChild(th);
    });
  }

  function renderTable() {
    const data = getFilteredData();
    const start = (state.currentPage - 1) * state.rowsPerPage;
    const subset = data.slice(start, start + state.rowsPerPage);
    elements.tableBody.innerHTML = "";

    subset.forEach((rowData, rowIndex) => {
      const tr = document.createElement("tr");
      tr.dataset.rowId = rowData._row_id;

      // Delete button
      const action = document.createElement("td");
      const del = document.createElement("button");
      del.className = "row-action delete-row";
      del.textContent = "🗑️";
      del.addEventListener("click", () => openDeleteModal(rowIndex));
      action.appendChild(del);
      tr.appendChild(action);

      state.columns.forEach((col) => {
        const td = document.createElement("td");
        td.contentEditable = true;
        td.dataset.column = col;
        td.textContent = rowData[col] || "";

        // Save changes on blur
        td.addEventListener("blur", (e) => {
          const idx = state.data.indexOf(rowData);
          if (idx >= 0) state.data[idx][col] = e.target.textContent.trim();
        });

        // 🟢 Excel-style paste handler
        td.addEventListener("paste", (e) => {
          e.preventDefault();
          const clipboardData = e.clipboardData.getData("text/plain");
          if (!clipboardData) return;

          const rows = clipboardData
            .split(/\r?\n/)
            .map((r) => r.split("\t"))
            .filter((r) => r.length && r.some((v) => v.trim() !== ""));

          const startRow = rowIndex;
          const startCol = state.columns.indexOf(col);

          rows.forEach((rValues, rOffset) => {
            const targetRowIndex = startRow + rOffset;
            if (targetRowIndex >= state.data.length) {
              const newRow = {};
              state.columns.forEach((c) => (newRow[c] = ""));
              newRow._row_id = `row-${Date.now()}-${rOffset}`;
              state.data.push(newRow);
              state.rowOrder.push(newRow._row_id);
            }

            const targetRow = state.data[targetRowIndex];
            rValues.forEach((val, cOffset) => {
              const targetColIndex = startCol + cOffset;
              const targetCol = state.columns[targetColIndex];
              if (targetCol) targetRow[targetCol] = val.trim();
            });
          });

          renderTable();
        });

        tr.appendChild(td);
      });

      elements.tableBody.appendChild(tr);
    });

    updatePagination();
  }

  function getFilteredData() {
    const t = state.searchTerm.toLowerCase();
    return !t
      ? state.data
      : state.data.filter((r) =>
          Object.values(r).some((v) => v && String(v).toLowerCase().includes(t))
        );
  }

  function updatePagination() {
    const total = Math.ceil(getFilteredData().length / state.rowsPerPage);
    elements.currentPageEl.textContent = state.currentPage;
    elements.totalPagesEl.textContent = total;
    elements.prevPageBtn.disabled = state.currentPage <= 1;
    elements.nextPageBtn.disabled = state.currentPage >= total;
  }

  // ============================================================
  // DELETE MODAL
  // ============================================================
  function openDeleteModal(index) {
    state.deleteRowIndex = index;
    elements.deleteModal.style.display = "flex";
  }

  function closeDeleteModal() {
    elements.deleteModal.style.display = "none";
    state.deleteRowIndex = null;
  }

  elements.deleteCancel.addEventListener("click", closeDeleteModal);
  elements.deleteConfirm.addEventListener("click", async () => {
    if (state.deleteRowIndex !== null) await deleteRowAt(state.deleteRowIndex);
    closeDeleteModal();
  });

  async function deleteRowAt(index) {
    if (index < 0 || index >= state.data.length) return;
    const row = state.data[index];
    state.data.splice(index, 1);
    state.rowOrder = state.rowOrder.filter((r) => r !== row._row_id);
    renderTable();

    if (row.id) {
      try {
        await deleteDocumentApi(row.id);
        showToast("Row deleted from collection", "success");
      } catch (error) {
        console.error("Error deleting row:", error);
        showToast("Failed to delete row", "error");
      }
    }

    await saveMeta();
  }

  // ============================================================
  // ROW CONTEXT MENU
  // ============================================================
  let clickedRowIndex = null;
  elements.tableBody.addEventListener("contextmenu", (e) => {
    const tr = e.target.closest("tr");
    if (!tr) return;
    e.preventDefault();
    clickedRowIndex = Array.from(elements.tableBody.children).indexOf(tr);
    elements.rowMenu.style.left = `${e.pageX}px`;
    elements.rowMenu.style.top = `${e.pageY}px`;
    elements.rowMenu.style.display = "block";
  });

  document.getElementById("insert-row-above").addEventListener("click", () => {
    insertRowAt(clickedRowIndex);
    elements.rowMenu.style.display = "none";
  });
  document.getElementById("insert-row-below").addEventListener("click", () => {
    insertRowAt(clickedRowIndex + 1);
    elements.rowMenu.style.display = "none";
  });
  document.getElementById("delete-this-row").addEventListener("click", () => {
    openDeleteModal(clickedRowIndex);
    elements.rowMenu.style.display = "none";
  });

  function insertRowAt(i) {
    const newRow = Object.fromEntries(state.columns.map((c) => [c, ""]));
    newRow._row_id = `row-${Date.now()}`;
    state.data.splice(i, 0, newRow);
    state.rowOrder.splice(i, 0, newRow._row_id);
    renderTable();
    showToast("Row inserted");
  }

  // ============================================================
  // COLUMN CONTEXT MENU + MODAL
  // ============================================================
  let clickedColIndex = null;
  elements.columnHeaders.addEventListener("contextmenu", (e) => {
    e.preventDefault();
    const th = e.target.closest("th");
    if (!th || !th.dataset.column) return;
    clickedColIndex = Array.from(elements.columnHeaders.children).indexOf(th) - 1;
    elements.colMenu.style.left = `${e.pageX}px`;
    elements.colMenu.style.top = `${e.pageY}px`;
    elements.colMenu.style.display = "block";
  });

  document.getElementById("insert-left").addEventListener("click", () => {
    openColumnModal("add", clickedColIndex);
    elements.colMenu.style.display = "none";
  });
  document.getElementById("insert-right").addEventListener("click", () => {
    openColumnModal("add", clickedColIndex + 1);
    elements.colMenu.style.display = "none";
  });
  document.getElementById("rename-column").addEventListener("click", () => {
    openColumnModal("rename", clickedColIndex);
    elements.colMenu.style.display = "none";
  });
  document.getElementById("delete-column").addEventListener("click", () => {
    deleteColumn(clickedColIndex);
    elements.colMenu.style.display = "none";
  });

  // ============================================================
  // COLUMN MODAL
  // ============================================================
  let modalAction = null;
  let modalColumnIndex = null;

  function openColumnModal(action, index = null) {
    modalAction = action;
    modalColumnIndex = index;
    elements.columnInput.value = "";
    elements.columnTitle.textContent =
      action === "rename" ? "Rename Column" : "Add Column";
    elements.columnModal.classList.add("show");
    elements.columnInput.focus();
  }

  function closeColumnModal() {
    elements.columnModal.classList.remove("show");
    modalAction = null;
    modalColumnIndex = null;
  }

  elements.cancelBtn.addEventListener("click", closeColumnModal);
  elements.closeBtn.addEventListener("click", closeColumnModal);
  elements.columnModal.addEventListener("click", (e) => {
    if (e.target === elements.columnModal) closeColumnModal();
  });

  elements.confirmBtn.addEventListener("click", () => {
    const name = elements.columnInput.value.trim();
    if (!name) return showToast("Please enter a valid column name", "error");
    if (modalAction === "add") handleAddColumn(name, modalColumnIndex);
    else if (modalAction === "rename") handleRenameColumn(name, modalColumnIndex);
    closeColumnModal();
  });

  function handleAddColumn(name, index) {
    if (state.columns.includes(name)) return showToast("Column already exists", "error");
    state.columns.splice(index, 0, name);
    state.data.forEach((r) => (r[name] = ""));
    renderColumnHeaders();
    renderTable();
    showToast(`Column "${name}" added`);
  }

  function handleRenameColumn(newName, index) {
    const oldName = state.columns[index];
    if (newName === oldName) return;
    if (state.columns.includes(newName)) return showToast("Column already exists", "error");
    state.columns[index] = newName;
    state.data.forEach((r) => {
      r[newName] = r[oldName];
      delete r[oldName];
    });
    renderColumnHeaders();
    renderTable();
    showToast(`Renamed "${oldName}" → "${newName}"`);
  }

  function deleteColumn(i) {
    const col = state.columns[i];
    if (!confirm(`Delete column "${col}"?`)) return;
    state.columns.splice(i, 1);
    state.data.forEach((r) => delete r[col]);
    renderColumnHeaders();
    renderTable();
    showToast(`Column "${col}" deleted`);
  }

  // ============================================================
  // EXPORT UTILITIES
  // ============================================================
  function exportData(format) {
    const data = getFilteredData();
    if (!data.length) return showToast("No data to export", "error");

    switch (format) {
      case "csv":
        return exportCSV(data);
      case "json":
        return exportJSON(data);
      case "excel":
        return exportExcel(data);
      case "pdf":
        return exportPDF(data);
    }
  }

  function exportCSV(data) {
    const headers = state.columns.join(",");
    const rows = data.map((r) =>
      state.columns.map((c) => `"${r[c] || ""}"`).join(",")
    );
    const blob = new Blob([headers + "\n" + rows.join("\n")], { type: "text/csv" });
    const url = URL.createObjectURL(blob);
    downloadFile(url, "export.csv");
  }

  function exportJSON(data) {
    const blob = new Blob([JSON.stringify(data, null, 2)], {
      type: "application/json",
    });
    const url = URL.createObjectURL(blob);
    downloadFile(url, "export.json");
  }

  function exportExcel(data) {
    if (typeof XLSX === "undefined") return showToast("SheetJS not loaded", "error");
    const ws = XLSX.utils.json_to_sheet(data, { header: state.columns });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Data");
    XLSX.writeFile(wb, "export.xlsx");
    showToast("Excel exported", "success");
  }

  function exportPDF(data) {
    if (typeof jsPDF === "undefined" || typeof autoTable === "undefined")
      return exportCSV(data);
    const doc = new jsPDF();
    const tableData = data.map((r) => state.columns.map((c) => r[c] || ""));
    doc.autoTable({ head: [state.columns], body: tableData });
    doc.save("export.pdf");
    showToast("PDF exported", "success");
  }

  function downloadFile(url, filename) {
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 100);
  }

  // ============================================================
  // HELPERS & EVENTS
  // ============================================================
  function showToast(msg, type = "success") {
    const t = document.createElement("div");
    t.className = `toast toast-${type}`;
    t.textContent = msg;
    elements.toastContainer.appendChild(t);
    t.classList.add("show");
    setTimeout(() => t.remove(), 3000);
  }

  document.addEventListener("click", () => {
    elements.colMenu.style.display = "none";
    elements.rowMenu.style.display = "none";
  });

  elements.addRowBtn.addEventListener("click", () => insertRowAt(state.data.length));
  elements.saveAllBtn.addEventListener("click", saveAllRows);

  elements.exportBtn.addEventListener("click", () => {
    elements.dropdownMenu.style.display =
      elements.dropdownMenu.style.display === "block" ? "none" : "block";
  });

  elements.dropdownMenu.querySelectorAll(".dropdown-item").forEach((item) => {
    item.addEventListener("click", () => {
      exportData(item.dataset.format);
      elements.dropdownMenu.style.display = "none";
    });
  });

  elements.searchInput.addEventListener("input", (e) => {
    state.searchTerm = e.target.value.trim();
    state.currentPage = 1;
    renderTable();
  });

  elements.prevPageBtn.addEventListener("click", () => {
    if (state.currentPage > 1) {
      state.currentPage--;
      renderTable();
    }
  });

  elements.nextPageBtn.addEventListener("click", () => {
    const total = Math.ceil(getFilteredData().length / state.rowsPerPage);
    if (state.currentPage < total) {
      state.currentPage++;
      renderTable();
    }
  });

  elements.themeToggle.addEventListener("click", () => {
    state.darkMode = !state.darkMode;
    document.body.classList.toggle("dark-mode", state.darkMode);
    localStorage.setItem("darkMode", state.darkMode);
  });

  // ============================================================
  // INIT
  // ============================================================
  loadData();
});
