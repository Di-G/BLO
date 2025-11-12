(() => {
	const MAX_FORMS = 10000;
	const STATUS_OPTIONS = [
		"",
		"Shifted",
		"Married",
		"Divorced",
		"Dead",
		"No-one by this name",
		"House Locked"
	];

	const els = {
		numInput: document.getElementById("numForms"),
		generateBtn: document.getElementById("generateBtn"),
		clearAllBtn: document.getElementById("clearAllBtn"),
		settingsBtn: document.getElementById("settingsBtn"),
		settingsModal: document.getElementById("settingsModal"),
		settingsClose: document.getElementById("settingsClose"),
		themeToggle: document.getElementById("themeToggle"),
		installBtn: document.getElementById("installBtn"),
		controls: document.querySelector(".controls"),
		warning: document.getElementById("warning"),
		listHeader: document.getElementById("listHeader"),
		list: document.getElementById("list"),
		hdrStatus1: document.getElementById("hdrStatus1"),
		hdrStatus2: document.getElementById("hdrStatus2"),
		societyName: document.getElementById("societyName"),
		societyStart: document.getElementById("societyStart"),
		societyEnd: document.getElementById("societyEnd"),
		addSocietyBtn: document.getElementById("addSocietyBtn"),
		societyError: document.getElementById("societyError"),
		societyList: document.getElementById("societyList"),
		firstRunControls: document.getElementById("firstRunControls"),
		settingsNumInput: document.getElementById("settingsNumForms"),
		settingsSaveFormsBtn: document.getElementById("settingsSaveForms"),
		settingsFormsWarning: document.getElementById("settingsFormsWarning"),
		confirmModal: document.getElementById("confirmModal"),
		confirmMessage: document.getElementById("confirmMessage"),
		confirmOk: document.getElementById("confirmOk"),
		confirmCancel: document.getElementById("confirmCancel"),
		exportToggle: document.getElementById("exportToggle"),
		exportOptions: document.getElementById("exportOptions"),
		exportPdfBtn: document.getElementById("exportPdfBtn"),
		exportExcelBtn: document.getElementById("exportExcelBtn"),
		importExcelBtn: document.getElementById("importExcelBtn"),
		importExcelInput: document.getElementById("importExcelInput"),
		importExportMessage: document.getElementById("importExportMessage"),
		exportPdfPanel: document.getElementById("exportPdfPanel"),
		exportPdfModeAll: document.getElementById("exportPdfModeAll"),
		exportPdfModeSociety: document.getElementById("exportPdfModeSociety"),
		exportPdfModeRange: document.getElementById("exportPdfModeRange"),
		exportPdfSocietyWrap: document.getElementById("exportPdfSocietyWrap"),
		exportPdfSocieties: document.getElementById("exportPdfSocieties"),
		exportPdfRangeWrap: document.getElementById("exportPdfRangeWrap"),
		exportPdfFrom: document.getElementById("exportPdfFrom"),
		exportPdfTo: document.getElementById("exportPdfTo"),
		exportPdfConfirm: document.getElementById("exportPdfConfirm"),
		exportPdfCancel: document.getElementById("exportPdfCancel")
	};

	const storageKey = (id) => `blo_forms_tracker_v1_row_${id}`;
	const storageCountKey = "blo_forms_tracker_v1_count";
	const storageThemeKey = "blo_forms_tracker_v1_theme";
	const storageGroupsKey = "blo_forms_tracker_v1_groups";
	const storageGroupOpenKey = (id) => `blo_forms_tracker_v1_group_open_${id}`;

	let deferredInstallPrompt = null;
	let installMessageTimeout = null;

	function toggleInstallButton(show) {
		if (!els.installBtn) return;
		if (show) {
			els.installBtn.classList.remove("hidden");
			document.body.classList.add("has-install-banner");
		} else {
			els.installBtn.classList.add("hidden");
			document.body.classList.remove("has-install-banner");
		}
	}

	function announceInstallSuccess() {
		if (installMessageTimeout) {
			clearTimeout(installMessageTimeout);
			installMessageTimeout = null;
		}
		showWarning("App installed successfully! You can launch it from your home screen.");
		installMessageTimeout = setTimeout(() => {
			showWarning("");
			installMessageTimeout = null;
		}, 5000);
	}

	function coerceBoolean(value) {
		if (typeof value === "boolean") return value;
		if (typeof value === "number") return value !== 0;
		if (typeof value === "string") {
			const normalized = value.trim().toLowerCase();
			if (!normalized) return false;
			if (["true", "yes", "y", "1", "✓", "✔"].includes(normalized)) return true;
			if (["false", "no", "n", "0"].includes(normalized)) return false;
		}
		return false;
	}

	function coerceString(value) {
		return typeof value === "string" ? value : value == null ? "" : String(value);
	}

	function refreshExportSocietyOptions() {
		if (!els.exportPdfSocieties) return;
		const select = els.exportPdfSocieties;
		const previousSelection = new Set(Array.from(select.selectedOptions).map((opt) => opt.value));
		select.textContent = "";
		const groups = getGroupsSorted();
		if (!groups.length) {
			const option = document.createElement("option");
			option.value = "";
			option.textContent = "No societies available";
			option.disabled = true;
			select.appendChild(option);
			select.disabled = true;
			return;
		}
		select.disabled = false;
		for (const group of groups) {
			const option = document.createElement("option");
			option.value = group.id;
			option.textContent = `${group.name} (${group.start}–${group.end})`;
			if (previousSelection.has(group.id)) {
				option.selected = true;
			}
			select.appendChild(option);
		}
	}

	function getSelectedPdfMode() {
		if (!els.exportPdfPanel) return "all";
		const checked = els.exportPdfPanel.querySelector('input[name="exportPdfMode"]:checked');
		return checked ? checked.value : "all";
	}

	function syncPdfExportModeUI() {
		const mode = getSelectedPdfMode();
		if (els.exportPdfSocietyWrap) {
			if (mode === "societies") {
				els.exportPdfSocietyWrap.classList.remove("hidden");
			} else {
				els.exportPdfSocietyWrap.classList.add("hidden");
			}
		}
		if (els.exportPdfRangeWrap) {
			if (mode === "range") {
				els.exportPdfRangeWrap.classList.remove("hidden");
			} else {
				els.exportPdfRangeWrap.classList.add("hidden");
			}
		}
	}

	function preparePdfExportPanel() {
		if (!els.exportPdfPanel) return;
		refreshExportSocietyOptions();
		if (els.exportPdfModeAll) {
			els.exportPdfModeAll.checked = true;
		}
		if (els.exportPdfModeSociety) {
			els.exportPdfModeSociety.checked = false;
		}
		if (els.exportPdfModeRange) {
			els.exportPdfModeRange.checked = false;
		}
		const count = getSavedCount();
		if (els.exportPdfFrom) {
			els.exportPdfFrom.value = count ? 1 : "";
		}
		if (els.exportPdfTo) {
			els.exportPdfTo.value = count ? count : "";
		}
		syncPdfExportModeUI();
	}

	function openPdfPanel() {
		if (!els.exportPdfPanel) return;
		preparePdfExportPanel();
		els.exportPdfPanel.classList.remove("hidden");
	}

	function closePdfPanel() {
		if (!els.exportPdfPanel) return;
		els.exportPdfPanel.classList.add("hidden");
	}

	function collectPdfFilter() {
		const mode = getSelectedPdfMode();
		if (mode === "societies") {
			if (!els.exportPdfSocieties || els.exportPdfSocieties.disabled) {
				showImportExportMessage("No societies available to export.", "warn");
				return null;
			}
			const selected = Array.from(els.exportPdfSocieties.selectedOptions || [])
				.map((opt) => opt.value)
				.filter(Boolean);
			if (!selected.length) {
				showImportExportMessage("Please select at least one society.", "warn");
				return null;
			}
			return { mode: "societies", societyIds: selected };
		}
		if (mode === "range") {
			const start = Number.parseInt((els.exportPdfFrom ? els.exportPdfFrom.value : "").trim(), 10);
			const end = Number.parseInt((els.exportPdfTo ? els.exportPdfTo.value : "").trim(), 10);
			if (!Number.isFinite(start) || !Number.isFinite(end)) {
				showImportExportMessage("Enter both start and end numbers.", "warn");
				return null;
			}
			if (start < 1 || end < 1 || start > MAX_FORMS || end > MAX_FORMS) {
				showImportExportMessage(`Numbers must be between 1 and ${MAX_FORMS}.`, "warn");
				return null;
			}
			if (start > end) {
				showImportExportMessage("Start number cannot be greater than end number.", "warn");
				return null;
			}
			return { mode: "range", start, end };
		}
		return { mode: "all" };
	}

	function buildGroupedRows(rows, groups) {
		const rowsByGroup = new Map();
		for (const group of groups) {
			rowsByGroup.set(group.id, []);
		}
		const ungroupedRows = [];
		for (const row of rows) {
			if (row.groupId && rowsByGroup.has(row.groupId)) {
				rowsByGroup.get(row.groupId).push(row);
			} else {
				ungroupedRows.push(row);
			}
		}
		const groupedRows = [];
		for (const group of groups) {
			const groupRows = rowsByGroup.get(group.id) || [];
			if (groupRows.length > 0) {
				groupedRows.push({ group, rows: groupRows });
			}
		}
		return { groupedRows, ungroupedRows };
	}

	function showWarning(message) {
		els.warning.textContent = message || "";
	}

	// Custom confirmation dialog
	function showConfirm(message, isDanger = false) {
		return new Promise((resolve) => {
			if (!els.confirmModal || !els.confirmMessage || !els.confirmOk || !els.confirmCancel) {
				resolve(false);
				return;
			}
			els.confirmMessage.textContent = message;
			els.confirmOk.classList.remove("danger");
			if (isDanger) {
				els.confirmOk.classList.add("danger");
			}
			els.confirmModal.setAttribute("aria-hidden", "false");
			document.body.classList.add("modal-open");

			const cleanup = () => {
				els.confirmModal.setAttribute("aria-hidden", "true");
				document.body.classList.remove("modal-open");
				els.confirmOk.removeEventListener("click", handleOk);
				els.confirmCancel.removeEventListener("click", handleCancel);
				if (els.confirmModal) {
					els.confirmModal.removeEventListener("click", handleBackdrop);
				}
				window.removeEventListener("keydown", handleEscape);
			};

			const handleOk = () => {
				cleanup();
				resolve(true);
			};

			const handleCancel = () => {
				cleanup();
				resolve(false);
			};

			const handleBackdrop = (e) => {
				if (e.target === els.confirmModal) {
					handleCancel();
				}
			};

			const handleEscape = (e) => {
				if (e.key === "Escape" && els.confirmModal.getAttribute("aria-hidden") === "false") {
					handleCancel();
				}
			};

			els.confirmOk.addEventListener("click", handleOk);
			els.confirmCancel.addEventListener("click", handleCancel);
			if (els.confirmModal) {
				els.confirmModal.addEventListener("click", handleBackdrop);
			}
			window.addEventListener("keydown", handleEscape);
			els.confirmOk.focus();
		});
	}

	function showSettingsMessage(message, type = "error") {
		if (!els.settingsFormsWarning) return;
		const el = els.settingsFormsWarning;
		el.textContent = message || "";
		el.classList.remove("message-success", "message-warn");
		if (!message) return;
		if (type === "success") {
			el.classList.add("message-success");
		} else if (type === "warn") {
			el.classList.add("message-warn");
		}
	}

	function showImportExportMessage(message, type = "error") {
		if (!els.importExportMessage) return;
		const el = els.importExportMessage;
		el.textContent = message || "";
		el.classList.remove("message-success", "message-warn");
		if (!message) return;
		if (type === "success") {
			el.classList.add("message-success");
		} else if (type === "warn") {
			el.classList.add("message-warn");
		}
	}

	function syncFormInputs(value) {
		const strValue = Number.isFinite(value) && value > 0 ? String(value) : "";
		if (els.numInput) els.numInput.value = strValue;
		if (els.settingsNumInput) els.settingsNumInput.value = strValue;
	}

	function setFirstRunMode(isFirstRun) {
		if (!els.firstRunControls) return;
		if (isFirstRun) {
			els.firstRunControls.classList.remove("hidden");
			if (els.controls) {
				els.controls.classList.remove("hidden");
			}
		} else {
			els.firstRunControls.classList.add("hidden");
			if (els.controls) {
				els.controls.classList.add("hidden");
			}
		}
	}

	function cleanupRowsAbove(limit, previousCount) {
		if (!Number.isFinite(limit) || !Number.isFinite(previousCount)) return;
		if (previousCount <= limit) return;
		for (let id = limit + 1; id <= previousCount; id++) {
			localStorage.removeItem(storageKey(id));
		}
	}

	async function processCountSubmission(rawValue, context) {
		const rawString = rawValue == null ? "" : String(rawValue);
		const trimmed = rawString.trim();
		const parsed = Number.parseInt(trimmed, 10);
		const isSettings = context === "settings";

		if (isSettings) {
			showSettingsMessage("");
		} else {
			showWarning("");
		}

		if (!Number.isFinite(parsed) || parsed < 1) {
			if (isSettings) {
				showSettingsMessage("Please enter a valid number between 1 and 10,000.");
			} else {
				showWarning("Please enter a valid number between 1 and 10,000.");
			}
			return false;
		}
		if (parsed > MAX_FORMS) {
			if (isSettings) {
				showSettingsMessage("More than 10,000 is not allowed.");
			} else {
				showWarning("More than 10,000 is not allowed.");
			}
			return false;
		}

		const currentCount = getSavedCount();
		if (Number.isFinite(currentCount) && currentCount > 0) {
			if (parsed === currentCount) {
				if (isSettings) {
					showSettingsMessage(`Number of forms is already ${parsed}.`, "warn");
				} else {
					showWarning("Number of forms is already set to that value.");
				}
				return false;
			}
			if (parsed < currentCount) {
				const confirmMessage = `You currently have ${currentCount} forms. Reducing this to ${parsed} will permanently remove data for forms ${parsed + 1} to ${currentCount}. Continue?`;
				const proceed = await showConfirm(confirmMessage, true);
				if (!proceed) {
					if (isSettings) {
						showSettingsMessage("No changes made.", "warn");
					} else {
						showWarning("No changes made.");
					}
					return false;
				}
				cleanupRowsAbove(parsed, currentCount);
			}
		}

		saveCount(parsed);
		syncFormInputs(parsed);
		renderList(parsed);
		setFirstRunMode(false);
		if (isSettings) {
			showSettingsMessage(`Number of forms updated to ${parsed}.`, "success");
		}
		return true;
	}

	function renderSocietyManager() {
		if (!els.societyList) return;
		const groups = getGroupsSorted();
		els.societyList.textContent = "";
		if (groups.length === 0) {
			const placeholder = document.createElement("li");
			placeholder.className = "society-placeholder";
			placeholder.textContent = "No societies yet. Add one above.";
			els.societyList.appendChild(placeholder);
			refreshExportSocietyOptions();
			return;
		}
		const totalForms = getSavedCount();
		for (const group of groups) {
			const item = document.createElement("li");
			item.className = "society-item";
			const info = document.createElement("div");
			info.className = "society-item-info";
			const name = document.createElement("span");
			name.className = "society-item-name";
			name.textContent = group.name;
			const range = document.createElement("span");
			range.className = "society-item-range";
			let countLabel = "";
			if (Number.isFinite(totalForms)) {
				const start = Math.max(1, group.start);
				const end = Math.min(totalForms, group.end);
				const count = Math.max(0, end - start + 1);
				countLabel = count > 0 ? ` · ${count} form${count === 1 ? "" : "s"}` : " · 0 forms in current list";
			}
			range.textContent = `Numbers ${group.start}–${group.end}${countLabel}`;
			info.appendChild(name);
			info.appendChild(range);
			const removeBtn = document.createElement("button");
			removeBtn.type = "button";
			removeBtn.className = "secondary";
			removeBtn.textContent = "Remove";
			removeBtn.addEventListener("click", async () => {
				const proceed = await showConfirm(`Remove society "${group.name}"?`);
				if (!proceed) return;
				const remaining = loadGroups().filter((g) => g.id !== group.id);
				saveGroups(remaining);
				renderSocietyManager();
				const count = getSavedCount();
				if (count) renderList(count);
			});
			item.appendChild(info);
			item.appendChild(removeBtn);
			els.societyList.appendChild(item);
		}
		refreshExportSocietyOptions();
	}

	function showSocietyError(message) {
		if (!els.societyError) return;
		els.societyError.textContent = message || "";
	}

	function handleAddSociety() {
		if (!els.societyName || !els.societyStart || !els.societyEnd) return;
		const name = els.societyName.value.trim();
		const start = Number.parseInt((els.societyStart.value || "").trim(), 10);
		const end = Number.parseInt((els.societyEnd.value || "").trim(), 10);
		if (!name) {
			showSocietyError("Please enter a society name.");
			return;
		}
		if (!Number.isFinite(start) || !Number.isFinite(end)) {
			showSocietyError("Please enter valid start and end numbers.");
			return;
		}
		if (start < 1 || end < 1 || start > MAX_FORMS || end > MAX_FORMS) {
			showSocietyError(`Numbers must be between 1 and ${MAX_FORMS}.`);
			return;
		}
		if (start > end) {
			showSocietyError("Start number cannot be greater than end number.");
			return;
		}
		const groups = loadGroups();
		const id = `soc_${Date.now()}_${Math.floor(Math.random() * 1000)}`;
		groups.push({ id, name, start, end });
		saveGroups(groups);
		setGroupOpenState(id, false);
		showSocietyError("");
		els.societyName.value = "";
		els.societyStart.value = "";
		els.societyEnd.value = "";
		renderSocietyManager();
		const count = getSavedCount();
		if (count) renderList(count);
	}

	function getSavedCount() {
		const raw = localStorage.getItem(storageCountKey);
		const parsed = Number.parseInt(raw || "", 10);
		return Number.isFinite(parsed) && parsed > 0 ? parsed : null;
	}

	function saveCount(n) {
		localStorage.setItem(storageCountKey, String(n));
	}

	function loadGroups() {
		const raw = localStorage.getItem(storageGroupsKey);
		if (!raw) return [];
		try {
			const parsed = JSON.parse(raw);
			return Array.isArray(parsed) ? parsed : [];
		} catch {
			return [];
		}
	}

	function saveGroups(groups) {
		localStorage.setItem(storageGroupsKey, JSON.stringify(groups));
	}

	function getGroupsSorted() {
		return loadGroups()
			.slice()
			.sort((a, b) => {
				if (a.start !== b.start) return a.start - b.start;
				if (a.end !== b.end) return a.end - b.end;
				return a.name.localeCompare(b.name);
			});
	}

	function getGroupOpenState(id) {
		return localStorage.getItem(storageGroupOpenKey(id)) === "1";
	}

	function setGroupOpenState(id, isOpen) {
		localStorage.setItem(storageGroupOpenKey(id), isOpen ? "1" : "0");
	}

	function saveRowState(id, state) {
		localStorage.setItem(storageKey(id), JSON.stringify(state));
	}

	function loadRowState(id) {
		const raw = localStorage.getItem(storageKey(id));
		try {
			return raw ? JSON.parse(raw) : null;
		} catch {
			return null;
		}
	}

	function normalizeRowState(raw) {
		const state = {
			distributed: false,
			takenBack: false,
			haveForm: false,
			showSecondStatus: false,
			status1: "",
			status2: "",
			comments: "",
			showComments: false
		};
		if (!raw || typeof raw !== "object") return state;
		state.distributed = coerceBoolean(raw.distributed);
		state.takenBack = coerceBoolean(raw.takenBack);
		state.haveForm = coerceBoolean(raw.haveForm);
		state.showSecondStatus = coerceBoolean(raw.showSecondStatus);
		state.status1 = coerceString(raw.status1).trim();
		state.status2 = coerceString(raw.status2).trim();
		state.comments = coerceString(raw.comments);
		state.showComments = coerceBoolean(raw.showComments);
		if (state.status2 && !state.showSecondStatus) {
			state.showSecondStatus = true;
		}
		if (state.comments.trim() && !state.showComments) {
			state.showComments = true;
		}
		return state;
	}

	function getRowState(id) {
		const saved = loadRowState(id);
		return normalizeRowState(saved);
	}

	function clearAllRowData() {
		const keysToRemove = [];
		for (let i = 0; i < localStorage.length; i++) {
			const key = localStorage.key(i);
			if (key && key.startsWith("blo_forms_tracker_v1_row_")) {
				keysToRemove.push(key);
			}
		}
		for (const key of keysToRemove) {
			localStorage.removeItem(key);
		}
	}

	function clearGroupOpenStates() {
		const keysToRemove = [];
		for (let i = 0; i < localStorage.length; i++) {
			const key = localStorage.key(i);
			if (key && key.startsWith("blo_forms_tracker_v1_group_open_")) {
				keysToRemove.push(key);
			}
		}
		for (const key of keysToRemove) {
			localStorage.removeItem(key);
		}
	}

	function findGroupForForm(formNumber, groups) {
		return groups.find((group) => formNumber >= group.start && formNumber <= group.end) || null;
	}

	function determineRowColorHex(state) {
		if (!state || typeof state !== "object") return "#ffffff";
		if (state.haveForm) return "#f5d0fe"; // brighter purple
		if (state.distributed && state.takenBack) return "#dcfce7"; // light green
		if (state.distributed) return "#bae6fd"; // brighter blue
		if (state.takenBack) return "#fef9c3"; // light yellow
		return "#ffffff";
	}

	function hexToRgb(hex) {
		if (typeof hex !== "string") return null;
		const normalized = hex.trim().replace("#", "");
		if (normalized.length === 3) {
			const r = normalized[0];
			const g = normalized[1];
			const b = normalized[2];
			return [
				Number.parseInt(r + r, 16),
				Number.parseInt(g + g, 16),
				Number.parseInt(b + b, 16)
			];
		}
		if (normalized.length !== 6) return null;
		const r = Number.parseInt(normalized.slice(0, 2), 16);
		const g = Number.parseInt(normalized.slice(2, 4), 16);
		const b = Number.parseInt(normalized.slice(4, 6), 16);
		if ([r, g, b].some((value) => Number.isNaN(value))) return null;
		return [r, g, b];
	}

	function getDataSnapshot() {
		const count = getSavedCount();
		if (!Number.isFinite(count) || count <= 0) {
			return {
				count: 0,
				rows: [],
				groups: [],
				groupedRows: [],
				ungroupedRows: []
			};
		}
		const groups = getGroupsSorted();
		const rows = [];
		for (let i = 1; i <= count; i++) {
			const state = getRowState(i);
			const group = findGroupForForm(i, groups);
			const rowColor = determineRowColorHex(state);
			const row = {
				formNumber: i,
				...state,
				groupId: group ? group.id : null,
				groupName: group ? group.name : "",
				colorHex: rowColor
			};
			rows.push(row);
		}
		const { groupedRows, ungroupedRows } = buildGroupedRows(rows, groups);
		return { count, rows, groups, groupedRows, ungroupedRows };
	}

	function applyPdfFilter(snapshot, filter) {
		if (!filter || filter.mode === "all") {
			return snapshot;
		}
		const baseRows = Array.isArray(snapshot.rows) ? snapshot.rows : [];
		const baseGroups = Array.isArray(snapshot.groups) ? snapshot.groups : [];
		let filteredRows = baseRows;
		let filteredGroups = baseGroups;
		if (filter.mode === "societies") {
			const selectedIds = new Set((filter.societyIds || []).filter(Boolean));
			filteredRows = baseRows.filter((row) => row.groupId && selectedIds.has(row.groupId));
			filteredGroups = baseGroups.filter((group) => selectedIds.has(group.id));
		} else if (filter.mode === "range") {
			const start = Number.isFinite(filter.start) ? filter.start : 1;
			const end = Number.isFinite(filter.end) ? filter.end : start;
			filteredRows = baseRows.filter((row) => row.formNumber >= start && row.formNumber <= end);
			const groupIdSet = new Set(
				filteredRows.filter((row) => row.groupId).map((row) => row.groupId)
			);
			filteredGroups = baseGroups.filter((group) => groupIdSet.has(group.id));
		}
		const { groupedRows, ungroupedRows } = buildGroupedRows(filteredRows, filteredGroups);
		return {
			count: filteredRows.length,
			rows: filteredRows,
			groups: filteredGroups,
			groupedRows,
			ungroupedRows
		};
	}

	function formatDateForFilename(prefix, extension) {
		const now = new Date();
		const datePart = now.toISOString().replace(/[:.]/g, "-");
		return `${prefix}-${datePart}.${extension}`;
	}

	function formatReadableDateTime() {
		return new Date().toLocaleString();
	}

	function buildSelect(currentValue) {
		const select = document.createElement("select");
		select.className = "status-select";
		for (const opt of STATUS_OPTIONS) {
			const option = document.createElement("option");
			option.value = opt;
			option.textContent = opt === "" ? "— Optional —" : opt;
			if (opt === currentValue) option.selected = true;
			select.appendChild(option);
		}
		return select;
	}

	function buildRow(id) {
		const saved = getRowState(id);

		const row = document.createElement("div");
		row.className = "row";
		function refreshRowVisuals() {
			row.classList.remove("state-dist-yes", "state-back-yes", "state-have-yes", "state-dist-and-back");
			if (saved.haveForm) {
				row.classList.add("state-have-yes");
			} else if (saved.distributed && saved.takenBack) {
				row.classList.add("state-dist-and-back");
			} else if (saved.distributed) {
				row.classList.add("state-dist-yes");
			} else if (saved.takenBack) {
				row.classList.add("state-back-yes");
			}
		}
		refreshRowVisuals();

		// Number column
		const numCol = document.createElement("div");
		numCol.className = "col-num";
		numCol.textContent = String(id);

		// Actions column
		const actionsCol = document.createElement("div");
		actionsCol.className = "col-actions";
		const actionsWrap = document.createElement("div");
		actionsWrap.className = "actions";
		const btnDist = document.createElement("button");
		btnDist.className = saved.distributed ? "btn-dist" : "btn-muted";
		btnDist.textContent = "Distributed form";
		const btnHave = document.createElement("button");
		btnHave.className = saved.haveForm ? "btn-have" : "btn-muted";
		btnHave.textContent = "I have form";
		const btnBack = document.createElement("button");
		btnBack.className = saved.takenBack ? "btn-back" : "btn-muted";
		btnBack.textContent = "Taken back form";
		const btnComment = document.createElement("button");
		btnComment.className = "secondary";
		btnComment.textContent = "Comment";
		actionsWrap.appendChild(btnDist);
		actionsWrap.appendChild(btnHave);
		actionsWrap.appendChild(btnBack);
		actionsWrap.appendChild(btnComment);
		actionsCol.appendChild(actionsWrap);

		// Status selects
		const status1Col = document.createElement("div");
		status1Col.className = "col-status status-slot-1";
		const status1Wrap = document.createElement("div");
		status1Wrap.className = "status-wrap";
		const select1 = buildSelect(saved.status1);
		const addSecondBtn = document.createElement("button");
		addSecondBtn.type = "button";
		addSecondBtn.className = "btn-plus";
		addSecondBtn.textContent = "+";
		status1Wrap.appendChild(select1);
		status1Wrap.appendChild(addSecondBtn);
		status1Col.appendChild(status1Wrap);
		// Only show when haveForm is true
		status1Col.style.display = saved.haveForm ? "" : "none";

		const status2Col = document.createElement("div");
		status2Col.className = "col-status status-slot-2";
		const select2 = buildSelect(saved.status2);
		status2Col.appendChild(select2);
		// Only show when haveForm is true and user chose to add second status
		const showSecond = Boolean(saved.haveForm && saved.showSecondStatus);
		status2Col.style.display = showSecond ? "" : "none";
		addSecondBtn.style.display = saved.haveForm && !showSecond ? "" : "none";

		// Comments
		const commentsCol = document.createElement("div");
		commentsCol.className = "col-comments";
		const textarea = document.createElement("textarea");
		textarea.className = "comments";
		textarea.placeholder = "Comments (optional)";
		textarea.value = saved.comments || "";
		commentsCol.appendChild(textarea);
		// Hide comments by default unless there is saved text or explicit toggle
		const shouldShowComments = Boolean(saved.comments && saved.comments.trim().length > 0) || Boolean(saved.showComments);
		commentsCol.style.display = shouldShowComments ? "" : "none";
		btnComment.textContent = shouldShowComments ? "Hide comment" : "Comment";

		// Handlers
		btnDist.addEventListener("click", async () => {
			if (saved.haveForm) {
				const confirmMessage = "You're changing from 'I have form' to 'Distributed form'. Continue?";
				const proceed = await showConfirm(confirmMessage);
				if (!proceed) return;
				saved.haveForm = false;
				saved.distributed = true;
				saved.takenBack = false;
				btnHave.className = "btn-muted";
				btnBack.className = "btn-muted";
				status1Col.style.display = "none";
				status2Col.style.display = "none";
				addSecondBtn.style.display = "none";
				btnDist.className = "btn-dist";
				refreshRowVisuals();
				saveRowState(id, saved);
				recalcHeaderStatusVisibility();
				return;
			}
			if (saved.distributed) {
				const confirmMessage = "You're changing from 'Distributed form' to 'Not distributed'. Continue?";
				const proceed = await showConfirm(confirmMessage);
				if (!proceed) return;
			}
			const next = !saved.distributed;
			saved.distributed = next;
			btnDist.className = next ? "btn-dist" : "btn-muted";
			refreshRowVisuals();
			saveRowState(id, saved);
		});

		btnHave.addEventListener("click", async () => {
			const next = !saved.haveForm;
			if (saved.haveForm && !next) {
				const confirmMessage = "You're changing from 'I have form' to 'Not holding the form'. Continue?";
				const proceed = await showConfirm(confirmMessage);
				if (!proceed) return;
			}
			if (!saved.haveForm && next) {
				if (saved.distributed) {
					const confirmMessage = "You're changing from 'Distributed form' to 'I have form'. Continue?";
					const proceed = await showConfirm(confirmMessage);
					if (!proceed) return;
				}
				if (saved.takenBack) {
					const confirmMessage = "You're changing from 'Taken back form' to 'I have form'. Continue?";
					const proceed = await showConfirm(confirmMessage);
					if (!proceed) return;
				}
			}
			saved.haveForm = next;
			btnHave.className = next ? "btn-have" : "btn-muted";

			// When officer has the form, it means not distributed and not taken back
			if (next) {
				saved.distributed = false;
				saved.takenBack = false;
				btnDist.className = "btn-muted";
				btnBack.className = "btn-muted";
				refreshRowVisuals();
				status1Col.style.display = "";
				// Show second status only if previously added
				status2Col.style.display = saved.showSecondStatus ? "" : "none";
				addSecondBtn.style.display = saved.showSecondStatus ? "none" : "";
			} else {
				status1Col.style.display = "none";
				status2Col.style.display = "none";
				addSecondBtn.style.display = "none";
				btnDist.className = saved.distributed ? "btn-dist" : "btn-muted";
				btnBack.className = saved.takenBack ? "btn-back" : "btn-muted";
				refreshRowVisuals();
			}
			saveRowState(id, saved);
			recalcHeaderStatusVisibility();
		});

		btnBack.addEventListener("click", async () => {
			if (saved.haveForm) return; // disabled when have form
			if (saved.takenBack) {
				const confirmMessage = "You're changing from 'Taken back form' to 'Not taken back'. Continue?";
				const proceed = await showConfirm(confirmMessage);
				if (!proceed) return;
			}
			const next = !saved.takenBack;
			saved.takenBack = next;
			btnBack.className = next ? "btn-back" : "btn-muted";
			refreshRowVisuals();
			saveRowState(id, saved);
		});

		btnComment.addEventListener("click", () => {
			const visible = commentsCol.style.display !== "none";
			const nextVisible = !visible;
			commentsCol.style.display = nextVisible ? "" : "none";
			btnComment.textContent = nextVisible ? "Hide comment" : "Comment";
			saved.showComments = nextVisible;
			saveRowState(id, saved);
			if (nextVisible) {
				textarea.focus();
			}
		});

		select1.addEventListener("change", () => {
			saved.status1 = select1.value;
			saveRowState(id, saved);
		});
		select2.addEventListener("change", () => {
			saved.status2 = select2.value;
			saveRowState(id, saved);
		});
		addSecondBtn.addEventListener("click", () => {
			saved.showSecondStatus = true;
			status2Col.style.display = "";
			addSecondBtn.style.display = "none";
			saveRowState(id, saved);
			select2.focus();
			recalcHeaderStatusVisibility();
		});
		textarea.addEventListener("input", () => {
			saved.comments = textarea.value;
			saveRowState(id, saved);
		});

		// Assemble
		row.appendChild(numCol);
		row.appendChild(actionsCol);
		row.appendChild(status1Col);
		row.appendChild(status2Col);
		row.appendChild(commentsCol);

		return row;
	}

	function buildRowSummary(row) {
		const summary = {
			title: "-",
			details: []
		};
		const hasComments = typeof row.comments === "string" && row.comments.trim().length > 0;
		const trimmedComment = hasComments ? row.comments.trim() : "";

		if (row.haveForm) {
			summary.title = "I have form";
			const statuses = [row.status1, row.status2].map((s) => (typeof s === "string" ? s.trim() : "")).filter(Boolean);
			if (statuses.length) summary.details.push(statuses.join(", "));
			if (hasComments) summary.details.push(trimmedComment);
		} else if (row.distributed && row.takenBack) {
			summary.title = "Completed";
			if (hasComments) summary.details.push(trimmedComment);
		} else if (row.distributed) {
			summary.title = "Distributed";
			if (hasComments) summary.details.push(trimmedComment);
		} else if (row.takenBack) {
			summary.title = "Taken back form";
			if (hasComments) summary.details.push(trimmedComment);
		} else if (hasComments) {
			summary.details.push(trimmedComment);
		}

		return summary;
	}

	function renderList(n) {
		renderSocietyManager();
		els.list.textContent = "";
		if (!Number.isFinite(n) || n <= 0) {
			els.listHeader.classList.add("hidden");
			return;
		}
		els.listHeader.classList.remove("hidden");

		const groups = getGroupsSorted().map((group) => ({ ...group, numbers: [] }));
		const ungrouped = [];

		for (let i = 1; i <= n; i++) {
			const match = groups.find((group) => i >= group.start && i <= group.end);
			if (match) {
				match.numbers.push(i);
			} else {
				ungrouped.push(i);
			}
		}

		const fragment = document.createDocumentFragment();
		const hasNamedGroups = groups.length > 0;
		const buildGroupSection = (group, numbers, isUngrouped = false) => {
			if (numbers.length === 0) return;
			const details = document.createElement("details");
			details.className = "society";
			const groupId = isUngrouped ? "ungrouped" : group.id;
			if (getGroupOpenState(groupId)) {
				details.open = true;
			}
			const summary = document.createElement("summary");
			const summaryText = document.createElement("span");
			summaryText.textContent = isUngrouped
				? hasNamedGroups ? "Ungrouped forms" : "All forms"
				: `${group.name} (${group.start}–${group.end})`;
			const countSpan = document.createElement("span");
			countSpan.className = "society-summary-count";
			countSpan.textContent = `${numbers.length} form${numbers.length === 1 ? "" : "s"}`;
			summary.appendChild(summaryText);
			summary.appendChild(countSpan);
			details.appendChild(summary);
			const body = document.createElement("div");
			body.className = "society-body";
			const renderRowsIntoBody = () => {
				// Avoid re-rendering if already populated
				if (body.childElementCount > 0) return;
				const rowsFragment = document.createDocumentFragment();
				for (const num of numbers) {
					rowsFragment.appendChild(buildRow(num));
				}
				body.appendChild(rowsFragment);
			};
			// Lazy render: only populate when opened
			if (details.open) {
				renderRowsIntoBody();
			}
			details.appendChild(body);
			details.addEventListener("toggle", () => {
				setGroupOpenState(groupId, details.open);
				if (details.open) {
					renderRowsIntoBody();
				}
				recalcHeaderStatusVisibility();
			});
			fragment.appendChild(details);
		};

		for (const group of groups) {
			buildGroupSection(group, group.numbers, false);
		}
		buildGroupSection({ id: "ungrouped", name: "Ungrouped" }, ungrouped, true);

		els.list.appendChild(fragment);
		recalcHeaderStatusVisibility();
	}

	async function handleGenerate() {
		await processCountSubmission(els.numInput ? els.numInput.value : "", "initial");
	}

	async function handleSettingsSave() {
		await processCountSubmission(els.settingsNumInput ? els.settingsNumInput.value : "", "settings");
	}

	async function handleClearAll() {
		const proceed = await showConfirm("Clear All Data will remove every saved form, society, and local setting (except theme). This cannot be undone. Do you want to continue?", true);
		if (!proceed) return;
		const theme = localStorage.getItem(storageThemeKey); // keep theme
		const groupsRaw = localStorage.getItem(storageGroupsKey); // keep group definitions
		localStorage.clear();
		if (theme) localStorage.setItem(storageThemeKey, theme);
		if (groupsRaw) localStorage.setItem(storageGroupsKey, groupsRaw);
		syncFormInputs(null);
		setFirstRunMode(true);
		els.list.textContent = "";
		els.listHeader.classList.add("hidden");
		showWarning("All data cleared.");
		showSettingsMessage("All data cleared. Enter a new number to start again.", "warn");
		renderSocietyManager();
	}

	// Settings modal
	function openSettings() {
		if (!els.settingsModal) return;
		els.settingsModal.setAttribute("aria-hidden", "false");
		document.body.classList.add("modal-open");
		renderSocietyManager();
		if (els.societyName) els.societyName.focus();
	}
	function closeSettings() {
		if (!els.settingsModal) return;
		els.settingsModal.setAttribute("aria-hidden", "true");
		document.body.classList.remove("modal-open");
		hideExportOptions();
		showImportExportMessage("");
	}

	function recalcHeaderStatusVisibility() {
		// Show Status 1/2 headers only if any row shows its status selects
		const statusCols = Array.from(els.list.querySelectorAll(".col-status"));
		const anyVisible = statusCols.some((col) => getComputedStyle(col).display !== "none");
		if (els.hdrStatus1) els.hdrStatus1.style.display = anyVisible ? "" : "none";
		if (els.hdrStatus2) els.hdrStatus2.style.display = anyVisible ? "" : "none";
	}

	function hideExportOptions() {
		if (els.exportOptions && !els.exportOptions.classList.contains("hidden")) {
			els.exportOptions.classList.add("hidden");
			closePdfPanel();
		}
	}

	function toggleExportOptions() {
		if (!els.exportOptions) return;
		els.exportOptions.classList.toggle("hidden");
		closePdfPanel();
	}

	function exportToPdf(filter) {
		const baseSnapshot = getDataSnapshot();
		const titleFontSize = 11;
		const detailsFontSize = 10;
		if (baseSnapshot.count === 0) {
			showImportExportMessage("Generate forms before exporting.", "warn");
			return;
		}
		if (!window.jspdf || !window.jspdf.jsPDF) {
			showImportExportMessage("PDF export library failed to load. Check your internet connection.", "error");
			return;
		}
		const snapshot = applyPdfFilter(baseSnapshot, filter);
		if (snapshot.count === 0) {
			showImportExportMessage("No forms match the selected filters.", "warn");
			return;
		}
		const { jsPDF } = window.jspdf;
		const doc = new jsPDF({ orientation: "landscape", unit: "pt", format: "a4" });
		const marginLeft = 40;
		const marginRight = 40;
		const pageWidth = doc.internal.pageSize.getWidth();
		const tableWidth = pageWidth - marginLeft - marginRight;
		const colWidthNum = 60;
		const colWidthStatus = 240;
		const colWidthDetails = Math.max(140, tableWidth - colWidthNum - colWidthStatus);
		const titleY = 50;
		doc.setFontSize(18);
		doc.text("BLO Forms Report", marginLeft, titleY);
		doc.setFontSize(12);
		const infoStart = titleY + 20;
		doc.text(`Generated: ${formatReadableDateTime()}`, marginLeft, infoStart);
		doc.text(`Total Forms: ${snapshot.count}`, marginLeft, infoStart + 18);
		if (snapshot.groups.length > 0) {
			doc.text(`Societies: ${snapshot.groups.length}`, marginLeft, infoStart + 36);
		}
		if (typeof doc.autoTable !== "function") {
			showImportExportMessage("PDF table helper not available. Please reload and try again.", "error");
			return;
		}
		const sections = [];
		for (const entry of snapshot.groupedRows) {
			if (!entry || !entry.group || entry.rows.length === 0) continue;
			const countLabel = `${entry.rows.length} form${entry.rows.length === 1 ? "" : "s"}`;
			const title = `${entry.group.name} (${entry.group.start}–${entry.group.end}) • ${countLabel}`;
			sections.push({ title, rows: entry.rows });
		}
		if (snapshot.groups.length > 0) {
			if (snapshot.ungroupedRows.length > 0) {
				const countLabel = `${snapshot.ungroupedRows.length} form${snapshot.ungroupedRows.length === 1 ? "" : "s"}`;
				sections.push({ title: `Ungrouped forms • ${countLabel}`, rows: snapshot.ungroupedRows });
			}
		} else {
			const countLabel = `${snapshot.count} form${snapshot.count === 1 ? "" : "s"}`;
			sections.push({ title: `All forms • ${countLabel}`, rows: snapshot.rows });
		}
		let currentY = infoStart + (snapshot.groups.length > 0 ? 72 : 54);
		const pageHeight = doc.internal.pageSize.getHeight();
		for (const section of sections) {
			if (!section.rows || section.rows.length === 0) continue;
			if (currentY > pageHeight - 80) {
				doc.addPage();
				currentY = 60;
			}
			doc.setFontSize(14);
			doc.setTextColor(34, 34, 34);
			doc.text(section.title, marginLeft, currentY);
			currentY += 16;
			const body = section.rows.map((row) => {
				const summaryBlocks = buildRowSummary(row);
				const arr = [
					row.formNumber,
					summaryBlocks.title,
					summaryBlocks.details.length ? summaryBlocks.details : []
				];
				arr.rowData = row;
				return arr;
			});
			doc.autoTable({
				startY: currentY,
				head: [["Form #", "Status", "Details"]],
				body,
				margin: { left: marginLeft, right: marginRight },
				styles: {
					fontSize: 10,
					cellPadding: 6,
					valign: "top",
					textColor: [0, 0, 0]
				},
				headStyles: { fillColor: [34, 150, 243], textColor: 255, fontStyle: "bold" },
				columnStyles: {
					0: { halign: "center", cellWidth: colWidthNum },
					1: { cellWidth: colWidthStatus },
					2: { cellWidth: colWidthDetails }
				},
				didParseCell: (data) => {
					if (data.section !== "body") return;
					const rowData = data.row && data.row.raw && data.row.raw.rowData;
					if (!rowData || !rowData.colorHex) return;
					const rgb = hexToRgb(rowData.colorHex);
					if (!rgb) return;
					data.cell.styles.fillColor = rgb;
					if (data.column.index === 1 && data.row.raw[1] && typeof data.row.raw[1] === "string") {
						data.cell.styles.fontSize = titleFontSize;
						data.cell.styles.fontStyle = "normal";
						data.cell.styles.textColor = [15, 23, 42];
						data.cell.text = [data.row.raw[1].toUpperCase()];
					}
					if (data.column.index === 2 && data.row.raw[2] && Array.isArray(data.row.raw[2])) {
						data.cell.styles.fontSize = detailsFontSize;
						data.cell.styles.fontStyle = "normal";
						data.cell.styles.textColor = [31, 41, 55];
						data.cell.styles.cellPadding = { top: 6, right: 10, bottom: 6, left: 10 };
						data.cell.text = data.row.raw[2];
					}
				}
			});
			currentY = (doc.lastAutoTable && doc.lastAutoTable.finalY ? doc.lastAutoTable.finalY : currentY) + 18;
			doc.setFontSize(12);
			doc.setTextColor(0, 0, 0);
		}
		const filename = formatDateForFilename("blo-forms-report", "pdf");
		doc.save(filename);
		showImportExportMessage(`PDF exported: ${filename}`, "success");
		hideExportOptions();
	}

	function exportToExcel() {
		const snapshot = getDataSnapshot();
		if (snapshot.count === 0) {
			showImportExportMessage("Generate forms before exporting.", "warn");
			return;
		}
		if (typeof XLSX === "undefined") {
			showImportExportMessage("Excel export library failed to load. Check your internet connection.", "error");
			return;
		}
		const workbook = XLSX.utils.book_new();
		const formsSheetData = [
			["FormNumber", "GroupName", "Distributed", "TakenBack", "HaveForm", "Status1", "Status2", "Comments", "ShowSecondStatus", "ShowComments"]
		];
		for (const row of snapshot.rows) {
			formsSheetData.push([
				row.formNumber,
				row.groupName || "",
				row.distributed ? "Yes" : "No",
				row.takenBack ? "Yes" : "No",
				row.haveForm ? "Yes" : "No",
				row.status1,
				row.status2,
				row.comments,
				row.showSecondStatus ? "Yes" : "No",
				row.showComments ? "Yes" : "No"
			]);
		}
		const formsSheet = XLSX.utils.aoa_to_sheet(formsSheetData);
		XLSX.utils.book_append_sheet(workbook, formsSheet, "Forms");

		const groupsSheetData = [["Name", "Start", "End"]];
		for (const group of snapshot.groups) {
			groupsSheetData.push([group.name, group.start, group.end]);
		}
		const groupsSheet = XLSX.utils.aoa_to_sheet(groupsSheetData);
		XLSX.utils.book_append_sheet(workbook, groupsSheet, "Societies");

		const settingsSheetData = [
			["Key", "Value"],
			["formsCount", snapshot.count],
			["generatedAt", formatReadableDateTime()]
		];
		const settingsSheet = XLSX.utils.aoa_to_sheet(settingsSheetData);
		XLSX.utils.book_append_sheet(workbook, settingsSheet, "Settings");

		const filename = formatDateForFilename("blo-forms-data", "xlsx");
		XLSX.writeFile(workbook, filename);
		showImportExportMessage(`Excel exported: ${filename}`, "success");
		hideExportOptions();
	}

	async function importFromExcelFile(file) {
		if (!file) return;
		if (typeof XLSX === "undefined") {
			showImportExportMessage("Excel library not available. Please check your internet connection.", "error");
			return;
		}
		const proceed = await showConfirm(
			`Importing "${file.name}" will replace your current forms and societies. Continue?`,
			true
		);
		if (!proceed) {
			if (els.importExcelInput) {
				els.importExcelInput.value = "";
			}
			return;
		}
		showImportExportMessage("Importing data…", "warn");
		try {
			const buffer = await file.arrayBuffer();
			const workbook = XLSX.read(buffer, { type: "array" });

			const settingsSheet = workbook.Sheets.Settings;
			if (!settingsSheet) {
				throw new Error("Settings sheet not found in file.");
			}
			const settingsRows = XLSX.utils.sheet_to_json(settingsSheet, { header: ["key", "value"], defval: "" });
			const countRow = settingsRows.find(
				(entry) => coerceString(entry.key).trim().toLowerCase() === "formscount"
			);
			const parsedCount = Number.parseInt(coerceString(countRow ? countRow.value : "").trim(), 10);
			if (!Number.isFinite(parsedCount) || parsedCount < 1 || parsedCount > MAX_FORMS) {
				throw new Error("Invalid forms count in file.");
			}

			const formsSheet = workbook.Sheets.Forms;
			if (!formsSheet) {
				throw new Error("Forms sheet not found in file.");
			}
			const formsRows = XLSX.utils.sheet_to_json(formsSheet, { defval: "" });
			const stateMap = new Map();
			for (const row of formsRows) {
				const rawNumber =
					row.FormNumber ?? row.formNumber ?? row["Form #"] ?? row["Number"] ?? row["FormNumber"];
				const formNumber = Number.parseInt(rawNumber, 10);
				if (!Number.isFinite(formNumber) || formNumber < 1 || formNumber > parsedCount) continue;
				const normalized = normalizeRowState({
					distributed: row.Distributed,
					takenBack: row.TakenBack,
					haveForm: row.HaveForm,
					showSecondStatus: row.ShowSecondStatus,
					status1: row.Status1,
					status2: row.Status2,
					comments: row.Comments,
					showComments: row.ShowComments
				});
				stateMap.set(formNumber, normalized);
			}

			clearAllRowData();

			for (let i = 1; i <= parsedCount; i++) {
				const state = stateMap.get(i) || normalizeRowState(null);
				saveRowState(i, state);
			}

			const groupsSheet = workbook.Sheets.Societies || workbook.Sheets.Groups || null;
			const importedGroups = [];
			if (groupsSheet) {
				const groupsRows = XLSX.utils.sheet_to_json(groupsSheet, { defval: "" });
				let index = 0;
				for (const row of groupsRows) {
					const name = coerceString(row.Name ?? row.name).trim();
					const start = Number.parseInt(row.Start ?? row.start, 10);
					const end = Number.parseInt(row.End ?? row.end, 10);
					if (!name || !Number.isFinite(start) || !Number.isFinite(end)) continue;
					if (start < 1 || end < 1 || start > MAX_FORMS || end > MAX_FORMS || start > end) continue;
					index += 1;
					importedGroups.push({
						id: `imported_${Date.now()}_${index}`,
						name,
						start,
						end
					});
				}
			}

			saveGroups(importedGroups);
			clearGroupOpenStates();
			saveCount(parsedCount);
			syncFormInputs(parsedCount);
			setFirstRunMode(false);
			renderSocietyManager();
			renderList(parsedCount);
			showWarning("");
			showSettingsMessage("");
			showImportExportMessage("Import successful. Your data has been updated.", "success");
		} catch (error) {
			console.error("Import failed", error);
			showImportExportMessage(
				error && error.message ? `Import failed: ${error.message}` : "Import failed due to an unknown error.",
				"error"
			);
		} finally {
			if (els.importExcelInput) {
				els.importExcelInput.value = "";
			}
		}
	}

	// Theme handling
	function applyTheme(theme) {
		// theme: 'light' | 'dark' | 'system'
		let resolved = theme;
		if (theme === "system") {
			const prefersDark = window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches;
			resolved = prefersDark ? "dark" : "light";
		}
		document.documentElement.setAttribute("data-theme", resolved === "light" ? "light" : "dark");
		if (els.themeToggle) {
			// Use emoji icon to indicate the action (moon shown in light theme, sun shown in dark theme)
			els.themeToggle.textContent = resolved === "light" ? "🌙" : "☀️";
			const label = resolved === "light" ? "Switch to Night Mode" : "Switch to Day Mode";
			els.themeToggle.title = label;
			els.themeToggle.setAttribute("aria-label", label);
		}
	}

	function getSavedTheme() {
		return localStorage.getItem(storageThemeKey);
	}

	function setSavedTheme(theme) {
		localStorage.setItem(storageThemeKey, theme);
	}

	function toggleTheme() {
		// Toggle only between light and dark for simplicity
		const current = document.documentElement.getAttribute("data-theme") || "dark";
		const next = current === "light" ? "dark" : "light";
		setSavedTheme(next);
		applyTheme(next);
	}

	// Init
	els.generateBtn.addEventListener("click", handleGenerate);
	els.numInput.addEventListener("keyup", (e) => {
		if (e.key === "Enter") handleGenerate();
	});
	els.clearAllBtn.addEventListener("click", handleClearAll);
	if (els.themeToggle) {
		els.themeToggle.addEventListener("click", toggleTheme);
	}
	if (els.settingsBtn) {
		els.settingsBtn.addEventListener("click", openSettings);
	}
	if (els.settingsClose) {
		els.settingsClose.addEventListener("click", closeSettings);
	}
	if (els.settingsSaveFormsBtn) {
		els.settingsSaveFormsBtn.addEventListener("click", handleSettingsSave);
	}
	if (els.settingsModal) {
		els.settingsModal.addEventListener("click", (e) => {
			const target = e.target;
			if (target && target.getAttribute && target.getAttribute("data-close") === "true") {
				closeSettings();
			}
		});
		window.addEventListener("keydown", (e) => {
			if (e.key === "Escape" && els.settingsModal.getAttribute("aria-hidden") === "false") {
				closeSettings();
			}
		});
	}
	if (els.addSocietyBtn) {
		els.addSocietyBtn.addEventListener("click", handleAddSociety);
	}
	if (els.installBtn) {
		els.installBtn.addEventListener("click", async () => {
			if (!deferredInstallPrompt) {
				showWarning("To install: Use your browser menu to Install/Add to Home Screen.");
				return;
			}
			toggleInstallButton(false);
			deferredInstallPrompt.prompt();
			try {
				const outcome = await deferredInstallPrompt.userChoice;
				if (outcome && outcome.outcome === "accepted") {
					announceInstallSuccess();
				}
			} catch (error) {
				console.error("Install prompt failed:", error);
			}
			deferredInstallPrompt = null;
		});
	}
	const societyInputs = [els.societyName, els.societyStart, els.societyEnd].filter(Boolean);
	for (const input of societyInputs) {
		input.addEventListener("keyup", (event) => {
			if (event.key === "Enter") {
				handleAddSociety();
			}
		});
	}
	if (els.settingsNumInput) {
		els.settingsNumInput.addEventListener("keyup", (event) => {
			if (event.key === "Enter") {
				handleSettingsSave();
			}
		});
	}
	if (els.exportToggle) {
		els.exportToggle.addEventListener("click", (event) => {
			event.preventDefault();
			event.stopPropagation();
			showImportExportMessage("");
			toggleExportOptions();
		});
	}
	if (els.exportPdfBtn) {
		els.exportPdfBtn.addEventListener("click", (event) => {
			event.preventDefault();
			event.stopPropagation();
			showImportExportMessage("");
			openPdfPanel();
		});
	}
	if (els.exportExcelBtn) {
		els.exportExcelBtn.addEventListener("click", (event) => {
			event.preventDefault();
			event.stopPropagation();
			exportToExcel();
		});
	}
	const pdfModeRadios = [els.exportPdfModeAll, els.exportPdfModeSociety, els.exportPdfModeRange].filter(Boolean);
	for (const radio of pdfModeRadios) {
		radio.addEventListener("change", () => {
			showImportExportMessage("");
			syncPdfExportModeUI();
		});
	}
	if (els.exportPdfConfirm) {
		els.exportPdfConfirm.addEventListener("click", () => {
			const filter = collectPdfFilter();
			if (!filter) return;
			closePdfPanel();
			exportToPdf(filter);
		});
	}
	if (els.exportPdfCancel) {
		els.exportPdfCancel.addEventListener("click", () => {
			closePdfPanel();
		});
	}
	if (els.importExcelBtn && els.importExcelInput) {
		els.importExcelBtn.addEventListener("click", (event) => {
			event.preventDefault();
			if (els.importExcelInput) {
				els.importExcelInput.click();
			}
		});
		els.importExcelInput.addEventListener("change", async (event) => {
			const file = event.target && event.target.files ? event.target.files[0] : null;
			if (file) {
				await importFromExcelFile(file);
			}
		});
	}
	window.addEventListener("beforeinstallprompt", (event) => {
		event.preventDefault();
		deferredInstallPrompt = event;
		const isStandalone = window.matchMedia && window.matchMedia("(display-mode: standalone)").matches;
		const isIosStandalone = typeof window.navigator.standalone === "boolean" && window.navigator.standalone;
		if (!isStandalone && !isIosStandalone) {
			toggleInstallButton(true);
		}
	});
	window.addEventListener("appinstalled", () => {
		deferredInstallPrompt = null;
		toggleInstallButton(false);
		announceInstallSuccess();
	});
	document.addEventListener("click", (event) => {
		if (!els.exportOptions || !els.exportToggle) return;
		if (els.exportOptions.classList.contains("hidden")) return;
		if (els.exportOptions.contains(event.target) || els.exportToggle.contains(event.target)) return;
		hideExportOptions();
	});

	// Restore prior state if present
	const savedTheme = getSavedTheme();
	if (savedTheme) {
		applyTheme(savedTheme);
	} else {
		applyTheme("system");
	}

	renderSocietyManager();

	const savedCount = getSavedCount();
	syncFormInputs(savedCount);
	setFirstRunMode(!savedCount);
	if (savedCount) {
		renderList(savedCount);
	}

	const isAlreadyInstalled =
		(window.matchMedia && window.matchMedia("(display-mode: standalone)").matches) ||
		(typeof window.navigator.standalone === "boolean" && window.navigator.standalone);
	// Default: keep banner hidden; it will be shown when beforeinstallprompt fires
	if (isAlreadyInstalled) {
		toggleInstallButton(false);
	} else {
		toggleInstallButton(false);
	}

	if ("serviceWorker" in navigator) {
		window.addEventListener("load", async () => {
			navigator.serviceWorker.register("sw.js").catch((error) => {
				console.error("Service worker registration failed:", error);
			});
		});
	}
})();


