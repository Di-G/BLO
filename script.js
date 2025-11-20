(() => {
	const MAX_FORMS = 10000;
	const ASSET_VERSION = "12";
	const DATA_EXPORT_VERSION = 3;
	const ROW_STATE_EXPORT_KEYS = [
		"distributed",
		"takenBack",
		"haveForm",
		"onlineUploaded",
		"missedForm",
		"distributionType",
		"showSecondStatus",
		"status1",
		"status2",
		"comments",
		"showComments",
		"deleted",
		"missing",
		"name",
		"phone",
		"epic"
	];
	const FORM_EXPORT_HEADERS = [
		"FormNumber",
		"GroupId",
		"GroupName",
		"GroupRangeStart",
		"GroupRangeEnd",
		"Name",
		"PhoneNumber",
		"EpicNumber",
		"Distributed",
		"DistributionType",
		"Collected",
		"TakenBack",
		"HaveForm",
		"OnlineUploaded",
		"MissedForm",
		"NotReturned",
		"Deleted",
		"FormMissing",
		"ShowSecondStatus",
		"Status1",
		"Status2",
		"ShowComments",
		"Comments",
		"NoEntry",
		"ColorHex",
		"RawStateJson"
	];
	const GROUP_EXPORT_HEADERS = ["Id", "Name", "Start", "End"];
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
		topShell: document.querySelector(".top-shell"),
		searchPanel: document.getElementById("searchPanel"),
		searchInput: document.getElementById("searchInput"),
		searchSubmit: document.getElementById("searchSubmit"),
		searchStatus: document.getElementById("searchStatus"),
		searchPrev: document.getElementById("searchPrev"),
		searchNext: document.getElementById("searchNext"),
		searchDone: document.getElementById("searchDone"),
		installBtn: document.getElementById("installBtn"),
		controls: document.querySelector(".controls"),
		warning: document.getElementById("warning"),
		listHeader: document.getElementById("listHeader"),
		list: document.getElementById("list"),
		analysisPanel: document.getElementById("analysisPanel"),
		formStats: document.getElementById("formStats"),
		statsDistributed: document.getElementById("statsDistributed"),
		statsCollected: document.getElementById("statsCollected"),
		statsHave: document.getElementById("statsHave"),
		statsOnlineUploaded: document.getElementById("statsOnlineUploaded"),
		statsMissing: document.getElementById("statsMissing"),
		statsDeleted: document.getElementById("statsDeleted"),
		statsNoEntry: document.getElementById("statsNoEntry"),
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
	const storageInstalledKey = "blo_forms_tracker_v1_installed";
	const storageGroupOpenKey = (id) => `blo_forms_tracker_v1_group_open_${id}`;

	let deferredInstallPrompt = null;
	let installMessageTimeout = null;
	let societyEditState = null;
	let searchState = null;
	let lastSearchRow = null;
	const groupRenderers = new Map();
	let searchHeightRaf = null;
	let statsUpdateScheduled = false;
	let groupStatsUpdateScheduled = false;
	const pendingGroupStatUpdates = new Set();

	function isAppMarkedInstalled() {
		return localStorage.getItem(storageInstalledKey) === "1";
	}

	function markAppInstalled() {
		localStorage.setItem(storageInstalledKey, "1");
		toggleInstallButton(false);
	}

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
		markAppInstalled();
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

	function startSocietyEdit(group) {
		if (!group) return;
		societyEditState = {
			id: group.id,
			name: group.name,
			start: String(group.start),
			end: String(group.end),
			error: ""
		};
		renderSocietyManager();
		requestAnimationFrame(() => {
			const editor = els.societyList
				? els.societyList.querySelector('.society-item.editing input[name="societyEditName"]')
				: null;
			if (editor && typeof editor.focus === "function") {
				editor.focus();
			}
		});
	}

	function cancelSocietyEdit() {
		societyEditState = null;
		renderSocietyManager();
	}

	function updateSocietyEditState(field, value) {
		if (!societyEditState) return;
		societyEditState[field] = value;
	}

	function handleSocietyEditSave(groupId) {
		if (!societyEditState || societyEditState.id !== groupId) return;
		const name = societyEditState.name ? societyEditState.name.trim() : "";
		const startValue = societyEditState.start != null ? String(societyEditState.start).trim() : "";
		const endValue = societyEditState.end != null ? String(societyEditState.end).trim() : "";
		const start = Number.parseInt(startValue, 10);
		const end = Number.parseInt(endValue, 10);

		const setError = (message) => {
			societyEditState.error = message;
			renderSocietyManager();
		};

		if (!name) {
			setError("Please enter a society name.");
			return;
		}
		if (!Number.isFinite(start) || !Number.isFinite(end)) {
			setError("Please enter valid start and end numbers.");
			return;
		}
		if (start < 1 || end < 1 || start > MAX_FORMS || end > MAX_FORMS) {
			setError(`Numbers must be between 1 and ${MAX_FORMS}.`);
			return;
		}
		if (start > end) {
			setError("Start number cannot be greater than end number.");
			return;
		}

		const groups = loadGroups();
		const targetIndex = groups.findIndex((group) => group.id === groupId);
		if (targetIndex === -1) {
			societyEditState = null;
			renderSocietyManager();
			return;
		}
		groups[targetIndex] = {
			...groups[targetIndex],
			name,
			start,
			end
		};
		saveGroups(groups);
		societyEditState = null;
		renderSocietyManager();
		const count = getSavedCount();
		if (count) renderList(count);
		showSocietyError("");
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
		const computeCountLabel = (startValue, endValue) => {
			if (!Number.isFinite(totalForms)) return "";
			const startNum = Number.parseInt(startValue, 10);
			const endNum = Number.parseInt(endValue, 10);
			if (!Number.isFinite(startNum) || !Number.isFinite(endNum)) {
				return " · 0 forms in current list";
			}
			const boundedStart = Math.max(1, startNum);
			const boundedEnd = Math.min(totalForms, endNum);
			const count = Math.max(0, boundedEnd - boundedStart + 1);
			return count > 0 ? ` · ${count} form${count === 1 ? "" : "s"}` : " · 0 forms in current list";
		};
		const formatRangeValue = (value) => {
			const trimmed = value == null ? "" : String(value).trim();
			return trimmed === "" ? "?" : trimmed;
		};
		for (const group of groups) {
			const item = document.createElement("li");
			item.className = "society-item";
			const isEditing = societyEditState && societyEditState.id === group.id;
			if (isEditing) {
				item.classList.add("editing");
			}
			const header = document.createElement("div");
			header.className = "society-item-header";
			const info = document.createElement("div");
			info.className = "society-item-info";
			const name = document.createElement("span");
			name.className = "society-item-name";
			const displayName = isEditing ? societyEditState.name : group.name;
			name.textContent = displayName || "Untitled society";
			const range = document.createElement("span");
			range.className = "society-item-range";
			const setRangeDisplay = (startValue, endValue) => {
				const countLabel = computeCountLabel(startValue, endValue);
				range.textContent = `Numbers ${formatRangeValue(startValue)}–${formatRangeValue(endValue)}${countLabel}`;
			};
			setRangeDisplay(isEditing ? societyEditState.start : group.start, isEditing ? societyEditState.end : group.end);
			info.appendChild(name);
			info.appendChild(range);
			const actions = document.createElement("div");
			actions.className = "society-item-actions";
			const editBtn = document.createElement("button");
			editBtn.type = "button";
			editBtn.className = "society-icon-btn";
			editBtn.innerHTML = '<span aria-hidden="true">✏️</span>';
			editBtn.setAttribute("aria-label", `Edit ${group.name}`);
			editBtn.title = "Edit";
			editBtn.addEventListener("click", () => {
				startSocietyEdit(group);
			});
			if (isEditing) {
				editBtn.disabled = true;
			}
			const removeBtn = document.createElement("button");
			removeBtn.type = "button";
			removeBtn.className = "society-icon-btn";
			removeBtn.innerHTML = '<span aria-hidden="true">🗑️</span>';
			removeBtn.setAttribute("aria-label", `Remove ${group.name}`);
			removeBtn.title = "Remove";
			removeBtn.addEventListener("click", async () => {
				const proceed = await showConfirm(`Remove society "${group.name}"?`);
				if (!proceed) return;
				if (societyEditState && societyEditState.id === group.id) {
					societyEditState = null;
				}
				const remaining = loadGroups().filter((g) => g.id !== group.id);
				saveGroups(remaining);
				renderSocietyManager();
				const count = getSavedCount();
				if (count) renderList(count);
			});
			actions.appendChild(editBtn);
			actions.appendChild(removeBtn);
			header.appendChild(info);
			header.appendChild(actions);
			item.appendChild(header);
			if (isEditing) {
				const editBlock = document.createElement("div");
				editBlock.className = "society-edit";
				const fields = document.createElement("div");
				fields.className = "society-edit-fields";
				const error = document.createElement("p");
				error.className = "society-edit-error";
				error.textContent = societyEditState.error || "";
				const nameInput = document.createElement("input");
				nameInput.type = "text";
				nameInput.name = "societyEditName";
				nameInput.value = societyEditState.name;
				nameInput.placeholder = "Society name";
				nameInput.addEventListener("input", (event) => {
					updateSocietyEditState("name", event.target.value);
					name.textContent = event.target.value.trim() ? event.target.value : "Untitled society";
					if (societyEditState.error) {
						societyEditState.error = "";
						error.textContent = "";
					}
				});
				const startInput = document.createElement("input");
				startInput.type = "number";
				startInput.name = "societyEditStart";
				startInput.value = societyEditState.start;
				startInput.placeholder = "Start";
				startInput.addEventListener("input", (event) => {
					updateSocietyEditState("start", event.target.value);
					setRangeDisplay(societyEditState.start, societyEditState.end);
					if (societyEditState.error) {
						societyEditState.error = "";
						error.textContent = "";
					}
				});
				const endInput = document.createElement("input");
				endInput.type = "number";
				endInput.name = "societyEditEnd";
				endInput.value = societyEditState.end;
				endInput.placeholder = "End";
				endInput.addEventListener("input", (event) => {
					updateSocietyEditState("end", event.target.value);
					setRangeDisplay(societyEditState.start, societyEditState.end);
					if (societyEditState.error) {
						societyEditState.error = "";
						error.textContent = "";
					}
				});
				fields.appendChild(nameInput);
				fields.appendChild(startInput);
				fields.appendChild(endInput);
				editBlock.appendChild(fields);
				editBlock.appendChild(error);
				const editActions = document.createElement("div");
				editActions.className = "society-edit-actions";
				const saveBtn = document.createElement("button");
				saveBtn.type = "button";
				saveBtn.textContent = "Save";
				saveBtn.addEventListener("click", () => handleSocietyEditSave(group.id));
				const cancelBtn = document.createElement("button");
				cancelBtn.type = "button";
				cancelBtn.className = "secondary";
				cancelBtn.textContent = "Cancel";
				cancelBtn.addEventListener("click", cancelSocietyEdit);
				editActions.appendChild(saveBtn);
				editActions.appendChild(cancelBtn);
				editBlock.appendChild(editActions);
				item.appendChild(editBlock);
			}
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
			missedForm: false,
			onlineUploaded: false,
			distributionType: "self",
			showSecondStatus: false,
			status1: "",
			status2: "",
			comments: "",
			showComments: false,
			deleted: false,
			missing: false,
			name: "",
			phone: "",
			epic: ""
		};
		if (!raw || typeof raw !== "object") return state;
		state.distributed = coerceBoolean(raw.distributed);
		state.takenBack = coerceBoolean(raw.takenBack);
		state.haveForm = coerceBoolean(raw.haveForm);
		state.showSecondStatus = coerceBoolean(raw.showSecondStatus);
		state.missedForm = coerceBoolean(
			raw.missedForm ??
				raw.didntReceiveBack ??
				raw.didNotReceiveBack ??
				raw.notReturned ??
				raw.notReturnedForm
		);
		state.onlineUploaded = coerceBoolean(
			raw.onlineUploaded ??
				raw.uploadedOnline ??
				raw.onlineUpload ??
				raw.onlineUploadedFlag ??
				raw.formUploadedOnline
		);
		const normalizedDistributionType = coerceString(raw.distributionType)
			.trim()
			.toLowerCase()
			.replace(/[\s_-]+/g, " ");
		if (normalizedDistributionType === "given to" || normalizedDistributionType === "given") {
			state.distributionType = "given";
		} else if (normalizedDistributionType === "self taken" || normalizedDistributionType === "self") {
			state.distributionType = "self";
		}
		state.status1 = coerceString(raw.status1).trim();
		state.status2 = coerceString(raw.status2).trim();
		state.comments = coerceString(raw.comments);
		state.showComments = coerceBoolean(raw.showComments);
		state.deleted = coerceBoolean(raw.deleted ?? raw.isDeleted ?? raw.removed);
		state.missing = coerceBoolean(raw.missing ?? raw.formMissing ?? raw.isMissing);
		state.name = coerceString(raw.name ?? raw.personName ?? raw.fullName).trim();
		state.phone = coerceString(raw.phone ?? raw.phoneNumber ?? raw.contactNumber ?? raw.mobile).trim();
		state.epic = coerceString(raw.epic ?? raw.epicNumber ?? raw.epicNo ?? raw.epicId ?? raw.epicCode).trim();
		if (state.status2 && !state.showSecondStatus) {
			state.showSecondStatus = true;
		}
		if (state.comments.trim() && !state.showComments) {
			state.showComments = true;
		}
		if (state.missedForm) {
			state.haveForm = false;
			state.takenBack = false;
		}
		if (state.deleted || state.missing) {
			state.distributed = false;
			state.takenBack = false;
			state.haveForm = false;
			state.missedForm = false;
			state.onlineUploaded = false;
			state.showSecondStatus = false;
			state.status1 = "";
			state.status2 = "";
			state.showComments = false;
		}
		if (state.distributionType === "given") {
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
		if (state.deleted) return "#e2e8f0"; // muted gray for deleted
		if (state.missing) return "#fef3c7"; // pale amber for missing
		if (state.onlineUploaded) return "#bbf7d0"; // mint green for online uploaded
		if (state.missedForm) return "#fecaca"; // deeper red for not returned
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

	function formatBooleanForExcel(value) {
		return value ? "Yes" : "No";
	}

	function safeStringForExcel(value) {
		return value == null ? "" : String(value);
	}

	function firstPopulatedValue(...values) {
		for (const value of values) {
			if (value === undefined || value === null) continue;
			if (typeof value === "string" && value.trim() === "") continue;
			return value;
		}
		return undefined;
	}

	function buildRowStateExportPayload(row) {
		const payload = {};
		if (!row || typeof row !== "object") {
			for (const key of ROW_STATE_EXPORT_KEYS) {
				payload[key] = null;
			}
			return payload;
		}
		for (const key of ROW_STATE_EXPORT_KEYS) {
			payload[key] = row[key] ?? null;
		}
		return payload;
	}

	function stringifyRowStateForExport(row) {
		try {
			const payload = buildRowStateExportPayload(row);
			return JSON.stringify(payload);
		} catch {
			return "";
		}
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
		row.dataset.formId = String(id);
		const infoFieldControllers = [];
		function refreshRowVisuals() {
			row.classList.remove(
				"state-dist-yes",
				"state-back-yes",
				"state-have-yes",
				"state-dist-and-back",
				"state-missed",
				"state-ignored",
				"state-deleted",
				"state-missing",
				"state-online-uploaded"
			);
			if (saved.deleted || saved.missing) {
				row.classList.add("state-ignored");
				if (saved.deleted) row.classList.add("state-deleted");
				if (saved.missing) row.classList.add("state-missing");
				return;
			}
			if (saved.onlineUploaded) {
				row.classList.add("state-online-uploaded");
				return;
			}
			if (saved.missedForm) {
				row.classList.add("state-missed");
			} else if (saved.haveForm) {
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
		const numValue = document.createElement("span");
		numValue.className = "col-num__value";
		numValue.textContent = String(id);
		const onlineToggle = document.createElement("button");
		onlineToggle.type = "button";
		onlineToggle.className = "online-upload-toggle";
		onlineToggle.setAttribute("aria-label", `Mark form ${id} as uploaded online`);
		const onlineTick = document.createElement("span");
		onlineTick.className = "online-upload-icon";
		onlineTick.textContent = "✔";
		onlineToggle.appendChild(onlineTick);
		const onlineLabel = document.createElement("span");
		onlineLabel.className = "online-upload-label";
		onlineLabel.textContent = "uploaded online";
		numCol.appendChild(numValue);
		numCol.appendChild(onlineToggle);
		numCol.appendChild(onlineLabel);

		// Actions column
		const actionsCol = document.createElement("div");
		actionsCol.className = "col-actions";
		const actionsWrap = document.createElement("div");
		actionsWrap.className = "actions";
		const infoBar = document.createElement("div");
		infoBar.className = "row-info-bar";
		const infoFields = document.createElement("div");
		infoFields.className = "row-info-fields";
		infoBar.appendChild(infoFields);

		const createInfoField = (fieldKey, config) => {
			const fieldContainer = document.createElement("div");
			fieldContainer.className = "row-info-field";

			const triggerButton = document.createElement("button");
			triggerButton.type = "button";
			triggerButton.className = "row-info-add";
			triggerButton.textContent = config.addLabel;
			triggerButton.setAttribute("aria-haspopup", "true");
			triggerButton.setAttribute("aria-expanded", "false");

			const displayWrap = document.createElement("div");
			displayWrap.className = "row-info-display";
			let labelSpan = null;
			if (config.showLabel !== false) {
				labelSpan = document.createElement("span");
				labelSpan.className = "row-info-label";
				labelSpan.textContent = `${config.label}:`;
				displayWrap.appendChild(labelSpan);
			}
			const valueSpan = document.createElement("span");
			valueSpan.className = "row-info-value";
			displayWrap.appendChild(valueSpan);

			const inputWrap = document.createElement("div");
			inputWrap.className = "row-info-input";
			const input = document.createElement("input");
			input.type = config.type || "text";
			input.placeholder = config.placeholder;
			input.autocomplete = config.autocomplete || "off";
			input.enterKeyHint = "done";
			if (config.inputMode) {
				input.setAttribute("inputmode", config.inputMode);
			}
			if (config.pattern) {
				input.setAttribute("pattern", config.pattern);
			}
			const doneButton = document.createElement("button");
			doneButton.type = "button";
			doneButton.className = "row-info-done";
			doneButton.textContent = "Done";
			inputWrap.appendChild(input);
			inputWrap.appendChild(doneButton);

			const setMode = (mode) => {
				if (saved.deleted || saved.missing) {
					triggerButton.style.display = "none";
					displayWrap.style.display = "none";
					inputWrap.style.display = "none";
					return;
				}
				triggerButton.style.display = mode === "add" ? "" : "none";
				displayWrap.style.display = mode === "display" ? "" : "none";
				inputWrap.style.display = mode === "edit" ? "" : "none";
				triggerButton.setAttribute("aria-expanded", mode === "edit" ? "true" : "false");
			};

			const getCurrentValue = () => {
				const value = saved[fieldKey];
				return typeof value === "string" ? value.trim() : "";
			};

			const refresh = () => {
				const current = getCurrentValue();
				if (current) {
					valueSpan.textContent = current;
					setMode("display");
				} else {
					valueSpan.textContent = "";
					input.value = "";
					setMode("add");
				}
			};

			const openEditor = () => {
				input.value = getCurrentValue();
				setMode("edit");
				requestAnimationFrame(() => {
					input.focus();
					const len = input.value.length;
					if (typeof input.setSelectionRange === "function") {
						try {
							input.setSelectionRange(len, len);
						} catch {
							// ignore
						}
					}
				});
			};

			const handleSubmit = () => {
				const currentValue = getCurrentValue();
				const hadValue = Boolean(currentValue);
				const nextValue = input.value.trim();
				if (!nextValue) {
					saved[fieldKey] = "";
					saveRowState(id, saved);
					refresh();
					if (hadValue) {
						scheduleFormStatsUpdate();
					}
					return;
				}
				saved[fieldKey] = nextValue;
				valueSpan.textContent = nextValue;
				saveRowState(id, saved);
				setMode("display");
				if (!hadValue) {
					scheduleFormStatsUpdate();
				}
			};

			const handleCancel = () => {
				const current = getCurrentValue();
				if (current) {
					valueSpan.textContent = current;
					setMode("display");
				} else {
					refresh();
				}
			};

			triggerButton.addEventListener("click", (event) => {
				event.preventDefault();
				openEditor();
			});

			displayWrap.addEventListener("click", (event) => {
				event.preventDefault();
				openEditor();
			});

			doneButton.addEventListener("click", (event) => {
				event.preventDefault();
				handleSubmit();
			});

			input.addEventListener("keydown", (event) => {
				if (event.key === "Enter") {
					event.preventDefault();
					handleSubmit();
				} else if (event.key === "Escape") {
					event.preventDefault();
					handleCancel();
				}
			});

			fieldContainer.appendChild(triggerButton);
			fieldContainer.appendChild(displayWrap);
			fieldContainer.appendChild(inputWrap);
			infoFields.appendChild(fieldContainer);

			refresh();
			infoFieldControllers.push({ refresh });
		};

		createInfoField("name", {
			label: "Name",
			addLabel: "Name",
			placeholder: "Enter name",
			type: "text",
			autocomplete: "name",
			showLabel: false
		});

		createInfoField("phone", {
			label: "Phone #",
			addLabel: "Phone #",
			placeholder: "Enter phone number",
			type: "tel",
			inputMode: "numeric",
			pattern: "[0-9]*",
			autocomplete: "tel"
		});

		createInfoField("epic", {
			label: "EPIC #",
			addLabel: "EPIC #",
			placeholder: "Enter EPIC number",
			type: "tel",
			inputMode: "numeric",
			pattern: "[0-9]*",
			autocomplete: "off"
		});

		if (!saved.deleted && !saved.missing) {
			row.appendChild(infoBar);
		}
		const btnDist = document.createElement("button");
		btnDist.className = saved.distributed ? "btn-dist" : "btn-muted";
		btnDist.textContent = "Distributed form";
		const btnHave = document.createElement("button");
		btnHave.className = saved.haveForm ? "btn-have" : "btn-muted";
		btnHave.textContent = "I have form";
		const btnBack = document.createElement("button");
		btnBack.className = saved.takenBack ? "btn-back" : "btn-muted";
		btnBack.textContent = "Collected form";
		const btnMissed = document.createElement("button");
		btnMissed.className = saved.missedForm ? "btn-missed" : "btn-muted";
		btnMissed.innerHTML = "Didn't get<br>back form";
		actionsWrap.appendChild(btnDist);
		actionsWrap.appendChild(btnHave);
		actionsWrap.appendChild(btnBack);
		actionsWrap.appendChild(btnMissed);
		const distributedMeta = document.createElement("div");
		distributedMeta.className = "distributed-meta";
		const distributedLabel = document.createElement("span");
		distributedLabel.className = "distributed-label";
		distributedLabel.textContent = "Distributed:";
		const distributedSelect = document.createElement("select");
		distributedSelect.className = "distributed-select";
		distributedSelect.setAttribute("aria-label", "Distributed assignment");
		const distSelf = document.createElement("option");
		distSelf.value = "self";
		distSelf.textContent = "Self Taken";
		const distGiven = document.createElement("option");
		distGiven.value = "given";
		distGiven.textContent = "Given to";
		distributedSelect.appendChild(distSelf);
		distributedSelect.appendChild(distGiven);
		const normalizedDistributionType = saved.distributionType === "given" ? "given" : "self";
		saved.distributionType = normalizedDistributionType;
		distributedSelect.value = normalizedDistributionType;
		distributedMeta.appendChild(distributedLabel);
		distributedMeta.appendChild(distributedSelect);
		actionsWrap.appendChild(distributedMeta);
		const syncDistributedMetaVisibility = () => {
			if (!saved.distributed) {
				distributedMeta.style.display = "none";
				distributedSelect.value = "self";
				return;
			}
			distributedMeta.style.display = "";
			distributedSelect.value = saved.distributionType === "given" ? "given" : "self";
		};
		syncDistributedMetaVisibility();
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
		commentsCol.className = "col-comments comment-inline";
		const commentControls = document.createElement("div");
		commentControls.className = "comment-inline-controls";
		const commentToggle = document.createElement("button");
		commentToggle.type = "button";
		commentToggle.className = "comment-inline-toggle";
		commentToggle.setAttribute("aria-expanded", "false");
		commentControls.appendChild(commentToggle);
		const btnDeleted = document.createElement("button");
		btnDeleted.type = "button";
		btnDeleted.className = "comment-flag comment-flag-deleted";
		btnDeleted.textContent = "DELETED";
		commentControls.appendChild(btnDeleted);
		const btnMissing = document.createElement("button");
		btnMissing.type = "button";
		btnMissing.className = "comment-flag comment-flag-missing";
		btnMissing.textContent = "FORM MISSING";
		commentControls.appendChild(btnMissing);
		commentsCol.appendChild(commentControls);
		const textarea = document.createElement("textarea");
		textarea.className = "comments";
		textarea.placeholder = "Comments (optional)";
		textarea.value = saved.comments || "";
		textarea.setAttribute("aria-label", "Comment");
		commentsCol.appendChild(textarea);
		const focusCommentInput = () => {
			textarea.focus();
			const length = textarea.value.length;
			if (typeof textarea.setSelectionRange === "function") {
				try {
					textarea.setSelectionRange(length, length);
				} catch {
					// Ignore if the browser does not support setSelectionRange on textarea
				}
			}
		};
		const refreshCommentToggleLabel = () => {
			commentToggle.textContent = "Comment";
		};
		const updateCommentVisibility = (visible, { persist = true, focus = false } = {}) => {
			const previous = Boolean(saved.showComments);
			const nextVisible = Boolean(visible);
			saved.showComments = nextVisible;
			commentsCol.classList.toggle("comment-inline--collapsed", !nextVisible);
			commentToggle.setAttribute("aria-pressed", nextVisible ? "true" : "false");
			refreshCommentToggleLabel();
			if (nextVisible && focus) {
				requestAnimationFrame(focusCommentInput);
			}
			if (persist) {
				saveRowState(id, saved);
				if (previous !== nextVisible) {
					scheduleFormStatsUpdate();
				}
			}
		};
		const highlightComment = ({ focus = false } = {}) => {
			updateCommentVisibility(true, { persist: false, focus });
			commentsCol.classList.add("comment-inline--active");
			setTimeout(() => {
				commentsCol.classList.remove("comment-inline--active");
			}, 600);
		};
		const initialCommentVisible =
			Boolean(saved.showComments) || Boolean(saved.comments && saved.comments.trim().length > 0);
		updateCommentVisibility(initialCommentVisible, { persist: false });
		refreshCommentToggleLabel();
		commentToggle.addEventListener("click", () => {
			const nextVisible = commentToggle.getAttribute("aria-pressed") !== "true";
			updateCommentVisibility(nextVisible, { focus: nextVisible });
		});

		const ignoredInfoCol = document.createElement("div");
		ignoredInfoCol.className = "col-ignored-info";

		const applyOnlineUploadedVisuals = () => {
			const active = Boolean(saved.onlineUploaded);
			const controlsDisabled = Boolean(saved.deleted || saved.missing);
			row.classList.toggle("row-online-uploaded", active);
			onlineToggle.classList.toggle("online-upload-toggle--active", active);
			onlineTick.classList.toggle("online-upload-icon--visible", active);
			onlineToggle.disabled = controlsDisabled;
			onlineToggle.setAttribute("aria-pressed", active ? "true" : "false");
			onlineToggle.setAttribute(
				"aria-label",
				active ? `Unmark form ${id} as uploaded online` : `Mark form ${id} as uploaded online`
			);
			onlineToggle.setAttribute(
				"title",
				active
					? "Marked as uploaded online. Click to undo."
					: "Click to mark this form as uploaded online."
			);
			onlineLabel.style.display = active ? "" : "none";
			const hideControls = controlsDisabled || active;
			if (infoBar) {
				infoBar.style.display = hideControls ? "none" : "";
			}
			actionsCol.style.display = hideControls ? "none" : "";
			commentsCol.style.display = hideControls ? "none" : "";
			if (hideControls) {
				status1Col.style.display = "none";
				status2Col.style.display = "none";
				addSecondBtn.style.display = "none";
				updateCommentVisibility(false, { persist: false });
			} else {
				status1Col.style.display = saved.haveForm ? "" : "none";
				const showSecond = Boolean(saved.haveForm && saved.showSecondStatus);
				status2Col.style.display = showSecond ? "" : "none";
				addSecondBtn.style.display = saved.haveForm && !showSecond ? "" : "none";
				const shouldShowComments =
					Boolean(saved.comments && saved.comments.trim().length > 0) || Boolean(saved.showComments);
				updateCommentVisibility(shouldShowComments, { persist: false });
			}
			recalcHeaderStatusVisibility();
		};

		const renderIgnoredBadge = () => {
			ignoredInfoCol.textContent = "";
			const label = saved.deleted ? "DELETED" : saved.missing ? "FORM MISSING" : "";
			if (!label) return;
			const badge = document.createElement("button");
			badge.type = "button";
			badge.className = "ignored-pill";
			if (saved.deleted) badge.classList.add("ignored-pill--deleted");
			if (saved.missing) badge.classList.add("ignored-pill--missing");
			badge.textContent = label;
			badge.title = "Click to restore this form";
			badge.addEventListener("click", () => {
				if (saved.deleted) {
					toggleIgnoredState("deleted");
				} else if (saved.missing) {
					toggleIgnoredState("missing");
				}
			});
			ignoredInfoCol.appendChild(badge);
		};

		const applyIgnoredState = () => {
			const ignored = saved.deleted || saved.missing;
			row.classList.toggle("row-ignored", ignored);
			btnDeleted.classList.toggle("active", saved.deleted);
			btnMissing.classList.toggle("active", saved.missing);
			btnDeleted.setAttribute("aria-pressed", saved.deleted ? "true" : "false");
			btnMissing.setAttribute("aria-pressed", saved.missing ? "true" : "false");
			renderIgnoredBadge();
			ignoredInfoCol.style.display = ignored ? "" : "none";
			if (infoBar) {
				infoBar.style.display = ignored ? "none" : "";
			}
			if (!ignored) {
				for (const controller of infoFieldControllers) {
					if (controller && typeof controller.refresh === "function") {
						controller.refresh();
					}
				}
			}
			applyOnlineUploadedVisuals();
		};

		const toggleIgnoredState = (mode) => {
			const activating = mode === "deleted" ? !saved.deleted : !saved.missing;
			if (mode === "deleted") {
				saved.deleted = activating;
				if (activating) saved.missing = false;
			} else {
				saved.missing = activating;
				if (activating) saved.deleted = false;
			}
			if (saved.deleted || saved.missing) {
				saved.distributed = false;
				saved.takenBack = false;
				saved.haveForm = false;
				saved.missedForm = false;
				saved.onlineUploaded = false;
				saved.showSecondStatus = false;
				saved.status1 = "";
				saved.status2 = "";
				saved.showComments = false;
				btnDist.className = "btn-muted";
				btnBack.className = "btn-muted";
				btnHave.className = "btn-muted";
				btnMissed.className = "btn-muted";
			} else {
				btnDist.className = saved.distributed ? "btn-dist" : "btn-muted";
				btnBack.className = saved.takenBack ? "btn-back" : "btn-muted";
				btnHave.className = saved.haveForm ? "btn-have" : "btn-muted";
				btnMissed.className = saved.missedForm ? "btn-missed" : "btn-muted";
			}
			refreshRowVisuals();
			syncDistributedMetaVisibility();
			applyIgnoredState();
			saveRowState(id, saved);
			scheduleFormStatsUpdate();
			scheduleGroupStatsUpdate(id);
			recalcHeaderStatusVisibility();
		};

		onlineToggle.addEventListener("click", async () => {
			if (saved.deleted || saved.missing) return;
			const next = !saved.onlineUploaded;
			if (!next) {
				const proceed = await showConfirm("You're marking this form as not uploaded online");
				if (!proceed) return;
			}
			saved.onlineUploaded = next;
			saveRowState(id, saved);
			refreshRowVisuals();
			applyOnlineUploadedVisuals();
			if (!next) {
				for (const controller of infoFieldControllers) {
					if (controller && typeof controller.refresh === "function") {
						controller.refresh();
					}
				}
			}
			scheduleFormStatsUpdate();
			scheduleGroupStatsUpdate(id);
		});

		btnDeleted.addEventListener("click", () => toggleIgnoredState("deleted"));
		btnMissing.addEventListener("click", () => toggleIgnoredState("missing"));

		// Handlers
		btnDist.addEventListener("click", async () => {
			if (saved.haveForm) {
				const confirmMessage = "You're changing from 'I have form' to 'Distributed form'. Continue?";
				const proceed = await showConfirm(confirmMessage);
				if (!proceed) return;
				saved.haveForm = false;
				saved.distributed = true;
				saved.takenBack = false;
				saved.missedForm = false;
				saved.distributionType = "self";
				btnHave.className = "btn-muted";
				btnBack.className = "btn-muted";
				btnMissed.className = "btn-muted";
				status1Col.style.display = "none";
				status2Col.style.display = "none";
				addSecondBtn.style.display = "none";
				updateCommentVisibility(false, { persist: false });
				btnDist.className = "btn-dist";
				syncDistributedMetaVisibility();
				refreshRowVisuals();
				saveRowState(id, saved);
				scheduleFormStatsUpdate();
				scheduleGroupStatsUpdate(id);
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
			saved.distributionType = "self";
			if (!next) {
				saved.missedForm = false;
				btnMissed.className = "btn-muted";
			}
			btnDist.className = next ? "btn-dist" : "btn-muted";
			syncDistributedMetaVisibility();
			refreshRowVisuals();
			saveRowState(id, saved);
			scheduleFormStatsUpdate();
			scheduleGroupStatsUpdate(id);
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
				const confirmMessage = "You're changing from 'Collected form' to 'I have form'. Continue?";
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
				saved.missedForm = false;
				saved.distributionType = "self";
				btnDist.className = "btn-muted";
				btnBack.className = "btn-muted";
				btnMissed.className = "btn-muted";
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
				btnMissed.className = saved.missedForm ? "btn-missed" : "btn-muted";
				refreshRowVisuals();
			}
			syncDistributedMetaVisibility();
			saveRowState(id, saved);
			scheduleFormStatsUpdate();
			scheduleGroupStatsUpdate(id);
			recalcHeaderStatusVisibility();
		});

		btnBack.addEventListener("click", async () => {
			if (saved.haveForm) return; // disabled when have form
			if (saved.takenBack) {
				const confirmMessage = "You're changing from 'Collected form' to 'Not collected'. Continue?";
				const proceed = await showConfirm(confirmMessage);
				if (!proceed) return;
			} else if (saved.missedForm) {
				const proceed = await showConfirm(
					"This form is marked as 'Didn't get back form'. Mark it as 'Collected form' instead?"
				);
				if (!proceed) return;
			}
			const next = !saved.takenBack;
			saved.takenBack = next;
			if (next) {
				saved.missedForm = false;
				btnMissed.className = "btn-muted";
			} else if (saved.missedForm) {
				btnMissed.className = "btn-missed";
			}
			btnBack.className = next ? "btn-back" : "btn-muted";
			refreshRowVisuals();
			saveRowState(id, saved);
			scheduleFormStatsUpdate();
			scheduleGroupStatsUpdate(id);
		});

		btnMissed.addEventListener("click", async () => {
			const next = !saved.missedForm;
			if (next) {
				if (saved.haveForm) {
					const proceed = await showConfirm("Marking 'Didn't get back form' will clear 'I have form'. Continue?");
					if (!proceed) return;
				}
				if (saved.takenBack) {
					const proceed = await showConfirm("Marking 'Didn't get back form' will clear 'Collected form'. Continue?");
					if (!proceed) return;
				}
		} else {
			const proceed = await showConfirm("You're unchecking 'Didn't get back form'. Continue?");
			if (!proceed) return;
		}
			saved.missedForm = next;
			if (next) {
				saved.haveForm = false;
				saved.takenBack = false;
				saved.distributed = true;
				saved.distributionType = "self";
				btnHave.className = "btn-muted";
				btnBack.className = "btn-muted";
				btnDist.className = "btn-dist";
				status1Col.style.display = "none";
				status2Col.style.display = "none";
				addSecondBtn.style.display = "none";
				highlightComment({ focus: true });
				recalcHeaderStatusVisibility();
			} else {
				const keepOpen = Boolean(saved.comments && saved.comments.trim().length > 0);
				updateCommentVisibility(keepOpen, { persist: false });
				btnBack.className = saved.takenBack ? "btn-back" : "btn-muted";
				btnDist.className = saved.distributed ? "btn-dist" : "btn-muted";
			}
			btnMissed.className = next ? "btn-missed" : "btn-muted";
			syncDistributedMetaVisibility();
			refreshRowVisuals();
			saveRowState(id, saved);
			recalcHeaderStatusVisibility();
			scheduleFormStatsUpdate();
			scheduleGroupStatsUpdate(id);
		});

		distributedSelect.addEventListener("change", () => {
			const nextValue = distributedSelect.value === "given" ? "given" : "self";
			saved.distributionType = nextValue;
			if (nextValue === "given") {
				highlightComment({ focus: true });
			} else if (!saved.comments || !saved.comments.trim().length) {
				updateCommentVisibility(false, { persist: false });
			}
			saveRowState(id, saved);
		});

		select1.addEventListener("change", () => {
			const hadStatus = Boolean(saved.status1 && saved.status1.trim());
			saved.status1 = select1.value;
			const hasStatus = Boolean(saved.status1 && saved.status1.trim());
			saveRowState(id, saved);
			if (hadStatus !== hasStatus) {
				scheduleFormStatsUpdate();
			}
		});
		select2.addEventListener("change", () => {
			const hadStatus = Boolean(saved.status2 && saved.status2.trim());
			saved.status2 = select2.value;
			const hasStatus = Boolean(saved.status2 && saved.status2.trim());
			saveRowState(id, saved);
			if (hadStatus !== hasStatus) {
				scheduleFormStatsUpdate();
			}
		});
		addSecondBtn.addEventListener("click", () => {
			saved.showSecondStatus = true;
			status2Col.style.display = "";
			addSecondBtn.style.display = "none";
			saveRowState(id, saved);
			select2.focus();
			recalcHeaderStatusVisibility();
			scheduleFormStatsUpdate();
		});
		textarea.addEventListener("input", () => {
			const hadComment = Boolean(saved.comments && saved.comments.trim());
			saved.comments = textarea.value;
			const hasComment = Boolean(saved.comments && saved.comments.trim());
			const isVisible = !commentsCol.classList.contains("comment-inline--collapsed");
			refreshCommentToggleLabel(isVisible);
			saveRowState(id, saved);
			if (hadComment !== hasComment) {
				scheduleFormStatsUpdate();
			}
		});

		applyIgnoredState();

		// Assemble
		row.appendChild(numCol);
		row.appendChild(actionsCol);
		row.appendChild(status1Col);
		row.appendChild(status2Col);
		row.appendChild(commentsCol);
		row.appendChild(ignoredInfoCol);

		return row;
	}

	function waitForFrame() {
		return new Promise((resolve) => requestAnimationFrame(resolve));
	}

	function queueSearchBarHeightSync() {
		if (!els.searchPanel) return;
		if (searchHeightRaf != null) return;
		searchHeightRaf = requestAnimationFrame(() => {
			searchHeightRaf = null;
			if (!els.searchPanel) return;
			const searchHeight = els.searchPanel.offsetHeight || 0;
			const topShellHeight = els.topShell ? els.topShell.offsetHeight || 0 : searchHeight;
			document.documentElement.style.setProperty("--search-bar-height", `${searchHeight}px`);
			document.documentElement.style.setProperty("--top-shell-height", `${topShellHeight}px`);
			document.documentElement.style.setProperty("--sticky-offset", `${topShellHeight}px`);
			if (document.body) {
				document.body.style.paddingTop = `${topShellHeight}px`;
			}
		});
	}

	function setSearchStatus(message) {
		if (els.searchStatus) {
			els.searchStatus.textContent = message || "";
		}
		queueSearchBarHeightSync();
	}

	function clearSearchHighlight() {
		if (lastSearchRow) {
			lastSearchRow.classList.remove("row-search-highlight");
			lastSearchRow = null;
		}
	}

	function updateSearchNavigation() {
		if (!els.searchPrev || !els.searchNext || !els.searchDone) return;
		if (!searchState || !Array.isArray(searchState.matches) || searchState.matches.length === 0) {
			els.searchPrev.classList.add("hidden");
			els.searchNext.classList.add("hidden");
			els.searchDone.classList.add("hidden");
			els.searchPrev.disabled = true;
			els.searchNext.disabled = true;
			els.searchDone.disabled = false;
			queueSearchBarHeightSync();
			return;
		}
		const { matches, index } = searchState;
		const total = matches.length;
		const hasMultiple = total > 1;
		els.searchPrev.classList.toggle("hidden", !hasMultiple);
		els.searchNext.classList.toggle("hidden", !hasMultiple);
		els.searchDone.classList.remove("hidden");
		els.searchPrev.disabled = index <= 0;
		els.searchNext.disabled = index >= total - 1;
		els.searchDone.disabled = false;
		queueSearchBarHeightSync();
	}

	async function revealRowElement(formNumber) {
		if (!Number.isFinite(formNumber)) return null;
		const selector = `.row[data-form-id="${formNumber}"]`;
		let row = els.list.querySelector(selector);
		if (row) {
			const container = row.closest(".society");
			if (container && !container.open) {
				container.open = true;
				await waitForFrame();
				queueSearchBarHeightSync();
			}
			return row;
		}
		const groups = getGroupsSorted();
		const group = findGroupForForm(formNumber, groups);
		const targetGroupId = (group ? group.id : "ungrouped") || "ungrouped";
		let details = els.list.querySelector(`.society[data-group-id="${targetGroupId}"]`);
		const rendererEntry = groupRenderers.get(targetGroupId);
		if (!details && rendererEntry && rendererEntry.details) {
			details = rendererEntry.details;
		}
		if (rendererEntry && rendererEntry.body && rendererEntry.body.childElementCount === 0) {
			rendererEntry.render();
		}
		if (!details) return null;
		if (!details.open) {
			details.open = true;
			queueSearchBarHeightSync();
		}
		await waitForFrame();
		row = els.list.querySelector(selector);
		if (!row) {
			if (rendererEntry) {
				rendererEntry.render();
			}
			await waitForFrame();
			row = els.list.querySelector(selector);
		}
		if (row) {
			const container = row.closest(".society");
			if (container && !container.open) {
				container.open = true;
				await waitForFrame();
			}
		}
		queueSearchBarHeightSync();
		return row || null;
	}

	function formatSearchStatus(index, total, formNumber, query, mode) {
		if (total <= 1) {
			if (mode === "number") {
				return `Form ${formNumber} found.`;
			}
			return `Found form ${formNumber} for "${query}".`;
		}
		return `Result ${index + 1} of ${total} for "${query}" • Form ${formNumber}`;
	}

	async function focusSearchResult(index) {
		if (!searchState || !Array.isArray(searchState.matches) || searchState.matches.length === 0) return;
		if (index < 0 || index >= searchState.matches.length) return;
		searchState.index = index;
		updateSearchNavigation();
		const formNumber = searchState.matches[index];
		const row = await revealRowElement(formNumber);
		if (!row) {
			setSearchStatus(`Unable to locate form ${formNumber}.`);
			return;
		}
		if (lastSearchRow && lastSearchRow !== row) {
			lastSearchRow.classList.remove("row-search-highlight");
		}
		row.classList.add("row-search-highlight");
		lastSearchRow = row;
		const topShellHeight = els.topShell ? els.topShell.offsetHeight || 0 : 0;
		const stickyOffset = topShellHeight + 16;
		const targetTop = row.getBoundingClientRect().top + window.scrollY - stickyOffset;
		window.scrollTo({
			top: targetTop < 0 ? 0 : targetTop,
			behavior: "smooth"
		});
		setSearchStatus(
			formatSearchStatus(index, searchState.matches.length, formNumber, searchState.query, searchState.mode)
		);
	}

	function collectFieldMatches(query) {
		const normalized = query.trim().toLowerCase();
		const count = getSavedCount();
		if (!normalized || !Number.isFinite(count) || count <= 0) {
			return [];
		}
		const matches = [];
		for (let formNumber = 1; formNumber <= count; formNumber++) {
			const state = getRowState(formNumber);
			if (state.deleted || state.missing) continue;
			const nameMatches = state.name && state.name.toLowerCase().includes(normalized);
			const phoneMatches = state.phone && state.phone.toLowerCase().includes(normalized);
			const epicMatches = state.epic && state.epic.toLowerCase().includes(normalized);
			if (nameMatches || phoneMatches || epicMatches) {
				matches.push(formNumber);
			}
		}
		return matches;
	}

	async function performSearch() {
		if (!els.searchInput) return;
		const rawQuery = els.searchInput.value || "";
		const query = rawQuery.trim();
		if (!query) {
			setSearchStatus("");
			searchState = null;
			updateSearchNavigation();
			clearSearchHighlight();
			return;
		}
		const count = getSavedCount();
		if (!Number.isFinite(count) || count <= 0) {
			setSearchStatus("Generate forms before searching.");
			searchState = null;
			updateSearchNavigation();
			clearSearchHighlight();
			return;
		}
		const matchesSet = new Set();
		let mode = "text";
		if (/^\d+$/.test(query)) {
			mode = "number";
			const numberMatch = Number.parseInt(query, 10);
			if (Number.isFinite(numberMatch) && numberMatch >= 1 && numberMatch <= count) {
				matchesSet.add(numberMatch);
			}
		}
		for (const formNumber of collectFieldMatches(query)) {
			matchesSet.add(formNumber);
		}
		const matches = Array.from(matchesSet).sort((a, b) => a - b);
		if (!matches.length) {
			setSearchStatus(`No matches for "${query}".`);
			searchState = null;
			updateSearchNavigation();
			clearSearchHighlight();
			return;
		}
		clearSearchHighlight();
		searchState = {
			query,
			mode,
			matches,
			index: 0
		};
		updateSearchNavigation();
		await focusSearchResult(0);
	}

	function completeSearchSession() {
		if (!searchState || !Array.isArray(searchState.matches) || searchState.matches.length === 0) {
			searchState = null;
			updateSearchNavigation();
			return;
		}
		const formNumber = searchState.matches[searchState.index] || searchState.matches[0];
		setSearchStatus(
			`Search paused at form ${formNumber} (${searchState.index + 1} of ${searchState.matches.length}) for "${searchState.query}".`
		);
		searchState = null;
		updateSearchNavigation();
	}

	function buildRowSummary(row) {
		const summary = {
			title: "-",
			details: []
		};
		const hasComments = typeof row.comments === "string" && row.comments.trim().length > 0;
		const trimmedComment = hasComments ? row.comments.trim() : "";
		const personalDetails = [];
		const appendIfPresent = (label, value) => {
			if (typeof value !== "string") return;
			const trimmed = value.trim();
			if (!trimmed) return;
			personalDetails.push(`${label}: ${trimmed}`);
		};
		appendIfPresent("Name", row.name);
		appendIfPresent("Phone #", row.phone);
		appendIfPresent("EPIC #", row.epic);

		if (row.onlineUploaded) {
			summary.title = "Online uploaded";
			if (hasComments) summary.details.push(trimmedComment);
		} else if (row.deleted) {
			summary.title = "Deleted";
			if (hasComments) summary.details.push(trimmedComment);
		} else if (row.missing) {
			summary.title = "Form missing";
			if (hasComments) summary.details.push(trimmedComment);
		} else if (row.haveForm) {
			summary.title = "I have form";
			const statuses = [row.status1, row.status2].map((s) => (typeof s === "string" ? s.trim() : "")).filter(Boolean);
			if (statuses.length) summary.details.push(statuses.join(", "));
			if (hasComments) summary.details.push(trimmedComment);
		} else if (row.missedForm) {
			summary.title = "Not returned";
			if (hasComments) summary.details.push(trimmedComment);
		} else if (row.distributed && row.takenBack) {
			summary.title = "Collected";
			if (row.distributionType === "given") summary.details.push("Given to");
			if (hasComments) summary.details.push(trimmedComment);
		} else if (row.distributed) {
			summary.title = "Distributed";
			if (row.distributionType === "given") summary.details.push("Given to");
			if (hasComments) summary.details.push(trimmedComment);
		} else if (row.takenBack) {
			summary.title = "Collected form";
			if (hasComments) summary.details.push(trimmedComment);
		} else if (hasComments) {
			summary.details.push(trimmedComment);
		}

		if (personalDetails.length) {
			summary.details = [...personalDetails, ...summary.details];
		}

		return summary;
	}

	function computeRowStats(rows) {
		const totals = {
			distributed: 0,
			takenBack: 0,
			haveForm: 0,
			missed: 0,
			onlineUploaded: 0
		};
		if (!Array.isArray(rows)) {
			return totals;
		}
		for (const row of rows) {
			if (!row || typeof row !== "object") continue;
			const isUploaded = Boolean(row.onlineUploaded);
			if (isUploaded) totals.onlineUploaded += 1;
			if (isUploaded) continue;
			if (row.distributed) totals.distributed += 1;
			if (row.takenBack) totals.takenBack += 1;
			if (row.haveForm) totals.haveForm += 1;
			if (row.missedForm) totals.missed += 1;
		}
		return totals;
	}

	function computeGroupStats(formNumbers) {
		const stats = {
			distributed: 0,
			have: 0,
			collected: 0,
			uploaded: 0
		};
		if (!Array.isArray(formNumbers)) {
			return stats;
		}
		for (const formNumber of formNumbers) {
			if (!Number.isFinite(formNumber)) continue;
			const state = getRowState(formNumber);
			if (!state || typeof state !== "object") continue;
			if (state.deleted || state.missing) continue;
			const isUploaded = Boolean(state.onlineUploaded);
			if (isUploaded) {
				stats.uploaded += 1;
				continue;
			}
			if (state.distributed) stats.distributed += 1;
			if (state.haveForm) stats.have += 1;
			if (state.takenBack) stats.collected += 1;
		}
		return stats;
	}

	function isRowConsideredNoEntry(state) {
		if (!state || typeof state !== "object") return true;
		const hasPrimaryAction =
			Boolean(state.distributed) ||
			Boolean(state.takenBack) ||
			Boolean(state.haveForm) ||
			Boolean(state.missedForm) ||
			Boolean(state.onlineUploaded);
		const hasDetails =
			Boolean(state.status1 && state.status1.trim()) ||
			Boolean(state.status2 && state.status2.trim()) ||
			Boolean(state.comments && state.comments.trim()) ||
			Boolean(state.name && state.name.trim()) ||
			Boolean(state.phone && state.phone.trim()) ||
			Boolean(state.epic && state.epic.trim()) ||
			Boolean(state.showSecondStatus) ||
			Boolean(state.showComments);
		return !hasPrimaryAction && !hasDetails;
	}

	function computeFormStatsTotals() {
		const savedCount = getSavedCount();
		const total = Number.isFinite(savedCount) && savedCount > 0 ? savedCount : 0;
		const stats = {
			total,
			distributed: 0,
			collected: 0,
			have: 0,
			uploaded: 0,
			missing: 0,
			deleted: 0,
			noEntry: 0
		};
		if (total === 0) {
			return stats;
		}
		for (let id = 1; id <= total; id++) {
			const state = getRowState(id);
			if (!state || typeof state !== "object") continue;
			const isDeleted = Boolean(state.deleted);
			const isMissing = Boolean(state.missing);
			if (isDeleted) stats.deleted += 1;
			if (isMissing) stats.missing += 1;
			if (isDeleted || isMissing) continue;
			const isUploaded = Boolean(state.onlineUploaded);
			if (isUploaded) {
				stats.uploaded += 1;
				continue;
			}
			if (state.distributed) stats.distributed += 1;
			if (state.takenBack) stats.collected += 1;
			if (state.haveForm) stats.have += 1;
			if (isRowConsideredNoEntry(state)) {
				stats.noEntry += 1;
			}
		}
		return stats;
	}

	function updateFormStats() {
		if (!els.formStats) return;
		const stats = computeFormStatsTotals();
		if (els.statsDistributed) {
			els.statsDistributed.textContent = String(stats.distributed);
		}
		if (els.statsCollected) {
			els.statsCollected.textContent = String(stats.collected);
		}
		if (els.statsHave) {
			els.statsHave.textContent = String(stats.have);
		}
		if (els.statsOnlineUploaded) {
			els.statsOnlineUploaded.textContent = String(stats.uploaded);
		}
		if (els.statsMissing) {
			els.statsMissing.textContent = String(stats.missing);
		}
		if (els.statsDeleted) {
			els.statsDeleted.textContent = String(stats.deleted);
		}
		if (els.statsNoEntry) {
			els.statsNoEntry.textContent = String(stats.noEntry);
		}
		els.formStats.setAttribute("data-total-forms", String(stats.total));
	}

	function scheduleFormStatsUpdate() {
		if (statsUpdateScheduled) return;
		statsUpdateScheduled = true;
		const hasWindow = typeof window !== "undefined";
		const raf =
			hasWindow && typeof window.requestAnimationFrame === "function"
				? window.requestAnimationFrame.bind(window)
				: null;
		const runner =
			raf ||
			((callback) => {
				if (hasWindow && typeof window.setTimeout === "function") {
					window.setTimeout(callback, 16);
				} else if (typeof setTimeout === "function") {
					setTimeout(callback, 16);
				} else {
					callback();
				}
			});
		runner(() => {
			statsUpdateScheduled = false;
			updateFormStats();
		});
	}

	function refreshGroupStats(groupId) {
		if (!groupId) return;
		const entry = groupRenderers.get(groupId);
		if (!entry) return;
		const numbers = Array.isArray(entry.numbers) ? entry.numbers : [];
		const stats = computeGroupStats(numbers);
		if (entry.statValues && typeof entry.statValues === "object") {
			const entries = [
				["distributed", stats.distributed],
				["have", stats.have],
				["collected", stats.collected],
				["uploaded", stats.uploaded]
			];
			for (const [key, value] of entries) {
				const statInfo = entry.statValues[key];
				if (!statInfo) continue;
				if (statInfo.valueEl) {
					statInfo.valueEl.textContent = String(value);
				}
				if (statInfo.root) {
					statInfo.root.setAttribute("title", `${statInfo.label}: ${value}`);
				}
			}
		}
	}

	function scheduleGroupStatsUpdate(formNumber) {
		if (!Number.isFinite(formNumber)) return;
		const total = getSavedCount();
		if (!Number.isFinite(total) || total <= 0) return;
		if (formNumber < 1 || formNumber > total) return;
		const groups = getGroupsSorted();
		const group = findGroupForForm(formNumber, groups);
		const groupId = (group ? group.id : "ungrouped") || "ungrouped";
		pendingGroupStatUpdates.add(groupId);
		if (!groupStatsUpdateScheduled) {
			groupStatsUpdateScheduled = true;
			const hasWindow = typeof window !== "undefined";
			const raf =
				hasWindow && typeof window.requestAnimationFrame === "function"
					? window.requestAnimationFrame.bind(window)
					: null;
			const runner =
				raf ||
				((callback) => {
					if (hasWindow && typeof window.setTimeout === "function") {
						window.setTimeout(callback, 16);
					} else if (typeof setTimeout === "function") {
						setTimeout(callback, 16);
					} else {
						callback();
					}
				});
			runner(() => {
				groupStatsUpdateScheduled = false;
				const pending = Array.from(pendingGroupStatUpdates);
				pendingGroupStatUpdates.clear();
				for (const id of pending) {
					refreshGroupStats(id);
				}
			});
		}
	}

	function formatRowStatsLine(stats) {
		const totals =
			stats || { distributed: 0, takenBack: 0, haveForm: 0, missed: 0, onlineUploaded: 0 };
		return `Distributed: ${totals.distributed}  •  Not Returned: ${totals.missed}  •  Collected: ${totals.takenBack}  •  I Have: ${totals.haveForm}  •  Online Uploaded: ${totals.onlineUploaded}`;
	}

	function renderList(n) {
		renderSocietyManager();
		updateFormStats();
		els.list.textContent = "";
		if (!Number.isFinite(n) || n <= 0) {
			els.listHeader.classList.add("hidden");
			groupRenderers.clear();
			queueSearchBarHeightSync();
			return;
		}
		els.listHeader.classList.remove("hidden");
		groupRenderers.clear();

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
			const hasRows = numbers.length > 0;
			if (!hasRows && isUngrouped) return;
			const details = document.createElement("details");
			details.className = "society";
			const groupIdRaw = isUngrouped ? "ungrouped" : group.id;
			const groupId = groupIdRaw || "ungrouped";
			details.dataset.groupId = groupId;
			if (numbers.length > 0) {
				details.dataset.rangeStart = String(numbers[0]);
				details.dataset.rangeEnd = String(numbers[numbers.length - 1]);
			} else if (!isUngrouped && group) {
				details.dataset.rangeStart = String(group.start);
				details.dataset.rangeEnd = String(group.end);
			} else {
				details.dataset.rangeStart = "";
				details.dataset.rangeEnd = "";
			}
			if (getGroupOpenState(groupId)) {
				details.open = true;
			}
			const summary = document.createElement("summary");
			summary.className = "society-summary";
			const summaryMain = document.createElement("div");
			summaryMain.className = "society-summary-main";
			const nameSpan = document.createElement("span");
			nameSpan.className = "society-summary-name";
			nameSpan.textContent = isUngrouped
				? hasNamedGroups ? "Ungrouped forms" : "All forms"
				: `${group.name} (${group.start}–${group.end})`;
			const countSpan = document.createElement("span");
			countSpan.className = "society-summary-count";
			countSpan.textContent = `${numbers.length} form${numbers.length === 1 ? "" : "s"}`;
			summaryMain.appendChild(nameSpan);
			summaryMain.appendChild(countSpan);
			const statsWrap = document.createElement("div");
			statsWrap.className = "society-summary-stats";
			const statValues = {};
			const createSummaryStat = (variant, value, label) => {
				const stat = document.createElement("span");
				stat.className = "society-summary-stat";
				stat.title = `${label}: ${value}`;
				const srLabel = document.createElement("span");
				srLabel.className = "sr-only";
				srLabel.textContent = `${label}: `;
				const dot = document.createElement("span");
				dot.className = `society-summary-dot society-summary-dot--${variant}`;
				dot.setAttribute("aria-hidden", "true");
				const valueSpan = document.createElement("span");
				valueSpan.className = "society-summary-stat-count";
				valueSpan.textContent = String(value);
				stat.append(srLabel, dot, valueSpan);
				statsWrap.appendChild(stat);
				statValues[variant] = {
					valueEl: valueSpan,
					root: stat,
					label
				};
			};
			const initialStats = computeGroupStats(numbers);
			createSummaryStat("distributed", initialStats.distributed, "Distributed forms");
			createSummaryStat("have", initialStats.have, "I have forms");
			createSummaryStat("collected", initialStats.collected, "Collected forms");
			createSummaryStat("uploaded", initialStats.uploaded, "Online uploaded forms");
			summary.appendChild(summaryMain);
			summary.appendChild(statsWrap);
			details.appendChild(summary);
			const body = document.createElement("div");
			body.className = "society-body";
			const renderRowsIntoBody = () => {
				// Avoid re-rendering if already populated
				if (body.childElementCount > 0) return;
				if (!hasRows) {
					const emptyMessage = document.createElement("p");
					emptyMessage.className = "society-empty";
					emptyMessage.textContent = "No forms currently fall within this range.";
					body.appendChild(emptyMessage);
					return;
				}
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
			groupRenderers.set(groupId, {
				render: renderRowsIntoBody,
				body,
				details,
				summary,
				statValues,
				numbers: numbers.slice()
			});
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
		if (searchState && Array.isArray(searchState.matches) && searchState.matches.length > 0) {
			const safeIndex = Math.min(searchState.index, searchState.matches.length - 1);
			searchState.index = safeIndex;
			void focusSearchResult(safeIndex);
		} else {
			updateSearchNavigation();
		}
		queueSearchBarHeightSync();
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
		updateFormStats();
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
		const overallStats = computeRowStats(snapshot.rows);
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
		doc.setTextColor(0, 0, 0);
		const infoStart = titleY + 20;
		let infoY = infoStart;
		const infoLines = [
			`Generated: ${formatReadableDateTime()}`,
			`Total Forms: ${snapshot.count}`
		];
		if (snapshot.groups.length > 0) {
			infoLines.push(`Societies: ${snapshot.groups.length}`);
		}
		infoLines.push(formatRowStatsLine(overallStats));
		for (const line of infoLines) {
			doc.text(line, marginLeft, infoY);
			infoY += 18;
		}
		if (typeof doc.autoTable !== "function") {
			showImportExportMessage("PDF table helper not available. Please reload and try again.", "error");
			return;
		}
		const toSection = (title, rows) => ({
			title,
			rows,
			stats: computeRowStats(rows)
		});
		const sections = [];
		for (const entry of snapshot.groupedRows) {
			if (!entry || !entry.group || entry.rows.length === 0) continue;
			const countLabel = `${entry.rows.length} form${entry.rows.length === 1 ? "" : "s"}`;
			const title = `${entry.group.name} (${entry.group.start}–${entry.group.end}) • ${countLabel}`;
			sections.push(toSection(title, entry.rows));
		}
		if (snapshot.groups.length > 0) {
			if (snapshot.ungroupedRows.length > 0) {
				const countLabel = `${snapshot.ungroupedRows.length} form${snapshot.ungroupedRows.length === 1 ? "" : "s"}`;
				sections.push(toSection(`Ungrouped forms • ${countLabel}`, snapshot.ungroupedRows));
			}
		} else {
			const countLabel = `${snapshot.count} form${snapshot.count === 1 ? "" : "s"}`;
			sections.push(toSection(`All forms • ${countLabel}`, snapshot.rows));
		}
		let currentY = infoY + 12;
		const pageHeight = doc.internal.pageSize.getHeight();
		for (const section of sections) {
			if (!section.rows || section.rows.length === 0) continue;
			if (currentY > pageHeight - 120) {
				doc.addPage();
				currentY = 60;
			}
			doc.setFontSize(14);
			doc.setTextColor(34, 34, 34);
			doc.text(section.title, marginLeft, currentY);
			currentY += 16;
			const statsLine = formatRowStatsLine(section.stats);
			doc.setFontSize(11);
			doc.setTextColor(71, 85, 105);
			doc.text(statsLine, marginLeft, currentY);
			currentY += 14;
			doc.setFontSize(10);
			doc.setTextColor(0, 0, 0);
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
			const tableFinalY = doc.lastAutoTable && doc.lastAutoTable.finalY ? doc.lastAutoTable.finalY : currentY;
			currentY = tableFinalY + 18;
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
		const groupLookup = new Map(
			Array.isArray(snapshot.groups) ? snapshot.groups.map((group) => [group.id, group]) : []
		);
		const formsSheetData = [FORM_EXPORT_HEADERS.slice()];
		for (const row of snapshot.rows) {
			const group = row && row.groupId ? groupLookup.get(row.groupId) : null;
			const isDeleted = Boolean(row && row.deleted);
			const isMissing = Boolean(row && row.missing);
			const noEntry = !isDeleted && !isMissing && isRowConsideredNoEntry(row);
			const distributionLabel =
				row && row.distributionType === "given"
					? "Given to"
					: row && row.distributionType === "self"
					? "Self Taken"
					: safeStringForExcel(row && row.distributionType);
			const colorHex = row && row.colorHex ? row.colorHex : determineRowColorHex(row);
			formsSheetData.push([
				row.formNumber,
				row.groupId || "",
				row.groupName || (group ? group.name : ""),
				group && Number.isFinite(group.start) ? group.start : "",
				group && Number.isFinite(group.end) ? group.end : "",
				safeStringForExcel(row.name),
				safeStringForExcel(row.phone),
				safeStringForExcel(row.epic),
				formatBooleanForExcel(row.distributed),
				distributionLabel,
				formatBooleanForExcel(row.takenBack),
				formatBooleanForExcel(row.takenBack),
				formatBooleanForExcel(row.haveForm),
				formatBooleanForExcel(row.onlineUploaded),
				formatBooleanForExcel(row.missedForm),
				formatBooleanForExcel(row.missedForm),
				formatBooleanForExcel(row.deleted),
				formatBooleanForExcel(row.missing),
				formatBooleanForExcel(row.showSecondStatus),
				safeStringForExcel(row.status1),
				safeStringForExcel(row.status2),
				formatBooleanForExcel(row.showComments),
				safeStringForExcel(row.comments),
				formatBooleanForExcel(noEntry),
				safeStringForExcel(colorHex),
				stringifyRowStateForExport(row)
			]);
		}
		const formsSheet = XLSX.utils.aoa_to_sheet(formsSheetData);
		XLSX.utils.book_append_sheet(workbook, formsSheet, "Forms");

		const groupsSheetData = [GROUP_EXPORT_HEADERS.slice()];
		for (const group of snapshot.groups) {
			groupsSheetData.push([
				group && typeof group.id === "string" ? group.id : "",
				group ? group.name : "",
				group ? group.start : "",
				group ? group.end : ""
			]);
		}
		const groupsSheet = XLSX.utils.aoa_to_sheet(groupsSheetData);
		XLSX.utils.book_append_sheet(workbook, groupsSheet, "Societies");

		const settingsSheetData = [
			["Key", "Value"],
			["dataVersion", DATA_EXPORT_VERSION],
			["assetVersion", ASSET_VERSION],
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
			const versionRow = settingsRows.find(
				(entry) => coerceString(entry.key).trim().toLowerCase() === "dataversion"
			);
			const dataVersionValue =
				Number.parseInt(coerceString(versionRow ? versionRow.value : "").trim(), 10) || 1;
			const supportsRowStateJson = dataVersionValue >= 2;

			const formsSheet =
				workbook.Sheets.Forms || workbook.Sheets.forms || workbook.Sheets.FORMS || null;
			if (!formsSheet) {
				throw new Error("Forms sheet not found in file.");
			}
			const formsRows = XLSX.utils.sheet_to_json(formsSheet, { defval: "" });
			const stateMap = new Map();
			for (const row of formsRows) {
				const rawNumber = firstPopulatedValue(
					row.FormNumber,
					row.formNumber,
					row["Form #"],
					row.Number,
					row["FormNumber"]
				);
				const formNumber = Number.parseInt(rawNumber, 10);
				if (!Number.isFinite(formNumber) || formNumber < 1 || formNumber > parsedCount) continue;
				const rawStateJsonCandidate = coerceString(
					firstPopulatedValue(row.RawStateJson, row.rawStateJson, row.StateJson, row.stateJson)
				);
				let normalized = null;
				if (rawStateJsonCandidate) {
					try {
						const parsedState = JSON.parse(rawStateJsonCandidate);
						if (parsedState && typeof parsedState === "object") {
							normalized = normalizeRowState(parsedState);
						}
					} catch {
						// Ignore JSON parse errors and fall back to column-based parsing
					}
				}
				if (!normalized) {
					const fallback = {
						distributed: firstPopulatedValue(row.Distributed, row.distributed),
						takenBack: firstPopulatedValue(row.TakenBack, row.takenBack, row.Collected, row.collected),
						missedForm: firstPopulatedValue(
							row.MissedForm,
							row.missedForm,
							row.NotReturned,
							row.notReturned,
							row["Not Returned"],
							row["NotReturned"],
							row.Missed,
							row.DidntReceiveBack,
							row.didntReceiveBack
						),
						deleted: firstPopulatedValue(row.Deleted, row.deleted, row["IsDeleted"]),
						missing: firstPopulatedValue(
							row.FormMissing,
							row.formMissing,
							row.Missing,
							row.missing,
							row["Form Missing"]
						),
						haveForm: firstPopulatedValue(
							row.HaveForm,
							row.haveForm,
							row.IHaveForm,
							row["I Have"],
							row["IHave"]
						),
						onlineUploaded: firstPopulatedValue(
							row.OnlineUploaded,
							row.onlineUploaded,
							row["Online Uploaded"],
							row.UploadedOnline
						),
						distributionType: firstPopulatedValue(
							row.DistributionType,
							row.distributionType,
							row["Distribution Type"],
							row["Distribution"]
						),
						showSecondStatus: firstPopulatedValue(row.ShowSecondStatus, row.showSecondStatus),
						status1: firstPopulatedValue(row.Status1, row.status1),
						status2: firstPopulatedValue(row.Status2, row.status2),
						comments: firstPopulatedValue(row.Comments, row.comments),
						showComments: firstPopulatedValue(row.ShowComments, row.showComments),
						name: firstPopulatedValue(row.Name, row.name, row.FullName, row.fullName),
						phone: firstPopulatedValue(
							row.PhoneNumber,
							row.phoneNumber,
							row.Phone,
							row.phone,
							row.ContactNumber,
							row.contactNumber,
							row.Mobile,
							row.mobile
						),
						epic: firstPopulatedValue(
							row.EpicNumber,
							row.epicNumber,
							row.Epic,
							row.epic,
							row.EPICNumber,
							row.epicNo,
							row["EPIC Number"]
						)
					};
					normalized = normalizeRowState(fallback);
				}
				stateMap.set(formNumber, normalized);
			}

			clearAllRowData();

			for (let i = 1; i <= parsedCount; i++) {
				const state = stateMap.get(i) || normalizeRowState(null);
				saveRowState(i, state);
			}

			const groupsSheet =
				workbook.Sheets.Societies || workbook.Sheets.Groups || workbook.Sheets.societies || null;
			const importedGroups = [];
			const usedGroupIds = new Set();
			if (groupsSheet) {
				const groupsRows = XLSX.utils.sheet_to_json(groupsSheet, { defval: "" });
				let index = 0;
				for (const row of groupsRows) {
					const name = coerceString(firstPopulatedValue(row.Name, row.name)).trim();
					const start = Number.parseInt(firstPopulatedValue(row.Start, row.start), 10);
					const end = Number.parseInt(firstPopulatedValue(row.End, row.end), 10);
					if (!name || !Number.isFinite(start) || !Number.isFinite(end)) continue;
					if (start < 1 || end < 1 || start > MAX_FORMS || end > MAX_FORMS || start > end) continue;
					index += 1;
					const rawId = coerceString(firstPopulatedValue(row.Id, row.ID, row.id)).trim();
					const fallbackPrefix = supportsRowStateJson ? "imported" : "legacy";
					let groupId = rawId || `${fallbackPrefix}_${Date.now()}_${index}`;
					while (usedGroupIds.has(groupId)) {
						groupId = `${groupId}_${index}`;
					}
					usedGroupIds.add(groupId);
					importedGroups.push({
						id: groupId,
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
				} else if (outcome && outcome.outcome === "dismissed") {
					showWarning("Install cancelled. You can try again from the menu.");
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
	if (els.searchSubmit) {
		els.searchSubmit.addEventListener("click", () => {
			void performSearch();
		});
	}
	if (els.searchInput) {
		els.searchInput.addEventListener("keyup", (event) => {
			if (event.key === "Enter") {
				event.preventDefault();
				void performSearch();
			}
		});
	}
	if (els.searchNext) {
		els.searchNext.addEventListener("click", () => {
			if (!searchState || !Array.isArray(searchState.matches) || searchState.matches.length === 0) return;
			const nextIndex = Math.min(searchState.index + 1, searchState.matches.length - 1);
			if (nextIndex !== searchState.index) {
				void focusSearchResult(nextIndex);
			}
		});
	}
	if (els.searchPrev) {
		els.searchPrev.addEventListener("click", () => {
			if (!searchState || !Array.isArray(searchState.matches) || searchState.matches.length === 0) return;
			const prevIndex = Math.max(searchState.index - 1, 0);
			if (prevIndex !== searchState.index) {
				void focusSearchResult(prevIndex);
			}
		});
	}
	if (els.searchDone) {
		els.searchDone.addEventListener("click", () => {
			completeSearchSession();
		});
	}
	if (els.analysisPanel) {
		els.analysisPanel.addEventListener("toggle", () => {
			queueSearchBarHeightSync();
		});
	}
	window.addEventListener("resize", queueSearchBarHeightSync);
	window.addEventListener("scroll", () => {
		queueSearchBarHeightSync();
	});
	queueSearchBarHeightSync();
	updateSearchNavigation();
	window.addEventListener("beforeinstallprompt", (event) => {
		event.preventDefault();
		if (isAppMarkedInstalled()) {
			return;
		}
		deferredInstallPrompt = event;
		const isStandalone = window.matchMedia && window.matchMedia("(display-mode: standalone)").matches;
		const isIosStandalone = typeof window.navigator.standalone === "boolean" && window.navigator.standalone;
		if (!isStandalone && !isIosStandalone) {
			toggleInstallButton(true);
		}
	});
	window.addEventListener("appinstalled", () => {
		deferredInstallPrompt = null;
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
	} else {
		updateFormStats();
	}

	const isAlreadyInstalled =
		(window.matchMedia && window.matchMedia("(display-mode: standalone)").matches) ||
		(typeof window.navigator.standalone === "boolean" && window.navigator.standalone);
	if (isAlreadyInstalled) {
		markAppInstalled();
	} else if (isAppMarkedInstalled()) {
		toggleInstallButton(false);
	}

	if ("serviceWorker" in navigator) {
		const swUrl = `sw.js?v=${ASSET_VERSION}`;
		window.addEventListener("load", async () => {
			try {
				await navigator.serviceWorker.register(swUrl);
			} catch (error) {
				console.error("Service worker registration failed:", error);
			}
		});
	}
})();


