/* global Office, __APP_VERSION__, __BACKEND_BASE__ */

// Build-time injected version (from package.json via webpack DefinePlugin)
const APP_VERSION = (typeof __APP_VERSION__ !== "undefined") ? __APP_VERSION__ : "0.0.0";

// HTML-escape user-supplied data before interpolating into innerHTML.
// All contact names/emails, attachment names, error messages from the backend,
// and other externally-sourced strings must pass through this.
function escHtml(s) {
    return String(s ?? "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#39;");
}

// =====================================================================
// Office initialization
// =====================================================================
Office.onReady(() => {
    console.log("SaveAsPDF taskpane loaded — v" + APP_VERSION);

    // Credit footer version badge
    const versionEl = document.getElementById("appVersion");
    if (versionEl) versionEl.textContent = APP_VERSION;

    document.getElementById("saveBtn").onclick      = onSaveAsPdf;
    document.getElementById("pickLeaderBtn").onclick = () => openContactPicker("leader");

    // Settings panel
    document.getElementById("openSettingsBtn").onclick = openSettingsPanel;

    // Help link → /help on the backend, forced into the user's default browser.
    // window.open with _blank is what Outlook routes to the system shell.
    const helpLink = document.getElementById("helpLink");
    if (helpLink) {
        const helpUrl = `${BACKEND_BASE}/help`;
        helpLink.href = helpUrl;
        helpLink.onclick = (e) => {
            e.preventDefault();
            window.open(helpUrl, "_blank", "noopener,noreferrer");
        };
    }
    document.getElementById("spCloseBtn").onclick      = closeSettingsPanel;
    document.getElementById("spCancelBtn").onclick     = closeSettingsPanel;
    document.getElementById("spSaveBtn").onclick       = saveSettingsPanel;
    document.getElementById("spNewCatBtn").onclick     = showNewCategoryForm;
    document.getElementById("spNewCatCancel").onclick  = hideNewCategoryForm;
    document.getElementById("spNewCatCreate").onclick  = createNewCategory;
    document.getElementById("spDispatchMode").onchange = (e) => {
        document.getElementById("spForwardDefaultRow").style.display =
            e.target.checked ? "block" : "none";
    };
    document.getElementById("spSubfolderAdd").onclick    = addSubfolder;
    document.getElementById("spSubfolderRename").onclick = renameSubfolder;
    document.getElementById("spSubfolderDelete").onclick = deleteSubfolder;
    document.getElementById("spSubfolderClear").onclick  = clearSubfolderDefault;
    document.getElementById("spSubfolderInput").onkeydown = (e) => {
        if (e.key === "Enter") { e.preventDefault(); addSubfolder(); }
    };
    setupSettingsTabs();

    document.getElementById("cpCloseBtn").onclick   = closeContactPicker;
    document.getElementById("cpCancelBtn").onclick  = closeContactPicker;
    document.getElementById("cpSelectBtn").onclick  = confirmPickerSelection;

    document.getElementById("cpSearch").oninput = () => {
        const hasText = document.getElementById("cpSearch").value.length > 0;
        document.getElementById("cpClearSearch").style.display = hasText ? "block" : "none";
        renderPickerList();
    };
    document.getElementById("cpClearSearch").onclick = () => {
        document.getElementById("cpSearch").value = "";
        document.getElementById("cpClearSearch").style.display = "none";
        renderPickerList();
    };

    setupLeaderAutocomplete();
    setupProjectIdLookup();
    setupFolderContextMenu();
    initializeTabs();
    initializeEmployeesTab();
    loadAttachments();
    tryRestoreProjectInfo();
    preloadContacts(); // background – feeds autocomplete & speeds up picker

    loadPrefs();
    applyPrefsToCheckboxes();
    updateCategoryLabel();
    updateDestBanner();

    fetchAdminPolicy();         // populates _adminPolicy, then re-applies prefs
    startBackendHealthCheck();
});

// =====================================================================
// Backend health check — polls /api/info; shows a red banner when down
// =====================================================================
// Webpack injects __BACKEND_BASE__: "http://localhost:5176" in dev,
// "" (same origin) in production builds. Empty fallback lets pure-browser tests work.
const BACKEND_BASE       = (typeof __BACKEND_BASE__ !== "undefined") ? __BACKEND_BASE__ : "";
const HEALTH_INTERVAL_OK = 30000; // ms – when backend is healthy
const HEALTH_INTERVAL_KO = 5000;  // ms – when backend is down (retry sooner)
let _backendUp = true;

function startBackendHealthCheck() {
    pingBackend();
    schedule();

    function schedule() {
        const next = _backendUp ? HEALTH_INTERVAL_OK : HEALTH_INTERVAL_KO;
        setTimeout(async () => {
            await pingBackend();
            schedule();
        }, next);
    }
}

async function pingBackend() {
    try {
        const ctrl = new AbortController();
        const t = setTimeout(() => ctrl.abort(), 4000);
        const res = await fetch(`${BACKEND_BASE}/api/info`, { signal: ctrl.signal });
        clearTimeout(t);
        setBackendUp(res.ok);
    } catch {
        setBackendUp(false);
    }
}

function setBackendUp(up) {
    if (up === _backendUp) return; // no transition
    _backendUp = up;
    const banner = document.getElementById("backendBanner");
    if (banner) banner.style.display = up ? "none" : "block";
    const saveBtn = document.getElementById("saveBtn");
    if (saveBtn) {
        saveBtn.disabled = !up;
        saveBtn.title = up ? "" : "השרת הראשי אינו זמין";
    }
    // Backend just came online — refresh admin policy in case it changed.
    if (up) fetchAdminPolicy();
}

// =====================================================================
// Project lookup — fetch project data from backend when project number changes
// =====================================================================
let _projectLookupTimer = null;

function setupProjectIdLookup() {
    const input = document.getElementById("projectId");
    input.addEventListener("input", () => {
        clearTimeout(_projectLookupTimer);
        _projectLookupTimer = setTimeout(loadProjectByNumber, 600);
    });
    input.addEventListener("blur", () => {
        clearTimeout(_projectLookupTimer);
        loadProjectByNumber();
    });
}

async function loadProjectByNumber() {
    const num = document.getElementById("projectId").value.trim();
    const status = document.getElementById("status");
    if (!num) return;

    try {
        const res = await fetch(`${BACKEND_BASE}/api/project/${encodeURIComponent(num)}`);
        if (!res.ok) {
            const err = await res.json().catch(() => ({}));
            status.innerText = err.error || `שגיאת שרת ${res.status}`;
            return;
        }
        const data = await res.json();

        if (!data.exists) {
            status.innerText = `הפרויקט לא נמצא בתיקייה: ${data.expectedPath || ""}`;
            _folderTree = null;
            renderFolderTree();
            updateDestBanner();

            const r = await dlgConfirm(
                "הפרויקט לא קיים",
                `הפרויקט "${num}" לא נמצא בנתיב הצפוי:\n${data.expectedPath || ""}\n\nליצור את התיקיה?`,
                "צור פרויקט"
            );
            if (r.ok) {
                try {
                    const cr = await fetch(
                        `${BACKEND_BASE}/api/project/${encodeURIComponent(num)}`,
                        { method: "POST" }
                    );
                    if (!cr.ok) {
                        const e = await cr.json().catch(() => ({}));
                        await dlgAlert("שגיאה ביצירת פרויקט", e.error || ("HTTP " + cr.status));
                        return;
                    }
                    const created = await cr.json();
                    status.innerText = `הפרויקט נוצר בנתיב: ${created.folderPath}`;
                    // Re-trigger the lookup so the tree + project data load
                    loadProjectByNumber();
                } catch (e) {
                    await dlgAlert("שגיאה ביצירת פרויקט", e.message);
                }
            }
            return;
        }

        status.innerText = "";

        // Load the folder tree for this project (independent of project XML)
        loadFolderTree(num);
        if (data.projectName) {
            document.getElementById("projectName").value = data.projectName;
        }

        // Replace employees with the loaded list (if non-empty)
        if (Array.isArray(data.employees) && data.employees.length > 0) {
            _employees.length = 0;
            data.employees.forEach(e => _employees.push({
                displayName: e.displayName,
                email: e.email,
                isLeader: !!e.isLeader
            }));

            const leader = data.employees.find(e => e.isLeader);
            if (leader) {
                const lin = document.getElementById("projectLeader");
                lin.value = `${leader.displayName}  <${leader.email}>`;
                lin.dataset.email = leader.email;
            }
            renderEmployees();
        }
    } catch (e) {
        console.warn("Project lookup failed:", e);
    }
}

// =====================================================================
// Contacts cache (shared by picker + autocomplete)
// =====================================================================
let _allContacts      = [];   // flat deduplicated list
let _cachedPickerData = null; // full {folders, contacts} from backend

async function preloadContacts() {
    try {
        const res = await fetch(`${BACKEND_BASE}/api/contacts`);
        if (!res.ok) return;
        _cachedPickerData = await res.json();
        const seen = new Set();
        _allContacts = Object.values(_cachedPickerData.contacts || {})
            .flat()
            .filter(c => { if (seen.has(c.email)) return false; seen.add(c.email); return true; });
    } catch { }
}

// =====================================================================
// Project leader field
// =====================================================================
function setProjectLeader(contact) {
    const input = document.getElementById("projectLeader");
    input.value = `${contact.displayName}  <${contact.email}>`;
    input.dataset.email = contact.email;

    // Demote previous leader badge
    _employees.forEach(e => { e.isLeader = false; });

    // Add / promote to leader in the employees list
    const existing = _employees.find(e => e.email === contact.email);
    if (existing) {
        existing.isLeader = true;
    } else {
        _employees.unshift({ ...contact, isLeader: true });
    }
    renderEmployees();
    hideLeaderDropdown();
}

// =====================================================================
// Leader autocomplete
// =====================================================================
function setupLeaderAutocomplete() {
    const input    = document.getElementById("projectLeader");
    const dropdown = document.getElementById("leaderDropdown");

    input.addEventListener("input", () => {
        const q = input.value.toLowerCase().trim();
        if (!q || q.includes("<")) { hideLeaderDropdown(); return; }

        const hits = _allContacts
            .filter(c => c.displayName.toLowerCase().includes(q) || c.email.toLowerCase().includes(q))
            .slice(0, 8);

        if (!hits.length) { hideLeaderDropdown(); return; }

        dropdown.innerHTML = "";
        hits.forEach(c => {
            const item = document.createElement("div");
            item.style.cssText = "padding:7px 10px;cursor:pointer;border-bottom:1px solid #f0f0f0";
            item.innerHTML =
                `<div style="font-weight:600;font-size:13px">${escHtml(c.displayName)}</div>` +
                `<div style="font-size:11px;color:#666;direction:ltr;text-align:left">${escHtml(c.email)}</div>`;
            item.onmouseenter = () => item.style.background = "#f3f3f3";
            item.onmouseleave = () => item.style.background = "";
            item.onmousedown  = e => { e.preventDefault(); setProjectLeader(c); };
            dropdown.appendChild(item);
        });
        dropdown.style.display = "block";
    });

    input.addEventListener("blur", () => setTimeout(hideLeaderDropdown, 150));
}

function hideLeaderDropdown() {
    document.getElementById("leaderDropdown").style.display = "none";
}

// =====================================================================
// Attachment list
// =====================================================================
function loadAttachments() {
    const item      = Office.context.mailbox.item;
    const container = document.getElementById("attachmentsTab");

    if (!item || !item.attachments || item.attachments.length === 0) {
        container.innerHTML = "<p>אין קבצים מצורפים</p>";
        return;
    }

    container.innerHTML = item.attachments.map(att => `
        <label style="display:flex;align-items:center;gap:8px;margin-top:8px">
            <input type="checkbox" class="att-check" data-id="${escHtml(att.id)}" checked />
            <span>${escHtml(att.name)} <span style="color:#888;font-size:12px">(${escHtml(formatSize(att.size))})</span></span>
        </label>
    `).join("");
}

function formatSize(bytes) {
    if (bytes < 1024) return bytes + " B";
    if (bytes < 1024 * 1024) return Math.round(bytes / 1024) + " KB";
    return (bytes / (1024 * 1024)).toFixed(1) + " MB";
}

// =====================================================================
// Employees tab
// =====================================================================
const _employees = [];

function initializeEmployeesTab() { renderEmployees(); }

function addEmployee(contact) {
    if (_employees.some(e => e.email === contact.email)) return;
    _employees.push({ ...contact });
    renderEmployees();
}

function renderEmployees() {
    const container = document.getElementById("employeesTab");

    const rows = _employees.length === 0
        ? `<div style="color:#888;padding:8px 0;font-size:13px">לא נבחרו עובדים</div>`
        : _employees.map((e, i) => `
            <div style="display:flex;align-items:center;justify-content:space-between;
                        padding:6px 0;border-bottom:1px solid #f0f0f0;
                        ${e.isLeader ? "background:#fffbe6;" : ""}">
                <div>
                    <div style="font-weight:600;font-size:13px">
                        ${e.isLeader
                            ? `<span style="background:#f0c000;color:#333;font-size:10px;
                                    padding:1px 5px;border-radius:8px;margin-left:4px;
                                    font-weight:700">מנהל</span>`
                            : ""}
                        ${escHtml(e.displayName)}
                    </div>
                    <div style="font-size:11px;color:#666;direction:ltr;text-align:right">${escHtml(e.email)}</div>
                </div>
                <button type="button" class="remove-emp-btn" data-idx="${i}"
                        style="border:none;background:none;cursor:pointer;color:#cc0000;
                               font-size:16px;padding:2px 6px">✕</button>
            </div>
        `).join("");

    container.innerHTML = `
        <div style="margin-bottom:8px">
            <button type="button" id="addEmployeeBtn" style="margin-top:0">+ הוסף עובד</button>
        </div>
        ${rows}
    `;

    document.getElementById("addEmployeeBtn").onclick = () => openContactPicker("employee");

    container.querySelectorAll(".remove-emp-btn").forEach(btn => {
        btn.onclick = () => {
            _employees.splice(parseInt(btn.dataset.idx, 10), 1);
            renderEmployees();
        };
    });
}

// =====================================================================
// User preferences (Office.context.roamingSettings) + Settings panel
// =====================================================================
const PREFS_KEY = "saveAsPdfPrefs";

const DEFAULT_PREFS = {
    defaultStamp:    true,
    stampFields:     { projectId: true, projectName: true, leader: true, date: true,
                       user: true, employees: true, attachments: true },
    stampNotes:      "",
    stampTemplate:   "",
    defaultCategory: true,
    categoryName:    "SaveAsPDF – Processed",
    folderNaming:    { mode: "default", customPrefix: "" },   // default | date | custom
    dispatchMode:    false,                                    // enables forward-to-leader feature
    forwardToLeaderByDefault: false,                           // pre-checks the forward checkbox each time
    defaultSubfolders: {
        list:       ["מכתבים", "התקבל", "אישור ציוד"],
        selected:   "",      // "" = no default sub-folder
        appendDate: false    // append today's date as inner sub-folder
    }
};

let _prefs = { ...DEFAULT_PREFS, stampFields: { ...DEFAULT_PREFS.stampFields } };

function loadPrefs() {
    try {
        const raw = Office.context.roamingSettings?.get(PREFS_KEY);
        if (!raw || typeof raw !== "object") return;
        _prefs = {
            ...DEFAULT_PREFS,
            ...raw,
            stampFields:  { ...DEFAULT_PREFS.stampFields,  ...(raw.stampFields  || {}) },
            folderNaming: { ...DEFAULT_PREFS.folderNaming, ...(raw.folderNaming || {}) },
            defaultSubfolders: {
                ...DEFAULT_PREFS.defaultSubfolders,
                ...(raw.defaultSubfolders || {}),
                list: Array.isArray(raw.defaultSubfolders?.list)
                    ? raw.defaultSubfolders.list.filter(s => typeof s === "string" && s.trim())
                    : DEFAULT_PREFS.defaultSubfolders.list
            }
        };
    } catch (e) { console.warn("loadPrefs failed", e); }
}

function savePrefs() {
    try {
        Office.context.roamingSettings.set(PREFS_KEY, _prefs);
        Office.context.roamingSettings.saveAsync();
    } catch (e) { console.warn("savePrefs failed", e); }
}

// ---------- Admin policy (server-enforced stamp overrides) ----------
// Map of admin-forced stamp fields. Shape:
//   { defaultStamp: bool|null, includeProjectId: bool|null, ... }
// null/undefined = user controls; true/false = forced value.
let _adminPolicy = {};

// Mapping from policy key -> taskpane checkbox id (used to lock the UI).
const POLICY_FIELD_MAP = {
    defaultStamp:        "spDefaultStamp",
    includeProjectId:    "spFieldProjectId",
    includeProjectName:  "spFieldProjectName",
    includeLeader:       "spFieldLeader",
    includeDate:         "spFieldDate",
    includeUser:         "spFieldUser",
    includeEmployees:    "spFieldEmployees",
    includeAttachments:  "spFieldAttachments"
};

async function fetchAdminPolicy() {
    try {
        const ctrl = new AbortController();
        const t = setTimeout(() => ctrl.abort(), 4000);
        const res = await fetch(`${BACKEND_BASE}/api/policy`, { signal: ctrl.signal });
        clearTimeout(t);
        if (!res.ok) return;
        const body = await res.json();
        _adminPolicy = body.stamp || {};
        applyPrefsToCheckboxes();   // re-apply so the master "stamp" toggle reflects the policy
    } catch { /* backend may be down — fall back to user prefs only */ }
}

// Returns the effective value of a policy field: null (user-controlled) or bool.
function policyForcedValue(key) {
    const v = _adminPolicy?.[key];
    return (v === true || v === false) ? v : null;
}

// Lock an input to the forced value and decorate with a lock indicator.
function lockField(checkboxId, forcedValue) {
    const cb = document.getElementById(checkboxId);
    if (!cb) return;
    cb.checked  = forcedValue;
    cb.disabled = true;
    const label = cb.closest("label");
    if (label && !label.querySelector(".policy-lock")) {
        const tag = document.createElement("span");
        tag.className = "policy-lock";
        tag.textContent = " 🔒";
        tag.title = "ההגדרה נקבעה על-ידי מנהל המערכת";
        tag.style.cssText = "color:#888;font-size:11px;margin-right:4px";
        label.appendChild(tag);
        label.style.opacity = "0.75";
    }
}

function unlockField(checkboxId) {
    const cb = document.getElementById(checkboxId);
    if (!cb) return;
    cb.disabled = false;
    const label = cb.closest("label");
    if (label) {
        label.style.opacity = "";
        label.querySelectorAll(".policy-lock").forEach(n => n.remove());
    }
}

// Apply locks for any field the admin has forced. Called at the end of openSettingsPanel.
function applyPolicyLocks() {
    Object.entries(POLICY_FIELD_MAP).forEach(([key, id]) => {
        const forced = policyForcedValue(key);
        if (forced === null) unlockField(id);
        else                 lockField(id, forced);
    });
}

function applyPrefsToCheckboxes() {
    const optStamp = document.getElementById("optStamp");
    const forcedStamp = policyForcedValue("defaultStamp");
    optStamp.checked = forcedStamp !== null ? forcedStamp : !!_prefs.defaultStamp;
    optStamp.disabled = forcedStamp !== null;
    optStamp.title = forcedStamp !== null ? "🔒 ההגדרה נקבעה על-ידי מנהל המערכת" : "";

    document.getElementById("optCategory").checked = !!_prefs.defaultCategory;

    // Dispatch / forward-to-leader visibility + default state
    const row = document.getElementById("forwardToLeaderRow");
    const cb  = document.getElementById("optForwardToLeader");
    if (row && cb) {
        row.style.display = _prefs.dispatchMode ? "block" : "none";
        cb.checked = _prefs.dispatchMode && _prefs.forwardToLeaderByDefault;
    }
}

function updateCategoryLabel() {
    // Update the visible label next to the optCategory checkbox to reflect chosen category
    const lbl = document.querySelector('label > input#optCategory');
    if (!lbl) return;
    const parent = lbl.parentElement;
    // Replace text node after the checkbox with the chosen category name
    const name = _prefs.categoryName || DEFAULT_PREFS.categoryName;
    parent.childNodes.forEach(n => {
        if (n.nodeType === 3) parent.removeChild(n);
    });
    parent.appendChild(document.createTextNode(` שיוך קטגוריה «${name}»`));
}

// ---------- Settings panel UI ----------
function openSettingsPanel() {
    document.body.classList.add("overlay-open");
    document.getElementById("settingsPanel").style.display = "flex";

    // Populate UI from current prefs
    document.getElementById("spDefaultStamp").checked    = !!_prefs.defaultStamp;
    document.getElementById("spFieldProjectId").checked   = !!_prefs.stampFields.projectId;
    document.getElementById("spFieldProjectName").checked = !!_prefs.stampFields.projectName;
    document.getElementById("spFieldLeader").checked      = !!_prefs.stampFields.leader;
    document.getElementById("spFieldDate").checked        = !!_prefs.stampFields.date;
    document.getElementById("spFieldUser").checked        = !!_prefs.stampFields.user;
    document.getElementById("spFieldEmployees").checked   = !!_prefs.stampFields.employees;
    document.getElementById("spFieldAttachments").checked = !!_prefs.stampFields.attachments;
    document.getElementById("spStampNotes").value        = _prefs.stampNotes || "";
    document.getElementById("spStampTemplate").value     = _prefs.stampTemplate || "";

    document.getElementById("spDefaultCategory").checked = !!_prefs.defaultCategory;

    // Folder naming radios + custom prefix
    const fnMode = _prefs.folderNaming?.mode || "default";
    document.querySelectorAll('input[name="spFolderNaming"]').forEach(r => {
        r.checked = (r.value === fnMode);
    });
    document.getElementById("spCustomPrefix").value = _prefs.folderNaming?.customPrefix || "";

    // Dispatch
    document.getElementById("spDispatchMode").checked = !!_prefs.dispatchMode;
    document.getElementById("spForwardDefault").checked = !!_prefs.forwardToLeaderByDefault;
    document.getElementById("spForwardDefaultRow").style.display =
        _prefs.dispatchMode ? "block" : "none";

    // Default sub-folder destination (tentative copies edited until Save)
    _spSubfolderList     = [...(_prefs.defaultSubfolders?.list || [])];
    _spSubfolderSelected = _prefs.defaultSubfolders?.selected || "";
    document.getElementById("spSubfolderAppendDate").checked = !!_prefs.defaultSubfolders?.appendDate;
    document.getElementById("spSubfolderInput").value = "";
    renderSubfolderList();

    // Restore the cat-list class in case the previous open switched to fallback
    const catContainer = document.getElementById("spCategoryList");
    if (catContainer) catContainer.classList.add("cat-list");

    // Tentative selection starts from the saved prefs
    _spSelectedCategory = _prefs.categoryName || "";

    hideNewCategoryForm();
    populateCategoryList();

    // Locks any fields the admin has forced (must run after the inputs are populated).
    applyPolicyLocks();
}

function closeSettingsPanel() {
    document.body.classList.remove("overlay-open");
    document.getElementById("settingsPanel").style.display = "none";
}

function saveSettingsPanel() {
    _prefs.defaultStamp    = document.getElementById("spDefaultStamp").checked;
    _prefs.stampFields = {
        projectId:   document.getElementById("spFieldProjectId").checked,
        projectName: document.getElementById("spFieldProjectName").checked,
        leader:      document.getElementById("spFieldLeader").checked,
        date:        document.getElementById("spFieldDate").checked,
        user:        document.getElementById("spFieldUser").checked,
        employees:   document.getElementById("spFieldEmployees").checked,
        attachments: document.getElementById("spFieldAttachments").checked
    };
    _prefs.stampNotes      = document.getElementById("spStampNotes").value.trim();
    _prefs.stampTemplate   = document.getElementById("spStampTemplate").value.trim();
    _prefs.defaultCategory = document.getElementById("spDefaultCategory").checked;
    if (_spSelectedCategory) _prefs.categoryName = _spSelectedCategory;

    const fnRadio = document.querySelector('input[name="spFolderNaming"]:checked');
    _prefs.folderNaming = {
        mode:         fnRadio ? fnRadio.value : "default",
        customPrefix: document.getElementById("spCustomPrefix").value.trim()
    };

    _prefs.dispatchMode             = document.getElementById("spDispatchMode").checked;
    _prefs.forwardToLeaderByDefault = document.getElementById("spForwardDefault").checked;

    _prefs.defaultSubfolders = {
        list:       _spSubfolderList.slice(),
        selected:   _spSubfolderSelected,
        appendDate: document.getElementById("spSubfolderAppendDate").checked
    };

    savePrefs();
    applyPrefsToCheckboxes();
    updateCategoryLabel();
    updateDestBanner();
    closeSettingsPanel();
}

// ---------- Default sub-folder list (settings panel) ----------
let _spSubfolderList     = [];   // tentative list while panel is open
let _spSubfolderSelected = "";   // tentative selected default

function renderSubfolderList() {
    const c = document.getElementById("spSubfolderList");
    if (!c) return;
    if (_spSubfolderList.length === 0) {
        c.innerHTML = `<div class="sf-empty">אין תיקיות ברירת מחדל. הוסף בעזרת תיבת הקלט מטה.</div>`;
        return;
    }
    c.innerHTML = "";
    _spSubfolderList.forEach(name => {
        const row = document.createElement("div");
        row.className = "sf-row" + (name === _spSubfolderSelected ? " selected" : "");
        row.textContent = name;
        row.onclick = () => {
            _spSubfolderSelected = (_spSubfolderSelected === name) ? "" : name;
            renderSubfolderList();
        };
        c.appendChild(row);
    });
}

function addSubfolder() {
    const input = document.getElementById("spSubfolderInput");
    const name = (input.value || "").trim();
    if (!name) return;
    if (_spSubfolderList.includes(name)) {
        input.value = "";
        _spSubfolderSelected = name;
        renderSubfolderList();
        return;
    }
    _spSubfolderList.push(name);
    _spSubfolderSelected = name;
    input.value = "";
    renderSubfolderList();
}

async function renameSubfolder() {
    if (!_spSubfolderSelected) {
        await dlgAlert("שינוי שם", "בחר/י תיקיה מהרשימה תחילה.");
        return;
    }
    const r = await dlgPrompt("שינוי שם תיקיה", _spSubfolderSelected, "שמור");
    if (!r.ok) return;
    const newName = (r.value || "").trim();
    if (!newName || newName === _spSubfolderSelected) return;
    if (_spSubfolderList.includes(newName)) {
        await dlgAlert("שינוי שם", "השם כבר קיים ברשימה.");
        return;
    }
    const idx = _spSubfolderList.indexOf(_spSubfolderSelected);
    if (idx >= 0) _spSubfolderList[idx] = newName;
    _spSubfolderSelected = newName;
    renderSubfolderList();
}

async function deleteSubfolder() {
    if (!_spSubfolderSelected) {
        await dlgAlert("הסרה", "בחר/י תיקיה מהרשימה תחילה.");
        return;
    }
    const ok = await dlgConfirm("הסרה מהרשימה", `להסיר את "${_spSubfolderSelected}" מהרשימה?`);
    if (!ok) return;
    _spSubfolderList = _spSubfolderList.filter(n => n !== _spSubfolderSelected);
    _spSubfolderSelected = "";
    renderSubfolderList();
}

function clearSubfolderDefault() {
    _spSubfolderSelected = "";
    renderSubfolderList();
}

// Resolves the destination sub-folder for a save operation.
// Manual selection in the folder tree (if any) wins over the prefs default.
function resolveSaveDestination() {
    if (_selectedFolderPath) return _selectedFolderPath;
    const ds = _prefs.defaultSubfolders;
    if (!ds || !ds.selected) return "";
    if (!ds.appendDate) return ds.selected;
    const d = new Date();
    const dateStr = `${pad2(d.getDate())}.${pad2(d.getMonth() + 1)}.${d.getFullYear()}`;
    return `${ds.selected}/${dateStr}`;
}

// ---------- Master categories list (settings panel) ----------
let _categoryFallback   = false; // true = listing API blocked, free-text fallback
let _spSelectedCategory = "";    // tentative selection while panel is open
let _spLoadedCategories = [];    // cached list from masterCategories.getAsync

function populateCategoryList() {
    const container = document.getElementById("spCategoryList");
    if (!container) return;

    container.innerHTML = `<div class="cat-empty">טוען רשימת קטגוריות...</div>`;
    _categoryFallback = false;

    if (!Office.context.mailbox.masterCategories?.getAsync) {
        showCategoryFallback();
        return;
    }

    try {
        Office.context.mailbox.masterCategories.getAsync(r => {
            if (r.status !== Office.AsyncResultStatus.Succeeded) {
                showCategoryFallback();
                return;
            }
            _spLoadedCategories = r.value || [];
            renderCategoryList();
        });
    } catch (e) {
        showCategoryFallback();
    }
}

function renderCategoryList() {
    const container = document.getElementById("spCategoryList");
    if (!container) return;
    container.innerHTML = "";

    let cats = [..._spLoadedCategories];

    // Always include the currently chosen category, even if not in master list yet
    if (_spSelectedCategory && !cats.some(c => c.displayName === _spSelectedCategory)) {
        cats.unshift({ displayName: _spSelectedCategory, color: null, _autoCreate: true });
    }

    if (cats.length === 0) {
        container.innerHTML = `<div class="cat-empty">אין קטגוריות מוגדרות</div>`;
        return;
    }

    cats.forEach(cat => {
        const row = document.createElement("div");
        row.className = "cat-row" + (cat.displayName === _spSelectedCategory ? " selected" : "");

        const swatch  = document.createElement("span");
        swatch.className = "cat-swatch";
        const palette = CATEGORY_PALETTE.find(p => p.name === cat.color);
        swatch.style.background = palette ? palette.hex : "#ccc";
        if (!palette) swatch.title = "אין צבע מוגדר";
        row.appendChild(swatch);

        const name = document.createElement("span");
        name.className   = "cat-name";
        name.textContent = cat.displayName + (cat._autoCreate ? "  (יווצר אוטומטית)" : "");
        row.appendChild(name);

        if (cat.displayName === _spSelectedCategory) {
            const check = document.createElement("span");
            check.className = "cat-check";
            check.textContent = "✓";
            row.appendChild(check);
        }

        row.onclick = () => {
            _spSelectedCategory = cat.displayName;
            renderCategoryList();
        };
        container.appendChild(row);
    });
}

// When masterCategories API is denied, swap the list for a free-text input.
// On save the tag is applied via item.categories.addAsync; Outlook auto-creates
// the category by name (without a custom color, but it works as a tag).
function showCategoryFallback() {
    if (_categoryFallback) return;
    _categoryFallback = true;

    const container = document.getElementById("spCategoryList");
    if (!container) return;

    container.innerHTML = "";
    container.classList.remove("cat-list");

    const input = document.createElement("input");
    input.type = "text";
    input.id = "spCategoryFallbackInput";
    input.value = _spSelectedCategory || "";
    input.placeholder = "שם קטגוריה (יווצר אוטומטית בשימוש)";
    input.oninput = () => { _spSelectedCategory = input.value.trim(); };
    container.appendChild(input);

    const note = document.createElement("div");
    note.className = "cat-fallback-note";
    note.textContent =
        "סביבת ה‑Outlook הזו אינה מאפשרת לטעון רשימת קטגוריות. " +
        "הזן שם קטגוריה ידנית — Outlook ייצור אותה אוטומטית בעת השימוש.";
    container.appendChild(note);

    const newBtn = document.getElementById("spNewCatBtn");
    if (newBtn) newBtn.style.display = "none";
}

// Settings panel tab switcher
function setupSettingsTabs() {
    const buttons = document.querySelectorAll(".sp-tab-btn");
    buttons.forEach(btn => {
        btn.addEventListener("click", () => {
            buttons.forEach(b => b.classList.remove("active"));
            document.querySelectorAll(".sp-tab-panel").forEach(p => p.classList.remove("active"));
            btn.classList.add("active");
            document.getElementById(btn.dataset.spTab).classList.add("active");
        });
    });
}

// 25-color palette mapped to Office.MailboxEnums.CategoryColor.Preset0..24
// Hex values are approximations of Outlook's canonical category colors.
const CATEGORY_PALETTE = [
    { name: "Preset0",  hex: "#E91D63" }, { name: "Preset1",  hex: "#FF8C00" },
    { name: "Preset2",  hex: "#A0522D" }, { name: "Preset3",  hex: "#FFC83D" },
    { name: "Preset4",  hex: "#5DBA47" }, { name: "Preset5",  hex: "#0FB5B5" },
    { name: "Preset6",  hex: "#7E9C2E" }, { name: "Preset7",  hex: "#0078D4" },
    { name: "Preset8",  hex: "#8764B8" }, { name: "Preset9",  hex: "#C30052" },
    { name: "Preset10", hex: "#69797E" }, { name: "Preset11", hex: "#4F6228" },
    { name: "Preset12", hex: "#525252" }, { name: "Preset13", hex: "#A6A6A6" },
    { name: "Preset14", hex: "#000000" }, { name: "Preset15", hex: "#9C0000" },
    { name: "Preset16", hex: "#C75300" }, { name: "Preset17", hex: "#6F4123" },
    { name: "Preset18", hex: "#B89230" }, { name: "Preset19", hex: "#226B2A" },
    { name: "Preset20", hex: "#0B7A78" }, { name: "Preset21", hex: "#42561F" },
    { name: "Preset22", hex: "#003F8A" }, { name: "Preset23", hex: "#5B348B" },
    { name: "Preset24", hex: "#7A0033" }
];

let _selectedSwatch = "Preset7"; // default to blue

function showNewCategoryForm() {
    document.getElementById("spNewCatForm").style.display = "block";
    document.getElementById("spNewCatName").value = "";
    _selectedSwatch = "Preset7";
    renderSwatches();
}

function hideNewCategoryForm() {
    document.getElementById("spNewCatForm").style.display = "none";
}

function renderSwatches() {
    const grid = document.getElementById("spSwatchGrid");
    grid.innerHTML = "";
    CATEGORY_PALETTE.forEach(c => {
        const s = document.createElement("div");
        s.className = "swatch" + (c.name === _selectedSwatch ? " selected" : "");
        s.style.background = c.hex;
        s.title = c.name;
        s.onclick = () => { _selectedSwatch = c.name; renderSwatches(); };
        grid.appendChild(s);
    });
}

function createNewCategory() {
    const name = document.getElementById("spNewCatName").value.trim();
    if (!name) { alert("יש להזין שם קטגוריה"); return; }

    if (!Office.context.mailbox.masterCategories?.addAsync) {
        alert("יצירת קטגוריות דורשת Outlook 2019 ומעלה (Mailbox 1.8+)");
        return;
    }

    try {
        Office.context.mailbox.masterCategories.addAsync(
            [{ displayName: name, color: Office.MailboxEnums.CategoryColor[_selectedSwatch] }],
            r => {
                if (r.status !== Office.AsyncResultStatus.Succeeded) {
                    dlgAlert("שגיאה ביצירת הקטגוריה",
                        (r.error?.message || "") +
                        "\n\nניתן עדיין להזין שם קטגוריה ידנית — Outlook ייצור אותה בשימוש (ללא צבע מותאם).");
                    _spSelectedCategory = name;
                    hideNewCategoryForm();
                    showCategoryFallback();
                    return;
                }
                // Optimistically add to the cached list so the user sees it immediately
                _spLoadedCategories.push({ displayName: name, color: _selectedSwatch });
                _spSelectedCategory = name;
                hideNewCategoryForm();
                renderCategoryList();
            }
        );
    } catch (e) {
        dlgAlert("יצירת קטגוריה אינה זמינה",
            e.message + "\n\nניתן להזין שם קטגוריה ידנית — Outlook ייצור אותה אוטומטית.");
        _spSelectedCategory = name;
        hideNewCategoryForm();
        showCategoryFallback();
    }
}

// =====================================================================
// Custom modal dialog (Outlook task panes block prompt/confirm/alert)
// =====================================================================
function dlg(opts) {
    return new Promise(resolve => {
        const bd     = document.getElementById("modalBackdrop");
        const titleE = document.getElementById("modalTitle");
        const msgE   = document.getElementById("modalMsg");
        const inp    = document.getElementById("modalInput");
        const ok     = document.getElementById("modalOk");
        const cancel = document.getElementById("modalCancel");

        titleE.textContent = opts.title || "";
        msgE.textContent   = opts.message || "";
        msgE.style.display = opts.message ? "block" : "none";

        if (opts.input) {
            inp.style.display = "block";
            inp.value = opts.defaultValue || "";
        } else {
            inp.style.display = "none";
        }

        ok.textContent     = opts.okLabel     || "אישור";
        cancel.textContent = opts.cancelLabel || "ביטול";
        cancel.style.display = opts.hideCancel ? "none" : "inline-block";

        bd.style.display = "flex";
        if (opts.input) setTimeout(() => { inp.focus(); inp.select(); }, 50);

        const close = (result) => {
            bd.style.display = "none";
            ok.onclick = cancel.onclick = inp.onkeydown = null;
            resolve(result);
        };
        ok.onclick     = () => close(opts.input ? { ok: true, value: inp.value.trim() } : { ok: true });
        cancel.onclick = () => close({ ok: false });
        inp.onkeydown  = (e) => {
            if (e.key === "Enter")  { e.preventDefault(); ok.click(); }
            if (e.key === "Escape") { e.preventDefault(); cancel.click(); }
        };
    });
}
const dlgPrompt  = (title, defaultValue, okLabel) => dlg({ title, input: true, defaultValue, okLabel });
const dlgConfirm = (title, message, okLabel)      => dlg({ title, message, okLabel });
const dlgAlert   = (title, message)               => dlg({ title, message, hideCancel: true, okLabel: "סגור" });

// =====================================================================
// Project folder tree (in the third tab)
// =====================================================================
let _folderTree         = null;   // root node from backend
let _selectedFolderPath = "";     // empty = project root
let _expandedPaths      = new Set([""]);     // root always expanded
let _treeProjectNumber  = "";     // project the tree was loaded for

async function loadFolderTree(projectNumber) {
    _treeProjectNumber = projectNumber;
    const treeEl = document.getElementById("folderTree");
    treeEl.innerHTML = `<div style="color:#888;padding:10px 0">טוען עץ תיקיות...</div>`;

    try {
        const res = await fetch(
            `${BACKEND_BASE}/api/project/${encodeURIComponent(projectNumber)}/folders/tree`
        );
        if (!res.ok) {
            const err = await res.json().catch(() => ({}));
            treeEl.innerHTML =
                `<div style="color:#c53030;padding:10px 0;font-size:12px">${escHtml(err.error || "שגיאה בטעינת עץ התיקיות")}</div>`;
            _folderTree = null;
            updateDestBanner();
            return;
        }
        const body = await res.json();
        _folderTree = body.tree;
        _selectedFolderPath = ""; // reset to project root on every load
        renderFolderTree();
        updateDestBanner();
    } catch (e) {
        treeEl.innerHTML = `<div style="color:#c53030;padding:10px 0">${escHtml(e.message)}</div>`;
        _folderTree = null;
        updateDestBanner();
    }
}

function renderFolderTree() {
    const treeEl = document.getElementById("folderTree");
    if (!_folderTree) {
        treeEl.innerHTML = `<div style="color:#888;padding:10px 0">אין נתונים</div>`;
        return;
    }
    treeEl.innerHTML = "";
    treeEl.appendChild(renderFolderNode(_folderTree, 0));
}

function renderFolderNode(node, depth) {
    const wrap = document.createElement("div");

    const row = document.createElement("div");
    row.className = "folder-row" + (node.path === _selectedFolderPath ? " selected" : "");

    const hasChildren = (node.children?.length || 0) > 0;
    const expanded    = _expandedPaths.has(node.path);

    const chev = document.createElement("span");
    chev.className = "folder-chevron" + (hasChildren ? "" : " empty") + (expanded ? " expanded" : "");
    chev.textContent = "▶";
    chev.onclick = (e) => {
        e.stopPropagation();
        if (!hasChildren) return;
        if (_expandedPaths.has(node.path)) _expandedPaths.delete(node.path);
        else _expandedPaths.add(node.path);
        renderFolderTree();
    };
    row.appendChild(chev);

    const icon = document.createElement("span");
    icon.className = "folder-icon";
    icon.textContent = (hasChildren && expanded) ? "📂" : "📁";
    row.appendChild(icon);

    const label = document.createElement("span");
    label.className = "folder-name";
    label.textContent = node.path === "" ? `${node.name} (root)` : node.name;
    label.title = node.path || "(project root)";
    row.appendChild(label);

    row.onclick = () => {
        _selectedFolderPath = node.path;
        renderFolderTree();
        updateDestBanner();
    };
    row.oncontextmenu = (e) => {
        e.preventDefault();
        _selectedFolderPath = node.path;
        renderFolderTree();
        updateDestBanner();
        showFolderContextMenu(e.clientX, e.clientY, node);
    };

    wrap.appendChild(row);

    if (hasChildren && expanded) {
        const children = document.createElement("div");
        children.className = "folder-children";
        node.children.forEach(c => children.appendChild(renderFolderNode(c, depth + 1)));
        wrap.appendChild(children);
    }
    return wrap;
}

function updateDestBanner() {
    const banner = document.getElementById("destBanner");
    const pathEl = document.getElementById("destPath");
    const dest   = resolveSaveDestination();
    if (!_folderTree && !dest) {
        banner.style.display = "none";
        return;
    }
    banner.style.display = "block";
    pathEl.textContent = dest || "(project root)";
}

// ---------- Context menu ----------
let _ctxNode = null;

function showFolderContextMenu(x, y, node) {
    _ctxNode = node;
    const menu = document.getElementById("folderCtxMenu");
    menu.style.display = "block";
    menu.style.left = x + "px";
    menu.style.top  = y + "px";

    // Disable rename / delete on the project root
    menu.querySelector('[data-action="rename"]').style.display = node.path ? "block" : "none";
    menu.querySelector('[data-action="delete"]').style.display = node.path ? "block" : "none";

    // Adjust if it overflows the viewport
    requestAnimationFrame(() => {
        const r = menu.getBoundingClientRect();
        if (r.right > window.innerWidth)  menu.style.left = (window.innerWidth - r.width - 8) + "px";
        if (r.bottom > window.innerHeight) menu.style.top  = (y - r.height) + "px";
    });
}

function hideFolderContextMenu() {
    document.getElementById("folderCtxMenu").style.display = "none";
    _ctxNode = null;
}

function setupFolderContextMenu() {
    const menu = document.getElementById("folderCtxMenu");
    menu.querySelectorAll(".ctx-item").forEach(item => {
        item.onclick = (e) => {
            e.stopPropagation();
            const action = item.dataset.action;
            const node = _ctxNode;
            hideFolderContextMenu();
            if (!node) return;
            if (action === "create") doCreateFolder(node);
            else if (action === "rename") doRenameFolder(node);
            else if (action === "delete") doDeleteFolder(node);
            else if (action === "refresh") loadFolderTree(_treeProjectNumber);
        };
    });
    document.addEventListener("click", hideFolderContextMenu);
    document.addEventListener("contextmenu", (e) => {
        if (!e.target.closest(".folder-row") && !e.target.closest("#folderCtxMenu"))
            hideFolderContextMenu();
    });
}

// ---------- Folder name suggestion (per settings) ----------
function suggestNewFolderName() {
    const fn = _prefs.folderNaming || DEFAULT_PREFS.folderNaming;
    if (fn.mode === "default") return "New Folder";
    const d = new Date();
    const dateStr = `${pad2(d.getDate())}.${pad2(d.getMonth() + 1)}.${d.getFullYear()}`;
    if (fn.mode === "date") return dateStr;
    const prefix = (fn.customPrefix || "").trim();
    return prefix ? `${prefix} ${dateStr}` : dateStr;
}
function pad2(n) { return n < 10 ? "0" + n : "" + n; }

// ---------- Create / Rename / Delete ----------
async function doCreateFolder(parentNode) {
    const r = await dlgPrompt("שם תיקיה חדשה", suggestNewFolderName(), "צור");
    if (!r.ok || !r.value) return;

    try {
        const res = await fetch(
            `${BACKEND_BASE}/api/project/${encodeURIComponent(_treeProjectNumber)}/folders`,
            {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ parent: parentNode.path, name: r.value })
            }
        );
        if (!res.ok) {
            const err = await res.json().catch(() => ({}));
            await dlgAlert("שגיאה", err.error || ("שגיאה ביצירת תיקיה: " + res.status));
            return;
        }
        _expandedPaths.add(parentNode.path);
        await loadFolderTree(_treeProjectNumber);
    } catch (e) {
        await dlgAlert("שגיאה", e.message);
    }
}

async function doRenameFolder(node) {
    const r = await dlgPrompt("שם חדש", node.name, "שנה שם");
    if (!r.ok || !r.value || r.value === node.name) return;

    try {
        const res = await fetch(
            `${BACKEND_BASE}/api/project/${encodeURIComponent(_treeProjectNumber)}/folders`,
            {
                method: "PUT",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ path: node.path, newName: r.value })
            }
        );
        if (!res.ok) {
            const err = await res.json().catch(() => ({}));
            await dlgAlert("שגיאה", err.error || ("שגיאה בשינוי שם: " + res.status));
            return;
        }
        _selectedFolderPath = "";
        await loadFolderTree(_treeProjectNumber);
    } catch (e) {
        await dlgAlert("שגיאה", e.message);
    }
}

async function doDeleteFolder(node) {
    const r = await dlgConfirm(
        "אישור מחיקה",
        `למחוק את "${node.name}"?\n\nהמחיקה כוללת את כל התוכן ולא ניתן לבטלה.`,
        "מחק"
    );
    if (!r.ok) return;

    try {
        const res = await fetch(
            `${BACKEND_BASE}/api/project/${encodeURIComponent(_treeProjectNumber)}/folders`,
            {
                method: "DELETE",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ path: node.path, recursive: true })
            }
        );
        if (!res.ok) {
            const err = await res.json().catch(() => ({}));
            await dlgAlert("שגיאה", err.error || ("שגיאה במחיקה: " + res.status));
            return;
        }
        _selectedFolderPath = "";
        await loadFolderTree(_treeProjectNumber);
    } catch (e) {
        await dlgAlert("שגיאה", e.message);
    }
}

// =====================================================================
// Inline contact picker panel (Outlook-style sidebar layout)
// =====================================================================
let _pickerMode       = "leader";
let _pickerFolders    = [];
let _pickerContacts   = {};
let _pickerSelected   = [];
let _pickerActiveFolderId = null;

async function openContactPicker(mode) {
    _pickerMode     = mode;
    _pickerSelected = [];

    const panel = document.getElementById("contactPickerPanel");
    panel.style.display = "flex";
    document.body.classList.add("overlay-open");

    document.getElementById("cpTitle").textContent =
        mode === "leader" ? "בחירת מנהל פרויקט" : "הוספת עובדי פרויקט";
    document.getElementById("cpSearch").value = "";
    document.getElementById("cpClearSearch").style.display = "none";
    document.getElementById("cpSelectBtn").disabled = true;

    if (!_cachedPickerData) {
        document.getElementById("cpList").innerHTML =
            `<div class="cp-empty">טוען אנשי קשר...</div>`;
        try {
            const res = await fetch(`${BACKEND_BASE}/api/contacts`);
            if (!res.ok) throw new Error("שגיאת שרת " + res.status);
            _cachedPickerData = await res.json();
            if (!_allContacts.length) {
                const seen = new Set();
                _allContacts = Object.values(_cachedPickerData.contacts || {}).flat().filter(c => {
                    if (seen.has(c.email)) return false; seen.add(c.email); return true;
                });
            }
        } catch (e) {
            document.getElementById("cpList").innerHTML =
                `<div class="cp-empty" style="color:red">שגיאה: ${e.message}</div>`;
            return;
        }
    }

    _pickerFolders        = _cachedPickerData.folders  || [];
    _pickerContacts       = _cachedPickerData.contacts || {};
    _pickerActiveFolderId = _pickerFolders[0]?.id ?? null;

    renderPickerFolders();
    renderPickerList();
}

function renderPickerFolders() {
    const tree = document.getElementById("cpFolderTree");
    tree.innerHTML = "";
    _pickerFolders.forEach(f => {
        const count = (_pickerContacts[f.id] || []).length;
        const el = document.createElement("div");
        el.className = "cp-folder" + (f.id === _pickerActiveFolderId ? " active" : "");
        el.style.paddingLeft = (10 + (f.parentId ? 14 : 0)) + "px";
        el.innerHTML = `${escHtml(f.displayName)}<span class="cnt">(${count})</span>`;
        el.onclick = () => {
            _pickerActiveFolderId = f.id;
            renderPickerFolders();
            renderPickerList();
        };
        tree.appendChild(el);
    });
}

function closeContactPicker() {
    document.getElementById("contactPickerPanel").style.display = "none";
    document.body.classList.remove("overlay-open");
    _pickerSelected = [];
}

function renderPickerList() {
    const list     = document.getElementById("cpList");
    const query    = document.getElementById("cpSearch").value.toLowerCase().trim();
    const multiSel = _pickerMode === "employee";

    let pool;
    if (query) {
        const seen = new Set();
        pool = Object.values(_pickerContacts).flat().filter(c => {
            if (seen.has(c.email)) return false;
            seen.add(c.email);
            return c.displayName.toLowerCase().includes(query) ||
                   c.email.toLowerCase().includes(query);
        });
    } else {
        pool = _pickerContacts[_pickerActiveFolderId] || [];
    }

    if (!pool.length) {
        list.innerHTML = `<div class="cp-empty">אין אנשי קשר</div>`;
        return;
    }

    list.innerHTML = "";
    pool.forEach(c => {
        const isSel = _pickerSelected.some(s => s.email === c.email);
        const row   = document.createElement("div");
        row.className = "cp-row" + (isSel ? " selected" : "");

        if (multiSel) {
            const cb    = document.createElement("input");
            cb.type     = "checkbox";
            cb.checked  = isSel;
            cb.tabIndex = -1;
            row.appendChild(cb);
        }

        const avatar = document.createElement("div");
        avatar.className = "cp-avatar";
        avatar.style.background = avatarColor(c.email || c.displayName);
        avatar.textContent = avatarInitials(c.displayName);
        row.appendChild(avatar);

        const info = document.createElement("div");
        info.className = "cp-info";
        info.innerHTML =
            `<div class="cp-name">${escHtml(c.displayName)}</div>` +
            `<div class="cp-email">${escHtml(c.email)}</div>`;
        row.appendChild(info);

        row.onclick = () => {
            const idx = _pickerSelected.findIndex(s => s.email === c.email);
            if (idx >= 0) {
                _pickerSelected.splice(idx, 1);
                row.classList.remove("selected");
                const cb = row.querySelector("input"); if (cb) cb.checked = false;
            } else {
                if (!multiSel) _pickerSelected = [];
                _pickerSelected.push(c);
                if (!multiSel) { confirmPickerSelection(); return; }
                row.classList.add("selected");
                const cb = row.querySelector("input"); if (cb) cb.checked = true;
            }
            updatePickerBtn();
        };
        list.appendChild(row);
    });
}

// Generate consistent avatar color + initials (Outlook-style)
function avatarColor(seed) {
    let hash = 0;
    for (let i = 0; i < (seed || "").length; i++) {
        hash = (seed.charCodeAt(i) + ((hash << 5) - hash)) | 0;
    }
    const palette = ["#7B68EE","#FF6347","#3CB371","#FFA500","#1E90FF",
                     "#9370DB","#20B2AA","#FF69B4","#5F9EA0","#DA70D6","#CD5C5C","#4682B4"];
    return palette[Math.abs(hash) % palette.length];
}

function avatarInitials(name) {
    if (!name) return "?";
    const parts = name.trim().split(/\s+/).filter(Boolean);
    if (!parts.length) return "?";
    if (parts.length === 1) return parts[0][0].toUpperCase();
    return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
}

function updatePickerBtn() {
    const btn = document.getElementById("cpSelectBtn");
    btn.disabled    = _pickerSelected.length === 0;
    btn.textContent = _pickerMode === "employee" && _pickerSelected.length > 0
        ? `הוסף (${_pickerSelected.length})` : "בחר";
}

function confirmPickerSelection() {
    if (_pickerMode === "leader" && _pickerSelected.length > 0) {
        setProjectLeader(_pickerSelected[0]);
    } else if (_pickerMode === "employee") {
        _pickerSelected.forEach(addEmployee);
    }
    closeContactPicker();
}

// =====================================================================
// Main SaveAsPDF action
// =====================================================================
async function onSaveAsPdf() {
    const status = document.getElementById("status");
    status.innerText = "";

    if (!validateMarkingOptions()) {
        status.innerText = "יש לבחור לפחות אופן סימון אחד להודעה";
        return;
    }

    const projectId = document.getElementById("projectId").value.trim();
    if (!projectId) { status.innerText = "יש להזין מספר פרויקט"; return; }

    const projectName    = document.getElementById("projectName").value.trim();
    const leaderInput    = document.getElementById("projectLeader");
    const projectLeader  = leaderInput.dataset.email || leaderInput.value.trim();

    try {
        status.innerText = "אוסף נתוני הודעה…";
        const mailData = await collectMailData();

        const doStamp    = document.getElementById("optStamp").checked;
        const doCategory = document.getElementById("optCategory").checked;
        const fwdCbEl    = document.getElementById("optForwardToLeader");
        const willForward = !!(_prefs.dispatchMode && fwdCbEl?.checked);

        status.innerText = "שולח נתונים לשרת…";
        const result = await sendToBackend({
            projectId, projectName, projectLeader,
            employees: _employees,
            email: mailData.email,
            attachments: mailData.attachments,
            destinationFolder: resolveSaveDestination(),
            stamp: doStamp ? {
                includeProjectId:   !!_prefs.stampFields?.projectId,
                includeProjectName: !!_prefs.stampFields?.projectName,
                includeLeader:      !!_prefs.stampFields?.leader,
                includeDate:        !!_prefs.stampFields?.date,
                includeUser:        !!_prefs.stampFields?.user,
                includeEmployees:   !!_prefs.stampFields?.employees,
                includeAttachments: !!_prefs.stampFields?.attachments,
                notes:              _prefs.stampNotes || "",
                template:           _prefs.stampTemplate || "",
                forwarded:          willForward,
                forwardedTo:        willForward ? stripEmail(projectLeader) : "",
                userName:           getCurrentUserDisplay(),
                attachmentNames:    mailData.allAttachmentNames || []
            } : null
        });

        // Surface PDF generation outcome (server may have fallen back to .html)
        if (result?.pdf && result.pdf.pdfCreated === false) {
            const reason = result.pdf.fallbackReason || "המרה ל‑PDF נכשלה";
            await dlgAlert(
                "לא נוצר קובץ PDF",
                `נשמר קובץ HTML חלופי במקום ה‑PDF.\n\nסיבה: ${reason}\n\nנתיב: ${result.pdf.fullPath || ""}`
            );
        }

        const warnings = await markMessageProcessed(
            Office.context.mailbox.item,
            projectId, projectName, projectLeader,
            result, doStamp, doCategory
        );

        // Dispatch: forward to project leader if requested
        const fwdCb = document.getElementById("optForwardToLeader");
        const doForward = _prefs.dispatchMode && fwdCb?.checked;
        if (doForward) {
            const fr = await forwardToProjectLeader(
                Office.context.mailbox.item,
                projectId, projectName, projectLeader, mailData
            );
            if (!fr.ok) warnings.push("העברה למנהל פרויקט: " + fr.reason);
        }
        // Reset the forward checkbox if not pinned to "default checked"
        if (fwdCb && !_prefs.forwardToLeaderByDefault) fwdCb.checked = false;

        const savedName = result?.pdf?.fileName || "";
        if (warnings.length === 0) {
            status.innerText = savedName
                ? `נשמר ✅ ${savedName}`
                : "ההודעה נשמרה בהצלחה ✅";
        } else {
            status.innerText = savedName
                ? `נשמר ✅ ${savedName} (עם אזהרות)`
                : "ההודעה נשמרה ✅ (עם אזהרות)";
            await dlgAlert("ההודעה נשמרה — חלק מהפעולות לא בוצעו",
                warnings.map(w => "• " + w).join("\n"));
        }
    } catch (err) {
        console.error(err);
        status.innerText = "שגיאה בשמירת ההודעה ❌";
        await dlgAlert("שגיאה בשמירה", err?.message || String(err));
    }
}

// =====================================================================
// Validation / Collect / Send
// =====================================================================
function validateMarkingOptions() {
    return document.getElementById("optStamp").checked ||
           document.getElementById("optCategory").checked;
}

async function collectMailData() {
    const item = Office.context.mailbox.item;

    const bodyHtml = await new Promise((resolve, reject) =>
        item.body.getAsync(Office.CoercionType.Html,
            r => r.status === Office.AsyncResultStatus.Succeeded ? resolve(r.value) : reject(r.error))
    );

    const checkedIds = new Set(
        [...document.querySelectorAll(".att-check:checked")].map(cb => cb.dataset.id)
    );

    const attachments = [];
    for (const att of item.attachments) {
        if (!checkedIds.has(att.id)) continue;
        const content = await new Promise((resolve, reject) =>
            item.getAttachmentContentAsync(att.id,
                r => r.status === Office.AsyncResultStatus.Succeeded ? resolve(r.value) : reject(r.error))
        );
        attachments.push({ name: att.name, contentType: att.contentType, size: att.size, base64: content.content });
    }

    // All attachment names from the original message (independent of save selection),
    // used only for stamping the PDF.
    const allAttachmentNames = (item.attachments || [])
        .filter(a => !a.isInline)
        .map(a => a.name);

    return {
        email: { subject: item.subject, from: item.from?.emailAddress || "", bodyHtml },
        attachments,
        allAttachmentNames
    };
}

async function sendToBackend(payload) {
    const res = await fetch(`${BACKEND_BASE}/api/saveaspdf`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload)
    });
    if (!res.ok) {
        const text = await res.text().catch(() => "");
        let detail = text;
        try { detail = JSON.parse(text).error || JSON.parse(text).title || text; } catch {}
        throw new Error(`HTTP ${res.status}: ${detail || res.statusText}`);
    }
    return await res.json();
}

// =====================================================================
// Mark message after save
// =====================================================================
// Feature-detect post-save actions and report any issues without throwing.
// Returns an array of human-readable warnings (empty = all good).
//
// Read-mode markers used (no body.setAsync, no elevated perms required):
//   1. notificationMessages.addAsync - visible info banner pinned to the email
//   2. loadCustomPropertiesAsync     - invisible machine-readable metadata
//   3. categories.addAsync           - best-effort (may be denied without elevated perm)
async function markMessageProcessed(item, projectId, projectName, projectLeader, result, doStamp, doCategory) {
    const warnings = [];
    const dateStr  = new Date().toLocaleString("he-IL");

    // ---------- Persist invisible metadata (so tryRestoreProjectInfo can read it back) ----------
    if (typeof item?.loadCustomPropertiesAsync === "function") {
        await new Promise(resolve => {
            try {
                item.loadCustomPropertiesAsync(r => {
                    if (r.status !== Office.AsyncResultStatus.Succeeded) {
                        warnings.push("מאפיינים מותאמים: " + (r.error?.message || "נכשל"));
                        resolve();
                        return;
                    }
                    const props = r.value;
                    props.set("saveAsPdfProcessed",     "true");
                    props.set("saveAsPdfProjectId",     projectId || "");
                    props.set("saveAsPdfProjectName",   projectName || "");
                    props.set("saveAsPdfProjectLeader", projectLeader || "");
                    props.set("saveAsPdfDate",          new Date().toISOString());
                    props.saveAsync(r2 => {
                        if (r2.status !== Office.AsyncResultStatus.Succeeded) {
                            warnings.push("שמירת מאפיינים: " + (r2.error?.message || "נכשלה"));
                        }
                        resolve();
                    });
                });
            } catch (e) {
                warnings.push("מאפיינים מותאמים: " + (e.message || "לא נתמך"));
                resolve();
            }
        });
    }

    // ---------- Visible "stamp" -> notification banner (read-mode safe) ----------
    if (doStamp) {
        if (typeof item?.notificationMessages?.addAsync !== "function") {
            warnings.push("הודעת מצב אינה זמינה בסביבה זו");
        } else {
            const f = _prefs.stampFields || DEFAULT_PREFS.stampFields;
            const parts = [];
            if (f.projectId   && projectId)     parts.push(`פרויקט ${projectId}`);
            if (f.projectName && projectName)   parts.push(projectName);
            if (f.leader      && projectLeader) parts.push(`מנהל: ${stripEmail(projectLeader)}`);
            if (f.date)                         parts.push(dateStr);
            if (_prefs.stampNotes?.trim())      parts.push(_prefs.stampNotes.trim());

            const message = `📁 SaveAsPDF • ${parts.join(" • ")}`;
            const truncated = message.length > 150 ? message.substring(0, 147) + "..." : message;

            // The icon parameter must be a 4–32 char manifest-defined icon NAME,
            // not a URL. The unified JSON manifest exposes "color"/"outline" at top
            // level — try those first; fall back to ErrorMessage type (no icon).
            const tryAdd = (opts) => new Promise(res => {
                try {
                    item.notificationMessages.addAsync("saveAsPdfStamp", opts, r => {
                        res(r.status === Office.AsyncResultStatus.Succeeded
                            ? null
                            : (r.error?.message || "נכשלה"));
                    });
                } catch (e) {
                    res(e.message || "לא נתמך");
                }
            });

            let err = await tryAdd({
                type:       Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
                message:    truncated,
                icon:       "color",   // matches manifest.icons.color
                persistent: true
            });
            if (err) {
                // Fallback: ErrorMessage type needs no icon and still displays a banner.
                err = await tryAdd({
                    type:    Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
                    message: truncated
                });
            }
            if (err) console.warn("notificationMessages.addAsync failed:", err);
        }
    }

    // ---------- Apply category (best-effort) ----------
    if (doCategory) {
        const cat = (_prefs.categoryName || DEFAULT_PREFS.categoryName).trim();
        if (!cat) {
            warnings.push("שם קטגוריה ריק");
        } else if (typeof item?.categories?.addAsync !== "function") {
            warnings.push("שיוך קטגוריה אינו זמין במצב זה");
        } else {
            await new Promise(resolve => {
                try {
                    item.categories.addAsync([cat], r => {
                        if (r.status !== Office.AsyncResultStatus.Succeeded) {
                            warnings.push("שיוך קטגוריה: " + (r.error?.message || "נדרשת הרשאה מורחבת"));
                        }
                        resolve();
                    });
                } catch (e) {
                    warnings.push("שיוך קטגוריה: " + (e.message || "לא נתמך"));
                    resolve();
                }
            });
        }
    }

    return warnings;
}

// "Name <email@x.co>" -> "email@x.co" (or just the original if no <...>)
function stripEmail(text) {
    if (!text) return "";
    const m = String(text).match(/<([^>]+)>/);
    return m ? m[1].trim() : String(text).trim();
}

function getCurrentUserDisplay() {
    try {
        const up = Office.context.mailbox.userProfile;
        return up?.displayName || up?.emailAddress || "";
    } catch { return ""; }
}

// Open a new compose window pre-populated with the project leader as recipient
// and the original message body forwarded.
async function forwardToProjectLeader(item, projectId, projectName, projectLeader, mailData) {
    const leaderEmail = stripEmail(projectLeader);
    if (!leaderEmail.includes("@")) {
        return { ok: false, reason: "מנהל פרויקט לא הוגדר או אין אימייל תקין" };
    }
    if (typeof Office.context.mailbox.displayNewMessageForm !== "function") {
        return { ok: false, reason: "פתיחת חלון הודעה חדשה אינה נתמכת" };
    }

    const subject = (item.subject || "").startsWith("FW:")
        ? item.subject
        : `FW: ${item.subject || ""}`;

    const intro = `
<div dir="rtl" style="font-family:Segoe UI,Arial,sans-serif">
  <p><b>הועבר אוטומטית מתחנת הפצה של SaveAsPDF</b></p>
  <table style="border-collapse:collapse;font-size:13px">
    <tr><td style="color:#666;padding:2px 8px 2px 0">פרויקט:</td><td>${projectId || ""} ${projectName ? "(" + projectName + ")" : ""}</td></tr>
    <tr><td style="color:#666;padding:2px 8px 2px 0">תאריך עיבוד:</td><td>${new Date().toLocaleString("he-IL")}</td></tr>
    <tr><td style="color:#666;padding:2px 8px 2px 0">שולח מקורי:</td><td>${item.from?.displayName || ""} &lt;${item.from?.emailAddress || ""}&gt;</td></tr>
  </table>
  <hr style="margin:12px 0;border:none;border-top:1px solid #ccc"/>
</div>
`;

    try {
        Office.context.mailbox.displayNewMessageForm({
            toRecipients: [leaderEmail],
            subject: subject,
            htmlBody: intro + (mailData?.email?.bodyHtml || "")
        });
        return { ok: true };
    } catch (e) {
        return { ok: false, reason: e.message || "פתיחת חלון נכשלה" };
    }
}

function buildStampHtml(ctx) {
    // Advanced template overrides field toggles when non-empty
    if (_prefs.stampTemplate?.trim()) {
        return _prefs.stampTemplate
            .replace(/\{\{projectId\}\}/g,   ctx.projectId   || "")
            .replace(/\{\{projectName\}\}/g, ctx.projectName || "")
            .replace(/\{\{leader\}\}/g,     ctx.projectLeader || "")
            .replace(/\{\{date\}\}/g,       ctx.date         || "")
            .replace(/\{\{notes\}\}/g,      _prefs.stampNotes || "");
    }

    const f = _prefs.stampFields || DEFAULT_PREFS.stampFields;
    const lines = [];
    if (f.projectId   && ctx.projectId)     lines.push(`מספר פרויקט: ${ctx.projectId}`);
    if (f.projectName && ctx.projectName)   lines.push(`שם פרויקט: ${ctx.projectName}`);
    if (f.leader      && ctx.projectLeader) lines.push(`מנהל/ת פרויקט: ${ctx.projectLeader}`);
    if (f.date)                             lines.push(`תאריך: ${ctx.date}`);
    if (_prefs.stampNotes?.trim())          lines.push(_prefs.stampNotes.trim());

    if (lines.length === 0) return "";
    return `\n<div style="border:1px solid #888;padding:8px;margin-bottom:10px">\n  <b>SaveAsPDF</b><br/>\n  ${lines.join("<br/>\n  ")}\n</div>`;
}

// =====================================================================
// Auto-restore from hidden comment
// =====================================================================
function tryRestoreProjectInfo() {
    const item = Office.context.mailbox.item;
    if (typeof item?.loadCustomPropertiesAsync !== "function") return;

    item.loadCustomPropertiesAsync(r => {
        if (r.status !== Office.AsyncResultStatus.Succeeded) return;
        const props  = r.value;
        const id     = props.get("saveAsPdfProjectId");
        const name   = props.get("saveAsPdfProjectName");
        const leader = props.get("saveAsPdfProjectLeader");

        if (id)   document.getElementById("projectId").value   = id;
        if (name) document.getElementById("projectName").value = name;
        if (leader) {
            const input = document.getElementById("projectLeader");
            input.value = leader;
            input.dataset.email = stripEmail(leader);
        }
    });
}

// =====================================================================
// Tabs
// =====================================================================
function initializeTabs() {
    const buttons = document.querySelectorAll(".tab-btn");
    buttons.forEach(btn => btn.addEventListener("click", () => {
        buttons.forEach(b => b.classList.remove("active"));
        document.querySelectorAll(".tab-panel").forEach(p => p.classList.remove("active"));
        btn.classList.add("active");
        document.getElementById(btn.dataset.tab).classList.add("active");
    }));
}
