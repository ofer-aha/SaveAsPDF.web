/* global Office, __APP_VERSION__, __BACKEND_BASE__ */

// Build-time injected version (from package.json via webpack DefinePlugin)
const APP_VERSION = (typeof __APP_VERSION__ !== "undefined") ? __APP_VERSION__ : "0.0.0";
// Live version fetched from /api/info — single source of truth; updated on load
let _liveVersion = APP_VERSION;

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
// "Backend is sleeping" fun screen — shown when the server can't be reached
// =====================================================================
function isNetworkError(err) {
    // sendToBackend rejects with a Hebrew message, so matching only English fetch
    // wording meant a network failure during Save skipped the "server is sleeping"
    // screen every other path shows and produced a generic red error instead.
    return err instanceof TypeError ||
           /failed to fetch|networkerror|load failed|abort|timeout/i.test(err?.message || "") ||
           /שגיאת רשת|לא ניתן להגיע לשרת|פסק זמן/.test(err?.message || "");
}

function showSleepingScreen() {
    if (document.getElementById("sleepOverlay")) return;
    const div = document.createElement("div");
    div.id  = "sleepOverlay";
    div.dir = "rtl";
    div.style.cssText =
        "position:fixed;inset:0;z-index:9999;background:linear-gradient(180deg,#1a2340 0%,#2c3a63 100%);" +
        "display:flex;flex-direction:column;align-items:center;justify-content:center;" +
        "text-align:center;padding:24px;font-family:'Segoe UI',Arial,sans-serif;color:#f5f2e8";
    div.innerHTML = `
      <svg width="210" height="170" viewBox="0 0 210 170" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
        <!-- moon + stars -->
        <circle cx="178" cy="26" r="14" fill="#f7e8a4"/>
        <circle cx="172" cy="22" r="12" fill="#2c3a63"/>
        <circle cx="30" cy="18" r="2" fill="#f7e8a4"/>
        <circle cx="60" cy="34" r="1.5" fill="#f7e8a4"/>
        <circle cx="140" cy="14" r="1.7" fill="#f7e8a4"/>
        <!-- Zzz -->
        <text x="150" y="66" font-size="20" fill="#9fd0ff" font-weight="bold">Z</text>
        <text x="163" y="52" font-size="14" fill="#9fd0ff" font-weight="bold">z</text>
        <text x="173" y="42" font-size="10" fill="#9fd0ff" font-weight="bold">z</text>
        <!-- bed -->
        <rect x="18" y="128" width="174" height="12" rx="6" fill="#7a5a3a"/>
        <rect x="22" y="138" width="10" height="20" rx="3" fill="#7a5a3a"/>
        <rect x="178" y="138" width="10" height="20" rx="3" fill="#7a5a3a"/>
        <!-- pillow -->
        <rect x="26" y="106" width="44" height="24" rx="10" fill="#e8e4f5"/>
        <!-- server with a sleepy face -->
        <rect x="52" y="78" width="64" height="52" rx="8" fill="#8fa3c8" stroke="#5c6f94" stroke-width="2"/>
        <circle cx="106" cy="88" r="3" fill="#65d06e"/>
        <rect x="60" y="116" width="40" height="4" rx="2" fill="#5c6f94"/>
        <!-- closed eyes -->
        <path d="M66 98 q5 6 10 0" stroke="#2c3a63" stroke-width="2.5" fill="none" stroke-linecap="round"/>
        <path d="M88 98 q5 6 10 0" stroke="#2c3a63" stroke-width="2.5" fill="none" stroke-linecap="round"/>
        <!-- sleepy mouth -->
        <circle cx="82" cy="110" r="3" fill="#2c3a63"/>
        <!-- nightcap -->
        <path d="M52 82 q10 -22 46 -14 l-8 12 z" fill="#d9534f"/>
        <circle cx="100" cy="66" r="5" fill="#f7e8a4"/>
        <!-- blanket over the server -->
        <path d="M44 130 q62 -26 148 -4 l0 4 z" fill="#4f74b3"/>
        <path d="M44 130 q62 -26 148 -4" stroke="#3c5c96" stroke-width="2" fill="none"/>
      </svg>
      <div style="font-size:22px;font-weight:600;margin-top:14px">אני ישן עכשיו... 😴</div>
      <div style="font-size:16px;margin-top:8px">תגידו לעופר שלא בא לי לעבוד 🙃</div>
      <div style="font-size:11px;opacity:.65;margin-top:10px">(השרת לא זמין כרגע — נסו שוב עוד כמה דקות)</div>
      <button id="sleepRetryBtn" type="button"
        style="margin-top:18px;padding:8px 22px;font-size:14px;border:none;border-radius:20px;
               background:#f7e8a4;color:#1a2340;font-weight:600;cursor:pointer">
        נסה להעיר אותי ⏰
      </button>`;
    document.body.appendChild(div);
    document.getElementById("sleepRetryBtn").onclick = async () => {
        const btn = document.getElementById("sleepRetryBtn");
        btn.disabled = true; btn.textContent = "מנסה להעיר... ⏰";
        try {
            const ctrl = new AbortController();
            const t = setTimeout(() => ctrl.abort(), 6000);
            const res = await fetch(`${BACKEND_BASE}/api/info`, { signal: ctrl.signal, cache: "no-store" });
            clearTimeout(t);
            if (!res.ok) throw new Error("still down");
            div.remove();
            location.reload();
        } catch {
            btn.textContent = "עדיין ישן... נסה שוב 😴";
            btn.disabled = false;
        }
    };
}

// =====================================================================
// Office initialization
// =====================================================================
Office.onReady(() => {
    console.log("SaveAsPDF taskpane loaded — v" + APP_VERSION);

    // Credit footer version badge — seed with bundle version, then update from /api/info
    const versionEl = document.getElementById("appVersion");
    if (versionEl) {
        versionEl.textContent = APP_VERSION;
        fetch(`${BACKEND_BASE}/api/info`)
            .then(r => r.json())
            .then(j => { if (j.version) { versionEl.textContent = j.version; _liveVersion = j.version; } })
            .catch(() => showSleepingScreen());
    }

    document.getElementById("saveBtn").onclick      = onSaveAsPdf;
    document.getElementById("pickLeaderBtn").onclick = () => openContactPicker("leader");
    document.getElementById("bugReportBtn").onclick  = onBugReport;

    // Settings panel
    document.getElementById("openSettingsBtn").onclick = openSettingsPanel;

    // About panel
    document.getElementById("aboutBtn").onclick      = openAboutPanel;
    document.getElementById("aboutCloseBtn").onclick = closeAboutPanel;

    // Pin support — when pinned, auto-reload email data on item change
    if (Office.context.mailbox.addHandlerAsync) {
        Office.context.mailbox.addHandlerAsync(
            Office.EventType.ItemChanged,
            () => { if (Office.context.mailbox.item) loadCurrentItem(); },
            () => {}
        );
    }

    // Help link → /help on the backend, forced into the user's default browser.
    // window.open with _blank is what Outlook routes to the system shell.
    const helpLink = document.getElementById("helpLink");
    if (helpLink) {
        const helpUrl = `${BACKEND_BASE}/help/index.html`;
        helpLink.href = helpUrl;
        helpLink.onclick = (e) => {
            e.preventDefault();
            window.open(helpUrl, "_blank", "noopener,noreferrer");
        };
    }
    document.getElementById("spCloseBtn").onclick      = closeSettingsPanel;
    document.getElementById("spCancelBtn").onclick     = closeSettingsPanel;
    document.getElementById("spSaveBtn").onclick       = saveSettingsPanel;
    document.getElementById("spNewCatBtn").onclick      = showNewCategoryForm;
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
    applyAddSelf();          // auto-add logged-in user if setting is on
    preloadContacts(); // background – feeds autocomplete & speeds up picker

    loadPrefs();
    applyPrefsToCheckboxes();
    updateCategoryLabel();
    updateDestBanner();

    fetchAdminPolicy();         // populates _adminPolicy, then re-applies prefs
    checkAdminAccess();         // shows admin link if user is in the admin AD group
    startBackendHealthCheck();

    // Pinned task panes stay open across messages instead of reloading the
    // page, so we must detect message changes ourselves and refresh the UI.
    registerItemChangedHandler();
});

// =====================================================================
// Item-changed handling (required for pinned task panes — see manifest.json
// "pinnable": true). When the pane isn't pinned, Outlook simply reloads the
// page on message switch and this handler is a harmless no-op in practice;
// when it IS pinned, the page survives across messages and every piece of
// per-message state must be reset and reloaded from scratch.
// =====================================================================
function registerItemChangedHandler() {
    if (typeof Office?.context?.mailbox?.addHandlerAsync !== "function") return;
    Office.context.mailbox.addHandlerAsync(
        Office.EventType.ItemChanged,
        reloadForCurrentItem,
        (r) => {
            if (r.status !== Office.AsyncResultStatus.Succeeded) {
                console.warn("Failed to register ItemChanged handler:", r.error);
            }
        }
    );
}

// Re-reads everything that depends on the message currently shown in the
// pane. Mirrors the per-item initialization done in Office.onReady.
function reloadForCurrentItem() {
    // Close overlays — they may be showing data tied to the previous message
    closeSettingsPanel();
    closeContactPicker();
    hideFolderContextMenu();

    // Reset project fields (will be re-populated by tryRestoreProjectInfo
    // from the new item's custom properties, or by the user)
    const idInput     = document.getElementById("projectId");
    const nameInput   = document.getElementById("projectName");
    const leaderInput = document.getElementById("projectLeader");
    idInput.value     = "";
    nameInput.value   = "";
    leaderInput.value = "";
    delete leaderInput.dataset.email;

    // Reset employees list
    _employees.length = 0;
    renderEmployees();

    // The unticked-attachment memory is per message, not per session.
    _uncheckedAttachmentIds = new Set();

    // Reset folder tree / destination banner back to "no project loaded" state
    _folderTree         = null;
    _treeProjectNumber  = "";
    _selectedFolderPath = "";
    document.getElementById("folderTree").innerHTML =
        `<div style="color:#888;font-size:13px;padding:10px 0">הזן מספר פרויקט כדי לטעון את עץ התיקיות</div>`;
    updateDestBanner();

    // Hide the saved-message info card until the new item's metadata is read
    renderSavedInfo(null);

    // Clear status / progress indicators left over from the previous message
    document.getElementById("status").innerText = "";
    document.getElementById("uploadProgress").style.display = "none";

    // Reload data scoped to the new message
    loadAttachments();
    tryRestoreProjectInfo();

    // Re-apply preference defaults (stamp / category / forward-to-leader…)
    applyPrefsToCheckboxes();
    updateCategoryLabel();

    // Auto-set current user as project leader if that setting is enabled
    // and the new message doesn't already have one
    applyAddSelf();
}

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
    input.addEventListener("keydown", async (e) => {
        if (e.key !== "Enter") return;
        e.preventDefault();
        clearTimeout(_projectLookupTimer);
        const exists = await loadProjectByNumber();
        if (!exists) {
            document.getElementById("projectName").focus();
            document.getElementById("projectName").select();
        }
    });
}

// Returns true if the project exists, false if not found, undefined on error.
async function loadProjectByNumber() {
    const num = document.getElementById("projectId").value.trim();
    const status = document.getElementById("status");
    if (!num) return undefined;

    try {
        const res = await fetch(`${BACKEND_BASE}/api/project/${encodeURIComponent(num)}`);
        if (!res.ok) {
            const err = await res.json().catch(() => ({}));
            status.innerText = err.error || `שגיאת שרת ${res.status}`;
            return undefined;
        }
        const data = await res.json();

        if (!data.exists) {
            status.innerText = `הפרויקט לא נמצא בתיקייה: ${data.expectedPath || ""}`;
            _folderTree = null;
            _selectedFolderPath = "";
            _treeProjectNumber  = "";
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
                        return false;
                    }
                    const created = await cr.json();
                    status.innerText = `הפרויקט נוצר בנתיב: ${created.folderPath}`;
                    // Re-trigger the lookup so the tree + project data load
                    loadProjectByNumber();
                } catch (e) {
                    await dlgAlert("שגיאה ביצירת פרויקט", e.message);
                }
            }
            return false;
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
                displayName: e.displayName || enrichDisplayName(e.email),
                email: e.email,
                isLeader: !!e.isLeader
            }));

            const leader = _employees.find(e => e.isLeader);
            if (leader) {
                const lin = document.getElementById("projectLeader");
                // Show the human name (<FirstName> <LastName> from the employees XML);
                // the address stays in dataset.email for sending/forwarding.
                lin.value = (leader.displayName || "").trim() || leader.email;
                lin.dataset.email = leader.email;
            }
            renderEmployees();
        }
        applyAddSelf();
        return true;
    } catch (e) {
        console.warn("Project lookup failed:", e);
        return undefined;
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

// Look up a display name from the preloaded contacts cache by email.
// Returns the name string, or "" if not found / contacts not yet loaded.
function enrichDisplayName(email) {
    if (!email) return "";
    return _allContacts.find(c => c.email === email)?.displayName || "";
}

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
// Returns the active sig-image threshold in bytes.
// Admin policy overrides the user pref when set (non-null).
function effectiveSigImgThreshold() {
    const forced = _adminAttachmentPolicy?.sigImgThreshold;
    if (forced !== null && forced !== undefined) return forced;
    return _prefs.sigImgThreshold ?? 8192;
}

function isSigImage(att) {
    const t = effectiveSigImgThreshold();
    if (t <= 0) return false;
    const isImg = /^image\//i.test(att.contentType || '') ||
                  /\.(png|jpe?g|gif|bmp|ico|tiff?|webp|svg)$/i.test(att.name || '');
    return isImg && att.size > 0 && att.size <= t;
}

// Ids the user explicitly unticked for the current message. loadAttachments is
// re-run whenever the backend health check flips down->up and after the policy
// fetch, and it used to hard-code checked on every box - so a transient server
// blip silently re-selected attachments the user had deliberately excluded, and
// they were written to the project folder on the next save.
let _uncheckedAttachmentIds = new Set();

function rememberAttachmentSelection() {
    document.querySelectorAll(".att-check").forEach(cb => {
        if (cb.checked) _uncheckedAttachmentIds.delete(cb.dataset.id);
        else            _uncheckedAttachmentIds.add(cb.dataset.id);
    });
}

function loadAttachments() {
    const item      = Office.context.mailbox.item;
    const container = document.getElementById("attachmentsTab");

    // Capture the current state before the list is rebuilt.
    rememberAttachmentSelection();

    if (!item || !item.attachments || item.attachments.length === 0) {
        container.innerHTML = "<p>אין קבצים מצורפים</p>";
        return;
    }

    // Inline attachments (cid: body embeds) are not saved separately
    const nonInline = item.attachments.filter(a => !a.isInline);
    const visible   = nonInline.filter(a => !isSigImage(a));
    const hidden    = nonInline.filter(a =>  isSigImage(a));

    if (visible.length === 0 && hidden.length === 0) {
        container.innerHTML = "<p>אין קבצים מצורפים</p>";
        return;
    }

    let html = visible.map(att => `
        <label style="display:flex;align-items:center;gap:8px;margin-top:8px">
            <input type="checkbox" class="att-check" data-id="${escHtml(att.id)}" ${_uncheckedAttachmentIds.has(att.id) ? "" : "checked"} />
            <span>${escHtml(att.name)} <span style="color:#888;font-size:12px">(${escHtml(formatSize(att.size))})</span></span>
        </label>
    `).join("");

    if (hidden.length > 0) {
        const kb = Math.round(effectiveSigImgThreshold() / 1024);
        html += `<div style="color:#aaa;font-size:11px;margin-top:6px">
            (${hidden.length} תמונת חתימה הוסתרה — מתחת ל-${kb} KB)
        </div>`;
    }

    container.innerHTML = html;
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
                       user: true, employees: true, attachments: true,
                       from: true, to: true, cc: true, sent: true, received: true, subject: true },
    stampNotes:      "",
    stampTemplate:   "",
    defaultCategory: true,
    categoryName:    "SaveAsPDF – Processed",
    folderNaming:    { mode: "default", customPrefix: "" },   // default | date | custom
    dispatchMode:    false,                                    // enables forward-to-leader feature
    forwardToLeaderByDefault: false,                           // pre-checks the forward checkbox each time
    addSelfAsEmployee: false,                                  // auto-add current Outlook user to employees
    sigImgThreshold:   8192,                                   // hide image attachments ≤ this size (bytes); 0 = off
    defaultSubfolders: {
        list:       ["מכתבים", "התקבל", "אישור ציוד"],
        selected:   "",      // "" = no default sub-folder
        appendDate: false    // append today's date as inner sub-folder
    },
    // PDF output settings the user may control (subject to admin lock policy).
    pdf: {
        pageSize:        "A4",     // A4 | Letter | Legal | A3
        landscape:       false,
        marginTopCm:     2.54,
        marginBottomCm:  2.54,
        marginLeftCm:    2.54,
        marginRightCm:   2.54,
        printBackground: true,
        // Append PDF attachments into the generated PDF. Off by default: the
        // attachments are always written as separate files anyway, so the
        // combined copy is opt-in. An admin lock (PdfPolicy) overrides this.
        mergePdfAttachments: false
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
            pdf:          { ...DEFAULT_PREFS.pdf,          ...(raw.pdf          || {}) },
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

// ---------- Admin policy (server-enforced overrides) ----------
// Stamp policy: { defaultStamp: bool|null, ... }
let _adminPolicy = {};
// Attachment policy: { sigImgThreshold: number|null }  null = user controls
let _adminAttachmentPolicy = {};
// PDF policy + admin defaults: { settings: {...}, policy: { pageSize, orientation, margins, printBackground } }
let _adminPdf = { settings: {}, policy: {} };

// Mapping from policy key -> taskpane checkbox id (used to lock the UI).
const POLICY_FIELD_MAP = {
    defaultStamp:        "spDefaultStamp",
    includeProjectId:    "spFieldProjectId",
    includeProjectName:  "spFieldProjectName",
    includeLeader:       "spFieldLeader",
    includeDate:         "spFieldDate",
    includeUser:         "spFieldUser",
    includeEmployees:    "spFieldEmployees",
    includeAttachments:  "spFieldAttachments",
    includeFrom:         "spFieldFrom",
    includeTo:           "spFieldTo",
    includeCc:           "spFieldCc",
    includeSent:         "spFieldSent",
    includeReceived:     "spFieldReceived",
    includeSubject:      "spFieldSubject"
};

// Folder DELETE is admin-gated server-side. Showing the menu item to everyone led
// users through a "cannot be undone" confirmation to a bare "401".
let _isAdmin = false;

async function checkAdminAccess() {
    try {
        const email = Office.context.mailbox?.userProfile?.emailAddress;
        if (!email) return;
        const res = await fetch(`${BACKEND_BASE}/api/policy/is-admin?email=${encodeURIComponent(email)}`);
        if (!res.ok) return;
        const data = await res.json();
        _isAdmin = !!data.isAdmin;
        if (_isAdmin) document.getElementById("adminLink").style.display = "";
    } catch { /* AD unavailable or non-domain machine — admin link stays hidden */ }
}

async function fetchAdminPolicy() {
    try {
        const ctrl = new AbortController();
        const t = setTimeout(() => ctrl.abort(), 4000);
        const res = await fetch(`${BACKEND_BASE}/api/policy`, { signal: ctrl.signal });
        clearTimeout(t);
        if (!res.ok) return;
        const body = await res.json();
        _adminPolicy           = body.stamp      || {};
        _adminAttachmentPolicy = body.attachment || {};
        _adminPdf              = body.pdf        || { settings: {}, policy: {} };
        applyPrefsToCheckboxes();
        loadAttachments(); // re-filter with the now-known admin threshold
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

    // Signature-image threshold lock
    const forcedThresh = _adminAttachmentPolicy?.sigImgThreshold;
    const cbSig  = document.getElementById("spHideSigImages");
    const inpKb  = document.getElementById("spSigImgThresholdKb");
    const rowSig = document.getElementById("spSigThresholdRow");
    if (forcedThresh !== null && forcedThresh !== undefined) {
        cbSig.checked  = forcedThresh > 0;
        cbSig.disabled = true;
        inpKb.value    = forcedThresh > 0 ? Math.round(forcedThresh / 1024) : 8;
        inpKb.disabled = true;
        if (rowSig) rowSig.style.display = forcedThresh > 0 ? "flex" : "none";
        if (!cbSig.closest("label")?.querySelector(".policy-lock")) {
            const tag = document.createElement("span");
            tag.className = "policy-lock";
            tag.textContent = " 🔒";
            tag.title = "ההגדרה נקבעה על-ידי מנהל המערכת";
            tag.style.cssText = "color:#888;font-size:11px;margin-right:4px";
            cbSig.closest("label")?.appendChild(tag);
        }
    } else {
        cbSig.disabled = false;
        inpKb.disabled = false;
        cbSig.closest("label")?.querySelectorAll(".policy-lock").forEach(n => n.remove());
    }
}

// The in-body stamp is no longer a task-pane checkbox - it is configured once on
// the settings page (and may be locked by admin policy). Resolve it from prefs.
function stampEnabled() {
    const forced = policyForcedValue("defaultStamp");
    return forced !== null ? !!forced : !!_prefs.defaultStamp;
}

function applyPrefsToCheckboxes() {
    const catCb = document.getElementById("optCategory");
    catCb.checked = !!_prefs.defaultCategory;
    document.getElementById("categoryRow").style.display = _prefs.defaultCategory ? "" : "none";

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
    document.getElementById("spFieldFrom").checked        = !!_prefs.stampFields.from;
    document.getElementById("spFieldTo").checked          = !!_prefs.stampFields.to;
    document.getElementById("spFieldCc").checked          = !!_prefs.stampFields.cc;
    document.getElementById("spFieldSent").checked        = !!_prefs.stampFields.sent;
    document.getElementById("spFieldReceived").checked    = !!_prefs.stampFields.received;
    document.getElementById("spFieldSubject").checked     = !!_prefs.stampFields.subject;
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
    document.getElementById("spDispatchMode").checked        = !!_prefs.dispatchMode;
    document.getElementById("spForwardDefault").checked      = !!_prefs.forwardToLeaderByDefault;
    document.getElementById("spAddSelfAsEmployee").checked   = !!_prefs.addSelfAsEmployee;
    document.getElementById("spForwardDefaultRow").style.display =
        _prefs.dispatchMode ? "block" : "none";

    // Signature-image filter
    const thresh = _prefs.sigImgThreshold ?? 8192;
    document.getElementById("spHideSigImages").checked = thresh > 0;
    document.getElementById("spSigImgThresholdKb").value = thresh > 0 ? Math.round(thresh / 1024) : 8;
    document.getElementById("spSigThresholdRow").style.display = thresh > 0 ? "flex" : "none";

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

    // PDF settings
    const pdf = _prefs.pdf || DEFAULT_PREFS.pdf;
    document.getElementById("spPdfPageSize").value = pdf.pageSize || "A4";
    (document.querySelector(`input[name="spPdfOrientation"][value="${pdf.landscape ? "landscape" : "portrait"}"]`) || {}).checked = true;
    document.getElementById("spPdfMarginTop").value    = Number(pdf.marginTopCm    ?? 2.54).toFixed(2);
    document.getElementById("spPdfMarginBottom").value = Number(pdf.marginBottomCm ?? 2.54).toFixed(2);
    document.getElementById("spPdfMarginLeft").value   = Number(pdf.marginLeftCm   ?? 2.54).toFixed(2);
    document.getElementById("spPdfMarginRight").value  = Number(pdf.marginRightCm  ?? 2.54).toFixed(2);
    document.getElementById("spPdfPrintBackground").checked = pdf.printBackground !== false;
    // Default off, so test for an explicit true rather than "not false".
    document.getElementById("spMergePdfAttachments").checked = pdf.mergePdfAttachments === true;

    // Locks any fields the admin has forced (must run after the inputs are populated).
    applyPolicyLocks();
    applyPdfPolicyLocks();
}

// Disable PDF controls the admin has locked, showing the admin's forced value.
function applyPdfPolicyLocks() {
    const pol = _adminPdf?.policy   || {};
    const adm = _adminPdf?.settings || {};

    const ps = document.getElementById("spPdfPageSize");
    if (pol.pageSize) { if (adm.pageSize) ps.value = adm.pageSize; ps.disabled = true; }
    else ps.disabled = false;

    const oRadios = document.querySelectorAll('input[name="spPdfOrientation"]');
    if (pol.orientation) {
        const el = document.querySelector(`input[name="spPdfOrientation"][value="${adm.landscape ? "landscape" : "portrait"}"]`);
        if (el) el.checked = true;
        oRadios.forEach(r => r.disabled = true);
    } else oRadios.forEach(r => r.disabled = false);

    const mIds  = ["spPdfMarginTop", "spPdfMarginBottom", "spPdfMarginLeft", "spPdfMarginRight"];
    const mVals = [adm.marginTopCm, adm.marginBottomCm, adm.marginLeftCm, adm.marginRightCm];
    mIds.forEach((id, i) => {
        const el = document.getElementById(id);
        if (pol.margins) { if (mVals[i] != null) el.value = Number(mVals[i]).toFixed(2); el.disabled = true; }
        else el.disabled = false;
    });

    const pb = document.getElementById("spPdfPrintBackground");
    if (pol.printBackground) { pb.checked = adm.printBackground !== false; pb.disabled = true; }
    else pb.disabled = false;

    // Lives in the "סימון מצב" tab rather than the PDF tab, so it carries its own
    // locked note instead of the shared spPdfLockNote below.
    const mg = document.getElementById("spMergePdfAttachments");
    if (mg) {
        if (pol.mergePdfAttachments) { mg.checked = adm.mergePdfAttachments === true; mg.disabled = true; }
        else mg.disabled = false;
        const mgNote = document.getElementById("spMergePdfLockNote");
        if (mgNote) mgNote.style.display = pol.mergePdfAttachments ? "block" : "none";
    }

    const anyLocked = !!(pol.pageSize || pol.orientation || pol.margins || pol.printBackground);
    document.getElementById("spPdfLockNote").style.display = anyLocked ? "block" : "none";
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
        attachments: document.getElementById("spFieldAttachments").checked,
        from:        document.getElementById("spFieldFrom").checked,
        to:          document.getElementById("spFieldTo").checked,
        cc:          document.getElementById("spFieldCc").checked,
        sent:        document.getElementById("spFieldSent").checked,
        received:    document.getElementById("spFieldReceived").checked,
        subject:     document.getElementById("spFieldSubject").checked
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
    _prefs.addSelfAsEmployee        = document.getElementById("spAddSelfAsEmployee").checked;

    const hideSig  = document.getElementById("spHideSigImages").checked;
    const threshKb = parseInt(document.getElementById("spSigImgThresholdKb").value, 10) || 8;
    _prefs.sigImgThreshold = hideSig ? threshKb * 1024 : 0;

    _prefs.defaultSubfolders = {
        list:       _spSubfolderList.slice(),
        selected:   _spSubfolderSelected,
        appendDate: document.getElementById("spSubfolderAppendDate").checked
    };

    // PDF settings (admin-locked fields are disabled; reading them is harmless —
    // the server re-forces locked fields from the admin policy on save).
    const clampCm = v => Math.max(0.5, Math.min(10, parseFloat(v) || 2.54));
    _prefs.pdf = {
        pageSize:        document.getElementById("spPdfPageSize").value || "A4",
        landscape:       document.querySelector('input[name="spPdfOrientation"]:checked')?.value === "landscape",
        marginTopCm:     clampCm(document.getElementById("spPdfMarginTop").value),
        marginBottomCm:  clampCm(document.getElementById("spPdfMarginBottom").value),
        marginLeftCm:    clampCm(document.getElementById("spPdfMarginLeft").value),
        marginRightCm:   clampCm(document.getElementById("spPdfMarginRight").value),
        printBackground: document.getElementById("spPdfPrintBackground").checked,
        mergePdfAttachments: document.getElementById("spMergePdfAttachments").checked
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

async function addSubfolder() {
    const input = document.getElementById("spSubfolderInput");
    // These names become real folders later, so they follow the same rules.
    const chk = checkFolderName(input.value || "");
    if (!chk.ok) {
        if (chk.code !== "empty") await dlgAlert("שם לא תקין", chk.message);
        return;
    }
    const name = chk.value;
    input.value = name;
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
    const r = await dlgFolderName("שינוי שם תיקיה", _spSubfolderSelected, "שמור", "rename");
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
    // dlg() resolves with an object, so the old `if (!ok)` was never true and
    // Cancel removed the entry anyway. Every other call site uses r.ok.
    const r = await dlgConfirm("הסרה מהרשימה", `להסיר את "${_spSubfolderSelected}" מהרשימה?`);
    if (!r.ok) return;
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
    // Second line of defence behind the reset in loadFolderTree: only trust the
    // selection if the tree it came from belongs to the project being saved to.
    const currentProject = (document.getElementById("projectId")?.value || "").trim();
    if (_selectedFolderPath && _treeProjectNumber === currentProject)
        return _selectedFolderPath;
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

    if (_spSelectedCategory && !cats.some(c => c.displayName === _spSelectedCategory)) {
        cats.unshift({ displayName: _spSelectedCategory, color: null, _autoCreate: true });
    }

    if (cats.length === 0) {
        container.innerHTML = `<div class="cat-empty">אין קטגוריות מוגדרות</div>`;
        return;
    }

    const canManage = !!Office.context.mailbox.masterCategories?.removeAsync;

    cats.forEach(cat => {
        const row = document.createElement("div");
        row.className = "cat-row" + (cat.displayName === _spSelectedCategory ? " selected" : "");
        row.style.cssText = "display:flex;align-items:center;gap:4px;padding:5px 8px;cursor:pointer";

        const swatch = document.createElement("span");
        swatch.className = "cat-swatch";
        const palette = CATEGORY_PALETTE.find(p => p.name === cat.color);
        swatch.style.background = palette ? palette.hex : "#ccc";
        row.appendChild(swatch);

        const name = document.createElement("span");
        name.className = "cat-name";
        name.style.flex = "1";
        name.textContent = cat.displayName + (cat._autoCreate ? "  (יווצר אוטומטית)" : "");
        row.appendChild(name);

        if (cat.displayName === _spSelectedCategory) {
            const check = document.createElement("span");
            check.className = "cat-check";
            check.textContent = "✓";
            row.appendChild(check);
        }

        // Inline management buttons (only for real categories, not auto-create placeholders)
        if (canManage && !cat._autoCreate) {
            const btnRename = document.createElement("button");
            btnRename.textContent = "✎";
            btnRename.title = "שנה שם";
            btnRename.style.cssText = "padding:1px 5px;font-size:11px;border:1px solid #ccc;border-radius:3px;background:#fff;cursor:pointer;flex-shrink:0";
            row.appendChild(btnRename);

            const btnDel = document.createElement("button");
            btnDel.textContent = "🗑";
            btnDel.title = "מחק";
            btnDel.style.cssText = "padding:1px 5px;font-size:11px;border:1px solid #fca5a5;border-radius:3px;background:#fff;cursor:pointer;color:#b91c1c;flex-shrink:0";
            row.appendChild(btnDel);

            btnRename.onclick = (e) => { e.stopPropagation(); startInlineRename(cat); };
            btnDel.onclick    = (e) => { e.stopPropagation(); deleteMasterCategory(cat.displayName); };
        }

        row.onclick = () => { _spSelectedCategory = cat.displayName; renderCategoryList(); };
        container.appendChild(row);
    });
}

function deleteMasterCategory(name) {
    dlgConfirm("מחיקת קטגוריה", `למחוק את "${name}"?`, "מחק", "ביטול").then(ok => {
        if (!ok) return;
        Office.context.mailbox.masterCategories.removeAsync([name], r => {
            if (r.status !== Office.AsyncResultStatus.Succeeded) {
                dlgAlert("שגיאה במחיקה", r.error?.message || "מחיקה נכשלה"); return;
            }
            _spLoadedCategories = _spLoadedCategories.filter(c => c.displayName !== name);
            if (_spSelectedCategory === name)
                _spSelectedCategory = _spLoadedCategories[0]?.displayName || "";
            renderCategoryList();
        });
    });
}

function startInlineRename(cat) {
    // Populate the new-category form pre-filled for rename
    _editingCategoryOldName = cat.displayName;
    document.getElementById("spNewCatName").value = cat.displayName;
    _selectedSwatch = cat.color || "Preset7";
    document.getElementById("spNewCatForm").style.display = "block";
    document.getElementById("spNewCatCreate").textContent = "שמור שם";
    renderSwatches();
    document.getElementById("spNewCatName").focus();
    document.getElementById("spNewCatName").select();
}

let _editingCategoryOldName = null; // non-null = rename mode

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
    _editingCategoryOldName = null;
    document.getElementById("spNewCatForm").style.display = "block";
    document.getElementById("spNewCatCreate").textContent = "צור";
    document.getElementById("spNewCatName").value = "";
    _selectedSwatch = "Preset7";
    renderSwatches();
    document.getElementById("spNewCatName").focus();
}

function hideNewCategoryForm() {
    document.getElementById("spNewCatForm").style.display = "none";
    document.getElementById("spNewCatCreate").textContent = "צור";
    _editingCategoryOldName = null;
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
    if (!name) { dlgAlert("שגיאה", "יש להזין שם קטגוריה"); return; }

    if (!Office.context.mailbox.masterCategories?.addAsync) {
        dlgAlert("לא זמין", "יצירת קטגוריות דורשת Outlook 2019 ומעלה (Mailbox 1.8+)");
        return;
    }

    const isRename = !!_editingCategoryOldName && _editingCategoryOldName !== name;
    const oldName  = _editingCategoryOldName;
    _editingCategoryOldName = null;
    document.getElementById("spNewCatCreate").textContent = "צור";

    const doAdd = () => {
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
                if (isRename) _spLoadedCategories = _spLoadedCategories.filter(c => c.displayName !== oldName);
                _spLoadedCategories.push({ displayName: name, color: _selectedSwatch });
                _spSelectedCategory = name;
                hideNewCategoryForm();
                renderCategoryList();
            }
        );
    };

    if (isRename && oldName) {
        // Rename = delete old + add new
        Office.context.mailbox.masterCategories.removeAsync([oldName], () => doAdd());
    } else {
        doAdd();
    }
}

// =====================================================================
// Folder name rules (mirror of FolderNameRules.cs on the backend)
// =====================================================================
// Two different failure modes, deliberately handled differently:
//   * illegal characters  -> stripped as the user types, no nagging
//   * escape attempts     -> blocked, red warning, and reported to the server
//                            so the attempt is logged
// This is UX only. The backend enforces the same rules and is the real
// boundary; never rely on these checks for security.

const RESERVED_NAME_RE  = /^(CON|PRN|AUX|NUL|COM[0-9]|LPT[0-9])(\..*)?$/i;
// Narrow on purpose: a bare "%2e" can occur in a legitimate name, an encoded
// separator cannot.
const ENCODED_TRAVERSAL = /%2f|%5c|%2e%2e/i;
const DRIVE_RE          = /^[A-Za-z]:/;
const FOLDER_NAME_MAX   = 200;

// Character filter only - safe to run on every keystroke. It must NOT trim, or
// the user could never type a space between two words.
function stripIllegalFolderChars(raw) {
    if (!raw) return "";
    let s = "";
    for (const ch of raw) {
        const code = ch.codePointAt(0);
        if (code < 32 || code === 127) continue;
        if ("<>:|?*\"".indexOf(ch) >= 0) continue;
        s += ch;
    }
    return s;
}

// Strip what Windows cannot store, leaving the user's intent intact.
function sanitizeFolderName(raw) {
    if (!raw) return "";
    let s = "";
    for (const ch of raw) {
        const code = ch.codePointAt(0);
        if (code < 32 || code === 127) continue;      // control chars
        if ("<>:|?*\"".indexOf(ch) >= 0) continue;    // illegal in a Windows name
        s += ch;
    }
    s = s.replace(/\s+/g, " ").trim().replace(/[. ]+$/, "");
    if (s.length > FOLDER_NAME_MAX) s = s.slice(0, FOLDER_NAME_MAX).replace(/[. ]+$/, "");
    return s;
}

// Returns { ok, value, code, security, message }.
// message is Hebrew - it is shown directly in the taskpane.
//
// Judged on the RAW text: sanitizing first would turn "..\..\x" into "....x"
// and hide the attempt.
//
// Note what is NOT blocked: ".." inside a name. "as..as" is an ordinary folder
// name. Dots only navigate when they form a whole path segment, and separators
// are refused outright, so the only dangerous case left is a name made of
// nothing but dots.
function checkFolderName(raw) {
    raw = raw || "";
    const trimmed = raw.trim();

    // Order matters. Each check below runs on the RAW text and must run before
    // sanitizing, which would erase the very evidence being looked for - it
    // strips trailing dots, so "..'" would arrive here as an empty string and be
    // reported as "no name" instead of as a traversal attempt.

    // A name made only of dots IS a path segment: ".." resolves to the parent.
    if (trimmed.length > 0 && /^\.+$/.test(trimmed))
        return { ok: false, value: "", code: "traversal", security: true,
                 message: "שם תיקיה לא יכול להיות מורכב מנקודות בלבד. בחר/י שם אחר." };

    // Before the separator check: "C:\\Windows" is a drive escape, not a typo.
    if (DRIVE_RE.test(trimmed))
        return { ok: false, value: "", code: "traversal", security: true,
                 message: "שם תיקיה לא יכול להתחיל באות כונן (למשל C:). בחר/י שם אחר." };

    if (ENCODED_TRAVERSAL.test(raw))
        return { ok: false, value: "", code: "traversal", security: true,
                 message: "השם מכיל תו נתיב מקודד. בחר/י שם אחר." };

    const hasSeparator = raw.indexOf("/") >= 0 || raw.indexOf("\\") >= 0;

    // A separator together with ".." is the classic escape attempt.
    if (hasSeparator && raw.indexOf("..") >= 0)
        return { ok: false, value: "", code: "traversal", security: true,
                 message: "השם מכיל רצף של מעבר בין תיקיות. בחר/י שם אחר." };

    // A separator on its own is a usability problem, not an attack: the user is
    // trying to create a folder and a subfolder in one step.
    if (hasSeparator)
        return { ok: false, value: "", code: "separator", security: false,
                 message: "לא ניתן ליצור תיקיה ותת-תיקיה בפעולה אחת. צור/י קודם את התיקיה, ואז תת-תיקיה בתוכה." };

    const value = sanitizeFolderName(raw);

    if (!value)
        return { ok: false, value: "", code: "empty", security: false,
                 message: "יש להזין שם תיקיה." };

    if (RESERVED_NAME_RE.test(value))
        return { ok: false, value: "", code: "reserved", security: false,
                 message: `"${value}" הוא שם שמור במערכת ההפעלה ולא ניתן להשתמש בו. בחר/י שם אחר.` };

    return { ok: true, value, code: null, security: false, message: "" };
}

// Fire-and-forget audit report for a name the taskpane blocked before it was
// ever submitted. The server re-validates, so this cannot inject log text.
function reportBlockedFolderName(name, action) {
    try {
        if (!_treeProjectNumber) return;   // no project context (e.g. settings panel)
        fetch(`${BACKEND_BASE}/api/project/${encodeURIComponent(_treeProjectNumber)}/folders/report-blocked`,
            {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ name, action, user: getCurrentUserDisplay() })
            }).catch(() => {});
    } catch { /* never let auditing break the UI */ }
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
        const warn   = document.getElementById("modalWarn");
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

        // Optional live validation (opts.strip / opts.check), used by the folder
        // name prompts: illegal characters disappear as they are typed, while a
        // name that must be refused shows a red warning and disables OK.
        const showWarn = (msg) => {
            if (!warn) return;
            warn.textContent   = msg || "";
            warn.style.display = msg ? "block" : "none";
        };
        showWarn("");
        ok.disabled        = false;
        inp.style.borderColor = "#ccc";

        const runCheck = () => {
            if (!opts.check) return { ok: true, value: inp.value.trim() };

            if (opts.strip) {
                const raw = inp.value;
                const pos = inp.selectionStart == null ? raw.length : inp.selectionStart;
                const cleaned = opts.strip(raw);
                if (cleaned !== raw) {
                    const removedBefore = pos - opts.strip(raw.slice(0, pos)).length;
                    inp.value = cleaned;
                    const np = Math.max(0, pos - removedBefore);
                    try { inp.setSelectionRange(np, np); } catch { /* older webview */ }
                }
            }

            const r = opts.check(inp.value);
            showWarn(r.ok ? "" : r.message);
            inp.style.borderColor = r.ok ? "#ccc" : "#a4262c";
            ok.disabled = !r.ok;
            return r;
        };
        // 'input' alone is not enough: in the Outlook task-pane webview a paste,
        // cut or drag-drop can land in the field without firing it, which let
        // illegal characters in through Ctrl+V even though typing was filtered.
        // Listening to those events too - deferred, so the field already holds
        // the pasted text when the check runs - closes that hole.
        const LIVE_EVENTS = ["input", "paste", "cut", "drop", "keyup", "change"];
        const liveHandler = () => setTimeout(runCheck, 0);
        if (opts.check) LIVE_EVENTS.forEach(ev => inp.addEventListener(ev, liveHandler));

        bd.style.display = "flex";
        if (opts.input) setTimeout(() => { inp.focus(); inp.select(); }, 50);

        const close = (result) => {
            bd.style.display = "none";
            ok.disabled = false;
            inp.style.borderColor = "#ccc";
            showWarn("");
            if (opts.check) LIVE_EVENTS.forEach(ev => inp.removeEventListener(ev, liveHandler));
            ok.onclick = cancel.onclick = inp.onkeydown = null;
            resolve(result);
        };
        ok.onclick = () => {
            if (!opts.input) return close({ ok: true });
            if (!opts.check) return close({ ok: true, value: inp.value.trim() });

            // Re-check on submit: the value may have been pasted, and a blocked
            // name must be reported for the audit log before we refuse it.
            const r = runCheck();
            if (!r.ok) {
                if (r.security && opts.onBlocked) opts.onBlocked(inp.value, r);
                return;                       // keep the dialog open
            }
            close({ ok: true, value: r.value });
        };
        cancel.onclick = () => close({ ok: false });
        inp.onkeydown  = (e) => {
            if (e.key === "Enter")  { e.preventDefault(); ok.click(); }
            if (e.key === "Escape") { e.preventDefault(); cancel.click(); }
        };
    });
}
const dlgPrompt  = (title, defaultValue, okLabel) => dlg({ title, input: true, defaultValue, okLabel });
// Prompt for a folder name: strips illegal characters live, blocks traversal and
// Windows-reserved names with a red warning, and reports blocked attempts.
const dlgFolderName = (title, defaultValue, okLabel, action) => dlg({
    title, input: true, defaultValue, okLabel,
    strip: stripIllegalFolderChars,
    check: checkFolderName,
    onBlocked: (raw) => reportBlockedFolderName(raw, action)
});
const dlgConfirm = (title, message, okLabel)      => dlg({ title, message, okLabel });
const dlgAlert   = (title, message)               => dlg({ title, message, hideCancel: true, okLabel: "סגור" });

// Show a folder path in a selectable text input with a Copy button.
// Falls back to document.execCommand('copy') which works without clipboard permissions.
function dlgPath(path) {
    return new Promise(resolve => {
        const bd     = document.getElementById("modalBackdrop");
        const titleE = document.getElementById("modalTitle");
        const msgE   = document.getElementById("modalMsg");
        const inp    = document.getElementById("modalInput");
        const ok     = document.getElementById("modalOk");
        const cancel = document.getElementById("modalCancel");

        titleE.textContent    = "נתיב התיקייה";
        msgE.style.display    = "none";
        inp.style.display     = "block";
        inp.value             = path;
        inp.readOnly          = true;
        inp.style.direction   = "ltr";
        inp.style.textAlign   = "left";

        ok.textContent        = "📋 העתק";
        cancel.textContent    = "סגור";
        cancel.style.display  = "inline-block";

        bd.style.display = "flex";
        setTimeout(() => { inp.focus(); inp.select(); }, 50);

        const close = () => {
            inp.readOnly       = false;
            inp.style.direction  = "";
            inp.style.textAlign  = "";
            bd.style.display   = "none";
            ok.onclick = cancel.onclick = inp.onkeydown = null;
            resolve();
        };

        ok.onclick = () => {
            inp.select();
            try { document.execCommand("copy"); } catch {}
            const orig = ok.textContent;
            ok.textContent = "✓ הועתק!";
            setTimeout(() => { ok.textContent = orig; }, 1400);
        };
        cancel.onclick = close;
        inp.onkeydown  = (e) => { if (e.key === "Escape") { e.preventDefault(); close(); } };
    });
}

// Convert a Windows path to a file:// URL the browser can follow.
//   \\FS01\jobs\24  -> file://FS01/jobs/24
//   C:\Apps\X       -> file:///C:/Apps/X
function fileUrlFromPath(p) {
    if (!p) return "";
    const s = String(p);
    if (s.startsWith("\\\\")) return "file://" + s.replace(/^\\+/, "").replace(/\\/g, "/");
    return "file:///" + s.replace(/\\/g, "/");
}

// Popup that shows the folder as a clickable link the user opens themselves.
// The backend runs as a service and can't open Explorer in the user's session,
// and window.open("file://…") is blocked in WebView2 — so we hand the user a
// real <a href="file://…"> link (plus a Copy-path fallback if the link is
// blocked by their Outlook/WebView2 policy).
function dlgFolderLink(path) {
    return new Promise(resolve => {
        const bd     = document.getElementById("modalBackdrop");
        const titleE = document.getElementById("modalTitle");
        const msgE   = document.getElementById("modalMsg");
        const inp    = document.getElementById("modalInput");
        const ok     = document.getElementById("modalOk");
        const cancel = document.getElementById("modalCancel");

        const url = fileUrlFromPath(path);
        titleE.textContent = "פתיחת תיקייה";
        msgE.style.display = "block";
        msgE.innerHTML =
            "לחצו על הקישור לפתיחת התיקייה בסייר הקבצים:" +
            '<a href="' + escHtml(url) + '" target="_blank" rel="noopener" ' +
            'style="display:block;margin-top:8px;color:#0078d4;direction:ltr;text-align:left;word-break:break-all">' +
            "📂 " + escHtml(path) + "</a>" +
            '<div style="margin-top:8px;font-size:11px;color:#888">' +
            "אם הקישור אינו נפתח, השתמשו ב«העתק נתיב» והדביקו בסייר הקבצים.</div>";

        inp.style.display = "none";

        ok.textContent       = "📋 העתק נתיב";
        cancel.textContent   = "סגור";
        cancel.style.display  = "inline-block";

        bd.style.display = "flex";

        const close = () => {
            bd.style.display = "none";
            msgE.innerHTML   = "";
            ok.onclick = cancel.onclick = null;
            resolve();
        };

        ok.onclick = () => {
            const ta = document.createElement("textarea");
            ta.value = path;
            ta.style.cssText = "position:fixed;opacity:0;top:0;left:0";
            document.body.appendChild(ta);
            ta.select();
            try { document.execCommand("copy"); } catch {}
            document.body.removeChild(ta);
            const orig = ok.textContent;
            ok.textContent = "✓ הועתק!";
            setTimeout(() => { ok.textContent = orig; }, 1400);
        };
        cancel.onclick = close;
    });
}

// =====================================================================
// Project folder tree (in the third tab)
// =====================================================================
let _folderTree         = null;   // root node from backend
let _folderRootPath     = "";     // absolute server path of the project root
let _selectedFolderPath = "";     // empty = project root
let _expandedPaths      = new Set([""]);     // root always expanded
let _treeProjectNumber  = "";     // project the tree was loaded for

async function loadFolderTree(projectNumber) {
    // Reset here, not only on the success path. _treeProjectNumber was reassigned
    // immediately while _selectedFolderPath was cleared 20 lines later, so a tree
    // load that failed (share hiccup, project does not exist) left a folder chosen
    // in the PREVIOUS project as the save destination for the new one - and the
    // backend would happily create that subfolder there and file the e-mail in it.
    _treeProjectNumber  = projectNumber;
    _selectedFolderPath = "";
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
        _folderTree     = body.tree;
        _folderRootPath = body.rootPath || "";
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

    // Disable rename / delete on the project root. Delete additionally requires an
    // admin session on the server, so do not offer it to users who cannot use it.
    menu.querySelector('[data-action="rename"]').style.display = node.path ? "block" : "none";
    menu.querySelector('[data-action="delete"]').style.display =
        (node.path && _isAdmin) ? "block" : "none";

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
            if (action === "open")    doOpenFolder(node);
            else if (action === "create") doCreateFolder(node);
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

// ---------- Open folder for the user ----------
// The backend runs as a Windows service and cannot open Explorer in the user's
// interactive session, so we present a clickable folder link (with a copy-path
// fallback) the user opens themselves.
function openFolderForUser(fullPath) {
    return dlgFolderLink(fullPath);
}

async function doOpenFolder(node) {
    if (!_folderRootPath) return;
    const rel  = node.path ? node.path.replace(/\//g, "\\") : "";
    const full = rel ? _folderRootPath.replace(/[\\\/]+$/, "") + "\\" + rel : _folderRootPath;
    await openFolderForUser(full);
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
    const r = await dlgFolderName("שם תיקיה חדשה", suggestNewFolderName(), "צור", "create");
    if (!r.ok || !r.value) return;

    try {
        const res = await fetch(
            `${BACKEND_BASE}/api/project/${encodeURIComponent(_treeProjectNumber)}/folders`,
            {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ parent: parentNode.path, name: r.value, user: getCurrentUserDisplay() })
            }
        );
        if (!res.ok) {
            const err = await res.json().catch(() => ({}));
            await dlgAlert("שגיאה", err.error || ("שגיאה ביצירת תיקיה: " + res.status));
            return;
        }
        _expandedPaths.add(parentNode.path);
        // Compute new path before tree reload (loadFolderTree resets _selectedFolderPath)
        const newFolderPath = parentNode.path ? `${parentNode.path}/${r.value}` : r.value;
        await loadFolderTree(_treeProjectNumber);
        // Auto-select the newly created folder as the save destination
        _selectedFolderPath = newFolderPath;
        renderFolderTree();
        updateDestBanner();
    } catch (e) {
        await dlgAlert("שגיאה", e.message);
    }
}

async function doRenameFolder(node) {
    const r = await dlgFolderName("שם חדש", node.name, "שנה שם", "rename");
    if (!r.ok || !r.value || r.value === node.name) return;

    try {
        const res = await fetch(
            `${BACKEND_BASE}/api/project/${encodeURIComponent(_treeProjectNumber)}/folders`,
            {
                method: "PUT",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ path: node.path, newName: r.value, user: getCurrentUserDisplay() })
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
                body: JSON.stringify({ path: node.path, recursive: true, user: getCurrentUserDisplay() })
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
// A save takes seconds (body fetch, attachment encoding, Chromium render) and the
// button used to stay live the whole time. A second click duplicated the
// attachments on disk - AttachmentService deliberately never overwrites, so you
// got scan.pdf, scan(1).pdf, scan(2).pdf - opened a second forward window, and
// wrote two log rows racing on the same metadata file.
let _saveInFlight = false;

async function onSaveAsPdf() {
    const status = document.getElementById("status");
    if (_saveInFlight) return;
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

    const saveBtn = document.getElementById("saveBtn");
    _saveInFlight = true;
    if (saveBtn) saveBtn.disabled = true;

    try {
        status.innerText = "אוסף נתוני הודעה…";
        const mailData = await collectMailData();

        const doStamp    = stampEnabled();
        const doCategory = document.getElementById("optCategory").checked;
        const fwdCbEl    = document.getElementById("optForwardToLeader");
        const willForward = !!(_prefs.dispatchMode && fwdCbEl?.checked);

        status.innerText = "שולח נתונים לשרת…";
        const result = await sendToBackend({
            projectId, projectName, projectLeader,
            savedBy:   getCurrentUserDisplay(),
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
                includeFrom:        !!_prefs.stampFields?.from,
                includeTo:          !!_prefs.stampFields?.to,
                includeCc:          !!_prefs.stampFields?.cc,
                includeSent:        !!_prefs.stampFields?.sent,
                includeReceived:    !!_prefs.stampFields?.received,
                includeSubject:     !!_prefs.stampFields?.subject,
                notes:              _prefs.stampNotes || "",
                template:           _prefs.stampTemplate || "",
                forwarded:          willForward,
                forwardedTo:        willForward ? (leaderDisplayName(projectLeader) || stripEmail(projectLeader)) : "",
                userName:           getCurrentUserDisplay(),
                attachmentNames:    mailData.allAttachmentNames || []
            } : null,
            // User-chosen PDF output settings; the server merges these with the
            // admin lock policy (locked fields are forced server-side).
            pdfSettings: _prefs.pdf || DEFAULT_PREFS.pdf
        });

        // Surface PDF generation outcome (server may have fallen back to .html)
        if (result?.pdf && result.pdf.pdfCreated === false) {
            const reason = result.pdf.fallbackReason || "המרה ל‑PDF נכשלה";
            await dlgAlert(
                "לא נוצר קובץ PDF",
                `נשמר קובץ HTML חלופי במקום ה‑PDF.\n\nסיבה: ${reason}\n\nנתיב: ${result.pdf.fullPath || ""}`
            );
        }

        const savedAttachmentNames = (mailData.attachments || []).map(a => a.name);

        const warnings = await markMessageProcessed(
            Office.context.mailbox.item,
            projectId, projectName, projectLeader,
            result, doStamp, doCategory, savedAttachmentNames
        );

        // Attachments that could not be saved are a user-visible outcome, not a
        // console detail: the PDF still lists every attachment the message had.
        if (mailData.attachmentWarnings?.length)
            warnings.push(...mailData.attachmentWarnings);

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
        const saveDir   = result?.project?.saveDir || result?.project?.fullName || "";

        // Show the "already saved" info card straight away (mirrors what the user
        // will see next time this message is opened).
        renderSavedInfo({
            processed:   "true",
            projectId,
            projectName,
            dateIso:     new Date().toISOString(),
            folder:      saveDir,
            file:        savedName,
            attachments: savedAttachmentNames
        });

        if (warnings.length > 0) {
            await dlgAlert("ההודעה נשמרה — חלק מהפעולות לא בוצעו",
                warnings.map(w => "• " + w).join("\n"));
        }

        // Build status line: text + open-folder button
        status.innerHTML = "";
        status.appendChild(document.createTextNode(
            savedName
                ? `נשמר ✅ ${savedName}${warnings.length > 0 ? " (עם אזהרות)" : ""}`
                : `ההודעה נשמרה בהצלחה ✅${warnings.length > 0 ? " (עם אזהרות)" : ""}`
        ));

        if (saveDir) {
            const openBtn = document.createElement("button");
            openBtn.type        = "button";
            openBtn.title       = saveDir;
            openBtn.textContent = "📂";
            openBtn.style.cssText =
                "margin-right:8px;margin-top:0;padding:2px 7px;font-size:13px;" +
                "border-radius:4px;cursor:pointer;vertical-align:middle";
            openBtn.onclick = () => openFolderForUser(saveDir);
            status.appendChild(openBtn);
        }
    } catch (err) {
        console.error(err);
        if (isNetworkError(err)) {
            status.innerText = "השרת ישן 😴";
            showSleepingScreen();
            return;
        }
        status.innerText = "שגיאה בשמירת ההודעה ❌";
        await dlgAlert("שגיאה בשמירה", err?.message || String(err));
    } finally {
        _saveInFlight = false;
        // Re-enable only if the backend is still reachable; the health check owns
        // the disabled state in that case.
        const btn = document.getElementById("saveBtn");
        if (btn && _backendUp !== false) btn.disabled = false;
    }
}

// =====================================================================
// Validation / Collect / Send
// =====================================================================
function validateMarkingOptions() {
    return stampEnabled() ||
           document.getElementById("optCategory").checked;
}

// ---------------------------------------------------------------------------
// Attachment payloads
// ---------------------------------------------------------------------------
// getAttachmentContentAsync does not always hand back base64. An e-mail attached
// to an e-mail comes back as `eml` (a raw MIME string), a meeting item as
// `iCalendar`, and a OneDrive/SharePoint link attachment as `url`. This code used
// to skip every non-base64 format with a bare `continue`, so embedded messages
// vanished from the project folder without a word - while the PDF cover page went
// on listing them by name, because that list is read straight off item.attachments.
// Text formats are real content: encode them and save them like any other file,
// and never drop an attachment silently.
const TEXT_ATTACHMENT_FORMATS = {
    eml:       { ext: ".eml", contentType: "message/rfc822" },
    icalendar: { ext: ".ics", contentType: "text/calendar"  }
};

function ensureExtension(name, ext) {
    const n = (name || "attachment").trim();
    return n.toLowerCase().endsWith(ext) ? n : n + ext;
}

// UTF-8 safe string -> base64. btoa() on its own throws on any code point above
// U+00FF, which is every Hebrew subject line inside an .eml.
//
// TextEncoder is missing from the IE11 webview that classic Outlook still uses on
// some desktops, and core-js does not polyfill it, so fall back to the
// encodeURIComponent trick there rather than throwing away the attachment on the
// one host where this matters most.
function utf8ToBase64(text) {
    if (typeof TextEncoder !== "undefined") {
        const bytes = new TextEncoder().encode(text);
        let binary = "";
        const CHUNK = 0x8000;   // fromCharCode blows its argument limit on large buffers
        for (let i = 0; i < bytes.length; i += CHUNK)
            binary += String.fromCharCode.apply(null, bytes.subarray(i, i + CHUNK));
        return btoa(binary);
    }
    return btoa(encodeURIComponent(text).replace(/%([0-9A-F]{2})/g,
        (_, hex) => String.fromCharCode(parseInt(hex, 16))));
}

// Returns { payload } for something that can be written to disk, or { reason }
// explaining why it cannot be - the caller turns that into a visible warning.
function buildAttachmentPayload(att, content) {
    const format = String(content?.format || "").toLowerCase();
    const raw    = content?.content;

    if (!raw) return { reason: "הקובץ המצורף חזר ריק מ-Outlook" };

    if (format === "base64") {
        return { payload: {
            name:        att.name,
            contentType: att.contentType,
            size:        att.size,
            base64:      raw.replace(/[\r\n\t ]/g, "")
        } };
    }

    const asText = TEXT_ATTACHMENT_FORMATS[format];
    if (asText) {
        return { payload: {
            name:        ensureExtension(att.name, asText.ext),
            contentType: asText.contentType,
            size:        raw.length,
            base64:      utf8ToBase64(raw)
        } };
    }

    if (format === "url") {
        return { reason: "קובץ ענן (OneDrive/SharePoint) — Outlook מחזיר קישור בלבד, אין תוכן לשמירה" };
    }

    return { reason: `פורמט לא נתמך (${content.format})` };
}

async function collectMailData() {
    const item = Office.context.mailbox.item;

    let bodyHtml = await new Promise((resolve, reject) =>
        item.body.getAsync(Office.CoercionType.Html,
            r => r.status === Office.AsyncResultStatus.Succeeded ? resolve(r.value) : reject(r.error))
    );

    // Resolve cid: inline image references to data: URIs - Chromium is never
    // allowed to fetch anything, so a cid: left in the HTML prints as a broken
    // image with only its alt text ("Inline image").
    //
    // The Content-ID is NOT reliably the file name. Mail composed in Outlook uses
    // "image001.png@01DC1234.56789ABC", so the part before the @ happened to equal
    // the attachment name and a name-keyed lookup worked. Mail that passes through
    // Yahoo / Gmail / Apple Mail carries an opaque Content-ID
    // ("1349457626.1555239@mail.yahoo.com") while Exchange still names the part
    // image0NN.png - the lookup missed, the old code wrote src="" and the
    // signature logo printed as a broken link. So: try every sensible key, then
    // fall back to document order (Outlook lists inline attachments in the order
    // the body references them), and if even that fails leave the tag untouched
    // and warn - blanking the src is what drew the broken-link box.
    const inlineImageWarnings = [];
    if (bodyHtml.includes('cid:')) {
        const inlineAtts   = (item.attachments || []).filter(a => a.isInline);
        const inlineImages = [];          // document order, as Outlook reports them
        const byKey        = new Map();

        const addKey = (key, image) => {
            const k = String(key || '').trim().toLowerCase();
            if (k && !byKey.has(k)) byKey.set(k, image);
        };

        for (const att of inlineAtts) {
            try {
                const result = await new Promise((res, rej) =>
                    item.getAttachmentContentAsync(att.id,
                        r => r.status === Office.AsyncResultStatus.Succeeded ? res(r.value) : rej(r.error))
                );
                // Compare case-insensitively rather than against the enum object:
                // the value is "base64", but do not rely on the host echoing that
                // exact casing.
                if (String(result?.format || '').toLowerCase() !== 'base64' || !result.content) {
                    inlineImageWarnings.push(
                        `${att.name} — תמונה מוטבעת שלא ניתן להטמיע (${result?.format || 'empty'})`);
                    continue;
                }
                // Strip MIME line-breaks so the data URI stays valid
                const b64   = result.content.replace(/[\r\n\t ]/g, '');
                const name  = att.name || '';
                const image = {
                    name,
                    dataUri: `data:${att.contentType || 'image/png'};base64,${b64}`,
                    used: false
                };
                inlineImages.push(image);
                addKey(name, image);                            // image010.png
                addKey(name.replace(/\.[^.]+$/, ''), image);    // image010
            } catch {
                inlineImageWarnings.push(
                    `${att.name} — לא ניתן לקרוא תמונה מוטבעת מ-Outlook`);
            }
        }

        if (inlineImages.length > 0) {
            const resolveCid = (rawCid) => {
                const cid = String(rawCid || '').replace(/^<|>$/g, '').trim().toLowerCase();
                if (!cid) return null;

                // 1. the whole Content-ID  2. the part before the @ (classic Outlook)
                let hit = byKey.get(cid) || byKey.get(cid.split('@')[0]);

                // 3. an attachment name embedded anywhere in the Content-ID
                if (!hit) hit = inlineImages.find(
                    img => img.name && cid.includes(img.name.toLowerCase()));

                // 4. document order - the Content-ID is opaque to us
                if (!hit) hit = inlineImages.find(img => !img.used);

                if (!hit) return null;
                hit.used = true;
                return hit.dataUri;
            };

            let unresolved = 0;
            // Quoted and unquoted src, and the FULL Content-ID - never split on @.
            bodyHtml = bodyHtml.replace(
                /src\s*=\s*(?:(["'])cid:([^"']*)\1|cid:([^\s">]+))/gi,
                (match, quote, quoted, bare) => {
                    const uri = resolveCid(quoted !== undefined ? quoted : bare);
                    if (uri) return `src="${uri}"`;
                    unresolved++;
                    return match;   // leave as-is; src="" is what drew the broken box
                }
            );
            if (unresolved > 0)
                inlineImageWarnings.push(
                    `${unresolved} תמונות מוטבעות בגוף ההודעה לא נפתרו ולא יופיעו ב-PDF`);
        } else if (inlineAtts.length > 0) {
            inlineImageWarnings.push(
                'לא ניתן היה לקרוא את התמונות המוטבעות בגוף ההודעה');
        }
    }

    const checkedIds = new Set(
        [...document.querySelectorAll(".att-check:checked")].map(cb => cb.dataset.id)
    );

    const attachments = [];
    const attachmentWarnings = [...inlineImageWarnings];
    for (const att of item.attachments) {
        if (!checkedIds.has(att.id)) continue;

        let content;
        try {
            content = await new Promise((resolve, reject) =>
                item.getAttachmentContentAsync(att.id,
                    r => r.status === Office.AsyncResultStatus.Succeeded ? resolve(r.value) : reject(r.error))
            );
        } catch (err) {
            attachmentWarnings.push(`${att.name} — לא ניתן לקרוא את הקובץ מ-Outlook (${err?.message || err})`);
            continue;
        }

        const built = buildAttachmentPayload(att, content);
        if (built.payload) attachments.push(built.payload);
        else               attachmentWarnings.push(`${att.name} — ${built.reason}`);
    }

    // All attachment names from the original message (independent of save selection),
    // used only for stamping the PDF.
    const allAttachmentNames = (item.attachments || [])
        .filter(a => !a.isInline)
        .map(a => a.name);

    // "Name <email>" when a display name is available, otherwise the bare address.
    const fmtAddr = (a) => {
        if (!a) return "";
        const email = a.emailAddress || "";
        const name  = (a.displayName || "").trim();
        return name && name !== email ? `${name} <${email}>` : email;
    };
    const fmtList = (arr) => (arr || []).map(fmtAddr).filter(Boolean);

    // Read-mode Outlook exposes the sent time as dateTimeCreated.
    const sentDate = item.dateTimeCreated
        ? new Date(item.dateTimeCreated).toISOString()
        : null;

    // receivedDate was never sent, so EmailDto.ReceivedDate was always null. Two
    // consequences: the PDF file name used the SAVE time instead of the message
    // time (two messages saved in the same minute collided), and the "התקבל" stamp
    // row was a permanent no-op because PdfService only emits it when the value is
    // present. dateTimeCreated is the received time for an incoming message; fall
    // back to the sent time when Outlook does not expose it.
    const receivedDate = item.dateTimeModified
        ? new Date(item.dateTimeModified).toISOString()
        : sentDate;

    return {
        email: {
            subject:  item.subject,
            from:     fmtAddr(item.from) || item.from?.emailAddress || "",
            to:       fmtList(item.to),
            cc:       fmtList(item.cc),
            sentDate,
            receivedDate,
            bodyHtml
        },
        attachments,
        allAttachmentNames,
        attachmentWarnings
    };
}

function sendToBackend(payload) {
    return new Promise((resolve, reject) => {
        const progressWrap  = document.getElementById("uploadProgress");
        const progressBar   = document.getElementById("uploadProgressBar");
        const progressLabel = document.getElementById("uploadProgressLabel");

        const showProgress = (pct) => {
            if (progressWrap)  progressWrap.style.display  = "block";
            if (progressBar)   progressBar.style.width     = pct + "%";
            if (progressLabel) progressLabel.textContent   = pct + "%";
        };
        const hideProgress = () => {
            if (progressWrap) progressWrap.style.display = "none";
        };

        showProgress(0);

        const xhr = new XMLHttpRequest();

        xhr.upload.onprogress = (e) => {
            if (e.lengthComputable) {
                showProgress(Math.round(e.loaded / e.total * 100));
            }
        };

        xhr.onload = () => {
            showProgress(100);
            hideProgress();
            if (xhr.status >= 200 && xhr.status < 300) {
                try { resolve(JSON.parse(xhr.responseText)); }
                catch { resolve({}); }
            } else {
                let detail = xhr.responseText;
                try { detail = JSON.parse(xhr.responseText).error ||
                               JSON.parse(xhr.responseText).title || detail; } catch {}
                reject(new Error(`HTTP ${xhr.status}: ${detail || xhr.statusText}`));
            }
        };

        xhr.onerror = () => {
            hideProgress();
            reject(new Error("שגיאת רשת — לא ניתן להגיע לשרת"));
        };

        xhr.ontimeout = () => {
            hideProgress();
            reject(new Error("הבקשה לשרת פגה בזמן"));
        };

        xhr.open("POST", `${BACKEND_BASE}/api/saveaspdf`);
        // ontimeout was dead code: XMLHttpRequest defaults to no timeout, so a hung
        // backend left the pane stuck on "שולח נתונים לשרת…" with no way out. The
        // server's own render budget tops out at 120 s, so 180 s is a real ceiling
        // rather than an impatient one.
        xhr.timeout = 180000;
        xhr.setRequestHeader("Content-Type", "application/json");
        xhr.send(JSON.stringify(payload));
    });
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
async function markMessageProcessed(item, projectId, projectName, projectLeader, result, doStamp, doCategory, savedAttachmentNames = []) {
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
                    props.set("saveAsPdfProjectLeaderName", leaderDisplayName(projectLeader));
                    props.set("saveAsPdfFolder",        result?.project?.saveDir || result?.project?.fullName || "");
                    props.set("saveAsPdfFile",          result?.pdf?.fileName || "");
                    props.set("saveAsPdfAttachments",   JSON.stringify(savedAttachmentNames || []));
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
            if (f.leader      && projectLeader) parts.push(`מנהל: ${leaderDisplayName(projectLeader) || stripEmail(projectLeader)}`);
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

// Resolve a human display name ("<FirstName> <LastName>") for a leader address.
// Looks in the project roster first - that is what .SaveAsPDF_Emploeeys.xml
// gives us - then in the global contacts cache. Returns "" when neither knows
// the address (callers fall back to showing the address itself).
function leaderDisplayName(addressOrText) {
    const addr = stripEmail(addressOrText).toLowerCase();
    if (!addr) return "";
    const emp = _employees.find(e => (e.email || "").toLowerCase() === addr);
    if (emp?.displayName) return emp.displayName.trim();
    const con = _allContacts.find(c => (c.email || "").toLowerCase() === addr);
    return (con?.displayName || "").trim();
}

// A saved message only stores the leader's ADDRESS, and on restore neither the
// roster nor the contacts cache is necessarily loaded yet. Fetch the project's
// own employees XML and show the name it holds.
async function enrichRestoredLeader(projectNumber, email) {
    const input = document.getElementById("projectLeader");
    const addr  = (email || "").toLowerCase();
    if (!addr) return;

    const local = leaderDisplayName(email);
    // dataset.email guard: the user may have switched messages mid-flight.
    if (local) { if (input.dataset.email === email) input.value = local; return; }
    if (!projectNumber) return;

    try {
        const res = await fetch(`${BACKEND_BASE}/api/project/${encodeURIComponent(projectNumber)}`);
        if (!res.ok) return;
        const data = await res.json();
        if (!data.exists || !Array.isArray(data.employees)) return;
        const hit = data.employees.find(e => (e.email || "").toLowerCase() === addr)
                 || data.employees.find(e => e.isLeader);
        const name = (hit?.displayName || "").trim();
        if (name && input.dataset.email === email) input.value = name;
    } catch { /* offline / backend down - keep showing the address */ }
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
    // Show the person, not the address.
    const leaderLabel = ctx.projectLeader
        ? (leaderDisplayName(ctx.projectLeader) || stripEmail(ctx.projectLeader))
        : "";
    // Advanced template overrides field toggles when non-empty
    if (_prefs.stampTemplate?.trim()) {
        return _prefs.stampTemplate
            .replace(/\{\{projectId\}\}/g,   ctx.projectId   || "")
            .replace(/\{\{projectName\}\}/g, ctx.projectName || "")
            .replace(/\{\{leader\}\}/g,     leaderLabel || "")
            .replace(/\{\{date\}\}/g,       ctx.date         || "")
            .replace(/\{\{notes\}\}/g,      _prefs.stampNotes || "");
    }

    const f = _prefs.stampFields || DEFAULT_PREFS.stampFields;
    const lines = [];
    if (f.projectId   && ctx.projectId)     lines.push(`מספר פרויקט: ${ctx.projectId}`);
    if (f.projectName && ctx.projectName)   lines.push(`שם פרויקט: ${ctx.projectName}`);
    if (f.leader      && leaderLabel)       lines.push(`מנהל/ת פרויקט: ${leaderLabel}`);
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
        const props    = r.value;
        const id       = props.get("saveAsPdfProjectId");
        const name     = props.get("saveAsPdfProjectName");
        const leader   = props.get("saveAsPdfProjectLeader");
        const folder   = props.get("saveAsPdfFolder");
        const file     = props.get("saveAsPdfFile");
        const dateIso  = props.get("saveAsPdfDate");
        const processed = props.get("saveAsPdfProcessed");
        let   atts     = [];
        try { atts = JSON.parse(props.get("saveAsPdfAttachments") || "[]"); } catch { atts = []; }

        if (id)   document.getElementById("projectId").value   = id;
        if (name) document.getElementById("projectName").value = name;
        if (leader) {
            const input = document.getElementById("projectLeader");
            const email = stripEmail(leader);
            // Prefer the name persisted with the message; otherwise show the
            // address for now and swap in the name from the project XML.
            const savedName = (props.get("saveAsPdfProjectLeaderName") || "").trim();
            input.dataset.email = email;
            input.value = savedName || leaderDisplayName(email) || email;
            if (!savedName) enrichRestoredLeader(id, email);
        }

        renderSavedInfo({
            processed, projectId: id, projectName: name,
            dateIso, folder, file, attachments: atts
        });
    });
}

// Format an ISO timestamp as "dd.MM.yyyy HH:mm" (Israeli display order).
function formatSavedDate(iso) {
    if (!iso) return "";
    const d = new Date(iso);
    if (isNaN(d.getTime())) return iso;
    return `${pad2(d.getDate())}.${pad2(d.getMonth() + 1)}.${d.getFullYear()} ` +
           `${pad2(d.getHours())}:${pad2(d.getMinutes())}`;
}

// Render (or hide) the "this message was already saved" info card. Mirrors the
// summary the user previously saw stamped into saved messages: project number,
// name, update date, and clickable links that open the destination folder —
// same behavior as the attachment links.
function renderSavedInfo(info) {
    const card = document.getElementById("savedInfo");
    const body = document.getElementById("savedInfoBody");
    if (!card || !body) return;

    const isSaved = info && (info.processed === "true" || info.folder || info.projectId);
    if (!isSaved) { card.style.display = "none"; body.innerHTML = ""; return; }

    const dateStr = formatSavedDate(info.dateIso);
    const folder  = info.folder || "";
    const fAttr   = escHtml(folder);

    let html = "";
    if (info.projectId)   html += `<div class="si-row"><span class="si-label">מס' פרויקט:</span> <b>${escHtml(info.projectId)}</b></div>`;
    if (info.projectName) html += `<div class="si-row"><span class="si-label">שם הפרויקט:</span> ${escHtml(info.projectName)}</div>`;
    if (dateStr)          html += `<div class="si-row"><span class="si-label">תאריך עדכון:</span> ${escHtml(dateStr)}</div>`;

    if (folder) {
        html += `<div class="si-sep"></div>`;
        // One clear "open folder" action, plus the location and file names as
        // plain info text (clicking opens a popup with a folder link — the
        // service backend can't open Explorer in the user's session).
        html += `<div class="si-row">🗂️ <span class="si-label">מיקום השמירה:</span> ` +
                `<a href="#" class="si-link" data-path="${fAttr}">📂 פתח תיקייה</a></div>`;
        html += `<div class="si-row"><span class="si-path" dir="ltr" title="${fAttr}">${escHtml(folder)}</span></div>`;
        if (info.file)
            html += `<div class="si-row">📄 <span class="si-label">קובץ:</span> ` +
                    `<span class="si-path" dir="ltr" title="${escHtml(info.file)}">${escHtml(info.file)}</span></div>`;

        if (Array.isArray(info.attachments) && info.attachments.length) {
            html += `<div class="si-row" style="margin-top:6px">📎 <span class="si-label">קבצים מצורפים:</span></div>`;
            html += info.attachments.map(n =>
                `<div class="si-att"><span class="si-path" dir="ltr">${escHtml(n)}</span></div>`
            ).join("");
        }
    }

    body.innerHTML = html;
    body.querySelectorAll(".si-link").forEach(a => {
        a.onclick = (e) => { e.preventDefault(); openFolderForUser(a.dataset.path); };
    });
    card.style.display = "block";
}

// =====================================================================
// Bug report
// =====================================================================
function onBugReport() {
    const version = _liveVersion;
    const date    = new Date().toLocaleDateString("he-IL");
    const body = `
<div dir="rtl" style="font-family:Segoe UI,Arial,sans-serif;font-size:14px;line-height:1.6">
  <p>שלום עופר,</p>
  <p>נתקלתי בתקלה בתוסף SaveAsPDF ואשמח לעזרה.</p>
  <p><strong>תיאור קצר של התקלה:</strong><br>
  [אנא החלף/י שורה זו בתיאור קצר של מה שקרה]</p>
  <p><strong>צילום מסך:</strong><br>
  [אנא צרף/י צילום מסך של הודעת השגיאה או ההתנהגות הלא-תקינה]</p>
  <hr style="border:none;border-top:1px solid #ddd;margin:12px 0">
  <p style="color:#888;font-size:12px">גרסה: ${escHtml(version)} &nbsp;|&nbsp; תאריך: ${escHtml(date)}</p>
</div>`;

    if (typeof Office?.context?.mailbox?.displayNewMessageForm === "function") {
        try {
            Office.context.mailbox.displayNewMessageForm({
                toRecipients: ["ofer@sw-eng.co.il"],
                subject:      "SaveAsPDF bug report",
                htmlBody:     body
            });
        } catch (e) {
            dlgAlert("דיווח על תקלה", "לא ניתן לפתוח חלון הודעה חדשה:\n" + e.message);
        }
    } else {
        dlgAlert("דיווח על תקלה",
            "פתיחת הודעה חדשה אינה נתמכת בסביבה זו.\n" +
            "אנא שלח/י מייל ידנית אל ofer@sw-eng.co.il עם נושא: SaveAsPDF bug report");
    }
}

// =====================================================================
// Auto-add current Outlook user as project leader (if setting is on)
// =====================================================================
function applyAddSelf() {
    if (!_prefs.addSelfAsEmployee) return;
    try {
        const up = Office.context.mailbox.userProfile;
        const email = up?.emailAddress || "";
        if (!email) return;
        // Don't act if the user is already in the list or a leader is already set
        if (_employees.some(e => e.email === email)) return;
        const leaderInput = document.getElementById("projectLeader");
        if (leaderInput.dataset.email || leaderInput.value.trim()) return;
        setProjectLeader({ displayName: up.displayName || email, email });
    } catch { }
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

        // Refresh the folder tree whenever the Folders tab is opened and a
        // project number is present, so newly-created folders on disk show up.
        if (btn.dataset.tab === "foldersTab") {
            const pid = document.getElementById("projectId").value.trim();
            if (pid) loadFolderTree(pid);
        }
    }));
}

// =====================================================================
// About panel
// =====================================================================
function openAboutPanel() {
    const panel = document.getElementById("aboutPanel");
    const verEl = document.getElementById("aboutVersion");
    if (verEl) verEl.textContent = _liveVersion || "…";
    panel.style.display = "flex";
    document.body.classList.add("overlay-open");
}
function closeAboutPanel() {
    document.getElementById("aboutPanel").style.display = "none";
    document.body.classList.remove("overlay-open");
}

// =====================================================================
// Pinned task pane — reload email-specific data on item change
// =====================================================================
function loadCurrentItem() {
    try {
        loadAttachments();
        tryRestoreProjectInfo();
        document.getElementById("status").textContent = "";
        updateDestBanner();
    } catch { }
}
