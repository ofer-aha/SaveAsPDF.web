/* global Office, __BACKEND_BASE__ */

const BACKEND = (typeof __BACKEND_BASE__ !== "undefined") ? __BACKEND_BASE__ : "http://localhost:5176";

Office.onReady(async () => {
    const mode = new URLSearchParams(window.location.search).get("mode") || "leader";
    const multiSelect = (mode === "employee");

    document.getElementById("statusBar").textContent = "טוען אנשי קשר...";

    // Fetch contacts directly from the backend — avoids localStorage isolation
    // between the taskpane and dialog (separate WebView2 processes in Outlook).
    let folders, contacts;
    try {
        const res = await fetch(`${BACKEND}/api/contacts`);
        if (!res.ok) {
            const body = await res.json().catch(() => ({}));
            throw new Error(body.error || `שגיאת שרת ${res.status}`);
        }
        ({ folders, contacts } = await res.json());
    } catch (e) {
        document.getElementById("loadingMsg").textContent = "שגיאה: " + e.message;
        document.getElementById("statusBar").textContent = "שגיאה בטעינת הנתונים";
        return;
    }

    if (!folders || folders.length === 0) {
        document.getElementById("loadingMsg").textContent = "לא נמצאו תיקיות אנשי קשר";
        document.getElementById("statusBar").textContent = "0 אנשי קשר";
        return;
    }

    let activeFolderId = folders[0].id;
    const selected = [];

    // ----------------------------------------------------------------
    // Folder tree
    // ----------------------------------------------------------------
    function renderFolders() {
        const tree = document.getElementById("folderTree");
        tree.innerHTML = "";

        folders.forEach(folder => {
            const depth = folder.parentId ? 1 : 0;
            const count = (contacts[folder.id] || []).length;

            const el = document.createElement("div");
            el.className = "folder-item" + (folder.id === activeFolderId ? " active" : "");
            el.style.paddingLeft = (10 + depth * 14) + "px";
            el.innerHTML = `
                <span class="folder-icon">${depth > 0 ? "📂" : "📁"}</span>
                <span class="folder-name">${folder.displayName}</span>
                <span class="folder-count">${count}</span>
            `;
            el.onclick = () => {
                activeFolderId = folder.id;
                renderFolders();
                renderContacts();
            };
            tree.appendChild(el);
        });
    }

    // ----------------------------------------------------------------
    // Contact list
    // ----------------------------------------------------------------
    function getVisibleContacts() {
        const query = document.getElementById("searchBox").value.toLowerCase().trim();
        if (query) {
            const all = Object.values(contacts).flat();
            return all.filter(c =>
                c.displayName.toLowerCase().includes(query) ||
                c.email.toLowerCase().includes(query)
            );
        }
        return contacts[activeFolderId] || [];
    }

    function isSelected(email) {
        return selected.some(s => s.email === email);
    }

    function renderContacts() {
        const list = document.getElementById("contactList");
        const items = getVisibleContacts();

        if (items.length === 0) {
            list.innerHTML = `<div class="empty-msg">אין אנשי קשר להצגה</div>`;
            updateStatus();
            return;
        }

        list.innerHTML = "";
        items.forEach(contact => {
            const row = document.createElement("div");
            row.className = "contact-item" + (isSelected(contact.email) ? " selected" : "");

            if (multiSelect) {
                const cb = document.createElement("input");
                cb.type = "checkbox";
                cb.checked = isSelected(contact.email);
                cb.tabIndex = -1;
                row.appendChild(cb);
            }

            const info = document.createElement("div");
            info.className = "contact-info";
            info.innerHTML = `
                <div class="name">${contact.displayName}</div>
                <div class="email">${contact.email}</div>
            `;
            row.appendChild(info);
            row.onclick = () => toggleContact(contact, row);
            list.appendChild(row);
        });

        updateStatus();
    }

    function toggleContact(contact, row) {
        const idx = selected.findIndex(s => s.email === contact.email);
        if (idx >= 0) {
            selected.splice(idx, 1);
            row.classList.remove("selected");
            const cb = row.querySelector("input[type=checkbox]");
            if (cb) cb.checked = false;
        } else {
            if (!multiSelect) {
                selected.length = 0;
                document.querySelectorAll(".contact-item.selected").forEach(el => {
                    el.classList.remove("selected");
                    const cb = el.querySelector("input[type=checkbox]");
                    if (cb) cb.checked = false;
                });
            }
            selected.push(contact);
            row.classList.add("selected");
            const cb = row.querySelector("input[type=checkbox]");
            if (cb) cb.checked = true;
        }
        updateStatus();
    }

    function updateStatus() {
        const total = getVisibleContacts().length;
        document.getElementById("statusBar").textContent =
            `${total} אנשי קשר | ${selected.length} נבחרו`;
        document.getElementById("btnSelect").disabled = selected.length === 0;
    }

    // ----------------------------------------------------------------
    // Search
    // ----------------------------------------------------------------
    document.getElementById("searchBox").oninput = renderContacts;

    // ----------------------------------------------------------------
    // Buttons
    // ----------------------------------------------------------------
    document.getElementById("btnSelect").onclick = () => {
        if (selected.length === 0) return;
        Office.context.ui.messageParent(JSON.stringify({
            type: "selection",
            mode,
            contacts: selected
        }));
    };

    document.getElementById("btnClose").onclick = () => {
        Office.context.ui.messageParent(JSON.stringify({ type: "cancel" }));
    };

    // ----------------------------------------------------------------
    // Init
    // ----------------------------------------------------------------
    renderFolders();
    renderContacts();
});
