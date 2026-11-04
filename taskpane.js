/* Quick Text – Word add-in
 * Snippets with folders (tabs), favorites, pins and multiple tags.
 * Everything is stored locally; the chosen snippet is inserted at the cursor.
 */

const KEY_V2 = "quicktext.data.v2";
const KEY_V1 = "quicktext.snippets.v1"; // old format, migrated automatically
const KEY_UI = "quicktext.ui.v1";
const ALL_TAGS = "__all__";

let data = { folders: [], snippets: [] };
let ui = { tab: "all", tag: ALL_TAGS, taW: 0, taH: 0 };
let editingId = null;
let editTags = [];
let folderMode = null; // "new" | "rename" | null
let pendingDelete = null; // snippet id or "folder:<id>"
let pendingTimer = null;
let toastTimer = null;
const expanded = new Set();
const memory = {}; // fallback if localStorage is unavailable

const $ = (id) => document.getElementById(id);

function uid() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
}

/* ---------- storage ---------- */

function readKey(key) {
  try { return window.localStorage.getItem(key); } catch (e) { return memory[key] || null; }
}
function writeKey(key, val) {
  try { window.localStorage.setItem(key, val); } catch (e) { memory[key] = val; }
}

function cleanTags(arr) {
  const out = [];
  (Array.isArray(arr) ? arr : []).forEach((t) => {
    const v = String(t).trim();
    if (v && !out.some((x) => x.toLowerCase() === v.toLowerCase())) out.push(v);
  });
  return out;
}

function normSnippet(s) {
  let tags = cleanTags(s.tags);
  if (!tags.length && typeof s.group === "string" && s.group.trim()) tags = [s.group.trim()]; // from v1
  return {
    id: typeof s.id === "string" && s.id ? s.id : uid(),
    title: String(s.title || ""),
    text: String(s.text || ""),
    tags,
    folder: typeof s.folder === "string" ? s.folder : "",
    fav: !!s.fav,
    pinned: !!s.pinned,
  };
}

function isValidSnippet(s) {
  return s && typeof s.title === "string" && typeof s.text === "string";
}

function load() {
  const rawV2 = readKey(KEY_V2);
  if (rawV2) {
    try {
      const p = JSON.parse(rawV2);
      if (p && Array.isArray(p.snippets)) {
        data.folders = (Array.isArray(p.folders) ? p.folders : [])
          .filter((f) => f && typeof f.id === "string" && typeof f.name === "string");
        data.snippets = p.snippets.filter(isValidSnippet).map(normSnippet);
        return;
      }
    } catch (e) { /* fall through */ }
  }
  const rawV1 = readKey(KEY_V1);
  if (rawV1) {
    try {
      const p = JSON.parse(rawV1);
      if (Array.isArray(p)) {
        data.snippets = p.filter(isValidSnippet).map(normSnippet);
        return;
      }
    } catch (e) { /* fall through */ }
  }
  const letters = { id: uid(), name: "Letters" };
  data.folders = [letters];
  data.snippets = [
    normSnippet({ title: "Greeting", tags: ["General", "Opening"], folder: letters.id, fav: true, text: "Dear Sir/Madam,\n\nThank you for your message." }),
    normSnippet({ title: "Closing", tags: ["General"], folder: letters.id, text: "Kind regards," }),
  ];
}

function save() { writeKey(KEY_V2, JSON.stringify(data)); }

function loadUi() {
  try {
    const p = JSON.parse(readKey(KEY_UI) || "{}");
    if (p && typeof p === "object") ui = Object.assign(ui, p);
  } catch (e) { /* ignore */ }
  const valid = ui.tab === "all" || ui.tab === "fav" || data.folders.some((f) => f.id === ui.tab);
  if (!valid) ui.tab = "all";
}
function saveUi() { writeKey(KEY_UI, JSON.stringify(ui)); }

/* ---------- helpers ---------- */

function allTags() {
  const map = new Map();
  data.snippets.forEach((s) => s.tags.forEach((t) => { if (!map.has(t.toLowerCase())) map.set(t.toLowerCase(), t); }));
  return Array.from(map.values()).sort((a, b) => a.localeCompare(b));
}

function folderById(id) { return data.folders.find((f) => f.id === id); }

function hasTag(s, tag) { return s.tags.some((t) => t.toLowerCase() === tag.toLowerCase()); }

function showView(name) {
  $("view-list").hidden = name !== "list";
  $("view-edit").hidden = name !== "edit";
  $("view-backup").hidden = name !== "backup";
  $("btn-new").hidden = name !== "list";
  $("btn-backup").hidden = name !== "list";
}

/* ---------- tabs & folders ---------- */

function countFor(tab) {
  if (tab === "all") return data.snippets.length;
  if (tab === "fav") return data.snippets.filter((s) => s.fav).length;
  return data.snippets.filter((s) => s.folder === tab).length;
}

function renderTabs() {
  const box = $("tabs");
  box.innerHTML = "";
  const tabs = [{ id: "all", label: "All" }, { id: "fav", label: "★ Favorites" }]
    .concat(data.folders.map((f) => ({ id: f.id, label: "📁 " + f.name })));
  tabs.forEach((t) => {
    const b = document.createElement("button");
    b.type = "button";
    b.className = "tab" + (ui.tab === t.id ? " active" : "");
    b.setAttribute("role", "tab");
    b.setAttribute("aria-selected", ui.tab === t.id ? "true" : "false");
    b.textContent = t.label + " (" + countFor(t.id) + ")";
    b.addEventListener("click", () => selectTab(t.id));
    box.appendChild(b);
  });
}

function selectTab(id) {
  ui.tab = id;
  saveUi();
  closeFolderForm();
  clearPending();
  renderList();
}

function renderFolderTools() {
  const f = folderById(ui.tab);
  $("folder-tools").hidden = !f;
  if (!f) return;
  $("folder-title").textContent = "📁 " + f.name;
  const del = $("btn-folder-delete");
  if (pendingDelete === "folder:" + f.id) {
    del.textContent = "Confirm delete";
    del.classList.add("confirm");
  } else {
    del.textContent = "Delete folder";
    del.classList.remove("confirm");
  }
}

function openFolderForm(mode) {
  folderMode = mode;
  const f = folderById(ui.tab);
  $("folder-name").value = mode === "rename" && f ? f.name : "";
  $("folder-form").hidden = false;
  $("folder-error").hidden = true;
  $("folder-name").focus();
}

function closeFolderForm() {
  folderMode = null;
  $("folder-form").hidden = true;
  $("folder-error").hidden = true;
}

function folderError(msg) {
  $("folder-error").textContent = msg;
  $("folder-error").hidden = false;
}

function saveFolder() {
  const name = $("folder-name").value.trim();
  if (!name) return folderError("Please type a folder name.");
  const clash = data.folders.some((f) => f.name.toLowerCase() === name.toLowerCase() &&
    !(folderMode === "rename" && f.id === ui.tab));
  if (clash) return folderError("A folder with that name already exists.");

  if (folderMode === "rename") {
    const f = folderById(ui.tab);
    if (f) f.name = name;
  } else {
    const f = { id: uid(), name };
    data.folders.push(f);
    ui.tab = f.id;
    saveUi();
  }
  save();
  closeFolderForm();
  renderList();
}

function deleteFolder() {
  const f = folderById(ui.tab);
  if (!f) return;
  const key = "folder:" + f.id;
  if (pendingDelete !== key) { setPending(key); return; }
  clearPending();
  data.snippets.forEach((s) => { if (s.folder === f.id) s.folder = ""; });
  data.folders = data.folders.filter((x) => x.id !== f.id);
  ui.tab = "all";
  saveUi();
  save();
  renderList();
  toast("Folder deleted (its snippets were kept in All)");
}

/* ---------- pending (two-step) delete ---------- */

function setPending(key) {
  pendingDelete = key;
  clearTimeout(pendingTimer);
  pendingTimer = setTimeout(() => { pendingDelete = null; renderList(); }, 3500);
  renderList();
}
function clearPending() {
  pendingDelete = null;
  clearTimeout(pendingTimer);
}

/* ---------- list ---------- */

function renderTagFilter() {
  const sel = $("tag-filter");
  sel.innerHTML = "";
  const o0 = document.createElement("option");
  o0.value = ALL_TAGS;
  o0.textContent = "All tags";
  sel.appendChild(o0);
  const tags = allTags();
  tags.forEach((t) => {
    const o = document.createElement("option");
    o.value = t;
    o.textContent = "# " + t;
    sel.appendChild(o);
  });
  if (ui.tag !== ALL_TAGS && !tags.some((t) => t.toLowerCase() === ui.tag.toLowerCase())) ui.tag = ALL_TAGS;
  sel.value = tags.find((t) => t.toLowerCase() === ui.tag.toLowerCase()) || ALL_TAGS;
}

function renderList() {
  renderTabs();
  renderFolderTools();
  renderTagFilter();

  const q = $("search").value.trim().toLowerCase();
  const list = $("snippet-list");
  list.innerHTML = "";

  const shown = data.snippets.filter((s) => {
    if (ui.tab === "fav" && !s.fav) return false;
    if (ui.tab !== "all" && ui.tab !== "fav" && s.folder !== ui.tab) return false;
    if (ui.tag !== ALL_TAGS && !hasTag(s, ui.tag)) return false;
    if (!q) return true;
    return (s.title + " " + s.text + " " + s.tags.join(" ")).toLowerCase().includes(q);
  });
  shown.sort((a, b) => (b.pinned ? 1 : 0) - (a.pinned ? 1 : 0)); // pinned first, otherwise keep order

  $("empty").hidden = shown.length > 0;
  $("empty").textContent = ui.tab === "fav" ? "No favorites yet. Tap ☆ on a snippet." : "No snippets here.";

  shown.forEach((s) => list.appendChild(buildCard(s)));
}

function iconButton(text, label, active, onClick) {
  const b = document.createElement("button");
  b.type = "button";
  b.className = "icon-btn" + (active ? " on" : "");
  b.textContent = text;
  b.title = label;
  b.setAttribute("aria-label", label);
  b.setAttribute("aria-pressed", active ? "true" : "false");
  b.addEventListener("click", onClick);
  return b;
}

function buildCard(s) {
  const li = document.createElement("li");
  li.className = "snippet" + (s.pinned ? " pinned" : "");

  const head = document.createElement("div");
  head.className = "snippet-head";

  const main = document.createElement("button");
  main.type = "button";
  main.className = "snippet-main";
  main.title = "Click to insert at the cursor";
  main.addEventListener("click", () => insertSnippet(s));

  const title = document.createElement("div");
  title.className = "snippet-title";
  title.textContent = s.title || "(untitled)";
  main.appendChild(title);

  const icons = document.createElement("div");
  icons.className = "snippet-icons";
  icons.appendChild(iconButton("📌", s.pinned ? "Unpin" : "Pin to top", s.pinned, () => toggle(s, "pinned")));
  icons.appendChild(iconButton(s.fav ? "★" : "☆", s.fav ? "Remove from favorites" : "Add to favorites", s.fav, () => toggle(s, "fav")));

  head.appendChild(main);
  head.appendChild(icons);
  li.appendChild(head);

  const meta = document.createElement("div");
  meta.className = "snippet-meta";
  const f = folderById(s.folder);
  if (f && ui.tab !== s.folder) {
    const fb = document.createElement("span");
    fb.className = "folder-badge";
    fb.textContent = "📁 " + f.name;
    meta.appendChild(fb);
  }
  s.tags.forEach((t) => {
    const tag = document.createElement("button");
    tag.type = "button";
    tag.className = "tag";
    tag.textContent = "# " + t;
    tag.title = "Filter by this tag";
    tag.addEventListener("click", () => { ui.tag = t; saveUi(); renderList(); });
    meta.appendChild(tag);
  });
  if (meta.childNodes.length) li.appendChild(meta);

  const open = expanded.has(s.id);
  const preview = document.createElement("div");
  preview.className = "snippet-preview" + (open ? " open" : "");
  preview.textContent = s.text;
  li.appendChild(preview);

  const actions = document.createElement("div");
  actions.className = "snippet-actions";

  if (s.text.length > 90 || s.text.indexOf("\n") !== -1) {
    const more = document.createElement("button");
    more.type = "button";
    more.textContent = open ? "Show less" : "Show more";
    more.addEventListener("click", () => {
      if (expanded.has(s.id)) expanded.delete(s.id); else expanded.add(s.id);
      renderList();
    });
    actions.appendChild(more);
  }

  const edit = document.createElement("button");
  edit.type = "button";
  edit.textContent = "Edit";
  edit.addEventListener("click", () => openEditor(s.id));

  const del = document.createElement("button");
  del.type = "button";
  del.className = "danger";
  if (pendingDelete === s.id) { del.classList.add("confirm"); del.textContent = "Confirm delete"; }
  else del.textContent = "Delete";
  del.addEventListener("click", () => requestDelete(s.id));

  actions.appendChild(edit);
  actions.appendChild(del);
  li.appendChild(actions);
  return li;
}

function toggle(s, field) {
  s[field] = !s[field];
  save();
  renderList();
}

function requestDelete(id) {
  if (pendingDelete === id) {
    clearPending();
    data.snippets = data.snippets.filter((s) => s.id !== id);
    save();
    renderList();
    toast("Deleted");
    return;
  }
  setPending(id);
}

/* ---------- insert ---------- */

async function insertSnippet(s) {
  try {
    await Word.run(async (context) => {
      const range = context.document.getSelection().insertText(s.text, Word.InsertLocation.replace);
      range.select(Word.SelectionMode.end); // leave the cursor after the inserted text
      await context.sync();
    });
    toast("Inserted: " + (s.title || "snippet"));
  } catch (err) {
    toast("Could not insert: " + (err && err.message ? err.message : err), true);
  }
}

/* ---------- editor ---------- */

function renderFolderSelect(selected) {
  const sel = $("f-folder");
  sel.innerHTML = "";
  const none = document.createElement("option");
  none.value = "";
  none.textContent = "(No folder)";
  sel.appendChild(none);
  data.folders.forEach((f) => {
    const o = document.createElement("option");
    o.value = f.id;
    o.textContent = f.name;
    sel.appendChild(o);
  });
  sel.value = folderById(selected) ? selected : "";
}

function renderTagEditor() {
  const chips = $("tag-chips");
  chips.innerHTML = "";
  editTags.forEach((t) => {
    const c = document.createElement("span");
    c.className = "chip";
    c.appendChild(document.createTextNode("# " + t));
    const x = document.createElement("button");
    x.type = "button";
    x.className = "chip-x";
    x.textContent = "×";
    x.setAttribute("aria-label", "Remove tag " + t);
    x.addEventListener("click", () => { editTags = editTags.filter((v) => v !== t); renderTagEditor(); });
    c.appendChild(x);
    chips.appendChild(c);
  });

  const sug = $("tag-suggest");
  sug.innerHTML = "";
  const free = allTags().filter((t) => !editTags.some((e) => e.toLowerCase() === t.toLowerCase()));
  if (free.length) {
    const lab = document.createElement("span");
    lab.className = "muted small-text";
    lab.textContent = "Existing: ";
    sug.appendChild(lab);
    free.forEach((t) => {
      const b = document.createElement("button");
      b.type = "button";
      b.className = "tag";
      b.textContent = "+ " + t;
      b.addEventListener("click", () => addTag(t));
      sug.appendChild(b);
    });
  }
}

function addTag(raw) {
  const t = String(raw).replace(/,/g, " ").trim();
  if (!t) return;
  if (!editTags.some((e) => e.toLowerCase() === t.toLowerCase())) editTags.push(t);
  renderTagEditor();
}

function commitTagInput() {
  const inp = $("f-tag-input");
  if (inp.value.trim()) addTag(inp.value);
  inp.value = "";
}

function openEditor(id) {
  editingId = id || null;
  const s = id ? data.snippets.find((x) => x.id === id) : null;
  $("edit-heading").textContent = s ? "Edit snippet" : "New snippet";
  $("f-title").value = s ? s.title : "";
  $("f-text").value = s ? s.text : "";
  $("f-fav").checked = s ? s.fav : ui.tab === "fav";
  $("f-pin").checked = s ? s.pinned : false;
  const defFolder = s ? s.folder : (folderById(ui.tab) ? ui.tab : "");
  renderFolderSelect(defFolder);
  editTags = s ? s.tags.slice() : [];
  $("f-tag-input").value = "";
  renderTagEditor();
  $("edit-error").hidden = true;
  showView("edit");
  applyTextSize();
  $("f-title").focus();
}

function saveEditor() {
  commitTagInput();
  const title = $("f-title").value.trim();
  const text = $("f-text").value.replace(/\r\n/g, "\n");
  if (!title) return editorError("Please give the snippet a title.");
  if (!text.trim()) return editorError("The text to insert cannot be empty.");

  const fields = {
    title,
    text,
    tags: cleanTags(editTags),
    folder: $("f-folder").value,
    fav: $("f-fav").checked,
    pinned: $("f-pin").checked,
  };
  if (editingId) {
    const s = data.snippets.find((x) => x.id === editingId);
    if (s) Object.assign(s, fields);
  } else {
    data.snippets.unshift(normSnippet(fields));
  }
  save();
  closeEditor();
  toast("Saved");
}

function editorError(msg) {
  $("edit-error").textContent = msg;
  $("edit-error").hidden = false;
}

function closeEditor() {
  editingId = null;
  showView("list");
  renderList();
}

/* ---------- resizable text box ---------- */

function applyTextSize() {
  const ta = $("f-text");
  if (ui.taH > 60) ta.style.height = ui.taH + "px";
  if (ui.taW > 120) ta.style.width = ui.taW + "px";
}

function watchTextSize() {
  const ta = $("f-text");
  if (typeof ResizeObserver === "undefined") return;
  let t = null;
  new ResizeObserver(() => {
    if (ta.offsetHeight === 0) return; // editor hidden
    clearTimeout(t);
    t = setTimeout(() => {
      ui.taH = Math.round(ta.offsetHeight);
      ui.taW = Math.round(ta.offsetWidth);
      saveUi();
    }, 300);
  }).observe(ta);
}

/* ---------- backup / restore ---------- */

function openBackup() {
  $("backup-text").value = JSON.stringify({ version: 2, folders: data.folders, snippets: data.snippets }, null, 2);
  $("backup-msg").hidden = true;
  showView("backup");
}

function backupMessage(msg) {
  $("backup-msg").textContent = msg;
  $("backup-msg").hidden = false;
}

function importBackup(replace) {
  let p;
  try { p = JSON.parse($("backup-text").value); } catch (e) { return backupMessage("That text is not a valid backup."); }

  let inFolders = [];
  let inSnippets;
  if (Array.isArray(p)) inSnippets = p; // old backup format
  else if (p && Array.isArray(p.snippets)) { inSnippets = p.snippets; inFolders = Array.isArray(p.folders) ? p.folders : []; }
  else return backupMessage("That text is not a valid backup.");

  inSnippets = inSnippets.filter(isValidSnippet);
  if (!inSnippets.length) return backupMessage("No snippets found in that text.");

  const base = replace ? { folders: [], snippets: [] } : { folders: data.folders.slice(), snippets: data.snippets.slice() };

  // map imported folder ids to (existing or new) folders, matching by name
  const idMap = {};
  inFolders.filter((f) => f && typeof f.id === "string" && typeof f.name === "string").forEach((f) => {
    let hit = base.folders.find((x) => x.name.toLowerCase() === f.name.toLowerCase());
    if (!hit) { hit = { id: uid(), name: f.name }; base.folders.push(hit); }
    idMap[f.id] = hit.id;
  });

  const added = inSnippets.map((s) => {
    const n = normSnippet(Object.assign({}, s, { id: uid() }));
    n.folder = idMap[n.folder] || "";
    return n;
  });

  data = { folders: base.folders, snippets: added.concat(base.snippets) };
  save();
  backupMessage("Imported " + added.length + " snippet(s).");
  $("backup-text").value = JSON.stringify({ version: 2, folders: data.folders, snippets: data.snippets }, null, 2);
}

function copyBackup() {
  const ta = $("backup-text");
  ta.focus();
  ta.select();
  let ok = false;
  try { ok = document.execCommand("copy"); } catch (e) { ok = false; }
  backupMessage(ok ? "Copied to clipboard." : "Press Ctrl+C to copy the selected text.");
}

/* ---------- toast ---------- */

function toast(msg, bad) {
  const el = $("toast");
  el.textContent = msg;
  el.classList.toggle("bad", !!bad);
  el.hidden = false;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { el.hidden = true; }, bad ? 5000 : 2200);
}

/* ---------- init ---------- */

function init() {
  load();
  save();
  loadUi();

  $("btn-new").addEventListener("click", () => openEditor(null));
  $("btn-save").addEventListener("click", saveEditor);
  $("btn-cancel").addEventListener("click", closeEditor);
  $("btn-backup").addEventListener("click", openBackup);
  $("btn-backup-close").addEventListener("click", () => { showView("list"); renderList(); });
  $("btn-copy").addEventListener("click", copyBackup);
  $("btn-import-merge").addEventListener("click", () => importBackup(false));
  $("btn-import-replace").addEventListener("click", () => importBackup(true));
  $("search").addEventListener("input", renderList);
  $("tag-filter").addEventListener("change", () => { ui.tag = $("tag-filter").value; saveUi(); renderList(); });

  $("btn-add-folder").addEventListener("click", () => openFolderForm("new"));
  $("btn-folder-rename").addEventListener("click", () => openFolderForm("rename"));
  $("btn-folder-delete").addEventListener("click", deleteFolder);
  $("btn-folder-save").addEventListener("click", saveFolder);
  $("btn-folder-cancel").addEventListener("click", closeFolderForm);
  $("folder-name").addEventListener("keydown", (e) => {
    if (e.key === "Enter") { e.preventDefault(); saveFolder(); }
    if (e.key === "Escape") closeFolderForm();
  });

  const tagInput = $("f-tag-input");
  tagInput.addEventListener("keydown", (e) => {
    if (e.key === "Enter" || e.key === ",") {
      e.preventDefault();
      commitTagInput();
    } else if (e.key === "Backspace" && !tagInput.value && editTags.length) {
      editTags.pop();
      renderTagEditor();
    }
  });
  $("tag-box").addEventListener("click", () => tagInput.focus());

  watchTextSize();
  showView("list");
  renderList();
}

if (typeof Office !== "undefined" && Office.onReady) {
  Office.onReady(() => init());
} else {
  document.addEventListener("DOMContentLoaded", init);
}
