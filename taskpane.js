/* Quick Text – Word add-in
 * Stores your snippets locally and inserts the chosen one at the cursor.
 */

const STORE_KEY = "quicktext.snippets.v1";
const ALL_GROUPS = "__all__";

let snippets = [];
let editingId = null;
let pendingDeleteId = null;
let pendingDeleteTimer = null;
let toastTimer = null;
let memoryFallback = null; // used only if localStorage is unavailable

const $ = (id) => document.getElementById(id);

function uid() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
}

/* ---------- storage ---------- */

function loadSnippets() {
  let raw = null;
  try {
    raw = window.localStorage.getItem(STORE_KEY);
  } catch (e) {
    raw = memoryFallback;
  }
  if (raw) {
    try {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed)) return parsed.filter(isValidSnippet);
    } catch (e) { /* fall through to samples */ }
  }
  return [
    { id: uid(), title: "Greeting", group: "General", text: "Dear Sir/Madam,\n\nThank you for your message." },
    { id: uid(), title: "Closing", group: "General", text: "Kind regards," },
  ];
}

function saveSnippets() {
  const raw = JSON.stringify(snippets);
  try {
    window.localStorage.setItem(STORE_KEY, raw);
  } catch (e) {
    memoryFallback = raw;
  }
}

function isValidSnippet(s) {
  return s && typeof s.title === "string" && typeof s.text === "string";
}

/* ---------- views ---------- */

function showView(name) {
  $("view-list").hidden = name !== "list";
  $("view-edit").hidden = name !== "edit";
  $("view-backup").hidden = name !== "backup";
  $("btn-new").hidden = name !== "list";
  $("btn-backup").hidden = name !== "list";
}

function groupsList() {
  const set = new Set();
  snippets.forEach((s) => { if (s.group) set.add(s.group); });
  return Array.from(set).sort((a, b) => a.localeCompare(b));
}

function renderGroupFilter() {
  const select = $("group-filter");
  const current = select.value || ALL_GROUPS;
  select.innerHTML = "";
  const all = document.createElement("option");
  all.value = ALL_GROUPS;
  all.textContent = "All groups";
  select.appendChild(all);
  groupsList().forEach((g) => {
    const o = document.createElement("option");
    o.value = g;
    o.textContent = g;
    select.appendChild(o);
  });
  select.value = groupsList().includes(current) ? current : ALL_GROUPS;

  const dl = $("group-options");
  dl.innerHTML = "";
  groupsList().forEach((g) => {
    const o = document.createElement("option");
    o.value = g;
    dl.appendChild(o);
  });
}

function renderList() {
  renderGroupFilter();
  const q = $("search").value.trim().toLowerCase();
  const group = $("group-filter").value;
  const list = $("snippet-list");
  list.innerHTML = "";

  const shown = snippets.filter((s) => {
    if (group !== ALL_GROUPS && (s.group || "") !== group) return false;
    if (!q) return true;
    return (s.title + " " + s.text + " " + (s.group || "")).toLowerCase().includes(q);
  });

  $("empty").hidden = shown.length > 0;

  shown.forEach((s) => {
    const li = document.createElement("li");
    li.className = "snippet";

    const main = document.createElement("button");
    main.type = "button";
    main.className = "snippet-main";
    main.title = "Click to insert at the cursor";
    main.addEventListener("click", () => insertSnippet(s));

    const title = document.createElement("div");
    title.className = "snippet-title";
    const t = document.createElement("span");
    t.textContent = s.title || "(untitled)";
    title.appendChild(t);
    if (s.group) {
      const tag = document.createElement("span");
      tag.className = "tag";
      tag.textContent = s.group;
      title.appendChild(tag);
    }
    const preview = document.createElement("div");
    preview.className = "snippet-preview";
    preview.textContent = s.text;
    main.appendChild(title);
    main.appendChild(preview);

    const actions = document.createElement("div");
    actions.className = "snippet-actions";

    const edit = document.createElement("button");
    edit.type = "button";
    edit.textContent = "Edit";
    edit.addEventListener("click", () => openEditor(s.id));

    const del = document.createElement("button");
    del.type = "button";
    del.className = "danger";
    if (pendingDeleteId === s.id) {
      del.classList.add("confirm");
      del.textContent = "Confirm delete";
    } else {
      del.textContent = "Delete";
    }
    del.addEventListener("click", () => requestDelete(s.id));

    actions.appendChild(edit);
    actions.appendChild(del);

    li.appendChild(main);
    li.appendChild(actions);
    list.appendChild(li);
  });
}

/* ---------- insert ---------- */

async function insertSnippet(s) {
  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      const range = selection.insertText(s.text, Word.InsertLocation.replace);
      range.select(Word.SelectionMode.end); // leave the cursor after the inserted text
      await context.sync();
    });
    toast("Inserted: " + (s.title || "snippet"));
  } catch (err) {
    toast("Could not insert: " + (err && err.message ? err.message : err), true);
  }
}

/* ---------- editor ---------- */

function openEditor(id) {
  editingId = id || null;
  const s = id ? snippets.find((x) => x.id === id) : null;
  $("edit-heading").textContent = s ? "Edit snippet" : "New snippet";
  $("f-title").value = s ? s.title : "";
  $("f-group").value = s ? s.group || "" : "";
  $("f-text").value = s ? s.text : "";
  $("edit-error").hidden = true;
  showView("edit");
  $("f-title").focus();
}

function saveEditor() {
  const title = $("f-title").value.trim();
  const group = $("f-group").value.trim();
  const text = $("f-text").value.replace(/\r\n/g, "\n");

  if (!title) return editorError("Please give the snippet a title.");
  if (!text.trim()) return editorError("The text to insert cannot be empty.");

  if (editingId) {
    const s = snippets.find((x) => x.id === editingId);
    if (s) { s.title = title; s.group = group; s.text = text; }
  } else {
    snippets.unshift({ id: uid(), title, group, text });
  }
  saveSnippets();
  closeEditor();
  toast("Saved");
}

function editorError(msg) {
  const el = $("edit-error");
  el.textContent = msg;
  el.hidden = false;
}

function closeEditor() {
  editingId = null;
  showView("list");
  renderList();
}

/* ---------- delete (two-step, no confirm() dialog) ---------- */

function requestDelete(id) {
  if (pendingDeleteId === id) {
    clearTimeout(pendingDeleteTimer);
    pendingDeleteId = null;
    snippets = snippets.filter((s) => s.id !== id);
    saveSnippets();
    renderList();
    toast("Deleted");
    return;
  }
  pendingDeleteId = id;
  clearTimeout(pendingDeleteTimer);
  pendingDeleteTimer = setTimeout(() => {
    pendingDeleteId = null;
    renderList();
  }, 3500);
  renderList();
}

/* ---------- backup / restore ---------- */

function openBackup() {
  $("backup-text").value = JSON.stringify(snippets, null, 2);
  $("backup-msg").hidden = true;
  showView("backup");
}

function backupMessage(msg) {
  const el = $("backup-msg");
  el.textContent = msg;
  el.hidden = false;
}

function importBackup(replace) {
  let parsed;
  try {
    parsed = JSON.parse($("backup-text").value);
  } catch (e) {
    return backupMessage("That text is not a valid backup.");
  }
  if (!Array.isArray(parsed)) return backupMessage("That text is not a valid backup.");
  const valid = parsed.filter(isValidSnippet).map((s) => ({
    id: uid(),
    title: String(s.title),
    group: typeof s.group === "string" ? s.group : "",
    text: String(s.text),
  }));
  if (!valid.length) return backupMessage("No snippets found in that text.");

  snippets = replace ? valid : valid.concat(snippets);
  saveSnippets();
  backupMessage("Imported " + valid.length + " snippet(s).");
  $("backup-text").value = JSON.stringify(snippets, null, 2);
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
  snippets = loadSnippets();
  saveSnippets();

  $("btn-new").addEventListener("click", () => openEditor(null));
  $("btn-save").addEventListener("click", saveEditor);
  $("btn-cancel").addEventListener("click", closeEditor);
  $("btn-backup").addEventListener("click", openBackup);
  $("btn-backup-close").addEventListener("click", () => { showView("list"); renderList(); });
  $("btn-copy").addEventListener("click", copyBackup);
  $("btn-import-merge").addEventListener("click", () => importBackup(false));
  $("btn-import-replace").addEventListener("click", () => importBackup(true));
  $("search").addEventListener("input", renderList);
  $("group-filter").addEventListener("change", renderList);

  showView("list");
  renderList();
}

if (typeof Office !== "undefined" && Office.onReady) {
  Office.onReady(() => init());
} else {
  document.addEventListener("DOMContentLoaded", init);
}
