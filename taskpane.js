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

  const list = $("snippet-list");
  list.innerHTML = "";

  const shown = visibleSnippets();

  $("empty").hidden = shown.length > 0;
  $("empty").textContent = ui.tab === "fav" ? "No favorites yet. Tap ☆ on a snippet." : "No snippets here.";

  shown.forEach((s) => list.appendChild(buildCard(s)));
}

/* Snippets currently shown (tab + tag + search), pinned first, otherwise in saved order. */
function visibleSnippets() {
  const q = $("search").value.trim().toLowerCase();
  const shown = data.snippets.filter((s) => {
    if (ui.tab === "fav" && !s.fav) return false;
    if (ui.tab !== "all" && ui.tab !== "fav" && s.folder !== ui.tab) return false;
    if (ui.tag !== ALL_TAGS && !hasTag(s, ui.tag)) return false;
    if (!q) return true;
    return (s.title + " " + s.text + " " + s.tags.join(" ")).toLowerCase().includes(q);
  });
  shown.sort((a, b) => (b.pinned ? 1 : 0) - (a.pinned ? 1 : 0));
  return shown;
}

/* ---------- drag to reorder ---------- */
/* Uses pointer events (mouse, pen and touch) instead of HTML5 drag-and-drop, which is unreliable in Office task panes. */

let drag = null; // { id, li, pid, startY, moved, targetId, after }

function clearDropMarks() {
  document.querySelectorAll(".drop-before, .drop-after").forEach((el) => el.classList.remove("drop-before", "drop-after"));
}

function snippetById(id) { return data.snippets.find((s) => s.id === id); }

/* Move snippet `id` next to `targetId` (before or after it). Works inside the current view:
 * the visible snippets keep occupying the same slots of the saved order, just rearranged. */
function reorder(id, targetId, after) {
  const dragged = snippetById(id);
  const target = snippetById(targetId);
  if (!dragged || !target || id === targetId || dragged.pinned !== target.pinned) return false;

  const group = visibleSnippets().filter((s) => s.pinned === dragged.pinned);
  const ids = group.map((s) => s.id).filter((x) => x !== id);
  let at = ids.indexOf(targetId);
  if (at < 0) return false;
  if (after) at++;
  ids.splice(at, 0, id);

  const inGroup = new Set(group.map((s) => s.id));
  const slots = [];
  data.snippets.forEach((s, i) => { if (inGroup.has(s.id)) slots.push(i); });
  const ordered = ids.map(snippetById);
  slots.forEach((pos, i) => { data.snippets[pos] = ordered[i]; });
  save();
  return true;
}

/* Move one step up (-1) or down (+1) within the same pinned group. */
function step(s, dir) {
  const group = visibleSnippets().filter((x) => x.pinned === s.pinned);
  const i = group.findIndex((x) => x.id === s.id);
  const j = i + dir;
  if (i < 0 || j < 0 || j >= group.length) return false;
  return reorder(s.id, group[j].id, dir > 0);
}

function focusControl(selector) {
  const el = document.querySelector(selector);
  if (el) el.focus();
}

function endDrag() {
  if (drag) {
    drag.li.classList.remove("dragging");
    drag = null;
  }
  document.body.classList.remove("is-dragging");
  clearDropMarks();
}

function onDragMove(e) {
  if (!drag || e.pointerId !== drag.pid) return;
  if (!drag.moved && Math.abs(e.clientY - drag.startY) < 4) return;
  drag.moved = true;
  e.preventDefault();
  drag.li.classList.add("dragging");
  document.body.classList.add("is-dragging");

  // find the drop position among cards of the same pinned state
  const pinned = drag.li.getAttribute("data-pinned");
  const cards = Array.prototype.filter.call($("snippet-list").children, (c) => c.getAttribute("data-pinned") === pinned);
  clearDropMarks();
  drag.targetId = null;
  for (let k = 0; k < cards.length; k++) {
    const c = cards[k];
    const r = c.getBoundingClientRect();
    if (e.clientY < r.top + r.height / 2) { drag.targetId = c.getAttribute("data-id"); drag.after = false; c.classList.add("drop-before"); break; }
  }
  if (!drag.targetId && cards.length) {
    const last = cards[cards.length - 1];
    drag.targetId = last.getAttribute("data-id");
    drag.after = true;
    last.classList.add("drop-after");
  }
  if (drag.targetId === drag.id) { clearDropMarks(); }

  // scroll the pane when near its top/bottom edge
  if (e.clientY < 50) window.scrollBy(0, -14);
  else if (e.clientY > window.innerHeight - 50) window.scrollBy(0, 14);
}

function onDragEnd(e) {
  if (!drag || e.pointerId !== drag.pid) return;
  const d = drag;
  const doMove = d.moved && d.targetId && e.type === "pointerup";
  try { if (e.target.releasePointerCapture) e.target.releasePointerCapture(d.pid); } catch (err) { /* ignore */ }
  endDrag();
  if (doMove && reorder(d.id, d.targetId, d.after)) {
    renderList();
    focusControl('.drag-handle[data-id="' + d.id + '"]');
  }
}

function buildHandle(s, li) {
  const h = document.createElement("button");
  h.type = "button";
  h.className = "drag-handle";
  h.setAttribute("data-id", s.id);
  h.textContent = "⠿";
  h.title = "Drag up or down to reorder";
  h.setAttribute("aria-label", "Reorder " + (s.title || "snippet") + ". Press up or down arrow to move.");

  h.addEventListener("pointerdown", (e) => {
    if (e.button !== undefined && e.button !== 0) return;
    e.preventDefault();
    drag = { id: s.id, li, pid: e.pointerId, startY: e.clientY, moved: false, targetId: null, after: false };
    try { if (h.setPointerCapture) h.setPointerCapture(e.pointerId); } catch (err) { /* ignore */ }
  });
  h.addEventListener("pointermove", onDragMove);
  h.addEventListener("pointerup", onDragEnd);
  h.addEventListener("pointercancel", onDragEnd);

  h.addEventListener("keydown", (e) => {
    if (e.key !== "ArrowUp" && e.key !== "ArrowDown") return;
    e.preventDefault();
    if (step(s, e.key === "ArrowUp" ? -1 : 1)) {
      renderList();
      focusControl('.drag-handle[data-id="' + s.id + '"]');
    }
  });
  return h;
}

function iconButton(text, label, active, onClick, extraClass) {
  const b = document.createElement("button");
  b.type = "button";
  b.className = "icon-btn" + (active ? " on" : "") + (extraClass ? " " + extraClass : "");
  b.textContent = text;
  b.title = label;
  b.setAttribute("aria-label", label);
  if (active !== null) b.setAttribute("aria-pressed", active ? "true" : "false"); // null = plain action, not a toggle
  b.addEventListener("click", onClick);
  return b;
}

/* True when a click landed on a button or the expanded preview, i.e. not a "plain" click on the card. */
function isControl(el, card) {
  while (el && el !== card) {
    if (el.tagName === "BUTTON" || (el.classList && el.classList.contains("snippet-detail"))) return true;
    el = el.parentNode;
  }
  return false;
}

function buildCard(s) {
  const li = document.createElement("li");
  li.className = "snippet" + (s.pinned ? " pinned" : "");
  li.setAttribute("data-id", s.id);
  li.setAttribute("data-pinned", s.pinned ? "1" : "0");
  li.tabIndex = 0;
  li.title = "Click to insert at the cursor";

  // the whole card inserts the snippet; buttons inside keep their own action
  li.addEventListener("click", (e) => {
    if (isControl(e.target, li)) return;
    insertSnippet(s);
  });
  li.addEventListener("keydown", (e) => {
    if (e.target !== li) return;
    if (e.key === "Enter" || e.key === " ") { e.preventDefault(); insertSnippet(s); }
  });

  const open = expanded.has(s.id);

  // one compact row: handle | title + folder/tag badges | icon buttons
  const head = document.createElement("div");
  head.className = "snippet-head";
  head.appendChild(buildHandle(s, li));

  const line = document.createElement("div");
  line.className = "snippet-line";
  const title = document.createElement("span");
  title.className = "snippet-title";
  title.textContent = s.title || "(untitled)";
  line.appendChild(title);

  const f = folderById(s.folder);
  if (f && ui.tab !== s.folder) {
    const fb = document.createElement("span");
    fb.className = "folder-badge";
    fb.textContent = "📁 " + f.name;
    line.appendChild(fb);
  }
  s.tags.forEach((t) => {
    const tag = document.createElement("button");
    tag.type = "button";
    tag.className = "tag";
    tag.textContent = "# " + t;
    tag.title = "Filter by this tag";
    tag.addEventListener("click", () => { ui.tag = t; saveUi(); renderList(); });
    line.appendChild(tag);
  });
  head.appendChild(line);

  const icons = document.createElement("div");
  icons.className = "snippet-icons";
  icons.appendChild(iconButton("📌", s.pinned ? "Unpin" : "Pin to top", s.pinned, () => toggle(s, "pinned")));
  icons.appendChild(iconButton(s.fav ? "★" : "☆", s.fav ? "Remove from favorites" : "Add to favorites", s.fav, () => toggle(s, "fav")));
  icons.appendChild(iconButton(open ? "▴" : "▾", open ? "Hide preview" : "Show preview", null, () => {
    if (expanded.has(s.id)) expanded.delete(s.id); else expanded.add(s.id);
    renderList();
  }, "preview-toggle"));
  icons.appendChild(iconButton("✎", "Edit", null, () => openEditor(s.id), "edit-btn"));
  const confirming = pendingDelete === s.id;
  icons.appendChild(iconButton(confirming ? "Delete?" : "🗑", confirming ? "Click again to confirm delete" : "Delete", null,
    () => requestDelete(s.id), "del-btn" + (confirming ? " confirm" : "")));
  head.appendChild(icons);
  li.appendChild(head);

  if (open) {
    const detail = document.createElement("div");
    detail.className = "snippet-detail";
    const preview = document.createElement("div");
    preview.className = "snippet-preview";
    preview.textContent = s.text;
    detail.appendChild(preview);

    const tools = document.createElement("div");
    tools.className = "detail-tools";
    [["↑", "Move up", -1], ["↓", "Move down", 1]].forEach((m) => {
      const mb = document.createElement("button");
      mb.type = "button";
      mb.className = "move-btn";
      mb.textContent = m[0];
      mb.title = m[1];
      mb.setAttribute("aria-label", m[1] + " " + (s.title || "snippet"));
      mb.setAttribute("data-move", s.id + ":" + m[2]);
      mb.addEventListener("click", () => {
        if (step(s, m[2])) {
          renderList();
          focusControl('.move-btn[data-move="' + s.id + ":" + m[2] + '"]');
        }
      });
      tools.appendChild(mb);
    });
    detail.appendChild(tools);
    li.appendChild(detail);
  }
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
      // The trailing line break ends the paragraph, so the cursor lands on the NEXT line.
      // A second click then inserts below instead of on the same line.
      const text = s.text.replace(/\n+$/, "") + "\n";
      const range = context.document.getSelection().insertText(text, Word.InsertLocation.replace);
      range.select(Word.SelectionMode.end); // cursor goes to the start of the new line
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
