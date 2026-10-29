"use strict";

const $ = id => document.getElementById(id);
const state = {
  document: null, proposal: null, busy: false, connected: false,
  aiProblem: "Checking configuration", graphProblem: "Checking configuration",
  folders: [{ id: null, name: "OneDrive" }], next: null,
  selected: new Set(),
  chat: [],
};

function showNotice(message, kind = "error") {
  $("notice-text").textContent = message;
  $("notice").className = `notice ${kind}`;
}

async function api(path, body, timeout = 90000) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeout);
  try {
    const options = { signal: controller.signal, credentials: "same-origin", headers: {} };
    if (body !== undefined) {
      options.method = "POST";
      options.headers["X-Demo-Request"] = "1";
      if (body instanceof FormData) {
        options.body = body;
      } else {
        options.headers["Content-Type"] = "application/json";
        options.body = JSON.stringify(body);
      }
    }
    const response = await fetch(path, options);
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || `Request failed (HTTP ${response.status}).`);
    return data;
  } catch (error) {
    if (error.name === "AbortError") {
      throw new Error("The request timed out; its result is not confirmed. Use Refresh state, and check OneDrive before retrying a save.");
    }
    if (error instanceof TypeError) throw new Error("Could not reach the local app. Check the Python server and refresh.");
    throw error;
  } finally {
    clearTimeout(timer);
  }
}

function syncControls() {
  document.querySelectorAll("button, input, textarea").forEach(element => { element.disabled = state.busy; });
  const loaded = Boolean(state.document);
  $("login").disabled = state.busy || Boolean(state.graphProblem);
  $("generate").disabled = state.busy || Boolean(state.aiProblem);
  $("apply").disabled = state.busy || !state.proposal || state.proposal.changes.length === 0;
  $("apply-copy").disabled = state.busy || !state.proposal || !state.document?.source || !state.connected;
  $("apply-copy").classList.toggle("hidden", !state.document?.source);
  $("apply").textContent = state.document?.source ? "Apply & save" : "Apply locally";
  $("apply-help").textContent = state.document?.source
    ? "Apply saves to the same OneDrive file automatically. Apply as new file saves a separate copy without updating the original."
    : "This is a local document. Apply updates the working copy; download it to keep your changes.";
  $("discard").disabled = state.busy || !state.proposal;
  document.querySelectorAll(".block-select").forEach(element => { element.disabled = state.busy; });
  $("clear-selection").disabled = state.busy || !state.selected.size;
  $("save").disabled = state.busy || !loaded || !state.document.source || !state.connected;
  $("update").disabled = state.busy || !loaded || !state.document.source || !state.connected || !state.document.modified;
  $("update").classList.toggle("hidden", !state.document?.source || !state.document.modified);
  $("create-report").disabled = state.busy || !loaded || !state.connected || !state.document.research_draft || Boolean(state.document.source) || Boolean(state.proposal);
  $("create-report").classList.toggle("hidden", !state.document?.research_draft || Boolean(state.document?.source));
  $("allow-research").disabled = state.busy || !state.connected;
  if (!state.connected) $("allow-research").checked = false;
  $("research-help").textContent = state.connected
    ? "Research creates a new report in OneDrive automatically. Existing files are never overwritten by research."
    : "Connect Microsoft 365 before researching. You can still chat about a loaded document.";
  $("drive-refresh").disabled = state.busy || !state.connected;
  $("reset").disabled = state.busy || !loaded;
  $("folder-up").disabled = state.busy || state.folders.length < 2;
  $("download").classList.toggle("disabled", state.busy || !loaded);
  $("download").setAttribute("aria-disabled", String(state.busy || !loaded));
  if (loaded) $("download").href = "/api/document/download";
  else $("download").removeAttribute("href");
  $("activity").classList.toggle("working", state.busy);
}

async function task(label, action) {
  if (state.busy) return;
  state.busy = true;
  $("activity").textContent = label;
  $("notice").classList.add("hidden");
  syncControls();
  try {
    await action();
  } catch (error) {
    showNotice(error.message || "The operation failed. Refresh the app and retry.");
  } finally {
    state.busy = false;
    $("activity").textContent = "Ready";
    syncControls();
  }
}

function safeLink(element, value) {
  element.classList.add("hidden");
  element.removeAttribute("href");
  if (!value) return;
  try {
    const url = new URL(value);
    if (url.protocol !== "https:" || url.username || url.password) return;
    element.href = url.href;
    element.classList.remove("hidden");
  } catch {
    showNotice("The saved document link is invalid. Locate the file directly in OneDrive.");
  }
}

function applyDocumentFormat(element, format = {}) {
  // These are validated Word document values, not application theme colors or model HTML.
  if (/^[0-9a-f]{6}$/i.test(format.color || "")) {
    element.style.setProperty("--cp-doc-color", `#${format.color}`);
    element.style.color = "var(--cp-doc-color)";
  }
  if (Number.isFinite(format.font_size_pt) && format.font_size_pt >= 1 && format.font_size_pt <= 400) {
    element.style.fontSize = `${format.font_size_pt}pt`;
  }
  if (typeof format.bold === "boolean") element.style.fontWeight = format.bold ? "700" : "400";
  if (typeof format.italic === "boolean") element.style.fontStyle = format.italic ? "italic" : "normal";
  if (typeof format.font_family === "string" && format.font_family.length <= 80) {
    element.style.fontFamily = `${JSON.stringify(format.font_family)}, serif`;
  }
  if (["left", "center", "right", "justify"].includes(format.alignment)) element.style.textAlign = format.alignment;
}

function updateSelection() {
  $("selection-status").textContent = state.selected.size
    ? `${state.selected.size} block(s) selected. AI will change only these blocks.`
    : "Optional: select blocks to limit the edit. For a table, select consecutive paragraphs or list items.";
  syncControls();
}

function renderNodes(container, nodes, selectable = false) {
  container.replaceChildren();
  function paragraph(block) {
    const item = document.createElement("div");
    item.className = "doc-item";
    if (selectable) {
      const checkbox = document.createElement("input");
      checkbox.type = "checkbox";
      checkbox.className = "block-select";
      checkbox.dataset.blockId = block.id;
      checkbox.setAttribute("aria-label", `Select ${block.id}: ${block.text.slice(0, 70)}`);
      checkbox.classList.toggle("hidden", !block.editable);
      checkbox.addEventListener("change", () => {
        if (checkbox.checked && state.selected.size >= 50) {
          checkbox.checked = false;
          showNotice("Select at most 50 blocks per proposal.");
          return;
        }
        if (checkbox.checked) state.selected.add(block.id);
        else state.selected.delete(block.id);
        item.classList.toggle("selected", checkbox.checked);
        updateSelection();
      });
      item.append(checkbox);
    }
    const p = document.createElement("div");
    p.className = `doc-block ${block.kind}${block.editable ? "" : " readonly"}`;
    p.dataset.blockId = block.id;
    applyDocumentFormat(p, block.format);
    if (!block.editable) {
      const label = document.createElement("span");
      label.className = "block-label";
      label.textContent = `READ-ONLY: ${block.reason}`;
      p.append(label);
    }
    for (const run of block.runs) {
      const span = document.createElement("span");
      span.textContent = run.text;
      applyDocumentFormat(span, run.format);
      p.append(span);
    }
    item.append(p);
    return item;
  }
  for (const node of nodes) {
    if (node.kind === "table") {
      const table = document.createElement("table");
      table.className = "word-table";
      table.setAttribute("aria-label", "Word document table");
      const body = document.createElement("tbody");
      for (const row of node.rows) {
        const tr = document.createElement("tr");
        for (const cell of row) {
          const td = document.createElement("td");
          td.colSpan = cell.colspan;
          for (const block of cell.blocks) td.append(paragraph(block));
          tr.append(td);
        }
        body.append(tr);
      }
      table.append(body);
      container.append(table);
    } else {
      for (const block of node.blocks) container.append(paragraph(block));
    }
  }
}

function renderDocument(documentView) {
  state.document = documentView;
  state.proposal = null;
  state.selected.clear();
  $("proposal").classList.add("hidden");
  $("proposed-content").classList.add("hidden");
  $("proposed-blocks").replaceChildren();
  $("saved-result").classList.add("hidden");
  $("empty-document").classList.toggle("hidden", Boolean(documentView));
  $("document-content").classList.toggle("hidden", !documentView);
  $("document-blocks").replaceChildren();
  $("document-warnings").replaceChildren();
  $("document-name").textContent = documentView ? documentView.name : "Your Word document";
  $("document-status").textContent = !documentView ? "No file loaded" : documentView.modified ? "Unsaved working edits" : "Document loaded";
  updateSelection();
  if (!documentView) return;
  safeLink($("source-link"), documentView.source?.web_url);
  renderNodes($("document-blocks"), documentView.nodes, true);
  for (const warning of documentView.warnings) {
    const p = document.createElement("p");
    p.textContent = warning;
    $("document-warnings").append(p);
  }
}

function changedParts(before, after) {
  const left = Array.from(before);
  const right = Array.from(after);
  let start = 0;
  while (start < left.length && start < right.length && left[start] === right[start]) start++;
  let suffix = 0;
  while (suffix < left.length - start && suffix < right.length - start && left[left.length - suffix - 1] === right[right.length - suffix - 1]) suffix++;
  return {
    prefix: left.slice(0, start).join(""),
    before: left.slice(start, left.length - suffix).join(""),
    after: right.slice(start, right.length - suffix).join(""),
    suffix: suffix ? left.slice(-suffix).join("") : "",
  };
}

function renderProposal(proposal) {
  if (!proposal.changes.length) {
    state.proposal = null;
    $("proposal").classList.add("hidden");
    showNotice(proposal.summary, "info");
    return;
  }
  state.proposal = proposal;
  $("proposal").open = false;
  $("proposal").classList.remove("hidden");
  $("proposal-summary").textContent = proposal.summary;
  $("change-count").textContent = `${proposal.changes.length} change${proposal.changes.length === 1 ? "" : "s"}`;
  $("proposed-content").classList.toggle("hidden", !proposal.preview);
  if (proposal.preview) renderNodes($("proposed-blocks"), proposal.preview.nodes);
  $("changes").replaceChildren();
  for (const change of proposal.changes) {
    const card = document.createElement("article");
    card.className = "change";
    const label = document.createElement("div");
    label.className = "change-label";
    label.textContent = `${change.kind.toUpperCase()} / ${change.block_id}`;
    card.append(label);
    if (change.description) {
      const description = document.createElement("p");
      description.className = "change-text";
      description.textContent = change.description;
      card.append(description);
    }
    if (change.kind === "format") {
      $("changes").append(card);
      continue;
    }
    const parts = changedParts(change.before, change.after);
    for (const [heading, text, tag] of [["BEFORE", parts.before, "del"], ["AFTER", parts.after, "ins"]]) {
      const p = document.createElement("p");
      p.className = "change-text";
      const title = document.createElement("small");
      title.textContent = heading;
      const highlight = document.createElement(tag);
      highlight.textContent = text;
      p.append(title, document.createTextNode(parts.prefix), highlight, document.createTextNode(parts.suffix));
      card.append(p);
    }
    $("changes").append(card);
  }
}

function renderChat(turns, stages = []) {
  state.chat = turns;
  const container = $("chat-history");
  container.replaceChildren();
  if (!turns.length) {
    const welcome = document.createElement("div");
    welcome.className = "chat-welcome";
    const title = document.createElement("strong");
    title.textContent = "What would you like to work on?";
    const description = document.createElement("p");
    description.textContent = "Open a document to refine it, or connect OneDrive and start researching. We can keep working together, one message at a time.";
    welcome.append(title, description);
    container.append(welcome);
  }
  for (const turn of turns) {
    const card = document.createElement("article");
    card.className = `chat-message ${turn.role === "user" ? "user" : "assistant"}`;
    const label = document.createElement("strong");
    label.textContent = turn.role === "user" ? "You" : "Wordsmith";
    if (turn.agent) label.title = turn.agent;
    if (turn.error) card.classList.add("error");
    const text = document.createElement("p");
    text.textContent = turn.text;
    card.append(label, text);
    container.append(card);
  }
  container.scrollTop = container.scrollHeight;
  $("agent-stages").replaceChildren();
  for (const stage of stages) {
    const p = document.createElement("p");
    p.textContent = stage;
    $("agent-stages").append(p);
  }
}

function destination() {
  const folder = state.folders[state.folders.length - 1];
  $("cloud-save-status").classList.remove("saving", "failed");
  $("cloud-save-title").textContent = !state.connected ? "Connect OneDrive to research"
    : state.document?.save_error || state.document?.research_draft ? "OneDrive save not confirmed"
    : state.document?.source && state.document.modified ? "Unsaved changes — retry save or download"
    : state.document?.source ? `In OneDrive: ${state.document.source.name}` : "Research reports save to OneDrive";
  if (state.document?.research_draft || state.document?.save_error) $("cloud-save-status").classList.add("failed");
  $("destination").textContent = state.connected
    ? `New reports save automatically to ${state.folders.map(item => item.name).join(" / ")}.`
    : "Connect Microsoft 365 in the files panel. Research is unavailable until connected.";
  return folder.id;
}

function canReplaceDocument() {
  return !(state.document?.modified || state.proposal) ||
    confirm("Replace the current working document and discard any unapplied proposal? Download or save edits first if you need them.");
}

async function browse(cursor = null) {
  const folder = state.folders[state.folders.length - 1];
  const params = new URLSearchParams();
  if (folder.id) params.set("folder_id", folder.id);
  if (cursor) params.set("cursor", cursor);
  const result = await api(`/api/drive?${params}`);
  state.next = result.next;
  $("folder-label").textContent = state.folders.map(item => item.name).join(" / ");
  destination();
  $("drive-items").replaceChildren();
  if (!result.entries.length) {
    const p = document.createElement("p");
    p.className = "hint";
    p.textContent = result.next ? "No supported files on this page. Try the next page." : "No folders or .docx files in this folder.";
    $("drive-items").append(p);
  }
  for (const entry of result.entries) {
    const button = document.createElement("button");
    button.className = "drive-entry";
    button.append(document.createTextNode(entry.name));
    const label = document.createElement("span");
    label.textContent = entry.kind === "folder" ? "FOLDER" : `WORD DOCUMENT / ${Math.ceil(entry.size / 1024)} KiB`;
    button.append(label);
    button.addEventListener("click", () => task(entry.kind === "folder" ? "Opening folder..." : "Downloading Word document...", async () => {
      if (entry.kind === "folder") {
        state.folders.push({ id: entry.id, name: entry.name });
        try { await browse(); } catch (error) { state.folders.pop(); throw error; }
      } else if (canReplaceDocument()) {
        renderDocument(await api("/api/document/onedrive", { item_id: entry.id }));
        renderChat([]);
      }
    }));
    $("drive-items").append(button);
  }
  $("drive-next").classList.toggle("hidden", !state.next);
}

async function refresh() {
  const result = await api("/api/status");
  state.aiProblem = result.ai_problem;
  state.graphProblem = result.graph_problem;
  state.connected = result.connected;
  $("ai-status").textContent = state.aiProblem ? "Setup needed" : "Configured";
  $("ai-status").classList.toggle("ready", !state.aiProblem);
  $("ai-help").textContent = state.aiProblem || `Azure authentication: ${result.ai_auth}. A real model call happens only when requested.`;
  $("cloud-status").textContent = state.connected ? "Connected" : "Not connected";
  $("cloud-status").classList.toggle("ready", state.connected);
  $("login").textContent = state.connected ? "Sign in again" : "Connect Microsoft 365";
  $("login-help").textContent = state.graphProblem || (state.connected
    ? "Connected. Open Test below and select a Word document."
    : "Connect authorizes this demo to access your OneDrive. Complete sign-in in the separate browser window.");
  $("logout").classList.toggle("hidden", !state.connected);
  $("drive-browser").classList.toggle("hidden", !state.connected);
  $("setup-items").replaceChildren();
  const problems = [state.graphProblem, state.aiProblem].filter(Boolean);
  for (const problem of problems) {
    const item = document.createElement("li");
    item.textContent = problem;
    $("setup-items").append(item);
  }
  $("setup").classList.toggle("hidden", problems.length === 0);
  renderDocument(result.document);
  renderChat(result.chat || []);
  if (result.proposal) renderProposal(result.proposal);
  destination();
  if (state.connected) await browse();
}

$("theme").addEventListener("click", () => {
  const next = document.documentElement.dataset.theme === "dark" ? "light" : "dark";
  document.documentElement.dataset.theme = next;
  const url = new URL(location.href);
  url.searchParams.set("clawpilotTheme", next);
  history.replaceState(null, "", url);
});
$("dismiss").addEventListener("click", () => $("notice").classList.add("hidden"));
$("refresh").addEventListener("click", () => task("Refreshing state...", refresh));
$("sample").addEventListener("click", () => task("Opening synthetic Word sample...", async () => {
  if (canReplaceDocument()) {
    renderDocument(await api("/api/document/sample", {}));
    renderChat([]);
  }
}));
$("upload").addEventListener("change", () => task("Opening Word document...", async () => {
  const file = $("upload").files[0];
  try {
    if (!file || !canReplaceDocument()) return;
    if (!file.name.toLowerCase().endsWith(".docx")) throw new Error("Select a .docx file.");
    if (file.size > 10 * 1024 * 1024) throw new Error("The demo accepts files up to 10 MiB.");
    const form = new FormData();
    form.append("file", file);
    renderDocument(await api("/api/document/upload", form));
    renderChat([]);
  } finally { $("upload").value = ""; }
}));
$("login").addEventListener("click", () => task("Complete Microsoft sign-in in the browser window...", async () => {
  $("login-pending").classList.remove("hidden");
  $("login-link").classList.add("hidden");
  let stopped = false;
  let polling;
  const poll = async () => {
    try {
      const result = await api("/api/auth/pending");
      if (!stopped && result.url) {
        const url = new URL(result.url);
        if (url.protocol !== "http:" || url.hostname !== "localhost" || url.username || url.password) {
          throw new Error("The server returned an invalid sign-in link.");
        }
        $("login-link").href = url.href;
        $("login-link").classList.remove("hidden");
      }
    } catch (error) {
      if (!stopped) showNotice(`Could not check sign-in progress: ${error.message}`);
    } finally {
      if (!stopped) polling = setTimeout(poll, 750);
    }
  };
  polling = setTimeout(poll, 500);
  try {
    await api("/api/auth/login", {}, 210000);
  } finally {
    stopped = true;
    clearTimeout(polling);
    $("login-pending").classList.add("hidden");
    $("login-link").removeAttribute("href");
  }
  state.folders = [{ id: null, name: "OneDrive" }];
  await refresh();
}));
$("logout").addEventListener("click", () => task("Disconnecting Microsoft 365...", async () => {
  if (!canReplaceDocument()) return;
  await api("/api/auth/logout", {});
  await refresh();
}));
$("folder-up").addEventListener("click", () => task("Opening parent folder...", async () => {
  const removed = state.folders.pop();
  try { await browse(); } catch (error) { state.folders.push(removed); throw error; }
}));
$("drive-next").addEventListener("click", () => task("Loading the next OneDrive page...", () => browse(state.next)));
$("drive-refresh").addEventListener("click", () => task("Refreshing this OneDrive folder...", () => browse()));
$("instruction").addEventListener("keydown", event => {
  if (event.key === "Enter" && !event.shiftKey && !event.isComposing) {
    event.preventDefault();
    if (!$("generate").disabled) $("chat-form").requestSubmit();
  }
});
$("chat-form").addEventListener("submit", event => {
  event.preventDefault();
  task("Thinking...", async () => {
  const instruction = $("instruction").value.trim();
  if (!instruction) throw new Error("Enter a research question, a follow-up, or an editing instruction.");
  const selectedIds = (state.document?.blocks || []).filter(block => state.selected.has(block.id)).map(block => block.id);
  const body = {
    message: instruction, request_id: crypto.randomUUID(), version: state.document?.version || null,
    selected_block_ids: selectedIds, allow_research: $("allow-research").checked,
    folder_id: destination(),
  };
  const previous = state.chat;
  renderChat([...previous, {role: "user", text: instruction}]);
  $("chat-history").setAttribute("aria-busy", "true");
  $("chat-progress").textContent = "Thinking...";
  $("chat-pending").classList.remove("hidden");
  let stopped = false;
  let polling;
  const poll = async () => {
    try {
      const progress = await api("/api/chat/activity");
      if (!stopped && progress.activity !== "Ready") {
        $("activity").textContent = progress.activity;
        $("chat-progress").textContent = progress.activity;
        if (progress.activity.startsWith("Saving")) {
          $("cloud-save-title").textContent = "Saving your report to OneDrive...";
          $("cloud-save-status").classList.add("saving");
        }
      }
    } catch (error) {
      if (!stopped) showNotice(`Could not read agent activity: ${error.message}`);
    } finally {
      if (!stopped) polling = setTimeout(poll, 1200);
    }
  };
  polling = setTimeout(poll, 500);
  try {
    const result = await api("/api/chat", body, 240000);
    renderDocument(result.document);
    renderChat(result.chat, result.stages);
    if (result.proposal) renderProposal(result.proposal);
    $("instruction").value = "";
    destination();
    if (result.save_error) showNotice(result.save_error);
    if (result.research_saved) {
      showNotice("The research report was created in OneDrive. Continue chatting to propose edits.", "success");
      await browse();
    }
  } catch (error) {
    renderChat([...previous, {role: "user", text: instruction}, {role: "assistant", text: error.message, error: true}]);
    destination();
    throw error;
  } finally {
    stopped = true;
    clearTimeout(polling);
    $("chat-history").setAttribute("aria-busy", "false");
    $("chat-pending").classList.add("hidden");
    $("instruction").focus();
  }
  });
});
$("clear-chat").addEventListener("click", () => task("Starting a new conversation...", async () => {
  await api("/api/chat/clear", {});
  $("allow-research").checked = false;
  renderChat([]);
  showNotice("New conversation started. Your current Word document is unchanged.", "success");
}));
$("create-report").addEventListener("click", () => task("Creating a new report in OneDrive...", async () => {
  renderDocument(await api("/api/document/create", {version: state.document.version, folder_id: destination()}));
  destination();
  showNotice("Created a new report in OneDrive. Future edits can be saved back to this file.", "success");
  await browse();
}));
function applyProposal(saveAsNew) {
  return task(saveAsNew ? "Applying and saving as a new OneDrive file..." : "Applying and saving reviewed changes...", async () => {
  $("cloud-save-title").textContent = state.document.source ? "Saving reviewed changes to OneDrive..." : "Applying local changes...";
  $("cloud-save-status").classList.add("saving");
  const result = await api("/api/proposal/apply", {
    version: state.document.version, proposal_id: state.proposal.id, save_as_new: saveAsNew,
  });
  renderDocument(result);
  destination();
  showNotice(result.apply_message, result.save_error ? "error" : "success");
  const status = await api("/api/status");
  renderChat(status.chat);
  });
}
$("apply").addEventListener("click", () => applyProposal(false));
$("apply-copy").addEventListener("click", () => applyProposal(true));
$("discard").addEventListener("click", () => task("Discarding proposal...", async () => {
  await api("/api/proposal/discard", {});
  state.proposal = null;
  $("proposal").classList.add("hidden");
  $("proposed-content").classList.add("hidden");
  $("proposed-blocks").replaceChildren();
  const status = await api("/api/status");
  renderChat(status.chat);
}));
$("clear-selection").addEventListener("click", () => {
  state.selected.clear();
  document.querySelectorAll(".block-select").forEach(checkbox => { checkbox.checked = false; });
  document.querySelectorAll(".doc-item.selected").forEach(item => item.classList.remove("selected"));
  updateSelection();
});
$("reset").addEventListener("click", () => task("Clearing the working document...", async () => {
  if (!canReplaceDocument()) return;
  await api("/api/document/reset", {});
  renderDocument(null);
  renderChat([]);
}));
$("download").addEventListener("click", event => {
  if (state.busy || !state.document) event.preventDefault();
});
$("save").addEventListener("click", () => task("Saving a new OneDrive copy. The original is not targeted...", async () => {
  if (!confirm("Create a new edited .docx in the source OneDrive folder? The original will not be overwritten.")) return;
  const saved = await api("/api/document/save", { version: state.document.version });
  $("saved-name").textContent = `Saved: ${saved.name}`;
  $("saved-link").textContent = "Open saved copy in Word";
  safeLink($("saved-link"), saved.web_url);
  $("saved-result").classList.remove("hidden");
  showNotice("A separate Word copy was saved to OneDrive.", "success");
}));
$("update").addEventListener("click", () => task("Saving reviewed changes back to the opened OneDrive document...", async () => {
  if (!confirm(`Update "${state.document.source.name}" in OneDrive? This replaces that file's contents. Newer cloud edits will block the save.`)) return;
  const saved = await api("/api/document/update", { version: state.document.version });
  renderDocument(saved.document);
  destination();
  $("saved-name").textContent = `Updated in OneDrive: ${saved.name}`;
  safeLink($("saved-link"), saved.web_url);
  $("saved-link").textContent = "Open saved document in Word";
  $("saved-result").classList.remove("hidden");
  showNotice("Your edits were saved back to the same OneDrive document.", "success");
}));
task("Checking local configuration...", refresh);
