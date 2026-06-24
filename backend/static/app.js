"use strict";

// --------------------------------------------------------------------------
// Small helpers
// --------------------------------------------------------------------------
const $ = (sel) => document.querySelector(sel);
const $$ = (sel) => Array.from(document.querySelectorAll(sel));

// Editable project columns (admin grid), in display order.
const PROJECT_FIELDS = [
  { key: "project_number", label: "Project #", type: "text" },
  { key: "name", label: "Name", type: "text", required: true },
  { key: "orig_substantial_completion", label: "Orig. Substantial", type: "date" },
  { key: "orig_final_completion", label: "Orig. Final", type: "date" },
  { key: "current_substantial_completion", label: "Current Substantial", type: "date" },
  { key: "current_final_completion", label: "Current Final", type: "date" },
  { key: "contract_amount_last_pay_app", label: "Contract $ (last pay app)", type: "number" },
  { key: "month_last_pay_app", label: "Month (last pay app)", type: "date" },
];

let projects = [];
let adminPassword = sessionStorage.getItem("adminPassword") || null;

function setStatus(el, message, kind) {
  el.textContent = message || "";
  el.className = "status center" + (kind ? " " + kind : "");
}

async function api(path, { method = "GET", body, admin = false } = {}) {
  const headers = {};
  if (body !== undefined) headers["Content-Type"] = "application/json";
  if (admin && adminPassword) headers["X-Admin-Password"] = adminPassword;
  const res = await fetch(path, {
    method,
    headers,
    body: body !== undefined ? JSON.stringify(body) : undefined,
  });
  if (!res.ok) {
    let detail = res.statusText;
    try {
      detail = (await res.json()).detail || detail;
    } catch (_) {}
    throw new Error(detail);
  }
  return res.status === 204 ? null : res.json();
}

// --------------------------------------------------------------------------
// Tabs
// --------------------------------------------------------------------------
$$(".tab").forEach((tab) => {
  tab.addEventListener("click", () => {
    $$(".tab").forEach((t) => t.classList.remove("active"));
    $$(".panel").forEach((p) => p.classList.remove("active"));
    tab.classList.add("active");
    $("#" + tab.dataset.tab).classList.add("active");
    if (tab.dataset.tab === "admin") refreshAdminView();
  });
});

// --------------------------------------------------------------------------
// Generate tab — pick a project, choose a CSV, download immediately.
// --------------------------------------------------------------------------
async function loadProjects() {
  try {
    projects = await api("/api/projects");
  } catch (e) {
    projects = [];
  }
  const sel = $("#project-select");
  const current = sel.value;
  sel.innerHTML =
    '<option value="">— Manual entry (no saved project) —</option>' +
    projects
      .map((p) => {
        const label = p.project_number
          ? `${p.project_number} — ${p.name}`
          : p.name;
        return `<option value="${p.id}">${escapeHtml(label)}</option>`;
      })
      .join("");
  if (current) sel.value = current;
}

const csvInput = $("#csv-input");
$("#upload-btn").addEventListener("click", () => csvInput.click());
csvInput.addEventListener("change", () => {
  if (csvInput.files.length) generateAndDownload(csvInput.files[0]);
});

async function generateAndDownload(file) {
  const status = $("#generate-status");
  const btn = $("#upload-btn");
  setStatus(status, "Generating…", "busy");
  btn.disabled = true;

  const fd = new FormData();
  fd.append("csv_file", file);
  const pid = $("#project-select").value;
  if (pid) fd.append("project_id", pid);

  try {
    const res = await fetch("/api/generate", { method: "POST", body: fd });
    if (!res.ok) {
      let detail = res.statusText;
      try {
        detail = (await res.json()).detail || detail;
      } catch (_) {}
      throw new Error(detail);
    }
    const blob = await res.blob();
    const filename = filenameFromDisposition(
      res.headers.get("Content-Disposition")
    );
    downloadBlob(blob, filename);
    setStatus(status, "Downloaded " + filename, "ok");
  } catch (e) {
    setStatus(status, "Error: " + e.message, "err");
  } finally {
    btn.disabled = false;
    csvInput.value = ""; // allow re-selecting the same file
  }
}

function filenameFromDisposition(disposition) {
  if (!disposition) return "Job Cost Projection.xlsx";
  const m = /filename="?([^"]+)"?/.exec(disposition);
  return m ? m[1] : "Job Cost Projection.xlsx";
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

// --------------------------------------------------------------------------
// Admin tab
// --------------------------------------------------------------------------
function refreshAdminView() {
  if (adminPassword) {
    $("#admin-login").classList.add("hidden");
    $("#admin-grid").classList.remove("hidden");
    renderGrid();
  } else {
    $("#admin-login").classList.remove("hidden");
    $("#admin-grid").classList.add("hidden");
  }
}

$("#admin-login-btn").addEventListener("click", async () => {
  const pwd = $("#admin-password").value;
  const status = $("#admin-login-status");
  if (!pwd) {
    setStatus(status, "Enter the admin password.", "err");
    return;
  }
  setStatus(status, "Signing in…", "busy");
  try {
    await api("/api/admin/login", { method: "POST", body: { password: pwd } });
    adminPassword = pwd;
    sessionStorage.setItem("adminPassword", pwd);
    setStatus(status, "", "");
    $("#admin-password").value = "";
    refreshAdminView();
  } catch (e) {
    setStatus(status, e.message, "err");
  }
});

$("#admin-signout-btn").addEventListener("click", () => {
  adminPassword = null;
  sessionStorage.removeItem("adminPassword");
  refreshAdminView();
});

$("#add-row-btn").addEventListener("click", () => {
  renderGrid({ id: null, name: "", _new: true });
});

function renderGridHead() {
  const head = $("#projects-head");
  head.innerHTML =
    PROJECT_FIELDS.map((f) => `<th>${escapeHtml(f.label)}</th>`).join("") +
    '<th class="col-actions">Actions</th>';
}

async function renderGrid(draftRow) {
  const status = $("#admin-grid-status");
  try {
    projects = await api("/api/projects");
  } catch (e) {
    setStatus(status, "Could not load projects: " + e.message, "err");
    return;
  }
  renderGridHead();
  const body = $("#projects-body");
  body.innerHTML = "";

  const rows = projects.slice();
  if (draftRow) rows.push(draftRow);

  if (rows.length === 0) {
    body.innerHTML = `<tr class="empty-row"><td colspan="${
      PROJECT_FIELDS.length + 1
    }">No projects yet. Click “+ Add project”.</td></tr>`;
    return;
  }

  rows.forEach((p) => body.appendChild(buildRow(p)));
}

function buildRow(p) {
  const tr = document.createElement("tr");
  tr.dataset.id = p.id == null ? "" : p.id;

  PROJECT_FIELDS.forEach((f) => {
    const td = document.createElement("td");
    const input = document.createElement("input");
    input.type = f.type;
    if (f.type === "number") input.step = "0.01";
    input.value = p[f.key] != null ? p[f.key] : "";
    input.dataset.field = f.key;
    if (f.required) input.placeholder = "Required";
    input.addEventListener("input", () => {
      td.classList.add("dirty");
      tr.querySelector(".btn-save").disabled = false;
    });
    td.appendChild(input);
    tr.appendChild(td);
  });

  const actions = document.createElement("td");
  actions.className = "row-actions";
  const save = document.createElement("button");
  save.className = "btn-save";
  save.textContent = p._new ? "Create" : "Save";
  save.disabled = !p._new;
  save.addEventListener("click", () => saveRow(tr, p));
  const del = document.createElement("button");
  del.className = "btn-del";
  del.textContent = "Delete";
  del.addEventListener("click", () => deleteRow(tr, p));
  actions.appendChild(save);
  actions.appendChild(del);
  tr.appendChild(actions);
  return tr;
}

function collectRow(tr) {
  const data = {};
  tr.querySelectorAll("input[data-field]").forEach((i) => {
    data[i.dataset.field] = i.value || null;
  });
  return data;
}

async function saveRow(tr, p) {
  const status = $("#admin-grid-status");
  const data = collectRow(tr);
  if (!data.name) {
    setStatus(status, "Project name is required.", "err");
    return;
  }
  setStatus(status, "Saving…", "busy");
  try {
    if (p.id == null) {
      await api("/api/projects", { method: "POST", body: data, admin: true });
    } else {
      await api("/api/projects/" + p.id, {
        method: "PUT",
        body: data,
        admin: true,
      });
    }
    setStatus(status, "Saved “" + data.name + "”.", "ok");
    await renderGrid();
    await loadProjects();
  } catch (e) {
    handleAdminError(status, e);
  }
}

async function deleteRow(tr, p) {
  const status = $("#admin-grid-status");
  if (p.id == null) {
    renderGrid(); // discard unsaved draft
    return;
  }
  if (!confirm('Delete project "' + (p.name || "") + '"?')) return;
  setStatus(status, "Deleting…", "busy");
  try {
    await api("/api/projects/" + p.id, { method: "DELETE", admin: true });
    setStatus(status, "Deleted.", "ok");
    await renderGrid();
    await loadProjects();
  } catch (e) {
    handleAdminError(status, e);
  }
}

function handleAdminError(status, e) {
  if (/auth/i.test(e.message) || /401/.test(e.message)) {
    adminPassword = null;
    sessionStorage.removeItem("adminPassword");
    refreshAdminView();
    setStatus($("#admin-login-status"), "Session expired, sign in again.", "err");
  } else {
    setStatus(status, "Error: " + e.message, "err");
  }
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, (c) => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#39;",
  }[c]));
}

// --------------------------------------------------------------------------
// Init
// --------------------------------------------------------------------------
loadProjects();
