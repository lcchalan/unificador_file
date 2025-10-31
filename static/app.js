// static/app.js
const $ = (sel) => document.querySelector(sel);
const log = (m) => {
  const el = $("#log");
  el.textContent += (m + "\n");
  el.scrollTop = el.scrollHeight;
};

let uploadToken = null;   // para modo "subir archivos"
let lastSource = "folder"; // folder | upload
let lastFilesMeta = [];   // [{name, count, titles}]

function setBadges(filesCount, titlesCount) {
  $("#badge-archivos").textContent = `${filesCount} archivo${filesCount===1?"":"s"}`;
  $("#badge-titulos").textContent = `${titlesCount} título${titlesCount===1?"":"s"}`;
}

function fillTitles(titles) {
  const sel = $("#titles");
  sel.innerHTML = "";
  titles.forEach(t => {
    const o = document.createElement("option");
    o.value = t; o.textContent = t;
    sel.appendChild(o);
  });
  setBadges(lastFilesMeta.length, titles.length);
}

function drawFilesTable(filesMeta) {
  const tbody = $("#files-table");
  tbody.innerHTML = "";
  if (!filesMeta.length) {
    tbody.innerHTML = `<tr><td colspan="3" class="text-muted small p-3">Sin datos. Escanea primero.</td></tr>`;
    return;
  }
  for (const f of filesMeta) {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td class="small">${f.name}</td>
      <td class="text-center"><span class="badge ${f.count>0?"bg-primary":"bg-secondary"}">${f.count}</span></td>
      <td class="text-end">
        <button class="btn btn-sm btn-outline-secondary" data-action="preview" data-name="${encodeURIComponent(f.name)}">Ver títulos</button>
      </td>
    `;
    tbody.appendChild(tr);
  }
}

function showPreview(name) {
  const decoded = decodeURIComponent(name);
  const meta = lastFilesMeta.find(x => x.name === decoded);
  $("#preview-file").textContent = decoded;
  const ul = $("#preview-list");
  ul.innerHTML = "";
  if (!meta || !meta.titles || !meta.titles.length) {
    ul.innerHTML = `<li class="list-group-item text-muted">No se encontraron títulos para el nivel seleccionado.</li>`;
  } else {
    meta.titles.forEach(t => {
      const li = document.createElement("li");
      li.className = "list-group-item small";
      li.textContent = t;
      ul.appendChild(li);
    });
  }
  const modal = new bootstrap.Modal(document.getElementById('modalPreview'));
  modal.show();
}

async function scanFolder() {
  lastSource = "folder";
  uploadToken = null;
  $("#titles").innerHTML = "";
  drawFilesTable([]);
  setBadges(0, 0);
  log("Escaneando carpeta...");

  const body = {
    folder: $("#folder").value.trim(),
    include_subs: $("#include_subs").checked,
    level: parseInt($("#level").value, 10)
  };
  const res = await fetch("/api/scan-folder", {
    method: "POST", headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body)
  });
  const js = await res.json();
  if (!js.ok) { log("ERROR: " + (js.error || res.status)); return; }
  lastFilesMeta = js.files_meta || [];
  log(`Archivos: ${js.count} | Títulos (únicos, nivel) : ${js.titles.length}`);
  fillTitles(js.titles || []);
  drawFilesTable(lastFilesMeta);
}

async function scanUpload() {
  lastSource = "upload";
  $("#titles").innerHTML = "";
  drawFilesTable([]);
  setBadges(0, 0);
  const files = $("#files").files;
  if (!files || !files.length) { log("Selecciona .docx primero."); return; }
  log(`Subiendo ${files.length} archivo(s)...`);
  const fd = new FormData();
  fd.append("level", $("#level").value);
  for (const f of files) fd.append("files", f);

  const res = await fetch("/api/scan-upload", { method: "POST", body: fd });
  const js = await res.json();
  if (!js.ok) { log("ERROR: " + (js.error || res.status)); return; }
  uploadToken = js.token;
  lastFilesMeta = js.files_meta || [];
  log(`Token: ${uploadToken} | Archivos: ${js.count} | Títulos (únicos): ${js.titles.length}`);
  fillTitles(js.titles || []);
  drawFilesTable(lastFilesMeta);
}

async function mergeNow() {
  const mode = document.querySelector('input[name="mode"]:checked').value; // unificado | grouped
  const use_all = $("#use_all").checked;
  const level = parseInt($("#level").value, 10);
  let selected = [];
  if (!use_all) {
    selected = Array.from($("#titles").selectedOptions).map(o => o.value);
    if (!selected.length) { log("Selecciona al menos un título o marca 'Usar todos'."); return; }
  }

  const payload = {
    source: lastSource, mode, level, use_all, titles: selected
  };

  if (lastSource === "folder") {
    const folder = $("#folder").value.trim();
    if (!folder) { log("Ingresa la ruta de la carpeta."); return; }
    payload.folder = folder;
    payload.include_subs = $("#include_subs").checked;
  } else {
    if (!uploadToken) { log("Primero sube y escanea archivos."); return; }
    payload.token = uploadToken;
  }

  log("Generando ZIP...");
  const res = await fetch("/api/merge", {
    method: "POST", headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload)
  });
  if (!res.ok) {
    const js = await res.json().catch(()=>({ok:false,error:res.statusText}));
    log("ERROR: " + (js.error || res.status));
    return;
  }
  const blob = await res.blob();
  const a = document.createElement("a");
  const url = URL.createObjectURL(blob);
  a.href = url;
  a.download = (mode === "unificado" ? "unificado.zip" : "por_titulo.zip");
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
  log("Descarga iniciada.");
}

async function cleanup() {
  if (!uploadToken) { log("No hay subidas por limpiar."); return; }
  await fetch("/api/cleanup", {
    method: "POST", headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ token: uploadToken })
  });
  uploadToken = null;
  log("Subidas temporales limpiadas.");
}

document.addEventListener("DOMContentLoaded", () => {
  $("#btn-scan-folder").addEventListener("click", scanFolder);
  $("#btn-scan-upload").addEventListener("click", scanUpload);
  $("#btn-merge").addEventListener("click", mergeNow);
  $("#btn-clean").addEventListener("click", cleanup);

  $("#level").addEventListener("change", () => {
    $("#titles").innerHTML = "";
    drawFilesTable([]);
    setBadges(0, 0);
    log("Nivel cambiado: vuelve a Escanear títulos.");
  });

  $("#use_all").addEventListener("change", (e) => {
    if (e.target.checked) {
      for (const opt of $("#titles").options) opt.selected = false;
    }
  });

  // Delegación para botón "Ver títulos"
  $("#files-table").addEventListener("click", (ev) => {
    const btn = ev.target.closest("button[data-action='preview']");
    if (!btn) return;
    showPreview(btn.dataset.name);
  });
});
