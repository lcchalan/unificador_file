// static/app.js
const $ = (sel) => document.querySelector(sel);
const log = (m) => { const el = $("#log"); el.textContent += (m + "\n"); el.scrollTop = el.scrollHeight; };

let uploadToken = null;    // para "subir .docx"
let uploadDirToken = null; // para "subir carpeta"
let lastSource = "folder"; // folder | upload | upload_dir
let lastFilesMeta = [];    // [{name, count, titles}]

// === WHITELIST (tu lista fija) ===
const TITULOS_WHITELIST = [
  "1. Introducción​",
  "2. Diagnóstico estratégico",
  "3. Misión, visión y valores de la carrera",
  "01. Plan de formación integral del estudiante",
  "06. Plan de admisión, acogida y acompañamiento académico de estudiantes",
  "23. Plan de seguimiento y mejora de indicadores del perfil docente",
  "25. Plan de formación integral del docente",
  "26. Plan de mejora del proceso de evaluación integral docente ",
  "03. Plan implantación del marco de competencias UTPL",
  "04. Plan de prospectiva y creación de nueva oferta",
  "07. Plan de acciones curriculares para el fortalecimiento de las competencias genéricas",
  "11. Plan de fortalecimiento de prácticas preprofesionales y proyectos de vinculación",
  "12. Plan de fortalecimiento de criterios para la evaluación de la calidad de carreras y programas académicos",
  "13. Plan de acciones curriculares para el fortalecimiento de la empleabilidad del graduado UTPL",
  "16. Plan de mejora del proceso de elaboración y seguimiento de planes docentes",
  "18. Plan de mejora de ambientes de aprendizaje",
  "19. Plan de mejora de evaluación de los aprendizajes",
  "20. Plan de mejora del proceso de integración curricular",
  "21. Plan de mejora del proceso de titulación",
  "22. Plan de seguimiento y mejora de la labor tutorial",
  "08. Plan de internacionalización del currículo",
  "24. Plan de intervención de personal académico en territorio",
  "05. Plan de acciones académicas orientadas a la comunicación y promoción de la oferta",
  "09. Plan de innovación educativa",
  "10. Plan de implantación de metodologías activas en el currículo",
  "28. Plan de formación de líderes académicos ",
  "29. Plan de posicionamiento institucional en innovación educativa",
  "30. Plan de investigación sobre innovación educativa, EaD, MP"
];

function setBadges(filesCount) {
  $("#badge-archivos").textContent = `${filesCount} archivo${filesCount===1?"":"s"}`;
}

function fillWhitelist() {
  const sel = $("#whitelist");
  sel.innerHTML = "";
  TITULOS_WHITELIST.forEach(t => {
    const o = document.createElement("option");
    o.value = t; o.textContent = t;
    sel.appendChild(o);
  });
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
    ul.innerHTML = `<li class="list-group-item text-muted">Sin coincidencias con la whitelist.</li>`;
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
  lastSource = "folder"; uploadToken = null; uploadDirToken = null;
  $("#whitelist").selectedIndex = -1;
  drawFilesTable([]); setBadges(0);
  log("Escaneando carpeta en servidor...");

  const body = {
    folder: $("#folder").value.trim(),
    include_subs: $("#include_subs").checked,
    whitelist: TITULOS_WHITELIST
  };
  const res = await fetch("/api/scan-folder", {
    method: "POST", headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body)
  });
  const js = await res.json();
  if (!js.ok) { log("ERROR: " + (js.error || res.status)); return; }
  lastFilesMeta = js.files_meta || [];
  setBadges(js.count || 0);
  drawFilesTable(lastFilesMeta);
  log(`Archivos: ${js.count} | títulos totales en whitelist: ${TITULOS_WHITELIST.length}`);
}

async function scanUpload() {
  lastSource = "upload"; uploadDirToken = null;
  $("#whitelist").selectedIndex = -1;
  drawFilesTable([]); setBadges(0);
  const files = $("#files").files;
  if (!files || !files.length) { log("Selecciona .docx primero."); return; }
  log(`Subiendo ${files.length} archivo(s)...`);
  const fd = new FormData();
  TITULOS_WHITELIST.forEach(t => fd.append("whitelist[]", t));
  for (const f of files) fd.append("files", f);

  const res = await fetch("/api/scan-upload", { method: "POST", body: fd });
  const js = await res.json();
  if (!js.ok) { log("ERROR: " + (js.error || res.status)); return; }
  uploadToken = js.token;
  lastFilesMeta = js.files_meta || [];
  setBadges(js.count || 0);
  drawFilesTable(lastFilesMeta);
  log(`Token: ${uploadToken} | Archivos: ${js.count}`);
}

async function scanUploadDir() {
  lastSource = "upload_dir"; uploadToken = null;
  $("#whitelist").selectedIndex = -1;
  drawFilesTable([]); setBadges(0);
  const files = $("#folderInput").files;
  if (!files || !files.length) { log("Selecciona una carpeta con .docx."); return; }
  log(`Subiendo carpeta con ${files.length} archivo(s)...`);
  const fd = new FormData();
  TITULOS_WHITELIST.forEach(t => fd.append("whitelist[]", t));
  for (const f of files) {
    if (f.name.toLowerCase().endsWith(".docx")) {
      fd.append("files", f);
    }
  }
  const res = await fetch("/api/scan-upload", { method: "POST", body: fd });
  const js = await res.json();
  if (!js.ok) { log("ERROR: " + (js.error || res.status)); return; }
  uploadDirToken = js.token;
  lastFilesMeta = js.files_meta || [];
  setBadges(js.count || 0);
  drawFilesTable(lastFilesMeta);
  log(`Token: ${uploadDirToken} | Archivos: ${js.count}`);
}

async function mergeNow() {
  const mode = document.querySelector('input[name="mode"]:checked').value; // unificado | grouped
  const use_all = $("#use_all").checked;
  let selected = [];
  if (!use_all) {
    selected = Array.from($("#whitelist").selectedOptions).map(o => o.value);
    if (!selected.length) { log("Selecciona al menos un título o marca 'Usar todos'."); return; }
  } else {
    selected = TITULOS_WHITELIST.slice();
  }

  const payload = { source: lastSource, mode, titles: selected };

  if (lastSource === "folder") {
    const folder = $("#folder").value.trim();
    if (!folder) { log("Ingresa la ruta de la carpeta."); return; }
    payload.folder = folder;
    payload.include_subs = $("#include_subs").checked;
  } else if (lastSource === "upload") {
    if (!uploadToken) { log("Primero sube y escanea archivos."); return; }
    payload.token = uploadToken;
  } else {
    if (!uploadDirToken) { log("Primero sube y escanea la carpeta."); return; }
    payload.token = uploadDirToken;
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
  const token = uploadToken || uploadDirToken;
  if (!token) { log("No hay subidas por limpiar."); return; }
  await fetch("/api/cleanup", {
    method: "POST", headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ token })
  });
  uploadToken = null; uploadDirToken = null;
  log("Subidas temporales limpiadas.");
}

document.addEventListener("DOMContentLoaded", () => {
  fillWhitelist();
  $("#btn-scan-folder").addEventListener("click", scanFolder);
  $("#btn-scan-upload").addEventListener("click", scanUpload);
  $("#btn-scan-upload-dir").addEventListener("click", scanUploadDir);
  $("#btn-merge").addEventListener("click", mergeNow);
  $("#btn-clean").addEventListener("click", cleanup);

  $("#files-table").addEventListener("click", (ev) => {
    const btn = ev.target.closest("button[data-action='preview']");
    if (!btn) return;
    showPreview(btn.dataset.name);
  });

  $("#use_all").addEventListener("change", (e) => {
    if (e.target.checked) {
      for (const opt of $("#whitelist").options) opt.selected = false;
    }
  });
});
