# app.py
import os, io, uuid, shutil, zipfile
from typing import List, Dict
from flask import Flask, render_template, request, jsonify, send_file

# === Tu lógica existente ===
from lector_word import procesar, procesar_grouped, headings_from_docx

app = Flask(__name__)

TMP_ROOT = os.path.join(os.path.dirname(__file__), "tmp")
os.makedirs(TMP_ROOT, exist_ok=True)

def _read_docx_from_folder(folder: str, include_subs: bool) -> List[Dict]:
    files = []
    if not os.path.isdir(folder):
        return files
    if include_subs:
        for root, _, fnames in os.walk(folder):
            for fn in fnames:
                if fn.lower().endswith(".docx") and not fn.startswith("~$"):
                    p = os.path.join(root, fn)
                    with open(p, "rb") as f:
                        files.append({"name": fn, "content": f.read()})
    else:
        for fn in os.listdir(folder):
            if fn.lower().endswith(".docx") and not fn.startswith("~$"):
                p = os.path.join(folder, fn)
                if os.path.isfile(p):
                    with open(p, "rb") as f:
                        files.append({"name": fn, "content": f.read()})
    return files

def _collect_overview(archivos: List[Dict], level: int):
    """
    Devuelve:
      - titles_unique: lista única (ordenada) de títulos del nivel
      - files_meta: [{name, count, titles}] por documento
    """
    seen = set()
    titles_unique = []
    files_meta = []
    for a in archivos:
        name = a.get("name", "archivo.docx")
        hs = []
        try:
            hs_all = headings_from_docx(a["content"])
            for h in hs_all:
                if int(h.get("level", 0)) == level:
                    t = (h.get("text") or "").strip()
                    if t:
                        hs.append(t)
                        if t not in seen:
                            seen.add(t)
                            titles_unique.append(t)
        except Exception:
            pass
        files_meta.append({
            "name": name,
            "count": len(hs),
            "titles": hs
        })
    titles_unique.sort(key=lambda s: s.casefold())
    files_meta.sort(key=lambda x: x["name"].casefold())
    return titles_unique, files_meta

@app.get("/")
def home():
    return render_template("index.html")

@app.post("/api/scan-folder")
def api_scan_folder():
    data = request.get_json(force=True)
    folder = (data.get("folder") or "").strip()
    include_subs = bool(data.get("include_subs", True))
    level = int(data.get("level", 1))
    if not folder or not os.path.isdir(folder):
        return jsonify({"ok": False, "error": "Carpeta inválida."}), 400

    archivos = _read_docx_from_folder(folder, include_subs)
    titles, files_meta = _collect_overview(archivos, level)
    return jsonify({
        "ok": True,
        "count": len(archivos),
        "titles": titles,
        "files": [{"name": a["name"]} for a in archivos],
        "files_meta": files_meta
    })

@app.post("/api/scan-upload")
def api_scan_upload():
    level = int(request.form.get("level", "1"))
    files = request.files.getlist("files")
    if not files:
        return jsonify({"ok": False, "error": "No enviaste .docx"}), 400

    token = uuid.uuid4().hex
    tmp_dir = os.path.join(TMP_ROOT, token)
    os.makedirs(tmp_dir, exist_ok=True)

    archivos = []
    for f in files:
        if not f.filename.lower().endswith(".docx"):
            continue
        raw = f.read()
        with open(os.path.join(tmp_dir, f.filename), "wb") as out:
            out.write(raw)
        archivos.append({"name": f.filename, "content": raw})

    titles, files_meta = _collect_overview(archivos, level)
    return jsonify({
        "ok": True,
        "token": token,
        "count": len(archivos),
        "titles": titles,
        "files": [{"name": a["name"]} for a in archivos],
        "files_meta": files_meta
    })

@app.post("/api/merge")
def api_merge():
    data = request.get_json(force=True)
    source = data.get("source")
    mode = (data.get("mode") or "unificado").lower()
    level = int(data.get("level", 1))
    use_all = bool(data.get("use_all", True))
    selected_titles = data.get("titles") or []
    titles = [] if use_all else selected_titles

    archivos = []
    if source == "folder":
        folder = (data.get("folder") or "").strip()
        include_subs = bool(data.get("include_subs", True))
        if not folder or not os.path.isdir(folder):
            return jsonify({"ok": False, "error": "Carpeta inválida."}), 400
        archivos = _read_docx_from_folder(folder, include_subs)
    elif source == "upload":
        token = data.get("token")
        if not token:
            return jsonify({"ok": False, "error": "Falta token de subida."}), 400
        tmp_dir = os.path.join(TMP_ROOT, token)
        if not os.path.isdir(tmp_dir):
            return jsonify({"ok": False, "error": "Token no válido o expirado."}), 400
        for fn in os.listdir(tmp_dir):
            if fn.lower().endswith(".docx"):
                p = os.path.join(tmp_dir, fn)
                with open(p, "rb") as f:
                    archivos.append({"name": fn, "content": f.read()})
    else:
        return jsonify({"ok": False, "error": "source inválido."}), 400

    if not archivos:
        return jsonify({"ok": False, "error": "No hay .docx para procesar."}), 400

    memzip = io.BytesIO()
    with zipfile.ZipFile(memzip, "w", zipfile.ZIP_DEFLATED) as zf:
        if mode == "unificado":
            res = procesar(archivos, [1, 2, 3], titles, enforce_whitelist=False)
            for fname, data_bytes in res.items():
                zf.writestr(fname, data_bytes)
        else:
            res = procesar_grouped(archivos, level, titles, enforce_whitelist=False)
            if not res:
                return jsonify({"ok": False, "error": "No se generó ningún archivo (verifica nivel/títulos)."}), 400
            for fname, data_bytes in res.items():
                safe = fname.replace("/", "_").replace("\\", "_")
                zf.writestr(safe, data_bytes)

    memzip.seek(0)
    return send_file(
        memzip,
        mimetype="application/zip",
        as_attachment=True,
        download_name=("unificado.zip" if mode == "unificado" else "por_titulo.zip")
    )

@app.post("/api/cleanup")
def api_cleanup():
    token = (request.get_json(force=True) or {}).get("token")
    if not token:
        return jsonify({"ok": True})
    tmp_dir = os.path.join(TMP_ROOT, token)
    if os.path.isdir(tmp_dir):
        shutil.rmtree(tmp_dir, ignore_errors=True)
    return jsonify({"ok": True})

@app.get("/health")
def health():
    return "OK", 200

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8000, debug=False)
