import os
import shutil
import threading
import uuid
import re
from flask import Flask, request, jsonify, send_from_directory, render_template

from paper_builder import build_docx, convert_to_pdf, default_instructions, get_logo_path, STATIC_DIR

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
GEN_DIR = os.path.join(BASE_DIR, "generated")
UPLOAD_DIR = os.path.join(BASE_DIR, "static", "uploads")
os.makedirs(GEN_DIR, exist_ok=True)
os.makedirs(UPLOAD_DIR, exist_ok=True)

# Word/LibreOffice automation can only reliably do one conversion at a time
# (that's what running conversions handles by giving each its own throwaway
# profile/process - but two people clicking "Generate" in the exact same
# instant would still race for the same automation). At 5-10 users this
# just means an occasional very-short wait, never a real bottleneck.
_convert_lock = threading.Lock()

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024  # 16MB (question image upload)


def safe_slug(text):
    text = re.sub(r"[^A-Za-z0-9]+", "_", text or "paper").strip("_")
    return text[:40] or "paper"


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/logo")
def current_logo():
    """Serves whichever logo is currently used in the footer (for the preview)."""
    p = get_logo_path()
    if not p:
        return ("", 404)
    resp = send_from_directory(os.path.dirname(p), os.path.basename(p))
    resp.headers["Cache-Control"] = "no-store"
    return resp


def _clear_logo():
    for ext in ("png", "jpg", "jpeg"):
        p = os.path.join(STATIC_DIR, f"logo.{ext}")
        if os.path.exists(p):
            os.remove(p)


@app.route("/api/upload_logo", methods=["POST"])
def upload_logo():
    """Replaces the fixed footer logo. It stays in place for all future papers."""
    f = request.files.get("logo")
    if not f or f.filename == "":
        return jsonify({"ok": False, "error": "No file"}), 400
    ext = os.path.splitext(f.filename)[1].lower()
    if ext not in (".png", ".jpg", ".jpeg"):
        return jsonify({"ok": False, "error": "Logo must be PNG or JPG"}), 400
    _clear_logo()
    f.save(os.path.join(STATIC_DIR, "logo" + ext))
    return jsonify({"ok": True})


@app.route("/api/reset_logo", methods=["POST"])
def reset_logo():
    """Restores the original Mount Litera Zee School logo."""
    _clear_logo()
    shutil.copyfile(os.path.join(STATIC_DIR, "default_logo.png"), os.path.join(STATIC_DIR, "logo.png"))
    return jsonify({"ok": True})


@app.route("/api/upload_asset", methods=["POST"])
def upload_asset():
    """Generic image upload used for images inserted inside a question."""
    f = request.files.get("file")
    if not f or f.filename == "":
        return jsonify({"ok": False, "error": "No file"}), 400
    ext = os.path.splitext(f.filename)[1].lower()
    if ext not in (".png", ".jpg", ".jpeg", ".gif", ".bmp"):
        return jsonify({"ok": False, "error": "Image must be PNG, JPG, GIF or BMP"}), 400
    token = uuid.uuid4().hex[:12]
    fname = f"asset_{token}{ext}"
    path = os.path.join(UPLOAD_DIR, fname)
    f.save(path)
    return jsonify({"ok": True, "asset_id": fname, "url": f"/static/uploads/{fname}"})


@app.route("/api/auto_instructions", methods=["POST"])
def auto_instructions():
    data = request.get_json(force=True)
    sections = data.get("sections", [])
    if not sections or any(len(s.get("questions", [])) == 0 for s in sections):
        return jsonify({"ok": False, "error": "Add at least one question to every section first."}), 400
    lines = default_instructions(sections, extra=data.get("extra") or None)
    return jsonify({"ok": True, "instructions": lines})


@app.route("/api/generate", methods=["POST"])
def generate():
    data = request.get_json(force=True)

    errors = []
    for field in ("exam_title", "subject", "grade", "max_marks"):
        if not str(data.get(field, "")).strip():
            errors.append(f"Missing '{field}'.")
    sections = data.get("sections") or []
    if not sections:
        errors.append("Add at least one section.")
    for s in sections:
        if not s.get("questions"):
            errors.append(f"Section {s.get('label', '?')} has no questions.")
        for q in s.get("questions", []):
            if not str(q.get("text", "")).strip():
                errors.append(f"Section {s.get('label', '?')} has an empty question.")
    if errors:
        return jsonify({"ok": False, "errors": errors}), 400

    # Resolve per-question image asset ids -> real file paths on disk.
    for s in sections:
        for q in s.get("questions", []):
            resolved = []
            for asset_id in (q.get("images") or []):
                p = os.path.join(UPLOAD_DIR, asset_id)
                if os.path.exists(p):
                    resolved.append(p)
            q["image_paths"] = resolved

    instructions = data.get("instructions") or default_instructions(sections)

    paper = {
        "exam_title": data["exam_title"],
        "subject": data["subject"],
        "grade": data["grade"],
        "max_marks": data["max_marks"],
        "date": data.get("date", ""),
        "time": data.get("time", ""),
        "instructions": instructions,
        "sections": sections,
    }

    token = uuid.uuid4().hex[:10]
    base = f"{safe_slug(data['subject'])}_{safe_slug(data['exam_title'])}_{token}"
    docx_path = os.path.join(GEN_DIR, base + ".docx")

    try:
        build_docx(paper, docx_path)
        with _convert_lock:
            pdf_path = convert_to_pdf(docx_path, GEN_DIR)
    except Exception as e:
        return jsonify({"ok": False, "errors": [f"Generation failed: {e}"]}), 500

    total_marks = sum(len(s["questions"]) * float(s["marks_each"]) for s in sections)

    return jsonify({
        "ok": True,
        "docx_url": f"/download/{os.path.basename(docx_path)}",
        "pdf_url": f"/download/{os.path.basename(pdf_path)}",
        "total_marks": total_marks,
    })


@app.route("/download/<path:fname>")
def download(fname):
    return send_from_directory(GEN_DIR, fname, as_attachment=True)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    try:
        from waitress import serve
        # threads=8 comfortably covers 5-10 people on a LAN generating papers
        # now and then; each request is quick except the PDF conversion step.
        print(f"Serving on http://0.0.0.0:{port} (press Ctrl+C to stop)")
        serve(app, host="0.0.0.0", port=port, threads=8)
    except ImportError:
        # waitress not installed - fall back to Flask's own dev server so the
        # app still runs, but this isn't recommended for more than one user.
        print("waitress not installed (pip install waitress) - "
              "falling back to Flask's dev server, not recommended for multiple users.")
        app.run(host="0.0.0.0", port=port, debug=False)
