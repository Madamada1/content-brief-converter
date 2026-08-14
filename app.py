import datetime
import os
import secrets
import tempfile
from datetime import timedelta

from flask import Flask, jsonify, redirect, render_template, request, send_from_directory, session, url_for
from werkzeug.utils import secure_filename

from docx_parser import BriefParseError, parse_docx
from xlsx_builder import build_workbook, safe_filename

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "outputs")
os.makedirs(OUTPUT_DIR, exist_ok=True)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50MB
app.secret_key = os.environ.get("SECRET_KEY", secrets.token_hex(32))
app.permanent_session_lifetime = timedelta(days=7)

# If APP_PASSWORD isn't set (e.g. running locally via Start App.bat), the
# login screen is skipped entirely. Set it as an environment variable on
# the hosting platform to require a password on the hosted version.
APP_PASSWORD = os.environ.get("APP_PASSWORD")


@app.before_request
def require_login():
    if not APP_PASSWORD:
        return
    if request.endpoint in ("login", "static"):
        return
    if not session.get("authenticated"):
        return redirect(url_for("login"))


@app.route("/login", methods=["GET", "POST"])
def login():
    error = None
    if request.method == "POST":
        if request.form.get("password") == APP_PASSWORD:
            session.permanent = True
            session["authenticated"] = True
            return redirect(url_for("index"))
        error = "Incorrect password."
    return render_template("login.html", error=error)


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/convert", methods=["POST"])
def convert():
    files = request.files.getlist("files")
    if not files:
        return jsonify({"results": [{"filename": None, "success": False, "error": "No files received."}]})

    results = []
    for upload in files:
        original_name = upload.filename or "upload.docx"
        result = {"filename": original_name, "success": False, "error": None, "download_url": None}

        if not original_name.lower().endswith(".docx"):
            result["error"] = "Not a .docx file."
            results.append(result)
            continue

        tmp_path = None
        try:
            fd, tmp_path = tempfile.mkstemp(suffix=".docx")
            os.close(fd)
            upload.save(tmp_path)

            parsed = parse_docx(tmp_path)
            parsed["source_filename"] = original_name
            workbook = build_workbook(parsed)

            base_name = safe_filename(parsed.get("content_page_name"), fallback="content-brief")
            timestamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
            out_filename = f"{base_name}-{timestamp}.xlsx"
            out_path = os.path.join(OUTPUT_DIR, out_filename)
            workbook.save(out_path)

            result["success"] = True
            result["download_url"] = f"/outputs/{out_filename}"
            result["output_filename"] = out_filename
            result["content_page_name"] = parsed.get("content_page_name")
        except BriefParseError as exc:
            result["error"] = str(exc)
        except Exception as exc:  # noqa: BLE001 - surface any failure to the UI
            result["error"] = f"Unexpected error: {exc}"
        finally:
            if tmp_path and os.path.exists(tmp_path):
                os.remove(tmp_path)

        results.append(result)

    return jsonify({"results": results})


@app.route("/outputs/<path:filename>")
def download(filename):
    safe_name = secure_filename(filename)
    return send_from_directory(OUTPUT_DIR, safe_name, as_attachment=True)


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5050, debug=False)
