import os
import shutil
import subprocess
import json
from flask import Flask, render_template, request, send_from_directory, redirect, url_for, flash
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = "supersecretkey"

# === Config ===
UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "outputs"
SCRIPT_FOLDER = "reporting_framework"
ALLOWED_EXTENSIONS = {"xlsx", "pptx", "png"}

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# === Script Map ===
SCRIPTS = {
    "Detailed NVT Report (Phase 1)": "phase_1/detailed_nvt_reports/detailed_nvt_reports.py",
    "Executive Report (Phase 1)": "phase_1/executive_report/executive_report.py",
    "Front Page (Phase 1)": "phase_1/front/front.py",
    "Executive Report (Phase 2)": "phase_2/executive_report/executive_report.py"
}

# === Helpers ===
def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS

def run_script(script_path, env):
    try:
        result = subprocess.run(
            ["python", script_path],
            check=True,
            text=True,
            capture_output=True,
            env=env
        )
        return True, result.stdout
    except subprocess.CalledProcessError as e:
        return False, e.stderr

# === Main Index Page ===
@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        report_type = request.form.get("report_type")
        custom_folder = secure_filename(request.form.get("custom_folder") or "default")
        custom_dir = os.path.join(OUTPUT_FOLDER, custom_folder)
        os.makedirs(custom_dir, exist_ok=True)

        # Set output directory environment variable
        env = os.environ.copy()
        env["CUSTOM_OUTPUT_DIR"] = os.path.abspath(custom_dir)

        # Define expected uploads
        uploads = {
            "sheet1": "sheet1.xlsx",
            "sheet2": "sheet2.xlsx",
            "template": "template.pptx"
        }

        # Determine target phase folder
        phase_folder = "phase_1" if "Phase 1" in report_type else "phase_2"

        # Upload handling
        for key, filename in uploads.items():
            uploaded_file = request.files.get(key)
            if uploaded_file and allowed_file(uploaded_file.filename):
                uploaded_data = uploaded_file.read()
                target_root = os.path.join(SCRIPT_FOLDER, phase_folder)
                for subdir in os.listdir(target_root):
                    sub_path = os.path.join(target_root, subdir)
                    if os.path.isdir(sub_path):
                        file_path = os.path.join(sub_path, filename)
                        with open(file_path, "wb") as f:
                            f.write(uploaded_data)

        # Special case: Front Page inputs
        if report_type == "Front Page (Phase 1)":
            input_data = {key: request.form.get(key, "") for key in request.form if key not in ["report_type", "custom_folder"]}
            front_input_path = os.path.join(SCRIPT_FOLDER, "phase_1", "front", "inputs.json")
            with open(front_input_path, "w") as f:
                json.dump(input_data, f, indent=2)

        # Run script
        script_rel_path = SCRIPTS.get(report_type)
        if not script_rel_path:
            flash("Invalid script selection.", "danger")
            return redirect(url_for("index"))

        script_path = os.path.join(SCRIPT_FOLDER, script_rel_path)
        success, msg = run_script(script_path, env)

        flash("Report generated successfully." if success else f"Script error: {msg}", "success" if success else "danger")
        return redirect(url_for("index"))

    return render_template("index.html", scripts=SCRIPTS)

# === Route: Front Page Form ===
@app.route("/front-page", methods=["GET"])
def front_page_form():
    return render_template("front_page.html")

@app.route("/front-page", methods=["GET", "POST"])
def front_page():
    if request.method == "POST":
        # Extract custom_folder from the form
        custom_folder = secure_filename(request.form.get("custom_folder", "default"))
        custom_dir = os.path.join(OUTPUT_FOLDER, custom_folder)
        os.makedirs(custom_dir, exist_ok=True)

        # Save all form inputs *excluding* 'custom_folder'
        input_data = {
            key: request.form.get(key, "")
            for key in request.form if key != "custom_folder"
        }

        input_path = os.path.join(SCRIPT_FOLDER, "phase_1", "front", "inputs.json")
        with open(input_path, "w") as f:
            json.dump(input_data, f, indent=2)

        # Set environment variable for output
        env = os.environ.copy()
        env["CUSTOM_OUTPUT_DIR"] = os.path.abspath(custom_dir)

        # Run front.py
        script_path = os.path.join(SCRIPT_FOLDER, "phase_1", "front", "front.py")
        success, msg = run_script(script_path, env)

        flash("Front page report generated successfully." if success else f"Script error: {msg}",
              "success" if success else "danger")
        return redirect(url_for("front_page"))

    return render_template("front_page.html")




# === Start Server ===
if __name__ == "__main__":
    app.run(debug=True, port=5000)
