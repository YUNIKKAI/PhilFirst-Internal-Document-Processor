from flask import Blueprint, render_template, request, send_file, flash, redirect, url_for, after_this_request
from soa_reinsurer.soa_reinsurer_processor import extract_soa_reinsurer
import shutil
import os

# ✅ Blueprint setup
soa_ri_bp = Blueprint("soa_reinsurer", __name__)

@soa_ri_bp.route("/", methods=["GET", "POST"])
def soa_handler():
    if request.method == "POST":
        # ✅ Grab uploaded files + form type
        files = request.files.getlist("files")
        file_type = request.form.get("type")

        # 🧩 Validate file presence
        if not files or all(f.filename.strip() == "" for f in files):
            flash("No CSV files uploaded.")
            return redirect(url_for("soa_reinsurer.soa_handler"))

        # 🧩 Validate selected type
        if not file_type:
            flash("Please select a file type (Cash Call or Premium).")
            return redirect(url_for("soa_reinsurer.soa_handler"))

        try:
            # 🧠 Process the uploaded files
            result = extract_soa_reinsurer(files, file_type)
            if not result:
                flash("No valid SOA data found.")
                return redirect(url_for("soa_reinsurer.soa_handler"))

            zip_path, zip_filename, temp_dir = result

            # 🧹 Clean up temp folder after response
            @after_this_request
            def cleanup(response):
                try:
                    if os.path.exists(temp_dir):
                        shutil.rmtree(temp_dir, ignore_errors=True)
                except Exception as cleanup_err:
                    print(f"Cleanup error: {cleanup_err}")
                return response

            # 📨 Send back the generated ZIP file with completion cookie
            response = send_file(zip_path, as_attachment=True, download_name=zip_filename)
            response.set_cookie('download_started', '1', max_age=3)
            return response

        except Exception as e:
            flash(f"Error processing files: {str(e)}")
            return redirect(url_for("soa_reinsurer.soa_handler"))

    # 🧾 Render upload page
    return render_template("soa_reinsurer.html")