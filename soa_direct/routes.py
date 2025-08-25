from flask import Blueprint, render_template, request, send_file, flash, redirect, url_for, after_this_request
from soa_direct.soa_direct_processor import extract_soa_direct
import shutil

soa_bp = Blueprint("soa_direct", __name__)

@soa_bp.route("/", methods=["GET", "POST"])
def soa_handler():
    if request.method == "POST":
        files = request.files.getlist("files")
        if not files or all(f.filename == "" for f in files):
            flash("No CSV files uploaded.")
            return redirect(url_for("soa.soa_handler"))

        result = extract_soa_direct(files)
        if not result:
            flash("No valid SOA data found.")
            return redirect(url_for("soa.soa_handler"))

        zip_path, zip_filename, temp_dir = result

        @after_this_request
        def cleanup(response):
            shutil.rmtree(temp_dir, ignore_errors=True)
            return response

        return send_file(zip_path, as_attachment=True, download_name=zip_filename)

    return render_template("soa_direct.html")
