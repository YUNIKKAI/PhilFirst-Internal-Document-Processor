from flask import Blueprint, render_template, request, send_file, flash, redirect, url_for, after_this_request
from renewal.renewal_notices import extract_renewal_notices
import shutil
import logging

renewal_bp = Blueprint("renewal", __name__)
logger = logging.getLogger("renewal")

@renewal_bp.route("/", methods=["GET", "POST"])
def renewal_handler():
    if request.method == "POST":
        files = request.files.getlist("pdf")
        if not files or all(f.filename == '' for f in files):
            flash("No PDF files uploaded.")
            return redirect(url_for("renewal.renewal_handler"))

        result = extract_renewal_notices(files, logger)
        if not result:
            flash("No valid renewal notices found.")
            return redirect(url_for("renewal.renewal_handler"))

        zip_path, zip_filename, temp_dir = result

        @after_this_request
        def cleanup(response):
            shutil.rmtree(temp_dir, ignore_errors=True)
            return response

        return send_file(zip_path, as_attachment=True, download_name=zip_filename)

    return render_template("renewal.html")