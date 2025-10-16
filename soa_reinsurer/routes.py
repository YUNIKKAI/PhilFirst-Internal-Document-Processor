from flask import Blueprint, render_template, request, send_file, flash, redirect, url_for, after_this_request
from soa_reinsurer.soa_reinsurer_premium import extract_soa_reinsurer_premium
from soa_reinsurer.soa_reinsurer_cashcall import extract_soa_reinsurer_cashcall
import shutil
import os
import traceback
import logging

# Setup logging
logger = logging.getLogger(__name__)

# ✅ Blueprint setup
soa_ri_bp = Blueprint("soa_reinsurer", __name__)

@soa_ri_bp.route("/", methods=["GET", "POST"])
def soa_handler():
    if request.method == "POST":
        # ✅ Get the selected type
        file_type = request.form.get("type")

        # 🧩 Validate selected type
        if not file_type:
            flash("Please select a file type (Cash Call or Premium).")
            return redirect(url_for("soa_reinsurer.soa_handler"))

        try:
            # 🧠 Process based on type
            if file_type == 'premium':
                # Get multiple premium files
                files = request.files.getlist("premium_files")
                
                if not files or all(f.filename.strip() == "" for f in files):
                    flash("No CSV files uploaded for Premium.")
                    return redirect(url_for("soa_reinsurer.soa_handler"))
                
                print(f"DEBUG: Processing {len(files)} premium file(s)")
                result = extract_soa_reinsurer_premium(files)
                
            elif file_type == 'cash-call':
                # Get the two separate files for Cash Call
                bulk_file = request.files.get("bulk_file")
                summary_file = request.files.get("summary_file")
                
                # Validate both files are present
                if not bulk_file or bulk_file.filename.strip() == "":
                    flash("Cash Call Bulk file is missing.")
                    return redirect(url_for("soa_reinsurer.soa_handler"))
                
                if not summary_file or summary_file.filename.strip() == "":
                    flash("Cash Call Summary file is missing.")
                    return redirect(url_for("soa_reinsurer.soa_handler"))
                
                # Pass both files as a list to maintain compatibility with your function
                files = [bulk_file, summary_file]
                result = extract_soa_reinsurer_cashcall(files)
                
            else:
                flash(f"Invalid file type: {file_type}")
                return redirect(url_for("soa_reinsurer.soa_handler"))
            
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
            error_msg = f"Error processing files: {str(e)}"
            print(f"ERROR: {error_msg}")
            print(f"TRACEBACK: {traceback.format_exc()}")
            logger.error(error_msg, exc_info=True)
            flash(error_msg)
            return redirect(url_for("soa_reinsurer.soa_handler"))

    # 🧾 Render upload page
    return render_template("soa_reinsurer.html")