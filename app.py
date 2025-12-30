import os
import io
import time
import uuid
import tempfile
import zipfile
import re

from flask import (
    Flask,
    render_template,
    request,
    send_file,
    redirect,
    url_for,
    flash,
)
import pandas as pd
from docxtpl import DocxTemplate
from docx2pdf import convert


app = Flask(
    __name__,
    static_folder="static",
    template_folder="templates",
)

app.secret_key = "super-secret-key"

# Store generated ZIPs in memory
GENERATED_ZIPS = {}

# Safe filename helper
def sanitize_filename(name: str) -> str:
    # remove path separators and control chars, trim
    name = re.sub(r"[\\/]+", "-", name)
    # replace multiple spaces with single space
    name = re.sub(r"\s+", " ", name).strip()
    # remove any characters that are problematic in filenames
    name = re.sub(r"[:\*\?\"<>\|]+", "", name)
    return name or "file"

# Render a filename from template and a row (and index)
def render_filename(template: str, row: pd.Series, idx: int) -> str:
    # Prepare values
    try:
        joining_date_val = pd.to_datetime(row.get("joining_date", "")).strftime("%d-%m-%Y")
    except Exception:
        joining_date_val = str(row.get("joining_date", ""))
    values = {
        "emp_name": str(row.get("emp_name", "")).strip(),
        "city": str(row.get("city", "")).strip(),
        "state": str(row.get("state", "")).strip(),
        "joining_date": joining_date_val,
        "index": str(idx),
    }
    # Replace placeholders
    result = template
    for k, v in values.items():
        result = result.replace("{" + k + "}", v)
    result = sanitize_filename(result)
    return result

@app.route("/", methods=["GET", "POST"])
def index():

    if request.method == "POST":

        # Highly accurate timer
        start_time = time.perf_counter()

        excel_file = request.files.get("employees_file")
        template_file = request.files.get("template_file")
        filename_template = request.form.get("filename_template", "{emp_name} LCF NDA Form").strip()
        preview_count_raw = request.form.get("preview_count", "5").strip()
        action = request.form.get("action", "generate")  # "preview" or "generate"

        # Validate uploads
        if not excel_file or not template_file:
            flash("Please upload both the Employees Excel file and the NDA Template DOCX.")
            return redirect(url_for("index"))

        # Validate preview_count
        try:
            preview_count = int(preview_count_raw)
            if preview_count < 1:
                preview_count = 5
        except Exception:
            preview_count = 5

        with tempfile.TemporaryDirectory() as tmpdir:

            # Save uploaded files to temp
            excel_path = os.path.join(tmpdir, "employees.xlsx")
            template_path = os.path.join(tmpdir, "template.docx")
            excel_file.save(excel_path)
            template_file.save(template_path)

            # Read Excel
            try:
                df = pd.read_excel(excel_path)
            except Exception as e:
                flash(f"Unable to read Excel file: {e}")
                return redirect(url_for("index"))

            required_cols = ["emp_name", "city", "state", "joining_date", "address", "gender"]
            missing = [c for c in required_cols if c not in df.columns]

            if missing:
                flash(f"Missing required Excel columns: {', '.join(missing)}")
                return redirect(url_for("index"))

            # If user wanted preview only -> compute filenames for first N rows and show them
            if action == "preview":
                preview_rows = []
                for i, (_, row) in enumerate(df.iterrows(), start=1):
                    preview_rows.append({
                        "index": i,
                        "emp_name": str(row.get("emp_name", "")).strip(),
                        "filename": render_filename(filename_template, row, i) + ".docx"
                    })
                    if i >= preview_count:
                        break

                result = {
                    "total": len(df),
                    "preview_count": preview_count,
                    "preview_rows": preview_rows,
                    "filename_template": filename_template,
                }

                # no heavy processing, just render preview list
                return render_template("index.html", preview=result)

            # Otherwise, proceed to generate DOCX, PDFs and ZIP using filename_template
            docx_dir = os.path.join(tmpdir, "DOCX")
            pdf_dir = os.path.join(tmpdir, "PDF")
            os.makedirs(docx_dir, exist_ok=True)
            os.makedirs(pdf_dir, exist_ok=True)

            total = len(df)
            success_count = 0
            error_count = 0
            report_rows = []

            # Process each employee
            for i, (_, row) in enumerate(df.iterrows(), start=1):

                emp_name = str(row["emp_name"]).strip()
                city = str(row["city"]).strip()
                state = str(row["state"]).strip()
                joining_date = row["joining_date"]
                address = str(row["address"]).strip()
                gender = str(row["gender"]).strip().lower()

                his_or_her = "his" if gender in ["male", "m"] else "her"

                # Format date
                try:
                    joining_date_str = pd.to_datetime(joining_date).strftime("%d-%m-%Y")
                except:
                    joining_date_str = str(joining_date)

                doc = DocxTemplate(template_path)

                context = {
                    "emp_name": emp_name,
                    "city": city,
                    "state": state,
                    "joining_date": joining_date_str,
                    "address": address,
                    "his_or_her": his_or_her,
                }

                # Use filename template
                base_name = render_filename(filename_template, row, i)

                word_path = os.path.join(docx_dir, base_name + ".docx")
                pdf_path = os.path.join(pdf_dir, base_name + ".pdf")

                try:
                    # DOCX
                    doc.render(context)
                    doc.save(word_path)

                    # PDF conversion - convert may raise; catch it
                    try:
                        convert(word_path, pdf_path)
                    except Exception as conv_e:
                        # If conversion fails, still mark error but keep DOCX
                        raise conv_e

                    status = "Success"
                    success_count += 1

                except Exception as e:
                    status = f"Error: {e}"
                    error_count += 1

                report_rows.append({"emp_name": emp_name, "filename": base_name, "status": status})

            # Create ZIP in memory
            zip_buffer = io.BytesIO()

            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:

                # Add DOCX
                for f in os.listdir(docx_dir):
                    zipf.write(os.path.join(docx_dir, f), f"DOCX/{f}")

                # Add PDF
                for f in os.listdir(pdf_dir):
                    zipf.write(os.path.join(pdf_dir, f), f"PDF/{f}")

                # Add report Excel
                report_df = pd.DataFrame(report_rows)
                report_io = io.BytesIO()

                with pd.ExcelWriter(report_io, engine="openpyxl") as writer:
                    report_df.to_excel(writer, index=False, sheet_name="Report")

                report_io.seek(0)
                zipf.writestr("NDA_Report.xlsx", report_io.read())

            zip_buffer.seek(0)

            # Unique ZIP ID
            zip_id = str(uuid.uuid4())
            GENERATED_ZIPS[zip_id] = zip_buffer

            # ⏳ Accurate time calculation
            elapsed = time.perf_counter() - start_time

            if elapsed < 60:
                time_taken = f"{elapsed:.2f} sec"
            else:
                minutes = int(elapsed // 60)
                seconds = elapsed % 60
                time_taken = f"{minutes} min {seconds:.2f} sec"

            # Final result
            result = {
                "total": total,
                "success": success_count,
                "error": error_count,
                "time": time_taken,
                "zip_id": zip_id,
                "filename_template": filename_template,
            }

            return render_template("index.html", result=result)

    return render_template("index.html")


@app.route("/download/<zip_id>")
def download_zip(zip_id):
    if zip_id not in GENERATED_ZIPS:
        return "Invalid or expired download link.", 404

    # Seek to start for send_file
    GENERATED_ZIPS[zip_id].seek(0)
    return send_file(
        GENERATED_ZIPS[zip_id],
        as_attachment=True,
        download_name="NDA_Forms.zip",
        mimetype="application/zip",
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)
