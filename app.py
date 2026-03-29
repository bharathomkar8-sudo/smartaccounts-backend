from flask import Flask, request, send_file, render_template_string
import pandas as pd
import zipfile
from io import BytesIO
from mapper import process_sheet

app = Flask(__name__)

uploaded_file = None

# =========================
# HOME PAGE
# =========================
@app.route("/")
def home():
    return """
    <h2>Smart Accounts</h2>
    <a href="/upload">Go to Upload</a>
    """

# =========================
# UPLOAD PAGE
# =========================
@app.route("/upload", methods=["GET", "POST"])
def upload():
    global uploaded_file

    if request.method == "POST":
        uploaded_file = request.files["file"]
        return """
        <h3>File uploaded successfully</h3>
        <a href="/process">Go to Process</a>
        """

    return """
    <h2>Upload Excel</h2>
    <form method="post" enctype="multipart/form-data">
        <input type="file" name="file" required><br><br>
        <button type="submit">Upload</button>
    </form>
    """

# =========================
# PROCESS FILE
# =========================
@app.route("/process", methods=["GET"])
def process():

    global uploaded_file

    if uploaded_file is None:
        return "No file uploaded"

    try:
        # Read Excel
        df = pd.read_excel(uploaded_file, header=None)

        # Process using mapper
        output_df = process_sheet(df)

        # Save to memory
        output = BytesIO()

        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            output_df.to_excel(writer, index=False, sheet_name="Output")

        output.seek(0)

        return send_file(
            output,
            download_name="output.xlsx",
            as_attachment=True
        )

    except Exception as e:
        return f"<h3>Error:</h3><pre>{str(e)}</pre>"

# =========================
# RUN APP
# =========================
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
