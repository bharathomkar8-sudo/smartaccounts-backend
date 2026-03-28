from flask import Flask, request, send_file
from flask_cors import CORS
import pandas as pd

app = Flask(__name__)
CORS(app)

@app.route("/")
def home():
    return "Backend Running ✅"

@app.route("/upload-sales", methods=["POST"])
def upload_sales():
    file = request.files.get("file")

    if not file:
        return "No file uploaded"

    filepath = "input.xlsx"
    file.save(filepath)

    df = pd.read_excel(filepath)

    # Example processing (you can change later)
    df["Processed"] = "Yes"

    output_path = "output.xlsx"
    df.to_excel(output_path, index=False)

    return send_file(output_path, as_attachment=True)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
