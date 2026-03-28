from flask import Flask, request, send_file
from flask_cors import CORS
import pandas as pd
import os

app = Flask(__name__)

# 🔥 FIX CORS
CORS(app, resources={r"/*": {"origins": "*"}})

@app.route("/")
def home():
    return "Backend Running ✅"

@app.route("/upload-sales", methods=["POST"])
def upload_sales():
    try:
        file = request.files.get("file")

        if not file:
            return "No file uploaded", 400

        filepath = "input.xlsx"
        file.save(filepath)

        df = pd.read_excel(filepath)

        # Example processing
        df["Processed"] = "Yes"

        output_path = "output.xlsx"
        df.to_excel(output_path, index=False)

        return send_file(output_path, as_attachment=True)

    except Exception as e:
        return f"Error: {str(e)}", 500


# 🔥 VERY IMPORTANT FOR RAILWAY
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
