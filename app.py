from flask import Flask, request, send_file, render_template_string, session
import pandas as pd
import zipfile
import io
import os

app = Flask(__name__)
app.secret_key = "super_secret_key" # Needed to store sheet names in session

@app.route('/', methods=['GET', 'POST'])
def upload():
    if request.method == 'POST':
        file = request.files['file']
        if not file: return "No file uploaded"

        # Read file into memory immediately
        file_bytes = file.read()
        
        # We use a trick: store the bytes in a global-ish way 
        # or process it to get sheet names
        file_handle = io.BytesIO(file_bytes)
        xls = pd.ExcelFile(file_handle)
        sheets = xls.sheet_names
        
        # Save the bytes to a temporary location or keep in session 
        # (Note: For very large files, session might be too small)
        # For now, let's pass the data through a hidden field or temp storage
        # Simplest fix for Railway: Save with a UNIQUE name
        unique_name = f"temp_{os.getpid()}.xlsx"
        with open(unique_name, "wb") as f:
            f.write(file_bytes)
            
        session['current_file'] = unique_name

        return render_template_string('''
            <h2>Select Sheets</h2>
            <form method="POST" action="/process">
                {% for s in sheets %}
                    <input type="checkbox" name="sheets" value="{{s}}" checked> {{s}}<br>
                {% endfor %}
                <br><button type="submit">Process</button>
            </form>
        ''', sheets=sheets)

    return '<h2>Upload Excel</h2><form method="POST" enctype="multipart/form-data"><input type="file" name="file"><button type="submit">Upload</button></form>'

@app.route('/process', methods=['POST'])
def process():
    selected_sheets = request.form.getlist('sheets')
    temp_filename = session.get('current_file')

    if not temp_filename or not os.path.exists(temp_filename):
        return "File expired or not found. Please upload again."

    output_zip = io.BytesIO()
    
    # Use 'with' to ensure the file is closed properly
    with pd.ExcelFile(temp_filename) as xls:
        with zipfile.ZipFile(output_zip, 'w') as zf:
            for sheet in selected_sheets:
                try:
                    df = pd.read_excel(xls, sheet_name=sheet, header=None)
                    
                    # --- YOUR MAPPING LOGIC START ---
                    # (Keep your existing iloc logic here)
                    # --- YOUR MAPPING LOGIC END ---
                    
                    # Write result to memory
                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                        # Replace 'out_df' with your actual processed dataframe
                        df.to_excel(writer, index=False) 
                    
                    zf.writestr(f"{sheet}.xlsx", buffer.getvalue())
                except Exception as e:
                    print(f"Error processing {sheet}: {e}")

    # Cleanup temp file
    if os.path.exists(temp_filename):
        os.remove(temp_filename)

    output_zip.seek(0)
    return send_file(output_zip, download_name='output.zip', as_attachment=True)
