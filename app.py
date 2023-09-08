import os
from flask import Flask, render_template, request, send_file
import pandas as pd
from io import BytesIO

app = Flask(__name__)

# Define the upload folder
UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Function to process the uploaded Excel file
def process_excel(input_file):
    # Read the Excel file (you may need to adjust this depending on your processing logic)
    df = pd.read_excel(input_file)

    # Perform some processing on the DataFrame
    # For example, add a new column or modify the data

    # Create a new Excel file
    output_buffer = BytesIO()
    with pd.ExcelWriter(output_buffer, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Processed_Data', index=False)

    output_buffer.seek(0)

    return output_buffer

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'excelFile' not in request.files:
        return "No file part"

    file = request.files['excelFile']

    if file.filename == '':
        return "No selected file"

    if file:
        # Save the uploaded file
        filename = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
        file.save(filename)

        # Process the Excel file
        output_buffer = process_excel(filename)

        # Return the processed file for download
        return send_file(output_buffer, as_attachment=True, download_name='processed_excel.xlsx')

if __name__ == '__main__':
    app.run(debug=True)
