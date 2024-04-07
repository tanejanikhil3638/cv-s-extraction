from flask import Flask, render_template, request, send_file
import os
import zipfile
import re
import PyPDF2
from docx import Document
import pandas as pd
from werkzeug.utils import secure_filename

app = Flask(__name__)

def extract_info_from_pdf(pdf_file):
    reader = PyPDF2.PdfReader(pdf_file)
    text = ""
    for page_num in range(len(reader.pages)):
        text += reader.pages[page_num].extract_text()
    return text

def extract_info_from_docx(docx_file):
    doc = Document(docx_file)
    text = ""
    for para in doc.paragraphs:
        text += para.text + "\n"
    return text

def extract_email_and_phone(text):
    emails = re.findall(r'[\w\.-]+@[\w\.-]+', text)
    phones = re.findall(r'\b\d{10}\b', text)
    return emails, phones

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        zip_file = request.files['zip_file']
        zip_filename = secure_filename(zip_file.filename)
        
        # Create a folder to extract zip contents
        zip_folder = os.path.join(app.config['UPLOAD_FOLDER'], zip_filename)
        os.makedirs(zip_folder, exist_ok=True)
        
        # Save zip file and extract contents
        zip_filepath = os.path.join(zip_folder, zip_filename)
        zip_file.save(zip_filepath)
        
        with zipfile.ZipFile(zip_filepath, 'r') as zip_ref:
            zip_ref.extractall(zip_folder)
        
        # Process extracted files
        data = []
        pdf_text = ""
        docx_text =""
        for filename in os.listdir(zip_folder):
            file_ext = os.path.splitext(filename)[1]
            
            if file_ext == '.pdf':
                pdf_filepath = os.path.join(zip_folder, filename)
                pdf_text = extract_info_from_pdf(pdf_filepath)
                emails, phones = extract_email_and_phone(pdf_text)
                
                overall_text = pdf_text
                data.append({
                    "Filename": filename,
                    "Email IDs": emails,
                    "Contact Numbers": phones,
                    "Overall Text": overall_text
                })
            elif file_ext == '.docx':
                docx_filepath = os.path.join(zip_folder, filename)
                docx_text = extract_info_from_docx(docx_filepath)
                emails, phones = extract_email_and_phone(docx_text)
                
                overall_text = docx_text
                data.append({
                    "Filename": filename,
                    "Email IDs": emails,
                    "Contact Numbers": phones,
                    "Overall Text": overall_text
                })
        
        df = pd.DataFrame(data)
        excel_filename = "CV_Info_All.xlsx"
        df.to_excel(excel_filename, index=False)
        
        return send_file(excel_filename, as_attachment=True, download_name=excel_filename)
    
    return render_template('index.html')

if __name__ == '__main__':
    app.config['UPLOAD_FOLDER'] = 'uploads'
    app.run(debug=True)
