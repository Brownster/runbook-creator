from flask import Flask, request, render_template, send_from_directory, redirect, url_for
import yaml
from docx import Document
from docx.shared import Pt
import os
import tempfile
import shutil

app = Flask(__name__)
UPLOAD_FOLDER = tempfile.mkdtemp()
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def add_run_with_font(paragraph, text, size):
    run = paragraph.add_run(text)
    run.font.size = Pt(size)
    return run

def create_runbook_file(file_path, filename):
    with open(file_path, 'r') as file:
        data = yaml.safe_load(file)
    
    if data is None or 'groups' not in data:
        raise ValueError("Invalid or empty YAML content")
    
    generated_filenames = []
    for group in data['groups']:
        for rule in group['rules']:
            alert_name = rule['alert']
            expr = rule['expr']
            description = rule['annotations']['description']
            severity = rule['labels']['severity']
            
            doc_filename = f"{filename}_{alert_name.replace(' ', '_')}.docx"
            doc_path = os.path.join(UPLOAD_FOLDER, doc_filename)
            doc = Document()

            para = doc.add_paragraph()
            add_run_with_font(para, f"SM3 Alert RunBook – {group['name']} – {alert_name}", 14)

            para = doc.add_paragraph('Alert Name: ', style='Heading 2')
            add_run_with_font(para, alert_name, 12)

            para = doc.add_paragraph('Alert Expression: ', style='Heading 2')
            add_run_with_font(para, expr, 12)

            para = doc.add_paragraph('Category: ', style='Heading 2')
            add_run_with_font(para, group['name'], 12)

            para = doc.add_paragraph('Description: ', style='Heading 2')
            add_run_with_font(para, description, 12)

            para = doc.add_paragraph('Severity: ', style='Heading 2')
            add_run_with_font(para, severity, 12)

            doc.save(doc_path)
            generated_filenames.append(doc_filename)
    
    return generated_filenames

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        if 'file' not in request.files:
            return redirect(request.url)
        file = request.files['file']
        if file.filename == '':
            return redirect(request.url)
        if file:
            filename = file.filename
            file_path = os.path.join(UPLOAD_FOLDER, filename)
            file.save(file_path)
            try:
                generated_filenames = create_runbook_file(file_path, filename)
                # Here you might want to handle or display multiple files
                return redirect(url_for('download_file', filename=generated_filenames[0]))
            except ValueError as e:
                return str(e), 400
    return render_template('upload.html')

@app.route('/downloads/<filename>')
def download_file(filename):
    return send_from_directory(path=UPLOAD_FOLDER, filename=f"{filename}.docx", as_attachment=True)

@app.route('/cleanup', methods=['POST'])
def cleanup():
    shutil.rmtree(UPLOAD_FOLDER)
    os.makedirs(UPLOAD_FOLDER)
    return redirect(url_for('upload_file'))

if __name__ == "__main__":
    app.run(debug=True)
