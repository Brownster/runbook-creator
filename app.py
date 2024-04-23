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

def add_heading(doc, text, size, level):
    heading = doc.add_heading(level=level)
    run = heading.add_run(text)
    run.font.size = Pt(size)
    return heading

def add_paragraph(doc, title, content, title_size, content_size):
    title_para = doc.add_paragraph()
    title_run = title_para.add_run(title)
    title_run.font.size = Pt(title_size)
    
    content_para = doc.add_paragraph(content)
    content_run = content_para.add_run()  # Creates a new run for content
    content_run.font.size = Pt(content_size)
    return content_para

def create_runbook_file(file_path, original_filename):
    with open(file_path, 'r') as file:
        data = yaml.safe_load(file)

    if data is None or 'groups' not in data:
        raise ValueError("Invalid or empty YAML content")
    
    generated_filenames = []
    for group in data['groups']:
        # Initialize a new Word document for each group
        doc = Document()
        doc_filename = f"{original_filename}_{group['name'].replace(' ', '_')}.docx"
        doc_path = os.path.join(UPLOAD_FOLDER, doc_filename)
        
        # Add a title page or a section header for the group
        doc.add_heading(f"RunBook for {group['name']}", level=1)

        for rule in group['rules']:
            alert_name = rule['alert']
            expr = rule['expr']
            description = rule['annotations']['description']
            severity = rule['labels']['severity']

            # Add each alert as a new section in the document
            doc.add_heading(f"Alert: {alert_name}", level=2)
            doc.add_paragraph('Alert Expression:').add_run(expr).font.size = Pt(12)
            doc.add_paragraph('Category:').add_run(group['name']).font.size = Pt(12)
            doc.add_paragraph('Description:').add_run(description).font.size = Pt(12)
            doc.add_paragraph('Notes:').add_run(severity).font.size = Pt(12)
            doc.add_paragraph()

        # Save the document after all alerts have been added
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
                return redirect(url_for('download_file', filename=generated_filenames[0]))
            except ValueError as e:
                return str(e), 400
    return render_template('upload.html')

@app.route('/downloads/<filename>')
def download_file(filename):
    try:
        return send_from_directory(app.config['UPLOAD_FOLDER'], filename, as_attachment=True)
    except Exception as e:
        return str(e), 500


@app.route('/cleanup', methods=['POST'])
def cleanup():
    shutil.rmtree(UPLOAD_FOLDER)
    os.makedirs(UPLOAD_FOLDER)
    return redirect(url_for('upload_file'))

if __name__ == "__main__":
    app.run(debug=True)
