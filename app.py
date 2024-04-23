from flask import Flask, request, render_template, send_from_directory, redirect, url_for
import yaml
from docx import Document
from docx.shared import Pt
import os
import tempfile
import shutil

app = Flask(__name__)
UPLOAD_FOLDER = tempfile.mkdtemp()  # Uses temporary directory for file uploads and generated docs
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def add_heading(doc, text, size, level):
    heading = doc.add_heading(level=level)
    run = heading.add_run(text)
    run.font.size = Pt(size)
    return heading

def add_paragraph(doc, title, content, font_size):
    paragraph = doc.add_paragraph()
    run = paragraph.add_run(title + " ")  # Add space after title for clarity
    run.font.size = Pt(font_size)
    content_run = paragraph.add_run(content)
    content_run.font.size = Pt(font_size)
    return paragraph

def create_runbook_file(file_path, original_filename):
    with open(file_path, 'r') as file:
        data = yaml.safe_load(file)

    if data is None or 'groups' not in data:
        raise ValueError("Invalid or empty YAML content")
    
    generated_filenames = []
    for group in data['groups']:
        doc = Document()
        doc_filename = f"{original_filename}_{group['name'].replace(' ', '_')}.docx"
        doc_path = os.path.join(UPLOAD_FOLDER, doc_filename)
        
        add_heading(doc, f"RunBook for {group['name']}", 14, level=1)

        for rule in group['rules']:
            add_heading(doc, f"Alert: {rule['alert']}", 12, level=2)
            add_paragraph(doc, "Alert Expression:", rule['expr'], 11)
            add_paragraph(doc, "Category:", group['name'], 11)
            add_paragraph(doc, "Description:", rule['annotations']['description'], 11)
            add_paragraph(doc, "Notes:", rule['labels']['severity'], 11)
            doc.add_paragraph()  # Add a space between sections for clarity

        doc.save(doc_path)
        generated_filenames.append(doc_filename)
    
    return generated_filenames

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        file = request.files.get('file')
        if file and file.filename:
            filename = file.filename
            file_path = os.path.join(UPLOAD_FOLDER, filename)
            file.save(file_path)
            generated_filenames = create_runbook_file(file_path, filename)
            return redirect(url_for('download_file', filename=generated_filenames[0]))
        else:
            return "No file selected", 400
    return render_template('upload.html')

@app.route('/downloads/<filename>')
def download_file(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename, as_attachment=True)

@app.route('/cleanup', methods=['POST'])
def cleanup():
    shutil.rmtree(UPLOAD_FOLDER)
    os.makedirs(UPLOAD_FOLDER)
    return redirect(url_for('upload_file'))

if __name__ == "__main__":
    app.run(debug=True)
