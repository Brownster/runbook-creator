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

def create_runbook_file(file_stream, filename):
    file_stream.seek(0)  # Ensure the stream is at the start
    data = yaml.safe_load(file_stream)
    if data is None or 'groups' not in data:
        raise ValueError("Invalid or empty YAML content")
    
    for group in data['groups']:
        for rule in group['rules']:
            alert_name = rule['alert']
            expr = rule['expr']
            description = rule['annotations']['description']
            severity = rule['labels']['severity']
            
            doc_filename = f"{filename}_{alert_name.replace(' ', '_')}.docx"
            doc_path = os.path.join(app.config['UPLOAD_FOLDER'], doc_filename)
            doc = Document()

            # Correctly setting the font size using a helper function
            para = doc.add_paragraph()
            run = para.add_run(f"SM3 Alert RunBook – {group['name']} – {alert_name}")
            run.font.size = Pt(14)

            para = doc.add_paragraph('Alert Name:', style='Heading 2')
            run = para.add_run(alert_name)
            run.font.size = Pt(12)

            doc.save(doc_path)

    return filename

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
            save_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(save_path)
            try:
                runbook_basename = create_runbook_file(file.stream, filename)
                return redirect(url_for('download_file', filename=runbook_basename))
            except ValueError as e:
                return str(e), 400
    return render_template('upload.html')

@app.route('/downloads/<filename>')
def download_file(filename):
    return send_from_directory(directory=app.config['UPLOAD_FOLDER'], filename=f"{filename}.docx", as_attachment=True)

@app.route('/cleanup', methods=['POST'])
def cleanup():
    shutil.rmtree(app.config['UPLOAD_FOLDER'])
    os.makedirs(app.config['UPLOAD_FOLDER'])
    return redirect(url_for('upload_file'))

if __name__ == "__main__":
    app.run(debug=True)
