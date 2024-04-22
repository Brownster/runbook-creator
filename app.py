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

def create_runbook_file(yaml_content, filename):
    data = yaml.safe_load(yaml_content)
    for group in data['groups']:
        for rule in group['rules']:
            alert_name = rule['alert']
            expr = rule['expr']
            description = rule['annotations']['description']
            severity = rule['labels']['severity']
            
            doc_filename = f"{filename}_{alert_name.replace(' ', '_')}.docx"
            doc_path = os.path.join(app.config['UPLOAD_FOLDER'], doc_filename)
            doc = Document()

            doc.add_heading(f'SM3 Alert RunBook – {group['name']} – {alert_name}', level=1).font.size = Pt(14)
            doc.add_heading('Alert Name:', level=2).font.size = Pt(12)
            doc.add_paragraph(alert_name)
            doc.add_heading('Alert Expression:', level=2).font.size = Pt(12)
            doc.add_paragraph(expr)
            doc.add_heading('Category:', level=2).font.size = Pt(12)
            doc.add_paragraph(group['name'])
            doc.add_heading('Description:', level=2).font.size = Pt(12)
            doc.add_paragraph(description)
            doc.add_heading('Possible Cause(s):', level=2).font.size = Pt(12)
            doc.add_heading('Impact:', level=2).font.size = Pt(12)
            doc.add_heading('Next Steps:', level=2).font.size = Pt(12)
            doc.add_heading('Extra notes:', level=2).font.size = Pt(12)
            severity_paragraph = doc.add_paragraph()
            run = severity_paragraph.add_run(f'Severity level is {severity}, as it may not immediately impact service but requires corrective action.')
            run.font.size = Pt(10)
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
            runbook_basename = create_runbook_file(file.stream, filename)
            return redirect(url_for('download_file', filename=runbook_basename))
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
