#!/usr/bin/env python3
"""
Simple Flask web server that lets a user upload a bank payment advice image.
It extracts structured information via OCR + LLM, displays the results,
then launches the CRM automation script to search by invoice number.
"""

from flask import Flask, render_template, request, redirect, url_for, flash
import os
import uuid
import subprocess
import sys
from imagedetect import extract_invoice

UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg', 'tiff', 'bmp', 'gif', 'pdf'}

app = Flask(__name__)
app.secret_key = 'supersecret'  # change in production
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        if 'file' not in request.files:
            flash('No file part', 'warning')
            return redirect(request.url)
        file = request.files['file']
        if file.filename == '':
            flash('No selected file', 'warning')
            return redirect(request.url)
        if file and allowed_file(file.filename):
            filename = f"{uuid.uuid4().hex}_{file.filename}"
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(file_path)
            
            # Extract invoice
            invoice, parsed_info = extract_invoice(file_path)
            if not invoice:
                flash('Invoice number not detected.', 'danger')
                return render_template('result.html', info=parsed_info, invoice=None, crm_rows=[])
            else:
                flash(f'Invoice detected: {invoice}. Launching CRM automation...', 'success')
                try:
                    proc = subprocess.run([
                        sys.executable, 'automate_crm_login.py', '--invoice', invoice, '--headless', '--json'
                    ], capture_output=True, text=True, check=True, timeout=300)
                    crm_rows = []
                    if proc.stdout:
                        import json
                        crm_rows = json.loads(proc.stdout)
                    flash('CRM automation completed successfully!', 'success')
                except subprocess.CalledProcessError as e:
                    flash(f'CRM script error: {e.stderr}', 'danger')
                    crm_rows = []
                except subprocess.TimeoutExpired:
                    flash('CRM script timed out after 5 minutes.', 'danger')
                    crm_rows = []
                return render_template('result.html', info=parsed_info, invoice=invoice, crm_rows=crm_rows)
        else:
            flash('File type not allowed. Please upload an image.', 'danger')
            return redirect(request.url)
    return render_template('index.html')

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5002, debug=True) 