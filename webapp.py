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
import json
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
            
            # Extract invoices (can be multiple from PDF)
            invoice_numbers, parsed_info_list = extract_invoice(file_path)
            if not parsed_info_list:
                flash('No information extracted from the file.', 'danger')
                return render_template('result.html', records=parsed_info_list, all_crm_rows=[], logs="")
            else:
                # Filter out None invoice numbers for CRM processing
                valid_invoice_numbers = [inv for inv in invoice_numbers if inv is not None]
                flash(f'Extracted {len(parsed_info_list)} page(s) with {len(valid_invoice_numbers)} invoice(s). Launching CRM automation...', 'success')
                
                try:
                    # Run CRM script once for all valid invoices (more efficient)
                    if valid_invoice_numbers:
                        cmd = [sys.executable, 'automate_crm_login.py', '--headless', '--no-interactive', '--web-output']
                        cmd.extend(['--invoices'] + valid_invoice_numbers)
                    else:
                        # No valid invoices found, skip CRM processing
                        flash('No valid invoice numbers found for CRM processing.', 'warning')
                        return render_template('result.html', records=parsed_info_list, all_crm_rows=[], logs="No valid invoice numbers found for CRM processing.")
                    
                    proc = subprocess.run(cmd, capture_output=True, text=True, check=False, timeout=600)  # Longer timeout for multiple invoices

                    # Combine stdout and stderr for logs
                    combined_logs = f"STDOUT:\n{proc.stdout}\n\nSTDERR:\n{proc.stderr}"

                    # Try to parse JSON from stdout if available
                    all_crm_rows = []
                    if proc.stdout:
                        try:
                            # Look for JSON in stdout
                            lines = proc.stdout.strip().split('\n')
                            for line in lines:
                                if line.strip().startswith('[') and line.strip().endswith(']'):
                                    all_crm_rows = json.loads(line)
                                    break
                        except:
                            pass

                    if proc.returncode == 0:
                        if all_crm_rows:
                            flash(f'CRM automation completed! Found {len(all_crm_rows)} total records.', 'success')
                        else:
                            flash('CRM automation completed but no records found.', 'warning')
                    else:
                        flash(f'CRM script exited with code {proc.returncode}. Check logs for details.', 'warning')
                except subprocess.CalledProcessError as e:
                    flash(f'CRM script error: {e.stderr}', 'danger')
                    all_crm_rows = []
                    combined_logs = f"ERROR: {str(e)}\nSTDERR: {e.stderr}"
                except subprocess.TimeoutExpired as e:
                    flash('CRM script timed out after 10 minutes.', 'danger')
                    all_crm_rows = []
                    combined_logs = f"TIMEOUT: {str(e)}\nSTDOUT: {e.stdout}\nSTDERR: {e.stderr}"

                flash(f'Processing completed! Total CRM records found: {len(all_crm_rows)}', 'info')
                return render_template('result.html', records=parsed_info_list, all_crm_rows=all_crm_rows, logs=combined_logs)
        else:
            flash('File type not allowed. Please upload an image.', 'danger')
            return redirect(request.url)
    return render_template('index.html')

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5002, debug=True) 