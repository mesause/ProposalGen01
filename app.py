#!/usr/bin/env python3
"""
Flask Web Application for Document Generation

This app lets users select a DOCX template (excluding any sanitized copies),
fill in the placeholder values via a web form, and then generates a new DOCX
document (with the filename automatically built from the values of "Client Company Name"
and "Proposal date"). All sanitized templates are stored in the "sanitized_templates"
folder, and generated documents in the "output" folder.
 
Dependencies:
  - Python 3.9+
  - Flask
  - docxtpl
  - docx2pdf (if you choose to implement PDF conversion; see notes below)
"""

import os
import re
import sys
import glob
import zipfile
import tempfile
import shutil
from flask import Flask, render_template, request, redirect, url_for, send_from_directory, flash
from docxtpl import DocxTemplate
# from docx2pdf import convert  # Uncomment if PDF conversion is set up on your Linux server

app = Flask(__name__)
app.secret_key = "supersecretkey"  # required for flash messages

# Define directories (all relative to the location of app.py)
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
# For our purposes the original templates are in BASE_DIR (or you can change this to a dedicated folder)
ORIGINAL_TEMPLATES_DIR = BASE_DIR  
SANITIZED_DIR = os.path.join(BASE_DIR, "sanitized_templates")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")

# Ensure the necessary folders exist.
os.makedirs(SANITIZED_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# --------------------
# Utility Functions
# --------------------

def extract_placeholders_from_xml(docx_path):
    """
    Extract placeholders from the document.xml inside the DOCX file.
    Uses a regex (with DOTALL) to find all text between {{ and }},
    strips out any embedded XML tags, and returns a list of unique placeholder strings.
    """
    try:
        with zipfile.ZipFile(docx_path, "r") as z:
            xml_content = z.read("word/document.xml").decode("utf-8")
    except Exception as e:
        print(f"Error reading the DOCX file: {e}")
        return []
    raw_matches = re.findall(r'{{(.*?)}}', xml_content, re.DOTALL)
    placeholders = set()
    for match in raw_matches:
        cleaned = re.sub(r'<[^>]+>', '', match).strip()
        if cleaned:
            placeholders.add(cleaned)
    return list(placeholders)

def sanitize_placeholder(placeholder):
    """
    Convert a placeholder (which may contain spaces or punctuation)
    to a valid Python identifier by replacing non-alphanumeric characters with underscores.
    """
    sanitized = re.sub(r'\W+', '_', placeholder)
    return sanitized.strip('_')

def sanitize_template_xml(template_path, mapping, sanitized_dir):
    """
    Create a sanitized version of the DOCX template.
    For each placeholder in the DOCX that, when cleaned, matches one of the keys in mapping,
    replace it with the corresponding sanitized version.
    The new DOCX is stored in the sanitized_dir.
    """
    os.makedirs(sanitized_dir, exist_ok=True)
    temp_dir = tempfile.mkdtemp()
    with zipfile.ZipFile(template_path, 'r') as zin:
        zin.extractall(temp_dir)
    doc_xml_path = os.path.join(temp_dir, "word", "document.xml")
    try:
        with open(doc_xml_path, "r", encoding="utf-8") as f:
            xml_content = f.read()
    except Exception as e:
        print(f"Error reading document.xml: {e}")
        shutil.rmtree(temp_dir)
        return None
    def replacement(match):
        full_match = match.group(0)
        inner = match.group(1)
        cleaned = re.sub(r'<[^>]+>', '', inner).strip()
        if cleaned in mapping:
            return "{{" + mapping[cleaned] + "}}"
        else:
            return full_match
    new_xml = re.sub(r'{{(.*?)}}', replacement, xml_content, flags=re.DOTALL)
    try:
        with open(doc_xml_path, "w", encoding="utf-8") as f:
            f.write(new_xml)
    except Exception as e:
        print(f"Error writing modified document.xml: {e}")
        shutil.rmtree(temp_dir)
        return None
    sanitized_template_path = os.path.join(sanitized_dir, "sanitized_" + os.path.basename(template_path))
    with zipfile.ZipFile(sanitized_template_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for foldername, subfolders, filenames in os.walk(temp_dir):
            for filename in filenames:
                file_path = os.path.join(foldername, filename)
                arcname = os.path.relpath(file_path, temp_dir)
                zout.write(file_path, arcname)
    shutil.rmtree(temp_dir)
    return sanitized_template_path

def get_value_case_insensitive(dictionary, target_key, default):
    """
    Retrieve a value from a dictionary using case-insensitive key matching.
    Returns the value if found and nonempty; otherwise, returns the default.
    """
    for key, value in dictionary.items():
        if key.strip().lower() == target_key.strip().lower():
            return value.strip() if value.strip() else default
    return default

# --------------------
# Flask Routes
# --------------------

@app.route('/')
def index():
    """Home page: list available original DOCX templates."""
    # List only original templates (exclude any file whose basename starts with "sanitized_")
    template_files = [f for f in glob.glob(os.path.join(ORIGINAL_TEMPLATES_DIR, "*Template*.docx"))
                      if not os.path.basename(f).startswith("sanitized_")]
    return render_template("index.html", template_files=template_files)

@app.route('/select_template', methods=["POST"])
def select_template():
    """After selecting a template, extract its placeholders and show a form to fill them."""
    template_file = request.form.get("template_file")
    if not template_file:
        flash("No template selected.")
        return redirect(url_for("index"))
    placeholders = extract_placeholders_from_xml(template_file)
    if not placeholders:
        flash("No placeholders found in the selected template.")
        return redirect(url_for("index"))
    placeholders.sort()
    # Pass the template_file and placeholders list to the next page.
    return render_template("fill_placeholders.html",
                           template_file=template_file,
                           placeholders=placeholders)

@app.route('/generate_document', methods=["POST"])
def generate_document():
    """Process the form data, generate the document, and provide a download link."""
    template_file = request.form.get("template_file")
    if not template_file:
        flash("Template file missing.")
        return redirect(url_for("index"))
    # Re-extract placeholders (in case the template changed)
    placeholders = extract_placeholders_from_xml(template_file)
    if not placeholders:
        flash("No placeholders found in the template.")
        return redirect(url_for("index"))
    placeholders.sort()
    mapping = {}
    raw_values = {}
    context = {}
    # For each placeholder, retrieve the user-supplied value.
    # We assume that the form field names equal the sanitized placeholder,
    # i.e. by replacing non-alphanumeric characters with underscores.
    for ph in placeholders:
        key = sanitize_placeholder(ph)
        mapping[ph] = key
        value = request.form.get(key)
        raw_values[ph] = value
        context[key] = value

    # Create a sanitized version of the template.
    sanitized_template_path = sanitize_template_xml(template_file, mapping, SANITIZED_DIR)
    if not sanitized_template_path:
        flash("Error sanitizing the template.")
        return redirect(url_for("index"))
    try:
        doc = DocxTemplate(sanitized_template_path)
        doc.render(context)
    except Exception as e:
        flash("Error rendering the document.")
        return redirect(url_for("index"))
    # Build the output filename using the raw values for "Client Company Name" and "Proposal date"
    client_name = get_value_case_insensitive(raw_values, "Client Company Name", "UnknownClient")
    proposal_date = get_value_case_insensitive(raw_values, "Proposal date", "UnknownDate")
    filename_base = f"Proposal_{client_name}_{proposal_date}"
    output_filename = filename_base + ".docx"
    output_path = os.path.join(OUTPUT_DIR, output_filename)
    try:
        doc.save(output_path)
    except Exception as e:
        flash("Error saving the generated document.")
        return redirect(url_for("index"))
    # (Optional: PDF conversion can be added here if desired.)
    return render_template("result.html", output_filename=output_filename)

@app.route('/download/<filename>')
def download(filename):
    """Serve a generated file for download."""
    return send_from_directory(OUTPUT_DIR, filename, as_attachment=True)

# --------------------
# Run the App
# --------------------
if __name__ == '__main__':
    # Set debug=True for development; remove or set False for production.
    app.run(debug=True)
