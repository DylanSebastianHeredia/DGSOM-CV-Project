# Sebastian Heredia
# DyHeredia@mednet.ucla.edu
# July 3, 2025

# Perhaps a more viable solution for 15-20 page documents (less data to process).

import streamlit as st              # Loads Streamlit GUI
from docx import Document           # For working with .docx documents
from docx.shared import Pt          # For setting font size
from docx.shared import Inches      # To create tables and adjust horizontal spacing
from io import BytesIO              # To hold input/output memory for download
from docx.enum.text import WD_ALIGN_PARAGRAPH       # To align text
from docx.enum.text import WD_LINE_SPACING          # To adjust vertical line spacing 
from docx.enum.text import WD_BREAK                 # Insert page break
from docx.oxml.ns import qn         # To enable .font.name update
from docx.oxml import OxmlElement   # Used to manually adjust width
# import json                       # Saving so that page refresh doesn't lose entries data
# import os                         # For file management

# Initialize all_cvs (multi-CV manager) and current cv name
if "all_cvs" not in st.session_state:
    st.session_state.all_cvs = {}
if "current_cv" not in st.session_state:
    st.session_state.current_cv = "Default_CV"

# If first load, create Default_CV
if st.session_state.current_cv not in st.session_state.all_cvs:
    st.session_state.all_cvs[st.session_state.current_cv] = {
        "BUSINESS INFORMATION": [
            {"name": "", "position": "", "order": 1, "company_name": "",
            "business_address": "", "business_phone": "", "email": "", },
        ],
        "EDUCATION": [
            {"degree": "", "year": "", "order": 1, "school": ""},      # Dictionary
        ]
    }

# Make cv_data point to the currently selected CV
st.session_state.cv_data = st.session_state.all_cvs[st.session_state.current_cv]

st.title("UCLA DGSOM CV Formatter")     # Website title

# CV selection + creation (Sidebar Navigation)
st.sidebar.header("CV Manager")
cv_list = list(st.session_state.all_cvs.keys())
selected_cv = st.sidebar.selectbox("Select CV", cv_list, index=cv_list.index(st.session_state.current_cv))
# Creates selection drop down with title Select CV

if selected_cv != st.session_state.current_cv:   
    st.session_state.current_cv = selected_cv       # Setting the current CV to the selected CV
    st.session_state.cv_data = st.session_state.all_cvs[selected_cv]
    st.rerun()

new_cv_name = st.sidebar.text_input("New CV name (no spaces):")
if st.sidebar.button("Create New CV"):
    if new_cv_name and new_cv_name not in st.session_state.all_cvs:
        st.session_state.all_cvs[new_cv_name] = {}
        st.session_state.current_cv = new_cv_name
        st.session_state.cv_data = st.session_state.all_cvs[new_cv_name]
        st.success(f"New CV '{new_cv_name}' created.")
        st.rerun()
    else:
        st.sidebar.warning("Enter a unique CV name.")      # Ensure no duplicates!

st.header("BUSINESS INFORMATION")

# Iterate over BUSINESS INFORMATION textboxes
for i, entry in enumerate(st.session_state.cv_data.get("BUSINESS INFORMATION", [])):
    name_display = entry["name"].strip() or ""
    with st.expander(f"Entry {i+1}: {name_display}"):
        with st.form(f"biz_form_{st.session_state.current_cv}_{i}"):
            new_name = st.text_input("Name", value=entry.get("name", ""), key=f"name_{st.session_state.current_cv}_{i}")
            new_position = st.text_input("Degree", value=entry.get("position", ""), key=f"position_{st.session_state.current_cv}_{i}")
            new_company = st.text_input("Company Name", value=entry.get("busineess_address", ""), key=f"company_{st.session_state.current_cv}_{i}")
            new_address = st.text_input("Business Address", value=entry.get("business_address", ""), key=f"address_{st.session_state.current_cv}_{i}")
            new_phone = st.text_input("Business Phone", value=entry.get("business_phone", ""), key=f"phone_{st.session_state.current_cv}_{i}")
            new_email = st.text_input("Email", value=entry.get("email", ""), key=f"email_{st.session_state.current_cv}_{i}")
            submitted = st.form_submit_button("Update Entry")
            if submitted:
                entry["name"] = new_name
                entry["position"] = new_position
                entry["company_name"] = new_company
                entry["business_address"] = new_address
                entry["business_phone"] = new_phone
                entry["email"] = new_email
                st.rerun()  # Refresh to show updated info
        
st.header("EDUCATION")      # Header in GUI

# EDUCATION: Incrementation + User input feature + Delete entry and confirmation for
for i, entry in enumerate(st.session_state.cv_data["EDUCATION"]):
    degree_display = entry["degree"].strip() or ""
    year_display = entry["year"].strip() or ""
    with st.expander(f"Entry {i+1}: {year_display} {degree_display}"):
        with st.form(f"form_{st.session_state.current_cv}_{i}"):
            new_degree = st.text_input("Degree", value=entry.get("degree", ""), key=f"degree_{st.session_state.current_cv}_degree_{i}")
            new_year = st.text_input("Year (Ex: 2000-2004)", value=entry.get("year", ""), key=f"year_{st.session_state.current_cv}_year_{i}")
            new_school = st.text_input("University Name", value=entry.get("school", ""), key=f"school_{st.session_state.current_cv}_school_{i}")
            submitted = st.form_submit_button("Update Entry")

            # Initialize delete confirmation flag if not present
            flag_key = f"delete_confirm_{st.session_state.current_cv}_{i}"
            if flag_key not in st.session_state:
                st.session_state[flag_key] = False

            if not st.session_state[flag_key]:
                delete_clicked = st.form_submit_button("Delete Entry")
                if delete_clicked:
                    st.session_state[flag_key] = True
                    st.rerun()
            else:
                st.write("Are you sure you want to delete this entry?")
                confirm = st.form_submit_button("Yes, delete")
                cancel = st.form_submit_button("Cancel")
                if confirm:
                    st.session_state.cv_data["EDUCATION"].pop(i)
                    st.session_state.pop(flag_key, None)
                    st.rerun()
                if cancel:
                    st.session_state[flag_key] = False
                    st.rerun()

            if submitted:
                entry["degree"] = new_degree
                entry["year"] = new_year
                entry["school"] = new_school
                st.rerun()

# For new Entry
if st.button("Add Entry", key="add_entry_EDUCATION"):
    st.session_state.cv_data.setdefault("EDUCATION", []).append({
        "degree": "",
        "year": "",
        "order": len(st.session_state.cv_data.get("EDUCATION", [])) + 1
    })
    st.rerun()

st.header("LECTURES AND PRESENTATIONS")



st.header("BIBLIOGRAPHY")

# Preview button

# Generate .docx and download button
def generate_docx(data):
    doc = Document()
    
    # Try to pull first name entry (or default if missing)
    biz_info = data.get("BUSINESS INFORMATION", [])
    name = biz_info[0]["name"].strip()
    degree = biz_info[0].get("position", "").strip() 

    name_line = f"{name}, {degree}".strip().rstrip(",")   

    # First line: Name, Degree
    name_para = doc.add_paragraph()
    name_run = name_para.add_run(name_line)
    name_run.font.name = 'Times New Roman'
    name_run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
    name_run.font.size = Pt(12)
    name_run.font.bold = True
    name_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    name_para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    name_para.paragraph_format.space_before = Pt(0)
    name_para.paragraph_format.space_after = Pt(0)

    # Second line: Curriculum Vitae
    title_para = doc.add_paragraph()
    title_run = title_para.add_run("Curriculum Vitae")
    title_run.font.name = 'Times New Roman'
    title_run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
    title_run.font.size = Pt(12)
    title_run.font.bold = True
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    title_para.paragraph_format.space_before = Pt(0)

    # BUSINESS INFORMATION section header
    biz_para = doc.add_paragraph()
    biz_run = biz_para.add_run("BUSINESS INFORMATION")
    biz_run.font.name = 'Times New Roman'
    biz_run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
    biz_run.font.size = Pt(12)
    biz_run.font.bold = True
    biz_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    biz_para.paragraph_format.space_after = Pt(3)
    biz_para.paragraph_format.space_before = Pt(16)

    # BUSINESS INFORMATION entries
    for entry in sorted(data.get("BUSINESS INFORMATION", []), key=lambda x: x["order"]):
        company = entry.get("company_name", "").strip() or "Company Missing"
        address = entry.get("business_address", "").strip() or "Address Missing"
        phone = entry.get("business_phone", "").strip() or "Phone Missing"
        email = entry.get("email", "").strip() or "Email Missing"

        para = doc.add_paragraph()
        para.paragraph_format.left_indent = Pt(18)

        lines = [company, address, f"Phone: {phone}", f"Email: {email}"]

        for i, line in enumerate(lines):
            run = para.add_run(line)
            run.font.name = 'Times New Roman'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
            run.font.size = Pt(12)
            if i < len(lines) - 1:
                run.add_break()

        para.alignment = WD_ALIGN_PARAGRAPH.LEFT
        para.paragraph_format.space_before= Pt(0)
        para.paragraph_format.space_after = Pt(0)
        
    # EDUCATION section header
    edu_para = doc.add_paragraph()
    edu_run = edu_para.add_run("EDUCATION")
    edu_run.font.name = 'Times New Roman'
    edu_run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
    edu_run.font.size = Pt(12)
    edu_run.font.bold = True
    edu_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    edu_para.paragraph_format.space_after = Pt(3)
    edu_para.paragraph_format.space_before = Pt(16)

    # EDUCATION entries
    # Create table with 3 columns: Year, Degree, School
    table = doc.add_table(rows=0, cols=3)
    para.paragraph_format.left_indent = Pt(18)          # Match Business Info indentation
    
    # Set column width in inches
    col_widths = [1.15, 0.5, 4.85]  # 8.5" - (1" margin x 2) = 6.5" usable

   # Add data rows for EDUCATION
    for entry in sorted(data["EDUCATION"], key=lambda x: x["order"]):
        year = entry["year"].strip()
        degree = entry["degree"].strip()
        school = entry["school"].strip()

        # Skip completely blank entries
        if not any([year, degree, school]):
            continue

        year = year or "Year (missing)"
        degree = degree or "Degree (missing)"
        school = school or "School (missing)"

        row_cells = table.add_row().cells
        row_data = [year, degree, school]

        for i, (cell, text) in enumerate(zip(row_cells, row_data)):
            # Set cell text
            paragraph = cell.paragraphs[0]
            run = paragraph.add_run(text)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(12)
            paragraph.paragraph_format.space_after = Pt(0)
            paragraph.paragraph_format.space_before = Pt(0)

            # Set column width via XML
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()
            tcW = OxmlElement('w:tcW')
            tcW.set(qn('w:w'), str(int(col_widths[i] * 1440)))  # twips: 1 inch = 1440 twips
            tcW.set(qn('w:type'), 'dxa')
            tcPr.append(tcW)

    # Apply left indentation to all paragraphs in the table
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                paragraph.paragraph_format.left_indent = Pt(16)

    # Page break before BIBLIOGRAPHY page
    page_break_para = doc.add_paragraph()
    page_break_run = page_break_para.add_run()
    page_break_run.add_break(WD_BREAK.PAGE)

    # BIBLIOGRAPHY section header
    bib_para = doc.add_paragraph()
    bib_run = bib_para.add_run("BIBLIOGRAPHY")
    bib_run.font.name = 'Times New Roman'
    bib_run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
    bib_run.font.size = Pt(12)
    bib_run.font.bold = True
    bib_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    bib_para.paragraph_format.space_after = Pt(3)
    bib_para.paragraph_format.space_before = Pt(16)

    # Save to buffer
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)  # Rewind the buffer to the beginning
    return buffer

if st.button("Export as .docx"):
    docx_buffer = generate_docx(st.session_state.cv_data)
    st.download_button(
        label="Download CV",
        data=docx_buffer,
        file_name=f"{st.session_state.current_cv}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
)
