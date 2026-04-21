import streamlit as st
from pptx import Presentation
from docx import Document
import PyPDF2
import os
import tempfile

# ---------------------------
# TEXT EXTRACTION
# ---------------------------
def extract_text_from_docx(file):
    doc = Document(file)
    return "\n".join([para.text for para in doc.paragraphs])

def extract_text_from_pdf(file):
    text = ""
    reader = PyPDF2.PdfReader(file)
    for page in reader.pages:
        text += page.extract_text() + "\n"
    return text

def extract_text(uploaded_file):
    if uploaded_file.name.endswith(".docx"):
        return extract_text_from_docx(uploaded_file)
    elif uploaded_file.name.endswith(".pdf"):
        return extract_text_from_pdf(uploaded_file)
    else:
        return ""

# ---------------------------
# SMART PARSER (IMPROVED)
# ---------------------------
def parse_cv(text):
    data = {
        "name": "",
        "location": "",
        "experience": "",
        "skills": "",
        "education": ""
    }

    lines = [l.strip() for l in text.split("\n") if l.strip()]

    # Name → first line
    if lines:
        data["name"] = lines[0]

    section = None

    for line in lines:
        l = line.lower()

        if "experience" in l:
            section = "experience"
            continue
        elif "skill" in l:
            section = "skills"
            continue
        elif "education" in l:
            section = "education"
            continue

        if any(loc in l for loc in ["bangalore", "india", "karnataka"]):
            data["location"] = line

        if section:
            data[section] += line + "\n"

    return data

# ---------------------------
# PPT GENERATION
# ---------------------------
def create_ppt(data, template_path):
    prs = Presentation(template_path)

    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                text = shape.text

                text = text.replace("FirstName SecondName", data["name"])
                text = text.replace("Location", data["location"])
                text = text.replace("RELEVANT EXPERIENCE", data["experience"])
                text = text.replace("RELEVANT SKILLS", data["skills"])
                text = text.replace("EDUCATION", data["education"])

                shape.text = text

    output_path = os.path.join(tempfile.gettempdir(), "generated.pptx")
    prs.save(output_path)

    return output_path

# ---------------------------
# STREAMLIT UI
# ---------------------------
st.title("📊 CV to PPT Generator")

uploaded_file = st.file_uploader("Upload CV (PDF/DOCX)", type=["pdf", "docx"])

if uploaded_file:
    st.success("File uploaded successfully!")

    text = extract_text(uploaded_file)
    data = parse_cv(text)

    ppt_path = create_ppt(data, "Deloitte Profile format 2026.pptx")

    with open(ppt_path, "rb") as f:
        st.download_button(
            label="📥 Download PPT",
            data=f,
            file_name="Generated_Profile.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )