
import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
from docx import Document
import io
import re

st.set_page_config(page_title="Smart CV Reader", layout="wide")
st.title("🚀 Smart CV Reader")
st.markdown("Upload resumes, extract key info, rank by skill match, and download as Excel.")

def extract_name(text):
    lines = text.strip().split('\n')
    ignore_words = {"resume", "curriculum vitae", "cv", "bio-data"}

    for line in lines[:10]:
        clean = line.strip().lower()
        if clean in ignore_words:
            continue
        if re.match(r'^[A-Z][a-z]+( [A-Z][a-z]+)+$', line.strip()):
            return line.strip()

    email_match = re.search(r"([a-zA-Z0-9._%+-]+)@([a-zA-Z0-9.-]+)", text)
    if email_match:
        username = email_match.group(1)
        name_parts = username.replace('.', ' ').replace('_', ' ').split()
        if len(name_parts) >= 2 and all(part.isalpha() for part in name_parts[:2]):
            return ' '.join([part.capitalize() for part in name_parts[:2]])

    for line in lines:
        clean = line.strip().lower()
        if clean not in ignore_words and len(line.strip()) > 4:
            return line.strip()

    return "Unknown"

def extract_email(text):
    match = re.search(r"[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+", text)
    return match.group() if match else "Not found"

def extract_phone(text):
    phone_regex = re.compile(r'(\+91[\s-]?)?\(?\d{3,5}\)?[\s-]?\d{3,5}[\s-]?\d{3,5}')
    match = phone_regex.search(text)
    return match.group() if match else "Not found"

def extract_skills(text, skills_list):
    found = []
    for skill in skills_list:
        if re.search(rf"\b{re.escape(skill.lower())}\b", text.lower()):
            found.append(skill)
    return found

def extract_section(text, keywords):
    for kw in keywords:
        if kw in text.lower():
            pattern = rf"{kw}.*?(?:\n\s*\n|$)"
            match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
            if match:
                return match.group().strip()
    return "Not found"

def match_score(extracted_skills, required_skills):
    matched = [skill for skill in required_skills if skill.lower() in [s.lower() for s in extracted_skills]]
    score = len(matched) / len(required_skills) if required_skills else 0
    return round(score * 100, 2), matched

def extract_text_from_pdf(file):
    doc = fitz.open(stream=file.read(), filetype="pdf")
    return "".join([page.get_text() for page in doc])

def extract_text_from_docx(file):
    doc = Document(file)
    return '\n'.join([p.text for p in doc.paragraphs])

uploaded_files = st.file_uploader("📄 Upload resumes (.pdf or .docx)", accept_multiple_files=True, type=['pdf', 'docx'])
required_skills_input = st.text_input("🛠️ Required skills (comma-separated)", "Python, SQL, Machine Learning, Data Analysis, Communication, Deep Learning, Excel, django, Html, CSS, Power BI")

if uploaded_files and required_skills_input:
    required_skills = [s.strip() for s in required_skills_input.split(',')]
    skill_keywords = ['python', 'sql', 'machine learning', 'data analysis', 'communication',
                      'deep learning', 'excel', 'django', 'html', 'css', 'power bi']

    results = []

    for file in uploaded_files:
        ext = file.name.split('.')[-1].lower()
        if ext == 'pdf':
            text = extract_text_from_pdf(file)
        elif ext == 'docx':
            text = extract_text_from_docx(file)
        else:
            st.warning(f"Unsupported file: {file.name}")
            continue

        name = extract_name(text)
        email = extract_email(text)
        phone = extract_phone(text)
        matched_skills = extract_skills(text, skill_keywords)
        education = extract_section(text, ["education", "academic", "qualifications"])
        experience = extract_section(text, ["experience", "employment", "work history"])
        score, matched = match_score(matched_skills, required_skills)

        results.append({
            "Filename": file.name,
            "Name": name,
            "Email": email,
            "Phone": phone,
            "Matched Skills": ', '.join(matched),
            "Education": education[:100] + "...",
            "Experience": experience[:100] + "...",
            "Score (%)": score
        })

    df = pd.DataFrame(results).sort_values(by="Score (%)", ascending=False)
    st.subheader("📊 Ranked Resumes")
    st.dataframe(df)

    output = io.BytesIO()
    df.to_excel(output, index=False, engine='openpyxl')
    st.download_button("⬇️ Download Excel", data=output.getvalue(), file_name="ranked_resumes.xlsx", mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
