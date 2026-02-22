# app.py
import streamlit as st
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO
import re
from datetime import datetime

st.set_page_config(page_title="Перевірка статті", layout="wide")
st.title("Перевірка наукової статті")

# ========================
# 1️⃣ Вибір мови та типу статті
# ========================
language = st.radio("Мова статті:", ('Українська', 'English'))
article_type = st.radio("Тип статті:", ('Оригінальне дослідження', 'Клінічний випадок', 'Огляд літератури'))

uploaded_file = st.file_uploader("Завантажте DOCX файл статті", type=["docx"])

# ========================
# 2️⃣ Обробка файлу
# ========================
if uploaded_file is not None:
    doc = Document(uploaded_file)
    paragraphs = doc.paragraphs
    report = []

    # -------------------
    # Поля сторінки
    # -------------------
    section = doc.sections[0]
    for attr, value in [('top_margin', 2), ('bottom_margin', 2), ('left_margin', 2), ('right_margin', 2)]:
        if getattr(section, attr) != Cm(value):
            setattr(section, attr, Cm(value))
            report.append(f"Виправлено поле {attr.replace('_',' ')} на {value} см")

    # -------------------
    # Формат тексту
    # -------------------
    for paragraph in paragraphs:
        para_format = paragraph.paragraph_format
        para_format.line_spacing = 1.5
        para_format.space_before = Pt(0)
        para_format.space_after = Pt(0)
        if paragraph.alignment == WD_ALIGN_PARAGRAPH.JUSTIFY:
            para_format.first_line_indent = Cm(1.25)
        for run in paragraph.runs:
            run.font.name = "Times New Roman"
            run.font.size = Pt(14)
    report.append("Формат тексту перевірено та виправлено")

    # -------------------
    # УДК, назва та автори
    # -------------------
    if len(paragraphs) >= 1:
        first = paragraphs[0]
        if not first.text.startswith("УДК"):
            first.text = "УДК 000.00"
            report.append("Додано УДК")
        for run in first.runs:
            run.font.bold = True
        report.append("УДК перевірено/виправлено")

    if len(paragraphs) >= 2:
        # Назва статті
        title_para = paragraphs[1]
        title_para.text = title_para.text.replace("\n"," ").strip().upper()
        for run in title_para.runs:
            run.font.bold = True
        report.append("Назва статті перевірена та приведена до формату")

    if len(paragraphs) >= 3:
        # Автори
        authors_para = paragraphs[2]
        authors_list = authors_para.text.split(',')
        new_authors = []
        for author in authors_list:
            author = author.strip()
            parts = author.split()
            if len(parts) >= 2:
                if parts[0].endswith("."):
                    initials, surname = parts[0], parts[1]
                    rest = " ".join(parts[2:])
                else:
                    surname, initials = parts[0], parts[1]
                    rest = " ".join(parts[2:])
                text = f"{initials} {surname}"
                if rest:
                    text += f" {rest}"
                new_authors.append(text)
            else:
                new_authors.append(author)
        authors_para.text = ", ".join(new_authors)
        for run in authors_para.runs:
            run.font.bold = True
            run.font.italic = True
        report.append("Автори відформатовані (жирний + курсив)")

    # -------------------
    # Афіліація та анотація
    # -------------------
    max_aff = 0
    numbers = re.findall(r'\d+', paragraphs[2].text if len(paragraphs)>=3 else "")
    if numbers:
        max_aff = max([int(n) for n in numbers])
    aff_start = 3
    aff_end = aff_start + max_aff if max_aff>0 else aff_start+1
    for para in paragraphs[aff_start:aff_end]:
        for run in para.runs:
            run.font.bold = False
            run.font.italic = False
    report.append(f"Афіліація авторів перевірена ({aff_end - aff_start} рядків)")

    # Анотація
    abstract_start = aff_end
    abstract_end = abstract_start
    for i in range(abstract_start, len(paragraphs)):
        if any(k.lower() in paragraphs[i].text.lower() for k in ["ключові слова","keywords"]):
            abstract_end = i+1
            break
    abstract_paras = paragraphs[abstract_start:abstract_end]
    for para in abstract_paras:
        para.paragraph_format.first_line_indent = None
        for run in para.runs:
            bold_prev = run.font.bold
            run.font.italic = True
            if bold_prev is not None:
                run.font.bold = bold_prev
    abstract_len = sum(len(p.text) for p in abstract_paras)
    if abstract_len < 1800 or abstract_len > 2500:
        report.append(f"⚠️ Довжина анотації {abstract_len} символів (рекомендовано 1800–2500)")
    report.append("Анотація перевірена та відформатована (курсив)")

    # ========================
    # Література (Vancouver + кількість)
    # ========================
    refs_start = None
    search_titles = ["список літератури"] if language=="Українська" else ["references"]
    for i, para in enumerate(paragraphs):
        if any(para.text.strip().lower().startswith(title) for title in search_titles):
            refs_start = i+1
            refs_title = paragraphs[i].text.strip()
            break
    if refs_start is None:
        report.append("❌ Не знайдено розділ літератури")
    else:
        ref_paras = []
        for para in paragraphs[refs_start:]:
            text = para.text.strip()
            if not text:
                continue
            if re.match(r"^\d+\.\s", text):
                ref_paras.append(para)
            else:
                break
        ref_count = len(ref_paras)
        numbering_errors = any(not re.match(r"^\d+\.\s", p.text.strip()) for p in ref_paras)
        vancouver_errors = any(not re.search(r"\b(19|20)\d{2}\b", p.text.strip()) or not re.match(r"^\d+\.\s", p.text.strip()) for p in ref_paras)
        if article_type in ["Оригінальне дослідження","Клінічний випадок"]:
            if ref_count>15:
                report.append(f"⚠️ Кількість джерел {ref_count} (не більше 15)")
            else:
                report.append(f"Кількість джерел: {ref_count} (норма)")
        elif article_type=="Огляд літератури":
            if ref_count<50:
                report.append(f"⚠️ Кількість джерел {ref_count} (не менше 50)")
            else:
                report.append(f"Кількість джерел: {ref_count} (відповідає вимогам)")
        if numbering_errors:
            report.append("⚠️ Порушена послідовність нумерації")
        if vancouver_errors:
            report.append("⚠️ Список літератури може не відповідати Vancouver style")
        else:
            report.append("Стиль літератури відповідає базовим вимогам Vancouver")
        report.append(f"Перевірено розділ: {refs_title}")

    # ========================
    # Вивід результатів
    # ========================
    st.header("Звіт про перевірку")
    for r in report:
        st.write(r)

    # ========================
    # Завантаження відформатованого файлу
    # ========================
    output = BytesIO()
    doc.save(output)
    output.seek(0)
    st.download_button(
        label="Завантажити відформатований DOCX",
        data=output,
        file_name="відформатована_стаття.docx"
    )