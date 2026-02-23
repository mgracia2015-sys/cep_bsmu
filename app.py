import streamlit as st
import re
from datetime import datetime
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io

st.set_page_config(page_title="Перевірка статті", layout="wide")

st.title("📄 Перевірка оформлення наукової статті")

# ============================================================
# 1️⃣ ВИБІР МОВИ ТА ТИПУ СТАТТІ
# ============================================================

language = st.radio(
    "Мова статті:",
    [("Українська", "uk"), ("English", "en")],
    format_func=lambda x: x[0]
)[1]

article_type = st.radio(
    "Тип статті:",
    [
        ("Оригінальне дослідження", "original"),
        ("Клінічний випадок", "case"),
        ("Огляд літератури", "review")
    ],
    format_func=lambda x: x[0]
)[1]

uploaded_file = st.file_uploader("Завантажте файл .docx", type=["docx"])

# ============================================================
# 2️⃣ ГОЛОВНА ЛОГІКА
# ============================================================

if uploaded_file is not None:

    if st.button("🔍 Перевірити статтю"):

        report = []

        doc = Document(uploaded_file)
        paragraphs = doc.paragraphs

        report.append(f"Файл завантажено: {uploaded_file.name}")

        # =====================================================
        # ПОЛЯ
        # =====================================================

        section = doc.sections[0]

        if section.top_margin != Cm(2):
            section.top_margin = Cm(2)
            report.append("Виправлено верхнє поле на 2 см")

        if section.bottom_margin != Cm(2):
            section.bottom_margin = Cm(2)
            report.append("Виправлено нижнє поле на 2 см")

        if section.left_margin != Cm(2):
            section.left_margin = Cm(2)
            report.append("Виправлено ліве поле на 2 см")

        if section.right_margin != Cm(2):
            section.right_margin = Cm(2)
            report.append("Виправлено праве поле на 2 см")

        # =====================================================
        # ФОРМАТ ТЕКСТУ
        # =====================================================

        for paragraph in paragraphs:
            paragraph.paragraph_format.line_spacing = 1.5
            paragraph.paragraph_format.space_before = Pt(0)
            paragraph.paragraph_format.space_after = Pt(0)

            if paragraph.alignment == WD_ALIGN_PARAGRAPH.JUSTIFY:
                paragraph.paragraph_format.first_line_indent = Cm(1.25)

            for run in paragraph.runs:
                run.font.name = "Times New Roman"
                run.font.size = Pt(14)

        report.append("Перевірено формат тексту")

        # =====================================================
        # ПЕРЕВІРКА ЛІТЕРАТУРИ
        # =====================================================

        references_start = None
        references_title = None

        for i, para in enumerate(paragraphs):
            text_lower = para.text.strip().lower()
            if text_lower.startswith("список літератури") or text_lower.startswith("references"):
                references_start = i + 1
                references_title = para.text.strip()
                break

        if references_start is None:
            report.append("❌ Не знайдено розділ літератури")
        else:

            reference_paragraphs = []

            for para in paragraphs[references_start:]:
                text = para.text.strip()

                if not text:
                    continue

                # зупинка якщо контактна інформація
                if re.search(r"(author|email|correspondence|адреса|контакт)", text.lower()):
                    break

                reference_paragraphs.append(text)

            reference_count = len(reference_paragraphs)

            # ----- кількість -----

            if article_type in ["original", "case"]:
                if reference_count > 15:
                    report.append(f"⚠️ Джерел: {reference_count} (не більше 15)")
                else:
                    report.append(f"Кількість джерел: {reference_count}")

            if article_type == "review":
                if reference_count < 50:
                    report.append(f"⚠️ Джерел: {reference_count} (не менше 50)")
                else:
                    report.append(f"Кількість джерел: {reference_count}")

            # ----- Vancouver -----

            numbering_errors = False
            vancouver_errors = False
            expected_number = 1

            for ref in reference_paragraphs:

                match = re.match(r"^(\d+)[\.\)]", ref)
                if match:
                    num = int(match.group(1))
                    if num != expected_number:
                        numbering_errors = True
                    expected_number += 1
                else:
                    numbering_errors = True

                if not re.search(r"\b(19|20)\d{2}\b", ref):
                    vancouver_errors = True

            if numbering_errors:
                report.append("⚠️ Порушена нумерація у списку літератури")

            if vancouver_errors:
                report.append("⚠️ Можливе порушення Vancouver style")
            else:
                report.append("Стиль літератури виглядає коректним")

            report.append(f"Перевірено розділ: {references_title}")

        # =====================================================
        # ЗБЕРЕЖЕННЯ ФАЙЛУ
        # =====================================================

        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)

        st.success("✅ Перевірку завершено")

        # =====================================================
        # ЗВІТ
        # =====================================================

        st.subheader("📋 Звіт")

        for item in report:
            st.write("•", item)

        # =====================================================
        # КНОПКА ЗАВАНТАЖЕННЯ
        # =====================================================

        st.download_button(
            label="⬇️ Завантажити відредагований файл",
            data=buffer,
            file_name=uploaded_file.name,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )