import streamlit as st
import re
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from io import BytesIO

# --- ФУНКЦІЇ ФОРМАТУВАННЯ ---

def apply_base_style(paragraph, first_line=1.25, space_before=0):
    paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    # Початкові налаштування (будуть перезаписані фінальним циклом, але залишаємо для структури)
    paragraph.paragraph_format.first_line_indent = Cm(first_line)
    paragraph.paragraph_format.space_before = Pt(space_before)

def add_run(paragraph, text, bold=False, italic=False):
    run = paragraph.add_run(text)
    run.font.name = 'Times New Roman'
    run.font.size = Pt(14) # Одразу ставимо 14
    run.bold, run.italic = bold, italic
    run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
    return run

def fix_authors_metadata(text):
    parts = re.split(r'[,;]', text)
    fixed = []
    for p in parts:
        p = p.strip()
        res = re.sub(r'([А-ЯЁІЇЄҐA-Z][а-яёіїєґa-z]+)\s+([А-ЯЁІЇЄҐA-Z]\.\s?[А-ЯЁІЇЄҐA-Z]\.)', r'\2 \1', p)
        fixed.append(res)
    return ", ".join(fixed)

def format_vancouver(text):
    text = text.replace('"', '').replace('«', '').replace('»', '')
    text = re.sub(r'([A-ZА-Я][a-zа-я]+)\s+([A-ZА-Я])\.\s?([A-ZА-Я])\.', r'\1 \2\3', text)
    text = re.sub(r'([A-ZА-Я][a-zа-я]+)\s+([A-ZА-Я])\.', r'\1 \2', text)
    text = re.sub(r'(\d{4})[\.\s,–—]*Vol\.?\s*(\d+)[\.\s,–—]*[Nn]o\.?\s*(\d+)[\.\s,–—]*[Pp]\.?\s*(\d+)[-–—](\d+)', r'\1;\2(\3):\4-\5', text)
    text = re.sub(r'(\d{4})[\.\s,–—]*Vol\.?\s*(\d+)[\.\s,–—]*[Pp]\.?\s*(\d+)[-–—](\d+)', r'\1;\2:\3-\4', text)
    return text.strip()

# --- ІНТЕРФЕЙС ---

st.set_page_config(page_title="Науковий Редактор", page_icon="📝")
st.title("📝 Науковий Редактор (Стандарт 14 пт, 1.5)")

article_type = st.radio(
    "Оберіть тип вашої статті:",
    ("Оригінальне дослідження", "Клінічний випадок", "Огляд літератури")
)

is_clinical = (article_type == "Клінічний випадок")
is_review = (article_type == "Огляд літератури")

uploaded_file = st.file_uploader("Завантажте файл .docx", type="docx")

if uploaded_file is not None:
    if st.button("Обробити статтю"):
        try:
            doc = Document(uploaded_file)
            report = []
            paras = doc.paragraphs
            text_indices = [i for i, p in enumerate(paras) if p.text.strip()]
            
            ua_kw_idx = next((i for i, p in enumerate(paras) if "Ключові слова" in p.text), -1)
            en_kw_idx = next((i for i, p in enumerate(paras) if "Key words" in p.text or "Keywords" in p.text), -1)

            if ua_kw_idx == -1 or en_kw_idx == -1:
                st.error("❌ Не знайдено ключові слова. Перевірте оформлення.")
            else:
                # 1. ШАПКА
                paras[text_indices[0]].text = paras[text_indices[0]].text.strip()
                paras[text_indices[1]].text = paras[text_indices[1]].text.strip().upper()
                new_authors_ua = fix_authors_metadata(paras[text_indices[2]].text)
                paras[text_indices[2]].clear()
                add_run(paras[text_indices[2]], new_authors_ua, bold=True, italic=True)

                # 2. РОЗДІЛИ (Картування)
                if is_review:
                    sections_map = [(r"^Вступ", "Вступ"), (r"^Мета", "Мета роботи"), (r"^Основна\s+частина", "Основна частина"), (r"^Висновок|^Висновки", "Висновки"), (r"^Список\s*літератури|^Література|^Список\s*використаних\s*джерел", "Список літератури")]
                    all_req = ["Вступ", "Мета роботи", "Основна частина", "Висновки", "Список літератури"]
                elif is_clinical:
                    sections_map = [(r"^Вступ", "Вступ"), (r"^Опис\s+клінічного\s+випадку", "Опис клінічного випадку"), (r"^Висновок|^Висновки", "Висновок"), (r"^Список\s*літератури|^Література|^Список\s*використаних\s*джерел", "Список літератури")]
                    all_req = ["Вступ", "Опис клінічного випадку", "Висновок", "Список літератури"]
                else:
                    sections_map = [(r"^Вступ", "Вступ"), (r"^Мета", "Мета роботи"), (r"^Матеріали\s*(і|та)\s*методи", "Матеріали та методи дослідження"), (r"^Результати\s*та\s*їх\s*обговорення", "Результати та їх обговорення"), (r"^Висновки", "Висновки"), (r"^Перспективи\s*подальших\s*досліджень", "Перспективи подальших досліджень"), (r"^Список\s*літератури|^Література|^Список\s*використаних\s*джерел", "Список літератури")]
                    all_req = ["Вступ", "Мета роботи", "Матеріали та методи дослідження", "Результати та їх обговорення", "Висновки", "Перспективи подальших досліджень", "Список літератури"]

                in_literature = False
                found_sections = set()

                for i in range(en_kw_idx + 1, len(paras)):
                    p = paras[i]
                    text = p.text.strip()
                    if not text: continue

                    if re.match(r"^References[:.\s]*$", text, re.IGNORECASE):
                        p.clear(); add_run(p, "References", bold=True); in_literature = True; continue

                    matched_std = None
                    for pattern, std_name in sections_map:
                        if re.match(pattern, text, re.IGNORECASE):
                            matched_std = std_name
                            text_after = re.sub(pattern + r"[:.\s-]*", "", text, count=1, flags=re.IGNORECASE).strip()
                            break
                    
                    if matched_std:
                        p.clear(); add_run(p, matched_std, bold=True)
                        found_sections.add(matched_std)
                        in_literature = (matched_std == "Список літератури")
                        if text_after:
                            new_p = p.insert_paragraph_before(format_vancouver(text_after) if in_literature else text_after)
                    else:
                        if in_literature:
                            for run in p.runs: run.text = format_vancouver(run.text)
                        elif not any(run.text.strip() for run in p.runs): # якщо параграф пустий або тільки з картинкою
                            pass 

                # --- ФІНАЛЬНИЙ ЦИКЛ: ПРИМУСОВЕ ФОРМАТУВАННЯ ВСЬОГО ТЕКСТУ ---
                # Це застосовується до ВСІХ параграфів (текст, анотації, таблиці)
                
                def final_format(paragraph):
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                    paragraph.paragraph_format.line_spacing = 1.5
                    # Встановлюємо шрифт для кожного фрагмента тексту
                    for run in paragraph.runs:
                        run.font.name = 'Times New Roman'
                        run.font.size = Pt(14)
                        run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')

                # Форматуємо основні параграфи
                for p in doc.paragraphs:
                    final_format(p)
                
                # Форматуємо текст всередині всіх таблиць
                for table in doc.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            for p in cell.paragraphs:
                                final_format(p)

                for r in all_req:
                    if r not in found_sections: report.append(f"❌ НЕ ЗНАЙДЕНО РОЗДІЛ: {r}")

                bio = BytesIO()
                doc.save(bio)
                
                st.subheader("Звіт:")
                if not report: st.success("✅ Все ідеально!")
                else:
                    for issue in report: st.error(issue) if "❌" in issue else st.warning(issue)

                st.download_button(label="📥 Завантажити статтю (14 пт, 1.5)", data=bio.getvalue(), file_name=f"fixed_{uploaded_file.name}", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

        except Exception as e:
            st.error(f"Помилка: {e}")