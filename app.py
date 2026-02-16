import streamlit as st
import re
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from io import BytesIO

# --- ФУНКЦІЇ ФОРМАТУВАННЯ (БЕЗ ЗМІН) ---

def apply_base_style(paragraph, first_line=1.25, space_before=0):
    paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    paragraph.paragraph_format.line_spacing = 1.15
    paragraph.paragraph_format.first_line_indent = Cm(first_line)
    paragraph.paragraph_format.space_before = Pt(space_before)

def add_run(paragraph, text, bold=False, italic=False):
    run = paragraph.add_run(text)
    run.font.name = 'Times New Roman'
    run.font.size = Pt(12)
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

def process_abstract_block(new_doc, raw_text, terms, forbidden_word, lang_label, report, skip_warnings=False):
    clean_text = re.sub(rf'^{forbidden_word}[:\s.-]*', '', raw_text, flags=re.IGNORECASE).strip()
    if lang_label == "Українська" and len(clean_text) > 1600:
        report.append(f"⚠️ {lang_label} анотація занадто велика ({len(clean_text)} зн. при ліміті 1600).")
    
    if not skip_warnings:
        for t in terms:
            if t not in clean_text: 
                report.append(f"❌ ВІДСУТНІЙ розділ у {lang_label} анотації: {t}")

    pattern = f"({'|'.join(re.escape(t) for t in terms)})"
    parts = re.split(pattern, clean_text)
    curr_term = None
    for pt in parts:
        if not pt or not pt.strip(): continue
        if pt in terms: curr_term = pt
        else:
            p = new_doc.add_paragraph()
            if curr_term: add_run(p, curr_term, bold=True, italic=True)
            add_run(p, " " + pt.strip())
            apply_base_style(p); curr_term = None

# --- ІНТЕРФЕЙС STREAMLIT ---

st.set_page_config(page_title="Науковий Редактор", page_icon="📝")
st.title("📝 Автоматичне форматування статті")

# Вибір типу статті (Додано третій варіант)
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
            new_doc = Document()
            report = []
            paras = [p for p in doc.paragraphs if p.text.strip()]
            
            # 1. ШАПКА
            p_udc = new_doc.add_paragraph(); add_run(p_udc, paras[0].text); apply_base_style(p_udc)
            p_t_ua = new_doc.add_paragraph(); add_run(p_t_ua, paras[1].text.upper()); apply_base_style(p_t_ua)
            p_a_ua = new_doc.add_paragraph(); add_run(p_a_ua, fix_authors_metadata(paras[2].text), bold=True, italic=True); apply_base_style(p_a_ua)
            p_aff = new_doc.add_paragraph(); add_run(p_aff, paras[3].text); apply_base_style(p_aff)

            ua_kw_idx = next((i for i, p in enumerate(paras) if "Ключові слова" in p.text), -1)
            en_kw_idx = next((i for i, p in enumerate(paras) if "Key words" in p.text or "Keywords" in p.text), -1)

            # 2. АНОТАЦІЇ (Логіка залежить від типу)
            if is_review:
                ua_terms = ["Мета", "Висновки"]
                en_terms = ["Aim", "Conclusions"]
            elif is_clinical:
                ua_terms = ["Мета", "Матеріали і методи", "Результати", "Висновки"]
                en_terms = ["Aim", "Material and methods", "Results", "Conclusions"]
            else: # Оригінальне дослідження
                ua_terms = ["Мета", "Матеріали і методи", "Результати", "Висновки"]
                en_terms = ["Aim", "Material and methods", "Results", "Conclusions"]

            process_abstract_block(new_doc, " ".join([paras[i].text for i in range(4, ua_kw_idx)]), 
                                   ua_terms, "Анотація|Реферат", "Українська", report, skip_warnings=is_clinical)
            
            p_kw_ua = new_doc.add_paragraph(); add_run(p_kw_ua, "Ключові слова:", bold=True, italic=True)
            add_run(p_kw_ua, " " + paras[ua_kw_idx].text.replace("Ключові слова", "").replace(":", "").strip()); apply_base_style(p_kw_ua)

            p_t_en = new_doc.add_paragraph(); add_run(p_t_en, paras[ua_kw_idx + 1].text.upper()); apply_base_style(p_t_en)
            p_a_en = new_doc.add_paragraph(); add_run(p_a_en, fix_authors_metadata(paras[ua_kw_idx + 2].text), bold=True, italic=True); apply_base_style(p_a_en)

            process_abstract_block(new_doc, " ".join([paras[i].text for i in range(ua_kw_idx + 3, en_kw_idx)]), 
                                   en_terms, "Abstract", "Англійська", report, skip_warnings=is_clinical)
            
            p_kw_en = new_doc.add_paragraph(); add_run(p_kw_en, "Key words:", bold=True, italic=True)
            add_run(p_kw_en, " " + paras[en_kw_idx].text.replace("Key words", "").replace("Keywords", "").replace(":", "").strip()); apply_base_style(p_kw_en)

            # 3. ОСНОВНИЙ ТЕКСТ (Налаштування розділів)
            if is_review:
                sections_map = [
                    (r"^Вступ", "Вступ"), 
                    (r"^Мета", "Мета роботи"), 
                    (r"^Основна\s+частина", "Основна частина"), 
                    (r"^Висновок|^Висновки", "Висновки"),
                    (r"^Список\s*літератури|^Література|^Список\s*використаних\s*джерел", "Список літератури")
                ]
                all_req = ["Вступ", "Мета роботи", "Основна частина", "Висновки", "Список літератури"]
            elif is_clinical:
                sections_map = [(r"^Вступ", "Вступ"), (r"^Опис\s+клінічного\s+випадку", "Опис клінічного випадку"), (r"^Висновок|^Висновки", "Висновок"), (r"^Список\s*літератури|^Література|^Список\s*використаних\s*джерел", "Список літератури")]
                all_req = ["Вступ", "Опис клінічного випадку", "Висновок", "Список літератури"]
            else: # Оригінальне дослідження
                sections_map = [(r"^Вступ", "Вступ"), (r"^Мета", "Мета роботи"), (r"^Матеріали\s*(і|та)\s*методи", "Матеріали та методи дослідження"), (r"^Результати\s*та\s*їх\s*обговорення", "Результати та їх обговорення"), (r"^Висновки", "Висновки"), (r"^Перспективи\s*подальших\s*досліджень", "Перспективи подальших досліджень"), (r"^Список\s*літератури|^Література|^Список\s*використаних\s*джерел", "Список літератури")]
                all_req = ["Вступ", "Мета роботи", "Матеріали та методи дослідження", "Результати та їх обговорення", "Висновки", "Перспективи подальших досліджень", "Список літератури"]
            
            in_literature = False
            in_references = False
            found_sections = set()

            for i in range(en_kw_idx + 1, len(paras)):
                text = paras[i].text.strip()
                if re.match(r"^References[:.\s]*$", text, re.IGNORECASE):
                    p_ref = new_doc.add_paragraph(); add_run(p_ref, "References", bold=True); apply_base_style(p_ref); in_references = True; in_literature = False; continue
                matched_std = None
                for pattern, std_name in sections_map:
                    if re.match(pattern, text, re.IGNORECASE):
                        matched_std = std_name; text = re.sub(pattern + r"[:.\s-]*", "", text, count=1, flags=re.IGNORECASE).strip(); break
                
                if matched_std:
                    p_h = new_doc.add_paragraph(); add_run(p_h, matched_std, bold=True); apply_base_style(p_h, space_before=10); found_sections.add(matched_std); in_literature = (matched_std == "Список літератури")
                    if text: p_c = new_doc.add_paragraph(); add_run(p_c, format_vancouver(text) if in_literature else text); apply_base_style(p_c)
                else:
                    p_txt = new_doc.add_paragraph()
                    if in_literature or in_references:
                        vanc_text = format_vancouver(text); add_run(p_txt, vanc_text)
                    else: add_run(p_txt, text)
                    apply_base_style(p_txt)

            for r in all_req:
                if r not in found_sections: report.append(f"❌ НЕ ЗНАЙДЕНО РОЗДІЛ: {r}")

            bio = BytesIO()
            new_doc.save(bio)
            
            st.subheader("Звіт про перевірку:")
            if not report: st.success("✅ Все виглядає чудово!")
            else:
                for issue in report:
                    if "❌" in issue: st.error(issue)
                    else: st.warning(issue)

            st.download_button(label="📥 Завантажити виправлену статтю", data=bio.getvalue(), file_name=f"fixed_{uploaded_file.name}", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        except Exception as e: st.error(f"Сталася помилка при обробці: {e}")