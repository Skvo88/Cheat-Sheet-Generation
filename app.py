import streamlit as st
import re
from docx import Document
from docx.shared import Cm, Pt
from docx.enum.table import WD_ROW_HEIGHT_RULE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from io import BytesIO

# --- ЛОГИКА ОБРАБОТКИ (Твой оригинальный алгоритм) ---

def set_cell_margins(cell, top, bottom, left, right):
    tcPr = cell._element.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for margin, val in zip(['top', 'left', 'bottom', 'right'], [top, left, bottom, right]):
        node = OxmlElement(f'w:{margin}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def get_word_stream(text_block):
    stream = []
    lines = text_block.split('\n')
    newline_weight = 35
    for line in lines:
        line = line.strip()
        if not line: continue
        
        is_term_line = False
        colon_idx = line.find(':')
        if line.startswith("СОДЕРЖАНИЕ:") or (re.match(r'^\d+', line) and 0 < colon_idx < 200):
            is_term_line = True
        
        if is_term_line:
            parts = line.split(':', 1)
            term_name = parts[0] + ":"
            for word in term_name.split():
                stream.append((word + " ", True))
            if len(parts) > 1:
                for word in parts[1].split():
                    stream.append((word + " ", False))
        else:
            for word in line.split():
                stream.append((word + " ", False))
        
        stream.append(("\n", False))
    return stream

def build_docx_in_memory(text_content, max_chars, margin_val, line_spacing):
    word_stream = get_word_stream(text_content)
    newline_weight = 35
    
    cells_data = []
    current_cell = []
    current_count = 0
    
    for word, is_bold in word_stream:
        weight = newline_weight if word == "\n" else len(word)
        if current_count + weight > max_chars:
            cells_data.append(current_cell)
            current_cell = []
            current_count = 0
            if word == "\n": continue 
        current_cell.append((word, is_bold))
        current_count += weight
        
    if current_cell:
        cells_data.append(current_cell)
        
    while len(cells_data) % 12 != 0:
        cells_data.append([])

    doc = Document()
    section = doc.sections[0]
    section.page_width, section.page_height = Cm(21.0), Cm(29.7)
    section.left_margin = section.right_margin = section.top_margin = section.bottom_margin = Cm(0.5)

    num_rows = len(cells_data) // 4
    table = doc.add_table(rows=num_rows, cols=4)
    table.style = 'Table Grid'
    
    for row in table.rows:
        row.height = Cm(9.2)
        row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY

    for idx, data in enumerate(cells_data):
        cell = table.cell(idx // 4, idx % 4)
        set_cell_margins(cell, top=margin_val, bottom=margin_val, left=margin_val, right=margin_val)
        p = cell.paragraphs[0]
        p.paragraph_format.line_spacing = line_spacing
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.space_before = Pt(0)
        
        for word, is_bold in data:
            if word == "\n":
                p.add_run().add_break()
            else:
                run = p.add_run(word)
                run.bold = is_bold
                run.font.name = 'Times New Roman'
                run.font.size = Pt(5)
    
    target = BytesIO()
    doc.save(target)
    target.seek(0)
    return target

# --- ИНТЕРФЕЙС STREAMLIT ---

st.set_page_config(page_title="Генератор карточек ЕГЭ", page_icon="📝")

# Боковая панель
with st.sidebar:
    st.header("⚙️ Настройки верстки")
    max_c = st.slider("Символов в карточке", 1500, 3500, 2700, help="Чем меньше, тем раньше текст перейдет в новую карточку")
    marg = st.slider("Поля (внутренний отступ)", 0, 200, 75, help="Отступ текста от границ рамки")
    l_space = st.slider("Межстрочный интервал", 0.7, 1.2, 0.82, step=0.01)
    st.info("Настройки применяются мгновенно при нажатии кнопки генерации.")

st.title("🖨 Генератор карточек для ЕГЭ")
st.write("Вставьте текст ниже. Скрипт разделит его по фразе `НИЗКАЯ ВЕРОЯТНОСТЬ ПОПАДАНИЯ` и создаст Word-файлы.")

txt_input = st.text_area("Текст для обработки:", height=400, placeholder="Вставьте сюда содержимое вашего input.txt...")

if st.button("🚀 Сгенерировать файлы"):
    if not txt_input.strip():
        st.warning("Пожалуйста, вставьте текст!")
    else:
        parts = re.split(r'НИЗКАЯ ВЕРОЯТНОСТЬ ПОПАДАНИЯ', txt_input)
        
        # Высокая вероятность
        st.subheader("Результаты:")
        file_high = build_docx_in_memory(parts[0], max_c, marg, l_space)
        st.download_button(
            label="📥 Скачать: Высокая вероятность",
            data=file_high,
            file_name="Cards_High.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        
        # Низкая вероятность
        if len(parts) > 1:
            file_low = build_docx_in_memory(parts[1], max_c, marg, l_space)
            st.download_button(
                label="📥 Скачать: Низкая вероятность",
                data=file_low,
                file_name="Cards_Low.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        st.success("Готово! Нажмите на кнопки выше, чтобы сохранить файлы.")