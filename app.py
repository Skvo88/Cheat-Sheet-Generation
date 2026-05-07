import streamlit as st
import re
import copy
from io import BytesIO
from docx import Document
from docx.shared import Cm, Pt, Mm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ROW_HEIGHT_RULE, WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# ==========================================
# 1. НИЗКОУРОВНЕВАЯ РАБОТА С XML (СЕТКА И РАМКИ)
# ==========================================

def set_cell_margins(cell, top_mm, bottom_mm, left_mm, right_mm):
    """Настраивает внутренние отступы ячейки."""
    def mm_to_dxa(mm): return int((mm / 25.4) * 1440)
    tcPr = cell._element.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for margin, val in zip(['top', 'bottom', 'left', 'right'], [top_mm, bottom_mm, left_mm, right_mm]):
        node = OxmlElement(f'w:{margin}')
        node.set(qn('w:w'), str(mm_to_dxa(val)))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_table_borders(table, outer_style, inner_style, border_size_pt=1, border_color="000000"):
    """Устанавливает стили границ таблицы."""
    if outer_style == "nil" and inner_style == "nil": return
    
    sz_val = str(int(border_size_pt * 8))
    tblPr = table._element.tblPr
    tblBorders = OxmlElement('w:tblBorders')

    for border_name in['top', 'left', 'bottom', 'right']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), outer_style)
        border.set(qn('w:sz'), sz_val)
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), border_color)
        tblBorders.append(border)

    for border_name in ['insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), inner_style)
        border.set(qn('w:sz'), sz_val)
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), border_color)
        tblBorders.append(border)

    tblPr.append(tblBorders)

def hex_to_rgb(hex_color):
    """Конвертация цвета для python-docx."""
    hex_color = hex_color.lstrip('#')
    return RGBColor(int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16))

# ==========================================
# 2. ЯДРО СОХРАНЕНИЯ ФОРМАТИРОВАНИЯ (МАГИЯ КЛОНИРОВАНИЯ)
# ==========================================

def copy_run_formatting(source_run, target_run):
    """
    ПОЛНОСТЬЮ копирует всё форматирование из исходного куска текста в новый.
    Переносит цвет, фон, маркер, шрифт, размер, подчеркивания и т.д.
    Основано на глубоком копировании XML-узла <w:rPr>.
    """
    if source_run is None or source_run._r.rPr is None:
        return
    
    # Удаляем стандартный стиль нового текста, если он есть
    if target_run._r.rPr is not None:
        target_run._r.remove(target_run._r.rPr)
        
    # Клонируем свойства из исходника
    target_run._r.append(copy.deepcopy(source_run._r.rPr))

def apply_smart_rules(target_run, text, rules):
    """
    Применяет пользовательские настройки поверх оригинального форматирования.
    """
    text_stripped = text.strip()
    if not text_stripped: return

    # Правило 1: КАПС
    if text_stripped.isupper() and any(c.isalpha() for c in text_stripped):
        if rules['caps_bold']: target_run.bold = True
        if rules['caps_italic']: target_run.italic = True
        if rules['caps_color_on']: target_run.font.color.rgb = hex_to_rgb(rules['caps_color'])

    # Правило 2: Нумерация и скобки (например "1.", "1)", "а)")
    if re.match(r'^[\dа-яА-Яa-zA-Z]+[.)]$', text_stripped):
        if rules['list_bold']: target_run.bold = True
        if rules['list_italic']: target_run.italic = True
        if rules['list_color_on']: target_run.font.color.rgb = hex_to_rgb(rules['list_color'])

# ==========================================
# 3. ПАРСЕРЫ ФАЙЛОВ И ТЕКСТА
# ==========================================

def parse_docx_stream(file_bytes):
    """
    Читает документ Word и возвращает поток: (Текст, Оригинальный Объект Run).
    """
    stream =[]
    doc = Document(file_bytes)
    
    for p in doc.paragraphs:
        if not p.text.strip():
            stream.append(("\n", None))
            continue
            
        for run in p.runs:
            text = run.text
            if not text: continue
            
            # Разбиваем по пробелам, чтобы карточки могли переносить текст по словам,
            # но привязываем к КАЖДОМУ слову его оригинальный Run для сохранения дизайна
            parts = re.split(r'( )', text)
            for part in parts:
                if part:
                    stream.append((part, run))
                    
        stream.append(("\n", None))
    return stream

def parse_raw_text_stream(text_block):
    """Парсер для сырого текста (без исходного форматирования)."""
    stream =[]
    lines = text_block.split('\n')
    for line in lines:
        if not line.strip():
            stream.append(("\n", None))
            continue
        parts = re.split(r'( )', line)
        for part in parts:
            if part: stream.append((part, None))
        stream.append(("\n", None))
    return stream

# ==========================================
# 4. ДВИЖОК СБОРКИ КАРТОЧЕК (BUILDER)
# ==========================================

def build_professional_docx(source_data, config, is_file=False):
    """Собирает документ с идеальной разбивкой на карточки."""
    
    if is_file:
        word_stream = parse_docx_stream(source_data)
    else:
        word_stream = parse_raw_text_stream(source_data)
        
    # Разделение потока на ячейки (карточки)
    cells_data = []
    current_cell =[]
    current_count = 0
    newline_weight = 40 # Штраф за перенос строки
    
    for text_chunk, original_run in word_stream:
        weight = newline_weight if text_chunk == "\n" else len(text_chunk)
        
        if current_count + weight > config['max_chars']:
            cells_data.append(current_cell)
            current_cell = []
            current_count = 0
            if text_chunk in ["\n", " "]: continue # Убираем висячие пробелы в начале карточки
            
        current_cell.append((text_chunk, original_run))
        current_count += weight
        
    if current_cell:
        cells_data.append(current_cell)
        
    # Добиваем пустыми ячейками для ровной таблицы
    while len(cells_data) % config['cols_num'] != 0:
        cells_data.append([])

    # Инициализация Word
    doc = Document()
    section = doc.sections[0]
    section.page_width = Mm(210)
    section.page_height = Mm(297)
    for margin in['left_margin', 'right_margin', 'top_margin', 'bottom_margin']:
        setattr(section, margin, Mm(config['page_margin']))

    num_rows = len(cells_data) // config['cols_num']
    table = doc.add_table(rows=num_rows, cols=config['cols_num'])
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False 
    
    set_table_borders(table, config['border_outer'], config['border_inner'], config['border_size'])
    
    # Заполнение таблицы
    for row_idx, row in enumerate(table.rows):
        row.height = Mm(config['row_height'])
        row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
        
        for col_idx, cell in enumerate(row.cells):
            cell.width = Mm(config['col_width'])
            set_cell_margins(cell, config['pad_mm'], config['pad_mm'], config['pad_mm'], config['pad_mm'])
            
            data_idx = row_idx * config['cols_num'] + col_idx
            if data_idx >= len(cells_data): continue
            
            data = cells_data[data_idx]
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY if config['justify'] else WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.line_spacing = config['line_spacing']
            p.paragraph_format.space_after = Pt(0)
            
            # ГЛОБАЛЬНЫЙ ШРИФТ ДЛЯ ВСЕЙ ЯЧЕЙКИ
            if config['apply_global_font']:
                p.style.font.name = config['font_name']
                p.style.font.size = Pt(config['font_size'])
            
            if not data: continue
            
            for text_chunk, original_run in data:
                if text_chunk == "\n":
                    p.add_run().add_break()
                else:
                    new_run = p.add_run(text_chunk)
                    
                    # 1. КОПИРУЕМ ИСХОДНОЕ ФОРМАТИРОВАНИЕ (Цвета, маркеры, жирность из файла)
                    if original_run is not None:
                        copy_run_formatting(original_run, new_run)
                    elif config['apply_global_font']:
                        # Если это сырой текст, применяем глобальный шрифт
                        new_run.font.name = config['font_name']
                        new_run.font.size = Pt(config['font_size'])
                        
                    # 2. ПРИМЕНЯЕМ УМНЫЕ ПРАВИЛА (Поверх исходника)
                    apply_smart_rules(new_run, text_chunk, config['rules'])
    
    target = BytesIO()
    doc.save(target)
    target.seek(0)
    return target

# ==========================================
# 5. ИНТЕРФЕЙС STREAMLIT
# ==========================================

st.set_page_config(page_title="Ultimate Карточки", page_icon="🖨️", layout="wide")

st.markdown("""
<style>
    .main-header { font-size: 2.2rem; font-weight: 800; color: #1E1E1E; margin-bottom: 5px; }
    .sub-header { font-size: 1.1rem; color: #666; margin-bottom: 25px; }
    .st-expander { border-radius: 8px !important; border: 1px solid #ddd !important; }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-header">🖨️ Pro Генератор Шпор (Сохранение Формата)</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Умный перенос всех выделений (цвет, маркер) из Word с возможностью добавить свои правила.</div>', unsafe_allow_html=True)

col_settings, col_main = st.columns([1.3, 2.2], gap="large")

# ----------------- БОКОВОЕ МЕНЮ (НАСТРОЙКИ) -----------------
with col_settings:
    st.write("### ⚙️ Правила и Настройки")
    
    # 1. Дополнительные правила
    with st.expander("✨ Умные правила (Добавочные)", expanded=True):
        st.info("Эти правила применятся **ПОВЕРХ** вашего текущего форматирования в документе.")
        
        st.markdown("**Что делать со словами КАПСОМ?**")
        c1, c2, c3 = st.columns(3)
        caps_bold = c1.checkbox("Жирный", value=False, key="cb")
        caps_italic = c2.checkbox("Курсив", value=False, key="ci")
        caps_color_on = c3.checkbox("Цвет", value=False, key="cc_on")
        caps_color = st.color_picker("Цвет КАПСА", "#FF0000", label_visibility="collapsed") if caps_color_on else "#000000"
            
        st.markdown("**Что делать со списками (1. 2) а))?**")
        l1, l2, l3 = st.columns(3)
        list_bold = l1.checkbox("Жирный", value=False, key="lb")
        list_italic = l2.checkbox("Курсив", value=False, key="li")
        list_color_on = l3.checkbox("Цвет", value=False, key="lc_on")
        list_color = st.color_picker("Цвет списков", "#0000FF", label_visibility="collapsed") if list_color_on else "#000000"

    # 2. Глобальный шрифт
    with st.expander("🔤 Глобальный шрифт", expanded=False):
        apply_global_font = st.checkbox("Принудительно изменить шрифт везде", value=False, help="Если выключено, скрипт сохранит шрифты из вашего исходного Word-файла.")
        font_preset = st.selectbox("Шрифт",["Arial", "Times New Roman", "Calibri", "Свой..."], disabled=not apply_global_font)
        final_font = st.text_input("Название (напр. Montserrat):", "Montserrat", disabled=not apply_global_font) if font_preset == "Свой..." else font_preset
        font_size = st.number_input("Кегль (pt)", min_value=3.0, max_value=14.0, value=5.0, step=0.5, disabled=not apply_global_font)
        
        line_spacing = st.number_input("Межстрочный интервал", min_value=0.5, max_value=2.0, value=0.92, step=0.01)
        justify = st.checkbox("Выравнивание по ширине", value=True)

    # 3. Сетка и границы
    with st.expander("🔲 Сетка и Карточки", expanded=False):
        cols_num = st.number_input("Колонок", min_value=1, max_value=10, value=4, step=1)
        col_width = st.number_input("Ширина (мм)", min_value=10.0, max_value=200.0, value=51.0, step=0.5)
        row_height = st.number_input("Высота (мм)", min_value=10.0, max_value=290.0, value=92.3, step=0.5)
        max_chars = st.slider("Вместимость (символов)", 500, 4000, 2600, step=50)
        
        st.markdown("---")
        border_styles = ["single", "dashed", "dotted", "double", "nil"]
        b1, b2 = st.columns(2)
        with b1: border_outer = st.selectbox("Рамка снаружи", border_styles, index=0)
        with b2: border_inner = st.selectbox("Внутренняя сетка", border_styles, index=1)
        border_size = st.number_input("Толщина линий (pt)", min_value=0.1, max_value=5.0, value=1.0, step=0.5)
        pad_mm = st.number_input("Отступ текста от краев (мм)", min_value=0.0, max_value=10.0, value=1.5, step=0.1)

    # Упаковка конфига
    config = {
        'page_margin': 10.0, 'col_width': col_width, 'row_height': row_height, 'cols_num': cols_num,
        'apply_global_font': apply_global_font, 'font_name': final_font, 'font_size': font_size,
        'justify': justify, 'line_spacing': line_spacing, 'max_chars': max_chars,
        'pad_mm': pad_mm, 'border_outer': border_outer, 'border_inner': border_inner, 'border_size': border_size,
        'rules': {
            'caps_bold': caps_bold, 'caps_italic': caps_italic, 'caps_color_on': caps_color_on, 'caps_color': caps_color,
            'list_bold': list_bold, 'list_italic': list_italic, 'list_color_on': list_color_on, 'list_color': list_color
        }
    }

# ----------------- ЦЕНТРАЛЬНАЯ ПАНЕЛЬ (ГЕНЕРАЦИЯ) -----------------
with col_main:
    tab1, tab2 = st.tabs(["📁 Создать из Word файла (Сохраняет форматы)", "✍️ Создать из текста"])
    
    with tab1:
        st.success("✅ **Этот режим бережно перенесет все ваши цвета, маркеры и подчеркивания из загруженного файла в таблицу.**")
        uploaded_file = st.file_uploader("Загрузите файл .docx", type=["docx"])
        if st.button("🚀 Сгенерировать из файла", use_container_width=True, type="primary", key="f_btn"):
            if uploaded_file:
                with st.spinner('Читаем форматирование и верстаем...'):
                    try:
                        result = build_professional_docx(uploaded_file, config, is_file=True)
                        st.download_button("📥 Скачать готовые карточки", result, f"Cards_{uploaded_file.name}", 
                                           mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
                    except Exception as e:
                        st.error(f"Ошибка: {e}")
            else:
                st.warning("Сначала загрузите файл!")

    with tab2:
        txt_input = st.text_area("Вставьте сырой текст:", height=300)
        if st.button("🚀 Сгенерировать из текста", use_container_width=True, type="primary", key="t_btn"):
            if txt_input.strip():
                with st.spinner('Верстаем...'):
                    result = build_professional_docx(txt_input, config, is_file=False)
                    st.download_button("📥 Скачать карточки", result, "TextCards.docx", 
                                       mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
