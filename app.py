import streamlit as st
import re
from io import BytesIO
from docx import Document
from docx.shared import Cm, Pt, Mm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ROW_HEIGHT_RULE, WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# ==========================================
# 1. ЯДРО: НИЗКОУРОВНЕВАЯ РАБОТА С XML DOCX
# ==========================================

def set_cell_margins(cell, top_mm, bottom_mm, left_mm, right_mm):
    """Настраивает внутренние отступы ячейки (padding) через OXML."""
    def mm_to_dxa(mm): return int((mm / 25.4) * 1440)
    tcPr = cell._element.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for margin, val in zip(['top', 'bottom', 'left', 'right'],[top_mm, bottom_mm, left_mm, right_mm]):
        node = OxmlElement(f'w:{margin}')
        node.set(qn('w:w'), str(mm_to_dxa(val)))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_table_borders(table, outer_style, inner_style, border_size_pt=1, border_color="000000"):
    """Устанавливает стили границ таблицы через OXML."""
    sz_val = str(int(border_size_pt * 8)) # 1 pt = 8 1/8ths of a point
    tblPr = table._element.tblPr
    tblBorders = OxmlElement('w:tblBorders')

    # Внешние границы
    for border_name in ['top', 'left', 'bottom', 'right']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), outer_style)
        border.set(qn('w:sz'), sz_val)
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), border_color)
        tblBorders.append(border)

    # Внутренние границы
    for border_name in ['insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), inner_style)
        border.set(qn('w:sz'), sz_val)
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), border_color)
        tblBorders.append(border)

    tblPr.append(tblBorders)

# ==========================================
# 2. УМНЫЙ ДВИЖОК ПРАВИЛ (SMART RULES ENGINE)
# ==========================================

def hex_to_rgb(hex_color):
    """Преобразует HEX цвет (#FF0000) в объект RGBColor для python-docx."""
    hex_color = hex_color.lstrip('#')
    if len(hex_color) != 6: return RGBColor(0, 0, 0)
    return RGBColor(int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16))

def apply_smart_rules(text, current_format, rules):
    """
    Анализирует кусок текста и применяет к нему правила (КАПС, списки и т.д.).
    Возвращает обновленный словарь форматов.
    """
    new_format = current_format.copy()
    text_stripped = text.strip()
    
    # ПРАВИЛО 1: КАПСЛОК
    # Если слово состоит из заглавных букв (минимум 2 буквы, чтобы отсеять предлоги типа "В", "С")
    # или если это короткое слово, но полностью заглавное.
    if text_stripped.isupper() and any(c.isalpha() for c in text_stripped):
        if rules['caps_action_bold']: new_format['bold'] = True
        if rules['caps_action_italic']: new_format['italic'] = True
        if rules['caps_action_color_apply']: new_format['color'] = rules['caps_color']
        
    # ПРАВИЛО 2: НУМЕРАЦИЯ СО СКОБКОЙ (напр. "1)", "25)", "а)")
    if re.match(r'^[\dа-яА-Яa-zA-Z]+\)$', text_stripped):
        if rules['list_action_bold']: new_format['bold'] = True
        if rules['list_action_italic']: new_format['italic'] = True
        if rules['list_action_color_apply']: new_format['color'] = rules['list_color']
        
    return new_format

# ==========================================
# 3. ПАРСЕРЫ (ИЗВЛЕЧЕНИЕ ТЕКСТА И ФОРМАТОВ)
# ==========================================

def parse_raw_text(text_block, rules):
    """Парсит обычный текст, разбивает на слова и применяет умные правила."""
    stream =[]
    lines = text_block.split('\n')
    
    for line in lines:
        if not line.strip(): 
            stream.append(("\n", {'bold': False, 'italic': False, 'color': None}))
            continue
            
        words = line.split(' ')
        for word in words:
            if not word: continue
            base_format = {'bold': False, 'italic': False, 'color': None}
            final_format = apply_smart_rules(word, base_format, rules)
            stream.append((word + " ", final_format))
            
        stream.append(("\n", {'bold': False, 'italic': False, 'color': None}))
    return stream

def parse_docx_file(file_bytes, rules):
    """
    Читает готовый DOCX. Сохраняет оригинальное форматирование 
    и накатывает поверх умные правила.
    """
    stream =[]
    doc = Document(file_bytes)
    
    for p in doc.paragraphs:
        if not p.text.strip():
            stream.append(("\n", {'bold': False, 'italic': False, 'color': None}))
            continue
            
        for run in p.runs:
            text = run.text
            if not text: continue
            
            # Извлекаем оригинальный формат
            orig_format = {
                'bold': bool(run.bold),
                'italic': bool(run.italic),
                'color': None # Чтение цвета сложнее, пока берем только жирность и курсив
            }
            
            # Разбиваем run на слова, чтобы применить правила к конкретным словам (например капсу)
            words = text.split(' ')
            for i, word in enumerate(words):
                if not word and i != len(words)-1:
                    stream.append((" ", orig_format))
                    continue
                if not word: continue
                
                # Применяем умные правила поверх оригинального формата
                final_format = apply_smart_rules(word, orig_format, rules)
                
                # Добавляем пробел, если он был "съеден" split-ом
                suffix = " " if i < len(words) - 1 else ""
                stream.append((word + suffix, final_format))
                
        stream.append(("\n", {'bold': False, 'italic': False, 'color': None}))
        
    return stream

# ==========================================
# 4. ДВИЖОК СБОРКИ ДОКУМЕНТА (BUILDER)
# ==========================================

def build_professional_docx(source_data, config, is_file=False):
    """
    Собирает идеальный Word-документ (шпору) на основе разобранного потока слов.
    """
    # 1. Получаем поток слов и форматов
    if is_file:
        word_stream = parse_docx_file(source_data, config['rules'])
    else:
        word_stream = parse_raw_text(source_data, config['rules'])
        
    # 2. Разбивка на карточки (ячейки)
    cells_data =[]
    current_cell =[]
    current_count = 0
    newline_weight = 35 # Штрафной вес за перенос строки (чтобы не переполняло по высоте)
    
    for text_chunk, fmt in word_stream:
        weight = newline_weight if text_chunk == "\n" else len(text_chunk)
        
        # Если лимит превышен - создаем новую карточку
        if current_count + weight > config['max_chars']:
            cells_data.append(current_cell)
            current_cell =[]
            current_count = 0
            if text_chunk == "\n": continue # Не начинаем новую карточку с пустой строки
            
        current_cell.append((text_chunk, fmt))
        current_count += weight
        
    if current_cell:
        cells_data.append(current_cell)
        
    # Добиваем пустые ячейки для ровной таблицы (кратно количеству колонок)
    while len(cells_data) % config['cols_num'] != 0:
        cells_data.append([])

    # 3. Сборка документа
    doc = Document()
    section = doc.sections[0]
    
    # Настройка листа А4
    section.page_width = Mm(210)
    section.page_height = Mm(297)
    for margin in ['left_margin', 'right_margin', 'top_margin', 'bottom_margin']:
        setattr(section, margin, Mm(config['page_margin']))

    # Создание таблицы
    num_rows = len(cells_data) // config['cols_num']
    table = doc.add_table(rows=num_rows, cols=config['cols_num'])
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False 
    
    # Применение XML стилей границ
    set_table_borders(
        table, 
        outer_style=config['border_outer'], 
        inner_style=config['border_inner'], 
        border_size_pt=config['border_size']
    )
    
    # Заполнение ячеек
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
            p.paragraph_format.space_before = Pt(0)
            
            if not data: continue
            
            for text_chunk, fmt in data:
                if text_chunk == "\n":
                    p.add_run().add_break()
                else:
                    run = p.add_run(text_chunk)
                    run.font.name = config['font_name']
                    run.font.size = Pt(config['font_size'])
                    
                    # Применяем форматы
                    if fmt['bold']: run.bold = True
                    if fmt['italic']: run.italic = True
                    if fmt['color']: run.font.color.rgb = hex_to_rgb(fmt['color'])
    
    target = BytesIO()
    doc.save(target)
    target.seek(0)
    return target


# ==========================================
# 5. ИНТЕРФЕЙС ПРИЛОЖЕНИЯ STREAMLIT
# ==========================================

st.set_page_config(page_title="Ultimate Карточки", page_icon="🎛️", layout="wide")

st.markdown("""
<style>
    .main-header { font-size: 2.2rem; font-weight: 800; color: #1E1E1E; margin-bottom: 5px; }
    .sub-header { font-size: 1.1rem; color: #666; margin-bottom: 25px; }
    .st-expander { border-radius: 10px !important; border: 1px solid #E0E0E0 !important; }
    div[data-baseweb="tooltip"] { font-size: 14px !important; }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-header">🎛️ Ultimate Генератор Шпор</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Умная система верстки: свои правила, точная сетка, поддержка готовых файлов.</div>', unsafe_allow_html=True)

col_settings, col_main = st.columns([1.3, 2.2], gap="large")

# ----------------- САЙДБАР: НАСТРОЙКИ -----------------
with col_settings:
    st.write("### ⚙️ Панель управления")
    
    # БЛОК 1: Умные правила (Главная фишка)
    with st.expander("🧠 Умные правила текста", expanded=True):
        st.markdown("**Что делать со словами ЗАГЛАВНЫМИ БУКВАМИ (КАПС)?**")
        c1, c2, c3 = st.columns(3)
        caps_bold = c1.checkbox("Жирный", value=True, help="Сделает все слова, написанные КАПСОМ, полужирными.")
        caps_italic = c2.checkbox("Курсив", value=False, help="Сделает КАПС курсивным.")
        caps_color_apply = c3.checkbox("Цветной", value=False, help="Покрасит КАПС в выбранный цвет.")
        caps_color = "#FF0000"
        if caps_color_apply:
            caps_color = st.color_picker("Выбрать цвет для КАПСА", "#FF0000")
            
        st.markdown("---")
        st.markdown("**Что делать с нумерацией списков (1), 2), а) и т.д.)?**")
        l1, l2, l3 = st.columns(3)
        list_bold = l1.checkbox("Жирный ", value=True, key="lb", help="Выделит цифры со скобками полужирным.")
        list_italic = l2.checkbox("Курсив ", value=False, key="li", help="Выделит цифры со скобками курсивом.")
        list_color_apply = l3.checkbox("Цветной ", value=False, key="lc", help="Покрасит цифры в выбранный цвет.")
        list_color = "#0000FF"
        if list_color_apply:
            list_color = st.color_picker("Выбрать цвет для нумерации", "#0000FF")

    # БЛОК 2: Базовый стиль шрифта
    with st.expander("🔤 Типографика (База)", expanded=False):
        font_preset = st.selectbox(
            "Семейство шрифта",
            ["Raleway", "Times New Roman", "Arial", "Calibri", "Свой шрифт..."],
            help="Шрифт, которым будет написан весь текст. Убедитесь, что он установлен на вашем ПК."
        )
        font_custom = ""
        if font_preset == "Свой шрифт...":
            font_custom = st.text_input("Название шрифта (напр. Montserrat):", "Montserrat")
        final_font = font_custom if font_preset == "Свой шрифт..." else font_preset
        
        font_size = st.number_input("Кегль шрифта (pt)", min_value=3.0, max_value=14.0, value=5.0, step=0.5)
        line_spacing = st.number_input("Межстрочный интервал", min_value=0.5, max_value=2.0, value=0.92, step=0.01)
        justify = st.checkbox("Выравнивание по ширине (Justify)", value=True, help="Делает края текста ровными, растягивая пробелы.")

    # БЛОК 3: Макет и карточки
    with st.expander("📄 Сетка и Размеры", expanded=False):
        cols_num = st.number_input("Количество колонок", min_value=1, max_value=10, value=4, step=1, help="На сколько частей бить лист по вертикали.")
        page_margin = st.number_input("Поля страницы А4 (мм)", min_value=0.0, max_value=50.0, value=10.0, step=1.0)
        col_width = st.number_input("Ширина карточки (мм)", min_value=10.0, max_value=200.0, value=51.0, step=0.5)
        row_height = st.number_input("Высота карточки (мм)", min_value=10.0, max_value=290.0, value=92.3, step=0.5)
        max_chars = st.slider(
            "Вместимость карточки (символов)", 500, 4000, 2600, step=50, 
            help="Максимальное количество текста в одной ячейке. Если текст не влезает, он перенесется в следующую."
        )

    # БЛОК 4: Рамки
    with st.expander("🔲 Границы и Отступы", expanded=False):
        pad_mm = st.number_input("Внутренний отступ текста (мм)", min_value=0.0, max_value=10.0, value=1.5, step=0.1, help="Расстояние от рамки до самого текста внутри карточки.")
        
        b1, b2 = st.columns(2)
        border_styles = ["single", "dashed", "dotted", "double", "nil"]
        with b1: border_outer = st.selectbox("Внешняя рамка", border_styles, index=0, help="Рамка по периметру всей таблицы.")
        with b2: border_inner = st.selectbox("Внутренняя сетка", border_styles, index=1, help="Линии реза между карточками.")
            
        border_size = st.number_input("Толщина линий (pt)", min_value=0.1, max_value=5.0, value=1.0, step=0.5)

    # Упаковка конфига
    config = {
        'page_margin': page_margin, 'col_width': col_width, 'row_height': row_height,
        'cols_num': cols_num, 'font_name': final_font, 'font_size': font_size,
        'justify': justify, 'line_spacing': line_spacing, 'max_chars': max_chars,
        'pad_mm': pad_mm, 'border_outer': border_outer, 'border_inner': border_inner,
        'border_size': border_size,
        'rules': {
            'caps_action_bold': caps_bold, 'caps_action_italic': caps_italic, 
            'caps_action_color_apply': caps_color_apply, 'caps_color': caps_color,
            'list_action_bold': list_bold, 'list_action_italic': list_italic, 
            'list_action_color_apply': list_color_apply, 'list_color': list_color
        }
    }

# ----------------- ЦЕНТРАЛЬНАЯ ПАНЕЛЬ: РАБОТА С ФАЙЛАМИ -----------------
with col_main:
    tab1, tab2 = st.tabs(["📁 Загрузить файл (Word .docx)", "✍️ Вставить сырой текст"])
    
    # ВКЛАДКА 1: Работа с готовым файлом Word
    with tab1:
        st.info("💡 **Режим загрузки файла:** Скрипт сохранит ваши полужирные шрифты и курсивы из оригинального документа, а затем дополнительно применит «Умные правила» из панели слева.")
        uploaded_file = st.file_uploader("Загрузите ваш документ (.docx)", type=["docx"])
        btn_file = st.button("🚀 Сгенерировать шпоры из файла", use_container_width=True, type="primary", key="btn1")
        
        if btn_file and uploaded_file:
            with st.spinner('Анализируем форматы, применяем правила и верстаем...'):
                try:
                    result_doc = build_professional_docx(uploaded_file, config, is_file=True)
                    st.success("✅ Документ идеально сверстан!")
                    st.download_button(
                        label="📥 Скачать готовые карточки (.docx)",
                        data=result_doc,
                        file_name=f"CheatSheets_{uploaded_file.name}",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )
                except Exception as e:
                    st.error(f"❌ Произошла ошибка при обработке файла: {e}")

    # ВКЛАДКА 2: Работа с сырым текстом
    with tab2:
        st.write("Сюда можно вставить любой текст (понятия, 20-е задание, 25-е и т.д.).")
        txt_input = st.text_area("Текст (разделение на блоки по фразе 'НИЗКАЯ ВЕРОЯТНОСТЬ ПОПАДАНИЯ'):", height=300)
        btn_text = st.button("🚀 Сгенерировать шпоры из текста", use_container_width=True, type="primary", key="btn2")
        
        if btn_text and txt_input.strip():
            with st.spinner('Применяем умные правила и верстаем текст...'):
                try:
                    parts = re.split(r'НИЗКАЯ ВЕРОЯТНОСТЬ ПОПАДАНИЯ', txt_input)
                    
                    file_high = build_professional_docx(parts[0], config, is_file=False)
                    st.success("✅ Генерация успешна!")
                    
                    d1, d2 = st.columns(2)
                    with d1:
                        st.download_button("📥 Скачать: БЛОК 1 (До разделения)", file_high, "Cards_Part1.docx", use_container_width=True)
                    
                    if len(parts) > 1:
                        file_low = build_professional_docx(parts[1], config, is_file=False)
                        with d2:
                            st.download_button("📥 Скачать: БЛОК 2 (После разделения)", file_low, "Cards_Part2.docx", use_container_width=True)
                except Exception as e:
                    st.error(f"❌ Произошла ошибка при сборке текста: {e}")

st.markdown("<br><hr>", unsafe_allow_html=True)
st.caption("🚀 Продвинутая система верстки. Построено на Python, Streamlit и python-docx. Архитектура: Пайплайн обработки форматов.")
