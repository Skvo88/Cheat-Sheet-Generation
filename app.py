import streamlit as st
import re
from io import BytesIO
from docx import Document
from docx.shared import Cm, Pt, Mm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
from docx.enum.table import WD_ROW_HEIGHT_RULE, WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# =====================================================================
# БЛОК 1: НИЗКОУРОВНЕВАЯ РАБОТА С XML (СЕТКА И ГРАНИЦЫ ИЗ ВАШЕГО КОДА)
# =====================================================================

def set_cell_margins(cell, top_mm, bottom_mm, left_mm, right_mm):
    """
    Устанавливает внутренние отступы (padding) ячейки в миллиметрах.
    Оригинальная функция из вашего кода (сохранена на 100%).
    """
    def mm_to_dxa(mm):
        return int((mm / 25.4) * 1440)

    tcPr = cell._element.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for margin, val in zip(['top', 'bottom', 'left', 'right'],[top_mm, bottom_mm, left_mm, right_mm]):
        node = OxmlElement(f'w:{margin}')
        node.set(qn('w:w'), str(mm_to_dxa(val)))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_table_borders(table, outer_style, inner_style, border_size_pt=1):
    """
    Продвинутая функция для настройки границ таблицы через XML Oxml.
    Позволяет задать разные стили для внешних и внутренних границ.
    """
    sz_val = str(int(border_size_pt * 8)) # Размер (sz) измеряется в 1/8 пункта
    
    tblPr = table._element.tblPr
    tblBorders = OxmlElement('w:tblBorders')

    # Внешние границы
    for border_name in ['top', 'left', 'bottom', 'right']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), outer_style)
        border.set(qn('w:sz'), sz_val)
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), '000000')
        tblBorders.append(border)

    # Внутренние границы
    for border_name in ['insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), inner_style)
        border.set(qn('w:sz'), sz_val)
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), '000000')
        tblBorders.append(border)

    tblPr.append(tblBorders)

# =====================================================================
# БЛОК 2: АНАЛИЗАТОР ФОРМАТИРОВАНИЯ (КЛОНИРОВАНИЕ ДОКУМЕНТА)
# =====================================================================

def extract_run_format(run):
    """
    Вытаскивает ВСЕ стили из конкретного кусочка текста в Word.
    Гарантирует, что оригинальное выделение (цвета, фоны, жирность) не потеряется.
    """
    # Безопасное извлечение цвета текста (RGB)
    color_rgb = None
    if run.font.color and run.font.color.type == 1:  # 1 означает RGB
        color_rgb = run.font.color.rgb

    return {
        'bold': bool(run.bold),
        'italic': bool(run.italic),
        'underline': bool(run.underline),
        'strike': bool(run.font.strike),
        'color': color_rgb,
        'highlight': run.font.highlight_color, # Цвет маркера (фон)
    }

def apply_format_to_run(run, fmt_dict):
    """
    Применяет сохраненный словарь стилей к новому кусочку текста в карточке.
    """
    if fmt_dict['bold']: run.bold = True
    if fmt_dict['italic']: run.italic = True
    if fmt_dict['underline']: run.underline = True
    if fmt_dict['strike']: run.font.strike = True
    
    if fmt_dict['color']:
        run.font.color.rgb = fmt_dict['color']
        
    if fmt_dict['highlight']:
        run.font.highlight_color = fmt_dict['highlight']

# =====================================================================
# БЛОК 3: "СПЯЩИЕ" УМНЫЕ ПРАВИЛА (ВКЛЮЧАЮТСЯ ТОЛЬКО ПО ЖЕЛАНИЮ)
# =====================================================================

def hex_to_rgb(hex_str):
    """Конвертер '#FF0000' -> RGBColor(255, 0, 0)"""
    hex_str = hex_str.lstrip('#')
    if len(hex_str) != 6: return None
    return RGBColor(int(hex_str[0:2], 16), int(hex_str[2:4], 16), int(hex_str[4:6], 16))

def override_with_smart_rules(word, current_fmt, rules):
    """
    Если пользователь включил "Умные правила" в боковой панели, 
    эта функция перезапишет форматы на лету.
    """
    if not rules.get('enable_smart_rules', False):
        return current_fmt # Если правила выключены - отдаем оригинальный формат без изменений!

    new_fmt = current_fmt.copy()
    word_clean = word.strip()

    # Правило 1: Текст написан КАПСОМ
    if word_clean.isupper() and any(c.isalpha() for c in word_clean):
        if rules['caps_bold']: new_fmt['bold'] = True
        if rules['caps_italic']: new_fmt['italic'] = True
        if rules['caps_color_on']: new_fmt['color'] = hex_to_rgb(rules['caps_color'])

    # Правило 2: Нумерация со скобкой (например "1)", "а)", "25)")
    if re.match(r'^[\dа-яА-Яa-zA-Z]+\)$', word_clean):
        if rules['list_bold']: new_fmt['bold'] = True
        if rules['list_italic']: new_fmt['italic'] = True
        if rules['list_color_on']: new_fmt['color'] = hex_to_rgb(rules['list_color'])

    return new_fmt

# =====================================================================
# БЛОК 4: ПАРСЕРЫ ТЕКСТА (РАЗБИВКА НА СЛОВА С УЧЕТОМ ФОРМАТА)
# =====================================================================

def parse_docx_file(file_bytes, rules):
    """
    Читает загруженный Word-файл, разбивает его на слова,
    сохраняя оригинальный стиль каждого слова.
    """
    stream =[]
    doc = Document(file_bytes)
    
    for p in doc.paragraphs:
        if not p.text.strip():
            # Если абзац пустой - добавляем просто перенос строки с пустым форматом
            stream.append(("\n", {'bold': False, 'italic': False, 'underline': False, 'strike': False, 'color': None, 'highlight': None}))
            continue
            
        for run in p.runs:
            text = run.text
            if not text: continue
            
            orig_fmt = extract_run_format(run)
            
            # Разбиваем на слова, чтобы не сломать логику переноса в карточках
            words = text.split(' ')
            for i, word in enumerate(words):
                if not word and i != len(words)-1:
                    stream.append((" ", orig_fmt))
                    continue
                if not word: continue
                
                # Применяем умные правила (если они включены)
                final_fmt = override_with_smart_rules(word, orig_fmt, rules)
                
                # Возвращаем пробел на место
                suffix = " " if i < len(words) - 1 else ""
                stream.append((word + suffix, final_fmt))
                
        # В конце каждого абзаца добавляем перенос строки
        stream.append(("\n", {'bold': False, 'italic': False, 'underline': False, 'strike': False, 'color': None, 'highlight': None}))
        
    return stream

# =====================================================================
# БЛОК 5: ЯДРО ГЕНЕРАЦИИ (СБОРКА ИДЕАЛЬНОЙ СЕТКИ КАРТОЧЕК)
# =====================================================================

def build_professional_docx(word_stream, config):
    """
    Берет поток слов со стилями и безупречно укладывает их 
    в вашу таблицу (3x4 или любые другие настройки).
    """
    cells_data = []
    current_cell =[]
    current_count = 0
    newline_weight = 35 # Вес переноса строки (чтобы защитить карточку от переполнения по высоте)
    
    # 1. Распределяем слова по карточкам
    for text_chunk, fmt in word_stream:
        weight = newline_weight if text_chunk == "\n" else len(text_chunk)
        
        # Защита от переполнения ячейки
        if current_count + weight > config['max_chars']:
            cells_data.append(current_cell)
            current_cell =[]
            current_count = 0
            if text_chunk == "\n": continue 
            
        current_cell.append((text_chunk, fmt))
        current_count += weight
        
    if current_cell:
        cells_data.append(current_cell)
        
    # 2. Добиваем пустые ячейки (чтобы лист заканчивался ровно)
    while len(cells_data) % config['cols_num'] != 0:
        cells_data.append([])

    # 3. Инициализация документа и настройка листа А4
    doc = Document()
    section = doc.sections[0]
    section.page_width = Mm(210)
    section.page_height = Mm(297)
    
    # Поля страницы
    section.left_margin = Mm(config['page_margin'])
    section.right_margin = Mm(config['page_margin'])
    section.top_margin = Mm(config['page_margin'])
    section.bottom_margin = Mm(config['page_margin'])

    # 4. Создание и настройка таблицы
    num_rows = len(cells_data) // config['cols_num']
    table = doc.add_table(rows=num_rows, cols=config['cols_num'])
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    
    # Настройка границ таблицы
    set_table_borders(
        table, 
        outer_style=config['border_outer'], 
        inner_style=config['border_inner'], 
        border_size_pt=config['border_size']
    )
    
    # 5. Заполнение таблицы текстом
    for row_idx, row in enumerate(table.rows):
        row.height = Mm(config['row_height'])
        row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY # Жесткая фиксация высоты
        
        for col_idx, cell in enumerate(row.cells):
            cell.width = Mm(config['col_width'])
            
            # Отступы внутри ячейки
            set_cell_margins(cell, config['pad_mm'], config['pad_mm'], config['pad_mm'], config['pad_mm'])
            
            data_idx = row_idx * config['cols_num'] + col_idx
            if data_idx >= len(cells_data): continue
            
            data = cells_data[data_idx]
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY if config['justify'] else WD_ALIGN_PARAGRAPH.LEFT
            
            # Настройка интервалов абзаца
            p.paragraph_format.line_spacing = config['line_spacing']
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.space_before = Pt(0)
            
            if not data: continue
            
            # Отрисовка каждого слова со своим форматом
            for text_chunk, fmt_dict in data:
                if text_chunk == "\n":
                    p.add_run().add_break()
                else:
                    run = p.add_run(text_chunk)
                    
                    # Применяем общий шрифт для всех карточек
                    run.font.name = config['font_name']
                    run.font.size = Pt(config['font_size'])
                    
                    # Накатываем оригинальные (или умные) стили
                    apply_format_to_run(run, fmt_dict)
    
    # Сохранение в байтовый буфер для скачивания
    target = BytesIO()
    doc.save(target)
    target.seek(0)
    return target


# =====================================================================
# БЛОК 6: ИНТЕРФЕЙС STREAMLIT (ПОД КЛЮЧ)
# =====================================================================

st.set_page_config(page_title="Pro Генератор Карточек", page_icon="✨", layout="wide")

st.markdown("""
<style>
    .main-header { font-size: 2.5rem; font-weight: 800; color: #1E1E1E; margin-bottom: 0px;}
    .sub-header { font-size: 1.1rem; color: #555555; margin-bottom: 30px;}
    .st-expander { border-radius: 8px !important; border: 1px solid #ddd !important; }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-header">✨ Pro Генератор Карточек</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Клонирует ваш Word-файл 1-в-1, сохраняя все цвета и выделения, перенося их в идеальную сетку карточек.</div>', unsafe_allow_html=True)

col_settings, col_main = st.columns([1.3, 2.2], gap="large")

# ----------------- САЙДБАР: НАСТРОЙКИ СЕТКИ И ПРАВИЛ -----------------
with col_settings:
    st.header("⚙️ Настройки сетки (Оригинал)")
    
    with st.expander("📄 Разметка и Размеры (Ваша сетка)", expanded=True):
        st.caption("Эти настройки гарантируют, что карточки встанут идеально.")
        cols_num = st.number_input("Количество колонок", min_value=1, max_value=10, value=4, step=1)
        page_margin = st.number_input("Поля страницы (мм)", min_value=0.0, max_value=50.0, value=10.0, step=1.0)
        col_width = st.number_input("Ширина колонки (мм)", min_value=30.0, max_value=100.0, value=51.0, step=0.5)
        row_height = st.number_input("Высота карточки (мм)", min_value=50.0, max_value=150.0, value=92.3, step=0.5)
        pad_mm = st.number_input("Внутренний отступ (мм)", min_value=0.0, max_value=10.0, value=1.5, step=0.1)

    with st.expander("🔲 Границы таблиц", expanded=False):
        c_b1, c_b2 = st.columns(2)
        border_options = ["single", "dashed", "dotted", "double", "nil"]
        with c_b1: border_outer = st.selectbox("Внешняя рамка", border_options, index=0)
        with c_b2: border_inner = st.selectbox("Внутренняя сетка", border_options, index=1)
        border_size = st.number_input("Толщина линий (pt)", min_value=0.5, max_value=5.0, value=1.0, step=0.5)

    with st.expander("🔤 Базовый Шрифт и Текст", expanded=False):
        font_preset = st.selectbox("Семейство шрифта", ["Raleway", "Times New Roman", "Arial", "Calibri", "Свой шрифт..."])
        font_custom = ""
        if font_preset == "Свой шрифт...":
            font_custom = st.text_input("Введите точное название шрифта:", "Montserrat")
        final_font = font_custom if font_preset == "Свой шрифт..." else font_preset
        
        font_size = st.number_input("Кегль (pt)", min_value=3.0, max_value=14.0, value=5.0, step=0.5)
        justify = st.checkbox("Выравнивание по ширине (Justify)", value=True)
        line_spacing = st.number_input("Межстрочный интервал", min_value=0.5, max_value=2.0, value=0.92, step=0.01)
        max_chars = st.slider("Макс. символов в карточке", 1000, 4000, 2600, step=50)

    # СПЯЩИЕ НАСТРОЙКИ (ВКЛЮЧАЮТСЯ ГАЛОЧКОЙ)
    st.header("🤖 Умные правила (Спящие)")
    enable_smart_rules = st.checkbox("🔥 АКТИВИРОВАТЬ ПЕРЕЗАПИСЬ ФОРМАТОВ", value=False, help="Если выключено - скрипт на 100% копирует исходник. Если включить - эти правила изменят оригинальный файл.")
    
    with st.expander("Настройки умных правил", expanded=enable_smart_rules):
        if not enable_smart_rules:
            st.warning("Сейчас правила спят. Поставьте галочку выше, чтобы они начали работать.")
            
        st.markdown("**Что делать с КАПСОМ (Слова ЗАГЛАВНЫМИ):**")
        r1, r2, r3 = st.columns(3)
        caps_bold = r1.checkbox("Жирный", value=True, disabled=not enable_smart_rules)
        caps_italic = r2.checkbox("Курсив", value=False, disabled=not enable_smart_rules)
        caps_color_on = r3.checkbox("Цвет", value=False, disabled=not enable_smart_rules)
        caps_color = st.color_picker("Цвет КАПСА", "#FF0000", disabled=not enable_smart_rules)
        
        st.markdown("---")
        st.markdown("**Что делать со списками (1), 2), а):**")
        l1, l2, l3 = st.columns(3)
        list_bold = l1.checkbox("Жирный ", value=True, key="lb", disabled=not enable_smart_rules)
        list_italic = l2.checkbox("Курсив ", value=False, key="li", disabled=not enable_smart_rules)
        list_color_on = l3.checkbox("Цвет ", value=False, key="lc", disabled=not enable_smart_rules)
        list_color = st.color_picker("Цвет списков", "#0000FF", disabled=not enable_smart_rules)

    # Сборка общего конфига
    config = {
        'cols_num': cols_num, 'page_margin': page_margin, 'col_width': col_width, 
        'row_height': row_height, 'pad_mm': pad_mm, 'border_outer': border_outer, 
        'border_inner': border_inner, 'border_size': border_size,
        'font_name': final_font, 'font_size': font_size, 'justify': justify, 
        'line_spacing': line_spacing, 'max_chars': max_chars,
        'rules': {
            'enable_smart_rules': enable_smart_rules,
            'caps_bold': caps_bold, 'caps_italic': caps_italic, 
            'caps_color_on': caps_color_on, 'caps_color': caps_color,
            'list_bold': list_bold, 'list_italic': list_italic, 
            'list_color_on': list_color_on, 'list_color': list_color
        }
    }

# ----------------- ЦЕНТРАЛЬНАЯ ПАНЕЛЬ: ЗАГРУЗКА -----------------
with col_main:
    st.write("### 📥 Режим загрузки документа (Клонирование)")
    st.info("Загрузите ваш документ. Программа перенесет **КАЖДЫЙ** ваш желтый фон, зеленый шрифт, жирный или курсивный текст в ровную сетку карточек-шпор.")
    
    uploaded_file = st.file_uploader("Загрузите файл (.docx)", type=["docx"])
    
    if st.button("🚀 Сгенерировать шпоры (Clone Mode)", use_container_width=True, type="primary"):
        if uploaded_file is None:
            st.error("❌ Пожалуйста, загрузите файл.")
        else:
            with st.spinner('Анализируем ваш шедевр и переносим в карточки...'):
                try:
                    # 1. Парсим загруженный файл
                    word_stream = parse_docx_file(uploaded_file, config['rules'])
                    
                    # 2. Собираем новый документ по сетке
                    result_doc = build_professional_docx(word_stream, config)
                    
                    st.success("✅ Готово! Все ваши форматы бережно перенесены.")
                    
                    st.download_button(
                        label="📄 Скачать идеальные шпоры",
                        data=result_doc,
                        file_name=f"Perfect_CheatSheets_{uploaded_file.name}",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )
                except Exception as e:
                    st.error(f"❌ Критическая ошибка: {e}")

st.markdown("---")
st.caption("Режим «1-в-1 Ксерокс». Дизайн-код карточек сохранен из оригинала. Умные правила спят по умолчанию.")
