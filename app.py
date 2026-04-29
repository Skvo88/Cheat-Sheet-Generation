import streamlit as st
import re
from io import BytesIO
from docx import Document
from docx.shared import Cm, Pt, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ROW_HEIGHT_RULE, WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# ==========================================
# ЯДРО ГЕНЕРАЦИИ (XML И ФОРМАТИРОВАНИЕ DOCX)
# ==========================================

def set_cell_margins(cell, top_mm, bottom_mm, left_mm, right_mm):
    """
    Устанавливает внутренние отступы (поля) ячейки в миллиметрах.
    Переводит миллиметры в dxa (1 дюйм = 25.4 мм = 1440 dxa).
    """
    def mm_to_dxa(mm):
        return int((mm / 25.4) * 1440)

    tcPr = cell._element.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for margin, val in zip(['top', 'bottom', 'left', 'right'], [top_mm, bottom_mm, left_mm, right_mm]):
        node = OxmlElement(f'w:{margin}')
        node.set(qn('w:w'), str(mm_to_dxa(val)))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_table_borders(table, outer_style, inner_style, border_size_pt=1):
    """
    Продвинутая функция для настройки границ таблицы через XML Oxml.
    Позволяет задать разные стили для внешних и внутренних границ.
    Размер (sz) в Word измеряется в 1/8 пункта. (т.е. 1 pt = 8).
    """
    sz_val = str(int(border_size_pt * 8))
    
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

def parse_text_to_stream(text_block, caps_headers=True):
    """
    Парсит текст, определяя заголовки (термины) и основной текст.
    Возвращает список кортежей: (слово, это_заголовок_ли).
    """
    stream =[]
    lines = text_block.split('\n')
    
    for line in lines:
        line = line.strip()
        if not line: continue
        
        is_term_line = False
        colon_idx = line.find(':')
        
        # Определяем, является ли строка началом термина (наличие номера или слова СОДЕРЖАНИЕ)
        if line.startswith("СОДЕРЖАНИЕ:") or (re.match(r'^\d+', line) and 0 < colon_idx < 200):
            is_term_line = True
        
        if is_term_line:
            parts = line.split(':', 1)
            term_name = parts[0] + ":"
            
            # Применяем КАПС к заголовку, если требует дизайнер
            if caps_headers:
                term_name = term_name.upper()
                
            for word in term_name.split():
                stream.append((word + " ", True))
                
            # Текст после двоеточия - обычный
            if len(parts) > 1:
                for word in parts[1].split():
                    stream.append((word + " ", False))
        else:
            for word in line.split():
                stream.append((word + " ", False))
        
        stream.append(("\n", False))
        
    return stream

def build_professional_docx(text_content, config):
    """
    Основная функция сборки документа по переданным настройкам.
    """
    word_stream = parse_text_to_stream(text_content, config['caps_headers'])
    
    # Разбивка на карточки (по лимиту символов)
    cells_data =[]
    current_cell =[]
    current_count = 0
    newline_weight = 35 
    
    for word, is_bold in word_stream:
        weight = newline_weight if word == "\n" else len(word)
        if current_count + weight > config['max_chars']:
            cells_data.append(current_cell)
            current_cell =[]
            current_count = 0
            if word == "\n": continue 
        current_cell.append((word, is_bold))
        current_count += weight
        
    if current_cell:
        cells_data.append(current_cell)
        
    # Добиваем пустые ячейки, чтобы таблица была ровной (кратно 4 колонкам и 3 строкам = 12)
    while len(cells_data) % 12 != 0:
        cells_data.append([])

    # Инициализация документа
    doc = Document()
    section = doc.sections[0]
    
    # 1. Формат страницы А4
    section.page_width = Mm(210)
    section.page_height = Mm(297)
    
    # 2. Поля страницы
    section.left_margin = Mm(config['page_margin'])
    section.right_margin = Mm(config['page_margin'])
    section.top_margin = Mm(config['page_margin'])
    section.bottom_margin = Mm(config['page_margin'])

    # Создание таблицы
    num_rows = len(cells_data) // config['cols_num']
    table = doc.add_table(rows=num_rows, cols=config['cols_num'])
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False # Отключаем автоподбор, жестко фиксируем ширину
    
    # Применение XML стилей границ (Сплошные внешние, пунктирные внутренние)
    set_table_borders(
        table, 
        outer_style=config['border_outer'], 
        inner_style=config['border_inner'], 
        border_size_pt=config['border_size']
    )
    
    # Настройка строк и ячеек
    for row_idx, row in enumerate(table.rows):
        row.height = Mm(config['row_height'])
        row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
        
        for col_idx, cell in enumerate(row.cells):
            cell.width = Mm(config['col_width'])
            
            # Внутренние отступы ячейки
            set_cell_margins(cell, config['pad_mm'], config['pad_mm'], config['pad_mm'], config['pad_mm'])
            
            # Индекс данных для текущей ячейки
            data_idx = row_idx * config['cols_num'] + col_idx
            if data_idx >= len(cells_data): continue
            
            data = cells_data[data_idx]
            
            p = cell.paragraphs[0]
            # Выравнивание: по ширине (Justify) или по левому краю
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY if config['justify'] else WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.line_spacing = config['line_spacing']
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.space_before = Pt(0)
            
            if not data: continue
            
            for word, is_bold in data:
                if word == "\n":
                    p.add_run().add_break()
                else:
                    run = p.add_run(word)
                    run.bold = is_bold
                    run.font.name = config['font_name']
                    run.font.size = Pt(config['font_size'])
    
    # Сохранение в буфер
    target = BytesIO()
    doc.save(target)
    target.seek(0)
    return target


# ==========================================
# ИНТЕРФЕЙС ПРИЛОЖЕНИЯ STREAMLIT
# ==========================================

st.set_page_config(page_title="Pro Генератор Карточек", page_icon="✨", layout="wide")

# -- CSS ДЛЯ КРАСОТЫ --
st.markdown("""
<style>
    .main-header { font-size: 2.5rem; font-weight: bold; color: #1E1E1E; margin-bottom: 0px;}
    .sub-header { font-size: 1.1rem; color: #555555; margin-bottom: 30px;}
</style>
""", unsafe_allow_html=True)

# Заголовки
st.markdown('<div class="main-header">✨ Pro Генератор Карточек</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Профессиональная верстка Word-файлов по дизайн-коду.</div>', unsafe_allow_html=True)

# Разделение экрана на две колонки (Боковое меню шире, Контент справа)
col_settings, col_main = st.columns([1.2, 2])

# ================= БОКОВАЯ ПАНЕЛЬ НАСТРОЕК =================
with col_settings:
    st.header("⚙️ Панель управления")
    
    # 1. Страница и Разметка
    with st.expander("📄 Разметка страницы", expanded=True):
        page_margin = st.number_input("Поля страницы со всех сторон (мм)", min_value=0.0, max_value=50.0, value=10.0, step=1.0)
        col_width = st.number_input("Ширина колонки (мм)", min_value=30.0, max_value=100.0, value=51.0, step=0.5, help="Для 4 колонок на А4 с полями 10мм идеальна ширина 47.5мм. 51мм может слегка вылезать.")
        row_height = st.number_input("Высота строки/карточки (мм)", min_value=50.0, max_value=150.0, value=92.3, step=0.5)

    # 2. Типографика (Шрифты)
    with st.expander("🔤 Типографика", expanded=True):
        font_preset = st.selectbox("Семейство шрифта",["Raleway", "Times New Roman", "Arial", "Calibri", "Свой шрифт..."])
        font_custom = ""
        if font_preset == "Свой шрифт...":
            font_custom = st.text_input("Введите точное название шрифта (напр. Montserrat):", "Montserrat")
        
        final_font = font_custom if font_preset == "Свой шрифт..." else font_preset
        
        font_size = st.number_input("Кегль (pt)", min_value=3.0, max_value=14.0, value=5.0, step=0.5)
        caps_headers = st.checkbox("Заголовки КАПСОМ (UPPERCASE)", value=True, help="Автоматически делает термины заглавными буквами")
        st.info(f"💡 Внимание: Шрифт **{final_font}** должен быть установлен на компьютере, где будут открывать Word.")

    # 3. Абзац и Текст
    with st.expander("📝 Абзац", expanded=True):
        justify = st.checkbox("Выравнивание по ширине (Justify)", value=True)
        line_spacing = st.number_input("Межстрочный интервал", min_value=0.5, max_value=2.0, value=0.92, step=0.01)
        max_chars = st.slider("Макс. символов в ячейке", 1500, 3500, 2600, step=50, help="Защита от переполнения карточки.")

    # 4. Ячейки и Границы
    with st.expander("🔲 Ячейки и Границы", expanded=True):
        pad_mm = st.number_input("Внутренний отступ текста от рамки (мм)", min_value=0.0, max_value=10.0, value=1.5, step=0.1)
        
        col_bord_1, col_bord_2 = st.columns(2)
        with col_bord_1:
            border_outer = st.selectbox("Внешняя рамка",["single", "dashed", "dotted", "double"], index=0)
        with col_bord_2:
            border_inner = st.selectbox("Внутренняя сетка", ["dashed", "single", "dotted", "double"], index=0)
            
        border_size = st.number_input("Толщина линий (pt)", min_value=0.5, max_value=5.0, value=1.0, step=0.5)

    # Сборка конфига
    config = {
        'page_margin': page_margin,
        'col_width': col_width,
        'row_height': row_height,
        'cols_num': 4,
        'font_name': final_font,
        'font_size': font_size,
        'caps_headers': caps_headers,
        'justify': justify,
        'line_spacing': line_spacing,
        'max_chars': max_chars,
        'pad_mm': pad_mm,
        'border_outer': border_outer,
        'border_inner': border_inner,
        'border_size': border_size
    }

# ================= ОСНОВНОЙ КОНТЕНТ =================
with col_main:
    st.write("### 📥 Ввод текста")
    txt_input = st.text_area("Вставьте сюда исходный текст (будет разделен по фразе 'НИЗКАЯ ВЕРОЯТНОСТЬ ПОПАДАНИЯ'):", height=500)
    
    if st.button("🚀 Сгенерировать шпоры по дизайн-коду", use_container_width=True, type="primary"):
        if not txt_input.strip():
            st.error("❌ Ошибка: Вставьте текст для обработки!")
        else:
            with st.spinner('Магия верстки в процессе...'):
                parts = re.split(r'НИЗКАЯ ВЕРОЯТНОСТЬ ПОПАДАНИЯ', txt_input)
                
                # Генерация файла 1
                file_high = build_professional_docx(parts[0], config)
                
                st.success("✅ Генерация успешно завершена!")
                st.write("### 💾 Скачать результаты")
                
                dl_col1, dl_col2 = st.columns(2)
                with dl_col1:
                    st.download_button(
                        label="📄 Скачать: ВЫСОКАЯ вероятность",
                        data=file_high,
                        file_name="Cards_High_Probability.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )
                
                # Генерация файла 2 (если есть разделение)
                if len(parts) > 1:
                    file_low = build_professional_docx(parts[1], config)
                    with dl_col2:
                        st.download_button(
                            label="📄 Скачать: НИЗКАЯ вероятность",
                            data=file_low,
                            file_name="Cards_Low_Probability.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True
                        )

st.markdown("---")
st.caption("Разработано для идеальной печати. Дизайнеры будут в восторге. 🎨")
