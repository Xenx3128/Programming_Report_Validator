import docx.document
import xml.etree.ElementTree as ET
from docx.api import Document
from docx.document import Document as doctwo
from docx.shared import Pt, Cm, Emu
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT as WD_ALIGN_VERTICAL
from docx.enum.section import WD_ORIENTATION
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table, _Cell
from docx.text.paragraph import Paragraph
from io import StringIO
import re



REGEX_TRANSFORM = {
    "Таблица <N> - <Название>.": "^Таблица [1-9]*(.[1-9])* (–|-) .*$",
    "<Название>.": ".*",
    "Рисунок <N> - <Название>.": "^Рисунок [1-9]*(.[1-9])* (–|-) .*$"
}

PARAM_TO_COMMENT = {
    "format_regex": "Формат",
    "font_name": "Тип шрифта",
    "font_size": "Размер шрифта",
    "font_bald": "Полужирный",
    "font_italic": "Курсив",
    "font_underline": "Подчёркивание",
    "font_color": "Цвет текста",
    "font_highlight": "Цвет подчёркивания",
    "alignment": "Выравнивание",
    "vert_alignment": "Выравнивание по вертикали",
    "keep_with_next": "Не отрывать от следующего",
    "page_break_before": "С новой страницы",
    "space_before": "Интервал перед",
    "space_after": "Интервал после",
    "left_indent": "Отступ слева",
    "right_indent": "Отступ справа",
    "first_line_indent": "Отступ первой строки",
    "line_spacing": "Межстрочный интервал",
    "is_list": "Нумерация",
    "top_margin": "Верхнее поле",
    "bottom_margin": "Нижнее поле",
    "left_margin": "Левое поле",
    "right_margin": "Правое поле",
    "orientation": "Ориентация",
}
VAR_TO_COMMENT = {
    True: "Да",
    False: "Нет",
    None: "Нет",
}
ALIGNMENT_TO_COMMENT = {
    WD_PARAGRAPH_ALIGNMENT.LEFT: "Левый край",
    WD_PARAGRAPH_ALIGNMENT.CENTER: "По центру",
    WD_PARAGRAPH_ALIGNMENT.RIGHT: "Правый край",
    WD_PARAGRAPH_ALIGNMENT.JUSTIFY: "По ширине",
}
VERT_ALIGNMENT_TO_COMMENT = {
    WD_ALIGN_VERTICAL.TOP: "Верхний край",
    WD_ALIGN_VERTICAL.CENTER: "По центру",
    WD_ALIGN_VERTICAL.BOTTOM: "Нижний край",
    WD_ALIGN_VERTICAL.BOTH: "Если вы видите это сообщение, что-то пошло не так",
}
ORIENTATION_TO_COMMENT = {
    WD_ORIENTATION.PORTRAIT: "Книжная",
    WD_ORIENTATION.LANDSCAPE: "Альбомная"
}

MODIFICATIONS = {
    "paragraph_after_table": {
        "space_before": 13.0
    }
}