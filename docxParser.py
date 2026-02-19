import os
from pathlib import Path

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

from default_settings import *
import inspect


class BaseChecklist:
    def set_settings(self, settings: dict):
        for par, value in settings.items():
            try:
                setattr(self, par, value)
            except:
                continue


class ParagraphChecklist(BaseChecklist):
    def __init__(self):
        super().__init__()
        self.format_regex = None
        self.font_name = None
        self.font_size = None
        self.font_bald = None
        self.font_italic = None
        self.font_underline = None
        self.font_color = None  # !!!
        self.font_back_color = None  # !!!
        self.alignment = None
        self.keep_with_next = None
        self.page_break_before = None
        self.space_before = None
        self.space_after = None
        self.left_indent = None
        self.right_indent = None
        self.first_line_indent = None
        self.line_spacing = None


class ListChecklist(ParagraphChecklist):
    def __init__(self):
        super().__init__()
        self.left_indent_base = None
        self.left_indent_mod = None
        self.list_reminder = (
            "Списку предшествует абзац текста с флагом \"не отрывать от следующего\"\n",
            "Нумерованный список:\n",
            "   Элемент списка начинается с заглавной буквы\n",
            "   Элемент списка заканчивается точкой\n",
            "   В качестве номера используется арабская цифра\n",
            "Маркированный список:\n",
            "   Элемент списка начинается с заглавной буквы\n",
            "   Элемент списка заканчивается точкой с запятой\n",
            "   В качестве маркера используется тире\n",
            "Выступ первой строки 0.75 см\n",
            "Отступ вычисляется по формуле: 1.25+0.75*(Уровень_элемента - 1)\n",
        )


class TableParagraphChecklist(ParagraphChecklist):
    def __init__(self):
        super().__init__()
        self.vert_alignment = None


class ImageChecklist(BaseChecklist):
    def __init__(self):
        super().__init__()
        self.alignment = None
        self.keep_with_next = None
        self.space_before = None
        self.space_after = None
        self.first_line_indent = None
        self.line_spacing = None


class MarginsChecklist(BaseChecklist):
    def __init__(self):
        super().__init__()
        self.top_margin = None
        self.bottom_margin = None
        self.left_margin = None
        self.right_margin = None
        self.orientation = None


regex_transform = {
    "Таблица <N> - <Название>.": "^Таблица [1-9]*(.[1-9])* (–|-) .*$",
    "<Название>.": ".*",
    "Рисунок <N> - <Название>.": "^Рисунок [1-9]*(.[1-9])* (–|-) .*$"
}

param_to_comment = {
    "format_regex": "Формат",
    "font_name": "Тип шрифта",
    "font_size": "Размер шрифта",
    "font_bald": "Полужирный",
    "font_italic": "Курсив",
    "font_underline": "Подчёркивание",
    "font_color": "Цвет текста",
    "font_back_color": "Цвет подчёркивания",
    "alignment": "Выравнивание",
    "vert_alignment": "Выравнивание по вертикали",
    "keep_with_next": "Не отрывать от следующего",
    "page_break_before": "С новой страницы",
    "space_before": "Верт отступ перед абзацем",
    "space_after": "Верт отступ после абзаца",
    "left_indent": "Отступ слева",
    "right_indent": "Отступ справа",
    "first_line_indent": "Отступ первой строки",
    "line_spacing": "Межстрочный интервал",
    "is_list": "Список",
    "top_margin": "Верхнее поле",
    "bottom_margin": "Нижнее поле",
    "left_margin": "Левое поле",
    "right_margin": "Правое поле",
    "orientation": "Ориентация",
}
var_to_comment = {
    True: "Да",
    False: "Нет",
    None: "Нет",
}
alignment_to_comment = {
    WD_PARAGRAPH_ALIGNMENT.LEFT: "Левый край",
    WD_PARAGRAPH_ALIGNMENT.CENTER: "По центру",
    WD_PARAGRAPH_ALIGNMENT.RIGHT: "Правый край",
    WD_PARAGRAPH_ALIGNMENT.JUSTIFY: "По ширине",
}
vert_alignment_to_comment = {
    WD_ALIGN_VERTICAL.TOP: "Верхний край",
    WD_ALIGN_VERTICAL.CENTER: "По центру",
    WD_ALIGN_VERTICAL.BOTTOM: "Нижний край",
    WD_ALIGN_VERTICAL.BOTH: "Ребята, не стоит вскрывать эту тему...",
}
orientation_to_comment = {
    WD_ORIENTATION.PORTRAIT: "Книжная",
    WD_ORIENTATION.LANDSCAPE: "Альбомная"
}

modifications = {
    "paragraph_after_table": {
        "space_before": 13.0
    }
}


class DocumentParser:
    def __init__(self):
        self.text_checklist = ParagraphChecklist()
        self.heading1_checklist = ParagraphChecklist()
        self.heading2_checklist = ParagraphChecklist()
        self.heading3_checklist = ParagraphChecklist()
        self.list_checklist = ListChecklist()
        self.table_headings_checklist = TableParagraphChecklist()
        self.table_text_checklist = TableParagraphChecklist()
        self.table_name_checklist = ParagraphChecklist()
        self.text_after_table_checklist = ParagraphChecklist()
        self.image_checklist = ImageChecklist()
        self.image_name_checklist = ParagraphChecklist()
        self.margins_checklist = MarginsChecklist()
        self.set_settings(default_text_checklist, default_heading1_checklist, default_heading2_checklist,
                          default_heading3_checklist, default_table_name_checklist, default_table_headings_checklist,
                          default_table_text_checklist, default_list_checklist, default_margins_checklist,
                          default_image_checklist, default_image_name_checklist)

        self.enable_optional_settings = {
            "table_headings_top": True,
            "table_headings_left": False,
            "table_title": True,
            "paragraph_after_table": True,
            "enable_pic_title": True,
            "list_reminder": True,
        }

    def set_enable_optional_settings(self, elements: dict):
        if elements is not None:
            for par, value in elements.items():
                if par in self.enable_optional_settings:
                    self.enable_optional_settings[par] = value

    def set_settings(self, text_check=None, h1_check=None, h2_check=None, h3_check=None, table_name_check=None,
                     table_heading_check=None, table_text_check=None, list_check=None,
                     page_check=None, pic_check=None, pic_name_check=None):
        if isinstance(text_check, dict):
            self.text_checklist.set_settings(text_check)
            self.text_after_table_checklist.set_settings(text_check)
            self.text_after_table_checklist.space_before = 13.0  # !!!!
            self.list_checklist.set_settings(text_check)
        if isinstance(h1_check, dict):
            self.heading1_checklist.set_settings(h1_check)
        if isinstance(h2_check, dict):
            self.heading2_checklist.set_settings(h2_check)
        if isinstance(h3_check, dict):
            self.heading3_checklist.set_settings(h3_check)
        if isinstance(table_name_check, dict):
            self.table_name_checklist.set_settings(table_name_check)
        if isinstance(table_heading_check, dict):
            self.table_headings_checklist.set_settings(table_heading_check)
        if isinstance(table_text_check, dict):
            self.table_text_checklist.set_settings(table_text_check)
        if isinstance(list_check, dict):
            self.list_checklist.set_settings(list_check)
            # if "left_indent_base" in list_check:
                # self.list_checklist.left_indent_base -= 0.75
        if isinstance(page_check, dict):
            self.margins_checklist.set_settings(page_check)
        if isinstance(pic_check, dict):
            self.image_checklist.set_settings(pic_check)
        if isinstance(pic_name_check, dict):
            self.image_name_checklist.set_settings(pic_name_check)

    @staticmethod
    def is_list(paragraph):
        return len(paragraph._element.xpath('./w:pPr/w:numPr')) > 0

    @staticmethod
    def get_error_comment(checklist, received: dict):
        attributes = inspect.getmembers(checklist, lambda a: not (inspect.isroutine(a)))
        expected = {a[0]: a[1] for a in attributes if not (a[0].startswith('__') and a[0].endswith('__'))}
        comments = set()
        for key, val in expected.items():
            comparison_res = True
            comment = ""
            if key in received.keys() and val is not None:
                # Отдельная обработка формата через RegEx
                if key == 'format_regex':
                    comparison_res = re.fullmatch(fr'{regex_transform[val]}', received[key])
                    if not comparison_res:
                        comment += f"{param_to_comment[key]}: "
                        comment += f"{val}; "
                        comment += f"Получено: {received[key]}.\n"
                        comments.add(comment)
                        continue
                elif "list_level" in received.keys() and key == "left_indent":
                    expected_indent = expected["left_indent_base"] + expected["left_indent_mod"] * received["list_level"]
                    if received["left_indent"] != expected_indent:
                        comment += f"{param_to_comment[key]}: "
                        comment += f"Ожидалось: {expected_indent}; "
                        comment += f"Получено: {received[key]}.\n"
                        comments.add(comment)
                        continue
                else:
                    comparison_res = val == received[key]
                if not comparison_res:
                    comment += f"{param_to_comment[key]}: "
                    # У выравниваний значения приравниваются к 0..3, поэтому нужен костыль
                    if key == "alignment":
                        comment += f"Ожидалось: {alignment_to_comment[val]}; "
                        comment += f"Получено: {alignment_to_comment[received[key]]}.\n"
                    elif key == "vert_alignment":
                        comment += f"Ожидалось: {vert_alignment_to_comment[val]}; "
                        comment += f"Получено: {vert_alignment_to_comment[received[key]]}.\n"
                    elif key == "orientation":
                        comment += f"Ожидалось: {orientation_to_comment[val]}; "
                        comment += f"Получено: {orientation_to_comment[received[key]]}.\n"
                    else:
                        if val is None or isinstance(val, bool):
                            comment += f"Ожидалось: {var_to_comment[val]}; "
                        else:
                            comment += f"Ожидалось: {val}; "
                        if received[key] is None or isinstance(received[key], bool):
                            comment += f"Получено: {var_to_comment[received[key]]}.\n"
                        else:
                            comment += f"Получено: {received[key]}.\n"
                    comments.add(comment)
        return comments

    def get_run_properties(self, document, p, run):
        """ Параметры параграфа """
        st = p.style
        formatting = p.paragraph_format  # Формат параграфа
        st_formatting = p.style.paragraph_format  # Формат параграфа, заданный стилем
        default = st.base_style if st.base_style else document.styles["Normal"]  # Стиль по умолчанию
        def_formatting = default.paragraph_format

        # Пришлось всему написать is not None, т.к может встретиться 0
        # Название шрифта
        font_name = run.font.name if run.font.name is not None else \
            st.font.name if st.font.name is not None else \
                default.font.name if default.font.name is not None else "Calibri"
        # Размер шрифта
        font_size = run.font.size.pt if run.font.size is not None else \
            st.font.size.pt if st.font.size is not None else \
                default.font.size.pt if default.font.size is not None else 11
        # Полужирный
        font_bald = run.bold if run.bold is not None else \
            st.font.bold if st.font.bold is not None else \
                default.font.bold if default.font.bold is not None else False
        # Курсив
        font_italic = run.italic if run.italic is not None else \
            st.font.italic if st.font.italic is not None else \
                default.font.italic if default.font.italic is not None else False
        # Подчёркивание
        font_underline = run.underline if run.underline is not None else \
            st.font.underline if st.font.underline is not None else \
                default.font.underline if default.font.underline is not None else False
        # Цвет текста
        font_color = run.font.color.rgb if run.font.color.rgb is not None else \
            st.font.color.rgb if st.font.color.rgb is not None else \
                default.font.color.rgb if default.font.color.rgb is not None else False

        if font_color and font_color[0] == font_color[1] == font_color[2] == 0:
            font_color = False
        # Цвет подчёркивания текста
        try:
            font_back_color = run.font.highlight_color if run.font.highlight_color is not None else \
                st.font.highlight_color if st.font.highlight_color is not None else \
                    default.font.highlight_color if default.font.highlight_color is not None else False
        except:
            font_back_color = False

        # Выравнивание (лево/центр/право/ширина)
        alignment = formatting.alignment if formatting.alignment is not None else \
            st_formatting.alignment if st_formatting.alignment is not None else \
                def_formatting.alignment if def_formatting.alignment is not None else WD_PARAGRAPH_ALIGNMENT.LEFT
        # Не отрывать от следующего
        keep_with_next = formatting.keep_with_next if formatting.keep_with_next is not None else \
            st_formatting.keep_with_next if st_formatting.keep_with_next is not None else \
                def_formatting.keep_with_next if def_formatting.keep_with_next is not None else False
        # С новой страницы
        page_break_before = formatting.page_break_before if formatting.page_break_before is not None else \
            st_formatting.page_break_before if st_formatting.page_break_before is not None else \
                def_formatting.page_break_before if def_formatting.page_break_before is not None else False
        # Верт отступ перед абзацем
        space_before = formatting.space_before.pt if formatting.space_before is not None else \
            st_formatting.space_before.pt if st_formatting.space_before is not None else \
                def_formatting.space_before.pt if def_formatting.space_before is not None else 0.0
        # Верт отступ перед абзацем
        space_after = formatting.space_after.pt if formatting.space_after is not None else \
            st_formatting.space_after.pt if st_formatting.space_after is not None else \
                def_formatting.space_after.pt if def_formatting.space_after is not None else 0.0
        # Отступ слева
        left_indent = round(formatting.left_indent.cm, 2) if formatting.left_indent is not None else \
            round(st_formatting.left_indent.cm, 2) if st_formatting.left_indent is not None else \
                round(def_formatting.left_indent.cm, 2) if def_formatting.left_indent is not None else 0.0
        # Отступ справа
        right_indent = round(formatting.right_indent.cm, 2) if formatting.right_indent is not None else \
            round(st_formatting.right_indent.cm, 2) if st_formatting.right_indent is not None else \
                round(def_formatting.right_indent.cm, 2) if def_formatting.right_indent is not None else 0.0
        # Красная строка
        first_line_indent = round(formatting.first_line_indent.cm, 2) if formatting.first_line_indent is not None else \
            round(st_formatting.first_line_indent.cm, 2) if st_formatting.first_line_indent is not None else \
                round(def_formatting.first_line_indent.cm, 2) if def_formatting.first_line_indent is not None else 0.0
        # Межстрочный интервал
        line_spacing = formatting.line_spacing if formatting.line_spacing is not None else \
            st_formatting.line_spacing if st_formatting.line_spacing is not None else \
                def_formatting.line_spacing if def_formatting.line_spacing is not None else 1.0

        is_list = self.is_list(p)

        paragraph_stats = {
            "font_name": font_name,
            "font_size": font_size,
            "font_bald": font_bald,
            "font_italic": font_italic,
            "font_underline": font_underline,
            "font_color": font_color,
            "font_back_color": font_back_color,
            "alignment": alignment,
            "keep_with_next": keep_with_next,
            "page_break_before": page_break_before,
            "space_before": space_before,
            "space_after": space_after,
            "left_indent": left_indent,
            "right_indent": right_indent,
            "first_line_indent": first_line_indent,
            "line_spacing": line_spacing,
            "is_list": is_list,
        }
        return paragraph_stats

    # This function extracts the tables and paragraphs from the document object
    @staticmethod
    def iter_block_items(parent):
        """
        Yield each paragraph and table child within *parent*, in document order.
        Each returned value is an instance of either Table or Paragraph. *parent*
        would most commonly be a reference to a main Document object, but
        also works for a _Cell object, which itself can contain paragraphs and tables.
        """
        if isinstance(parent, doctwo):
            parent_elm = parent.element.body
        elif isinstance(parent, _Cell):
            parent_elm = parent._tc
        else:
            raise ValueError("something's not right")

        for child in parent_elm.iterchildren():
            if isinstance(child, CT_P):
                yield Paragraph(child, parent)
            elif isinstance(child, CT_Tbl):
                yield Table(child, parent)

    @staticmethod
    def set_comment_name(name, prev_name, is_locked):
        if not is_locked:
            return name
        return prev_name

    def parse_margins(self, document):
        sections = document.sections
        paragraph_stats = {}
        stats_to_compare = self.margins_checklist
        margins_comments = {}
        for section_cnt, section in enumerate(sections):
            paragraph_stats["top_margin"] = round(section.top_margin.cm, 2)
            paragraph_stats["bottom_margin"] = round(section.bottom_margin.cm, 2)
            paragraph_stats["left_margin"] = round(section.left_margin.cm, 2)
            paragraph_stats["right_margin"] = round(section.right_margin.cm, 2)
            paragraph_stats["orientation"] = section.orientation
            section_comments = self.get_error_comment(stats_to_compare, paragraph_stats)
            if len(section_comments) > 0:
                margins_comments[f"Раздел {section_cnt + 1}:"] = section_comments
        return margins_comments

    def parse_document(self, filename):
        document = Document(filename)
        comment_count = 0
        written_comments = []
        paragraph_to_comment = ""
        comment_to_send = set()
        comment_type = "System"
        image_name_check = 0
        text_after_table_check = 0
        first_paragraph_not_reached = True
        table_name_not_found = False
        list_start = False
        paragraphs = document.paragraphs
        for block in self.iter_block_items(document):
            next_paragraph_to_comment = ""
            next_comment_to_send = set()
            next_comment_type = "System"
            lock_comment_name = False
            stats_to_compare = {}
            run_comments = set()
            if 'text' in str(block):
                if self.is_list(block):
                    if not list_start:
                        # print(block.text)
                        list_start = True
                        block.add_comment(''.join(self.list_checklist.list_reminder),
                                          author="Напоминание о формате списков")
                        written_comments.append(self.list_checklist.list_reminder)
                elif block.text != "":
                    list_start = False

                if first_paragraph_not_reached:
                    first_paragraph_not_reached = False
                    margin_comments = self.parse_margins(document)
                    for section, comments in margin_comments.items():
                        block.add_comment(''.join(comments), author=section)
                        written_comments.append(comments)
                for run in block.runs:
                    xmlstr = str(run.element.xml)
                    my_namespaces = dict([node for _, node in ET.iterparse(StringIO(xmlstr), events=['start-ns'])])
                    root = ET.fromstring(xmlstr)
                    if 'pic:pic' in xmlstr:  # Image run
                        for pic in root.findall('.//pic:pic', my_namespaces):
                            stats_to_compare = self.image_checklist
                            image_name_check = 2
                            paragraph_stats = self.get_run_properties(document, block, run)
                            run_comments = self.get_error_comment(stats_to_compare, paragraph_stats)
                            for comment in run_comments:
                                next_comment_to_send.add(comment)
                            next_comment_type = self.set_comment_name("Рисунок", next_comment_type, lock_comment_name)
                            lock_comment_name = True
                            next_paragraph_to_comment = block
                            # print("Image:", block)
                    else:  # Paragraph run
                        if run.text.strip() != "":

                            paragraph_stats = self.get_run_properties(document, block, run)
                            if block.style.name.startswith("Heading 1"):
                                paragraph_stats["is_list"] = self.is_list(block)
                                stats_to_compare = self.heading1_checklist
                                next_comment_type = self.set_comment_name("Заголовок 1", next_comment_type,
                                                                          lock_comment_name)
                            elif block.style.name.startswith("Heading 2"):
                                paragraph_stats["is_list"] = self.is_list(block)
                                stats_to_compare = self.heading2_checklist
                                next_comment_type = self.set_comment_name("Заголовок 2", next_comment_type,
                                                                          lock_comment_name)
                            elif block.style.name.startswith("Heading 3"):
                                paragraph_stats["is_list"] = self.is_list(block)
                                stats_to_compare = self.heading3_checklist
                                next_comment_type = self.set_comment_name("Заголовок 3", next_comment_type,
                                                                          lock_comment_name)
                            elif self.is_list(block):
                                stats_to_compare = self.list_checklist
                                if block.list_info[0]:
                                    paragraph_stats["list_level"] = block.list_info[2]
                                else:
                                    paragraph_stats["list_level"] = 0

                                next_comment_type = self.set_comment_name("Элемент списка", next_comment_type,
                                                                          lock_comment_name)

                            elif self.enable_optional_settings["enable_pic_title"] and image_name_check:
                                stats_to_compare = self.image_name_checklist
                                paragraph_stats["format_regex"] = block.text
                                next_comment_type = self.set_comment_name("Название рисунка", next_comment_type,
                                                                          lock_comment_name)
                                lock_comment_name = True
                            elif self.enable_optional_settings["paragraph_after_table"] and text_after_table_check:
                                stats_to_compare = self.text_after_table_checklist
                                next_comment_type = self.set_comment_name("Параграф после таблицы", next_comment_type,
                                                                          lock_comment_name)
                                lock_comment_name = True
                            else:
                                stats_to_compare = self.text_checklist
                                next_comment_type = self.set_comment_name("Абзац", next_comment_type, lock_comment_name)
                            if paragraph_stats["first_line_indent"] < 0:
                                paragraph_stats["left_indent"] += paragraph_stats["first_line_indent"]
                            run_comments = self.get_error_comment(stats_to_compare, paragraph_stats)
                            for comment in run_comments:
                                next_comment_to_send.add(comment)
                            next_paragraph_to_comment = block
            elif 'table' in str(block):  # Table
                text_after_table_check = 2
                """Перепроверка названия таблица"""
                if self.enable_optional_settings["table_title"]:
                    if isinstance(paragraph_to_comment, Paragraph):
                        comment_to_send.clear()
                        stats_to_compare = self.table_name_checklist
                        paragraph_stats["format_regex"] = paragraph_to_comment.text
                        run_comments = self.get_error_comment(stats_to_compare, paragraph_stats)
                        # print(run_comments)
                        for comment in run_comments:
                            comment_to_send.add(comment)
                        comment_type = "Название таблицы"
                    else:
                        next_comment_to_send.add("Нет названия таблицы!\n")

                # print("Table:", block)
                paragraph_chosen = False
                for row_count, row in enumerate(block.rows):
                    col_count = 0
                    for cell in row.cells:
                        if cell.text != "":
                            for par in cell.paragraphs:
                                if not paragraph_chosen:
                                    next_paragraph_to_comment = par
                                    paragraph_chosen = True
                                for run in par.runs:
                                    paragraph_stats = self.get_run_properties(document, par, run)
                                    vert_alignment = cell.vertical_alignment if cell.vertical_alignment else \
                                        WD_ALIGN_VERTICAL.TOP
                                    paragraph_stats['vert_alignment'] = vert_alignment
                                    if self.enable_optional_settings["table_headings_top"] and row_count == 0:
                                        stats_to_compare = self.table_headings_checklist
                                    elif self.enable_optional_settings["table_headings_left"] and col_count == 0:
                                        stats_to_compare = self.table_headings_checklist
                                    else:
                                        stats_to_compare = self.table_text_checklist

                                    run_comments = self.get_error_comment(stats_to_compare, paragraph_stats)
                                    for comment in run_comments:
                                        next_comment_to_send.add(comment)
                                    next_comment_type = self.set_comment_name("Таблица", next_comment_type,
                                                                              lock_comment_name)
                        col_count += 1

            if isinstance(paragraph_to_comment, Paragraph) and len(comment_to_send) > 0:
                paragraph_to_comment.add_comment(''.join(comment_to_send), author=comment_type)
                comment_count += 1
                written_comments.append(comment_to_send)

            paragraph_to_comment = next_paragraph_to_comment
            comment_to_send = next_comment_to_send
            comment_type = next_comment_type

            image_name_check -= 1 if image_name_check > 0 else 0
            text_after_table_check -= 1 if text_after_table_check > 0 else 0

        if isinstance(paragraph_to_comment, Paragraph) and len(comment_to_send) > 0:
            paragraph_to_comment.add_comment(''.join(comment_to_send), author=comment_type)
            comment_count += 1
            written_comments.append(comment_to_send)

        basename, extension = os.path.splitext(os.path.basename(filename))
        try:
            Path('./Results').mkdir(parents=True, exist_ok=False)
        except FileExistsError:
            pass
        count = 0
        while True:
            try:
                if count == 0:
                    document.save(f"Results/{basename}_Проверенный{extension}")
                else:
                    document.save(f"Results/{basename}_Проверенный_{count}{extension}")
                break
            except PermissionError:
                count += 1
        return written_comments


if __name__ == '__main__':
    parser = DocumentParser()
    '''parser.set_settings(default_text_checklist, default_heading1_checklist, default_heading2_checklist,
                        default_heading3_checklist, default_table_name_checklist, default_table_headings_checklist,
                        default_table_text_checklist, default_list_checklist, default_margins_checklist,
                        default_image_checklist, default_image_name_checklist)'''

    filename = "testDocs/testTables.docx"
    parser.parse_document(filename)
