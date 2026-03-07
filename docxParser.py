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
from docx.comments import Comment
from io import StringIO
import re

from default_settings import *
from constants import *
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

        
        self.doc_comments = {}

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
        comments = []
        
        for key, val in expected.items():
            comparison_res = True
            comment = []
            if key in received.keys() and val is not None:
                # Отдельная обработка формата через RegEx
                if key == 'format_regex':
                    comparison_res = re.fullmatch(fr'{REGEX_TRANSFORM[val]}', received[key])
                    if not comparison_res:
                        comment = [PARAM_TO_COMMENT[key], val, received[key]]
                        comments.append(comment)
                        continue
                elif "list_level" in received.keys() and key == "left_indent":
                    expected_indent = expected["left_indent_base"] + expected["left_indent_mod"] * received["list_level"]
                    if received["left_indent"] != expected_indent:
                        comment = [PARAM_TO_COMMENT[key], expected_indent, received[key]]
                        comments.append(comment)
                        continue
                else:
                    comparison_res = val == received[key]
                if not comparison_res:
                    comment = [PARAM_TO_COMMENT[key]]
                    # У выравниваний значения приравниваются к 0..3, поэтому нужен костыль
                    if key == "alignment":
                        comment += [ALIGNMENT_TO_COMMENT[val], ALIGNMENT_TO_COMMENT[received[key]]]
                    elif key == "vert_alignment":
                        comment += [VERT_ALIGNMENT_TO_COMMENT[val], VERT_ALIGNMENT_TO_COMMENT[received[key]]]
                    elif key == "orientation":
                        comment += [ORIENTATION_TO_COMMENT[val], ORIENTATION_TO_COMMENT[received[key]]]
                    else:
                        if val is None or isinstance(val, bool):
                            comment += [VAR_TO_COMMENT[val]]
                        else:
                            comment += [val]
                        if received is None or isinstance(received[key], bool):
                            comment += [VAR_TO_COMMENT[received[key]]]
                        else:
                            comment += [received[key]]
                    comments.append(comment)
        return comments
    
    @staticmethod
    def create_comment(document, runs, element, comments: list):
        if len(comments) > 0:
            comment = document.add_comment(
                runs=runs,
                text='',
                author=element
            )
            for cmt in comments:
                parameter, expected_value, received_value = cmt
                cmt_para = comment.add_paragraph()
                cmt_para.add_run(f"{parameter}: ").bold = True 
                cmt_para.add_run(f"{received_value} | ")
                cmt_para.add_run(f"({expected_value}).")
            return comment


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



    # ====================== НОВЫЕ ВСПОМОГАТЕЛЬНЫЕ МЕТОДЫ ======================
    @staticmethod
    def get_heading_level(paragraph):
        """Возвращает уровень заголовка (1-9) или None. Покрывает ВСЕ крайние случаи."""
        def extract_level_from_style(style):
            if not style:
                return None
            name = style.name or ""
            if name.startswith("Heading "):
                try:
                    return int(name.split()[-1])
                except (ValueError, IndexError):
                    pass
            style_id = getattr(style, 'style_id', None) or ""
            if style_id.startswith("Heading"):
                try:
                    return int(style_id.replace("Heading", "").strip())
                except ValueError:
                    pass
            if name.startswith("Heading") and len(name) > 7:
                try:
                    return int(name[7:])
                except ValueError:
                    pass
            return None

        # Быстрый путь
        level = extract_level_from_style(paragraph.style)
        if level is not None:
            return level

        # Цепочка наследования стилей (самый частый крайний случай!)
        current = paragraph.style
        depth = 0
        while current and depth < 20:
            level = extract_level_from_style(current)
            if level is not None:
                return level
            current = current.base_style
            depth += 1

        # outlineLvl (используется в некоторых шаблонах и экспортах из Google Docs)
        outline = paragraph._p.xpath('./w:pPr/w:outlineLvl/@w:val')
        if outline:
            try:
                lvl = int(outline[0])
                return lvl if 1 <= lvl <= 9 else None
            except (ValueError, IndexError):
                pass
        return None

    @staticmethod
    def is_list_item(paragraph):
        """Возвращает (is_list: bool, level: int). Работает даже если numPr отсутствует."""
        if not isinstance(paragraph, Paragraph):
            return False, 0

        # Самый надёжный способ — XML
        pPr = paragraph._p.pPr
        if pPr is not None and pPr.numPr is not None:
            ilvl = paragraph._p.xpath('.//w:ilvl/@w:val')
            level = int(ilvl[0]) + 1 if ilvl else 1
            return True, level

        # Стиль List / List Paragraph / пользовательские списки
        style_name = (paragraph.style.name or "").lower()
        if any(x in style_name for x in ("list", "нумерованный", "маркированный", "bullet", "number")):
            return True, 1

        return False, 0

    @staticmethod
    def has_image(paragraph):
        """Проверка наличия изображения (pic + drawing — покрывает все версии Word)"""
        for run in paragraph.runs:
            xml = str(run.element.xml).lower()
            if 'pic:pic' in xml:
                return True
        return False
    
    @staticmethod
    def has_image_run(run):
        """Проверяет, содержит ли именно этот run изображение"""
        if not run:
            return False
        xml = str(run.element.xml).lower()
        return 'pic:pic' in xml

    def collect_paragraph_errors(self, paragraph, checklist, document):
        """Проверяет КАЖДЫЙ run в параграфе и собирает ВСЕ ошибки.
        Возвращает список ошибок для одного общего комментария (дубликаты по параметру удаляются)."""
        if not isinstance(paragraph, Paragraph) or not paragraph.runs:
            return []

        all_errors = []
        seen = set()

        for run in paragraph.runs:
            # Пропускаем пустые run-ы (часто бывают пробелы/табуляция)
            if not run.text or not run.text.strip():
                continue

            stats = self.get_run_properties(document, paragraph, run)
            errors = self.get_error_comment(checklist, stats)

            for err in errors:
                err_key = tuple(str(x) for x in err)  # дедупликация
                if err_key not in seen:
                    seen.add(err_key)
                    all_errors.append(err)

        return all_errors

    def _classify_paragraph(self, paragraph):
        """Единая функция классификации — покрывает абсолютно все случаи."""
        if not isinstance(paragraph, Paragraph):
            return "unknown", None

        heading_level = self.get_heading_level(paragraph)
        is_list, list_level = self.is_list_item(paragraph)
        has_img = self.has_image(paragraph)

        if heading_level == 1:
            return "heading1", None
        if heading_level == 2:
            return "heading2", None
        if heading_level == 3:
            return "heading3", None
        if heading_level is not None:           # заголовки 4+ уровня
            return f"heading{heading_level}", None

        if is_list:
            return "list", list_level

        if has_img:
            return "image", None

        # Название рисунка/таблицы (стили Caption + наши настройки)
        style_name = (paragraph.style.name or "").lower()
        if "caption" in style_name or "подпись" in style_name:
            return "caption", None

        return "text", None

    # =========================================================================
    
    def parse_document(self, filename):
        document = Document(filename)
        written_comments = []
        comment_count = 0

        # Состояния
        image_name_check = 0
        text_after_table_check = 0
        first_paragraph_not_reached = True

        # Накопители для комментария
        current_paragraph = None
        current_errors = []
        current_author = "System"
        current_target_runs = []

        for block in self.iter_block_items(document):
            block_errors = []
            block_author = "System"
            target_paragraph = None
            target_runs = []

            # Проверка полей документа (один раз в начале)
            if first_paragraph_not_reached and isinstance(block, Paragraph):
                first_paragraph_not_reached = False
                margin_comments = self.parse_margins(document)
                for section_title, errors in margin_comments.items():
                    if errors:
                        self.create_comment(document, block.runs, section_title, errors)
                        comment_count += 1
                        written_comments.append([section_title, errors])

            # ====================== ПАРАГРАФ ======================
            if isinstance(block, Paragraph):
                p = block
                target_paragraph = p
                block_type, extra = self._classify_paragraph(p)
                
                if len(p.runs) == 0:
                    continue

                # Определяем чек-лист и автора
                if block_type.startswith("heading"):
                    level = int(block_type.replace("heading", ""))
                    checklist = getattr(self, f"heading{level}_checklist", self.heading3_checklist)
                    block_author = f"Заголовок {level}"

                elif block_type == "list":
                    checklist = self.list_checklist
                    block_author = "Элемент списка"

                elif self.has_image(p):                                      # ← РИСУНОК
                    checklist = self.image_checklist
                    block_author = "Рисунок"
                    image_name_check = 2

                    # Привязываем комментарий именно к run-у с картинкой
                    image_run = next((r for r in p.runs if self.has_image_run(r)), None)
                    target_runs = [image_run] if image_run else p.runs

                elif image_name_check > 0 or block_type == "caption":        # ← НАЗВАНИЕ РИСУНКА
                    checklist = self.image_name_checklist
                    block_author = "Название рисунка"
                    stats = {"format_regex": p.text.strip()}
                    block_errors.extend(self.get_error_comment(checklist, stats))
                    target_runs = p.runs

                elif self.enable_optional_settings.get("paragraph_after_table", False) and text_after_table_check > 0:
                    checklist = self.text_after_table_checklist
                    block_author = "Параграф после таблицы"
                else:
                    checklist = self.text_checklist
                    block_author = "Абзац"

                # Главная проверка — все run-ы
                if 'checklist' in locals() and not block_errors:   # для caption уже заполнили
                    block_errors = self.collect_paragraph_errors(p, checklist, document)

                # По умолчанию комментируем весь параграф
                if not target_runs:
                    target_runs = p.runs

            # ====================== ТАБЛИЦА ======================
            elif isinstance(block, Table):
                text_after_table_check = 2
                block_author = "Таблица"

                # Проверка названия таблицы
                if self.enable_optional_settings.get("table_title", True) and current_paragraph:
                    stats = {"format_regex": current_paragraph.text.strip()}
                    name_errors = self.get_error_comment(self.table_name_checklist, stats)
                    if name_errors:
                        current_errors.extend(name_errors)
                        current_author = "Название таблицы"

                # Проверка содержимого ячеек
                for r_idx, row in enumerate(block.rows):
                    for c_idx, cell in enumerate(row.cells):
                        if not cell.text.strip():
                            continue
                        for cell_p in cell.paragraphs:
                            if target_paragraph is None:
                                target_paragraph = cell_p
                                target_runs = cell_p.runs[:1] if cell_p.runs else []

                            cell_stats = self.get_run_properties(document, cell_p,
                                                                 cell_p.runs[0] if cell_p.runs else None)
                            cell_stats["vert_alignment"] = cell.vertical_alignment or WD_ALIGN_VERTICAL.TOP

                            checklist = self.table_headings_checklist if \
                                (self.enable_optional_settings.get("table_headings_top", True) and r_idx == 0) or \
                                (self.enable_optional_settings.get("table_headings_left", False) and c_idx == 0) \
                                else self.table_text_checklist

                            cell_errors = self.get_error_comment(checklist, cell_stats)
                            if cell_errors:
                                block_errors.extend(cell_errors)

            # ====================== СОХРАНЕНИЕ КОММЕНТАРИЯ ======================
            if current_paragraph and current_errors:
                runs_to_use = current_target_runs if current_target_runs else current_paragraph.runs
                self.create_comment(document, runs_to_use, current_author, current_errors)
                comment_count += 1
                written_comments.append([current_author, current_errors])

            # Переключаемся на следующий блок
            current_paragraph = target_paragraph
            current_errors = block_errors
            current_author = block_author
            current_target_runs = target_runs

            image_name_check = max(0, image_name_check - 1)
            text_after_table_check = max(0, text_after_table_check - 1)

        # Последний блок документа
        if current_paragraph and current_errors:
            runs_to_use = current_target_runs if current_target_runs else current_paragraph.runs
            self.create_comment(document, runs_to_use, current_author, current_errors)
            comment_count += 1
            written_comments.append([current_author, current_errors])

        # Сохранение результата
        basename = os.path.splitext(os.path.basename(filename))[0]
        out_path = f"Results/{basename}_Проверенный.docx"
        count = 1
        while os.path.exists(out_path):
            out_path = f"Results/{basename}_Проверенный_{count}.docx"
            count += 1

        os.makedirs("Results", exist_ok=True)
        document.save(out_path)
        return written_comments


if __name__ == '__main__':
    parser = DocumentParser()
    '''parser.set_settings(default_text_checklist, default_heading1_checklist, default_heading2_checklist,
                        default_heading3_checklist, default_table_name_checklist, default_table_headings_checklist,
                        default_table_text_checklist, default_list_checklist, default_margins_checklist,
                        default_image_checklist, default_image_name_checklist)'''

    filename = "testDocs/testTables.docx"
    parser.parse_document(filename)
