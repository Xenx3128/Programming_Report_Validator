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

from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_UNDERLINE
from docx.text.run import Run

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
        self.font_highlight = None  # !!!
        self.alignment = None
        self.keep_with_next = None
        self.page_break_before = None
        self.space_before = None
        self.space_after = None
        self.left_indent = None
        self.right_indent = None
        self.first_line_indent = None
        self.line_spacing = None
        self.header_numbering = None



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
            if element == 'Таблица':
                cmt_para = comment.add_paragraph()
                cmt_para.add_run("Выведена только первая ошибка в каждой категории\n").bold = True 
            for cmt in comments:
                parameter, expected_value, received_value = cmt
                cmt_para = comment.add_paragraph()
                cmt_para.add_run(f"{parameter}: ").bold = True 
                cmt_para.add_run(f"{received_value} | ")
                cmt_para.add_run(f"({expected_value}).")
            return comment




    def get_run_properties(self, document, paragraph, run):
        """
        Возвращает свойства для данного run
        """
        if run is None:
            # Пустой run или параграф без run-ов — возвращаем "нейтральные" значения
            return {
                "font_name": "Unknown (no run)",
                "font_size": None,
                "font_bold": False,
                "font_italic": False,
                "font_underline": False,
                "font_color": False,
                "font_highlight": False,
                "alignment": paragraph.alignment,
                "keep_with_next": paragraph.paragraph_format.keep_with_next,
                "page_break_before": paragraph.paragraph_format.page_break_before,
                "space_before": paragraph.paragraph_format.space_before.pt if paragraph.paragraph_format.space_before else 0.0,
                "space_after": paragraph.paragraph_format.space_after.pt if paragraph.paragraph_format.space_after else 0.0,
                "left_indent": paragraph.paragraph_format.left_indent.cm if paragraph.paragraph_format.left_indent else 0.0,
                "right_indent": paragraph.paragraph_format.right_indent.cm if paragraph.paragraph_format.right_indent else 0.0,
                "first_line_indent": paragraph.paragraph_format.first_line_indent.cm if paragraph.paragraph_format.first_line_indent else 0.0,
                "line_spacing": paragraph.paragraph_format.line_spacing,
                "is_list": False
            }

        # ── 1. Символьные свойства (run.font) ─────────────────────────────────────

        def resolve_font_property(getter_run, getter_style, default_value):
            """ Три-state разрешение: run → style chain → defaults """
            val = getter_run()
            if val is not None:
                return val

            # Идём по цепочке стилей параграфа
            style = paragraph.style
            while style is not None:
                val = getter_style(style)
                if val is not None:
                    return val
                style = style.base_style

            # Последний уровень — document defaults
            defaults = document.styles.element.xpath('./w:docDefaults/w:rPrDefault/w:rPr')
            if defaults:
                # Здесь нужно парсить XML, но для простоты часто возвращают default_value
                pass  # можно доработать при необходимости

            return default_value

        # ── Шрифт ──
        font_name = resolve_font_property(
            lambda: run.font.name,
            lambda s: s.font.name,
            "Calibri"   # или "Times New Roman" — зависит от твоих требований
        )

        # ── Размер шрифта ──
        font_size_pt = resolve_font_property(
            lambda: run.font.size.pt if run.font.size else None,
            lambda s: s.font.size.pt if s.font.size else None,
            11.0
        )

        # ── Bold / Italic / Underline (три-state) ──
        font_bold = resolve_font_property(
            lambda: run.bold,
            lambda s: s.font.bold,
            False
        )

        font_italic = resolve_font_property(
            lambda: run.italic,
            lambda s: s.font.italic,
            False
        )

        font_underline = resolve_font_property(
            lambda: run.underline,
            lambda s: s.font.underline,
            False
        )  # может быть WD_UNDERLINE.SINGLE и т.д.

        # ── Цвет текста ──
        def get_color_rgb(obj):
            if obj and obj.rgb:
                return obj.rgb
            return None

        font_color_rgb = resolve_font_property(
            lambda: get_color_rgb(run.font.color),
            lambda s: get_color_rgb(s.font.color),
            None
        )
        font_color = font_color_rgb if font_color_rgb else False
        if font_color_rgb and font_color_rgb == RGBColor(0, 0, 0):
            font_color = False  # часто чёрный = отсутствие явного цвета

        # ── Highlight (подсветка фона) ──
        
        ''' try:
            font_highlight = resolve_font_property(
            lambda: run.font.highlight_color,
            lambda s: s.font.highlight_color,
            None
        )
        except ValueError as e:
            font_highlight = None'''
            
        font_highlight = None

        # ── 2. Параграфные свойства (не зависят от run) ───────────────────────────

        pf = paragraph.paragraph_format
        st = paragraph.style
        def_style = st.base_style if st.base_style else document.styles["Normal"]
        def_pf = def_style.paragraph_format

        def resolve_para_property(getter):
            val = getter(pf)
            if val is not None:
                return val
            val = getter(st.paragraph_format)
            if val is not None:
                return val
            val = getter(def_pf)
            if val is not None:
                return val
            return None  # или дефолтное значение

        alignment = resolve_para_property(lambda fmt: fmt.alignment) or WD_PARAGRAPH_ALIGNMENT.LEFT
        keep_with_next = resolve_para_property(lambda fmt: fmt.keep_with_next) or False
        page_break_before = resolve_para_property(lambda fmt: fmt.page_break_before) or False

        space_before = (resolve_para_property(lambda fmt: fmt.space_before.pt if fmt.space_before else None) or 0.0)
        space_after  = (resolve_para_property(lambda fmt: fmt.space_after.pt  if fmt.space_after  else None) or 0.0)

        left_indent  = round(resolve_para_property(lambda fmt: fmt.left_indent.cm  if fmt.left_indent  else None) or 0.0, 2)
        right_indent = round(resolve_para_property(lambda fmt: fmt.right_indent.cm if fmt.right_indent else None) or 0.0, 2)
        first_line   = round(resolve_para_property(lambda fmt: fmt.first_line_indent.cm if fmt.first_line_indent else None) or 0.0, 2)

        line_spacing = resolve_para_property(lambda fmt: fmt.line_spacing) or 1.0

        is_list = self.is_list_item(paragraph)[0]

        return {
            "font_name": font_name,
            "font_size": font_size_pt,
            "font_bold": font_bold,
            "font_italic": font_italic,
            "font_underline": font_underline,
            "font_color": font_color,
            "font_highlight": font_highlight,  # было font_back_color
            "alignment": alignment,
            "keep_with_next": keep_with_next,
            "page_break_before": page_break_before,
            "space_before": space_before,
            "space_after": space_after,
            "left_indent": left_indent,
            "right_indent": right_indent,
            "first_line_indent": first_line,
            "line_spacing": line_spacing,
            "is_list": is_list
        }

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
    def is_list_item(paragraph: Paragraph, max_depth: int = 12):

        if not isinstance(paragraph, Paragraph):
            return False, None

        p = paragraph._p
        pPr = p.pPr

        # 1. Прямая нумерация в самом параграфе (самый надёжный случай)
        if pPr is not None and pPr.numPr is not None:
            ilvl_nodes = p.xpath('.//w:ilvl/@w:val')
            if ilvl_nodes:
                try:
                    level = int(ilvl_nodes[0]) + 1
                    return True, level
                except (ValueError, IndexError):
                    pass
            return True, 1  # ilvl нет → считаем первым уровнем

        # 2. Ищем в стиле и цепочке базовых стилей
        style = paragraph.style
        depth = 0

        while style is not None and depth < max_depth:
            style_elm = style._element
            if style_elm is not None:
                pPr_style = style_elm.pPr
                if pPr_style is not None and pPr_style.numPr is not None:
                    # Нашли numPr в стиле
                    ilvl_nodes_style = style_elm.xpath('.//w:ilvl/@w:val')
                    if ilvl_nodes_style:
                        try:
                            level = int(ilvl_nodes_style[0]) + 1
                            return True, level
                        except (ValueError, IndexError):
                            pass
                    return True, 1  # ilvl отсутствует → уровень 1

            style = style.base_style
            depth += 1

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
        written_comments = []           # для совместимости / отладки
        comment_count = 0

        # Переменные для отложенного комментирования
        prev_paragraph = None
        pending_errors = []
        pending_author = "System"

        image_name_check = 0
        text_after_table_check = 0
        first_paragraph_not_reached = True

        for block in self.iter_block_items(document):
            current_errors = []
            current_author = "System"
            current_paragraph = None

            if isinstance(block, Paragraph):
                p = block


                # Поля документа — один раз в начале
                if first_paragraph_not_reached and isinstance(block, Paragraph):
                    first_paragraph_not_reached = False
                    margin_comments = self.parse_margins(document)
                    for section_title, errors in margin_comments.items():
                        if errors:
                            self.create_comment(document, block.runs, section_title, errors)
                            comment_count += 1
                            written_comments.append([section_title, errors])

                # Классификация параграфа
                category, extra = self._classify_paragraph(p)

                # Выбор чек-листа и автора комментария
                if category in ("heading1", "heading2", "heading3") or category.startswith("heading"):
                    level = category.replace("heading", "") if category.startswith("heading") else "1"
                    checklist = getattr(self, f"heading{level}_checklist", self.heading3_checklist)
                    current_author = f"Заголовок {level}"
                elif category == "list":
                    checklist = self.list_checklist
                    current_author = "Элемент списка"
                elif category == "image":
                    checklist = self.image_checklist
                    current_author = "Рисунок"
                    image_name_check = 2
                elif category == "caption" and self.enable_optional_settings.get("enable_pic_title", False) and image_name_check > 0:
                    checklist = self.image_name_checklist
                    current_author = "Название рисунка"
                elif self.enable_optional_settings.get("paragraph_after_table", False) and text_after_table_check > 0:
                    checklist = self.text_after_table_checklist
                    current_author = "Параграф после таблицы"
                else:
                    checklist = self.text_checklist
                    current_author = "Абзац"

                # Сбор ошибок по всем run-ам
                current_errors = self.collect_paragraph_errors(p, checklist, document)

                if current_errors:
                    current_paragraph = p

            elif isinstance(block, Table):
                text_after_table_check = 2

                # Проверка названия таблицы (отложенно — на предыдущем параграфе)
                if self.enable_optional_settings.get("table_title", False) and isinstance(prev_paragraph, Paragraph):
                    caption_errors = self.collect_paragraph_errors(
                        prev_paragraph, self.table_name_checklist, document
                    )
                    if caption_errors:
                        self.create_comment(
                            document=document,
                            runs=prev_paragraph.runs,
                            element="Название таблицы",
                            comments=caption_errors
                        )
                        comment_count += 1
                        written_comments.append(caption_errors)
                    else:
                        # Можно добавить предупреждение "Нет названия таблицы"
                        pass

                # Проверка содержимого таблицы
                table_errors = []
                first_valid_paragraph = None

                for row_idx, row in enumerate(block.rows):
                    for col_idx, cell in enumerate(row.cells):
                        if not cell.text.strip():
                            continue
                        for cell_p in cell.paragraphs:
                            if not cell_p.runs or not cell_p.text.strip():
                                continue

                            if first_valid_paragraph is None:
                                first_valid_paragraph = cell_p

                            is_heading = (
                                (self.enable_optional_settings.get("table_headings_top", False) and row_idx == 0) or
                                (self.enable_optional_settings.get("table_headings_left", False) and col_idx == 0)
                            )
                            checklist = self.table_headings_checklist if is_heading else self.table_text_checklist

                            errs = self.collect_paragraph_errors(cell_p, checklist, document)
                            table_errors.extend(errs)
                            seen = set()
                            table_errors_filtered = []
                            
                            for err in table_errors:
                                err_key = tuple(str(x) for x in err)  # дедупликация
                                if err_key not in seen:
                                    seen.add(err_key)
                                    table_errors_filtered.append(err)


                if table_errors_filtered and first_valid_paragraph:
                    current_errors = table_errors_filtered
                    current_paragraph = first_valid_paragraph
                    current_author = "Таблица"

            # ── Сохранение предыдущего комментария (если есть) ───────────────────────
            if prev_paragraph is not None and pending_errors:
                self.create_comment(
                    document=document,
                    runs=prev_paragraph.runs,
                    element=pending_author,
                    comments=pending_errors
                )
                comment_count += 1
                written_comments.append(pending_errors)

            # Передаём текущие значения дальше
            prev_paragraph = current_paragraph
            pending_errors = current_errors
            pending_author = current_author

            # Счётчики
            if image_name_check > 0:
                image_name_check -= 1
            if text_after_table_check > 0:
                text_after_table_check -= 1

        # Последний отложенный комментарий
        if prev_paragraph is not None and pending_errors:
            self.create_comment(
                document=document,
                runs=prev_paragraph.runs,
                element=pending_author,
                comments=pending_errors
            )
            comment_count += 1
            written_comments.append(pending_errors)

        # Сохранение файла
        basename = os.path.splitext(os.path.basename(filename))[0]
        out_path = f"Results/{basename}_Проверенный.docx"
        count = 1
        while os.path.exists(out_path):
            out_path = f"Results/{basename}_Проверенный_{count}.docx"
            count += 1

        os.makedirs("Results", exist_ok=True)
        document.save(out_path)

        return written_comments  # или можно возвращать comment_count, save_path и т.д.

if __name__ == '__main__':
    parser = DocumentParser()
    '''parser.set_settings(default_text_checklist, default_heading1_checklist, default_heading2_checklist,
                        default_heading3_checklist, default_table_name_checklist, default_table_headings_checklist,
                        default_table_text_checklist, default_list_checklist, default_margins_checklist,
                        default_image_checklist, default_image_name_checklist)'''

    filename = "testDocs/testTables.docx"
    parser.parse_document(filename)
