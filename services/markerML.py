# encoding: utf-8
# !python
# ******************************************************************************
# *
# * Revision History:
# * -----------------
# * Code Version    YYYY-MM-DD    Author        Description
# * 0.1             2020-11-24    HaiNDM        Create this file
# ******************************************************************************
import copy
from docx import Document
from docx.table import _Cell, Table
from docx.enum.text import WD_COLOR_INDEX
from docx.document import Document as _Document
from docx.text.paragraph import Paragraph
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.oxml.text.paragraph import CT_P
from docx.oxml.text.run import CT_R
from docx.oxml.table import CT_Tbl
from docx.shared import Inches
from docx.shared import RGBColor


class Marker:

    not_disclosed_color = [RGBColor(0, 176, 80), RGBColor(0, 177, 80)]

    @staticmethod
    def highlight_text_keep(paragraph):
        for run in paragraph.runs:
            run.font.highlight_color = WD_COLOR_INDEX.BRIGHT_GREEN

    @staticmethod
    def highlight_text_delete(paragraph):
        copied_block = copy.deepcopy(paragraph)
        for run in copied_block.runs:
            run.font.highlight_color = WD_COLOR_INDEX.GRAY_50
        return copied_block

    @staticmethod
    def highlight_text_error(paragraph):
        for run in paragraph.runs:
            run.font.highlight_color = WD_COLOR_INDEX.PINK

    @staticmethod
    def highlight_cell_delete(cell):
        from docx.oxml.ns import nsdecls
        from docx.oxml import parse_xml

        shading_elm_1 = parse_xml(r'<w:shd {} w:fill="dcdcdc"/>'.format(nsdecls('w')))
        cell._tc.get_or_add_tcPr().append(shading_elm_1)

    @staticmethod
    def highlight_cell_error(cell):
        from docx.oxml.ns import nsdecls
        from docx.oxml import parse_xml
        shading_elm_1 = parse_xml(r'<w:shd {} w:fill="ff1493"/>'.format(nsdecls('w')))
        cell._tc.get_or_add_tcPr().append(shading_elm_1)

    @staticmethod
    def highlight_cell_keep(cell):
        from docx.oxml.ns import nsdecls
        from docx.oxml import parse_xml

        shading_elm_1 = parse_xml(r'<w:shd {} w:fill="7cfc00"/>'.format(nsdecls('w')))
        cell._tc.get_or_add_tcPr().append(shading_elm_1)

    @staticmethod
    def highlight_run_delete(run):
        from docx.oxml.ns import nsdecls
        from docx.oxml import parse_xml

        shading_elm_1 = parse_xml(r'<w:shd {} w:val="pct15" w:color="auto" w:fill="cccccc"/>'.format(nsdecls('w')))
        rPr = run.find('.//w:rPr', namespaces=run.nsmap)
        if rPr is not None:
            rPr.append(shading_elm_1)

    @staticmethod
    def highlight_table_replacement(block):
        if isinstance(block, Table):
            copied_block = copy.deepcopy(block)
            for row_index in range(len(copied_block.rows)):
                row_cells = copied_block.rows[row_index].cells
                for cell in row_cells:
                    Marker.highlight_cell_keep(cell)
            return copied_block

    @staticmethod
    def highlight_table_delete(block):
        if isinstance(block, Table):
            copied_block = copy.deepcopy(block)
            for row_index in range(len(copied_block.rows)):
                row_cells = copied_block.rows[row_index].cells
                for cell in row_cells:
                    Marker.highlight_cell_delete(cell)
            return copied_block
