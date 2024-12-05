# encoding: utf-8
import os
import re
import sys
from services.wordML import WordML
from services.elmHelper import ElmHelper
from docx.oxml.table import CT_Tbl
from docx.enum.text import WD_COLOR_INDEX
from docx.enum.text import WD_BREAK_TYPE
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
from docx.shared import RGBColor
from docx.oxml.xmlchemy import OxmlElement
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
WORD_NAMESPACE = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
PARA = WORD_NAMESPACE + 'p'
TEXT = WORD_NAMESPACE + 't'
BODY = WORD_NAMESPACE + 'body'
TBL = WORD_NAMESPACE + 'tbl'


class Finder:
    """Finder"""
    auto_worker_name = "Finder"

    @staticmethod
    def find_textbox_by_regex(block, regex_list=[r'R-car\s?\S+', r'<New>']):
        textbox_list = []
        if WordML.is_paragraph_block(block):
            for run in block.runs:
                for pict in run._element.findall(".//mc:AlternateContent/mc:Fallback/w:pict",
                                                          namespaces=run._element.nsmap):
                    for txbxContent in pict.findall(".//w:txbxContent",
                                                 namespaces=run._element.nsmap):
                        textbox_paragraphs = txbxContent.findall(".//w:p",
                                                                  namespaces=run._element.nsmap)
                        for p in textbox_paragraphs:
                            inbox_text = ''.join([n.text for n in p.findall(".//w:t", namespaces=p.nsmap)])
                            for regex in regex_list:
                                inner = re.search(regex, inbox_text, flags=re.I)
                                if inner:
                                    textbox_list.append(inbox_text)
        return textbox_list

    
    @staticmethod
    def find_textbox(block, regex_list=[r'R-car\s?\S+', r'<New>']):
        result_dict = dict()
        result_dict["t_lst"] = []
        result_dict["r_lst"] = []
        if WordML.is_paragraph_block(block):
            is_txbx_tag = False
            for run in block.runs:
                if run.text != "": continue
                result = ElmHelper.find_textbox(run._element)
                if result:
                    result_dict["r_lst"].append(run)
                for regex in regex_list:
                    inbox_text = "".join(result["t_lst"])
                    if re.search(regex, inbox_text, flags=re.I):
                        result_dict["t_lst"].extend(result["t_lst"])
                        is_txbx_tag = True
                        break
            if not is_txbx_tag:
                result_dict["t_lst"] = []
                result_dict["r_lst"] = []
        return result_dict

    @staticmethod
    def find_textbox_paragraphs_by_regex(block, regex_list=[r'R-car\s?\S+', r'<New>']):
        textbox_list = []
        if WordML.is_paragraph_block(block):
            for run in block.runs:
                for pict in run._element.findall(".//mc:AlternateContent/mc:Fallback/w:pict",
                                                          namespaces=run._element.nsmap):
                    for txbxContent in pict.findall(".//w:txbxContent",
                                                 namespaces=run._element.nsmap):
                        textbox_paragraphs = txbxContent.findall(".//w:p",
                                                                  namespaces=run._element.nsmap)
                        for p in textbox_paragraphs:
                            inbox_text = ''.join([n.text for n in p.findall(".//w:t", namespaces=p.nsmap)])
                            for regex in regex_list:
                                inner = re.search(regex, inbox_text, flags=re.I)
                                if inner:
                                    textbox_list.append(textbox_paragraphs)
        return textbox_list


    @staticmethod
    def find_all_elements_by_tags_from_file(docx_file, tag_list=['p']):
        package = docx.package.Package.open(docx_file)
        main_document_part = package.main_document_part
        assert isinstance(main_document_part, docx.parts.document.DocumentPart)
        document = main_document_part.document
        parent_elm = document.element.body
        elm_list = []
        for child in parent_elm.iterchildren():
            for tag in tag_list:
                if child.tag.endswith(tag):
                    elm_list.append(child)
                elms = child.findall("w:%s" % tag, namespaces = child.nsmap)
                elm_list.extend(elms)
        return elm_list

    @staticmethod
    def find_all_elements_by_tags_from_document(document, tag_list=['p']):
        parent_elm = document.element.body
        elm_list = []
        for child in parent_elm.iterchildren():
            for tag in tag_list:
                if child.tag.endswith(tag):
                    elm_list.append(child)
                elms = child.findall("w:%s" % tag, namespaces = child.nsmap)
                elm_list.extend(elms)
        return elm_list

    @staticmethod
    def find_all_elements_by_tags_from_parent(parent_elm, tag_list=['p']):
        elm_list = []
        for child in parent_elm.iterchildren():
            for tag in tag_list:
                if child.tag.endswith(tag):
                    elm_list.append(child)
                elms = child.findall("w:%s" % tag, namespaces = child.nsmap)
                elm_list.extend(elms)
        return elm_list

    @staticmethod
    def find_element_attrib(element, attrib):
        return element.attrib[WORD_NAMESPACE + attrib]

    @staticmethod
    def find_element_tag(element):
        return element.tag.replace(WORD_NAMESPACE, "")

    @staticmethod
    def is_paragraph_elm(block):
        if isinstance(block, CT_P):
            return True
        return False
    

