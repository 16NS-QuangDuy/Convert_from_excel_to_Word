# encoding: utf-8
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import copy
import re
from lxml import etree
XMLSPACE_NAMESPACE = '{http://www.w3.org/XML/1998/namespace}'
WORD_NAMESPACE = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
PARA = WORD_NAMESPACE + 'p'
TEXT = WORD_NAMESPACE + 't'
BODY = WORD_NAMESPACE + 'body'
TBL = WORD_NAMESPACE + 'tbl'

xmlns_dict = dict()
xmlns_dict["wpc"] = "{http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas}"
xmlns_dict["mc"] = "http://schemas.openxmlformats.org/markup-compatibility/2006}"
xmlns_dict["o"] = "{urn:schemas-microsoft-com:office:office}"
xmlns_dict["r"] = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
xmlns_dict["m"] = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
xmlns_dict["v"] = '{urn:schemas-microsoft-com:vml}'
xmlns_dict["wp14"] = "{http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing}"
xmlns_dict["wp"] = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
xmlns_dict["w10"] = "{urn:schemas-microsoft-com:office:word}"
xmlns_dict["w"] = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
xmlns_dict["w14"] = "{http://schemas.microsoft.com/office/word/2010/wordml}"
xmlns_dict["w15"] = "{http://schemas.microsoft.com/office/word/2012/wordml}"
xmlns_dict["wpg"] = "{http://schemas.microsoft.com/office/word/2010/wordprocessingGroup}"
xmlns_dict["wpi"] = "{http://schemas.microsoft.com/office/word/2010/wordprocessingInk}"
xmlns_dict["wne"] = "{http://schemas.microsoft.com/office/word/2006/wordml}"
xmlns_dict["wps"] = "{http://schemas.microsoft.com/office/word/2010/wordprocessingShape}"
xmlns_dict["pic"] = "{http://schemas.microsoft.com/office/word/2006/picture}"
XMLSPACE_NAMESPACE = '{http://www.w3.org/XML/1998/namespace}'

non_empty_elm_list = [
    ("drawing", "w"),
    ("inline", "wp"),
    ("anchor", "wp"),
    ("pic", "pic"),
    ("AlternateContent", "mc"),
    ("shape", "v"),
    ("shapetype", "v"),
    ("formulas", "v"),
    ("path", "v"),
    ("object", "w"),
    ("textbox", "v"),
    ("tbl", "w"),
    ("pict", "w"),
    ("line", "v"),
    ("group", "wpg"),
    ("tbl", "w"),
    ("hyperlink", "w"),
    ]


class ElmHelper:
    """Finder"""
    auto_worker_name = "ELM"

    @staticmethod
    def remove_elm(parent, xpath="", first=True, del_col=[]):
        """
        Remove child from parent element
        :param parent: tc
        :param xpath: ".//w:tcPr/w:textDirection"
        :return:
        """
        if xpath == "":
            parent.getparent().remove(parent)
        if first:
            elm = parent.find(xpath, namespaces=parent.nsmap)
            if elm is not None:
                elm.getparent().remove(elm)
        else:
            elm_list = parent.findall(xpath, namespaces=parent.nsmap)
            for idx, elm in enumerate(elm_list):
                if del_col == [] or idx in del_col:
                    elm.getparent().remove(elm)

    @staticmethod
    def replace_elm(parent, xpath, new_elm, first=True):
        """
        Remove child from parent element
        :param parent: tc
        :param xpath: ".//w:tcPr/w:textDirection"
        :return:
        """
        if first:
            elm = parent.find(xpath, namespaces=parent.nsmap)
            if elm is not None:
                elm.getparent().replace(elm, new_elm)
        else:
            elm_list = parent.findall(xpath, namespaces=parent.nsmap)
            for idx, elm in enumerate(elm_list):
                elm.getparent().replace(elm, new_elm)
    
    @staticmethod
    def findall(parent, tag):
        return [child for child in parent.getchildren() if child.tag == WORD_NAMESPACE + tag]

    @staticmethod
    def find_by_idx(parent, tag, idx):
        k = -1
        for child in parent.getchildren():
            if child.tag == WORD_NAMESPACE + tag:
                k += 1
                if k == idx:
                    return child
        return None

    @staticmethod
    def get_tc_text(tc):
        p_list = []
        for p in ElmHelper.findall(tc, 'p'):
            text = ""
            for r in ElmHelper.findall(p, 'r'):
                for t in ElmHelper.findall(r, 't'):
                    if t.text is not None:
                        text += t.text
            p_list.append(text)
        return "\n".join(p_list)

    @staticmethod
    def get_tr_text(tr):
        text_list = []
        for tc in ElmHelper.findall(tr, 'tc'):
            text_list.append(ElmHelper.get_tc_sym(tc))
        return text_list

    @staticmethod
    def get_row_cells_text(block, i):
        tr = ElmHelper.find_by_idx(block._tbl, "tr", i)
        if tr is not None:
            text_list = []
            for tc in ElmHelper.findall(tr, 'tc'):
                text_list.append(ElmHelper.get_tc_sym(tc))
        return text_list
        
    @staticmethod
    def get_tc_sym(tc):
        p_list = []
        for p in ElmHelper.findall(tc, 'p'):
            text = ""
            for r in ElmHelper.findall(p, 'r'):
                for child in r.getchildren():
                    if child.tag == (WORD_NAMESPACE + "sym"):
                        char = child.attrib[WORD_NAMESPACE + "char"]
                        symbol = chr(int(char, 16))
                        text += symbol
                    elif child.tag == (WORD_NAMESPACE + "t"):
                        if child.text is not None:
                            text += child.text
            p_list.append(text)
        return "\n".join(p_list)

    @staticmethod
    def findall(parent, tag, ns="w"):
        return [child for child in parent.getchildren() if child.tag == (xmlns_dict[ns] + tag)]

    @staticmethod
    def finditer(parent, tag, ns="w"):
        return [child for child in parent.iter() if child.tag == (xmlns_dict[ns] + tag)]

    @staticmethod
    def finditer2(parent, child_tag="r", tag_list=["t", "tab", "sym", "br", "cr"], ns="w"):
        tag_list = [(xmlns_dict[ns] + tag) for tag in tag_list]
        child_child_list = []
        for child in ElmHelper.finditer(parent, child_tag, ns="w"):
            child_child_list.extend([child_child for child_child in child.iter() if child_child.tag in tag_list])
        return child_child_list

    @staticmethod
    def finditer3(parent, tag_list=["t", "tab", "sym", "br", "cr"], ns="w"):
        tag_list = [(xmlns_dict[ns] + tag) for tag in tag_list]
        child_child_list = []
        child_child_list.extend([child_child for child_child in parent.iter() if child_child.tag in tag_list])
        return child_child_list
    
    @staticmethod
    def find_textbox(parent):
        result_dict = dict()
        result_dict["txbx_lst"] = [child for child in parent.iter() if child.tag == (xmlns_dict["w"] + "txbxContent")]
        result_dict["t_lst"] = []
        for txbx in result_dict["txbx_lst"]:
            # inbox_text = ElmHelper.get_elm_text(txbx.iter())
            t_list = [child for child in txbx.iter() if child.tag == (xmlns_dict["w"] + "t")]
            inbox_text = ''.join(filter(None,[t.text for t in t_list]))
            result_dict["t_lst"].append(inbox_text)
        result_dict["t_lst"] = list(dict.fromkeys(result_dict["t_lst"]))
        return result_dict

    @staticmethod
    def find_textbox_figure(parent):
        result_dict = dict()
        result_dict["txbx_lst"] = [child for child in parent.iter() if child.tag == (xmlns_dict["w"] + "txbxContent")]
        result_dict["p_lst"] = []
        result_dict["is_txbx_figure"] = False
        if len(result_dict["txbx_lst"]) >= 1:
            for txbx in result_dict["txbx_lst"]:
                t_list = [child for child in txbx.iter() if child.tag == (xmlns_dict["w"] + "t")]
                inbox_text = ''.join(filter(None, [t.text for t in t_list]))
                if "Figure" in inbox_text:
                    p_list = [child for child in txbx.iter() if child.tag == (xmlns_dict["w"] + "p")]
                    result_dict["p_lst"] = p_list
                    result_dict["is_txbx_figure"] = True
                    break
        result_dict["p_lst"] = list(dict.fromkeys(result_dict["p_lst"]))
        return result_dict

    @staticmethod
    def append_elm(parent, xpath, new_elm):
        elm = parent.find(xpath, namespaces=parent.nsmap)
        if elm is not None:
            elm.append(new_elm)

    @staticmethod
    def merge_tc_v(tc0, tc1, new_text):
        ElmHelper.append_elm(tc0, ".//w:tcPr", parse_xml(r'<w:gridSpan {} w:val="1"/>'.format(nsdecls('w'))))
        ElmHelper.append_elm(tc0, ".//w:tcPr", parse_xml(r'<w:vMerge {} w:val="restart"/>'.format(nsdecls('w'))))
        ElmHelper.append_elm(tc1, ".//w:tcPr", parse_xml(r'<w:gridSpan {} w:val="1"/>'.format(nsdecls('w'))))
        ElmHelper.append_elm(tc1, ".//w:tcPr", parse_xml(r'<w:vMerge {}/>'.format(nsdecls('w'))))
        ElmHelper.remove_elm(tc0, ".//w:p/w:r/w:t", first=False)
        t_elm = OxmlElement('w:t')
        t_elm.text = new_text
        ElmHelper.append_elm(tc0, ".//w:p/w:r", t_elm)

    @staticmethod
    def merge_tc_h(tc0, tc1, new_text):
        ElmHelper.append_elm(tc0, ".//w:tcPr", parse_xml(r'<w:gridSpan {} w:val="2"/>'.format(nsdecls('w'))))
        ElmHelper.remove_elm(tc0, ".//w:p/w:r/w:t", first=False)
        t_elm = OxmlElement('w:t')
        t_elm.text = new_text
        w0 = ElmHelper.read_elm_attr(tc0, r'<w:tcW w:w=\"(.*)\" .*\/>', 1)
        w1 = ElmHelper.read_elm_attr(tc1, r'<w:tcW w:w=\"(.*)\" .*\/>', 1)
        w = "%s" % (int(w0) + int(w1))
        ElmHelper.replace_elm(tc0, ".//w:tcPr/w:tcW", ElmHelper.create_elm("w:tcW", {"w:w": w, "w:type": "dxa"}))
        ElmHelper.append_elm(tc0, ".//w:p/w:r", t_elm)

    @staticmethod
    def format_tc(tc1):
        ElmHelper.remove_elm(tc1, ".//w:tcPr/w:vAlign")
        ElmHelper.remove_elm(tc1, ".//w:tcPr/w:shd")
        ElmHelper.append_elm(tc1, ".//w:tcPr", parse_xml(r'<w:vAlign {} w:val="bottom"/>'.format(nsdecls('w'))))
        ElmHelper.append_elm(tc1, ".//w:tcPr", 
                             parse_xml(r'<w:shd {} w:val="clear" w:color="auto" w:fill="auto"/>'.format(nsdecls('w'))))

    @staticmethod
    def format_tc_border_nil(tc1, pos='top'):
        if 'top' == pos:
            ElmHelper.replace_elm(tc1, ".//w:tcPr/w:tcBorders",  parse_xml(r'<w:top {} w:val="nil"/>'.format(nsdecls('w'))))
        elif 'bottom' == pos:
            ElmHelper.replace_elm(tc1, ".//w:tcPr/w:tcBorders",
                                  parse_xml(r'<w:bottom {} w:val="nil"/>'.format(nsdecls('w'))))

    @staticmethod
    def find_tc_by_index(tr, i):
        """
        convert index for merge cell case.
        :param tr:
        :param i: index from user view
        :return: index from xml data
        """
        real_idx = -1
        for idx, tc in enumerate(ElmHelper.findall(tr, 'tc')):
            gridSpan = [child for child in tc.iter() if child.tag == (xmlns_dict["w"] + "gridSpan")]
            if gridSpan:
                elm_val = gridSpan[0].attrib[WORD_NAMESPACE + 'val']
                real_idx += int(elm_val)
            else:
                real_idx += 1
                
            if i == real_idx:
                return tc
        return None

    @staticmethod
    def read_elm(parent, xpath, first=True, val=True):
        """
        Remove child from parent element
        :param parent: tc
        :param xpath: ".//w:tcPr/w:textDirection"
        :return:
        """
        if first:
            elm = parent.find(xpath, namespaces=parent.nsmap)
            if val:
                return elm.val if elm is not None else None
            return elm
        else:
            elm_list = parent.findall(xpath, namespaces=parent.nsmap)
            if val:
                return [elm.val if elm is not None else None for elm in elm_list]
            return elm_list
        
    @staticmethod
    def read_elm_attr(elm, regex, i):
        """
        Remove child from parent element
        :param parent: tc
        :param xpath: ".//w:tcPr/w:textDirection"
        :return:
        """
        match = re.search(regex, elm.xml)
        if match:
            return match.group(i)
        return None

    @staticmethod
    def read_elm_attr2(elm, attr):
        """
        """
        try:
            return elm.attrib[WORD_NAMESPACE + '%s' % attr]
        except:
            return None

    @staticmethod
    def set_elm_attr(elm, attr, new_val):
        """
        """
        try:
            elm.attrib[WORD_NAMESPACE + '%s' % attr] = new_val
            return elm
        except:
            return None

    @staticmethod
    def set_row_height(row, val):
        trHeight = OxmlElement('w:trHeight')
        trHeight.set(qn('w:val'), '%s' % val)
        trHeight.set(qn('w:hRule'), 'exact')
        row._tr.get_or_add_trPr().append(trHeight)

    @staticmethod
    def get_text_run_list(run_list):
        text = ""
        for run in run_list:
            text += ElmHelper.get_r_text(run)
        return text

    @staticmethod
    def get_r_text(r):
        text = ""
        for t in ElmHelper.findall(r, 't'):
            if t is not None and t.text is not None:
                text = text + t.text
        for t in ElmHelper.findall(r, 't', 'm'):
            if t is not None and t.text is not None:
                text = text + t.text
        return text

    @staticmethod
    def get_r_text2(r):
        text = ""
        for t in ElmHelper.findall(r, 't'):
            if t is not None and t.text is not None:
                text = text + t.text
        for t in ElmHelper.findall(r, 'delText'):
            if t is not None and t.text is not None:
                text = text + t.text
        return text

    @staticmethod
    def update_elm_t(t, new_t):
        if t is not None:
            t.text = new_t

    @staticmethod
    def get_t_text(t):
        text = ""
        if t is not None and t.text is not None:
            text = t.text
        return text

    @staticmethod
    def finditer(parent, tag, ns="w"):
        return [child for child in parent.iter() if child.tag == (xmlns_dict[ns] + tag)]

    @staticmethod
    def standardize_spaces(paragraph, remove_tab=False):
        """
        standardize spaces
        :param paragraph: a paragraph element
        :return: 
        """
        if remove_tab:
            # for tab in  ElmHelper.finditer2(paragraph, 'r', ['tab']):
            for run in ElmHelper.findall(paragraph, 'r'):
                for tab in ElmHelper.findall(run, 'tab'):
                    run.replace(tab, parse_xml(r'<w:t {} xml:space="preserve"> </w:t>'.format(nsdecls('w'))))
        t_list = ElmHelper.finditer2(paragraph)
        is_tag_t_list = [True if child.tag == qn('w:t') else False for child in t_list]
        is_greater_2 = (len(t_list) >= 2)
        for i in range(len(t_list)):
            if is_tag_t_list[i] and ElmHelper.get_t_text(t_list[i]) != "":
                t_list[i].text = re.sub(r'\s+', ' ', t_list[i].text)
        if is_greater_2:
            for i in range(len(t_list)-1):
                if is_tag_t_list[i] and is_tag_t_list[i+1] and ElmHelper.get_t_text(t_list[i]).endswith(" ") \
                        and ElmHelper.get_t_text(t_list[i+1]).startswith(" "):
                    t_list[i].text = t_list[i].text.rstrip()
        for i in range(len(t_list)):
            if is_tag_t_list[i] and ElmHelper.get_t_text(t_list[i]).startswith(" "):
                t_list[i].text = t_list[i].text.lstrip()
            if not is_tag_t_list[i] or t_list[i].text != "":
                break
        for i in range(len(t_list)-1, -1, -1):
            if is_tag_t_list[i] and ElmHelper.get_t_text(t_list[i]).endswith(" "):
                t_list[i].text = t_list[i].text.rstrip()
            if not is_tag_t_list[i] or t_list[i].text != "":
                break
        return paragraph

    @staticmethod
    def standardize_tbl_grid(tbl):
        elm_list = ElmHelper.read_elm(tbl, ".//w:tblGrid/w:gridCol", first=False, val=False)
        tr_1 = ElmHelper.find_by_idx(tbl, "tr", 1)
        cell_text_list = ElmHelper.get_tr_text(tr_1)
        if len(elm_list) > len(cell_text_list):
            elm_w_list = [ElmHelper.read_elm_attr(elm, r"<w:gridCol\s.*\sw:w=\"(.*)\"\/>", 1) for elm in elm_list]
            elm_w_list = [int(w) for w in elm_w_list]
            elm_del = [i for i, w in enumerate(elm_w_list) if w < 30]
            if len(elm_del) > 0:
                ElmHelper.remove_elm(tbl, ".//w:tblGrid/w:gridCol", del_col=elm_del)
                elm_list = ElmHelper.read_elm(tbl, ".//w:tblGrid/w:gridCol", first=False, val=False)

    @staticmethod
    def format_tc_center(tc):
        ElmHelper.append_elm(tc, ".//w:tcPr", parse_xml(r'<w:vAlign {} w:val="center"/>'.format(nsdecls('w'))))

    @staticmethod
    def replace_elm_text(paragraph, old_text, new_text):
        if old_text == new_text: return
        t_list = ElmHelper.finditer(paragraph, 't')
        p_text = ""
        end = len(t_list) - 1
        for i in range(len(t_list)):
            if t_list[i] is not None:
                p_text += ElmHelper.get_t_text(t_list[i])
                if old_text in p_text:
                    end = i
                    break
        start = 0
        p_text = ""
        for i in range(len(t_list)-1, -1, -1):
            if t_list[i] is not None:
                p_text = ElmHelper.get_t_text(t_list[i]) + p_text
                if old_text in p_text:
                    start = i
                    break
        change_t_list = t_list[start: end + 1]
        p_text = "".join([ElmHelper.get_t_text(t) for t in change_t_list if t is not None])
        if p_text == old_text:
            if len(change_t_list) > 0:
                if change_t_list[0] is not None:
                    change_t_list[0].text = new_text
            for t in change_t_list[1:]:
                if t is not None:
                    t.text = ""

        else:
            p_text = p_text.replace(old_text, new_text)
            if len(change_t_list) >= 1:
                if change_t_list[0] is not None:
                    change_t_list[0].text = p_text
                    for t in change_t_list[1:]:
                        if t is not None:
                            t.text = ""
        if new_text.startswith(" ") or new_text.endswith(" "):
            if len(change_t_list) >= 1:
                if change_t_list[0] is not None:
                    t = change_t_list[0]
                    t.attrib[XMLSPACE_NAMESPACE + 'space'] = "preserve"

    @staticmethod
    def set_elm_text(elm, new_text):
        t_list = ElmHelper.finditer(elm, 't')
        change_t_list = t_list
        if len(change_t_list) > 0:
            if change_t_list[0] is not None:
                change_t_list[0].text = new_text
        for t in change_t_list[1:]:
            if t is not None:
                t.text = ""

    @staticmethod
    def is_elm_tag(elm, tag, ns="w"):
        return elm is not None and elm.tag == (xmlns_dict[ns] + tag)

    @staticmethod
    def get_txbox_text(txbx):
        t_list = [child for child in txbx.iter() if child.tag == (xmlns_dict["w"] + "t")]
        inbox_text = ''.join([t.text for t in t_list])
        return inbox_text

    @staticmethod
    def get_tag_name(ns, name):
        return xmlns_dict[ns] + name

    @staticmethod
    def get_tag_name1(ns, name):
        return name.replace(xmlns_dict[ns], "")

    @staticmethod
    def get_elm_text(parent):
        """
        Modify from 'https://github.com/python-openxml/python-docx/blob/master/docx/oxml/text/run.py'
        Line 84
        :param parent: run
        :return:
        """
        text = ''
        for child in parent.getchildren():
            if child.tag in (qn('w:t'), qn('m:t')):
                text += ElmHelper.get_t_text(child)
            elif child.tag == qn('w:tab'):
                text += '\t'
            elif child.tag in (qn('w:br'), qn('w:cr')):
                text += '\n'
            elif child.tag in (qn('w:sym')):
                char = child.attrib[qn('w:char')]
                text += chr(int(char, 16))
            elif child.tag in (qn('w:noBreakHyphen')):
                text += '-'
        return text

    @staticmethod
    def get_elm_text2(parent, space=' '):
        """
        Modify from 'https://github.com/python-openxml/python-docx/blob/master/docx/oxml/text/run.py'
        Line 84
        :param parent: run
        :param space: replace space
        :return:
        """
        text = ''
        for child in parent.getchildren():
            if child.tag in (qn('w:t'), qn('m:t'), qn('w:delText')):
                text_i = ElmHelper.get_t_text(child)
                if space == '{SPACE}':
                    if text_i.count(" ") == len(text_i):
                        text_i = text_i.replace(" ", "{SPACE}")
                text += text_i
            elif child.tag == qn('w:tab'):
                text += '{TAB}'
            elif child.tag in (qn('w:cr')):
                text += '{ENTER}'
            elif child.tag in (qn('w:br')):
                br = child
                is_page = ElmHelper.read_elm_attr2(br, "type")
                if is_page == "page":
                    text += '{PAGE_BREAK}'
                else:
                    text += '{ENTER}'
            elif child.tag in (qn('w:sym')):
                char = child.attrib[qn('w:char')]
                text += chr(int(char, 16))
            elif child.tag in (qn('w:noBreakHyphen')):
                text += '-'
        return text

    @staticmethod
    def get_elm_text3(parent, space=' '):
        """
        get elm text for inserted paragraph
        :param parent: run
        :param space: replace space
        :return:
        """
        text = ''
        for child in parent.getchildren():
            if child.tag in (qn('w:t'), qn('m:t'), qn('w:ins')):
                text_i = ElmHelper.get_t_text(child)
                if space == '{SPACE}':
                    if text_i.count(" ") == len(text_i):
                        text_i = text_i.replace(" ", "{SPACE}")
                text += text_i
            elif child.tag == qn('w:tab'):
                text += '{TAB}'
            elif child.tag in (qn('w:cr')):
                text += '{ENTER}'
            elif child.tag in (qn('w:br')):
                br = child
                is_page = ElmHelper.read_elm_attr2(br, "type")
                if is_page == "page":
                    text += '{PAGE_BREAK}'
                else:
                    text += '{ENTER}'
            elif child.tag in (qn('w:sym')):
                char = child.attrib[qn('w:char')]
                text += chr(int(char, 16))
            elif child.tag in (qn('w:noBreakHyphen')):
                text += '-'
        return text

    @staticmethod
    def get_p_elm_text(p):
        run_list = ElmHelper.findall(p, 'r')
        run_list1 = ElmHelper.finditer2(p, 'hyperlink', ['r'])
        if len(run_list1) > 0:
            run_list.extend(run_list1)
        text = ''
        for r in run_list:
            text += ElmHelper.get_elm_text(r)
        # run_list1 = ElmHelper.finditer(p, 'r', 'm')
        # for r in run_list1:
        #     text += ElmHelper.get_elm_text(r)
        return text

    @staticmethod
    def find_text(parent):
        text = ''
        result_dict = dict()
        for child in parent.iter():
           if child.tag == qn('w:t'):
               text += ElmHelper.get_t_text(child)
           elif child.tag == qn('w:tab'):
               text += '\t'
           elif child.tag in (qn('w:br'), qn('w:cr')):
               text += '\n'
           elif child.tag in (qn('w:sym')):
               char = child.attrib[qn('w:char')]
               text += chr(int(char, 16))
        return text

    @staticmethod
    def findall2(parent, tag_list, ns="w"):
        tag_list = [xmlns_dict[ns] + tag for tag in tag_list]
        return [child for child in parent.getchildren() if child.tag in tag_list]

    @staticmethod
    def is_empty_p_elm(p_elm):
        is_blank_text = ElmHelper.get_p_elm_text(p_elm) == ""
        if not is_blank_text:
            return False
        for (tag, ns) in non_empty_elm_list:
            if len(ElmHelper.finditer(p_elm, tag, ns)) > 0:
                return False
        return True

    @staticmethod
    def is_empty_elm(elm):
        is_blank_text = ElmHelper.get_p_elm_text(elm).strip() in ["", "\t", "\n"]
        if not is_blank_text:
            return False
        for (tag, ns) in non_empty_elm_list:
            if len(ElmHelper.finditer(elm, tag, ns)) > 0:
                return False
        return True

    @staticmethod
    def is_empty_line(tr):
        cells_text = ElmHelper.get_tr_text(tr)
        is_empty_line = cells_text.count("") == len(cells_text)
        tc_list = ElmHelper.findall(tr, 'tc')
        empty_line = all([ElmHelper.is_empty_tc(tc) for tc in tc_list])
        if not empty_line:
            return False
        tc_border_list = []
        for tc in ElmHelper.finditer(tr, "tc"):
            for border in ["left", "right", "top", "bottom"]:
                border_elm = ElmHelper.read_elm(tc, ".//w:tcPr/w:tcBorders/w:%s" % border, val=False)
                border_val = ElmHelper.read_elm_attr2(border_elm, "val")
                if border_val is not None and border_val in ["dotted"]:
                    tc_border_list.append(border_val)
        if len(tc_border_list) > 0:
            is_empty_line = False
        return is_empty_line

    @staticmethod
    def create_elm(name, property_dict, text_contents=None):
        elm = OxmlElement(name)
        for k, v in property_dict.items():
            elm.set(qn(k), v)
        if text_contents is not None:
            elm.text = text_contents
        return elm

    @staticmethod
    def delete_paragraph(p_elm):
        try:
            parent = p_elm.getparent()
            # cell must have at least 1 paragraph
            if parent.tag == WORD_NAMESPACE + 'tc':
                p_list = [child for child in parent.getchildren() if child.tag == WORD_NAMESPACE + 'p']
                if len(p_list) > 1:
                    p = p_elm
                    p.getparent().remove(p)
                else:
                    ElmHelper.set_elm_text(p_elm, "")
            else:
                if ElmHelper.finditer3(p_elm, tag_list=["headerReference", "footerReference"], ns="w"):
                    pass
                else:
                    p = p_elm
                    p.getparent().remove(p)
        except:
            print("ERROR: ElmHelper.delete_paragraph...")
            pass

    @staticmethod
    def remove_break_page(p_elm):
        run_list = ElmHelper.findall(p_elm, 'r')
        for r in run_list:
            br = r.find('.//w:br', namespaces=r.nsmap)
            if br is not None:
                br.getparent().remove(br)

    @staticmethod
    def add_run_page_break(p_elm):
        pbreak_e = r'<w:br {} w:type="page"/>'.format(nsdecls('w'))
        p_elm.append(parse_xml(pbreak_e))

    @staticmethod
    def get_para_id(elm, text=None):
        regex = r".*%s=\"([A-F0-9]+)\".*" % "w14:paraId"
        para_id = ElmHelper.read_elm_attr(elm, regex, 1)
        if text is not None and text == ElmHelper.get_p_elm_text(elm):
            print(ElmHelper.get_p_elm_text(elm), para_id)
            print(elm.xml)
        return para_id

    @staticmethod
    def get_elm_xml(elm, debug=True):
        if debug:
            print(elm.xml)
        return elm.xml

    @staticmethod
    def get_elm_idx_list(document_elm, tag="p"):
        body_elm = ElmHelper.read_elm(document_elm, "./w:body", val=False)
        elm_idx_list = ElmHelper.findall2(body_elm, ["p", "tbl", "sdt"])
        elm_idx_list = [block_id for block_id, elm in enumerate(elm_idx_list) if ElmHelper.is_elm_tag(elm, tag)]
        return elm_idx_list

    @staticmethod
    def remove_break_page_before(p_elm):
        pPr = p_elm.find('.//w:pPr', namespaces=p_elm.nsmap)
        brbf = pPr.find('.//w:pageBreakBefore', namespaces=pPr.nsmap)
        if brbf is not None:
            pPr.remove(brbf)

    @staticmethod
    def is_empty_tc(elm):
        cond_1 = [ElmHelper.is_empty_elm(p) for p in ElmHelper.findall(elm, 'p')]
        cond_2 = [ElmHelper.is_empty_elm_properties(prop) for prop in ElmHelper.findall(elm, 'tcPr')]
        return all(cond_1) and all(cond_2)

    @staticmethod
    def is_empty_support_tc(elm):
        cond_1 = [ElmHelper.is_empty_elm(p) for p in ElmHelper.findall(elm, 'p')]
        return all(cond_1)

    @staticmethod
    def is_empty_r(elm):
        is_blank_text = ElmHelper.get_r_text(elm) in [""]
        if not is_blank_text:
            return False
        for (tag, ns) in non_empty_elm_list:
            if len(ElmHelper.finditer(elm, tag, ns)) > 0:
                return False
        return True

    @staticmethod
    def has_texbox_figure_elm(p_elm):
        for (tag, ns) in non_empty_elm_list:
            if len(ElmHelper.finditer(p_elm, tag, ns)) > 0:
                return True
        return False

    @staticmethod
    def has_tbl_elm(elm):
        tbl_elm_list = [
            ("tbl", "w"),
        ]
        for (tag, ns) in tbl_elm_list:
            if len(ElmHelper.finditer(elm, tag, ns)) > 0:
                return True
        return False

    @staticmethod
    def is_empty_elm_properties(elm):
        non_empty_prop_list = [("shd", "w")]
        for (tag, ns) in non_empty_prop_list:
            if len(ElmHelper.finditer(elm, tag, ns)) > 0:
                return False
        return True

    @staticmethod
    def is_page_break(elm):
        br_elm = ElmHelper.read_elm(elm, './/w:br', first=True, val=False)
        if br_elm is not None:
            type = ElmHelper.read_elm_attr2(br_elm, 'type')
            if type == 'page':
                return True
        return False

    @staticmethod
    def is_empty_line_by_content(tr):
        cells_text = ElmHelper.get_tr_text(tr)
        is_empty_line = cells_text.count("") == len(cells_text)
        return is_empty_line

    @staticmethod
    def is_vertical_merge_line(tr):
        tc_list = ElmHelper.findall(tr, "tc")
        tc_attr_list = [ElmHelper.read_elm_attr(tc, "<w:vMerge/>", 0) for tc in tc_list]
        vertical_merge_list = [True if attr == "<w:vMerge/>" else False for attr in tc_attr_list]
        return all(vertical_merge_list)

    @staticmethod
    def find_pos_col0_list(tr0, col_list):
        """
        find the actual column index by cell width
        it is needed in case of merge cell so refer the header cell position
        """
        tc0_list = ElmHelper.findall(tr0, 'tc')
        tc0_w_list = [ElmHelper.read_elm_attr(tc, r"<w:tcW\sw:w=\"(.*)\" w:type=\".*\"/>", 1) for tc in
                      tc0_list]
        tc0_w_list = [int(w) for w in tc0_w_list]
        pos_col0_list = [sum(tc0_w_list[:col + 1]) for col in col_list]
        return pos_col0_list

    @staticmethod
    def find_tc_list_by_pos_col0_list(tr, del_pos_col0):
        """
        convert index for merge cell case.
        :param tr:
        :param i: index from user view
        :return: index from xml data
        """
        total_w = 0
        total_w_list = []
        del_tc_list = []
        for idx, tc in enumerate(ElmHelper.findall(tr, 'tc')):
            tc_w = ElmHelper.read_elm_attr(tc, r"<w:tcW\sw:w=\"(.*)\" w:type=\".*\"/>", 1)
            total_w += int(tc_w)
            total_w_list.append(total_w)
            for j in del_pos_col0:
                if total_w >= j:
                    del_tc_list.append(tc)
        return del_tc_list

    @staticmethod
    def remove_br_from_p_elm(p_elm, idx_list=[0]):
        cut_list = []
        for r in ElmHelper.findall(p_elm, "r"):
            for child in r.getchildren():
                if child.tag in (qn('w:br'), qn('w:cr')):
                    cut_list.append(child)
        for idx in idx_list:
            if idx < len(cut_list):
                child = cut_list[idx]
                if child.getparent() is not None:
                    child.getparent().remove(child)

    @staticmethod
    def get_info_from_node(node, parent, level):
        """
        Args:
            node: node elements tree - child of "parent"
            parent: the first tag
            level: linked to ElmHelper.etree_to_list()
        Returns:
            A list of [lvl, title_num, title_text]

        EXAMPLE: ElmHelper.etree_to_list(t)
            -----
            title_ele_list = [child for child in t.getchildren() if child.tag == 'Title']
            for title_ele0 in title_ele_list:
                # lvl0
                bm_list.append(ElmHelper.get_info_from_node(title_ele0, t, level=0))
            -----
        """
        lvl, title, page, tp, left, top = int(), str(), int(), str(), float(), float()
        title_num, title_text = int(), str()
        if node.getparent() == parent:
            lvl = level
        if node.text:
            title = node.text.strip()
            match_title = re.match(r'^([\d.U]+|Appendix\s[AB.]+)\s+(.+)$', title)
            if match_title:
                title_num = match_title.group(1)
                title_text = match_title.group(2)
            else:
                title_num = 'Covers_'
                title_text = title
        if node.attrib:
            if 'Page' in node.attrib:
                # title_ele.attrib['Page']
                # tp, left, top
                match = re.match(r'^(\S+)\s(\S+)\s(\S+)\s(\S+)', node.attrib['Page'])
                if match:
                    page = int(match.group(1))
                    tp = '/' + match.group(2)
                    left = float(match.group(3))
                    top = float(match.group(4))
        return [lvl, title_num, title_text]

    @staticmethod
    def etree_to_list(t):
        """
        Args:
            t ():

        Returns:

        """
        # currently support lvl from 0 to 5 (6 levels)
        bm_list = []
        title_ele_list = [child for child in t.getchildren() if child.tag == 'Title']
        for title_ele0 in title_ele_list:
            # lvl0
            bm_list.append(ElmHelper.get_info_from_node(title_ele0, t, level=0))
            # lvl1
            children0 = list(title_ele0)
            if children0:
                for title_ele1 in children0:
                    bm_list.append(ElmHelper.get_info_from_node(title_ele1, title_ele0, level=1))
                    # lvl2
                    children1 = list(title_ele1)
                    if children1:
                        for title_ele2 in children1:
                            bm_list.append(ElmHelper.get_info_from_node(title_ele2, title_ele1, level=2))
                            # lvl3
                            children2 = list(title_ele2)
                            if children2:
                                for title_ele3 in children2:
                                    bm_list.append(ElmHelper.get_info_from_node(title_ele3, title_ele2, level=3))
                                    # lvl4
                                    children3 = list(title_ele3)
                                    if children3:
                                        for title_ele4 in children3:
                                            bm_list.append(
                                                ElmHelper.get_info_from_node(title_ele4, title_ele3, level=4))
                                            # lvl5
                                            children4 = list(title_ele4)
                                            if children4:
                                                for title_ele5 in children4:
                                                    bm_list.append(
                                                        ElmHelper.get_info_from_node(title_ele5, title_ele4,
                                                                                          level=5))
                                                    # lvl6
                                                    children5 = list(title_ele5)
                                                    if children5:
                                                        for title_ele6 in children5:
                                                            bm_list.append(
                                                                ElmHelper.get_info_from_node(title_ele6,
                                                                                                  title_ele5,
                                                                                                  level=6))
        return bm_list
