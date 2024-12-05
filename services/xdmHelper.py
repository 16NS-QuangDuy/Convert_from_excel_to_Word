# encoding: utf-8
import re
from lxml import etree
XMLSPACE_NAMESPACE = '{http://www.tresos.de/_projects/DataModel2/08/root.xsd}'

xmlns_dict = dict()
xmlns_dict["a"] = "{http://www.tresos.de/_projects/DataModel2/08/attribute.xsd}"
xmlns_dict["v"] = "{http://www.tresos.de/_projects/DataModel2/06/schema.xsd}"
xmlns_dict["d"] = "{http://www.tresos.de/_projects/DataModel2/06/data.xsd}"


class XdmHelper:
    """XdmHelper"""
    auto_worker_name = "XDM"

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
        return [child for child in parent.getchildren() if child.tag == XMLSPACE_NAMESPACE + tag]

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
    def findall(parent, tag, ns="w"):
        return [child for child in parent.getchildren() if child.tag == (xmlns_dict[ns] + tag)]

    @staticmethod
    def finditer(parent, tag, ns="w"):
        return [child for child in parent.iter() if child.tag == (xmlns_dict[ns] + tag)]

    @staticmethod
    def finditer2(parent, child_tag="r", tag_list=["t", "tab", "sym", "br", "cr"], ns="w"):
        tag_list = [(xmlns_dict[ns] + tag) for tag in tag_list]
        child_child_list = []
        for child in XdmHelper.finditer(parent, child_tag, ns="w"):
            child_child_list.extend([child_child for child_child in child.iter() if child_child.tag in tag_list])
        return child_child_list

    @staticmethod
    def finditer3(parent, tag_list=["t", "tab", "sym", "br", "cr"], ns="w"):
        tag_list = [(xmlns_dict[ns] + tag) for tag in tag_list]
        child_child_list = []
        child_child_list.extend([child_child for child_child in parent.iter() if child_child.tag in tag_list])
        return child_child_list

    @staticmethod
    def append_elm(parent, xpath, new_elm):
        elm = parent.find(xpath, namespaces=parent.nsmap)
        if elm is not None:
            elm.append(new_elm)

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
            return elm.attrib[XMLSPACE_NAMESPACE + '%s' % attr]
        except:
            return None

    @staticmethod
    def set_elm_attr(elm, attr, new_val):
        """
        """
        try:
            elm.attrib[XMLSPACE_NAMESPACE + '%s' % attr] = new_val
            return elm
        except:
            return None

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
    def findall2(parent, tag_list, ns="w"):
        tag_list = [xmlns_dict[ns] + tag for tag in tag_list]
        return [child for child in parent.getchildren() if child.tag in tag_list]

    @staticmethod
    def create_elm(name, property_dict, text_contents=None):
        elm = OxmlElement(name)
        for k, v in property_dict.items():
            elm.set(qn(k), v)
        if text_contents is not None:
            elm.text = text_contents
        return elm

    @staticmethod
    def get_elm_xml(elm, debug=True):
        if debug:
            print(elm.xml)
        return elm.xml

    @staticmethod
    def get_elm_idx_list(document_elm, tag="p"):
        body_elm = XdmHelper.read_elm(document_elm, "./w:body", val=False)
        elm_idx_list = XdmHelper.findall2(body_elm, ["p", "tbl", "sdt"])
        elm_idx_list = [block_id for block_id, elm in enumerate(elm_idx_list) if XdmHelper.is_elm_tag(elm, tag)]
        return elm_idx_list

    @staticmethod
    def replace_text_by_regex(prop_file, replace_list):
        app_xml_file = prop_file
        fin = open(app_xml_file, "rt")
        content = fin.read()
        fin.close()
        # output file to write the result to
        fout = open(app_xml_file, "wt")
        # read replace the string and write to output file
        for replace_i in replace_list:
            match = re.search(r"%s" % replace_i[0], content)
            if match:
                text = match.group(0)
                if replace_i[1] != text and match.group(1).strip() != "":
                    content = content.replace(text, replace_i[1])
        fout.write(content)
        fout.close()

    @staticmethod
    def get_text_by_regex(prop_file, regex, group):
        app_xml_file = prop_file
        fin = open(app_xml_file, "rt")
        content = fin.read()
        fin.close()
        match = re.search(r"%s" % regex, content)
        if match:
            return match.group(group)
        return match
