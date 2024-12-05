# encoding: utf-8
import xml.etree.ElementTree as ET


class XmlML:
    """handler for doing YAML file creation and reading"""
    auto_worker_name = __name__.split(".")[-1]
    auto_worker_name = auto_worker_name[0].upper() + auto_worker_name[1:]

    @staticmethod
    def modify_xml_file(filename, xpath, child1_dict, child2_dict, target_file):
        mytree = ET.parse(filename)
        myroot = mytree.getroot()
        for elm in myroot.findall(xpath):
            child1 = elm.find(child1_dict['name'])
            if child1.text == child1_dict['text']:
                # print(name.text)
                child2 = elm.find(child2_dict['name'])
                child2.text = str(child2_dict['text'])
        mytree.write(target_file, short_empty_elements=False)

    @staticmethod
    def modify_xml_file2(filename, tag, attr_dict1, attr_dict2, ns_dict, target_file):
        mytree = ET.parse(filename)
        myroot = mytree.getroot()
        for elm in myroot.iter():
            if elm.tag == tag:
                check_atrr_list = [k in elm.attrib and elm.attrib[k] == v for k, v in attr_dict1.items()]
                if all(check_atrr_list):
                    for k, v in attr_dict2.items():
                        elm.set(k, v)
        for k, v in ns_dict.items():
            ET.register_namespace(k, v)
        mytree.write(target_file, short_empty_elements=True)
        fin = open(target_file, "rt")
        content = fin.read()
        fin.close()
        text = "<?xml version='1.0'?>\n"
        text += content
        text = text.replace(" />", "/>")
        fout = open(target_file, "wt")
        fout.write(text)
        fout.close()

