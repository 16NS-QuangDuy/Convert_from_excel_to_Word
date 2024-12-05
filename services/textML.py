# encoding: utf-8
import yaml
import re
from settings.config import Config


class TextML:
    """Text"""
    auto_worker_name = "TextML"
    config_file = Config.get_default_config_file(the_script_file=__file__, the_config_name="textML_config.yaml")

    def __init__(self, rule_database_file=None, auto_worker_name="", description_key="data"):
        self.auto_worker_name = self.__class__.__name__
        if rule_database_file is None:
            self._rule_database_file = TextML.config_file
        if auto_worker_name !="":
            self.auto_worker_name = auto_worker_name
        self._rule_database_file = rule_database_file
        self._target_spec_config_dict = {}
        self._description_parts_list = []
        self._description_extract_dict = {}
        self._description_part_dict_list = []
        self._description_key = description_key
        self.read_config_info()

    def read_config_info(self):
        with open(self._rule_database_file, 'r', encoding="utf-8") as stream:
            target_spec_config_dict = yaml.safe_load(stream)
        if self.auto_worker_name in target_spec_config_dict:
            self._target_spec_config_dict = target_spec_config_dict[self.auto_worker_name]
            self._description_parts_list = self._target_spec_config_dict["descriptionPartList"]
            self._description_extract_dict = self._target_spec_config_dict["extractPartsDict"]
            self._description_part_dict_list = self._target_spec_config_dict["descriptionPartDictList"]

    def split_parts(self, description):
        """check description has oneline is parsed or not
        :param description: input of a description
        :return: a dictionary result
        """
        result_dict = {}
        for key in self._description_parts_list:
            result_dict[key] = ""
        description = re.sub(r"\n\s*\+", " +", description)
        result_dict["original_description"] = description
        result_dict[self._description_key] = description
        for description_part_dict_name in self._description_part_dict_list:
            real_description_part_dict = self._target_spec_config_dict[description_part_dict_name]
            description_part = str(result_dict[self._description_key])
            description_part = description_part.strip()
            for k, v in real_description_part_dict.items():
                regex = v['regex']
                if isinstance(regex, list):
                    regex = ''.join(regex)
                match = re.search(r'%s' % regex, description_part)
                if match:
                    for extract_elm in self._description_extract_dict[description_part_dict_name]:
                        if extract_elm in v:
                            extract_elm_value = str(v[extract_elm])
                            for key in v['matchgroup']:
                                try:
                                    value = match.group(v['matchgroup'][key])
                                except:
                                    value = None
                                if value is None:
                                    value = ""
                                extract_elm_value = extract_elm_value.replace('{%s}' % key, value)
                            result_dict[extract_elm] = extract_elm_value.strip()
                    break
        for key in self._description_parts_list:
            result_dict[key] = str(result_dict[key]).strip()
        if result_dict["descriptiontype"] == "":
            result_dict[self._description_key] = ""
        result_dict["displayname"] = result_dict[self._description_key]
        return result_dict

    @staticmethod
    def remove_brace(description):
        return description.replace("(", "").replace(")", "")

    @staticmethod
    def merge_multiple_lines_to_one_line(description):
        lines = description.splitlines()
        text = ''
        if len(lines) >= 2:
            for line in lines:
                text += line
        else:
            text = description
        text = text.replace("  ", " ")
        text = text.replace("_x000D_", "")
        return text

    @staticmethod
    def standardize(st, input_list=[]):
        ignore_list = ['in', 'on', 'the', 'of', 'and', 'for', 'with', 'from', 'to', 'via', 'along']
        if input_list:
            for li in input_list:
                ignore_list.extend(li)
        lst = []
        ignore_list = [word.lower() for word in ignore_list]
        for word in st.split():
            if word.lower() not in ignore_list and len(word) >= 2 and not re.match(r'^[a-z]=.*', word):
                lst.append(word[0].upper() + word[1:])
            elif word.lower() in ignore_list:
                lst.append(word.lower())
            else:
                lst.append(word)
        s = " ".join(lst)
        return s
