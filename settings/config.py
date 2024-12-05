# encoding: utf-8
import os
import yaml
import time
import re
from docx.shared import RGBColor


class Config:
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    SRV_BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    WorkingDir = os.path.join(BASE_DIR)
    InputDir = os.path.join(BASE_DIR, "Input")
    OutputDir = os.path.join(BASE_DIR, "Output")
    DatabaseDir = os.path.join(OutputDir, "YamlDataBase")
    Excel2YamlDir = os.path.join(OutputDir, "YamlOutputFromExcel")
    Doc2YamlDir = os.path.join(OutputDir, "YamlOutputFromDoc")
    Yaml2ExcelDir = os.path.join(OutputDir, "ExcelOutputFromYaml")
    Yaml2IpxactDir = os.path.join(OutputDir, "IpxactOutputFromYaml")
    Doc2DocDir = os.path.join(OutputDir, "WordOutputFromWord")
    ExcelReportDir = os.path.join(OutputDir, "ExcelReportFromChecker")
    Doc2ExcelDir = os.path.join(OutputDir, "ExcelReportFromTracer")
    InputComparatorADir = os.path.join(InputDir, "Comparator", "A")
    InputComparatorBDir = os.path.join(InputDir, "Comparator", "B")
    OutputComparatorDir = os.path.join(OutputDir, "Comparator2WordUM")
    TempDir = os.path.join(OutputDir, "Temp")
    SessionList = []
    Debug = False

    def __init__(self):
        self.__auto_worker_name = self.__class__.__name__
        self.SessionName = "Session"
        self.ConfigMode = "Manual"
        self.RunExcel2Yaml = False
        self.RunYaml2Doc = False
        self.RunDoc2Yaml = False
        self.RunYaml2Excel = False
        self.CreateTS = False
        self.CreateUM = False
        self.RunAllDoc = False
        self.RunAllExcel = False
        self.RunAll = True
        self.WorkingScriptList = "PRESET_AUTOMATION_TASK_17"
        self.WorkingTS = []
        self.ScriptList = []
        self.Description = "Description of preset task"
        self.Request = ""
        self.input_dir = Config.InputDir
        self.config_file = None
        self.MultipleInput = False
        self.InputNameRegexText = False
        self.rgb_list = []
        self.output_dir = Config.OutputDir
        self.report_dir = Config.OutputDir
        self.input_data_items = []
        self.working_dir = os.path.join(self.output_dir,
                                        "%s_%s" % (self.SessionName, time.strftime("%Y%m%d-%H%M%S")))
        while self.working_dir in self.SessionList:
            self.working_dir = os.path.join(self.output_dir,
                                            "%s_%s" % (self.SessionName, time.strftime("%Y%m%d-%H%M%S")))
        self.SessionList.append(self.working_dir)
        self.TargetName = ""
        self.ActionList = []
        self.UserConfigDict = dict()
        # self.UserConfigDict = dict()
        self.ManagerConfigDict = dict()
        self.OrganizerConfigDict = dict()
        self.WorkerConfigDict = dict()

    def update_config_from_yaml(self, the_config_name=None):
        if the_config_name is not None:
            config_file = Config.get_default_config_file(the_script_file=__file__, the_config_name=the_config_name)
            with open(config_file, 'r', encoding="utf-8") as stream:
                preset_config_dict = yaml.safe_load(stream)
            if self.__auto_worker_name in preset_config_dict:
                preset_config_dict = preset_config_dict[self.__auto_worker_name]
                for key, value in preset_config_dict.items():
                    if key in self.__dict__:
                        self.__setattr__(key, value)
                    elif key in self.__class__.__dict__:
                        setattr(self, key, value)
                    else:
                        setattr(self, key, value)
        return self

    def update_config(self, **kwargs):
        for key, value in kwargs.items():
            if key in self.__dict__:
                self.__setattr__(key, value)
            elif key in self.__class__.__dict__:
                setattr(self, key, value)
            else:
                setattr(self, key, value)
            self.UserConfigDict[key] = value
        return self

    def update_config_by_yaml(self, config_file, auto_worker_name=None):
        with open(config_file, 'r', encoding="utf-8") as stream:
            config_dict = yaml.safe_load(stream)
            if auto_worker_name is None:
                key_list = [k for k, v in config_dict.items()]
                if len(key_list) > 0:
                    auto_worker_name = key_list[0]
            config_dict = config_dict[auto_worker_name]
        for key, value in config_dict.items():
            if key in self.__dict__:
                self.__setattr__(key, value)
            elif key in self.__class__.__dict__:
                setattr(self, key, value)
            else:
                setattr(self, key, value)
            self.UserConfigDict[key] = value
        return self

    def update_config_by_dict(self, config_dict):
        for key, value in config_dict.items():
            if key in self.__dict__:
                self.__setattr__(key, value)
            elif key in self.__class__.__dict__:
                setattr(self, key, value)
            else:
                setattr(self, key, value)
            self.UserConfigDict[key] = value
        return self

    def prior_user_config_dict(self):
        for key, value in self.UserConfigDict.items():
            setattr(self, key, value)
        return self

    def refresh_config(self):
        self.DatabaseDir = os.path.join(self.OutputDir, "YamlDataBase")
        self.Excel2YamlDir = os.path.join(self.OutputDir, "YamlOutputFromExcel")
        self.Doc2YamlDir = os.path.join(self.OutputDir, "YamlOutputFromDoc")
        self.Yaml2ExcelDir = os.path.join(self.OutputDir, "ExcelOutputFromYaml")
        self.Yaml2IpxactDir = os.path.join(self.OutputDir, "IpxactOutputFromYaml")
        self.Doc2DocDir = os.path.join(self.OutputDir, "WordOutputFromWord")
        self.ExcelReportDir = os.path.join(self.OutputDir, "ExcelReportFromChecker")
        self.Doc2ExcelDir = os.path.join(self.OutputDir, "ExcelReportFromTracer")
        self.InputComparatorADir = os.path.join(self.InputDir, "Comparator", "A")
        self.InputComparatorBDir = os.path.join(self.InputDir, "Comparator", "B")
        self.OutputComparatorDir = os.path.join(self.OutputDir, "Comparator2WordUM")

    @staticmethod
    def get_default_config_file(the_script_file, the_config_name=None, depth=5, localcase=""):
        """find the default config file path of the given script path"""
        if the_config_name is None:
            config_name = os.path.basename(the_script_file).replace(".py", "_config.yaml")
        else:
            config_name = the_config_name
        config_path = os.path.dirname(os.path.abspath(the_script_file))
        config_file = os.path.join(config_path, "config", config_name)
        if os.path.exists(config_file):
            return config_file
        text = ''
        for i in range(depth):
            folder = os.path.basename(config_path)
            config_path = os.path.dirname(config_path)
            text += folder
            text += os.sep
            config_file = os.path.join(config_path, "config", text, config_name)
            if os.path.exists(config_file):
                return config_file
        return None

    def get_config(self, key):
        if key in self.__dict__:
            return self.__dict__[key]
        elif key in self.__class__.__dict__:
            return getattr(self, key)
        return "None"

    @staticmethod
    def set_attr_from_yaml(obj, config_file, auto_worker_name=None):
        with open(config_file, 'r', encoding="utf-8") as stream:
            config_dict = yaml.safe_load(stream)
            if auto_worker_name is None:
                key_list = [k for k, v in config_dict.items()]
                if len(key_list) > 0:
                    auto_worker_name = key_list[0]
        if auto_worker_name in config_dict:
            config_dict = config_dict[auto_worker_name]
            if config_dict is None:
                return obj
            for key, value in config_dict.items():
                if obj.__class__.__name__ == "Config":
                    update_value = obj.UserConfigDict.get(key, None)
                    if update_value is None:
                        setattr(obj, key, value)
                else:
                    setattr(obj, key, value)
            if "ColumnDict" in config_dict:
                column_dict = Config.convert_char2ord(config_dict["ColumnDict"])
                setattr(obj, "ColumnOrdDict", column_dict)
            if "VariableDict" in config_dict:
                variable_dict = config_dict["VariableDict"]
                variable_dict['Column'] = ord(variable_dict['Column']) - ord('A')
                setattr(obj, "VariableDict", variable_dict)
            if "ConfigForLocalCaseDict" not in config_dict:
                setattr(obj, "ConfigForLocalCaseDict", None)
            if "Template" in config_dict:
                template_file = Config.get_template_file(config_file, config_dict["Template"])
                setattr(obj, "Template", template_file)
            if "RGB_list" in config_dict:
                rgb_list = []
                rgb_dict = config_dict["RGB_list"]
                for rgb in rgb_dict:
                    rgb_list.append(RGBColor(rgb[0], rgb[1], rgb[2]))
                setattr(obj, "RGB_list", rgb_list)
            column_dict_list = [key for key, value in config_dict.items() if key.endswith("ColumnDict")]
            for ColumnDict in column_dict_list:
                if ColumnDict == "ColumnDict":
                    continue
                column_dict = Config.convert_char2ord(config_dict[ColumnDict])
                prefix = ColumnDict.replace("ColumnDict", "")
                setattr(obj, prefix + "ColumnOrdDict", column_dict)
        return obj

    @staticmethod
    def convert_char2ord(column_dict):
        for key in column_dict:
            try:
                column_dict[key] = ord(column_dict[key]) - ord('A')
            except:
                pass
        return column_dict

    @staticmethod
    def convert_ord2char(column_dict):
        out_column_dict = dict()
        for key in column_dict:
            try:
                out_column_dict[key] = chr(column_dict[key] + ord('A'))
            except:
                out_column_dict[key] = column_dict[key]
                pass
        return out_column_dict

    @staticmethod
    def get_converted_file_name(input_filepath, output_dir, new_ext):
        filename, file_extension = os.path.splitext(input_filepath)
        if file_extension == "":
            converted_name = "%s%s" % (os.path.basename(input_filepath), new_ext)
        else:
            converted_name = os.path.basename(input_filepath).replace(file_extension, new_ext)
        converted_file_name = os.path.join(output_dir, converted_name)
        return converted_file_name

    @staticmethod
    def get_template_file(the_config_file, the_template_name, depth=5):
        """find the template file path of the given script path"""
        script_name = os.path.basename(the_config_file).replace("_config.yaml", ".py")
        script_path = os.path.dirname(os.path.abspath(the_config_file))
        script_file = os.path.join(script_path, script_name)
        if os.path.exists(script_file):
            template_path = os.path.dirname(os.path.abspath(script_file))
            template_file = os.path.join(template_path, "Templates", the_template_name)
            if os.path.exists(template_file):
                return template_file
            else:
                return None
        text = ''
        for i in range(depth):
            folder = os.path.basename(script_path)
            script_path = os.path.dirname(script_path)
            if folder != "config":
                text = folder + os.sep + text
            else:
                script_file = os.path.join(script_path, text, script_name)
                if os.path.exists(script_file):
                    template_path = os.path.dirname(os.path.abspath(script_file))
                    template_file = os.path.join(template_path, "Templates", the_template_name)
                    if os.path.exists(template_file):
                        return template_file
        script_path = os.path.dirname(os.path.abspath(script_file))
        text = ''
        for i in range(depth):
            folder = os.path.basename(script_path)
            script_path = os.path.dirname(script_path)
            text += folder
            text += os.sep
            template_file = os.path.join(script_path, "Templates", text, the_template_name)
            if os.path.exists(template_file):
                return template_file
            template_file = os.path.join(script_path, "Templates", the_template_name)
            if os.path.exists(template_file):
                return template_file
        return None

    def add_action(self, action_name):
        self.ActionList.append(action_name)

    def end_action(self):
        if len(self.ActionList) > 0:
            del self.ActionList[0]

    def match_action(self, action_name):
        if len(self.ActionList) > 0:
            if self.ActionList[0] == action_name:
                return True
        return False

    def set_server_base_dir(self, srv_base_dir=None):
        if srv_base_dir is not None:
            self.SRV_BASE_DIR = srv_base_dir

