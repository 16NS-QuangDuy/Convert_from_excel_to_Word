# encoding: utf-8
import os
import re
import time
from datetime import datetime
from settings.config import Config
from services.folder import Folder


class MakeFileML:
    """MakeFileML"""
    auto_worker_name = "MakeFileML"

    @staticmethod
    def parse_make_records2(filename):
        """
        SRC_FILES = $(MODULE_PATH)/generic/include/Adc.h@outdir=include \
                    $(MODULE_PATH)/generic/include/Adc_Types.h@outdir=include \
                    $(MODULE_PATH)/generic/src/Adc.c@outdir=src \
                    $(MODULE_PATH)/specific/$(PLATFORM_NAME)/config/Adc.xdm@outdir=config \
        """
        make_file = filename
        fin = open(make_file, "rt")
        file_text = fin.read()
        fin.close()
        records = []
        regex1 = r"SRC_FILES\s*\+*=\s*((\$\(MODULE_PATH\).*) \\)\n([	\s]+(\$\(MODULE_PATH\).*) \\\n)+"
        regex = r"SRC_FILES\s*[\+]*=[\s\w\$\(\)\/\.\\@=]+"
        for match in re.finditer(regex, file_text):
            if match:
                lines = match.group(0).split("\\\n")
                for line in lines:
                    line = line.strip()
                    line_parts = line.split("@")
                    if len(line_parts) > 0:
                        line = line_parts[0]
                    if line.startswith("SRC_FILES"):
                        line_parts = line.split("=")
                        if len(line_parts) > 0:
                            line = line_parts[1]
                            line = line.strip()
                    if line != "":
                        records.append(line)
        result = dict()
        result['var_list'] = ["$(MODULE_PATH)", "$(PLATFORM_NAME)"]
        result['path_list'] = records
        print('\n'.join(result['path_list']))
        return result

    @staticmethod
    def parse_make_records(filename):
        """
        SRC_FILES = $(MODULE_PATH)/generic/include/Adc.h@outdir=include \
                    $(MODULE_PATH)/generic/include/Adc_Types.h@outdir=include \
                    $(MODULE_PATH)/generic/src/Adc.c@outdir=src \
                    $(MODULE_PATH)/specific/$(PLATFORM_NAME)/config/Adc.xdm@outdir=config \
        """
        make_file = filename
        fin = open(make_file, "rt")
        lines = fin.readlines()
        fin.close()
        out_lines = []
        start_tag = False
        close_tag = False
        for line in lines:
            if line.startswith("SRC_FILES"):
                start_tag = True
                close_tag = False
            if start_tag and not close_tag:
                if line != "":
                    out_lines.append(line)
            if line.strip() in ["", "\n"]:
                close_tag = True
                start_tag = False
        records = []
        for line in out_lines:
            line = line.split("\\\n")[0]
            line = line.strip()
            line_parts = line.split("@")
            if len(line_parts) > 0:
                line = line_parts[0]
            if line.startswith("SRC_FILES"):
                line_parts = line.split("=")
                if len(line_parts) > 0:
                    line = line_parts[1]
                    line = line.strip()
            if line != "":
                records.append(line)
        result = dict()
        result['var_list'] = ["$(MODULE_PATH)", "$(PLATFORM_NAME)"]
        result['path_list'] = records
        return result
