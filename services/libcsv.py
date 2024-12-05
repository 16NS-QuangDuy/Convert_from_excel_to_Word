# encoding: utf-8
import os
import re
import time
from datetime import datetime
from settings.config import Config
from services.folder import Folder
import csv


class CsvML:
    """CsvML"""
    auto_worker_name = "CsvML"

    @staticmethod
    def parse_records_to_csv(header, records, save_file_name):
        text = ''
        text += ",".join(header)
        text += "\n"
        for rc in records:
            rc_str_list = []
            for field in rc:
                rc_str = "%s" % field
                rc_str = rc_str.strip()
                if "\"" in rc_str:
                    rc_str = rc_str.replace("\"", "\"\"")
                    rc_str = "\"%s\"" % rc_str
                rc_str_list.append(rc_str)
            text += ",".join(rc_str_list)
            text += "\n"
        with open(save_file_name, 'w+', encoding="utf-8") as f:
            f.write(text)
        return text

    # @staticmethod
    # def parse_csv_records(filename, encoding="utf-8"):
    #     csv_file = filename
    #     fin = open(csv_file, "rt", encoding=encoding, errors='ignore')
    #     records = []
    #     for line in fin.readlines():
    #         rc = line.split(",")
    #         records.append(rc)
    #     fin.close()
    #     return records

    @staticmethod
    def filter_csv_by_field(CsvFileName, col, regex_list, out_CsvFileName):
        """
        This function is to filter csv file by handle text file and CSV format
        """
        def is_in_value_list(file_i):
            basename = os.path.basename(file_i)
            for regex in regex_list:
                match = re.match(r'%s' % regex, basename)
                if match:
                    return True
            return False
        out_row_list = []
        with open(CsvFileName) as csv_file:
            csv_reader = csv.reader(csv_file, delimiter=',')
            line_count = 0
            for row in csv_reader:
                if line_count == 0:
                    col_idx = row.index(col)
                    row[0] = 'CID'
                    out_row_list.append(row)
                    print(f'        Column names are {", ".join(row)}')
                    line_count += 1
                else:
                    line_count += 1
                    if len(row) > col_idx:
                        file_i = row[col_idx]
                        if is_in_value_list(file_i):
                            out_row_list.append(row)
            print(f'        Processed {line_count} lines.')
        with open(out_CsvFileName, mode='w') as out_csv:
            csv_writer = csv.writer(out_csv, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
            csv_writer.writerows(out_row_list)

        f1 = open(out_CsvFileName, "r")
        lines = f1.readlines()
        f1.close()
        f2 = open(out_CsvFileName, "w+")
        for line in lines:
            if line.strip() == '':
                continue
            f2.write(line)
        f2.close()

    @staticmethod
    def parse_csv_records(CsvFileName):
        """
        This function is to filter csv file by handle text file and CSV format
        """
        out_row_list = []
        with open(CsvFileName) as csv_file:
            csv_reader = csv.reader(csv_file, delimiter=',')
            line_count = 0
            for row in csv_reader:
                if line_count == 0:
                    row[0] = 'CID'
                    out_row_list.append(row)
                    print(f'        Column names are {", ".join(row)}')
                    line_count += 1
                else:
                    line_count += 1
                    out_row_list.append(row)
            print(f'        Processed {line_count} lines.')
        return out_row_list
