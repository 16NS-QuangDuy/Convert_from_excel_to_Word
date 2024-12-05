# encoding: utf-8
import re
import copy
import yaml
import time
from settings.config import Config
from services.folder import Folder


class YamlML:
    """handler for doing YAML file creation and reading"""
    auto_worker_name = __name__.split(".")[-1]
    auto_worker_name = auto_worker_name[0].upper() + auto_worker_name[1:]

    config_file = Config.get_default_config_file(the_script_file=__file__)
    with open(config_file, 'r', encoding="utf-8") as stream:
        target_spec_config_dict = yaml.safe_load(stream)
        target_spec_config_dict = target_spec_config_dict[auto_worker_name]

    @staticmethod
    def parse_text_to_yaml(input_text, key, indent, must=False):
        text = ''
        if input_text is None and must is False:
            return text
        elif input_text == "" and must is False:
            return text
        elif input_text is not None:
            input_text = input_text.replace('\t', "    ")
            input_text = input_text.strip()
            lines = input_text.splitlines()

            if "-" == key:
                if indent > 1:
                    text += '    ' * (indent - 1) + "- " + input_text + "\n"
                else:
                    text += "- " + input_text + "\n"
                return text
            if len(lines) <= 0:
                text += '    ' * indent + key + ":" + " " + "\n"
            elif len(lines) == 1:
                if lines[0].startswith("*"):
                    text += '    ' * indent + key + ": " + "\'" + lines[0].strip() + "\'" + "\n"
                elif lines[0] == "ON" or lines[0] == "OFF":
                    text += '    ' * indent + key + ": " + "\'" + lines[0].strip() + "\'" + "\n"
                elif ":" in lines[0] or lines[0].startswith("-") or lines[0].startswith("—") or ("[" in lines[0] and "]" in lines[0]):
                    if "\'" in lines[0]:
                        text += '    ' * indent + key + ": " + "\"" + lines[0] + "\"" + "\n"
                    else:
                        text += '    ' * indent + key + ": " + "\'" + lines[0] + "\'" + "\n"
                elif YamlML.is_reserved_bit_name(text):
                    text += '    ' * indent + key + ": " + "—" + "\n"
                else:
                    text += '    '*indent + key + ": " + lines[0].strip() + "\n"
            elif len(lines) >= 2:
                text += '    ' * indent + key + ": " + "|" + "\n"
                for para in lines:
                    text += '    ' * (indent+1) + para.strip() + "\n"
        return text

    @staticmethod
    def organize_bit(start, end):
        try:
            bit = ""
            if start == end:
                bit = "%s" % start
            elif abs(int(end)-int(start)) == 1:
                bit = "%s, %s" % (start, end)
            else:
                bit = "%s to %s" % (start, end)
            return bit
        except Exception as e:
            print ("       ERROR: ", str(e))
            return bit
            pass

    @staticmethod
    def read_record_field(rc_dict, rc_key):
        try:
            value = None
            if rc_key in rc_dict:
                value = rc_dict[rc_key]
            if isinstance(value, int):
                value = "%s" % value
            if value is None:
                value = ""
            if isinstance(value, dict):
                return value
            if isinstance(value, str):
                value = value.strip()
            return value
        except:
            print ("rc_dict", rc_dict)
            print ("rc_key", rc_key)
            pass

    @staticmethod
    def revert_bitfield_to_bit_tbl(src_record):
        dst_records = src_record
        src_full_hdr = ["Register name", "Bit Start", "Bit End", "Bit Name", "Initial Value", "R/W", "Description"]
        dst_full_hdr = ["Register name", "Bit Start", "Bit End", "Bit Name", "Initial Value", "R/W", "Description"]
        reg_bit_dict = {}
        start_bit = int(dst_records["Bit Start"])
        end_bit = int(dst_records["Bit End"])
        size_bit = start_bit - end_bit + 1
        reg_bit_dict["size_bit"] = size_bit
        if size_bit <= 1:
            reg_bit_dict["bit_list"] = [dst_records["Bit Start"]]
            reg_bit_dict["init_list"] = [str(dst_records["Initial Value"])]
            reg_bit_dict["rw_list"] = [str(dst_records["R/W"])]
            reg_bit_dict["name_list"] = [str(dst_records["Bit Name"])]
            if dst_records["Initial Value"] is None:
                reg_bit_dict["init_list"] = [""]
            elif dst_records["Initial Value"].startswith("B'"):
                initial = dst_records["Initial Value"].replace("B'", "").strip()
                reg_bit_dict["init_list"] = [initial]
            if dst_records["R/W"] is None:
                reg_bit_dict["R/W"] = [""]
            return reg_bit_dict
        if size_bit >= 2:
            bit_list = []
            rw_list = []
            name_list = []
            for bit_idx in range(start_bit, end_bit - 1, -1):
                bit_list.append(bit_idx)
                rw_list.append(dst_records["R/W"])
                # Variable bit name
                name_list.append(dst_records["Bit Name"])
            reg_bit_dict["bit_list"] = bit_list
            reg_bit_dict["rw_list"] = rw_list
            reg_bit_dict["name_list"] = name_list
            init_list = []
            initial = dst_records["Initial Value"]
            initial = initial.replace(" ", "").replace("’", "'").replace("-", "—")
            matchObj = re.match(r'([0-9]+)', initial, re.M | re.I)
            if matchObj and ("B" not in initial and "H" not in initial and "b" not in initial and not "h" in initial):
                number = int(initial, 10)
                initial = f'{number:032b}'
                initial = initial[32 - size_bit:]
                init_list.extend(initial)
                reg_bit_dict["init_list"] = init_list
                return reg_bit_dict
            if "ALL" in initial.upper():
                initial = initial.upper()
                initial = initial.replace("ALL", "").strip()
                for bit_idx in range(start_bit, end_bit - 1, -1):
                    init_list.append(initial)
                reg_bit_dict["init_list"] = init_list
                return reg_bit_dict
            if initial.startswith("B"):
                m = re.search('B.+?([01]+).*?', initial)
                if m:
                    initial = m.group(1)
                    number = int(initial, 2)
                    initial = f'{number:032b}'
                    initial = initial[32 - size_bit:]
                    init_list.extend(initial)
                    reg_bit_dict["init_list"] = init_list
                    return reg_bit_dict
            if initial.startswith("H"):
                m = re.search('H.+?([0-9A-F]+).+?', initial)
                if m:
                    initial = m.group(1)
                    number = int(initial, 16)
                    initial = f'{number:032b}'
                    initial = initial[32 - size_bit:]
                    init_list.extend(initial)
                    reg_bit_dict["init_list"] = init_list
                    return reg_bit_dict
            if "—" in initial:
                for bit_idx in range(start_bit, end_bit - 1, -1):
                    init_list.append("—")
                reg_bit_dict["init_list"] = init_list
                return reg_bit_dict
            if "*1" == initial:
                for bit_idx in range(start_bit, end_bit - 1, -1):
                    init_list.append(initial)
                reg_bit_dict["init_list"] = init_list
                return reg_bit_dict
            if "" == initial:
                for bit_idx in range(start_bit, end_bit - 1, -1):
                    init_list.append("0/1")
                reg_bit_dict["init_list"] = init_list
                return reg_bit_dict
            matchObj = re.match(r'([\d]+)([bB]\')([01]+)', initial, re.M | re.I)
            if matchObj:
                length = int(matchObj.group(1), 10)
                number = int(matchObj.group(3), 2)
                initial = f'{number:032b}'
                init_list.extend(initial[32 - length:])
                reg_bit_dict["init_list"] = init_list
                return reg_bit_dict
            matchObj = re.match(r'([\d]+)(\'[bB])([01]+)', initial, re.M | re.I)
            if matchObj:
                length = int(matchObj.group(1), 10)
                number = int(matchObj.group(3), 2)
                initial = f'{number:032b}'
                init_list.extend(initial[32 - length:])
                reg_bit_dict["init_list"] = init_list
                return reg_bit_dict
            matchObj = re.match(r'([\d]+)([hH]\')([0-9A-F]+)', initial, re.M | re.I)
            if matchObj:
                length = int(matchObj.group(1), 10)
                number = int(matchObj.group(3), 16)
                initial = f'{number:032b}'
                initial = initial[32 - length:]
                init_list.extend(initial)
                reg_bit_dict["init_list"] = init_list
                return reg_bit_dict
            matchObj = re.match(r'([\d]+)(\'[hH])([0-9A-F]+)', initial, re.M | re.I)
            if matchObj:
                length = int(matchObj.group(1), 10)
                number = int(matchObj.group(3), 16)
                initial = f'{number:032b}'
                initial = initial[32 - length:]
                init_list.extend(initial)
                reg_bit_dict["init_list"] = init_list
                return reg_bit_dict
            matchObj = re.match(r'^([0-9A-F]+)(h)$', initial, re.M | re.I)
            if matchObj:
                number = int(matchObj.group(1), 16)
                initial = f'{number:032b}'
                initial = initial[32 - size_bit:]
                init_list.extend(initial)
                reg_bit_dict["init_list"] = init_list
                return reg_bit_dict
            for bit_idx in range(start_bit, end_bit - 1, -1):
                init_list.append("")
            reg_bit_dict["init_list"] = init_list
            return reg_bit_dict

    @staticmethod
    def reorganize_field_records(src_records, hdr=None, bit_keys=None):
        if hdr is None or bit_keys is None:
            hdr = ["Bit", "Bit Name", "Initial Value", "R/W", "Description", "Ndisclosed"]
            bit_keys = ["Bit", "BitName", "Initial", "R/W", "BitDesc", "Ndisclosed"]
        dst_records = []
        reg_desc_dict = {}
        checked_keys = ["Bit Start", "Bit End", "Bit Name", "BitName", "Bit"]
        for rc_dict in src_records:
            bitfields = YamlML.parse_bit_field_bit(rc_dict["Bit"], rc_dict["Bit Name"])
            for bitfield in bitfields:
                reg_desc_dict["Bit Start"] = bitfield["Bit Start"]
                reg_desc_dict["Bit End"] = bitfield["Bit End"]
                reg_desc_dict["Bit Name"] = bitfield["Bit Name"]
                for idx in range (len(bit_keys)):
                    field = bit_keys[idx]
                    if field in checked_keys:
                        continue
                    else:
                        hdr_key = hdr[idx]
                        reg_desc_dict[hdr_key] = rc_dict[hdr_key]
                if "Ndisclosed" in rc_dict:
                    reg_desc_dict["Ndisclosed"] = rc_dict["Ndisclosed"]
                initial = YamlML.parse_initial_value(bitfield["Bit Start"], bitfield["Bit End"], rc_dict["Initial Value"])
                reg_desc_dict["Initial Value"] = initial
                dst_records.append(copy.deepcopy(reg_desc_dict))
        return dst_records

    @staticmethod
    def parse_initial_value(start, end, input_initial):
        try:
            size = abs(int(start) - int(end)) + 1
            input_initial = "%s" % input_initial
            if size >= 2:
                if "All" in input_initial:
                    return input_initial
                if input_initial in ['0', '1']:
                    number = int(input_initial, 2)
                    temp = f'{number:032b}'
                    initial = temp[32 - size:]
                    return "B'%s" % initial
                if input_initial in ['-', '—']:
                    return "All %s" % input_initial
                return input_initial
            elif input_initial in ['0', '1']:
                return "B'%s" % input_initial
            else:
                return input_initial
        except:
            print ("        ERROR: parse_initial_value", start, end, input_initial)
            return input_initial

    @staticmethod
    def parse_bit_field_bit(input_bit, input_bitname):
        bitfield_list = []
        input_bit = "%s" % input_bit
        input_bit = input_bit.strip()
        var_bitfield_list = YamlML.parse_var_bit_field_bit(input_bit, input_bitname)
        if len(var_bitfield_list) >= 1:
            bitfield = var_bitfield_list[0]
            start = bitfield["Bit Start"]
            end = bitfield["Bit End"]
            for k in range(int(start), int(end) - 1, -1):
                short_bitname = bitfield["Short Bit Name"]
                short_bitname = "%s%s" % (short_bitname[:len(short_bitname) - 1], k)
                bitfield_list.append({"Bit Start": "%s" % k, "Bit End": "%s" % k, "Bit Name": short_bitname})
            return bitfield_list
        for bitfield01 in input_bit.split("\n"):
            bit_list = []
            for bitfield in bitfield01.split(","):
                if "to" in bitfield:
                    (start, end) = bitfield.split("to")
                    bitfield_list.append({"Bit Start": start.strip(), "Bit End": end.strip(), "Bit Name": input_bitname})
                elif "～" in bitfield:
                    (start, end) = bitfield.split("～")
                    bitfield_list.append({"Bit Start": start.strip(), "Bit End": end.strip(), "Bit Name": input_bitname})
                else:
                    if bitfield.strip() != "":
                        bit_list.append(bitfield.strip())
            if len(bit_list) >= 1:
                if len(bit_list) % 2 == 1:
                    for bit in bit_list:
                        bitfield_list.append(
                            {"Bit Start": bit, "Bit End": bit, "Bit Name": input_bitname})
                else:
                    last = 0
                    for idx, bit in enumerate(bit_list):
                        if idx < last:
                            continue
                        if int(bit_list[idx]) == int(bit_list[idx + 1]) + 1:
                            (start, end) = (bit_list[idx], bit_list[idx + 1])
                            bitfield_list.append(
                                {"Bit Start": start.strip(), "Bit End": end.strip(), "Bit Name": input_bitname})
                            last = idx + 2
        return bitfield_list

    @staticmethod
    def parse_bit_field_bit_yaml2excel(input_bit, input_bitname):
        bitfield_list = []
        input_bit = "%s" % input_bit
        input_bit = input_bit.strip()
        var_bitfield_list = YamlML.parse_var_bit_field_bit(input_bit, input_bitname)
        if len(var_bitfield_list) >= 1:
            bitfield = var_bitfield_list[0]
            start = bitfield["Bit Start"]
            end = bitfield["Bit End"]
            for k in range(int(start), int(end) - 1, -1):
                short_bitname = bitfield["Short Bit Name"]
                short_bitname = "%s%s" % (short_bitname[:len(short_bitname) - 1], k)
                bitfield_list.append({"Bit Start": "%s" % k, "Bit End": "%s" % k, "Bit Name": short_bitname})
            return bitfield_list
        for bitfield01 in input_bit.split("\n"):
            bit_list = []
            for bitfield in bitfield01.split(","):
                match = re.match("\[(\d+):(\d+)\]", bitfield)
                if match:
                    (start, end) = (match.group(1), match.group(2))
                    bitfield_list.append(
                        {"Bit Start": start.strip(), "Bit End": end.strip(), "Bit Name": input_bitname})
                    continue
                match = re.match(r"\[(\d+)\]", bitfield)
                if match:
                    (start, end) = (match.group(1), match.group(1))
                    bitfield_list.append(
                        {"Bit Start": start.strip(), "Bit End": end.strip(), "Bit Name": input_bitname})
                    continue
                if "to" in bitfield:
                    (start, end) = bitfield.split("to")
                    bitfield_list.append({"Bit Start": start.strip(), "Bit End": end.strip(), "Bit Name": input_bitname})
                else:
                    bit_list.append(bitfield.strip())
            if len(bit_list) >= 1:
                if len(bit_list) == 1:
                    bitfield_list.append(
                        {"Bit Start": bit_list[0], "Bit End": bit_list[0], "Bit Name": input_bitname})
                else:
                    last = 0
                    for idx, bit in enumerate(bit_list):
                        if idx < last:
                            continue
                        try:
                            if idx + 1 >= len(bit_list):
                                break
                            if int(bit_list[idx]) == int(bit_list[idx + 1]) + 1:
                                (start, end) = (bit_list[idx], bit_list[idx + 1])
                                bitfield_list.append(
                                    {"Bit Start": start.strip(), "Bit End": end.strip(), "Bit Name": input_bitname})
                                last = idx + 2
                        except:
                            print("There is an error")
        return bitfield_list

    @staticmethod
    def parse_var_bit_field_bit(input_bit, input_bitname):
        bitfield_list = []
        input_bit = "%s" % input_bit
        input_bit = input_bit.strip()
        input_bitname = "%s" % input_bitname
        input_bitname = input_bitname.strip()
        if "k" in input_bitname or "m" in input_bitname or "k" in input_bit or "m" in input_bit:
            input_bit = input_bit.replace("\n", "").replace(" ", "")
            input_bitname = input_bitname.replace("\n", "").replace(" ", "")
            match01 = re.match(r"([A-Z]+)([km])\(([km]+)=(\d+)to(\d+)\)", input_bitname, re.M)
            if match01:
                start = match01.group(5)
                end = match01.group(4)
                bitname = "%s%s(%s[bit]=%s to %s)" % (match01.group(1), match01.group(2),  match01.group(3), match01.group(4), match01.group(5))
                short_bitname = "%s%s" % (match01.group(1), match01.group(2))
                bitfield_list.append({"Bit Start": start, "Bit End": end, "Bit Name": bitname, "Short Bit Name": short_bitname})
                return bitfield_list
            match02 = re.match(r"([kmKM])\(([km]+)=(\d+)to(\d+)\)", input_bit, re.M)
            if match02:
                start = match02.group(4)
                end = match02.group(3)
                bitname = "%s(%s[bit]=%s to %s)" % (input_bitname, match02.group(2), match02.group(3), match02.group(4))
                short_bitname = "%s%s" % (match02.group(1), match02.group(2))
                bitfield_list.append({"Bit Start": start, "Bit End": end, "Bit Name": bitname, "Short Bit Name": short_bitname})
                return bitfield_list
            match03 = re.match(r"([A-Z]+)([km])\(([km]+)=(\d+)~(\d+)\)", input_bitname, re.M)
            if match03:
                start = match03.group(3)
                end = match03.group(4)
                short_bitname = "%s%s" % (match03.group(1), match03.group(2))
                bitname = "%s(%s[bit]=%s to %s)" % (short_bitname, match03.group(2), match03.group(3), match03.group(4))
                bitfield_list.append({"Bit Start": start, "Bit End": end, "Bit Name": bitname, "Short Bit Name": short_bitname})
                return bitfield_list
            match04 = re.match(r"([A-Z]+)([km])\(([km]+\[bit\])=(\d+)to(\d+)\)", input_bitname, re.M) #STk(k[bit]=0to31
            if match04:
                start = match04.group(5)
                end = match04.group(4)
                short_bitname = "%s%s" % (match04.group(1), match04.group(2))
                bitname = "%s(%s[bit]=%s to %s)" % (short_bitname, match04.group(2), match04.group(4), match04.group(5))
                bitfield_list.append({"Bit Start": start, "Bit End": end, "Bit Name": bitname, "Short Bit Name": short_bitname})
                return bitfield_list
        return bitfield_list

    @staticmethod
    def create_all_records_from_field_dict(bits, hdr=None, bit_keys=None):
        try:
            if hdr is None or bit_keys is None:
                hdr = ["Bit", "Bit Name", "Initial Value", "R/W", "Description", "Ndisclosed"]
                bit_keys = ["Bit", "BitName", "Initial", "R/W", "BitDesc", "Ndisclosed"]
            records = []
            reg_desc_dict = {}
            if bits is not None:
                for bit_idx, key in enumerate(sorted(bits)):
                    for idx, bit_key in enumerate(bit_keys):
                        if bit_key in bits[key]:
                            bit_value = bits[key][bit_key]
                            if bit_value is not None:
                                reg_desc_dict[hdr[idx]] = bit_value
                                continue
                        reg_desc_dict[hdr[idx]] = ""
                    records.append(copy.deepcopy(reg_desc_dict))
            else:
                for key in hdr:
                    reg_desc_dict[key] = "ERR001"
                records.append(reg_desc_dict)
            dst_records = YamlML.reorganize_field_records(records, hdr, bit_keys)
        except Exception as e:
            print ("Error: create_all_records_from_field_dict" + str(e))
            print ("Error: create_all_records_from_field_dict", bits)
            pass
        return dst_records

    @staticmethod
    def create_all_records_from_field_dict_for_multil(bits, hdr=None, bit_keys=None):
        try:
            if hdr is None or bit_keys is None:
                hdr = ["Bit", "Bit Name", "Initial Value", "R/W", "Description", "Ndisclosed"]
                bit_keys = ["Bit", "BitName", "Initial", "R/W", "BitDesc", "Ndisclosed"]
            records = []
            reg_desc_dict = {}
            if bits is not None:
                for bit_idx, key in enumerate(sorted(bits)):
                    for idx, bit_key in enumerate(bit_keys):
                        if bit_key in bits[key]:
                            bit_value = bits[key][bit_key]
                            if bit_value is not None:
                                reg_desc_dict[hdr[idx]] = bit_value
                                continue
                        reg_desc_dict[hdr[idx]] = ""
                    records.append(copy.deepcopy(reg_desc_dict))
            else:
                for key in hdr:
                    reg_desc_dict[key] = "ERR001"
                records.append(reg_desc_dict)
            # dst_records = YamlML.reorganize_field_records(records, hdr, bit_keys)
        except Exception as e:
            print ("Error: create_all_records_from_field_dict" + str(e))
            print ("Error: create_all_records_from_field_dict", bits)
            pass
        return records

    @staticmethod
    def organize_format_bit_fields(bits):
        if len(bits["Bit"]) >= 32:
            bit_name_15 = bits["BitName"][15]
            bit_name_16 = bits["BitName"][16]
            if "[" in bit_name_15 and "[" in bit_name_16:
                (arr_name15, idx15) = bit_name_15.split("[")
                (arr_name16, idx16) = bit_name_16.split("[")
                if arr_name15 == arr_name16:
                    if ":" not in bit_name_15:
                        (start15, end15) = bit_name_15.split("]")
                    else:
                        (start15, end15) = bit_name_15.split(":")
                    if ":" not in bit_name_16:
                        (start16, end16) = bit_name_16.split("[")
                    else:
                        (start16, end16) = bit_name_16.split(":")
                    idx1516 = start15 + ':' + end16
                    for idx, bit_name in enumerate(bits["BitName"]):
                        if bit_name in [bit_name_15, bit_name_16]:
                            bits["BitName"][idx] = idx1516
        bitfields = []
        end = len(bits["BitName"])
        last = 0
        step = 1
        if len(bits["Initial"]) <= 0:
            for bit_idx in range(0, end, step):
                bits["Initial"].append("0/1")
        if len(bits["R/W"]) <= 0:
            for bit_idx in range(0, end, step):
                bits["R/W"].append("-")

        for bit_idx in range(0, end, step):
            if bit_idx < last:
                continue
            bit_idx_list = [bits["Bit"][bit_idx]]
            bit_name = bits["BitName"][bit_idx]
            initial = [bits["Initial"][bit_idx]]
            rw = [bits["R/W"][bit_idx]]
            for next_idx in range(bit_idx + 1, end, 1):
                if bits["BitName"][next_idx] == bits["BitName"][bit_idx]:
                    initial.append(bits["Initial"][next_idx])
                    bit_idx_list.append(bits["Bit"][next_idx])
                    rw.append(bits["R/W"][next_idx])
                else:
                    break
            last += len(bit_idx_list)
            if len(bit_idx_list) >= 2:
                from_bit_idx = 0
                to_bit_idx = len(bit_idx_list) - 1
                if len(bit_idx_list) >= 3:
                    bit_idx_list = [bit_idx_list[from_bit_idx] + " to " + bit_idx_list[to_bit_idx]]
                else:
                    bit_idx_list = [bit_idx_list[from_bit_idx] + ", " + bit_idx_list[to_bit_idx]]
                if initial.count(initial[0]) == len(initial):
                    if initial[0] in ['0', '1']:
                        initial = ["All %s" % initial[0]]
                    elif initial[0] in ['-', '—']:
                        initial = ["—"]
                    else:
                        initial = ["All %s" % initial[0]]
                else:
                    num_int = int(''.join(initial), 2)
                    initial = ["%sH'%X" % (len(initial), num_int)]
                if rw.count(rw[0]) == len(rw):
                    rw = [rw[0]]
            else:
                if initial[0] in ['0', '1']:
                    initial = ["B'%s" % initial[0]]
            bitfields.append([bit_idx_list, bit_name, initial, rw])
        return bitfields

    @staticmethod
    def reorganize_format_bit_fields(bitfields, format_type=1):
        bitfields1 = []
        if format_type == 2:
            reserved_list = []
            for j, bf in enumerate(bitfields):
                if bf[1] == chr(45) or bf[1] == '':
                    reserved_list.append(j)
            if len(reserved_list) >= 2:
                bit_idx_list = []
                for r in reserved_list:
                    bf = bitfields[r]
                    bit_idx_list.extend(bf[0])
                bit_idx_list_str = ", ".join(bit_idx_list)
                for j, bf in enumerate(bitfields):
                    if j in reserved_list:
                        if j <= reserved_list[0]:
                            bf = bitfields[j]
                            bitfields1.append([[bit_idx_list_str], bf[1], bf[2], bf[3]])
                    else:
                        bitfields1.append(bf)
            else:
                bitfields1 = bitfields
        return bitfields1

    @staticmethod
    def organize_format_hexa_string(baseaddress):
        newtext = baseaddress
        try:
            if len(newtext) >= 9:
                p_hexa = re.compile(r"([0-9ABCDEF]){8}$", re.IGNORECASE)
                phexa = p_hexa.search(baseaddress)
                if phexa is not None:
                    print ("     ", newtext)
                    temp = phexa.group(0)
                    newtext = "H'%s %s" % (temp[:4], temp[4:])
            return newtext
        except:
            print ("      ERROR: SPECIAL VALUE")
            return newtext

    @staticmethod
    def is_reserved_bit_name(bitname):
        bitname = bitname.strip()
        return bitname in ["-", "—", "¾", "—", "-"]

    @staticmethod
    def correct_reserved_symbol(bitname):
        bitname = bitname.strip()
        if bitname in ["-", "—", "¾", "—", "-"]:
            return "—"
        else:
            return bitname


    @staticmethod
    def format_intial_hexa(initial):
        newtext = initial
        int_regex = re.compile(r"^([+-]?[0-9]\d*|0)$")
        match = int_regex.match(str(newtext).strip())
        if match:
            value = int(match.group(1))
            newtext = format(value, "04X")
            newtext = "H'%s" % newtext
            return newtext
        matchObj = re.match(r'([\d]+)(hH\')([0-9A-F]+)', initial, re.M | re.I)
        if matchObj:
            newtext = "%sH'%s" % (matchObj.group(1), matchObj.group(3) )
            return newtext
        matchObj = re.match(r'([\d]+)(\'hH)([0-9A-F]+)', initial, re.M | re.I)
        if matchObj:
            newtext = "%sH'%s" % (matchObj.group(1), matchObj.group(3) )
            return newtext
        matchObj = re.match(r'^([0-9A-F]+)(h)$', initial, re.M | re.I)
        if matchObj:
            newtext = "H'%s" % (matchObj.group(1))
            return newtext
        return newtext

    @staticmethod
    def parse_second_text_to_yaml(input_text, indent):
        text = ''
        input_text = input_text.strip()
        lines = input_text.splitlines()
        for para in lines:
                text += '    ' * (indent+1) + para.strip() + "\n"
        return text

    @staticmethod
    def parse_list_to_yaml(input_text_list, key, indent, must=False):
        text = ''
        if input_text_list is None or len(input_text_list) <= 0 and must is False:
            return text
        if len(input_text_list) <= 0 and must is True:
            text += '    '*indent + key + ":" + " " + "\n"
        elif len(input_text_list) == 1:
            if ":" in input_text_list[0] or input_text_list[0].startswith("-"):
                text += '    ' * indent + key + ":" + " |" + "\n"
                text += YamlML.parse_second_text_to_yaml(input_text_list[0], indent)
            else:
                text += YamlML.parse_text_to_yaml(input_text_list[0], key, indent)
        elif len(input_text_list) >= 2:
            text += '    ' * indent + key + ":" + " |" + "\n"
            for para in input_text_list:
                text += YamlML.parse_second_text_to_yaml(para, indent)
        return text

    @staticmethod
    def filter_out_nondisclosed_reg_desc(input_register_dict):
        nondisclosed_dict = YamlML.target_spec_config_dict["NdisclosedDict"]
        sk = nondisclosed_dict["Key"]
        sv = nondisclosed_dict["Value"]
        register_dict = input_register_dict
        if register_dict is not None:
            for register_id in sorted(register_dict):
                cur_register_dict = register_dict[register_id]
                if sk in cur_register_dict:
                    if cur_register_dict[sk] == sv:
                        print("         del: ", "{0} : {1}".format(cur_register_dict['Name'].strip(), sv))
                        del register_dict[register_id]
            for register_id in sorted(register_dict):
                cur_register_dict = register_dict[register_id]
                bf_dict = cur_register_dict["Bits"]
                for bf_id in sorted(bf_dict):
                    cur_bf_dict = bf_dict[bf_id]
                    if sk in cur_bf_dict:
                        if cur_bf_dict[sk] == sv:
                            print("         del: ", "{0} : {1}".format(bf_id, sv))
                            del bf_dict[bf_id]
        return register_dict

    @staticmethod
    def filter_out_nondisclosed_reg_config(input_register_dict):
        nondisclosed_dict = YamlML.target_spec_config_dict["NdisclosedDict"]
        sk = nondisclosed_dict["Key"]
        sv = nondisclosed_dict["Value"]
        register_dict = input_register_dict
        if register_dict is not None:
            for register_tbl_id in register_dict:
                config_tbl_dict = register_dict[register_tbl_id]
                for register_id in sorted(config_tbl_dict):
                    cur_register_dict = config_tbl_dict[register_id]
                    if sk in cur_register_dict:
                        if cur_register_dict[sk] == sv:
                            print("         del: ", "{0} : {1}".format(cur_register_dict['Register'].strip(), sv))
                            del config_tbl_dict[register_id]
        return register_dict

    @staticmethod
    def get_nondisclosed_key_value():
        nondisclosed_dict = YamlML.target_spec_config_dict["NdisclosedDict"]
        sk = nondisclosed_dict["Key"]
        sv = nondisclosed_dict["Value"]
        return (sk, sv)

    @staticmethod
    def get_indent_dict_by_tbl_name(table_name="RegisterDescription"):
        if isinstance(table_name, str):
            table_name = table_name + "TableDict"
            if table_name in YamlML.target_spec_config_dict:
                return YamlML.target_spec_config_dict[table_name]["IndentDict"]
        return None

    @staticmethod
    def get_bit_description_table_type_info(bits):
        if bits is not None:
            table_dict = YamlML.target_spec_config_dict["BitDescriptionTableDict"]
            for k, v in table_dict.items():
                keys_list = v["KeyList"]
                match = 0
                for key in keys_list:
                    if key in bits:
                        match += 1
                if match >= len(keys_list):
                    return v
        return None

    @staticmethod
    def get_root_from_file(input_dir, file_name):
        yaml_root_dict = None
        files = Folder.get_all_files(input_dir, file_name)
        if len(files) >= 1:
            with open(files[0], 'r', encoding="utf-8") as stream:
                yaml_root_dict = yaml.safe_load(stream)
        return yaml_root_dict

    @staticmethod
    def create_header_yaml_file(top_yaml_key, info="Register Description", header=True):
        text = ''
        if header:
            text += '--- \n'
            text += '#%s of Target Spec\n' % info
            text += '#Created time: ' + time.strftime("%Y/%m/%d") + '\n'
        text += '%s: \n' % top_yaml_key
        return text

    @staticmethod
    def get_variable_dict_from_yaml_dict(top_yaml_dict):
        if "VariableDict" in top_yaml_dict:
            top_yaml_dict = top_yaml_dict["VariableDict"]
            variable_dict = dict()
            if top_yaml_dict is not None:
                for var, value in top_yaml_dict.items():
                    match = re.match(r'(\d+) to (\d+)', value)
                    if match:
                        start = int(match.group(1))
                        end = int(match.group(2))
                        val_list = [i for i in range(start, end+1)]
                        variable_dict[var] = val_list
            return variable_dict
        return None

    @staticmethod
    def get_yaml_root_dict_from_filepath(file_path):
        with open(file_path, 'r', encoding="utf-8") as stream:
            yaml_root_dict = yaml.safe_load(stream)
        return yaml_root_dict

    @staticmethod
    def add_indent_before_text(yaml_text, indent):
        line_list = ['    ' * indent + line + '\n' for line in yaml_text.split("\n")]
        return ''.join(line_list)

    @staticmethod
    def parse_records_to_yaml(header, records, save_file_name, root="Report"):
        text = root + "\n"
        for idx, rc in enumerate(records):
            key = "R" + "0" * (6 - len("%s" % idx)) + "%d" % (idx + 1)
            text += YamlML.parse_text_to_yaml("", key, 1, True)
            for col, col_name in enumerate(header):
                if col_name != "" and rc[col] != "":
                    text += YamlML.parse_text_to_yaml(rc[col], col_name, 2)
        with open(save_file_name, 'w+', encoding="utf-8") as f:
            f.write(text)
        return text

    @staticmethod
    def check_unaccept_line(input_dir, file_name):
        files = Folder.get_all_files(input_dir, file_name)
        unaccept_lines = []
        if len(files) >= 1:
            with open(files[0], 'r', encoding="utf-8") as stream:
                lines = stream.readlines()
                NON_PRINTABLE = re.compile('[^\x09\x0A\x0D\x20-\x7E\x85\xA0-\uD7FF\uE000-\uFFFD\U00010000-\U0010ffff]')
                for idx, line in enumerate(lines):
                    match = NON_PRINTABLE.search(line)
                    if match:
                        print (line.encode("utf-8"))
                        unaccept_lines.append("LINE %s: %s" % (idx, line))
        save_file_name = "REPORT.txt"
        with open(save_file_name, 'w+', encoding="utf-8") as f:
            f.write("\n".join(unaccept_lines))
