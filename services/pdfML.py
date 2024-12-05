# encoding: utf-8
import os
import re
import time
from datetime import datetime
from settings.config import Config
from services.folder import Folder
from PyPDF2 import PdfFileWriter, PdfFileReader, PdfFileMerger
from fpdf import FPDF

from collections import OrderedDict as o_dict


class PdfML(FPDF):
    """
    PdfML class
    """
    result_dict = dict()
    result_list = list()


    @staticmethod
    def merge_pdf_files(input_pdfs, output_pdf):
        """
        Combine multiple pdfs to single pdf.
        :param input_pdfs:
        :param output_pdf:
        :return:
        """
        pdf_merger = PdfFileMerger()
        for path in input_pdfs:
            pdf_merger.append(path, import_bookmarks=True)
        try:
            with open(output_pdf, 'wb') as fileobj:
                pdf_merger.write(fileobj)
        except:
            print("        ERROR: services.pdfML/PdfML..merge_pdf_files")
            print("        It could be Unicode Encode Error (latin-1 instead of utf-8)")
            print(r"       AppData\Local\Programs\Python\Python39\lib\site-packages\PyPDF2\utils.py")
            pass

    @staticmethod
    def extract_information(pdf_path):
        with open(pdf_path, 'rb') as f:
            pdf = PdfFileReader(f)
            information = pdf.getDocumentInfo()
            number_of_pages = pdf.getNumPages()
        try:
            created_date = information["/CreationDate"] # D:20210820195312+0700
            match = re.search(r"D:(.*)\+.*", created_date)
            created_date = datetime.strptime(match.group(1), "%Y%m%d%H%M%S") # "%Y-%b-%d %I:%M:%S %p"
            created_date = created_date.strftime("%Y-%b-%d %I:%M:%S %p")
        except:
            created_date = ""
            pass
        try:
            modified_date = information["/ModDate"]
            match1 = re.search(r"D:(.*)\+.*", modified_date)
            modified_date = datetime.strptime(match1.group(1), "%Y%m%d%H%M%S")
            modified_date = modified_date.strftime("%Y-%b-%d %I:%M:%S %p")
        except:
            modified_date = ""
            pass
        txt = f"""
        Information about {pdf_path}:    
        Author: {information.author}
        Creator: {information.creator}
        Producer: {information.producer}
        Subject: {information.subject}
        Title: {information.title}
        Number of pages: {number_of_pages}
        Created : {created_date}
        Last Modified : {modified_date}
        """
        result = {
            "about": pdf_path,
            "Author": information.author,
            "Creator": information.creator,
            "Producer": information.producer,
            "Subject": information.subject,
            "Title": information.title,
            "Number of pages": number_of_pages,
            "Created": created_date,
            "Last Modified": modified_date,
        }
        return result

    @staticmethod
    def extract_text(pdf_path, output_txt_file=None):
        text = ""
        with open(pdf_path, 'rb') as f:
            pdf = PdfFileReader(f)
            number_of_pages = pdf.getNumPages()
            for i in range(number_of_pages):
                page = pdf.getPage(i)
                text += "====Page %s===\n" % i
                text += page.extractText()
        if output_txt_file is None:
            out = open(pdf_path[:-4] + "_textextract.txt", "w", encoding="utf-8")
        else:
            out = open(output_txt_file, "w", encoding="utf-8")
        out.write(text)
        out.close()

    @staticmethod
    def split(path, name_of_split):
        pdf = PdfFileReader(path)
        for page in range(pdf.getNumPages()):
            pdf_writer = PdfFileWriter()
            pdf_writer.add_page(pdf.getPage(page))

            output = f'{name_of_split}{page}.pdf'
            with open(output, 'wb') as output_pdf:
                pdf_writer.write(output_pdf)

    @staticmethod
    def create_watermark(input_pdf, output, watermark):
        watermark_obj = PdfFileReader(watermark)
        watermark_page = watermark_obj.getPage(0)

        pdf_reader = PdfFileReader(input_pdf)
        pdf_writer = PdfFileWriter()

        # Watermark all the pages
        for page in range(pdf_reader.getNumPages()):
            page = pdf_reader.getPage(page)
            page.mergePage(watermark_page)
            pdf_writer.add_page(page)

        with open(output, 'wb') as out:
            pdf_writer.write(out)

    @staticmethod
    def add_encryption(input_pdf, output_pdf, password):
        pdf_writer = PdfFileWriter()
        pdf_reader = PdfFileReader(input_pdf)

        for page in range(pdf_reader.getNumPages()):
            pdf_writer.add_page(pdf_reader.getPage(page))

        pdf_writer.encrypt(user_pwd=password, owner_pwd=None,
                           use_128bit=True)

        with open(output_pdf, 'wb') as fh:
            pdf_writer.write(fh)

    @staticmethod
    def read_pdf_pages(input_pdf):
        pdfFileObject = open(input_pdf, 'rb')
        pdfReader = PdfFileReader(pdfFileObject)
        count = pdfReader.numPages
        for i in range(count):
            page = pdfReader.getPage(i)
            print(page.extractText())

    @staticmethod
    def get_bookmarks(filepath):
        PdfML.result_dict = dict()
        PdfML.result_list = list()
        reader = PdfFileReader(filepath)
        lvl = 0
        PdfML.__get_bookmarks(reader, lvl, reader.getOutlines())
        return PdfML.result_list

    @staticmethod
    def __get_bookmarks(reader, lvl, bookmark_list):
        for item in bookmark_list:
            if isinstance(item, list):
                # recursive call
                PdfML.result_dict.update(PdfML.__get_bookmarks(reader, lvl + 1, item))
            else:
                page = item.page if isinstance(item.page, int) else reader.getDestinationPageNumber(item)
                try:
                    PdfML.result_list.append([lvl, item.title, page, item.typ, item.left, item.top])
                except:
                    print ("        ERRROR: __get_bookmarks")
                    pass
        return PdfML.result_dict

    @staticmethod
    def bookmarks2txt(filepath, output_txt_file=None):
        """
        Read all bookmarks from a PDF document and create a dictionary with PageNumber and Title of the bookmark
        :param filepath:
        :return:
        """
        # WARNING! One page can have multiple bookmarks!
        bookmarks_txt_list = []
        toc = PdfML.get_bookmarks(filepath)
        for lvl, title, page, tp, left, top in toc:
            text_i = "%s\t%s" % ("\t"*lvl, title)
            bookmarks_txt_list.append(text_i)
        text = "\n".join(bookmarks_txt_list)
        if output_txt_file is None:
            out = open(filepath[:-4] + "_bookmarks.txt", "w", encoding="utf-8")
        else:
            out = open(output_txt_file, "w", encoding="utf-8")
        out.write(text)
        out.close()


    @staticmethod
    def bookmarks2txt_2(filepath, output_txt_file=None, level_max=0, exclude_list=[]):
        """
        Read all bookmarks from a PDF document and create a dictionary with PageNumber and Title of the bookmark
        :param filepath:
        :return:
        """
        # WARNING! One page can have multiple bookmarks!
        bookmarks_txt_list = []
        toc = PdfML.get_bookmarks(filepath)
        for lvl, title, page, tp, left, top in toc:
            match_list = [True if exclude in title else False for exclude in exclude_list]
            if match_list.count(True) > 0:
                continue
            if lvl <= level_max:
                text_i = "%s\t%s\t%s" % ("\t"*lvl, title, page)
                bookmarks_txt_list.append(text_i)
        text = "\n".join(bookmarks_txt_list)
        if output_txt_file is None:
            out = open(filepath[:-4] + "_bookmarks.txt", "w", encoding="utf-8")
        else:
            out = open(output_txt_file, "w", encoding="utf-8")
        out.write(text)
        out.close()

    @staticmethod
    def demote_bookmark(filepath, bm_list=["154.1", "154.2", "154.3"], output=""):
        pdf_writer = PdfFileWriter()
        pdf_reader = PdfFileReader(filepath)
        pdf_writer.appendPagesFromReader(pdf_reader)
        toc = PdfML.get_bookmarks(filepath)
        denote_bm_list = []
        if bm_list is not None:
            for bm in bm_list:
                found = False
                for idx, bm_i in enumerate(toc):
                    if bm_i[1].startswith(bm):
                        found = True
                        bm_start = idx
                        bm_end = idx
                        for j in range(idx+1, len(toc)):
                            if toc[j][0] > toc[bm_start][0] or (toc[j][0] == toc[bm_start][0] and toc[j][1].startswith(bm)):
                                bm_end = j
                            else:
                                break
                        denote_bm_list.extend(range(bm_start, bm_end+1))
                    if found:
                        break
        for idx in denote_bm_list:
            toc[idx][0] += 1
        perents = [None]
        oldperent = None
        nowLayer = 0
        for index in range(len(toc)):
            line = toc[index]
            count = line[0]
            if count > nowLayer:
                perents.append(oldperent)
                nowLayer += 1
            elif count < nowLayer:
                while(nowLayer > count):
                    perents.pop()
                    nowLayer -= 1
            elif count == 0:
                perents = [None]
                nowLayer = 0
            try:
                oldperent = pdf_writer.addBookmark(line[1], line[2], perents[-1], None, False, False,
                                                   line[3], line[4], line[5], 0)
            except:
                pass
        # Write out the merged PDF
        with open(output, 'wb') as out:
            pdf_writer.write(out)

    @staticmethod
    def modify_bookmark(filepath, bm_list_dict=None, output="", IgnoreBMList=["Section", "158. Control Domain Parts"]):
        pdf_writer = PdfFileWriter()
        pdf_reader = PdfFileReader(filepath)
        pdf_writer.appendPagesFromReader(pdf_reader)
        toc = PdfML.get_bookmarks(filepath)
        MAX = 3 if len(toc) > 3 else len(toc)
        perents = [None]
        oldperent = None
        insert = True
        for ignore in IgnoreBMList:
            regex = r"%s" % ignore
            if len(toc) > 0:
                bm = toc[0][1]
                match = re.match(regex, bm)
                if match:
                    insert = False
        if insert and bm_list_dict is not None:
            bm_list = [key for key, val in bm_list_dict.items()]
            for bm_i in bm_list:
                found = False
                for idx in range(MAX):
                    if toc[idx][1].startswith(bm_i):
                        found = True
                if found:
                    for line in bm_list_dict[bm_i]:
                        line[2] = toc[0][2]
                        line[3] = toc[0][3]
                        line[4] = toc[0][4]  # X
                        line[5] = toc[0][5] + line[6]  # Y
                        oldperent = pdf_writer.addBookmark(line[1], line[2], perents[-1], None, False, False,
                                                      line[3], line[4], line[5], 0)
                        perents.append(oldperent)
        if perents[-1] is not None:
            for idx, bm_i in enumerate(toc):
                toc[idx][0] += 1
        nowLayer = 0
        for index in range(len(toc)):
            line = toc[index]
            count = line[0]
            if count > nowLayer:
                perents.append(oldperent)
                nowLayer += 1
            elif count < nowLayer:
                while nowLayer > count:
                    perents.pop()
                    nowLayer -= 1
            elif count == 0:
                perents = [None]
                nowLayer = 0
            try:
                oldperent = pdf_writer.addBookmark(line[1], line[2], perents[-1], None, False, False,
                                                   line[3], line[4], line[5], 0)
            except:
                pass
        # Write out the merged PDF
        with open(output, 'wb') as out:
            pdf_writer.write(out)
        return True

    @staticmethod
    def modify_bookmark2(filepath, bm_list_dict=None, output="", IgnoreBMList=["Section", "158. Control Domain Parts"],
                         bm_list=["154.1", "154.2", "154.3"], DeleteMLList=None, BookmarkLevel=None):
        pdf_writer = PdfFileWriter()
        pdf_reader = PdfFileReader(filepath)
        pdf_writer.appendPagesFromReader(pdf_reader)
        toc = PdfML.get_bookmarks(filepath)
        MAX = 3 if len(toc) > 3 else len(toc)
        perents = [None]
        oldperent = None
        insert = True
        for ignore in IgnoreBMList:
            regex = r"%s" % ignore
            if len(toc) > 0:
                bm = toc[0][1]
                match = re.match(regex, bm)
                if match:
                    insert = False
        if insert and bm_list_dict is not None:
            bm_list = [key for key, val in bm_list_dict.items()]
            for bm_i in bm_list:
                found = False
                for idx in range(MAX):
                    if toc[idx][1].startswith(bm_i):
                        found = True
                if found:
                    for line in bm_list_dict[bm_i]:
                        line[2] = toc[0][2]
                        line[3] = toc[0][3]
                        line[4] = toc[0][4]  # X
                        line[5] = toc[0][5] + line[6]  # Y
                        oldperent = pdf_writer.addBookmark(line[1], line[2], perents[-1], None, False, False,
                                                      line[3], line[4], line[5], 0)
                        perents.append(oldperent)
        if perents[-1] is not None:
            for idx, bm_i in enumerate(toc):
                toc[idx][0] += 1
        denote_bm_list = []
        if bm_list is not None:
            for bm in bm_list:
                found = False
                for idx, bm_i in enumerate(toc):
                    if bm_i[1].startswith(bm):
                        found = True
                        bm_start = idx
                        bm_end = idx
                        for j in range(idx+1, len(toc)):
                            if toc[j][0] > toc[bm_start][0] or (toc[j][0] == toc[bm_start][0] and toc[j][1].startswith(bm)):
                                bm_end = j
                            else:
                                break
                        denote_bm_list.extend(range(bm_start, bm_end+1))
                    if found:
                        break
        for idx in denote_bm_list:
            toc[idx][0] += 1
        nowLayer = 0
        for index in range(len(toc)):
            line = toc[index]
            count = line[0]
            if count > nowLayer:
                perents.append(oldperent)
                nowLayer += 1
            elif count < nowLayer:
                while nowLayer > count:
                    perents.pop()
                    nowLayer -= 1
            elif count == 0:
                perents = [None]
                nowLayer = 0
            # delete bookmark by level
            if BookmarkLevel is not None:
                if line[0] not in BookmarkLevel:
                    continue
            # delete bookmark by text
            found = False
            for li in DeleteMLList:
                regex = "%s" % li
                match = re.match(r'%s' % regex, line[1])
                if match:
                    found = True
            if found:
                continue
            try:
                oldperent = pdf_writer.addBookmark(line[1], line[2], perents[-1], None, False, False,
                                                       line[3], line[4], line[5], 0)
            except:
                pass
        # Write out the merged PDF
        with open(output, 'wb') as out:
            pdf_writer.write(out)
        return True

    @staticmethod
    def modify_bookmark3(filepath, bm_list=None, output=""):
        pdf_writer = PdfFileWriter()
        pdf_reader = PdfFileReader(filepath)
        pdf_writer.appendPagesFromReader(pdf_reader)
        perents = [None]
        if bm_list is not None:
            for line in bm_list:
                oldperent = pdf_writer.addBookmark(line[1], line[2], perents[-1], None, False, False,
                                              line[3], line[4], -1, 0)
                perents.append(oldperent)
        # Write out the merged PDF
        with open(output, 'wb') as out:
            pdf_writer.write(out)
        return True

    @staticmethod
    def import_bookmark(filepath, bm_list=None, output=""):
        pdf_writer = PdfFileWriter()
        pdf_reader = PdfFileReader(filepath)
        pdf_writer.appendPagesFromReader(pdf_reader)
        perents = [None]
        toc = bm_list
        oldperent = None
        if bm_list is not None:
            # if perents[-1] is not None:
            #     for idx, bm_i in enumerate(toc):
            #         toc[idx][0] += 1
            nowLayer = 0
            for index in range(len(toc)):
                line = toc[index]
                count = line[0]
                if count > nowLayer:
                    perents.append(oldperent)
                    nowLayer += 1
                elif count < nowLayer:
                    while nowLayer > count:
                        perents.pop()
                        nowLayer -= 1
                elif count == 0:
                    perents = [None]
                    nowLayer = 0
                try:
                    oldperent = pdf_writer.addBookmark(line[1], line[2]-1, perents[-1], None, False, False,
                                                       line[3], line[4], line[5], 0)
                except:
                    pass
        with open(output, 'wb') as out:
            pdf_writer.write(out)
        return True

    @staticmethod
    def bookmarks2toc(filepath, output_pdf_file=None, total_page=1, level_list=[0], exclude_list=[],
                      replace_list=[], total_item_per_page=23,
                      start_new_page_list = [], extra_condition = {1: "Section [\d]+"},
                      default_heading_bm=[]):
        """
        :param filepath:
        :return:
        """
        has_default_heading_bm = True if default_heading_bm is not None else False
        bookmarks_0_tmp_list = []
        bookmarks_check_missing_tmp_list = []
        bookmarks_txt_list = []
        toc = PdfML.get_bookmarks(filepath)
        for lvl, title, page, tp, left, top in toc:
            match_list = [True if exclude in title else False for exclude in exclude_list]
            if match_list.count(True) > 0:
                continue
            if lvl in level_list:
                for [key, val] in replace_list:
                    if key in title:
                        title = title.replace(key, val)
                        break
                bookmarks_0_tmp_list.append([lvl, title, page])
        # if has_default_heading_bm is True:
        #     for default_num, default_title, default_page in default_heading_bm:
        #         for lvl, title, page in bookmarks_0_tmp_list:
        #             if default_title in title:
        #                 bookmarks_check_missing_tmp_list.append([lvl, title, page])
        #                 break
        #             else:
        #                 bookmarks_check_missing_tmp_list.append([0, default_num+" "+default_title, default_page])
        #                 break
        # else:
        #     bookmarks_check_missing_tmp_list = bookmarks_0_tmp_list
        # bookmarks_txt_list.extend(bookmarks_check_missing_tmp_list)
        bookmarks_txt_list.extend(bookmarks_0_tmp_list)
        # else:
        #     if extra_condition is not None:
        #         for k, v in extra_condition.items():
        #             if lvl == k and re.match(r'%s'% v, title):
        #                 bookmarks_txt_list.append([lvl, title, page])
        if output_pdf_file is None:
            output_pdf_file = filepath[:-4] + "_bookmarks2toc.pdf"
        top = 28
        left = 19
        right = 19
        bottom = 20
        height = 297
        width = 210
        pdf = FPDFML()
        pdf.add_page()
        toc_text = "Table of Contents"
        pdf.set_font('Arial', 'B', 20)
        pdf.set_text_color(0, 0, 0)
        w = pdf.get_string_width(toc_text) + 4
        pdf.set_xy((210 - w) / 2, 15)
        pdf.cell(w, 9, txt=toc_text, border=0, align='C', fill=False, link='')
        pdf.set_font('Arial', '', 14)
        pdf.set_text_color(0, 0, 0)
        pdf.set_y(30)
        w_tab = pdf.get_string_width(".")
        for i in range(len(bookmarks_txt_list)):
            pdf.set_x(left)
            [lvl, title, page] = bookmarks_txt_list[i]
            match_list = [True if start_new_page in title else False for start_new_page in start_new_page_list]
            if match_list.count(True) > 0:
                pdf.cell(40, 20, str(""), 0, 1)
                pdf.set_x(left)
            match = re.match(r'(.*\s[0-9\.]+)(\s)(.*)', title)
            indent_txt = "".join("    "*(lvl+1))
            if match:
                title_i = "%s%s%s" % (match.group(1).strip(), indent_txt, match.group(3).strip())
            else:
                title_i = "%s" % title
            w = pdf.get_string_width(title_i) + 4
            page_txt = "%s" % (page + 1 + total_page)
            page_num = page + 1 + total_page
            w_page = pdf.get_string_width(page_txt)
            tab_txt = "".join("." * round((210 - w - w_page-left-right-4) / w_tab))
            # text_i = title_i + tab_txt + page_txt
            text_i = title_i
            to_page = pdf.add_link()
            # pdf.set_link(to_page, page=page_num)
            pdf.set_link(to_page, page=1)
            pdf.cell(40, 10, str(text_i), 0, 1, link=to_page)
            # pdf.cell(40, 10, str(text_i), 0, 1)
        pdf.output(output_pdf_file, 'F')

    @staticmethod
    def bookmarks2toc_v2(filepath, output_pdf_file=None, level_list=[0], exclude_list=[],
                      replace_list=[]):
        bookmarks_txt_list = []
        toc = PdfML.get_bookmarks(filepath)
        for lvl, title, page, tp, left, top in toc:
            match_list = [True if exclude in title else False for exclude in exclude_list]
            if match_list.count(True) > 0:
                continue
            if lvl in level_list:
                for [key, val] in replace_list:
                    if key in title:
                        title = title.replace(key, val)
                        break
                bookmarks_txt_list.append([lvl, title, page])

        if output_pdf_file is None:
            output_pdf_file = filepath[:-4] + "_bookmarks2toc.pdf"

        pdf_toc = TOC()
        for lvl, title, page in bookmarks_txt_list:
            pdf_toc.TOC_Entry(level=lvl, txt=title, page=page)
        pdf_toc.insertTOC(location=4)
        pdf_toc.output(output_pdf_file, 'F')

    @staticmethod
    def create_pdf_toc(output_pdf_file='',total_page=3, start_new_page_list=[],replace_list=[], heading_toc=[], page_toc=[]):
        bookmarks_txt_list = []
        for title, page in zip(heading_toc, page_toc):
            for [key, val] in replace_list:
                if key in title:
                    title = title.replace(key, val)
                    break
            bookmarks_txt_list.append([0, title, page])

        top = 28
        left = 19
        right = 19
        bottom = 20
        height = 297
        width = 210
        pdf = FPDFML()
        pdf.core_fonts_encoding = "utf-8"
        pdf.add_page()
        toc_text = "Table of Contents"
        pdf.set_font('Arial', 'B', 20)
        pdf.set_text_color(0, 0, 0)
        w = pdf.get_string_width(toc_text) + 4
        pdf.set_xy((210 - w) / 2, 15)
        pdf.cell(w, 9, txt=toc_text, border=0, align='C', fill=False, link='')
        pdf.set_font('Arial', '', 14)
        pdf.set_text_color(0, 0, 0)
        pdf.set_y(30)
        w_tab = pdf.get_string_width(".")
        for i in range(len(bookmarks_txt_list)):
            pdf.set_x(left)
            [lvl, title, page] = bookmarks_txt_list[i]
            match_list = [True if start_new_page in title else False for start_new_page in start_new_page_list]
            if match_list.count(True) > 0:
                pdf.cell(40, 20, str(""), 0, 1)
                pdf.set_x(left)
            match = re.match(r'(.*\s[0-9\.]+)(\s)(.*)', title)
            indent_txt = "".join("    " * (lvl + 1))
            if match:
                title_i = "%s%s%s" % (match.group(1).strip(), indent_txt, match.group(3).strip())
            else:
                title_i = "%s" % title
            w = pdf.get_string_width(title_i) + 4
            page_txt = "%s" % (page + total_page)
            page_num = page + total_page
            w_page = pdf.get_string_width(page_txt)
            tab_txt = "".join("." * round((210 - w - w_page - left - right - 4) / w_tab))
            # text_i = title_i + tab_txt + page_txt
            text_i = title_i
            to_page = pdf.add_link()
            # pdf.set_link(to_page, page=page_num)
            pdf.set_link(to_page, page=1)
            pdf.cell(40, 10, str(text_i), 0, 1, link=to_page)
            # pdf.cell(40, 10, str(text_i), 0, 1)
        pdf.output(output_pdf_file, 'F')

    @staticmethod
    def delete_blank_pdf_pages(input, output):
        reader = PdfFileReader(open(input, 'rb'))
        writer = PdfFileWriter()
        pages = reader.getNumPages()

        for i in range(pages):
            page = reader.getPage(i)
            if "/XObject" in page["/Resources"].keys() or "/Font" in page["/Resources"].keys():
                writer.add_page(page)
        with open(output, 'wb') as out:
            writer.write(out)

    @staticmethod
    def get_number_page(input_pdf):
        reader = PdfFileReader(open(input_pdf, 'rb'))
        pages = reader.getNumPages()
        return pages

    @staticmethod
    def create_pdf_toc2(bookmarks_txt_list, output_pdf_file='', toc_split_list=[24, 48, 72, 98]):
        left = 19
        right = 19
        pdf = FPDFML()
        pdf.core_fonts_encoding = "utf-8"
        pdf.add_page()
        toc_text = "Table of Contents"
        pdf.set_font('Arial', 'B', 20)
        pdf.set_text_color(0, 0, 0)
        w = pdf.get_string_width(toc_text) + 4
        pdf.set_xy((210 - w) / 2, 15)
        pdf.cell(w, 9, txt=toc_text, border=0, align='C', fill=False, link='')
        pdf.set_font('Arial', '', 14)
        pdf.set_text_color(0, 0, 0)
        pdf.set_y(30)
        w_tab = pdf.get_string_width(".")
        page_count = 1
        for i in range(len(bookmarks_txt_list)):
            pdf.set_x(left)
            [lvl, title, page] = bookmarks_txt_list[i]
            if i in toc_split_list:
                # pdf.cell(40, 20, str(""), 0, 1)
                pdf.set_x(left)
                page_count += 1
            match = re.match(r'(.*\s[0-9\.]+)(\s)(.*)', title)
            indent_txt = "".join("    " * (lvl + 1))
            if match:
                title_i = "%s%s%s" % (match.group(1).strip(), indent_txt, match.group(3).strip())
            else:
                title_i = "%s" % title
            w = pdf.get_string_width(title_i) + 4
            page_txt = "%s" % (page)
            page_num = page
            w_page = pdf.get_string_width(page_txt)
            tab_txt = "".join("." * round((210 - w - w_page - left - right - 4) / w_tab))
            # text_i = title_i + tab_txt + page_txt
            text_i = title_i
            to_page = pdf.add_link()
            # pdf.set_link(to_page, page=page_num)
            pdf.set_link(to_page, page=1)
            pdf.cell(40, 10, str(text_i), 0, 1, link=to_page)
            # pdf.cell(40, 10, str(text_i), 0, 1)
        tmp_file = output_pdf_file + "0"
        pdf.output(tmp_file, 'F')
        PdfML.add_bookmark_simple(tmp_file, output_pdf_file, toc_text)

    @staticmethod
    def add_bookmark_simple(doc_input, output_pdf_file, toc_text):
        pdf_writer = PdfFileWriter()
        pdf_reader = PdfFileReader(doc_input)
        pdf_writer.appendPagesFromReader(pdf_reader)
        if isinstance(toc_text, str):
            pdf_writer.addBookmark(toc_text, 0, None, None, False, False,
                                   '/XYZ', 0, 0, 1)  # add bookmark
        elif isinstance(toc_text, list):
            for toc_text_i in toc_text:
                pdf_writer.addBookmark(toc_text_i[0], toc_text_i[1], None, None, False, False,
                                       '/XYZ', 0, 0, 1)  # add bookmark
        with open(output_pdf_file, 'wb') as out:
            pdf_writer.write(out)

    @staticmethod
    def import_bookmark_from_record(input_pdf, output_pdf, records):
        pdf_writer = PdfFileWriter()
        pdf_reader = PdfFileReader(input_pdf)
        pdf_writer.appendPagesFromReader(pdf_reader)

        for records_i in records:
            pdf_writer.addBookmark(records_i[1], records_i[2], None, None, False, False,
                                   '/XYZ', 0, 0, 1)  # add bookmark
        with open(output_pdf, 'wb') as out:
            pdf_writer.write(out)
        return
class PdfHeaderFooterML(FPDF):
    def header(self):
        top = 28
        left = 19
        right = 19
        bottom = 20
        # Header Separator
        self.set_draw_color(0, 9, 153)
        self.set_line_width(0.5)
        self.line(left, top-5, 210-right, top-5)
        # Under Development
        self.set_font('Arial', 'BI', 11)
        self.set_xy(left, 9)
        self.cell(40, 10, 'Under Development')
        # R-Car S4
        self.set_xy(left, top - 13)
        self.set_font('Arial', '', 10)
        self.cell(40, 10, 'R-Car S4')
        # Chapter title
        w = self.get_string_width("7. Pin Function (PFC/GPIO)") + 4
        self.set_y(15)
        self.set_x(210-w - left )
        self.cell(w, 9, txt="7. Pin Function (PFC/GPIO)", border=0, align='R', fill=False, link='')
        # # Century bold 12
        self.add_font('sysfont', '', r"c:\WINDOWS\Fonts\Century.ttf", uni=True)
        self.set_font('Arial', 'B', 12)
        # Calculate width of title and position
        w = self.get_string_width("RENESAS CONFIDENTIAL") + 4
        self.set_xy((210 - w) / 2, 10)
        # # Colors of frame, background and text
        self.set_draw_color(0, 0, 0)
        self.set_fill_color(255, 255, 255)
        # # Thickness of frame (1 mm)
        self.set_line_width(0.1)
        # # Title
        self.cell(w, 9, "RENESAS CONFIDENTIAL", 1, 1, 'C', 1)
        # # Line break
        # self.ln(10)

    def footer(self):
        top = 28
        left = 19
        right = 19
        bottom = 20
        height = 297
        width = 210
        # Header Separator
        self.set_y(height - bottom)
        self.set_draw_color(0, 9, 153)
        self.set_line_width(0.5)
        self.line(left, height - bottom, 210-right, height - bottom)
        # Rev 0.4
        self.set_xy(left, height - bottom)
        self.set_font('Arial', 'B', 10)
        self.cell(40, 10, 'Rev 0.4')
        self.set_xy(left, height - bottom + 10)
        self.set_font('Arial', '', 10)
        self.cell(40, 0, 'Apr. 19, 2021')
        # Page number
        w = self.get_string_width('Page %s of Section 7' % str(self.page_no())) + 4
        self.set_xy(210-w - left, height - bottom)
        self.cell(w, 9, txt='Page %s of Section 7' % str(self.page_no()), border=0, align='R', fill=False, link='')
        self.image('logo_renesas_mcu.png', 210/2 - 33/2,  height - bottom + 1, 33)

    def chapter_body(self, name):
        # Read text file
        with open(name, 'rb') as fh:
            txt = fh.read().decode('latin-1')
        # Times 12
        self.set_font('Times', '', 12)
        # Output justified text
        self.multi_cell(0, 5, txt)
        # Line break
        self.ln()
        # Mention in italics
        self.set_font('', 'I')
        self.cell(0, 5, '(end of excerpt)')

    def print_chapter(self, num, title, name):
        self.add_page()
        self.chapter_title(num, title)
        self.chapter_body(name)

    @staticmethod
    def merge_header_footer():
        bg_filename = 'Client_Name_Header_Footer.pdf'
        fg_filename = 'Client_Report.pdf'
        out_filename = 'Merged.pdf'

        with open(bg_filename, 'rb') as bg_file, open(fg_filename, 'rb') as fg_file:
            bg_page = PdfFileReader(bg_file).getPage(0)
            pdf_out = PdfFileWriter()
            for page in PdfFileReader(fg_file).pages:
                if page.extractText():  # Do not copy pages that have no text
                    page.mergePage(bg_page)
                    pdf_out.add_page(page)
            if pdf_out.getNumPages():
                with open(out_filename, 'wb') as out_file:
                    # Caution: All three files MUST be open when write() is called
                    pdf_out.write(out_file)


def create_dictionary_string(dict_):
    """format ordered dictionary as PDF dictionary"""
    return ''.join([
        '<<',
        '\n'.join([' '.join(f) for f in dict_.items()]),
        ' >>'
    ])


def create_list_string(list_):
    """format list of strings as PDF array"""
    return '[' + ' '.join(list_) + ']'


def iobj_ref(n):
    """format an indirect PDF Object reference from its id number"""
    return str(n) + ' 0 R'


class FPDFML(FPDF):
    def __init__(self, orientation='P', unit='mm', format='A4'):
        super(FPDFML, self).__init__(orientation, unit, format)
        self.outlines = []
        self.outline_root_obj_reference = None
        self.current_page_number = 0

    def bookmark(self, txt, level=0, y=0):
        if y == -1:
            y = self.get_y()
        self.outlines.append({
            't': txt, 'level': level, 'y': y,
            'p': self.current_page_number
        })

    def add_page(self, *a, **k):
        if not self.current_page_number:
            self.current_page_number = 1
        else:
            self.current_page_number += 1
        super(FPDFML, self).add_page(*a, **k)

    def _putbookmarks(self):
        """Print Bookmark tags into the resources section.

        If the accumulated list of bookmarks is empty, self is a no-op.

        First, some assembly required in order to organize a flat list into a
        tree. Then, after all nested indirect references are in place, make and
        dictionary for each bookmark, and make and insert the dictionary for
        the document outline.
        """
        if not self.outlines:
            return
        lru = {}
        last_level = 0
        for index, outline_el in enumerate(self.outlines):
            if outline_el['level'] > 0:
                # Set parent and last pointers
                parent = lru[outline_el['level'] - 1]
                outline_el['parent'] = parent
                self.outlines[parent]['last'] = index

                if outline_el['level'] > last_level:
                    # Level increasing: set first pointer
                    self.outlines[parent]['first'] = index
            else:
                outline_el['parent'] = len(self.outlines)
            if outline_el['level'] <= last_level and index > 0:
                # Set prev and next pointers
                prev = lru[outline_el['level']]
                self.outlines[prev]['next'] = index

                outline_el['prev'] = prev

            lru[outline_el['level']] = index
            last_level = outline_el['level']
        # Outline items
        n = self.n + 1

        for index, outline_el in enumerate(self.outlines):
            elem_fields = o_dict()
            elem_fields['/Title'] = (outline_el['t'])
            elem_fields['/Parent'] = iobj_ref(n + outline_el['parent'])

            for opt in ['prev', 'next', 'first', 'last']:
                if opt in outline_el:
                    tag = '/' + opt.capitalize()
                    elem_fields[tag] = iobj_ref(n + outline_el[opt])

            # page 582 PDF Reference 1.7
            elem_fields['/Dest'] = create_list_string([
                # page object reference
                iobj_ref(1 + 2 * outline_el['p']),
                '/XYZ',  # vs /Fit, /FitH, /FitV, /FitR, etc...
                # left, top
                '0', '%.2f' % ((self.h - outline_el['y']) * self.k),
                'null'  # zoom
            ])
            elem_fields['/Count'] = '0'

            self._newobj()
            self._out(create_dictionary_string(elem_fields))
            self._out('endobj')

        # Outline object (page 585 PDF Reference 1.7)
        outlines_dictionary = o_dict()
        outlines_dictionary['/Type'] = '/Outlines'
        outlines_dictionary['/First'] = iobj_ref(n)
        outlines_dictionary['/Last'] = iobj_ref(n + lru[0])

        self._newobj()
        outlines_string = create_dictionary_string(outlines_dictionary)

        # writing three lines of deletable code is easier than windows &
        # re-generating test hashes.
        outlines_mutable_string = bytearray(outlines_string, 'ascii')
        outlines_mutable_string[17] = ord(' ')
        outlines_string = outlines_mutable_string.decode('ascii')

        self._out(outlines_string)
        self._out('endobj')

        # Saved for Catalog Dictionary
        self.outline_root_obj_reference = self.n

    def _putresources(self):
        super(FPDFML, self)._putresources()
        self._putbookmarks()

    def _putcatalog(self):
        super(FPDFML, self)._putcatalog()
        if self.outlines:
            self._out('/Outlines ' + iobj_ref(self.outline_root_obj_reference))
            self._out('/PageMode /UseOutlines')


class TOC(FPDF):

    def __init__(self, orientation='P',unit='mm',format='A4'):
        self._toc = []
        self._numbering = 0
        self._numberingFooter = 0
        self._numPageNum = 1
        FPDF.__init__(self, orientation,unit,format)

    def add_page(self, orientation=''):
        FPDF.add_page(self, orientation)
        if self._numbering:
            self._numPageNum += 1

    def start_page_nums(self):
        self._numbering = 1
        self._numberingFooter = 1

    def stop_page_nums(self):
        self._numbering = 0

    def num_page_no(self):
        return self._numPageNum

    def TOC_Entry(self, txt, level=0, page=0):
        self._toc += [{'t': txt, 'l': level, 'p': page}]

    @staticmethod
    def str_repeat(s, t):
        return s * int(t)

    def insertTOC(self, location=1, labelSize=20, entrySize=10, tocfont='Times', label='Table of Contents'):
        # make toc at end
        self.stop_page_nums()
        self.add_page()
        tocstart = self.page

        self.set_font(tocfont, 'B', labelSize)
        self.cell(0, 5, label, 0, 1, 'C')
        self.ln(10)

        for t in self._toc:
            # Offset
            level = t['l']
            if level>0:
                self.cell(level*8)
            weight = ''
            if level == 0:
                weight = 'B'
            str_txt = t['t']
            self.set_font(tocfont, weight, entrySize)
            strsize = self.get_string_width(str_txt)
            to_page = self.add_link()
            link = self.set_link(to_page, page=t['p'])
            self.cell(strsize+2, self.font_size+2, str_txt, link=link)

            # Filling dots
            self.set_font(tocfont, '', entrySize)
            page_cell_size=self.get_string_width(str(t['p']))+2
            w = self.w-self.l_margin-self.r_margin-page_cell_size-(level*8)-(strsize+2)
            nb = w/self.get_string_width('.')
            dots = self.str_repeat('.', nb)
            self.cell(w, self.font_size+2, dots, 0, 0, 'R')

            # Page number
            self.cell(page_cell_size, self.font_size+2, str(t['p']), 0, 1, 'R')

        # grab it and move to selected location
        n = self.page
        n_toc = n - tocstart + 1
        last = []

        # store toc pages
        for i in range(tocstart,n+1):
            last += [self.pages[i]]

        # move pages
        for i in range(tocstart-1, location-1, -1):
            # ~ for(i = tocstart - 1;i>=location-1;i--)
            self.pages[i+n_toc] = self.pages[i]

        # Put toc pages at insert point
        for i in range(0, n_toc):
            self.pages[location + i] = last[i]

    def footer(self):
        if self._numberingFooter == 0:
            return
        # Go to 1.5 cm from bottom
        self.set_y(-15)
        # Select Arial italic 8
        self.set_font('Arial', 'I', 8)
        self.cell(0, 7, str(self.num_page_no()), 0, 0, 'C')
        if self._numbering == 0:
            self._numberingFooter = 0


if __name__ == '__main__':
    config = Config()
    input_pdfs = Folder.get_all_files(config.input_dir, "*.pdf")
    # output_pdf = os.path.join(config.output_dir, "output.pdf")
    # output_denote = os.path.join(config.output_dir, "output_denote.pdf")
    # PdfML.merge_pdf_files(input_pdfs, output_pdf)
    # PdfML.demote_bookmark(output_pdf, "154.1", output_denote)

    PdfML.extract_information(input_pdfs[0])
