# encoding: utf-8
import re
import os
import copy
from copy import deepcopy
from settings.config import Config
import win32com.client as win32
from win32com.client import Dispatch
import shutil
from services.folder import Folder
import time
import sys
import psutil
from win32com.client import constants
from dataclasses import dataclass
from services.reporter import Reporter


def clean_gen_cache(func):
    def inner(*args, **kwargs):
        try:
            MODULE_LIST = [m.__name__ for m in sys.modules.values()]
            for module in MODULE_LIST:
                if re.match(r'win32com\.gen_py\..+', module):
                    del sys.modules[module]
            shutil.rmtree(os.path.join(os.environ.get('LOCALAPPDATA'), 'Temp', 'gen_py'))
        except Exception as e:
            for process in psutil.process_iter():
                print(process)
            pass
        return func(*args, **kwargs)
    return inner


def clean_gen_py(func):
    def inner(*args, **kwargs):
        try:
            MODULE_LIST = [m.__name__ for m in sys.modules.values()]
            for module in MODULE_LIST:
                if re.match(r'win32com\.gen_py\..+', module):
                    del sys.modules[module]
            shutil.rmtree(os.path.join(os.environ.get('LOCALAPPDATA'), 'Temp', 'gen_py'))
        except Exception as e:
            pass
        return func(*args, **kwargs)
    return inner


@dataclass
class MyConstants:
    # Name	Value	Description
    # wdNoRevision	0	No revision.
    wdRevisionCellDeletion = 17  # Table cell deleted.
    wdRevisionCellInsertion = 16  # Table cell inserted.
    wdRevisionCellMerge = 18  #  Table cells merged.
    wdRevisionCellSplit = 19  #  This object, member, or enumeration is deprecated and is not intended to be used in your code.
    # wdRevisionConflict	7	Revision marked as a conflict.
    # wdRevisionConflictDelete	21	Deletion revision conflict in a coauthored document.
    # wdRevisionConflictInsert	20	Insertion revision conflict in a coauthored document
    wdRevisionDelete = 2    # Deletion.
    # wdRevisionDisplayField	5	Field display changed.
    wdRevisionInsert = 1    # Insertion.
    wdRevisionMovedFrom	= 14	# Content moved from.
    # wdRevisionMovedTo	15	Content moved to.
    wdRevisionParagraphNumber =	4	# Paragraph number changed.
    # wdRevisionParagraphProperty	10	Paragraph property changed.
    # wdRevisionProperty	3	Property changed.
    # wdRevisionReconcile	6	Revision marked as reconciled conflict.
    wdRevisionReplace =	9           # Replaced.
    # wdRevisionSectionProperty	12	Section property changed.
    wdRevisionStyle =	8	#  Style changed.
    wdRevisionStyleDefinition =	13	  #  Style definition changed.
    # wdRevisionTableProperty	11	Table property changed.

    wdActiveEndPageNumber = 3
    wdRefTypeHeading = 1
    wdRefTypeBookmark = 2
    wdRefTypeNumberedItem = 0
    wdRefTypeEndnote = 4
    wdRefTypeFootnote = 3

    wdGoToAbsolute = 1
    wdGoToNext = 2
    wdGoToPrevious = 3
    wdGoToRelative = 2
    wdGoToLast = -1
    wdGoToFirst = 1

    wdGoToPage = 1
    wdGoToHeading = 11
    wdGoToLine = 3
    wdGoToTable = 2
    wdGoToBookmark = -1

    wdDoNotSaveChanges = 0
    wdPromptToSaveChanges = -2
    wdSaveChanges = -1

    wdCollapseStart = -1
    wdCollapseEnd = 0

    wdFindContinue = 1

    # Specifies a unit of measure to use.
    wdParagraph = 4

    wdNumberOfPagesInDocument = 4 # Returns the number of pages in the document associated with the selection or range.

    wdWithInTable = 12


class WinWordML:
    """WinWordML"""
    auto_worker_name = "WinWordML"

    def __init__(self, filename=None, package=None):
        self.auto_worker_name = self.__class__.__name__

    @staticmethod
    def get_number_page(doc_input):
        """
        get total number of doc input file
        :param doc_input:
        :return: total_page
        """
        # open Word
        word = Dispatch('Word.Application')
        word.Visible = False
        doc = word.Documents.Open(doc_input)
        # get number of sheets
        # doc.Repaginate()
        # num_of_sheets = doc.ComputeStatistics(2)
        num_of_sheets = doc.ActiveWindow.Panes(1).Pages.Count
        doc.Close(True)
        word.Quit()
        return num_of_sheets

    @staticmethod
    def get_blank_pages(doc_input, total_page):
        """
        In order to get the entire Range for a page it's possible to use a built-in bookmark name Page".
        This only works for the page where the selection is, which is why it's necessary to first go to the page.
        :param doc_input:
        :param total_page:
        :return:
        """
        blank_page_list = []
        word = win32.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(doc_input)
        doc.Repaginate()
        for i in range(1, total_page + 1):
            word.Selection.GoTo(MyConstants.wdGoToPage, MyConstants.wdGoToAbsolute, "%s" % i)
            rngPage = doc.Bookmarks(r"\Page").Range
            if (rngPage.Text.strip()) == "":
                blank_page_list.append(i)
        doc.Close(True)
        word.Quit()
        # print(", ".join([str(j) for j in blank_page_list]))
        return blank_page_list

    @staticmethod
    def del_blank_pages(doc_input, blank_page_list):
        """
        In order delete blank page
        :param doc_input:
        :param blank_page_list:
        :return:
        """
        word = win32.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(doc_input)
        doc.Repaginate()
        for i in blank_page_list:
            word.Selection.GoTo(MyConstants.wdGoToPage, MyConstants.wdGoToAbsolute, "%s" % i)
            word.Selection.Delete()
        doc.Close(MyConstants.wdSaveChanges)
        word.Quit()
        return blank_page_list

    @staticmethod
    def create_outline(doc_input):
        """
        In order delete blank page
        :param doc_input:
        :param blank_page_list:
        :return:
        """
        word = win32.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(doc_input)
        doc.Content.ShowAll = False
        heading_list = []
        # GetCrossReferenceItems(wdRefTypeHeading) returns an array with references to all headings in the document
        astrHeadings = doc.GetCrossReferenceItems(MyConstants.wdRefTypeHeading)
        for intItem in range(len(astrHeadings)):
            # Get the text and the level.
            text = astrHeadings[intItem].strip()
            num = text.split(" ")[0]
            level = WinWordML.get_heading_level(text, num)
            text = text.strip()
            text_list = text.split(" ")
            if len(text_list) > 1:
                onlytext = text.replace(text_list[0], "").strip()
                onlytext = onlytext.strip()
            else:
                onlytext = text
            heading_list.append({"num": num, "text": text, "level": level, "onlytext": onlytext, "block_id": intItem})
        doc.Close(True)
        word.Quit()
        return heading_list

    @staticmethod
    def create_outline2(doc_input):
        """
        In order delete blank page
        http://www.vbaexpress.com/forum/showthread.php?37455-Solved-Check-whether-a-loop-has-encountered-a-table-and-then-skip-it
        :param doc_input:
        :param blank_page_list:
        :return:
        """
        word = win32.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(doc_input)
        heading_list = []
        block_id = -1
        saved_showall = doc.Content.ShowAll
        doc.Content.ShowAll = True
        oRng = doc.Paragraphs(1).Range
        doc_rng_end = doc.Content.End
        done = False
        while not done:
            block_id += 1
            if not oRng.Information(MyConstants.wdWithInTable):
                rng_para = oRng
                style = rng_para.Style
                if str(style).startswith("Heading") or str(style).startswith("List"):
                    text = "%s %s" % (rng_para.ListFormat.ListString, rng_para.Text)
                    only_text = rng_para.Text
                    level = rng_para.ParagraphFormat.OutlineLevel
                    num = rng_para.ListFormat.ListString
                    if num != "":
                        current_page = rng_para.Information(MyConstants.wdActiveEndPageNumber)
                        heading_list.append({"num": num, "text": text, "level": level, "page": current_page,
                                             "only_text": only_text, "block_id": block_id, "heading": str(style),
                                             "style": str(style)})
                para_rng_end = oRng.End
                done = para_rng_end == doc_rng_end
                oRng.Collapse(MyConstants.wdCollapseEnd)
                oRng.MoveEnd(MyConstants.wdParagraph, 1)
            else:
                oSkipRng = oRng.Tables(1).Range
                para_rng_end = oSkipRng.End
                done = para_rng_end == doc_rng_end
                oSkipRng.Collapse(MyConstants.wdCollapseEnd)
                oSkipRng.MoveEnd(MyConstants.wdParagraph, 1)
                oRng = oSkipRng.Paragraphs(1).Range
        # print("        Collecting heading numbering...", block_id, para_rng_end, para_rng_end / doc_rng_end * 100)
        doc.Content.ShowAll = saved_showall
        doc.Close(True)
        word.Quit()
        return heading_list

    @staticmethod
    def create_outline3(doc_input):
        """
        In order delete blank page
        http://www.vbaexpress.com/forum/showthread.php?37455-Solved-Check-whether-a-loop-has-encountered-a-table-and-then-skip-it
        :param doc_input:
        :param blank_page_list:
        :return:
        """
        word = win32.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(doc_input)
        heading_list = []
        block_id = -1
        saved_showall = doc.Content.ShowAll
        doc.Content.ShowAll = True
        oRng = doc.Paragraphs(1).Range
        para_rng_end = oRng.Paragraphs(1).Range.End
        doc_rng_end = doc.Content.End
        done = False
        while not done:
            block_id += 1
            #print("        Collecting heading numbering...", block_id, para_rng_end, para_rng_end/ doc_rng_end*100)
            #print("       ", block_id, oRng.Text)
            if not oRng.Information(MyConstants.wdWithInTable):
                rng_para = oRng
                style = rng_para.Style
                if str(style).startswith("Heading") or str(style).startswith("List"):
                    text = "%s %s" % (rng_para.ListFormat.ListString, rng_para.Text)
                    size = len(rng_para.Text)
                    only_text = rng_para.Text
                    level = rng_para.ParagraphFormat.OutlineLevel
                    num = rng_para.ListFormat.ListString
                    if num != "":
                        current_page = rng_para.Information(MyConstants.wdActiveEndPageNumber)
                        heading_list.append({"num": num, "text": text, "level": level, "page": current_page,
                                             "only_text": only_text, "block_id": block_id, "heading": str(style),
                                             "style": str(style)})
                else:
                    if len(heading_list) > 0:
                        lst_heading = copy.deepcopy(heading_list[-1])
                        lst_heading["block_id"] = block_id
                        current_page = rng_para.Information(MyConstants.wdActiveEndPageNumber)
                        lst_heading["page"] = current_page
                        heading_list.append(lst_heading)

                para_rng_end = oRng.End
                done = para_rng_end == doc_rng_end
                oRng.Collapse(MyConstants.wdCollapseEnd)
                oRng.MoveEnd(MyConstants.wdParagraph, 1)
            else:
                oSkipRng = oRng.Tables(1).Range
                if len(heading_list) > 0:
                    lst_heading = copy.deepcopy(heading_list[-1])
                    lst_heading["block_id"] = block_id
                    current_page = oSkipRng.Information(MyConstants.wdActiveEndPageNumber)
                    lst_heading["page"] = current_page
                    heading_list.append(lst_heading)
                para_rng_end = oSkipRng.End
                done = para_rng_end == doc_rng_end
                oSkipRng.Collapse(MyConstants.wdCollapseEnd)
                oSkipRng.MoveEnd(MyConstants.wdParagraph, 1)
                oRng = oSkipRng.Paragraphs(1).Range

        # print("        Collecting heading numbering...", block_id, para_rng_end, para_rng_end / doc_rng_end * 100)
        doc.Content.ShowAll = saved_showall
        doc.Close(True)
        word.Quit()
        return heading_list

    @staticmethod
    def generate_all_number_heading_dict2(doc_input, key="block_id"):
        """ return a dict of numbering of heading with key is block_id
        """
        heading_list = WinWordML.create_outline2(doc_input)
        number_heading_dict = dict()
        for i,  layout_i in enumerate(heading_list):
            num = layout_i["num"]
            text = layout_i["text"]
            style = layout_i["style"]
            level = layout_i["level"]
            block_id = layout_i["block_id"]
            heading = layout_i["heading"]
            page = layout_i["page"]
            only_text = layout_i["only_text"]
            if key == "block_id":
                number_heading_dict[block_id] = {"num": num, "text": text, "level": level, "style": style,
                                                 "heading": heading, "block_id": block_id, "page": page,
                                                 "only_text": only_text}
        return number_heading_dict

    @staticmethod
    def generate_all_number_heading_dict3(doc_input, key="block_id"):
        """ return a dict of numbering of heading with key is block_id
        """
        heading_list = WinWordML.create_outline3(doc_input)
        number_heading_dict = dict()
        for i,  layout_i in enumerate(heading_list):
            num = layout_i["num"]
            text = layout_i["text"]
            style = layout_i["style"]
            level = layout_i["level"]
            block_id = layout_i["block_id"]
            heading = layout_i["heading"]
            page = layout_i["page"]
            only_text = layout_i["only_text"]
            if key == "block_id":
                number_heading_dict[block_id] = {"num": num, "text": text, "level": level, "style": style,
                                                 "heading": heading, "block_id": block_id, "page": page,
                                                 "only_text": only_text}
        return number_heading_dict

    @staticmethod
    def get_heading_level(strItem, num):
        """
        ' Return the heading level of a header from the
        ' array returned by Word.

        ' The number of leading spaces indicates the
        ' outline level (2 spaces per level: H1 has
        ' 0 spaces, H2 has 2 spaces, H3 has 4 spaces.
        :param strItem:
        :return:
        """
        # Get rid of all trailing spaces.
        # Trim leading spaces, and then compare with
        # the original.
        # Subtract to find the number of
        # leading spaces in the original string.
        intDiff = len(strItem.rstrip()) - len(strItem.strip())
        GetLevel = (intDiff / 2) + 1
        GetLevel1 = num.count(".")
        if GetLevel1 > GetLevel:
            GetLevel = GetLevel1
        return GetLevel

    @staticmethod
    def get_toc_from_doc_file(doc_input):
        """
        In order delete blank page
        :param doc_input:
        :param blank_page_list:
        :return:
        """
        word = win32.Dispatch("Word.Application")
        word.Visible = True
        doc = word.Documents.Open(doc_input)
        toc = doc.TablesOfContents(1)
        toc.IncludePageNumbers = True
        toc_text = toc.Range.Text
        toc_list = toc_text.split(chr(13))
        heading_list = []
        for toc_i in toc_list:
            if toc_i != '':
                parts = toc_i.split("\t")
                if len(parts) <= 2:
                    heading_list.append({"num": "", "text": parts[0], "page": parts[1]})
                else:
                    heading_list.append({"num": parts[0], "text": parts[1], "page": parts[2]})
        doc.Close(True)
        word.Quit()
        time.sleep(1)
        return heading_list

    @staticmethod
    def iter_paragraph(doc_input):
        """
        In order delete blank page
        :param doc_input:
        :param blank_page_list:
        :return:
        """
        word = win32.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(doc_input)

        for paragraph in doc.Paragraphs:
            style = paragraph.Range.Style
            if str(style).startswith("Heading"):
                print(paragraph.Range.Text)
        doc.Close(True)
        word.Quit()

    @staticmethod
    def create_toc(doc_input, UpperHeadingLevel=1, LowerHeadingLevel=5):
        """
        create toc
        :param doc_input:
        :param blank_page_list:
        :return:
        """
        tmp_folder = os.path.dirname(doc_input)
        tmp_folder = os.path.join(tmp_folder, "TempTOC")
        Folder.create_directory(tmp_folder)
        doc_output = os.path.join(tmp_folder, os.path.basename(doc_input)[:-5] + "_%s.docx" % "TOC")
        shutil.copyfile(doc_input, doc_output)
        word = win32.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(doc_output)
        doc.Repaginate()
        total_pages_before = doc.ComputeStatistics(2)
        myRange = doc.Range(0, 0)
        doc.TablesOfContents.Add(Range=myRange,
                                 RightAlignPageNumbers=True,
                                 UseHeadingStyles=True,
                                 UpperHeadingLevel=UpperHeadingLevel,
                                 LowerHeadingLevel=LowerHeadingLevel,
                                 IncludePageNumbers=True,
                                 AddedStyles="",
                                 UseHyperlinks=True,
                                 HidePageNumbersInWeb=True,
                                 UseOutlineLevels=True)
        doc.Repaginate()
        total_pages_after = doc.ComputeStatistics(2)
        toc = []
        for paragraph in doc.TablesOfContents(1).Range.Paragraphs:
            toc_i = paragraph.Range.Text.strip()
            toc.append(toc_i)
        doc.Close(MyConstants.wdSaveChanges)
        word.Quit()
        toc = toc[:-1]
        # os.remove(doc_output)
        result = dict()
        result["input"] = doc_input
        result["total_pages_before"] = total_pages_before
        result["total_pages_after"] = total_pages_after
        result["offset_page"] = total_pages_after - total_pages_before
        result["toc"] = toc
        return result

    @staticmethod
    def create_toc_level1(doc_folder, template_file, toc_file):
        doc_files = Folder.get_all_files(doc_folder, "*.docx")
        toc = []
        for doc_input in doc_files:
            toc_i = WinWordML.create_toc(doc_input, UpperHeadingLevel=1, LowerHeadingLevel=1)
            toc.append(toc_i)
        word = win32.Dispatch("Word.Application")
        word.Visible = False
        shutil.copyfile(template_file, toc_file)
        word.Visible = False
        doc = word.Documents.Open(toc_file)
        for toc_i in toc:
            for toc_i_text in toc_i["toc"]:
                doc.Content.InsertAfter(Text=toc_i_text)
        doc.Close(MyConstants.wdSaveChanges)
        word.Quit()
        return toc

    @staticmethod
    def merge_docs(doc_folder, template_file, target_file):
        word = win32.Dispatch("Word.Application")
        word.Visible = False
        shutil.copyfile(template_file, target_file)
        doc_files = Folder.get_all_files(doc_folder, "*.docx")
        for doc_input in doc_files:
            main_doc = word.Documents.Open(target_file)
            Range = main_doc.Content
            Range.Collapse(Direction=MyConstants.wdCollapseEnd)
            Range.InsertFile(FileName=doc_input)
            main_doc.Close(win32.constants.wdSaveChanges)
        word.Quit()

    @staticmethod
    def compute_page_info(doc_input):
        word = win32.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(doc_input)

        for paragraph in doc.Paragraphs:
            text = paragraph.Range.Text.strip()
            current_page = paragraph.Range.Information(MyConstants.wdActiveEndPageNumber)
            print(text, " - Page-", current_page)
        doc.Close(True)
        word.Quit()

    @staticmethod
    def get_latest_heading(rng_para):
        text = rng_para.Text
        ins_list = []
        del_list = []
        for j in range(1, rng_para.Sentences.Count + 1):
            rng_sent = rng_para.Sentences(j)
            if rng_sent.Revisions.Count <= 0:
                text += rng_sent.Text
                continue
            for k in range(1, rng_sent.Revisions.Count + 1):
                rev = rng_sent.Revisions(k)
                if rev.Type == MyConstants.wdRevisionInsert:
                    ins_list.append(rev.Range.Text)
                elif rev.Type == MyConstants.wdRevisionDelete:
                    del_list.append(rev.Range.Text)
                    rev.Accept()
        del_text = "".join(del_list)
        ins_text = "".join(ins_list)
        if len(text) > len(del_text) > 0:
            for j in range(1, rng_para.Sentences.Count + 1):
                rng_sent = rng_para.Sentences(j)
                for k in range(1, rng_sent.Revisions.Count + 1):
                    rev = rng_sent.Revisions(k)
                    if rev.Type == MyConstants.wdRevisionDelete:
                        rev.Accept()
            text = rng_para.Text
        return text

    @staticmethod
    def track_change(doc_input, doc_output, report_file):
        shutil.copyfile(doc_input, doc_output)
        word = win32.Dispatch("Word.Application")
        word.Visible = True
        doc = word.Documents.Open(doc_output)
        change_sentence_list = []
        if doc.Revisions.Count <= 0:
            return change_sentence_list
        for i in range(1, doc.Paragraphs.Count + 1):
            rng_para = doc.Paragraphs(i).Range
            for j in range(1, rng_para.Sentences.Count + 1):
                rng_sent = rng_para.Sentences(j)
                if rng_sent.Revisions.Count <= 0:
                    continue
                before = rng_sent.Text.strip()
                after = rng_sent.Text.strip()
                ins_list = []
                del_list = []
                for k in range(1, rng_sent.Revisions.Count + 1):
                    rev = rng_sent.Revisions(k)
                    if rev.Type == MyConstants.wdRevisionInsert:
                        ins_list.append(rev.Range.Text)
                    elif rev.Type == MyConstants.wdRevisionDelete:
                        del_list.append(rev.Range.Text)
                    #########rev.Accept()
                if len(ins_list) <= 0 and len(del_list) <= 0:
                    break
                after = rng_sent.Text.strip()
                if len(ins_list) <= 0 < len(del_list):
                    change_type = "Deleted Sentence"
                elif len(del_list) <= 0 < len(ins_list):
                    change_type = "Inserted Sentence"
                else:
                    change_type = "Modified Sentence"
                current_page = rng_para.Information(MyConstants.wdActiveEndPageNumber)
                change_dict = {"ParaID": i, "SID": j, "before": before, "after": after, "type": change_type,
                               "page": current_page, "ins_list": ins_list, "del_list": del_list}
                change_sentence_list.append(change_dict)
                if len(change_sentence_list) >= 10:
                    break
        # update heading numbering
        change_sentence_list1 = []
        for change_dict in change_sentence_list:
            change_dict["heading_level"] = ""
            change_dict["text"] = ""
            change_dict["heading_text"] = ""
            change_dict["heading_number"] = ""
            i = change_dict["ParaID"]
            for k in range(i, 0, -1):
                rng_para = doc.Paragraphs(k).Range
                style = rng_para.Style
                if str(style).startswith("Heading"):
                    # https://wordmvp.com/FAQs/Numbering/ListString.htm
                    change_dict["heading_text"] = "%s%s" % (rng_para.ListFormat.ListString, rng_para.Text)
                    change_dict["text"] = rng_para.Text
                    change_dict["heading_level"] = rng_para.ParagraphFormat.OutlineLevel
                    change_dict["heading_number"] = rng_para.ListFormat.ListString
            change_sentence_list1.append(change_dict)

        doc.Close(False)
        word.Quit()
        if report_file is not None:
            from services.reporter2 import Reporter2
            Template = os.path.join(Config.BASE_DIR, "Tracers/Templates/Template_Tracer_Report.xlsx")
            config_file = os.path.join(Config.BASE_DIR, "Tracers/config/ins_del_trace_multifile_config.yaml")
            report = Reporter(Template, config_file, "InsDelTracerMulti")
            records = [change_dict["heading_text"] for change_dict in change_sentence_list1]
            report.create_detail_report(records)
            report.save(report_file)
        return change_sentence_list

    @staticmethod
    def clean_gen_cache():
        try:
            print('        clean_gen_cache...')
            MODULE_LIST = [m.__name__ for m in sys.modules.values()]
            for module in MODULE_LIST:
                if re.match(r'win32com\.gen_py\..+', module):
                    del sys.modules[module]
            shutil.rmtree(os.path.join(os.environ.get('LOCALAPPDATA'), 'Temp', 'gen_py'))
        except Exception as e:
            pass

    @staticmethod
    def kill_winword():
        PROCNAME = "WINWORD.EXE"
        for proc in psutil.process_iter():
            if proc.name() == PROCNAME:
                print('        Terminating %s...' % PROCNAME)
                proc.terminate()
                break

    @staticmethod
    def doc2pdf(doc_input, pdf_output):
        word = win32.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False
        doc = word.Documents.Open(doc_input)
        if doc.Comments.Count >= 1:
            doc.DeleteAllComments()
        doc.TrackRevisions = False
        doc.Fields.Locked = True
        doc.ExportAsFixedFormat(OutputFileName=pdf_output,
                                ExportFormat=17,
                                OpenAfterExport=False,
                                OptimizeFor=0,
                                CreateBookmarks=1,
                                DocStructureTags=True,
                                Item=0)
        doc.Save()
        doc.Close(False)
        time.sleep(1)
        word.DisplayAlerts = True
        word.Quit()

    @staticmethod
    def accept_all_changes(doc_input, doc_output):
        word = win32.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False
        doc = word.Documents.Open(doc_input)
        if doc.Comments.Count >= 1:
            doc.DeleteAllComments()
        doc.TrackRevisions = False
        doc.Revisions.AcceptAll()
        doc.Fields.Locked = True
        doc.SaveAs(FileName=doc_output)
        doc.Saved = True
        doc.Close(False)
        time.sleep(1)
        word.DisplayAlerts = True
        word.Quit()

    @staticmethod
    def doc2docx(doc_input, doc_output):
        word = win32.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False
        doc = word.Documents.Open(doc_input)
        doc.SaveAs2(doc_output, FileFormat=16)
        doc.Saved = True
        doc.Close(False)
        time.sleep(1)
        word.DisplayAlerts = True
        word.Quit()

    @staticmethod
    def close_doc(word_app, word_file):
        word_file.Saved = True
        word_app.DisplayAlerts = False
        word_file.Close(False)
        word_app.DisplayAlerts = True

    @staticmethod
    def send_keys(key="{ENTER}", times=1):
        """
        https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/sendkeys-statement
        :return:
        """
        shell = win32.Dispatch("WScript.Shell")
        for i in range(times):
            shell.SendKeys(key)


if __name__ == '__main__':
    testfile = r"D:\Tools\UMC\TS2UM\Input\TestWinWordML\R-CarV4H_TS_067_DSI_CSI2_TX.docx"
    testfile1 = r"D:\Tools\UMC\TS2UM\Input\TestWinWordML\R-CarV4H_TS_002_Area_map.docx"
    testfolder = r"D:\Tools\UMC\TS2UM\Input\TestWinWordML\TOC"
    template = r"D:\Tools\UMC\TS2UM\Input\TestWinWordML\TOC\Templates\Template_TOC.docx"
    toc_file = r"D:\Tools\UMC\TS2UM\Input\TestWinWordML\TOC\TOC.docx"
    target_file = r"D:\Tools\UMC\TS2UM\Input\TestWinWordML\TOC\TOC.docx"
    merge_folder = r"D:\Tools\UMC\TS2UM\Input\TestWinWordML\MergeDoc"
    track_file = r"D:\Tools\UMC\TS2UM\Input\TestWinWordML\TrackChanges\R-CarV4H_UM_002_Area_map.docx"
    # track_file = r"D:\Tool\UMC\TS2UM\Input\TrackChanges\04_S4_security_UM_IOComparator.docx"
    report_file = r"D:\Tool\UMC\TS2UM\Input\TrackChanges\04_S4_security_UM_IOComparator.xlsx"
    track_file_output = r"D:\Tools\UMC\TS2UM\Input\TestWinWordML\TrackChanges\R-CarV4H_UM_002_Area_map_Output.docx"
    track_file_output = r"D:\Tool\UMC\TS2UM\Input\TrackChanges\04_S4_security_UM_IOComparator_Output.docx"
    track_file = r"D:\Tool\UMC\TS2UM\Input\TrackChanges\TC13.docx"
    track_file_output = r"D:\Tool\UMC\TS2UM\Input\TrackChanges\TC13_output.docx"
    # print(WinWordML.get_number_page(testfile)) # 419
    # print(WinWordML.compute_page_info(testfile)) # 419
    # print(WinWordML.create_outline(testfile)) # 419
    # print(WinWordML.get_toc_from_doc_file(testfile)) # 419
    # print(WinWordML.iter_paragraph(testfile)) # 419
    print(WinWordML.track_change(track_file, track_file_output, report_file))
    # print(WinWordML.create_toc(testfile)) # 419
    # print(WinWordML.create_toc(testfile1)) # 419
    # print(WinWordML.create_toc_level1(merge_folder, template, toc_file)) # 419
    # print(WinWordML.merge_docs(merge_folder, template, target_file)) # 419
    # print(WinWordML.track_folder(merge_folder, template, target_file)) # 419
    #testfile = r"D:\Tools\UMC\TS2UM\Input\R-CarV4H_UM_009_Module_Standby_Software_Reset_D.docx"
    #BlockNumberHeadingDict = WinWordML.generate_all_number_heading_dict(testfile, "block_id")
