# encoding: utf-8
import os
import glob
import time
import re
import shutil
from settings.config import Config
from services.reporter import Reporter
from fnmatch import fnmatch


class Folder:
    """handler for doing folder operation"""
    filter_out_list = ["_Tracker"]

    @staticmethod
    def create_directory(directory, exist_del=False):
        if not os.path.exists(directory):
            os.makedirs(directory)
        elif exist_del:
            try:
                Folder.empty_directory(directory)
            except:
                pass

    @staticmethod
    def empty_directory(directory):
        files = Folder.scan_all_files(directory)
        files += Folder.scan_all_files("%s/*/" % directory)
        files += Folder.scan_all_files("%s/*/*/" % directory)
        for file in files:
            os.remove(file)
        folders = os.listdir(directory)
        for folder in folders:
            os.rmdir(os.path.join(directory, folder))

    @staticmethod
    def empty_directory1(directory):
        shutil.rmtree(directory, ignore_errors=True)

    @staticmethod
    def filter_out_temp_files(files):
        rc = []
        for f in files:
            if "~" not in f:
                rc.append(f)
        return rc

    @staticmethod
    def filter_out_tracker(files):
        rc = []
        for f in files:
            is_not_input = False
            for afilter in Folder.filter_out_list:
                if afilter in f.split('/')[-1]:
                    is_not_input = True
                    break
            if is_not_input:
                continue
            rc.append(f)
        return rc

    @staticmethod
    def filter_out_not_working_ts(files, working_ts_list):
        files = Folder.filter_out_tracker(files)
        rc = []
        if working_ts_list is not None:
            for f in files:
                f1 = os.path.splitext(f)[0]
                for working_ts in working_ts_list:
                    if f1.endswith(working_ts):
                        rc.append(f)
                        break
            return rc
        return files

    @staticmethod
    def get_all_docx(input_dir, display=True):
        doc_input_list = []
        filter = "%s%s" % (input_dir, "\\*.docx")
        files = glob.glob(filter)
        files = Folder.filter_out_temp_files(files)
        if len(files) <= 0:
            print("There is no docx in the folder")
        else:
            doc_input_list = files
        if Config.Debug:
            auto_worker_name = Folder.__class__.__name__
            Reporter.create_command_window_report(auto_worker_name, doc_input_list,
                                                  "file")
        global_doc_input_list = []
        for item in doc_input_list:
            fullpath = os.path.join(input_dir, item).replace("\\", "/")
            global_doc_input_list.append(fullpath)
        return global_doc_input_list

    @staticmethod
    def get_all_doc(input_dir, display=True):
        doc_input_list = []
        filter = "%s%s" % (input_dir, "\\*.doc")
        files = glob.glob(filter)
        files = Folder.filter_out_temp_files(files)
        if len(files) <= 0:
            print("There is no docx in the folder")
        else:
            for f in files:
                if f.endswith("doc"):
                    doc_input_list.append(f)
        if Config.Debug:
            auto_worker_name = Folder.__class__.__name__
            Reporter.create_command_window_report(auto_worker_name, doc_input_list,
                                                  "file")
        global_doc_input_list = []
        for item in doc_input_list:
            fullpath = os.path.join(input_dir, item).replace("\\", "/")
            global_doc_input_list.append(fullpath)
        return global_doc_input_list

    @staticmethod
    def get_all_xlsx(input_dir, contain="", display=True):
        doc_input_list = []
        filter = "%s%s" % (input_dir, "\\*.xlsx")
        files = glob.glob(filter)
        files = Folder.filter_out_temp_files(files)
        if len(files) <= 0:
            print("There is no excel in the folder")
        else:
            for f in files:
                if f.endswith("xlsx") and contain in f:
                    doc_input_list.append(f)
        if Config.Debug:
            auto_worker_name = Folder.__class__.__name__
            Reporter.create_command_window_report(auto_worker_name, doc_input_list,
                                                  "file")
        global_doc_input_list = []
        for item in doc_input_list:
            fullpath = os.path.join(input_dir, item).replace("\\", "/")
            global_doc_input_list.append(fullpath)
        return global_doc_input_list

    @staticmethod
    def get_all_yaml(input_dir, contain="", display=True):
        doc_input_list = []
        filter = "%s%s" % (input_dir, "\\*.yaml")
        files = glob.glob(filter)
        files = Folder.filter_out_temp_files(files)
        if len(files) <= 0:
            print("        There is no yaml in the folder, access database")
        else:
            for f in files:
                if f.endswith("yaml") and contain in f:
                    doc_input_list.append(f)
        if Config.Debug:
            auto_worker_name = Folder.__class__.__name__
            Reporter.create_command_window_report(auto_worker_name, doc_input_list,
                                                  "file")
        global_doc_input_list = []
        for item in doc_input_list:
            fullpath = os.path.join(input_dir, item).replace("\\", "/")
            global_doc_input_list.append(fullpath)
        return global_doc_input_list

    @staticmethod
    def get_all_basenames(files):
        file_list = []
        for f in files:
            if "~" not in f:
                file_list.append(os.path.basename(f))
        return file_list

    @staticmethod
    def get_yaml_by_docx(input_dir, contain="", match_doc_file=None, yaml_suffix=None, display=True):
        doc_input_list = []
        filter = "%s%s" % (input_dir, "\\*.yaml")
        files = glob.glob(filter)
        files = Folder.filter_out_temp_files(files)
        if len(files) <= 0:
            print("        There is no yaml in the folder, access database")
        else:
            for f in files:
                if f.endswith("yaml") and contain in f:
                    doc_input_list.append(f)
        for f in  doc_input_list:
            saved_f = f
            f = os.path.basename(f)
            if yaml_suffix is not None:
                f = f.replace(yaml_suffix, "")
            if match_doc_file is not None:
                if f in match_doc_file:
                    f = saved_f
                    f = os.path.join(input_dir, f).replace("\\", "/")
                    return f

    @staticmethod
    def get_all_working_ts(input_dir, display=True, ext=".docx"):
        filter = "%s%s" % (input_dir, "\\*%s" % ext)
        files = glob.glob(filter)
        files = Folder.filter_out_temp_files(files)
        doc_input_list = files
        global_doc_input_list = []
        for item in doc_input_list:
            working_ts = os.path.basename(item)
            working_ts = working_ts.replace(ext, "")
            global_doc_input_list.append(working_ts)
        if Config.Debug:
            auto_worker_name = Folder.__class__.__name__
            Reporter.create_command_window_report(auto_worker_name, doc_input_list,
                                                  "file")
        return global_doc_input_list

    @staticmethod
    def get_all_files(input_dir, search_pattern):
        """
        Params::
            input_dir: self.config.input_dir
            search_pattern: self.InputExtension: "*.<pattern>"
                eg: ["*.docx","*.doc"] or just "*.docx"
        Returns:: 
            files: list of files
        """
        files = []
        if isinstance(search_pattern, list):
            for ext in search_pattern:
                pathname = input_dir + '.' + os.sep + ext
                _files = glob.glob(pathname)
                files += Folder.filter_out_temp_files(_files)
        else:
            pathname = input_dir + '.' + os.sep + search_pattern
            _files = glob.glob(pathname)
            files += Folder.filter_out_temp_files(_files)
        if Config.Debug:
            print("%-6s%s" % (" ", "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"))
            print("%-6s%s" % (" ", "SUMMARY REPORT"))
            print("%-6s%s%s" % (" ", "- NUMBER OF ITEMS:", len(files)))
            print("%-6s%s" % (" ", "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"))
            if len(files) >= 1:
                print("%-6s%s" % (" ", "DETAIL REPORT"))
                for item in files:
                    print("%-6s%-10s %-50s" % (" ", "- file:", os.path.basename(item).encode("utf-8")))
                print("%-6s%s" % (" ", "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"))
        return files

    @staticmethod
    def scan_all_files(input_dir):
        filter = "%s%s" % (input_dir, "\\*")
        return [file for file in glob.glob(filter) if os.path.isfile(file)]

    @staticmethod
    def scan_all_folder(input_dir):
        filter = "%s%s" % (input_dir, "\\*")
        return [file for file in glob.glob(filter) if os.path.isdir(file)]

    @staticmethod
    def is_server_path(path):
        if path.startswith(r"\\"):
            return True
        return False

    @staticmethod
    def get_all_files2(input_dir, search_pattern):
        """
        Params::
            input_dir: self.config.input_dir
            search_pattern: self.InputExtension: "*.<pattern>"
                eg: ["*.docx","*.doc"] or just "*.docx"
        Returns::
            files: list of files
        """
        my_files = []
        if isinstance(search_pattern, list):
            for ext in search_pattern:
                for path, subdirs, files in os.walk(input_dir):
                    for name in files:
                        if fnmatch(name, ext) and "~" not in name:
                            my_files.append(Folder.standardize_path(os.path.join(path, name)))
        else:
            for path, subdirs, files in os.walk(input_dir):
                for name in files:
                    if fnmatch(name, search_pattern) and "~" not in name:
                        my_files.append(Folder.standardize_path(os.path.join(path, name)))
        if Config.Debug:
            print("%-6s%s" % (" ", "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"))
            print("%-6s%s" % (" ", "SUMMARY REPORT"))
            print("%-6s%s%s" % (" ", "- NUMBER OF ITEMS:", len(my_files)))
            print("%-6s%s" % (" ", "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"))
            if len(my_files) >= 1:
                print("%-6s%s" % (" ", "DETAIL REPORT"))
                for item in my_files:
                    print("%-6s%-10s %-50s" % (" ", "- file:", os.path.basename(item).encode("utf-8")))
                print("%-6s%s" % (" ", "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"))
        return my_files

    @staticmethod
    def standardize_path(doc_file, type="Windows"):
        doc_file = os.path.abspath(doc_file)
        if type=="Windows":
            doc_file = doc_file.replace('/', '\\')
        else:
            doc_file = doc_file.replace('\\', '/')
        return doc_file

    @staticmethod
    def get_has_env_var_path(path):
        match = re.search(r'\{(.*)\}', path)
        match1 = re.search(r'\%(.*)\%', path)
        if match1:
            match = match1
        if match:
            var_name = match.group(1)
            var_value = os.environ.get(var_name)
            return True, path.replace(match.group(0), var_value)
        return False, path

    @staticmethod
    def scan_all_files_with_recursive(path, ext):
        # Returns a list of names in list files.
        path_name = '%s/**/%s' % (path, ext)
        files = glob.glob(path_name, recursive=True)
        return files

    @staticmethod
    def copy_tree(src_folder_path, dst_folder_path):
        src_files = Folder.get_all_files2(src_folder_path, "*.*")
        src_folder_path = os.path.abspath(src_folder_path)
        src_folder_path = src_folder_path.replace('\\', '/')
        for fi in src_files:
            fi = fi.replace('\\', '/')
            link = fi.replace(src_folder_path, "")
            target_file = dst_folder_path + link
            target_dir = os.path.dirname(target_file)
            Folder.create_directory(target_dir)
            basename = os.path.basename(fi)
            target_file = os.path.join(target_dir, basename)
            shutil.copyfile(fi, target_file)
        return dst_folder_path

    @staticmethod
    def collect_files(input_dir, output_dir, extension_list):
        """ collect_files
        """
        for extention in extension_list:
            file_list = Folder.get_all_files2(input_dir, extention)
            for file_i in file_list:
                target_file = Folder.get_target_output(file_i, input_dir, output_dir)
                try:
                    shutil.copyfile(file_i, target_file)
                except Exception as ex:
                    print(str(ex))
                    pass

    @staticmethod
    def get_target_output(doc_file, input_dir, output_dir):
        doc_file = os.path.abspath(doc_file)
        input_dir = os.path.abspath(input_dir)
        doc_file = doc_file.replace('\\', '/')
        input_dir = input_dir.replace('\\', '/')
        link = doc_file.replace(input_dir, "")
        target_file = output_dir + link
        target_dir = os.path.dirname(target_file)
        Folder.create_directory(target_dir)
        basename = os.path.basename(doc_file)
        target_file = os.path.join(target_dir, basename)
        parts = os.path.splitext(target_file)
        if len(parts) > 1:
            ext = parts[1]
        else:
            ext = parts[0]
        target_file = os.path.splitext(target_file)[0] + ext

        return target_file

    @staticmethod
    def clean_files(folder_path, extension, contain_list):
        file_list = Folder.get_all_files2(folder_path, extension)
        if len(file_list) > 0:
            for file_i in file_list:
                f_name = os.path.basename(file_i)
                for contain in contain_list:
                    if contain in f_name:
                        os.remove(file_i)


if __name__ == '__main__':
    folderpath = r"\10.46.31.46\Bacnt4\S4_UM04"
    files = Folder.scan_all_folder(folderpath)
    print(files)
    print(Folder.is_server_path(folderpath))

