# encoding: utf-8
import threading
import sys
import os


class Capture:
    def __init__(self):
        self._stdout = None
        self._stderr = None
        self._r = None
        self._w = None
        self._thread = None
        self._on_readline_cb = None
        self._text = ""

    def _handler(self):
        while not self._w.closed:
            try:
                while True:
                    line = self._r.readline()
                    if len(line) == 0: break
                    if self._on_readline_cb: self._on_readline_cb(line)
                    self._text += line
            except:
                break

    def on_readline(self, callback):
        self._on_readline_cb = callback

    def start(self):
        self._stdout = sys.stdout
        self._stderr = sys.stderr
        r, w = os.pipe()
        r, w = os.fdopen(r, 'r'), os.fdopen(w, 'w', 1)
        self._r = r
        self._w = w
        sys.stdout = self._w
        sys.stderr = self._w
        self._thread = threading.Thread(target=self._handler)
        self._thread.start()
        self._text = ""

    def stop(self):
        self._w.close()
        if self._thread: self._thread.join()
        self._r.close()
        sys.stdout = self._stdout
        sys.stderr = self._stderr

    def get_text(self):
        return self._text.rstrip()

    def reset_text(self):
        self._text = ""

    def save_to_file(self, folder="", file_name="capture.txt", file_path=""):
        if folder != "" and file_name != "":
            file_path = os.path.join(folder, file_name)
        if os.path.exists(file_path):
            os.remove(file_path)
        if not os.path.exists(os.path.dirname(file_path)):
            os.makedirs(os.path.dirname(file_path))
        with open(file_path, 'w+', encoding="utf-8") as f:
            f.write(self.get_text())
            f.close()

    @staticmethod
    def read_text_file(out_file):
        if os.path.exists(out_file):
            with open(out_file, 'r', encoding="utf-8", errors='ignore') as f:
                out_content = "".join(f.readlines())
                f.close()
                return out_content
        return "FILE DOES NOT EXIST"

    @staticmethod
    def print_records_of_dict_in_cvs(records, empty_text="__EMPTY__"):
        if len(records) <= 0:
            print(empty_text)
        else:
            print(",".join([k for k, v in records[0].items()]))
            for layout_i in records:
                v_list = []
                for k, v in layout_i.items():
                    if isinstance(v, list):
                        v = " ".join(v)
                    else:
                        v = str(v)
                    v_list.append(v)
                print(",".join(v_list))

