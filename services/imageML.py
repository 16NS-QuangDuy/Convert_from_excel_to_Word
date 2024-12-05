# encoding: utf-8
import os
import re
import zipfile
# from PIL import Image, ImageChops, ImageDraw
from settings.config import Config
from services.folder import Folder


class ImageML:
    """ImageML"""

    config_file = Config.get_default_config_file(__file__)
    DIFFERENT_SIZE = 120
    DIFFERENT_TYPE = 130
    SAME_IMAGE_THRESH = 10

    def __init__(self):
        self.auto_worker_name = self.__class__.__name__
        self.config = Config()
        Config.set_attr_from_yaml(self.__dict__["config"], self.config_file, self.auto_worker_name)

    @staticmethod
    def diff_image_a_b(image_a, image_b, image_c, opacity=0.85):
        """check if the two images is has difference or not
        :param image_a: first image file path
        :type image_a: str
        :param image_b: second image file path
        :type image_b: str
        :param image_c: different image file path
        :type image_c: str
        :param opacity: gray rate
        :type opacity: double
        """
        if os.path.basename(image_a).endswith(".wmf") or os.path.basename(image_a).endswith(".emf"):
            image_a = ImageML.convert_wmf_emf_2png(image_a)
            image_b = ImageML.convert_wmf_emf_2png(image_b)
            image_c = os.path.splitext(image_c)[0] + ".png"
        a = Image.open(image_a)
        b = Image.open(image_b)
        completely_black, diff = ImageML.customized_diff_image_a_b(a, b)
        if not completely_black:
            diff = ImageML.customized_diff_image(a, b, diff, opacity)
            diff.save(image_c)
        a.close()
        b.close()
        diff.close()
        return completely_black

    @staticmethod
    def customized_diff_image_a_b(a, b):
        """
        will test to see whether an image is completely black.
        (Image.getbbox() returns the falsy None if there are no non-black pixels in the image,
        otherwise it returns a tuple of points, which is truthy.)
        """
        diff = ImageChops.difference(a, b)
        diff = diff.convert('L')
        completely_black = False
        if not diff.getbbox():
            completely_black = True
        return completely_black, diff

    @staticmethod
    def new_gray(size, color):
        img = Image.new('L', size)
        dr = ImageDraw.Draw(img)
        dr.rectangle((0, 0) + size, color)
        return img

    @staticmethod
    def customized_diff_image(a, b, diff, opacity=0.85):
        # Hack: there is no threshold in PILL,
        # so we add the difference with itself to do
        # a poor man's thresholding of the mask:
        # (the values for equal pixels-  0 - don't add up)
        thresholded_diff = diff
        for repeat in range(3):
            thresholded_diff = ImageChops.add(thresholded_diff, thresholded_diff)
        h, w = size = diff.size
        mask = ImageML.new_gray(size, int(255 * (opacity)))
        shade = ImageML.new_gray(size, 0)
        new = a.copy()
        new.paste(shade, mask=mask)
        # To have the original image show partially
        # on the final result, simply put "diff" instead of thresholded_diff bellow
        new.paste(b, mask=thresholded_diff)
        return new

    @staticmethod
    def extract_images_from_word_file(filepath, output_dir="", image_ext_list=[]):
        output_file_names = []
        default_output_dir = os.path.join(Config.OutputComparatorDir, "WordImages", "A")
        default_image_ext = ".*"
        if output_dir != "":
            default_output_dir = output_dir
        if len(image_ext_list) >= 1:
            default_image_ext = '|'.join(image_ext_list)
        myzip = zipfile.ZipFile(filepath)
        namelist = [name for name in myzip.namelist() if re.match(r'word/media/.*\.(%s)' % default_image_ext, name)]
        Folder.create_directory(default_output_dir)
        for name in namelist:
            out_file_name = os.path.join(default_output_dir, os.path.basename(name))
            with myzip.open(name) as myfile:
                with open(out_file_name, 'wb') as destination:
                    destination.write(myfile.read())
            output_file_names.append(out_file_name)
        for file in output_file_names:
            if os.path.splitext(file)[1] in [".wmf", ".emf"]:
                image_a = ImageML.convert_wmf_emf_2png(file)
                output_file_names.append(image_a)
        return output_file_names

    @staticmethod
    def convert_wmf_emf_2png(image_wmf_emf):
        new_image_a = os.path.splitext(image_wmf_emf)[0] + ".png"
        with open(image_wmf_emf, 'rb') as image_file:
            with Image.open(image_file) as img:
                img.save(new_image_a)
        return new_image_a

    @staticmethod
    def get_image_size(image_a):
        image = Image.open(image_a)
        h, w = image.size
        size = os.path.getsize(image_a)
        image.close()
        return h, w, size

    @staticmethod
    def copy_file(input_file, output_file):
        with open(input_file, 'rb') as image_file:
            with Image.open(image_file) as img:
                img.save(output_file)

    @staticmethod
    def diff_image_dir(dir_a, dir_b, output_dir, opacity=0.85):
        dir_a_list = Folder.get_all_files(dir_a, "*.png") + Folder.get_all_files(dir_a, "*.jpg")
        dir_b_list = Folder.get_all_files(dir_b, "*.png") + Folder.get_all_files(dir_b, "*.jpg")
        kept_a_dict = dict()
        kept_b_dict = dict()
        diff_a_dict = dict()
        diff_b_dict = dict()
        for file_a in dir_a_list:
            min = 110
            candidate = None
            for file_b in dir_b_list:
                if file_b in kept_b_dict:
                    continue
                percent = ImageML.diff_image_percentage(file_a, file_b)
                if percent == 0:
                    kept_a_dict[file_a] = file_b
                    kept_b_dict[file_b] = file_a
                    Folder.create_directory(os.path.join(output_dir, "kepta_img"))
                    Folder.create_directory(os.path.join(output_dir, "keptb_img"))
                    ImageML.copy_file(file_a, os.path.join(output_dir, "kepta_img", os.path.basename(file_a)))
                    ImageML.copy_file(file_b, os.path.join(output_dir, "keptb_img", os.path.basename(file_b)))
                    break
                elif 0 < percent < min:
                    min = percent
                    candidate = file_b
            if candidate is not None:
                diff_a_dict[file_a] = candidate
                diff_b_dict[candidate] = file_a
        for file_a in dir_a_list:
            if file_a not in kept_a_dict and file_a not in diff_a_dict:
                Folder.create_directory(os.path.join(output_dir, "del_img"))
                ImageML.copy_file(file_a, os.path.join(output_dir, "del_img", os.path.basename(file_a)))
        for file_b in dir_b_list:
            if file_b not in kept_b_dict and file_b not in diff_b_dict:
                Folder.create_directory(os.path.join(output_dir, "ins_img"))
                ImageML.copy_file(file_b, os.path.join(output_dir, "ins_img", os.path.basename(file_b)))
        for file_a in diff_a_dict:
            Folder.create_directory(os.path.join(output_dir, "diff_img"))
            new_name = "diff_{%s}_{%s}.png" % (os.path.basename(file_a), os.path.basename(diff_a_dict[file_a]))
            ImageML.diff_image_a_b(file_a, diff_a_dict[file_a], os.path.join(output_dir, "diff_img", new_name), opacity)

    @staticmethod
    def diff_image_percentage(image_a, image_b):
        i1 = Image.open(image_a)
        i2 = Image.open(image_b)
        if not i1.mode == i2.mode:
            print("Different kinds of images.")
            return ImageML.DIFFERENT_TYPE
        if not i1.size == i2.size:
            print("Different sizes.")
            return ImageML.DIFFERENT_SIZE
        pairs = zip(i1.getdata(), i2.getdata())
        if len(i1.getbands()) == 1:
            # for gray-scale jpegs
            dif = sum(abs(p1 - p2) for p1, p2 in pairs)
        else:
            dif = sum(abs(c1 - c2) for p1, p2 in pairs for c1, c2 in zip(p1, p2))
        ncomponents = i1.size[0] * i1.size[1] * 3
        print("Difference (percentage):", (dif / 255.0 * 100) / ncomponents)
        return (dif / 255.0 * 100) / ncomponents

    @staticmethod
    def get_diff_image_dir_result(output_dir):
        def create_init_layout_i():
            return {"change": "", "before": "", "after": "", "ins": "", "del": ""}
        records = []
        ins_list = [os.path.basename(f) for f in Folder.scan_all_files(os.path.join(output_dir, "ins_img"))]
        del_list = [os.path.basename(f) for f in Folder.scan_all_files(os.path.join(output_dir, "del_img"))]
        diff_list = [os.path.basename(f) for f in Folder.scan_all_files(os.path.join(output_dir, "diff_img"))]
        for ins_img in ins_list:
            layout_i = create_init_layout_i()
            layout_i["change"] = "ins"
            layout_i["after"] = ins_img
            layout_i["ins"] = ins_img
            records.append(layout_i)
        for del_img in del_list:
            layout_i = create_init_layout_i()
            layout_i["change"] = "del"
            layout_i["before"] = del_img
            layout_i["del"] = del_img
            records.append(layout_i)
        for diff_img in diff_list:
            match = re.match(r"diff_{(.*)}_{(.*)}.png", os.path.basename(diff_img))
            layout_i = create_init_layout_i()
            layout_i["change"] = "kept"
            layout_i["before"] = match.group(1)
            layout_i["after"] = match.group(2)
            layout_i["ins"] = diff_img
            records.append(layout_i)
        return records
