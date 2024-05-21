import json
import os
import sys
import uuid
from copy import deepcopy
from zipfile import ZipFile, ZIP_DEFLATED

from natsort import natsorted
from win32com import client
from get_image_size import get_image_size

VERSION = "1.3.1"
YEAR = "2024"
AUTHOR = "Martin Lehmann"
target_ratio = 2  # target aspect ratio for slides in h5p


def ppt2image(file):
    try:
        powerpoint = client.Dispatch("Powerpoint.Application")
    except Exception as e:
        print("Powerpoint could not be opened", file=sys.stderr)
        raise e
    try:  # look if an active presentation is open
        assert powerpoint.ActivePresentation is not None
        QUIT = False  # don't quit powerpoint later
    except:
        QUIT = True  # quit powerpoint later
    ppt = powerpoint.Presentations.Open(file)
    ppt.Export(file, "jpg")
    ppt.Close()
    if QUIT:  # quit only if required
        powerpoint.Quit()


def add_to_json(newfile, image_folder, images, title):
    reserved_files = ["content\content.json", ".\h5p.json"]
    if getattr(sys, "frozen", False):  # calling packaged binary
        basedir = sys._MEIPASS
    else:  # calling local script
        basedir = os.path.dirname(os.path.abspath(__file__))
    template = os.path.join(basedir, "template")

    img_width, img_height = get_image_size(os.path.join(image_folder, images[0]))

    with ZipFile(newfile, "w", compression=ZIP_DEFLATED) as zout:
        # add all other files to zip
        for dir_, _, files in os.walk(template):
            for file in files:
                rel_dir = os.path.relpath(dir_, template)
                rel_file = os.path.join(rel_dir, file)
                abs_file = os.path.join(dir_, file)
                if rel_file in reserved_files:
                    continue
                zout.write(abs_file, rel_file)

        # add image filenames to content.json
        with open(template + "/content/content.json") as fp:
            content = json.load(fp)
        print(f"adding images from::\n\t {image_folder} {images}")
        for i, image in enumerate(images):
            slides = content["presentation"]["slides"]
            # clone first element
            if i > 0 and len(slides) <= i:
                slides.append(deepcopy(slides[0]))
            elements = slides[i]["elements"][0]
            params = elements["action"]["params"]
            # add new image filename
            params["file"]["path"] = "images/" + image
            # add random uuid
            elements["action"]["subContentId"] = str(uuid.uuid4())
            # set width & height
            params["file"]["width"] = img_width
            params["file"]["height"] = img_height
            ratio = img_width / img_height
            if ratio > target_ratio:  # wider, need to shrink y
                elements["y"] = 100 * (1 - target_ratio / ratio) / 2
                elements["height"] = 100 * target_ratio / ratio
            elif ratio < target_ratio:  # higher, need to shrink x
                elements["x"] = 100 * (1 - ratio / target_ratio) / 2
                elements["width"] = 100 * ratio / target_ratio
        # save file
        zout.writestr("content/content.json", json.dumps(content))

        # add image files to tip
        for image in images:
            zout.write(os.path.join(image_folder, image), "content/images/" + image)

        # change presentation title
        with open(template + "/h5p.json", "r") as h5p:
            content = json.load(h5p)
            content["title"] = title
        zout.writestr("h5p.json", json.dumps(content))


if __name__ == "__main__":
    try:
        # Manifest
        print("Powerpoint to h5p Converter.")
        print(f"Author: {AUTHOR}")
        print(f"Version: {VERSION}, {YEAR}")
        print("Licence: BSD-2-Clause")
        print("Source code: https://github.com/MM-Lehmann/pptx2h5p")
        if len(sys.argv) != 2:
            print("Usage : python pptx2h5p.py [file]", file=sys.stderr)
            sys.exit(-1)

        # extract metadata
        filepath = os.path.abspath(sys.argv[1])
        if not os.path.exists(filepath):
            print(f"No such file: {filepath}", file=sys.stderr)
            sys.exit(-1)
        folder = os.path.dirname(filepath)
        filename = os.path.basename(filepath)
        title = os.path.splitext(filename)[0]

        # extract images
        print(f"extracting images from:\n\t {filepath}")
        ppt2image(filepath)
        image_folder = os.path.join(folder, title)
        images = natsorted(os.listdir(image_folder))

        # compile .hp5 file
        newfilename = os.path.splitext(filepath)[0] + ".h5p"
        print(f"building new file:\n\t {newfilename}")
        add_to_json(newfilename, image_folder, images, title)
        print("Converting successfully finished.")
        input(
            "Press Enter to delete temporary image (export) folder and close this window."
        )
        for image in images:
            os.remove(os.path.join(image_folder, image))
        os.rmdir(image_folder)
    except Exception as e:
        print(e, file=sys.stderr)
        input("An error has occured. Press Enter to close this window.")
