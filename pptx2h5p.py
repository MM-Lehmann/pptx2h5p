import json
import os
import sys
import uuid
from copy import deepcopy
from zipfile import ZipFile

from natsort import natsorted
from win32com import client
from get_image_size import get_image_size

VERSION = "1.2"
YEAR = "2021"
target_ratio = 2  # target aspect ratio for slides in h5p


def ppt2png(file):
    try:
        powerpoint = client.Dispatch("Powerpoint.Application")
    except Exception as e:
        print("Powerpoint could not be opened", file=sys.stderr)
        raise e
    try: # look if an active presentation is open
        assert powerpoint.ActivePresentation is not None
        QUIT = False # don't quit powerpoint later
    except:
        QUIT = True # quit powerpoint later
    ppt = powerpoint.Presentations.Open(file)
    ppt.Export(file, "PNG")
    ppt.Close()
    if QUIT: # quit only if required
        powerpoint.Quit()


def add_to_json(newfile, image_folder, images, title):
    exclude_files = ['content/content.json', 'h5p.json']
    if getattr(sys, 'frozen', False):
        basedir = sys._MEIPASS
    else:
        basedir = os.path.dirname(os.path.abspath(__file__))
    template = os.path.join(basedir, "template.h5p")
    
    img_width, img_height = get_image_size(os.path.join(image_folder, images[0]))

    with ZipFile(template, "r") as zin:
        with ZipFile(newfile, "w") as zout:
            # copy all other files
            for item in zin.infolist():
                if item.filename not in exclude_files:
                    zout.writestr(item, zin.read(item.filename))

            # add image filenames to content.json
            with zin.open("content/content.json") as fp:
                content = json.load(fp)
                print(f"adding images from {image_folder}: {images}")
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
            with zin.open("h5p.json", "r") as h5p:
                content = json.load(h5p)
                content["title"] = title
            zout.writestr("h5p.json", json.dumps(content))


if __name__ == "__main__":
    try:
        # Manifest
        print("Powerpoint to h5p Converter.")
        print(f"Version: {VERSION}")
        print(f"Martin Lehmann, {YEAR}")
        print("Licence: BSD-2-Clause")
        print("Source code: https://github.com/MM-Lehmann/pptx2h5p")
        if len(sys.argv) != 2:
            print("Usage : python pptx2h5p.py [file]", file=sys.stderr)
            sys.exit(-1)

        # extract metadata
        filepath = os.path.abspath(sys.argv[1])
        print(f"extracting images from {filepath}.")
        folder = os.path.dirname(filepath)
        filename = os.path.basename(filepath)
        title = os.path.splitext(filename)[0]
        if not os.path.exists(filepath):
            print("No such file!", file=sys.stderr)
            sys.exit(-1)

        # extract images
        ppt2png(filepath)
        image_folder = os.path.join(folder, title)
        images = natsorted(os.listdir(image_folder))

        # compile .hp5 file
        newfilename = os.path.splitext(filepath)[0] + ".h5p"
        print(f"building new {newfilename} file")
        add_to_json(
            newfilename, image_folder, images, title
        )
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
