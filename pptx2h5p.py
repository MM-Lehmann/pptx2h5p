import sys
import os
from win32com import client
import json
from zipfile import ZipFile
from copy import deepcopy
import uuid
from natsort import natsorted

VERSION = '1.1'


def ppt2png(f):
    powerpoint = client.Dispatch('Powerpoint.Application')
    powerpoint.Presentations.Open(f)
    powerpoint.ActivePresentation.Export(f, 'PNG')
    powerpoint.ActivePresentation.Close()
    powerpoint.Quit()


def add_to_json(newfile, image_folder, images, title):
    exclude_files = ['content/content.json', 'h5p.json']
    if getattr(sys, 'frozen', False):
        basedir = sys._MEIPASS
    else:
        basedir = os.path.dirname(os.path.abspath(__file__))
    template = os.path.join(basedir, 'template.h5p')
    
    with ZipFile(template, 'r') as zin:
        with ZipFile(newfile, 'w') as zout:
            # copy all other files
            for item in zin.infolist():
                if item.filename not in exclude_files:
                    zout.writestr(item, zin.read(item.filename))

            # bilder dateinamen dem content.json zufügen
            with zin.open('content/content.json') as fp:
                content = json.load(fp)
                print(f"adding images from {image_folder}: {images}")
                for i, image in enumerate(images):
                    if i > 0 and len(content['presentation']['slides']) <= i:
                        content['presentation']['slides'].append(deepcopy(content['presentation']['slides'][0]))
                    content['presentation']['slides'][i]['elements'][0]['action']['params']['file']['path'] = \
                        'images/' + image
                    content['presentation']['slides'][i]['elements'][0]['action']['subContentId'] = str(uuid.uuid4())
            zout.writestr('content/content.json', json.dumps(content))

            # bilder dem zip zufügen
            for image in images:
                zout.write(os.path.join(image_folder, image), 'content/images/' + image)

            # titel ändern
            with zin.open('h5p.json', 'r') as h5p:
                content = json.load(h5p)
                content['title'] = title
            zout.writestr('h5p.json', json.dumps(content))


if __name__ == '__main__':
    try:
        print("Powerpoint to h5p Converter.")
        print(f"Version: {VERSION}")
        print("Martin Lehmann, 2020")
        print("Licence: BSD-2-Clause")
        print("Source code: https://github.com/MM-Lehmann/pptx2h5p")
        if len(sys.argv) != 2:
            print("Usage : python pptx2h5p.py [file]")
            sys.exit(-1)

        filepath = os.path.abspath(sys.argv[1])
        print(f"extracting images from {filepath}.")
        folder = os.path.dirname(filepath)
        filename = os.path.basename(filepath)
        title = os.path.splitext(filename)[0]
        if not os.path.exists(filepath):
            print("No such file!")
            sys.exit(-1)

        ppt2png(filepath)
        image_folder = os.path.join(folder, title)
        images = natsorted(os.listdir(image_folder))
        print("building new .h5p file")
        add_to_json(os.path.splitext(filepath)[0] + '.h5p', image_folder, images, title)
        print("Converting successfully finished.")
        input("Press Enter to delete temporary image (export) folder and close this window.")
        for image in images:
            os.remove(os.path.join(image_folder, image))
        os.rmdir(image_folder)
    except Exception as e:
        print(e)
        input("An error has occured. Press Enter to close this window.")
