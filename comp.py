# required: use pip to install python-pptx, pywin32, opencv-python, pytesseract, pil, tesseract, tesseract-ocr, natsort
# resource used: https://github.com/sujiujiu/ppt_to_jpg/blob/master/ppt_to_jpg.py

from pptx import Presentation
import win32com
import win32com.client
import os
import sys
import cv2
import glob
import natsort
import argparse
import pytesseract
from difflib import SequenceMatcher
from datetime import datetime
import argparse

parser = argparse.ArgumentParser()
parser.add_argument('-video', '--video', help='Video file name', required=True)
parser.add_argument('-ppt', '--ppt', help='Powerpoint file name', required=True)
args = vars(parser.parse_args())
startTime = datetime.now()


pytesseract.pytesseract.tesseract_cmd = r'F:\Program Files\Tesseract-OCR\tesseract.exe'

filename = args['ppt']
titles = []
width = 0
height = 0


def get_text(img):
    return pytesseract.image_to_string(img)


def similar(a, b):
    return SequenceMatcher(None, a, b).ratio()


def extractImages(pathIn, pathOut):
    count = 0
    vidcap = cv2.VideoCapture(pathIn)
    success, image = vidcap.read()
    success = True
    width = vidcap.get(cv2.CAP_PROP_FRAME_WIDTH)
    height = vidcap.get(cv2.CAP_PROP_FRAME_HEIGHT)
    print("WxH" + str(width) + str(height))
    prev = None
    mid = int(width) / 2
    while success:
        vidcap.set(cv2.CAP_PROP_POS_MSEC, (count * 1000))    # added this line
        success, image = vidcap.read()
        print ('Read a new frame: ', success)
        if (success and count % 5 == 0):
            if (image[20, int(mid), 1] > 127):
                image = (255 - image)
            image = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)

            ret, image = cv2.threshold(image, 200, 255, 1)
            cv2.imwrite(pathOut + "\\%d.jpg" % count, image)     # save frame as JPEG file
            prev = image
        count = count + 1


def filename_to_time(name):
    time_str = name[:len(name) - 4]
    time_int = int(time_str)
    mins = time_int // 60
    seconds = time_int % 60
    time = str(mins) + ":" + str(seconds)
    return time


def get_text_from_image(filename):
    full_name = r'ss/' + str(filename)
    print(full_name)
    image = cv2.imread(full_name)
    text = get_text(image)
    text = text.replace('\n', ' ').replace('\x0c', '').replace('—', '–')
    return text


def get_inverted_text(filename):
    full_name = r'ss/' + str(filename)
    print(full_name)
    image = cv2.imread(full_name)
    image = (255 - image)
    text = get_text(image)
    text = text.replace('\n', ' ').replace('\x0c', '')
    return text


def clear_files():
    files = glob.glob('ss/*')
    for f in files:
        os.remove(f)


prs = Presentation(filename)


for slide in prs.slides:
    title = slide.shapes.title
    if title != None:
        title = slide.shapes.title.text
        title = title.replace("\x0b", "").strip()
        titles.append(title)

if titles == []:
    for slide in prs.slides:  # iterate over each slide
        title_shape = slide.shapes[0]  # consider the zeroth indexed shape as the title
        if title_shape.has_text_frame:  # is this shape has textframe attribute true then
            # check if the slide title already exists in the slide_title container
            if title_shape.text.strip(""" !@#$%^&*)(_-+=}{][:;<,>.?"'/<,""") not in titles:
                titles.append(title_shape.text.strip(""" !@#$%^&*)(_-+=}{][:;<,>.?"'/<,"""))

titles = list(dict.fromkeys(titles))

titles_length = len(titles)


clear_files()
if not os.path.exists('ss'):
    os.mkdir('ss')
extractImages(args['video'], r'ss')


files = [f for f in os.listdir('ss') if os.path.isfile(os.path.join('ss', f))]
sorted_files = natsort.natsorted(files)
print(sorted_files)


def remove_similar():
    prev = None
    for i in range(0, len(sorted_files)):
        if prev == None:
            prev = sorted_files[i]
        else:
            curr_name = r'ss/' + str(sorted_files[i])
            prev_name = r'ss/' + prev
            a = cv2.imread(curr_name)
            b = cv2.imread(prev_name)
            if (cv2.absdiff(a, b).mean() < 1):
                os.remove(curr_name)
            else:
                prev = sorted_files[i]


remove_similar()
files = [f for f in os.listdir('ss') if os.path.isfile(os.path.join('ss', f))]
sorted_files = natsort.natsorted(files)
print(sorted_files)

images_and_strings = []

for i in sorted_files:
    text = get_text_from_image(i)
    tup = (i, text)
    images_and_strings.append(tup)

result = []
found_list = []

for i in titles:
    found = False
    j = 0
    while not found and j < len(images_and_strings):
        if i.casefold() in images_and_strings[j][1].casefold():
            repeat = False
            for g in found_list:
                if g.casefold() in images_and_strings[j][1].casefold():
                    if (images_and_strings[j][1].casefold().index(g.casefold()) < images_and_strings[j][1].casefold().index(i.casefold())):
                        repeat = True
            if not repeat:
                found = True
                res_tup = (i, filename_to_time(images_and_strings[j][0]))
                result.append(res_tup)
                found_list.append(i)
            j += 1
        else:
            j += 1

print(result)

res_path = "result.txt"
if os.path.isfile(res_path):
    os.remove(res_path)

re = open(res_path, 'w')
for i in result:
    re.write(i[0] + ' - ' + i[1] + '\n')
re.close()

print("Time taken: " + str((datetime.now() - startTime)))
