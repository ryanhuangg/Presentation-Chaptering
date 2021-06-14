# required: use pip to install python-pptx, pywin32, opencv-python, pytesseract, tesseract, tesseract-ocr, natsort
# resource used: https://github.com/sujiujiu/ppt_to_jpg/blob/master/ppt_to_jpg.py

from pptx import Presentation
import win32com
import win32com.client
import os
import sys
import cv2
import natsort
import argparse
import pytesseract
from PIL import Image
from difflib import SequenceMatcher
from datetime import datetime
startTime = datetime.now()


pytesseract.pytesseract.tesseract_cmd = r'F:\Program Files\Tesseract-OCR\tesseract.exe'

filename = "F://Comparison/phantomgv.pptx"
folder = "F://Comparison"
titles = []


def get_text(img):
    return pytesseract.image_to_string(img)


def similar(a, b):
    return SequenceMatcher(None, a, b).ratio()


def extractImages(pathIn, pathOut):
    count = 0
    vidcap = cv2.VideoCapture(pathIn)
    success, image = vidcap.read()
    success = True
    while success:
        vidcap.set(cv2.CAP_PROP_POS_MSEC, (count * 1000))    # added this line
        success, image = vidcap.read()
        print ('Read a new frame: ', success)
        if (count % 5 == 0):
            image = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
            cv2.imwrite(pathOut + "\\%d.jpg" % count, image)     # save frame as JPEG file
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
    text = text.replace('\n', ' ').replace('\x0c', '')
    return text


def get_inverted_text(filename):
    full_name = r'ss/' + str(filename)
    print(full_name)
    image = cv2.imread(full_name)
    image = (255 - image)
    text = get_text(image)
    text = text.replace('\n', ' ').replace('\x0c', '')
    return text


prs = Presentation(filename)


for slide in prs.slides:
    title = slide.shapes.title
    if title != None:
        title = slide.shapes.title.text
        titles.append(title)

titles = list(dict.fromkeys(titles))

titles_length = len(titles)


os.system('py convert.py')


extractImages('recording.mp4', r'F:\Comparison\ss')

files = [f for f in os.listdir('F:\\Comparison\ss') if os.path.isfile(os.path.join('F:\\Comparison\ss', f))]
sorted_files = natsort.natsorted(files)
print(sorted_files)

images_and_strings = []

for i in sorted_files:
    tup = (i, get_text_from_image(i) + get_inverted_text(i))
    images_and_strings.append(tup)

result = []

for i in titles:
    found = False
    j = 0
    while not found and j < len(images_and_strings):
        if i in images_and_strings[j][1]:
            found = True
            res_tup = (i, filename_to_time(images_and_strings[j][0]))
            result.append(res_tup)
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
