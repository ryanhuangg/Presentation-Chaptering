# Presentation-Chaptering
Automatically creates timestamps for presentation videos by using OCR to identify text within the video that match powerpoint titles.

# Requirements:
Download and install Tesseract from here:
https://github.com/UB-Mannheim/tesseract/wiki

After installation, note down the path and change the tesseract.exe path in the comp.py file at line 25.

Use pip to install the following packages:
python-pptx, pywin32, opencv-python, pytesseract, pil, tesseract, tesseract-ocr, natsort
OR
pip install -r requirements.txt

# How to run:
Place the powerpoint file and the video recording file in the same folder as the python script. Then, run the following command from powershell or command line:

python comp.py -video [video-filename] -ppt [ppt-filename]

For example:
python comp.py -video sample.mp4 -ppt presentation.pptx
