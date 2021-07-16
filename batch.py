import glob, os
from datetime import datetime

startTime = datetime.now()

vid_list = []
for file in glob.glob("*.mp4"):
    vid_list.append(file)

ppt_list = []
for file in glob.glob("*.pptx"):
    ppt_list.append(file)
print(vid_list)
print(ppt_list)

if len(vid_list) != len(ppt_list):
    print("Mismatch in number of pptx and mp4 files")
    exit()
for j in range(0, len(vid_list)):
    video = vid_list[j]
    ppt = ""
    try:
        ppt = [item for item in ppt_list if item.split(".")[0] == video.split('.')[0]][0]
    except IndexError:
        print("Matching powerpoint file not found for " + video)
        exit()
    os.system('python comp.py -video ' + video + ' -ppt ' + ppt)

print("Total batch time taken: " + str((datetime.now() - startTime)))
