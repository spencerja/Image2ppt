import os, os.path
from pptx import Presentation
from pptx.util import Inches
import tkinter
from tkinter import filedialog
#Utilizes python-pptx: https://python-pptx.readthedocs.io/

tkinter.Tk().withdraw()
#Input path then output path
path = filedialog.askdirectory()
print(path)
img = os.listdir(path)
img_count = len(img)
print(img_count)
output_path = filedialog.askdirectory()

#Generate ppt and specify slide types
prs = Presentation()
blank_slide =  prs.slide_layouts[6]

#Specify dimensions
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
height = Inches(3.33)
left = top = Inches(0)
left2 = Inches(3.33)
right = Inches(6.67)
right2 = Inches(10)
bottom = Inches(4.17)

#insertion loop
i = 0
while i//8 < img_count//8:
    image_slide = prs.slides.add_slide(blank_slide)
#add 4 image pairs to 4 corners. Note this is alphabetical insert top left to bottom right
    image_slide.shapes.add_picture(path+'/'+img[i+0], left, top, height=height)
    image_slide.shapes.add_picture(path+'/'+img[i+1], left2, top, height=height)
    image_slide.shapes.add_picture(path+'/'+img[i+2], right, top, height=height)
    image_slide.shapes.add_picture(path+'/'+img[i+3], right2, top, height=height)
    image_slide.shapes.add_picture(path+'/'+img[i+4], left, bottom, height=height)
    image_slide.shapes.add_picture(path+'/'+img[i+5], left2, bottom, height=height)
    image_slide.shapes.add_picture(path+'/'+img[i+6], right, bottom, height=height)
    image_slide.shapes.add_picture(path+'/'+img[i+7], right2, bottom, height=height)
    i+=8


#specify savename
prs.save(output_path+'/test.pptx')




