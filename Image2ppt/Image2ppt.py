import os, os.path
from pptx import Presentation
from pptx.util import Inches
import tkinter
from tkinter import filedialog
#Utilizes python-pptx: https://python-pptx.readthedocs.io/

def main():
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


    for i in range(img_count):
        if(i%8 == 0):
            image_slide = prs.slides.add_slide(blank_slide)
        
        if(i%4 ==0):
            arg1 = left
        elif(i%4 ==1):
            arg1 = left2
        elif(i%4 ==2):
            arg1 = right
        elif(i%4 ==3):
            arg1 = right2

        if(i%8 < 4):
            arg2 = top
        elif(i%8 > 3):
            arg2 = bottom

        image_slide.shapes.add_picture(path+'/'+img[i], arg1, arg2, height=height)

    #specify savename
    prs.save(output_path+'/test.pptx')

if __name__ == "__main__":
    main()


