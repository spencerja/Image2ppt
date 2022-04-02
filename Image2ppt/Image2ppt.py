import os, os.path
from pptx import Presentation
from pptx.util import Inches
import tkinter
from tkinter import filedialog
#Utilizes python-pptx: https://python-pptx.readthedocs.io/

def main():
    
    append_ppt = AppendPPT()
    append_ppt.whole_process()


class AppendPPT():

    def whole_process(self):
        tkinter.Tk().withdraw()
        #Input path then output path
        input_path = filedialog.askdirectory()
        print(input_path)
        img = os.listdir(input_path)
        img_count = len(img)

        print(img_count)

        output_path = filedialog.askdirectory()

        append_slide = AppendSlide()

        prs = append_slide.append_images_in_ppt(img_count,input_path,img)

        #specify savename
        prs.save(output_path+'/test.pptx')

class AppendSlide():

    def __init__(self):
        self.column = 4
        self.row = 2
        self.ppt_width = 13.333
        self.ppt_height = 7.5
        self.img_iter = 8

    def append_images_in_ppt(self,img_count,input_path,img):
        #Generate ppt and specify slide types
        prs = Presentation()
        blank_slide =  prs.slide_layouts[6]
        prs.slide_width = Inches(self.ppt_width)
        prs.slide_height = Inches(self.ppt_height)      

        for i in range(img_count):
            if(i%self.img_iter == 0):
                image_slide = prs.slides.add_slide(blank_slide)
            arg1 = Inches((i%self.column)*(self.ppt_width/self.column))
            arg2 = Inches((i%self.img_iter//self.column)*self.ppt_height/self.row)
            image_slide.shapes.add_picture(input_path+'/'+img[i], arg1, arg2, height=Inches(self.ppt_width/self.column))

        return prs
    

if __name__ == "__main__":
    main()


