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


    def get_path(self):
        selected_path = filedialog.askdirectory()
        return selected_path

    def whole_process(self):
        tkinter.Tk().withdraw()

        input_path = self.get_path()
        

        output_path = self.get_path()

        append_slide = AppendSlide(input_path)

        prs = append_slide.append_images_in_ppt()

        prs.save(output_path+'/test.pptx')

class AppendSlide():

    #initial value loading
    def __init__(self,input_path):
        self.column = 4
        self.row = 2
        self.ppt_width = 13.333
        self.ppt_height = 7.5
        self.img_iter = 8
        self.img = self.get_images(input_path)
        self.img_count = len(self.img)
        self.input_path = input_path
        

    def get_images(self,input_path):
        img = os.listdir(input_path)
        return img

    #generate ppt and add images to the ppt
    def append_images_in_ppt(self):        
        prs = Presentation()
        prs.slide_width = Inches(self.ppt_width)
        prs.slide_height = Inches(self.ppt_height)      
        prs = self.append_images(prs)
        return prs

    def append_images(self,prs):
        blank_slide =  prs.slide_layouts[6]

        for i in range(self.img_count):
            if(i%self.img_iter == 0):                
                image_slide = prs.slides.add_slide(blank_slide)

            arg1 = Inches((i%self.column)*(self.ppt_width/self.column))
            arg2 = Inches((i%self.img_iter//self.column)*self.ppt_height/self.row)

            image_slide.shapes.add_picture(self.input_path+'/'+self.img[i], arg1, arg2, height=Inches(self.ppt_width/self.column))

        return prs
    

if __name__ == "__main__":
    main()


