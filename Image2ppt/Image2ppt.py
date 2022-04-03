import os, os.path
from pptx import Presentation
from pptx.util import Inches
import tkinter
from tkinter import filedialog
from tkinter import *
from tkinter import ttk

# Utilizes python-pptx: https://python-pptx.readthedocs.io/

def main():

    create_gui = CreateGUI()
    create_gui.construct_form()

class CreateGUI:
    def __init__(self):
        self.root = Tk()
        self.root.minsize(width=500, height=300)
        self.root.title("Image2ppt")
        self.frame = None
        self.input_path_label = None
        self.ppt_name_textbox = None
        self.path_list = [r"C:\Users\Fridge\Documents\PYGit\Image2ppt\Image2ppt\Input",
                          r"C:\Users\Fridge\Documents\PYGit\Image2ppt\Image2ppt\Output"]

    def construct_form(self):
        self.create_widget()
        self.input_path_label = self.create_label(self.path_list[0], 0, 0)
        input_path_button = self.create_button("Input path",0,1)
        input_path_button.bind("<ButtonPress>", lambda event: self.get_path(event, self.input_path_label))

        self.output_path_label = self.create_label(self.path_list[1], 1, 0)
        output_path_button = self.create_button("Output path", 1, 1)
        output_path_button.bind("<ButtonPress>", lambda event: self.get_path(event, self.output_path_label))

        self.ppt_name_textbox = self.create_textbox("test", 2, 0)

        start_process_button = self.create_button("Start",2,1)
        start_process_button.bind("<ButtonPress>", lambda event: self.whole_process(event))

        self.start_widget()

    def create_widget(self):
        self.frame = ttk.Frame(self.root, padding=40)
        self.frame.grid()

    def update_path_list(self):
        self.path_list = [self.input_path_label.cget("text"), self.output_path_label.cget("text")]

    def create_textbox(self,text,row,column):
        txt = ttk.Entry(self.frame,width=40)
        txt.insert(tkinter.END,text)
        txt.grid(row=row,column=column)
        return txt

    def create_label(self,text,row,column):
        label1 = ttk.Label(
            self.frame,
            text=text,
            padding=(5, 10))
        label1.grid(row=row, column=column)
        return label1

    def create_button(self,text,row,column):
        button1 = ttk.Button(
            self.frame,
            text=text,
            )
        button1.grid(row=row, column=column)
        return button1


    def get_path(self, event,arg):
        selected_path = filedialog.askdirectory()
        if not selected_path=="":
            arg.configure(text=selected_path)

    def whole_process(self,event):
        tkinter.Tk().withdraw()
        path_list = [self.input_path_label.cget("text"), self.output_path_label.cget("text")]
        input_path = path_list[0]
        output_path = path_list[1]
        append_slide = AppendSlide(input_path)
        prs = append_slide.append_images_in_ppt()
        prs.save(output_path + '/' + self.ppt_name_textbox.get() + '.pptx')
        os.startfile(output_path + '/' + self.ppt_name_textbox.get() + '.pptx')

    def start_widget(self):
        self.root.mainloop()


#this is not used in production
class AddTest:
    def add_test(self,a,b):
        return a+b


class AppendSlide:
    # initial value loading
    def __init__(self, input_path):
        self.column = 4
        self.row = 2

        self.ppt_width = 13.333
        self.ppt_height = 7.5

        self.img_iter = self.column * self.row

        self.img_list = self.get_images(input_path)
        self.img_count = len(self.img_list)

        self.input_path = input_path

    def get_images(self, input_path):
        img_list = os.listdir(input_path)
        return img_list

    # generate ppt and add images to the ppt
    def append_images_in_ppt(self):
        prs = Presentation()
        prs.slide_width = Inches(self.ppt_width)
        prs.slide_height = Inches(self.ppt_height)
        prs = self.append_images(prs)
        return prs

    def append_images(self, prs):
        blank_slide = prs.slide_layouts[6]

        for i in range(self.img_count):
            if i % self.img_iter == 0:
                image_slide = prs.slides.add_slide(blank_slide)

            horizontal = Inches((i % self.column) * (self.ppt_width / self.column))
            vertical = Inches((i % self.img_iter // self.column) * self.ppt_height / self.row)

            image_slide.shapes.add_picture(self.input_path + '/' + self.img_list[i], horizontal, vertical,
                                           height=Inches(self.ppt_width / self.column))

        return prs



if __name__ == "__main__":
    main()
