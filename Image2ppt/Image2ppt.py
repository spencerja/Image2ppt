import os, os.path
from pptx import Presentation
from pptx.util import Inches
import tkinter
from tkinter import filedialog
from tkinter import *
from tkinter import ttk
import io
from PIL import Image
import math
# Utilizes python-pptx: https://python-pptx.readthedocs.io/


def main():

    create_gui = CreateGUI()
    create_gui.construct_form()


class Components:

    def create_textbox(self,frame,text,row,column):
        txt = ttk.Entry(frame,width=40)
        txt.insert(tkinter.END,text)
        txt.grid(row=row,column=column)
        return txt

    def create_label(self,frame,text,row,column):
        label1 = ttk.Label(
            frame,
            text=text,
            padding=(5, 10))
        label1.grid(row=row, column=column)
        return label1

    def create_button(self,frame, text, row, column):
        button1 = ttk.Button(
            frame,
            text=text,
            )
        button1.grid(row=row, column=column)
        return button1

    def create_frame(self,root):
        frame = ttk.Frame(root, padding=40)
        frame.grid()
        return frame

class CreateGUI:
    def __init__(self):

        self.input_path_label = "initial path"
        self.output_path_label = "initial path"
        self.ppt_name_textbox = "test"
        self.path_list = [r"C:\Users\Fridge\Documents\PYGit\Image2ppt\Image2ppt\Input",
                          r"C:\Users\Fridge\Documents\PYGit\Image2ppt\Image2ppt\Output"]

    def construct_form(self):
        components = Components()
        root = Tk()
        root.minsize(width=500, height=300)
        root.title("Image2ppt")
        frame = components.create_frame(root)

        self.input_path_label = components.create_label(frame, self.path_list[0], 0, 0)

        input_path_button = components.create_button(frame, "Input path", 0, 1)
        input_path_button.bind("<ButtonPress>", lambda event: self.get_path(event, self.input_path_label))

        self.output_path_label = components.create_label(frame, self.path_list[1], 1, 0)

        output_path_button = components.create_button(frame,"Output path", 1, 1)
        output_path_button.bind("<ButtonPress>", lambda event: self.get_path(event, self.output_path_label))

        self.ppt_name_textbox = components.create_textbox(frame, "test", 2, 0)

        start_process_button = components.create_button(frame, "Start", 2, 1)
        start_process_button.bind("<ButtonPress>", lambda event: self.ppt_generation_process(event))

        self.start_gui(root)


    def get_path(self, event, arg):
        selected_path = filedialog.askdirectory()
        if not selected_path=="":
            arg.configure(text=selected_path)

    def ppt_generation_process(self, event):
        #tkinter.Tk().withdraw()
        path_list = [self.input_path_label.cget("text"), self.output_path_label.cget("text")]
        input_path = path_list[0]
        output_path = path_list[1]
        append_slide = AppendSlide(input_path)
        prs = append_slide.append_images_in_ppt()
        prs.save(output_path + '/' + self.ppt_name_textbox.get() + '.pptx')
        os.startfile(output_path + '/' + self.ppt_name_textbox.get() + '.pptx')

    def start_gui(self,root):
        root.mainloop()



class AppendSlide:
    # initial value loading
    def __init__(self, input_path):
        self.column = 4
        self.row = 2
        self.iter = self.row*self.column
        self.ppt_width = 13.333
        self.ppt_height = 7.5

        self.img_width = self.ppt_width/self.column
        self.img_height = self.ppt_height/self.row

        self.img_iter = self.column * self.row

        self.img_list = []
        self.img_count = 0

        self.input_path = input_path




    def get_images(self, input_path):
        folder_files = os.listdir(input_path)
        img_list = []

        for file in folder_files:
            if file.endswith('.png') or file.endswith(".tif") or file.endswith(".jpg") or file.endswith(".jpeg"):
                img_list.append(file)

        print(img_list)
        return img_list

    # generate ppt and add images to the ppt
    def append_images_in_ppt(self):
        prs = Presentation()
        prs.slide_width = Inches(self.ppt_width)
        prs.slide_height = Inches(self.ppt_height)
        prs = self.append_images(prs)
        return prs

    def append_images(self, prs):

        self.img_list = self.get_images(self.input_path)
        self.img_count = len(self.img_list)

        blank_slide = prs.slide_layouts[6]
        #in pixels
        pixel_width = int(960 / self.column)
        pixel_height = int(540 / self.row)
        #in inches
        width = self.ppt_width / self.column
        height = self.ppt_height / self.row
        dpi = 72

        for i in range(self.img_count):
            if i % self.img_iter == 0:
                image_slide = prs.slides.add_slide(blank_slide)

            current_img = Image.open(self.input_path + '/' + self.img_list[i])
            #working in pixels
            ratio = self.get_resize_ratio(current_img.width,current_img.height,pixel_width,pixel_height)
            resized_img = current_img.resize((int(current_img.width * ratio), int(current_img.height * ratio)))

            margin_width = Inches(width-resized_img.width/dpi)/2
            margin_height = Inches(height-resized_img.height/dpi)/2
            # in inches
            horizontal = Inches((i % self.column) * (self.ppt_width / self.column))
            vertical = Inches((i % self.img_iter // self.column) * self.ppt_height / self.row)

            if 0 <= i%self.iter and i%self.iter < self.column:
                #horizontal_position = horizontal+ margin_width
                vertical_position = vertical
            elif self.column*(self.row-1) <= i%self.iter and i%self.iter < self.column*self.row:
                #horizontal_position = horizontal + margin_width * 2
                vertical_position = vertical + margin_height*2
            else:
                #horizontal_position = horizontal + margin_width
                vertical_position = vertical + margin_height

            horizontal_position = horizontal + margin_width

            with io.BytesIO() as output:
                resized_img.save(output, format="GIF")
                image_slide.shapes.add_picture(output, horizontal_position, vertical_position)


        return prs

    #get ratio for both and use the smaller one to ensure that the image would fit in the slide panel
    def get_resize_ratio(self,img_width,img_height,pixel_width,pixel_height):
        ratio_width = pixel_width/img_width
        ratio_height = pixel_height/img_height
        return min(ratio_width,ratio_height)
        return ratio



if __name__ == "__main__":
    main()
