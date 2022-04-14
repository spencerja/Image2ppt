import os, os.path
from pptx import Presentation
from pptx.util import Inches, Pt
import tkinter
from tkinter import filedialog
from tkinter import *
from tkinter import ttk
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
import io
from PIL import Image
from math import ceil
import math
import time


class View():
    def __init__(self, root):
        #self.path_list = [r"C:\Users\Fridge\Documents\PYGit\Image2ppt\Image2ppt\Input",
       #                   r"C:\Users\Fridge\Documents\PYGit\Image2ppt\Image2ppt\Output"]
        self.path_list = ["Please choose an input folder","Please choose an output folder"]
        self.components = Components()
        self.frame = self.components.create_frame(root)
        self.notebook = self.components.create_notebook(root)
        self.prepare_general_tab()
        self.prepare_advanced_tab()

    def prepare_general_tab(self):
        self.general_frame = self.components.create_tab(self.notebook, "General")
        self.input_path_label = self.components.create_label(self.general_frame, self.path_list[0], 0, 0)
        self.input_path_button = self.components.create_button(self.general_frame, "Input path", 0, 1)
        self.output_path_label = self.components.create_label(self.general_frame, self.path_list[1], 1, 0)
        self.output_path_button = self.components.create_button(self.general_frame, "Output path", 1, 1)
        self.gui_column_desc = self.components.create_label(self.general_frame, "Column Number:", 2, 0)
        self.gui_column = self.components.create_textbox(self.general_frame, 4, 2, 1)
        self.gui_row_desc = self.components.create_label(self.general_frame, "Row Number:", 3, 0)
        self.gui_row = self.components.create_textbox(self.general_frame, 2, 3, 1)
        self.ppt_name_label = self.components.create_label(self.general_frame, "Save Name:", 7, 0)
        self.gui_ppt_name_textbox = self.components.create_textbox(self.general_frame, "test", 7, 1)
        self.combobox = self.components.create_combobox(self.general_frame,['Ascending','Descending'],8,0)
        self.start_process_button = self.components.create_button(self.general_frame, "Start", 8, 1)
        self.save_config_button  = self.components.create_button(self.general_frame,"Save Config",8,2)

    def prepare_advanced_tab(self):
        self.advanced_frame = self.components.create_tab(self.notebook, "Advanced")
        self.gui_ppt_width_desc = self.components.create_label(self.advanced_frame, "Slide Width (inches):", 4, 0)
        self.gui_ppt_width = self.components.create_textbox(self.advanced_frame, 26.6666, 4, 1)
        self.gui_ppt_height_desc = self.components.create_label(self.advanced_frame, "Slide Height (inches):", 5, 0)
        self.gui_ppt_height = self.components.create_textbox(self.advanced_frame, 15, 5, 1)
        self.gui_slide_counter_desc = self.components.create_label(self.advanced_frame, "Images for each cell:", 6, 0)
        self.gui_slide_counter = self.components.create_textbox(self.advanced_frame, 16, 6, 1)


class Components:

    def create_combobox(self,frame,list,row,column):
        combobox = ttk.Combobox(frame,values = list)
        combobox.set(list[0])
        combobox.grid(row =row,column=column)
        return combobox

    def create_textbox(self,frame,text,row,column):
        txt = ttk.Entry(frame,width=20)
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

    def create_notebook(self, root):
        notebook = ttk.Notebook(root)
        notebook.grid()
        return notebook

    def create_tab(self, notebook,text):
        frame = ttk.Frame(notebook)
        notebook.add(frame, text=text)
        return frame