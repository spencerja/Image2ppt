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
    """View class.
    """
    def __init__(self, root):
        """Initialize the class instance

        Args:
            root (Tk()): main window

        """
        self.path_list = ["Please choose an input folder","Please choose an output folder"]
        self.components = Components()
        self.frame = self.components.create_frame(root)
        self.notebook = self.components.create_notebook(root)
        self.prepare_general_tab()
        self.prepare_advanced_tab()
        self.start_process_button = self.components.create_button(self.frame, "Start", 1, 1)
        self.save_config_button = self.components.create_button(self.frame, "Save Config", 1, 2)

    def prepare_general_tab(self):
        """

        :return:
        """
        self.general_frame = self.components.create_tab(self.notebook, "General")
        self.input_path_label = self.components.create_label(self.general_frame, self.path_list[0], 0, 0)
        self.input_path_button = self.components.create_button(self.general_frame, "Input path", 0, 1)
        self.output_path_label = self.components.create_label(self.general_frame, self.path_list[1], 1, 0)
        self.output_path_button = self.components.create_button(self.general_frame, "Output path", 1, 1)

        self.ppt_name_label = self.components.create_label(self.general_frame, "Save Name:", 7, 0)
        self.gui_ppt_name_textbox = self.components.create_textbox(self.general_frame, "test", 7, 1)
        #self.start_process_button = self.components.create_button(self.general_frame, "Start", 8, 1)
        #self.save_config_button  = self.components.create_button(self.general_frame,"Save Config",8,2)

    def prepare_advanced_tab(self):
        """

        :return:
        """
        self.advanced_frame = self.components.create_tab(self.notebook, "Advanced")
        self.gui_ppt_width_desc = self.components.create_label(self.advanced_frame, "Slide Width (inches):", 4, 0)
        self.gui_ppt_width = self.components.create_textbox(self.advanced_frame, 26.6666, 4, 1)
        self.gui_ppt_height_desc = self.components.create_label(self.advanced_frame, "Slide Height (inches):", 5, 0)
        self.gui_ppt_height = self.components.create_textbox(self.advanced_frame, 15, 5, 1)
        self.gui_cell_image_total_desc = self.components.create_label(self.advanced_frame, "Images for each cell:", 6, 0)
        self.gui_cell_image_total = self.components.create_textbox(self.advanced_frame, 16, 6, 1)
        self.gui_column_desc = self.components.create_label(self.advanced_frame, "Column Number:", 2, 0)
        self.gui_column = self.components.create_textbox(self.advanced_frame, 4, 2, 1)
        self.gui_row_desc = self.components.create_label(self.advanced_frame, "Row Number:", 3, 0)
        self.gui_row = self.components.create_textbox(self.advanced_frame, 2, 3, 1)
        self.combo_label = self.components.create_label(self.advanced_frame, "Sorting:", 8, 0)
        self.combobox = self.components.create_combobox(self.advanced_frame, ['Alphabetical A-Z','Alphabetical Z-A', "Oldest-Newest","Newest-Oldest"], 8, 1)
        self.label_checkbox = self.components.create_checkbox(self.advanced_frame, "Label Slides?", 1, 4)


class Components:
    """Components
    Tkinter GUI component methods are stored in this class.

    """
    def create_combobox(self,frame,list,row,column):
        """

        :param frame:
        :param list:
        :param row: vertical grid location
        :param column: horizontal grid location
        :return: combobox
        """
        combobox = ttk.Combobox(frame,values = list)
        combobox.set(list[0])
        combobox.grid(row =row,column=column)
        return combobox

    def create_textbox(self,frame,text,row,column):
        """

        :param frame:
        :param text:
        :param row:
        :param column:
        :return:
        """
        txt = ttk.Entry(frame,width=20)
        txt.insert(tkinter.END,text)
        txt.grid(row=row,column=column)
        return txt

    def create_label(self,frame,text,row,column):
        """

        :param frame:
        :param text:
        :param row:
        :param column:
        :return:
        """
        label1 = ttk.Label(
            frame,
            text=text,
            padding=(5, 10))
        label1.grid(row=row, column=column)
        return label1

    def create_button(self,frame, text, row, column):
        """

        :param frame:
        :param text:
        :param row:
        :param column:
        :return:
        """
        button1 = ttk.Button(
            frame,
            text=text,
            )
        button1.grid(row=row, column=column)
        return button1

    def create_frame(self,root):
        """

        :param root:
        :return:
        """
        frame = ttk.Frame(root, padding=40)
        frame.grid(row=1,column=1)
        return frame

    def create_notebook(self, root):
        """

        :param root:
        :return:
        """
        notebook = ttk.Notebook(root)
        notebook.grid(row=0,column=0,columnspan = 2)
        return notebook

    def create_tab(self, notebook,text):
        """

        :param notebook:
        :param text:
        :return:
        """
        frame = ttk.Frame(notebook)
        notebook.add(frame, text=text)
        return frame

    def create_checkbox(self, frame, text, row, column):
        toggle_label = BooleanVar()
        toggle_label.set(True)
        checkbox = Checkbutton(frame, text=text, variable=toggle_label, command = toggle_label)
        checkbox.grid(row =row,column=column)
        return toggle_label