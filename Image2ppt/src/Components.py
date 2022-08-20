import tkinter
from tkinter import *
from tkinter import ttk


class Components:

    def __init__(self):
        pass

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

    def create_textbox(self,frame,text,row,column,width):
        """
        :param frame:
        :param text:
        :param row:
        :param column:
        :return:
        """
        txt = ttk.Entry(frame,width=width)
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
            padding=(5, 5))
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
        frame = ttk.Frame(root, padding=0)
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
        checkbox = Checkbutton(frame, text=text, variable=toggle_label, command=toggle_label)
        checkbox.grid(row =row,column=column)
        return toggle_label