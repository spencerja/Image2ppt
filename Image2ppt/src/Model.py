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

class Model():
    # get ratio for both and use the smaller one to ensure that the image would fit in the slide panel
    def get_resize_ratio(self, img_width, img_height, pixel_width, pixel_height):
        ratio_width = pixel_width / img_width
        ratio_height = pixel_height / img_height
        return min(ratio_width, ratio_height)

    def get_margin(self, length, resized_length, dpi):
        return Inches(length - resized_length / dpi) / 2
