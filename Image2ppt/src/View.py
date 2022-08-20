import Components


class View:
    """View class.
    """
    def __init__(self, root):
        """Initialize the class instance

        Args:
            root (Tk()): main window

        """
        self.path_list = ["Please choose an input folder","Please choose an output folder"]
        self.components = Components.Components()
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
        self.gui_ppt_name_textbox = self.components.create_textbox(self.general_frame, "test", 7, 1, 11)

    def prepare_advanced_tab(self):
        """

        :return:
        """
        self.advanced_frame = self.components.create_tab(self.notebook, "Advanced")
        self.gui_ppt_width_desc = self.components.create_label(self.advanced_frame, "Slide Width (inches):", 0, 2)
        self.gui_ppt_width = self.components.create_textbox(self.advanced_frame, 26.6666, 0, 3, 5)
        self.gui_ppt_height_desc = self.components.create_label(self.advanced_frame, "Slide Height (inches):", 1, 2)
        self.gui_ppt_height = self.components.create_textbox(self.advanced_frame, 15, 1, 3, 5)
        self.gui_cell_image_total_desc = self.components.create_label(self.advanced_frame, "Images for each cell:", 2, 0)
        self.gui_cell_image_total = self.components.create_textbox(self.advanced_frame, 16, 2, 1, 5)
        self.gui_column_desc = self.components.create_label(self.advanced_frame, "Column Number:", 0, 0)
        self.gui_column = self.components.create_textbox(self.advanced_frame, 4, 0, 1, 5)
        self.gui_row_desc = self.components.create_label(self.advanced_frame, "Row Number:", 1, 0)
        self.gui_row = self.components.create_textbox(self.advanced_frame, 2, 1, 1, 5)
        self.combo_label = self.components.create_label(self.advanced_frame, "Sorting:", 2, 2)
        self.combobox = self.components.create_combobox(self.advanced_frame, ['Alphabetical A-Z','Alphabetical Z-A', "Oldest-Newest","Newest-Oldest"], 2, 3)
        self.label_checkbox = self.components.create_checkbox(self.advanced_frame, "Label Slides?", 3, 0)


