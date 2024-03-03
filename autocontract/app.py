import configparser
import os
import webbrowser
from tkinter import filedialog, END, NORMAL, DISABLED
import pandas as pd
from docx import Document
from docx2pdf import convert
import gettext

import customtkinter
from CTkMenuBar import CTkMenuBar, CustomDropdownMenu
from CTkToolTip import CTkToolTip
from CTkXYFrame import CTkXYFrame
from CTkMessagebox import CTkMessagebox
from PIL import Image

# CONSTANTS
SEPARATOR_BACKGROUND_COLOR = "grey90", "black"

# Default options
options = {
    "themes": ["system", "light", "dark"],
    "default_theme": "system",
    "sounds": ["True", "False"],
    "default_sounds": "True",
    "languages": ["english", "portuguese"],
    "default_language": "english",
}

# Import _ function to translate
_ = gettext.gettext


class ConfigHolder:
    config_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.ini")
    config = configparser.ConfigParser()


class AutoContract(customtkinter.CTk, ConfigHolder):
    APP_WIDTH = 800
    APP_HEIGHT = 600

    TEMPLATE_FRAME_HEIGHT = 75
    DATA_CONTROLS_FRAME_HEIGHT = 75
    DATA_ADD_BUTTON_SIZE = 30
    DESTINATION_FRAME_HEIGHT = 75

    START_ENTRY_ROW_NUM = 4
    START_ENTRY_COL_NUM = 2

    START_ROW_SPAN = START_ENTRY_ROW_NUM

    # Plus 2 because the labels column
    START_COL_SPAN = START_ENTRY_COL_NUM + 1

    INFO_BUTTON_ICON_SIZE = (16, 16)
    ADD_ENTRY_ICON_SIZE = (16, 16)

    # Colors first correspond to Light Mode and second to Dark Mode
    MENUBAR_BACKGROUND_COLOR = "white", "black"
    DROPDOWN_BACKGROUND_COLOR = "white", "grey20"
    FRAMES_BACKGROUND_COLOR = "transparent"

    TOOLTIP_BORDER_COLOR = "black", "white"

    HOVER_COLOR = "grey85", "grey25"

    NO_FOLDER_SELECTED = _("No folder selected")
    NO_FILE_SELECTED = _("No file selected")

    DATA_INFO_MESSAGE = _(
        "The names of the variables (in the table) must be equal to the word to be replaced in the template. Is not mandatory to import a data file. If you still want to import, make sure the first the row is for variables and the rest for the data"
    )

    INFO_BUTTON_LIGHT_PATH = os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        "assets/icons/info-black.png",
    )
    INFO_BUTTON_DARK_PATH = os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        "assets/icons/info-white.png",
    )

    ADD_ENTRY_LIGHT_PATH = os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        "assets/icons/add-black.png",
    )
    ADD_ENTRY_DARK_PATH = os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        "assets/icons/add-white.png",
    )

    SOURCE_CODE_URL = "https://github.com/P3rdigas/AutoContract"

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        # Configure window
        self.title("AutoContract")
        self.iconbitmap(
            os.path.join(
                os.path.dirname(os.path.abspath(__file__)),
                "assets/logos/logo128x128.ico",
            )
        )
        self.geometry(f"{self.APP_WIDTH}x{self.APP_HEIGHT}")
        self.resizable(width=False, height=False)

        # Loading config file
        self.config_file = ConfigHolder.config_file
        self.config = ConfigHolder.config
        self.app_sounds = False
        self.load_configuration()

        # Set Menubar
        # TODO: Finish MenuBar
        toolbar = CTkMenuBar(master=self, bg_color=self.MENUBAR_BACKGROUND_COLOR)
        file_button = toolbar.add_cascade(text=_("File"), hover_color=self.HOVER_COLOR)

        self.settings_window = None

        toolbar.add_cascade(
            text=_("Settings"), hover_color=self.HOVER_COLOR, command=self.open_settings
        )
        about_button = toolbar.add_cascade(
            text=_("About"), hover_color=self.HOVER_COLOR
        )

        file_button_dropdown = CustomDropdownMenu(
            widget=file_button,
            corner_radius=0,
            bg_color=self.DROPDOWN_BACKGROUND_COLOR,
            hover_color=self.HOVER_COLOR,
        )
        file_button_dropdown.add_option(option=_("Import"), corner_radius=0)

        clear_data = file_button_dropdown.add_submenu(
            submenu_name=_("Clear"), corner_radius=0
        )
        clear_data.add_option(
            option=_("Template"), command=self.reset_template_file, corner_radius=0
        )
        clear_data.add_option(
            option=_("Data"), command=self.reset_data_content, corner_radius=0
        )
        clear_data.add_option(
            option=_("Folder"), command=self.reset_destination_folder, corner_radius=0
        )
        clear_data.add_option(option=_("All"), command=self.reset_all, corner_radius=0)

        file_button_dropdown.add_option(
            option=_("Generate"), command=self.generate_contracts, corner_radius=0
        )
        file_button_dropdown.add_option(
            option=_("Exit"), command=self.destroy, corner_radius=0
        )

        github_icon_light_path = os.path.join(
            os.path.dirname(os.path.abspath(__file__)), "assets/icons/github-mark.png"
        )

        github_icon_dark_path = os.path.join(
            os.path.dirname(os.path.abspath(__file__)),
            "assets/icons/github-mark-white.png",
        )

        github_image = customtkinter.CTkImage(
            light_image=Image.open(github_icon_light_path),
            dark_image=Image.open(github_icon_dark_path),
        )

        about_button_dropdown = CustomDropdownMenu(
            widget=about_button,
            corner_radius=0,
            bg_color=self.DROPDOWN_BACKGROUND_COLOR,
            hover_color=self.HOVER_COLOR,
        )
        about_button_dropdown.add_option(
            option="Source Code",
            command=self.open_browser,
            image=github_image,
            corner_radius=0,
        )

        # Layout (Template Frame(100px), Data Frame(what's left), Destination Frame(100px with generate button))
        # Template Frame
        template_frame = customtkinter.CTkFrame(
            self,
            height=self.TEMPLATE_FRAME_HEIGHT,
            corner_radius=0,
            fg_color=self.FRAMES_BACKGROUND_COLOR,
        )

        template_label_frame = customtkinter.CTkFrame(
            template_frame, corner_radius=0, fg_color=self.FRAMES_BACKGROUND_COLOR
        )

        # TODO: Should be?
        # Unicode symbols: http://xahlee.info/comp/unicode_circled_numbers.html
        template_label = customtkinter.CTkLabel(
            template_label_frame, text=_("① Template")
        )

        template_controls_frame = customtkinter.CTkFrame(
            template_frame, corner_radius=0, fg_color=self.FRAMES_BACKGROUND_COLOR
        )

        template_file_label = customtkinter.CTkLabel(
            template_controls_frame, text=_("Template file:")
        )

        self.template_file = None

        # FIXME: If translate doesn't work because of the text being a const
        # Default width of the label (can provide font),
        self.TEMPLATE_FILENAME_WIDTH = self.get_width_text(self.NO_FILE_SELECTED)

        # FIXME: If translate doesn't work because of the text being a const
        # TODO: Should be?
        # Width: plus 1 (if width of label equals the width of the text the text you jump out by 1 px) and plus 10 (margin, 5 to the left, 5 to the right)
        self.template_filename_label = customtkinter.CTkLabel(
            template_controls_frame,
            width=self.TEMPLATE_FILENAME_WIDTH + 1 + 10,
            text=self.NO_FILE_SELECTED,
        )

        self.template_filename_label_tooltip = CTkToolTip(
            self.template_filename_label,
            message=None,
            corner_radius=5,
            border_width=1,
            border_color=self.TOOLTIP_BORDER_COLOR,
        )

        # If the text is smaller than the label width no need to show a tooltip
        self.template_filename_label_tooltip.hide()

        self.template_filename_label.bind(
            "<Enter>",
            lambda event: self.on_enter(event, self.template_filename_label_tooltip),
        )

        # TODO: Set a better color
        # Width to 1 so the button takes the width of the text
        template_file_button = customtkinter.CTkButton(
            template_controls_frame,
            width=1,
            text=_("Choose file"),
            # fg_color="transparent",
            hover_color=self.HOVER_COLOR,
            command=self.choose_template_file,
        )

        # TODO: Set a better color
        # Width to 1 so the button takes the width of the text
        template_file_clear_button = customtkinter.CTkButton(
            template_controls_frame,
            width=1,
            text=_("Reset"),
            # fg_color="transparent",
            hover_color=self.HOVER_COLOR,
            command=self.reset_template_file,
        )

        # TODO: Add button for a tooltip to explain what a template should have

        # Create Separator for Template and Data Frames
        separator_t_d = customtkinter.CTkFrame(
            self,
            corner_radius=0,
            height=1,
            fg_color=SEPARATOR_BACKGROUND_COLOR,
            border_width=1,
        )

        # Data Frame
        # TODO: Check if data is valid
        self.data = []
        data_frame = customtkinter.CTkFrame(
            self, corner_radius=0, fg_color=self.FRAMES_BACKGROUND_COLOR
        )

        data_header_frame = customtkinter.CTkFrame(
            data_frame,
            height=self.DATA_CONTROLS_FRAME_HEIGHT,
            corner_radius=0,
            fg_color=self.FRAMES_BACKGROUND_COLOR,
        )

        data_header_label_frame = customtkinter.CTkFrame(
            data_header_frame,
            height=self.DATA_CONTROLS_FRAME_HEIGHT,
            corner_radius=0,
            fg_color=self.FRAMES_BACKGROUND_COLOR,
        )

        # TODO: Should be?
        # Unicode symbols: http://xahlee.info/comp/unicode_circled_numbers.html
        data_header_label = customtkinter.CTkLabel(
            data_header_label_frame, text=_("② Data")
        )

        info_button_image = customtkinter.CTkImage(
            light_image=Image.open(self.INFO_BUTTON_LIGHT_PATH),
            dark_image=Image.open(self.INFO_BUTTON_DARK_PATH),
            size=self.INFO_BUTTON_ICON_SIZE,
        )

        data_header_label_info = customtkinter.CTkButton(
            data_header_label_frame,
            text="",
            image=info_button_image,
            fg_color="transparent",
            width=0,
            height=0,
            hover=False,
        )

        # FIXME: If translate doesn't work because of the text being a const
        # TODO: Refactor width and spacing for the tooltip
        info_button_tooltip = CTkToolTip(
            data_header_label_info,
            message=self.DATA_INFO_MESSAGE,
            corner_radius=5,
            border_width=1,
            border_color=self.TOOLTIP_BORDER_COLOR,
        )

        data_header_label_info.bind(
            "<Enter>",
            lambda event: self.on_enter(event, info_button_tooltip),
        )

        data_header_controls_frame = customtkinter.CTkFrame(
            data_header_frame,
            corner_radius=0,
            fg_color=self.FRAMES_BACKGROUND_COLOR,
        )

        # TODO: Add import csv files as data
        data_file_label = customtkinter.CTkLabel(
            data_header_controls_frame, text=_("Data file:")
        )

        self.data_file = None

        # FIXME: If translate doesn't work because of the text being a const
        # Default width of the label (can provide font),
        self.DATA_FILENAME_WIDTH = self.get_width_text(self.NO_FILE_SELECTED)

        # TODO: Should be?
        # # Width: plus 1 (if width of label equals the width of the text the text you jump out by 1 px) and plus 10 (margin, 5 to the left, 5 to the right)
        self.data_filename_label = customtkinter.CTkLabel(
            data_header_controls_frame,
            width=self.DATA_FILENAME_WIDTH + 1 + 10,
            text=self.NO_FILE_SELECTED,
        )

        self.data_filename_label_tooltip = CTkToolTip(
            self.data_filename_label,
            message=None,
            corner_radius=5,
            border_width=1,
            border_color=self.TOOLTIP_BORDER_COLOR,
        )

        # If the text is smaller than the label width no need to show a tooltip
        self.data_filename_label_tooltip.hide()

        self.data_filename_label.bind(
            "<Enter>",
            lambda event: self.on_enter(event, self.data_filename_label_tooltip),
        )

        # TODO: Set a better color
        # Width to 1 so the button takes the width of the text
        data_file_button = customtkinter.CTkButton(
            data_header_controls_frame,
            width=1,
            text=_("Choose file"),
            # fg_color="transparent",
            hover_color=self.HOVER_COLOR,
            command=self.choose_data_file,
        )

        # TODO: Set a better color
        # Width to 1 so the button takes the width of the text
        data_file_clear_button = customtkinter.CTkButton(
            data_header_controls_frame,
            width=1,
            text=_("Reset"),
            # fg_color="transparent",
            hover_color=self.HOVER_COLOR,
            command=self.reset_data_content,
        )

        # TODO: Change to scrollable X & Y
        # TODO: Scrollable X & Y background color transparent
        # TODO: Stop scroll when has few elements
        # TODO: Enable horizontal and vertical scroll with Shift+Wheel
        # TODO: Add button to delete row/column
        # TODO: Add tooltip for each Entry
        # TODO: Size of column labels must be the same (Row1000000000 should cut the text like the folders name in 1 and 3)
        # TODO: Add button to change variable type (string, number, date, etc)
        # TODO: Search by field and/or button to go to the end
        self.data_entry_scroll_frame = CTkXYFrame(
            data_frame,
            corner_radius=0,
            # fg_color=self.FRAMES_BACKGROUND_COLOR,
        )

        labels = [_("Variables")] + [
            _("Row {i}").format(i=i) for i in range(1, self.START_ENTRY_ROW_NUM)
        ]  # [
        #     _(f"Row {i}" for i in range(1, self.START_ENTRY_ROW_NUM))
        # ]
        # FIXME: If translate doesn't work because of the text being a const
        for i, label_text in enumerate(labels):
            label = customtkinter.CTkLabel(
                self.data_entry_scroll_frame, text=label_text, padx=10, pady=5
            )
            label.grid(row=i, column=0, sticky="w")

            # Plus one because the column in index 0 is for the labels
            for j in range(1, self.START_ENTRY_COL_NUM + 1):
                entry = customtkinter.CTkEntry(self.data_entry_scroll_frame)
                entry.grid(row=i, column=j, padx=10, pady=5)

        add_entry_image = customtkinter.CTkImage(
            light_image=Image.open(self.ADD_ENTRY_LIGHT_PATH),
            dark_image=Image.open(self.ADD_ENTRY_DARK_PATH),
            size=self.ADD_ENTRY_ICON_SIZE,
        )

        self.add_row_button = customtkinter.CTkButton(
            self.data_entry_scroll_frame,
            height=self.DATA_ADD_BUTTON_SIZE,
            text="",
            image=add_entry_image,
            command=self.add_row,
        )

        self.add_column_button = customtkinter.CTkButton(
            self.data_entry_scroll_frame,
            width=self.DATA_ADD_BUTTON_SIZE,
            text="",
            image=add_entry_image,
            command=self.add_column,
        )

        # Create Separator for Data and Destination Frames
        separator_d_d = customtkinter.CTkFrame(
            self,
            corner_radius=0,
            height=1,
            fg_color=SEPARATOR_BACKGROUND_COLOR,
            border_width=1,
        )

        # Destination Frame
        destination_frame = customtkinter.CTkFrame(
            self,
            height=self.DESTINATION_FRAME_HEIGHT,
            corner_radius=0,
            fg_color=self.FRAMES_BACKGROUND_COLOR,
        )

        destination_label_frame = customtkinter.CTkFrame(
            destination_frame, corner_radius=0, fg_color=self.FRAMES_BACKGROUND_COLOR
        )

        # TODO: Should be?
        # Unicode symbols: http://xahlee.info/comp/unicode_circled_numbers.html
        destination_label = customtkinter.CTkLabel(
            destination_label_frame, text=_("③ Destination")
        )

        destination_controls_frame = customtkinter.CTkFrame(
            destination_frame, corner_radius=0, fg_color=self.FRAMES_BACKGROUND_COLOR
        )

        destination_folder_label = customtkinter.CTkLabel(
            destination_controls_frame, text=_("Destination folder:")
        )

        self.destination_folder = None

        # FIXME: If translate doesn't work because of the text being a const
        # Default width of the label (can provide font),
        self.DESTINATION_FOLDER_NAME_WIDTH = self.get_width_text(
            self.NO_FOLDER_SELECTED
        )

        # FIXME: If translate doesn't work because of the text being a const
        # TODO: Should be?
        # Width: plus 1 (if width of label equals the width of the text the text you jump out by 1 px) and plus 10 (margin, 5 to the left, 5 to the right)
        self.destination_folder_name_label = customtkinter.CTkLabel(
            destination_controls_frame,
            width=self.DESTINATION_FOLDER_NAME_WIDTH + 1 + 10,
            text=self.NO_FOLDER_SELECTED,
        )

        self.destination_folder_name_label_tooltip = CTkToolTip(
            self.destination_folder_name_label,
            message=None,
            corner_radius=5,
            border_width=1,
            border_color=self.TOOLTIP_BORDER_COLOR,
        )

        # If the text is smaller than the label width no need to show a tooltip
        self.destination_folder_name_label_tooltip.hide()

        self.destination_folder_name_label.bind(
            "<Enter>",
            lambda event: self.on_enter(
                event, self.destination_folder_name_label_tooltip
            ),
        )

        # TODO: Set a better color
        # Width to 1 so the button takes the width of the text
        destination_folder_button = customtkinter.CTkButton(
            destination_controls_frame,
            width=1,
            text=_("Choose folder"),
            # fg_color="transparent",
            hover_color=self.HOVER_COLOR,
            command=self.choose_destination_folder,
        )

        # TODO: Set a better color
        # Width to 1 so the button takes the width of the text
        destination_folder_clear_button = customtkinter.CTkButton(
            destination_controls_frame,
            width=1,
            text=_("Reset"),
            # fg_color="transparent",
            hover_color=self.HOVER_COLOR,
            command=self.reset_destination_folder,
        )

        self.pdf_value = customtkinter.StringVar(value="0")
        self.pdf_checkbox = customtkinter.CTkCheckBox(
            destination_controls_frame,
            text=_("Create PDF file"),
            variable=self.pdf_value,
            onvalue="1",
            offvalue="0",
        )

        # TODO: Set a better color
        # Width to 1 so the button takes the width of the text
        destination_generate_button = customtkinter.CTkButton(
            destination_controls_frame,
            width=50,
            text=_("Generate"),
            # fg_color="transparent",
            hover_color=self.HOVER_COLOR,
            command=self.generate_contracts,
        )

        # Template Layout
        template_label.pack(side="left", padx=10)
        template_label_frame.place(x=0, y=0, relheight=0.5, relwidth=1)
        template_file_label.pack(side="left", padx=10)
        # Left padding 0px and right padding 10px
        self.template_filename_label.pack(side="left", padx=(0, 10))
        template_file_button.pack(side="left", padx=(0, 5))
        template_file_clear_button.pack(side="left")
        template_controls_frame.place(x=0, rely=0.5, relheight=0.5, relwidth=1)
        template_frame.pack(fill="x")

        # Separator Layout for Template Frame and Data Frame
        separator_t_d.pack(fill="x")

        # Data Layout
        num_cols = self.data_entry_scroll_frame.grid_size()[0]
        num_rows = self.data_entry_scroll_frame.grid_size()[1]

        self.add_row_button.grid(
            row=num_rows,
            column=0,
            columnspan=self.START_COL_SPAN,
            pady=10,
            sticky="we",
        )

        self.add_column_button.grid(
            row=0,
            column=num_cols,
            rowspan=self.START_ROW_SPAN,
            padx=10,
            sticky="ns",
        )

        data_header_label.pack(side="left", padx=(10, 0))
        data_header_label_info.pack(side="left")
        data_header_label_frame.place(x=0, y=0, relheight=0.5, relwidth=1)
        data_header_controls_frame.place(x=0, rely=0.5, relheight=0.5, relwidth=1)
        data_file_label.pack(side="left", padx=10)
        self.data_filename_label.pack(side="left", padx=(0, 10))
        data_file_button.pack(side="left", padx=(0, 5))
        data_file_clear_button.pack(side="left")
        data_header_frame.pack(fill="x")
        self.data_entry_scroll_frame.pack(expand=True, fill="both")
        data_frame.pack(expand=True, fill="both")

        # Separator Layout for Data Frame and Destination Frame
        separator_d_d.pack(fill="x")

        # Destination Layout
        destination_label.pack(side="left", padx=10)
        destination_label_frame.place(x=0, y=0, relheight=0.5, relwidth=1)
        destination_folder_label.pack(side="left", padx=10)
        self.destination_folder_name_label.pack(side="left", padx=(0, 10))
        destination_folder_button.pack(side="left", padx=(0, 5))
        destination_folder_clear_button.pack(side="left")
        destination_generate_button.pack(side="right", padx=(0, 10))
        self.pdf_checkbox.pack(side="right", padx=(0, 10))
        destination_controls_frame.place(x=0, rely=0.5, relheight=0.5, relwidth=1)
        destination_frame.pack(fill="x")

    def load_configuration(self):
        theme, sounds, language = self._read_config()

        # FIXME: Seems that set_appearance_mode needs to be set after MessageBox is drawn, but happens only sometimes
        # Wait for MessageBox is drawn to change appearance mode
        self.update()
        customtkinter.set_appearance_mode(theme)
        self.app_sounds = eval(sounds)
        # TODO: load language

    def _read_config(self):
        if os.path.isfile(self.config_file):
            self.config.read(self.config_file)

            if not self.config.has_section("Settings"):
                self._wrong_config_warning(type="section", section="Settings")
                return self._default_config()
            else:
                theme = self._check_option(
                    "theme", options["themes"], options["default_theme"]
                ).lower()

                sounds = (
                    self._check_option(
                        "sounds", options["sounds"], options["default_sounds"]
                    )
                    .lower()
                    .capitalize()
                )

                lang = self._check_option(
                    "language", options["languages"], options["default_language"]
                ).lower()

                return theme, sounds, lang
        else:
            return self._default_config()

    def _default_config(self):
        self.config["Settings"] = {
            "theme": "system",
            "sounds": "True",
            "language": "english",
        }
        with open(self.config_file, "w") as config_file:
            self.config.write(config_file)

        return "system", "True", "english"

    def _check_option(self, option, valid_values, default_value):
        if not self.config.has_option("Settings", option):
            value = default_value
            self._wrong_config_warning(type="option", section="Settings", option=option)
            self.config.set("Settings", option, default_value)
            with open(self.config_file, "w") as config_file:
                self.config.write(config_file)
        else:
            value = self.config.get("Settings", option)
            if value not in valid_values:
                value = default_value
                self._wrong_config_warning(
                    type="value", option=option, values=", ".join(valid_values)
                )
                self.config.set("Settings", option, default_value)
                with open(self.config_file, "w") as config_file:
                    self.config.write(config_file)
        return value

    def _wrong_config_warning(self, type, section="", option="", values=""):
        if type == "section":
            message = _(
                f"Config file missing {section} section! File will be restored!"
            )
        elif type == "option":
            message = _(
                f"Config file missing {option} in {section}! File will be restored!"
            )
        elif type == "value":
            message = _(
                f"Config file with wrong value in {option}! Valid values: {values}. File will be restored!"
            )
        else:
            message = _("Something wrong in config file. Will be restored!")

        CTkMessagebox(
            title=_("Bad config file!"),
            message=message,
            icon="warning",
            option_1="Ok",
        )

    def open_browser(self):
        webbrowser.open_new(self.SOURCE_CODE_URL)

    def get_width_text(self, text, font=None):
        temp_label = customtkinter.CTkLabel(self, text=text, font=font)
        temp_label.update_idletasks()
        width = temp_label.winfo_reqwidth()
        temp_label.destroy()

        return width

    def on_enter(self, event, tooltip):
        tooltip.get()

    def add_row(self):
        entry_cols = self.data_entry_scroll_frame.grid_size()[0] - 2
        num_rows = self.data_entry_scroll_frame.grid_size()[1]

        new_entry_row = num_rows - 1

        self.add_column_button.grid_configure(rowspan=num_rows)

        self.add_row_button.grid(row=num_rows, pady=10, sticky="we")

        label_text = _(f"Row {new_entry_row}")
        label = customtkinter.CTkLabel(
            self.data_entry_scroll_frame, text=label_text, padx=10, pady=5
        )
        label.grid(row=new_entry_row, column=0, sticky="w")

        for j in range(1, entry_cols + 1):
            entry = customtkinter.CTkEntry(self.data_entry_scroll_frame)
            entry.grid(row=new_entry_row, column=j, padx=10, pady=5)

    def add_column(self):
        num_cols = self.data_entry_scroll_frame.grid_size()[0]
        entry_rows = self.data_entry_scroll_frame.grid_size()[1] - 1

        new_entry_col = num_cols - 1

        self.add_row_button.grid_configure(columnspan=num_cols)

        self.add_column_button.grid(column=num_cols, pady=10, sticky="ns")

        for j in range(0, entry_rows):
            entry = customtkinter.CTkEntry(self.data_entry_scroll_frame)
            entry.grid(row=j, column=new_entry_col, padx=10, pady=5)

    # TODO: Refactor this function with choose_destination_folder()
    def choose_template_file(self):
        app_path = os.path.dirname(os.path.abspath(__file__))

        file = filedialog.askopenfile(
            initialdir=app_path, filetypes=[(_("Word files"), ".docx")]
        )

        if file:
            self.template_file = file

            filename = os.path.basename(file.name)

            text_width = self.get_width_text(filename)

            label_width = self.TEMPLATE_FILENAME_WIDTH
            if text_width > self.TEMPLATE_FILENAME_WIDTH:
                # Less one (approximately three dots size in pixels)
                # TODO: Problems with chars outside ASCII could raise problems (like ã, é ...)
                text = (
                    filename[: int(label_width / (text_width / len(filename))) - 1]
                    + "..."
                )

                self.template_filename_label.configure(text=text)
                self.template_filename_label_tooltip.configure(message=filename)

                if self.template_filename_label_tooltip.is_disabled():
                    self.template_filename_label_tooltip.show()
            else:
                self.template_filename_label.configure(text=filename)
                self.template_filename_label_tooltip.hide()

    # FIXME: If translate doesn't work because of the text being a const
    # TODO: Refactor this function with reset_destination_folder()
    def reset_template_file(self):
        if self.template_file is not None:
            self.template_file = None
            self.template_filename_label.configure(text=self.NO_FILE_SELECTED)
            self.template_filename_label_tooltip.configure(message=None)
            self.template_filename_label_tooltip.hide()

    # TODO: Refactor this function with reset_template_file() and reset_destination_folder()
    def choose_data_file(self):
        app_path = os.path.dirname(os.path.abspath(__file__))

        file = filedialog.askopenfile(
            initialdir=app_path, filetypes=[(_("Excel files"), ".xlsx")]
        )

        if file:
            self.data_file = file

            filename = os.path.basename(file.name)

            text_width = self.get_width_text(filename)

            label_width = self.DATA_FILENAME_WIDTH
            if text_width > self.DATA_FILENAME_WIDTH:
                # Less one (approximately three dots size in pixels)
                # TODO: Problems with chars outside ASCII could raise problems (like ã, é ...)
                text = (
                    filename[: int(label_width / (text_width / len(filename))) - 1]
                    + "..."
                )

                self.data_filename_label.configure(text=text)
                self.data_filename_label_tooltip.configure(message=filename)

                if self.data_filename_label_tooltip.is_disabled():
                    self.data_filename_label_tooltip.show()
            else:
                self.data_filename_label.configure(text=filename)
                self.data_filename_label_tooltip.hide()

            self.import_data(file.name)

    # TODO: Import if the user type in the entries
    def import_data(self, file_path):
        entry_cols = self.data_entry_scroll_frame.grid_size()[0] - 2
        entry_rows = self.data_entry_scroll_frame.grid_size()[1] - 1

        df = pd.read_excel(file_path)

        excel_num_rows, excel_num_columns = df.shape
        variables = df.columns

        # Plus one for variables
        self.data = [[None] * excel_num_columns for _ in range(excel_num_rows + 1)]

        while entry_cols < excel_num_columns:
            self.add_column()
            entry_cols += 1

        while entry_rows <= excel_num_rows:
            self.add_row()
            entry_rows += 1

        frame_widgets = self.data_entry_scroll_frame.winfo_children()
        for widget in frame_widgets:
            if isinstance(widget, customtkinter.CTkEntry):
                row_index = widget.grid_info()["row"]
                col_index = widget.grid_info()["column"]

                if row_index == 0:
                    widget.insert(0, variables[col_index - 1])
                    self.data[row_index][col_index - 1] = variables[col_index - 1]
                else:
                    data = df.iloc[row_index - 1, col_index - 1]
                    widget.insert(0, data)
                    self.data[row_index][col_index - 1] = data

    # FIXME: If translate doesn't work because of the text being a const
    # TODO: Refactor this function with reset_template_file() and reset_destination_folder()
    def reset_data_content(self):
        if self.data_file is not None:
            self.data_file = None
            self.data_filename_label.configure(text=self.NO_FILE_SELECTED)
            self.data_filename_label_tooltip.configure(message=None)
            self.data_filename_label_tooltip.hide()

            frame_widgets = self.data_entry_scroll_frame.winfo_children()
            for widget in frame_widgets:
                row_index = widget.grid_info()["row"]
                col_index = widget.grid_info()["column"]

                if row_index > 3 or col_index > 2:
                    if not isinstance(widget, customtkinter.CTkButton):
                        widget.destroy()
                else:
                    if isinstance(widget, customtkinter.CTkEntry):
                        widget.delete(0, END)

                # TODO: If the user is in the right extreme and bottom extreme and the scroll is bigger
                #       then the initially table, the user has to scroll up of to the left to update the frame
                # Update row and column span to update the grid size
                if isinstance(widget, customtkinter.CTkButton):
                    if row_index == 0:
                        # Plus 1 because the column with index 0 is for the labels
                        widget.grid_configure(
                            column=self.START_ENTRY_COL_NUM + 1,
                            rowspan=self.START_ROW_SPAN,
                        )
                    elif col_index == 0:
                        widget.grid_configure(
                            row=self.START_ENTRY_ROW_NUM, columnspan=self.START_COL_SPAN
                        )
                self.data = []

    # TODO: Refactor this function with choose_template_file() and choose_data_file()
    def choose_destination_folder(self):
        app_path = os.path.dirname(os.path.abspath(__file__))

        folder = filedialog.askdirectory(initialdir=app_path)

        if folder:
            self.destination_folder = folder

            text_width = self.get_width_text(folder)

            label_width = self.DESTINATION_FOLDER_NAME_WIDTH
            if text_width > self.DESTINATION_FOLDER_NAME_WIDTH:
                # Less one (approximately three dots size in pixels)
                text = (
                    folder[: int(label_width / (text_width / len(folder))) - 1] + "..."
                )
                self.destination_folder_name_label.configure(text=text)
                self.destination_folder_name_label_tooltip.configure(message=folder)

                if self.destination_folder_name_label_tooltip.is_disabled():
                    self.destination_folder_name_label_tooltip.show()
            else:
                self.destination_folder_name_label.configure(text=folder)
                self.destination_folder_name_label_tooltip.hide()

    # TODO: Refactor this function with reset_template_file() and reset_destination_folder()
    def reset_destination_folder(self):
        if self.destination_folder is not None:
            self.destination_folder = None
            self.destination_folder_name_label.configure(text=self.NO_FOLDER_SELECTED)
            self.destination_folder_name_label_tooltip.configure(message=None)
            self.destination_folder_name_label_tooltip.hide()

    def reset_all(self):
        self.reset_template_file()
        self.reset_data_content()
        self.reset_destination_folder()

    # TODO: Error boxes for warnings or errors
    # TODO: Setting to disable sound (like the theme)
    def generate_contracts(self):
        if self.template_file is None:
            CTkMessagebox(
                title="Error",
                message=_("No template file selected!"),
                icon="cancel",
                sound=self.app_sounds,
                option_focus=1,
                cancel_button=None,
                # cancel_button_color="transparent",
            )
        elif not self.data:
            CTkMessagebox(
                title="Error",
                message=_("No data!"),
                icon="cancel",
                sound=self.app_sounds,
                option_focus=1,
                cancel_button=None,
                # cancel_button_color="transparent",
            )
        elif self.destination_folder is None:
            CTkMessagebox(
                title="Error",
                message=_("No destination folder selected!"),
                icon="cancel",
                sound=self.app_sounds,
                option_focus=1,
                cancel_button=None,
                # cancel_button_color="transparent",
            )
        else:
            self.pdf_checkbox.configure(state=DISABLED)
            # if self.pdf_value.get() == 1
            variables = self.data[0]
            for i, data_list in enumerate(self.data[1:], start=1):
                doc = Document(self.template_file.name)
                for j, data in enumerate(data_list):
                    self.replace_text_in_docx(doc, variables[j], data)

                # TODO: Better name for the output file
                # filename = os.path.basename(self.template_file.name)
                output_path = os.path.join(self.destination_folder, f"OUTPUT_{i}.docx")
                doc.save(output_path)

                # TODO: Convert to pdf is slow, maybe use threads
                # TODO: Test errors, like permission errors
                if self.pdf_value.get() == "1":
                    convert(output_path)

            # TODO: Progress bar and final message
            self.pdf_checkbox.configure(state=NORMAL)

    def replace_text_in_docx(self, doc, old_text, new_text):
        for paragraph in doc.paragraphs:
            if old_text in paragraph.text:
                for run in paragraph.runs:
                    run.text = run.text.replace(old_text, new_text)

    def open_settings(self):
        if self.settings_window is None or not self.settings_window.winfo_exists():
            self.settings_window = SettingsTopLevel(self)
            self.settings_window.bind(
                "<Visibility>", lambda event: self.settings_window.focus()
            )
        else:
            self.settings_window.focus()


class SettingsTopLevel(customtkinter.CTkToplevel, ConfigHolder):
    SETTINGS_WIDTH = 300
    SETTINGS_HEIGHT = 300

    def __init__(self, auto_contract, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.parent = auto_contract

        self.config_file = ConfigHolder.config_file
        self.config = ConfigHolder.config

        self.title(_("Settings"))

        # CTk wait to 200ms to set the icon - FIXME: could be fixed in the repo
        self.after(
            200,
            lambda: self.iconbitmap(
                os.path.join(
                    os.path.dirname(os.path.abspath(__file__)),
                    "assets/logos/settings128x128.ico",
                )
            ),
        )

        self.geometry(f"{self.SETTINGS_WIDTH}x{self.SETTINGS_HEIGHT}")
        self.resizable(width=False, height=False)

        theme, sounds, language = self.parent._read_config()

        appearance_frame = customtkinter.CTkFrame(
            self, corner_radius=0, fg_color="transparent"
        )

        appearance_title = customtkinter.CTkLabel(
            appearance_frame, text=_("Appearance")
        )

        appearance_options_frame = customtkinter.CTkFrame(
            appearance_frame, corner_radius=0, fg_color="transparent"
        )

        appearance_mode_label = customtkinter.CTkLabel(
            appearance_options_frame, text=_("Mode:")
        )

        appearance_mode_var = customtkinter.StringVar(value=theme.capitalize())
        # FIXME: If breaks change_appearance_mode (different values)
        appearance_mode_combobox = customtkinter.CTkComboBox(
            appearance_options_frame,
            values=[_("Light"), _("Dark"), _("System")],
            command=self.change_appearance_mode,
            variable=appearance_mode_var,
        )

        # Create Separator for Appearance and Sound Frames
        separator_a_s = customtkinter.CTkFrame(
            self,
            corner_radius=0,
            height=1,
            fg_color=SEPARATOR_BACKGROUND_COLOR,
            border_width=1,
        )

        sound_frame = customtkinter.CTkFrame(
            self, corner_radius=0, fg_color="transparent"
        )

        sound_title = customtkinter.CTkLabel(sound_frame, text=_("Sound"))

        sound_options_frame = customtkinter.CTkFrame(
            sound_frame, corner_radius=0, fg_color="transparent"
        )

        self.sound_var = customtkinter.StringVar(value=("1" if eval(sounds) else "0"))
        sound_checkbox = customtkinter.CTkCheckBox(
            sound_options_frame,
            text=_("Enable sound"),
            variable=self.sound_var,
            onvalue="1",
            offvalue="0",
            command=self.change_sounds,
        )

        # TODO: Try to set the volume

        # Create Separator for Sound and Language Frames
        separator_s_l = customtkinter.CTkFrame(
            self,
            corner_radius=0,
            height=1,
            fg_color=SEPARATOR_BACKGROUND_COLOR,
            border_width=1,
        )

        lang_frame = customtkinter.CTkFrame(
            self, corner_radius=0, fg_color="transparent"
        )

        lang_title = customtkinter.CTkLabel(lang_frame, text=_("Language"))

        lang_options_frame = customtkinter.CTkFrame(
            lang_frame, corner_radius=0, fg_color="transparent"
        )

        lang_label = customtkinter.CTkLabel(lang_options_frame, text=_("Choose:"))

        lang_var = customtkinter.StringVar(value=language.capitalize())
        # FIXME: If breaks change_appearance_mode (different values)
        lang_combobox = customtkinter.CTkComboBox(
            lang_options_frame,
            values=[_("English"), _("Portuguese")],
            command=self.change_lang,
            variable=lang_var,
        )

        # TODO: Select as default generate pdf (init as checked or not)

        # TODO: Select a default folder to output

        # Layout for appearance
        appearance_title.pack(anchor="w", padx=(10, 0))
        appearance_options_frame.pack(fill="x")
        appearance_mode_label.pack(side="left", padx=(10, 10))
        appearance_mode_combobox.pack(side="left")
        appearance_frame.pack(expand=True, fill="x")

        # Separator
        separator_a_s.pack(fill="x")

        # Layout for sound
        sound_title.pack(anchor="w", padx=(10, 0))
        sound_options_frame.pack(fill="x")
        sound_checkbox.pack(anchor="w", padx=(10, 0))
        sound_frame.pack(expand=True, fill="x")

        # Separator
        separator_s_l.pack(fill="x")

        # Layout for language
        lang_title.pack(anchor="w", padx=(10, 0))
        lang_options_frame.pack(fill="x")
        lang_label.pack(side="left", padx=(10, 10))
        lang_combobox.pack(side="left")
        lang_frame.pack(expand=True, fill="x")

    def change_appearance_mode(self, appearance_mode):
        customtkinter.set_appearance_mode(appearance_mode.lower())

        self.config.set("Settings", "theme", appearance_mode.lower())

        with open(self.config_file, "w") as config_file:
            self.config.write(config_file)

    def change_sounds(self):
        self.parent.app_sounds = self.sound_var.get() == "1"

        self.config.set("Settings", "sounds", str(self.parent.app_sounds))
        with open(self.config_file, "w") as config_file:
            self.config.write(config_file)

    def change_lang(self, lang):
        # TODO: Translate to different languages
        # IN PROGRESS, but is better to finish the app and then create the translations
        self.config.set("Settings", "language", lang.lower())

        with open(self.config_file, "w") as config_file:
            self.config.write(config_file)


def main():
    # TODO: Credits for <a href="https://www.flaticon.com/free-icons/plus" title="plus icons">Plus icons created by Fuzzee - Flaticon</a>
    # TODO: Credits for <a href="https://www.flaticon.com/free-icons/info" title="info icons">Info icons created by Graphics Plazza - Flaticon</a>
    # TODO: Credits for <a href="https://www.flaticon.com/free-icons/settings" title="settings icons">Settings icons created by Freepik - Flaticon</a>
    app = AutoContract()
    app.mainloop()


if __name__ == "__main__":
    main()
