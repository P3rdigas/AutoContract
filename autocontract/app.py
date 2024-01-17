import configparser
import os
import webbrowser
from tkinter import filedialog

import customtkinter
from CTkMenuBar import CTkMenuBar, CustomDropdownMenu
from CTkToolTip import CTkToolTip
from PIL import Image


class AutoContract(customtkinter.CTk):
    APP_WIDTH = 800
    APP_HEIGHT = 600

    TEMPLATE_FRAME_HEIGHT = 75
    DATA_CONTROLS_FRAME_HEIGHT = 75
    DESTINATION_FRAME_HEIGHT = 75

    # Colors first correspond to Light Mode and second to Dark Mode
    MENUBAR_BACKGROUND_COLOR = "white", "black"
    DROPDOWN_BACKGROUND_COLOR = "white", "grey20"
    FRAMES_BACKGROUND_COLOR = "transparent"
    SEPARATOR_BACKGROUND_COLOR = "grey90", "black"

    TOOLTIP_BORDER_COLOR = "black", "white"

    HOVER_COLOR = "grey85", "grey25"

    NO_FOLDER_SELECTED = "No folder selected"
    NO_FILE_SELECTED = "No file selected"

    SOURCE_CODE_URL = "https://github.com/P3rdigas/AutoContract"

    def __init__(self):
        super().__init__()

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
        self.config_file = os.path.join(
            os.path.dirname(os.path.abspath(__file__)), "config.ini"
        )
        self.config = configparser.ConfigParser()
        self.load_configuration()

        # Set Menubar
        self.toolbar = CTkMenuBar(master=self, bg_color=self.MENUBAR_BACKGROUND_COLOR)
        self.file_button = self.toolbar.add_cascade(
            "File",
            hover_color=self.HOVER_COLOR,
        )
        self.settings_button = self.toolbar.add_cascade(
            "Settings",
            hover_color=self.HOVER_COLOR,
        )
        self.about_button = self.toolbar.add_cascade(
            "About",
            hover_color=self.HOVER_COLOR,
        )

        self.file_button_dropdown = CustomDropdownMenu(
            widget=self.file_button,
            corner_radius=0,
            bg_color=self.DROPDOWN_BACKGROUND_COLOR,
            hover_color=self.HOVER_COLOR,
        )
        self.file_button_dropdown.add_option(option="Import")
        self.file_button_dropdown.add_option(option="Export")
        self.file_button_dropdown.add_option(option="Exit", command=self.destroy)

        self.settings_button_dropdown = CustomDropdownMenu(
            widget=self.settings_button,
            corner_radius=0,
            bg_color=self.DROPDOWN_BACKGROUND_COLOR,
            hover_color=self.HOVER_COLOR,
        )
        appearance_sub_menu = self.settings_button_dropdown.add_submenu("Appearance")
        appearance_sub_menu.add_option(
            option="Light", command=lambda: self.change_appearance_mode_event("light")
        )
        appearance_sub_menu.add_option(
            option="Dark", command=lambda: self.change_appearance_mode_event("dark")
        )
        appearance_sub_menu.add_option(
            option="System", command=lambda: self.change_appearance_mode_event("system")
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

        self.about_button_dropdown = CustomDropdownMenu(
            widget=self.about_button,
            corner_radius=0,
            bg_color=self.DROPDOWN_BACKGROUND_COLOR,
            hover_color=self.HOVER_COLOR,
        )
        self.about_button_dropdown.add_option(
            option="Source Code", image=github_image, command=self.open_browser
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

        template_label = customtkinter.CTkLabel(template_label_frame, text="① Template")

        template_controls_frame = customtkinter.CTkFrame(
            template_frame, corner_radius=0, fg_color=self.FRAMES_BACKGROUND_COLOR
        )

        template_file_label = customtkinter.CTkLabel(
            template_controls_frame, text="Template file:"
        )

        self.template_file = None

        # Default width of the label (can provide font),
        self.TEMPLATE_FILE_NAME_WIDTH = self.get_width_text(self.NO_FILE_SELECTED)

        # Width: plus 1 (if width of label equals the width of the text the text you jump out by 1 px) and plus 10 (margin, 5 to the left, 5 to the right)
        self.template_file_name_label = customtkinter.CTkLabel(
            template_controls_frame,
            width=self.TEMPLATE_FILE_NAME_WIDTH + 1 + 10,
            text=self.NO_FILE_SELECTED,
        )

        self.template_file_name_label_tooltip = CTkToolTip(
            self.template_file_name_label,
            message=None,
            corner_radius=5,
            border_width=1,
            border_color=self.TOOLTIP_BORDER_COLOR,
        )

        # If the text is smaller than the label width no need to show a tooltip
        self.template_file_name_label_tooltip.hide()

        self.template_file_name_label.bind(
            "<Enter>",
            lambda event: self.on_enter(event, self.template_file_name_label_tooltip),
        )

        # TODO: Set a better color
        # Width to 1 so the button takes the width of the text
        template_file_button = customtkinter.CTkButton(
            template_controls_frame,
            width=1,
            text="Choose file",
            # fg_color="transparent",
            hover_color=self.HOVER_COLOR,
            command=self.choose_template_file,
        )

        # TODO: Set a better color
        # Width to 1 so the button takes the width of the text
        template_file_clear_button = customtkinter.CTkButton(
            template_controls_frame,
            width=1,
            text="Reset",
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
            fg_color=self.SEPARATOR_BACKGROUND_COLOR,
            border_width=1,
        )

        # Data Frame
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

        # TODO: Add button for a tooltip to explain how to add data

        data_header_label = customtkinter.CTkLabel(
            data_header_label_frame, text="② Data"
        )

        data_header_controls_frame = customtkinter.CTkFrame(
            data_header_frame,
            corner_radius=0,
            fg_color=self.FRAMES_BACKGROUND_COLOR,
        )

        # TODO: Add import csv files as data

        # TODO: Change to scrollable X & Y
        data_entry_frame = customtkinter.CTkFrame(
            data_frame,
            corner_radius=0,
            fg_color=self.FRAMES_BACKGROUND_COLOR,
        )

        # Create Separator for Data and Destination Frames
        separator_d_d = customtkinter.CTkFrame(
            self,
            corner_radius=0,
            height=1,
            fg_color=self.SEPARATOR_BACKGROUND_COLOR,
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

        destination_label = customtkinter.CTkLabel(
            destination_label_frame, text="③ Destination"
        )

        destination_controls_frame = customtkinter.CTkFrame(
            destination_frame, corner_radius=0, fg_color=self.FRAMES_BACKGROUND_COLOR
        )

        destination_folder_label = customtkinter.CTkLabel(
            destination_controls_frame, text="Destination folder:"
        )

        self.destination_folder = None

        # Default width of the label (can provide font),
        self.DESTINATION_FOLDER_NAME_WIDTH = self.get_width_text(
            self.NO_FOLDER_SELECTED
        )

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
            text="Choose folder",
            # fg_color="transparent",
            hover_color=self.HOVER_COLOR,
            command=self.choose_destination_folder,
        )

        # TODO: Set a better color
        # Width to 1 so the button takes the width of the text
        destination_folder_clear_button = customtkinter.CTkButton(
            destination_controls_frame,
            width=1,
            text="Reset",
            # fg_color="transparent",
            hover_color=self.HOVER_COLOR,
            command=self.reset_destination_folder,
        )

        # TODO: Set a better color
        # Width to 1 so the button takes the width of the text
        destination_generate_button = customtkinter.CTkButton(
            destination_controls_frame,
            width=50,
            text="Generate",
            # fg_color="transparent",
            hover_color=self.HOVER_COLOR,
            command=self.generate_contract,
        )

        # Template Layout
        template_label.pack(side="left", padx=10)
        template_label_frame.place(x=0, y=0, relheight=0.5, relwidth=1)
        template_file_label.pack(side="left", padx=10)
        # Left padding 0px and right padding 10px
        self.template_file_name_label.pack(side="left", padx=(0, 10))
        template_file_button.pack(side="left", padx=(0, 5))
        template_file_clear_button.pack(side="left")
        template_controls_frame.place(x=0, rely=0.5, relheight=0.5, relwidth=1)
        template_frame.pack(fill="x")

        # Separator Layout for Template Frame and Data Frame
        separator_t_d.pack(fill="x")

        # Data Layout
        data_header_label.pack(side="left", padx=10)
        data_header_label_frame.place(x=0, y=0, relheight=0.5, relwidth=1)
        data_header_controls_frame.place(x=0, rely=0.5, relheight=0.5, relwidth=1)
        data_header_frame.pack(fill="x")
        data_entry_frame.pack(expand=True, fill="both")
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
        destination_controls_frame.place(x=0, rely=0.5, relheight=0.5, relwidth=1)
        destination_frame.pack(fill="x")

    def load_configuration(self):
        if os.path.isfile(self.config_file):
            self.config.read(self.config_file)
            theme = self.config.get("Settings", "theme")

            if theme == "system":
                customtkinter.set_appearance_mode("system")
            elif theme == "light":
                customtkinter.set_appearance_mode("light")
            else:
                customtkinter.set_appearance_mode("dark")
        else:
            self.config["Settings"] = {"theme": "system"}
            with open(self.config_file, "w") as configfile:
                self.config.write(configfile)

    def change_appearance_mode_event(self, new_appearance_mode):
        customtkinter.set_appearance_mode(new_appearance_mode)

        self.config.set("Settings", "theme", new_appearance_mode)

        with open(self.config_file, "w") as configfile:
            self.config.write(configfile)

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

    # TODO: Refactor this function with choose_destination_folder()
    def choose_template_file(self):
        app_path = os.path.dirname(os.path.abspath(__file__))

        file = filedialog.askopenfile(
            initialdir=app_path, filetypes=[("Word files", ".docx")]
        )

        if file:
            self.template_file = file

            filename = os.path.basename(file.name)

            text_width = self.get_width_text(filename)

            label_width = self.TEMPLATE_FILE_NAME_WIDTH
            if text_width > self.TEMPLATE_FILE_NAME_WIDTH:
                # Less one (approximately three dots size in pixels)
                text = (
                    filename[: int(label_width / (text_width / len(filename))) - 1]
                    + "..."
                )

                self.template_file_name_label.configure(text=text)
                self.template_file_name_label_tooltip.configure(message=filename)

                if self.template_file_name_label_tooltip.is_disabled():
                    self.template_file_name_label_tooltip.show()
            else:
                self.template_file_name_label.configure(text=filename)
                self.template_file_name_label_tooltip.hide()

    # TODO: Refactor this function with reset_destination_folder()
    def reset_template_file(self):
        if self.template_file is not None:
            self.template_file = None
            self.template_file_name_label.configure(text=self.NO_FILE_SELECTED)
            self.template_file_name_label_tooltip.configure(message=None)
            self.template_file_name_label_tooltip.hide()

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

    def reset_destination_folder(self):
        if self.destination_folder is not None:
            self.destination_folder = None
            self.destination_folder_name_label.configure(text=self.NO_FOLDER_SELECTED)
            self.destination_folder_name_label_tooltip.configure(message=None)
            self.destination_folder_name_label_tooltip.hide()

    def generate_contract(self):
        print("Generating contract")


def main():
    app = AutoContract()
    app.mainloop()


if __name__ == "__main__":
    main()
