import configparser
import os
import webbrowser

import customtkinter
from CTkMenuBar import CTkMenuBar, CustomDropdownMenu
from PIL import Image


class AutoContract(customtkinter.CTk):
    APP_WIDTH = 800
    APP_HEIGHT = 600

    MENUBAR_BACKGROUND_COLOR_LIGHT = "white"
    MENUBAR_BACKGROUND_COLOR_DARK = "black"

    DROPDOWN_BACKGROUND_COLOR_LIGHT = "white"
    DROPDOWN_BACKGROUND_COLOR_DARK = "grey20"

    DRAG_AND_DROP_BG_COLOR_LIGHT = "grey95"
    DRAG_AND_DROP_BG_COLOR_DARK = "grey15"
    DRAG_AND_DROP_SEPARATOR_COLOR_LIGHT = "grey90"
    DRAG_AND_DROP_SEPARATOR_COLOR_DARK = "black"

    TEXT_COLOR_LIGHT = "black"
    TEXT_COLOR_DARK = "white"
    HOVER_COLOR_LIGHT = "grey85"
    HOVER_COLOR_DARK = "grey25"

    GITHUB_LOGO_LIGHT_PATH = os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "assets/icons/github-mark.png"
    )
    GITHUB_LOGO_DARK_PATH = os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "assets/icons/github-mark-white.png"
    )

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
        self.toolbar = CTkMenuBar(
            master=self,
            bg_color=(
                self.MENUBAR_BACKGROUND_COLOR_LIGHT,
                self.MENUBAR_BACKGROUND_COLOR_DARK,
            ),
        )
        self.file_button = self.toolbar.add_cascade(
            "File",
            text_color=(self.TEXT_COLOR_LIGHT, self.TEXT_COLOR_DARK),
            hover_color=(self.HOVER_COLOR_LIGHT, self.HOVER_COLOR_DARK),
        )
        self.settings_button = self.toolbar.add_cascade(
            "Settings",
            text_color=(self.TEXT_COLOR_LIGHT, self.TEXT_COLOR_DARK),
            hover_color=(self.HOVER_COLOR_LIGHT, self.HOVER_COLOR_DARK),
        )
        self.about_button = self.toolbar.add_cascade(
            "About",
            text_color=(self.TEXT_COLOR_LIGHT, self.TEXT_COLOR_DARK),
            hover_color=(self.HOVER_COLOR_LIGHT, self.HOVER_COLOR_DARK),
        )

        self.file_button_dropdown = CustomDropdownMenu(
            widget=self.file_button,
            corner_radius=0,
            bg_color=(
                self.DROPDOWN_BACKGROUND_COLOR_LIGHT,
                self.DROPDOWN_BACKGROUND_COLOR_DARK,
            ),
            text_color=(self.TEXT_COLOR_LIGHT, self.TEXT_COLOR_DARK),
            hover_color=(self.HOVER_COLOR_LIGHT, self.HOVER_COLOR_DARK),
        )
        self.file_button_dropdown.add_option(option="Import")
        self.file_button_dropdown.add_option(option="Export")
        self.file_button_dropdown.add_option(option="Exit", command=self.destroy)

        self.settings_button_dropdown = CustomDropdownMenu(
            widget=self.settings_button,
            corner_radius=0,
            bg_color=(
                self.DROPDOWN_BACKGROUND_COLOR_LIGHT,
                self.DROPDOWN_BACKGROUND_COLOR_DARK,
            ),
            text_color=(self.TEXT_COLOR_LIGHT, self.TEXT_COLOR_DARK),
            hover_color=(self.HOVER_COLOR_LIGHT, self.HOVER_COLOR_DARK),
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

        github_image = customtkinter.CTkImage(
            light_image=Image.open(self.GITHUB_LOGO_LIGHT_PATH),
            dark_image=Image.open(self.GITHUB_LOGO_DARK_PATH),
        )

        self.about_button_dropdown = CustomDropdownMenu(
            widget=self.about_button,
            corner_radius=0,
            bg_color=(
                self.DROPDOWN_BACKGROUND_COLOR_LIGHT,
                self.DROPDOWN_BACKGROUND_COLOR_DARK,
            ),
            text_color=(self.TEXT_COLOR_LIGHT, self.TEXT_COLOR_DARK),
            hover_color=(self.HOVER_COLOR_LIGHT, self.HOVER_COLOR_DARK),
        )
        self.about_button_dropdown.add_option(
            option="Source Code", image=github_image, command=self.open_browser
        )

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


def main():
    app = AutoContract()
    app.mainloop()


if __name__ == "__main__":
    main()
