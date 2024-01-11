import configparser
import os

import customtkinter


class AutoContract(customtkinter.CTk):
    APP_WIDTH = 800
    APP_HEIGHT = 600

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


def main():
    app = AutoContract()
    app.mainloop()


if __name__ == "__main__":
    main()
