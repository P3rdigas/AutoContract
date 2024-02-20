import customtkinter

from CTkXYFrame import CTkXYFrame


class App(customtkinter.CTk):
    APP_WIDTH = 800
    APP_HEIGHT = 600

    def __init__(self):
        super().__init__()

        self.title("Frame example")
        self.geometry(f"{self.APP_WIDTH}x{self.APP_HEIGHT}")
        self.resizable(width=False, height=False)

        self.check_var = customtkinter.StringVar(value=1)
        pdf_checkbox = customtkinter.CTkCheckBox(
            self,
            text="Create PDF file",
            command=self.checkbox_event,
            variable=self.check_var,
            onvalue=1,
            offvalue=0,
        )

        pdf_checkbox.pack()

    def checkbox_event(self):
        print("checkbox toggled, current value:", self.check_var.get())


if __name__ == "__main__":
    customtkinter.set_appearance_mode("dark")
    app = App()
    app.mainloop()
