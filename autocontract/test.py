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

        self.frame = CTkXYFrame(self, corner_radius=0)

        label1 = customtkinter.CTkLabel(self.frame, text="Hello", padx=10, pady=5)
        label1.grid(row=0, column=0, sticky="w")
        button1 = customtkinter.CTkButton(self.frame, text="Hello")
        button1.grid(row=0, column=1, sticky="w")

        label2 = customtkinter.CTkLabel(self.frame, text="Bye", padx=10, pady=5)
        label2.grid(row=1, column=0, sticky="w")
        button2 = customtkinter.CTkButton(self.frame, text="Bye")
        button2.grid(row=1, column=1, sticky="w")

        label3 = customtkinter.CTkLabel(self.frame, text="Again", padx=10, pady=5)
        label3.grid(row=2, column=0, sticky="w")
        button3 = customtkinter.CTkButton(self.frame, text="Again")
        button3.grid(row=2, column=1, sticky="w")
        button4 = customtkinter.CTkButton(self.frame, text="I")
        button4.grid(row=2, column=2, sticky="w")
        button5 = customtkinter.CTkButton(self.frame, text="Will")
        button5.grid(row=2, column=3, sticky="w")

        button6 = customtkinter.CTkButton(
            self.frame, text="Delete", command=self.delete
        )
        button6.grid(row=3, column=0)
        self.frame.pack(expand=True, fill="both")

    def delete(self):
        frame_widgets = self.frame.winfo_children()

        for widget in frame_widgets:
            row_index = widget.grid_info()["row"]
            col_index = widget.grid_info()["column"]
            if col_index >= 1:
                widget.destroy()

        num_cols = self.frame.grid_size()[0]
        num_rows = self.frame.grid_size()[1]
        print(f"Num_Cols: {num_cols} Num_Rows: {num_rows}")


if __name__ == "__main__":
    customtkinter.set_appearance_mode("dark")
    app = App()
    app.mainloop()
