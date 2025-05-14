import customtkinter as ctk
from tkinter import Entry

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("System")  # Modes: "System", "Dark", "Light"
        ctk.set_default_color_theme("blue")  # Themes: "blue", "dark-blue", "green"

        # Configurer la fenêtre
        self.title("Money")
        center_window(self, 800, 535)

        self.bind("<Button-3>", self.remove_focus)
        self.bind("<Button-1>", self.remove_focus)

        # Enable window dragging
        self.offset_x = 0
        self.offset_y = 0

        self.CHECK_CREATE = 1
        self.CHECK_EXTRACT = 2
        self.CHECK_UPDATE = 3

        ### FRAME ###
        frame_choice = ctk.CTkFrame(self, fg_color='transparent')
        frame_choice.grid(row=0, column=0, sticky="nsew")
        frame_choice.grid_rowconfigure((0, 1), weight=1)
        frame_choice.grid_columnconfigure(0, weight=1)

        frame_ok_cancel = ctk.CTkFrame(self, fg_color='transparent')
        frame_ok_cancel.grid(row=1, column=0, sticky="ew")

        frame_loading = ctk.CTkFrame(self, fg_color='transparent', width=0, height=0)
        frame_loading.place(relx=0, rely=1, anchor='sw')

        ### FRAME TOP ###
        frame_top = ctk.CTkFrame(frame_choice)
        frame_top.grid(row=0, column=0, sticky="nsew")

        ### FRAME BOTTOM ###
        frame_bot = ctk.CTkFrame(frame_choice)
        frame_bot.grid(row=1, column=0, sticky="nsew")
        frame_bot.grid_columnconfigure((0, 1), weight=1)

        ### CREATE FRAME ###
        frame_choice_create = ctk.CTkFrame(frame_top, fg_color='transparent')
        frame_choice_create.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="ew")

        self.label_create_title = ctk.CTkLabel(frame_choice_create, text="Create", font=("Arial", 30, 'bold'))
        self.label_create_title.pack(pady=(0, 20))

        frame_choice_create_ym = ctk.CTkFrame(frame_choice_create, fg_color='transparent')
        frame_choice_create_ym.pack(pady=(10, 0), padx=20, fill='both')

        self.label_year_create = ctk.CTkLabel(frame_choice_create_ym, text="Year", font=("Arial", 17))
        self.label_year_create.grid(row=0, column=0, sticky="w")
        self.entry_year_create = ctk.CTkEntry(frame_choice_create_ym, width=75)
        self.entry_year_create.grid(row=0, column=1, padx=10, sticky="w")

        self.label_month_create = ctk.CTkLabel(frame_choice_create_ym, text="Month", font=("Arial", 17))
        self.label_month_create.grid(row=0, column=2, padx=10, sticky="w")
        self.entry_month_create = ctk.CTkComboBox(frame_choice_create_ym, width=120)
        self.entry_month_create.grid(row=0, column=3, padx=10, sticky="w")

        self.checkbox_create = ctk.CTkCheckBox(frame_choice_create, text='', onvalue=self.CHECK_CREATE, offvalue=0)
        self.checkbox_create.place(relx=0.93, rely=0.19, anchor='ne')

        ### EXTRACT FRAME ###
        frame_choice_extract = ctk.CTkFrame(frame_bot, fg_color='transparent')
        frame_choice_extract.grid(row=0, column=0, sticky="ew", padx=10)

        self.label_extract_title = ctk.CTkLabel(frame_choice_extract, text="Extraction", font=("Arial", 30, 'bold'))
        self.label_extract_title.grid(row=0, column=0, sticky="ew", pady=(10, 20))

        self.entry_pdf = ctk.CTkEntry(frame_choice_extract, placeholder_text="PDF Path", width=10)
        self.entry_pdf.grid(row=1, column=0, sticky="ew", padx=(20, 60), pady=(0, 10))

        self.button_browse_pdf = ctk.CTkButton(frame_choice_extract, text="Browse", width=80, height=40, font=("Arial", 15))
        self.button_browse_pdf.grid(row=2, column=0, sticky="e", padx=(0, 20), pady=(0, 10))

        self.checkbox_extract = ctk.CTkCheckBox(frame_choice_extract, text='', onvalue=self.CHECK_EXTRACT, offvalue=0)
        self.checkbox_extract.place(relx=0.96, rely=0.12, anchor='ne')

        ### UPDATE FRAME ###
        frame_choice_update = ctk.CTkFrame(frame_bot, fg_color='transparent')
        frame_choice_update.grid(row=0, column=1, sticky="nsew", padx=10)

        self.label_update_title = ctk.CTkLabel(frame_choice_update, text="Update", font=("Arial", 30, 'bold'))
        self.label_update_title.grid(row=0, column=0, sticky="ew", pady=(10, 20))

        frame_year_month_update = ctk.CTkFrame(frame_choice_update)
        frame_year_month_update.grid(row=1, column=0, sticky="ew")
        frame_year_month_update.grid_columnconfigure((0, 1), weight=1)

        self.label_year = ctk.CTkLabel(frame_year_month_update, text="Year", font=("Arial", 17, 'bold'))
        self.label_year.grid(row=0, column=0, sticky="ew")
        self.entry_year = ctk.CTkEntry(frame_year_month_update, width=7)
        self.entry_year.grid(row=1, column=0, sticky="ew", padx=(15, 15), pady=(5, 10))

        self.label_month = ctk.CTkLabel(frame_year_month_update, text="Month", font=("Arial", 17, 'bold'))
        self.label_month.grid(row=0, column=1, sticky="ew")
        self.entry_month = ctk.CTkComboBox(frame_year_month_update, width=100, values=[''])
        self.entry_month.grid(row=1, column=1, sticky="ew", padx=(15, 15), pady=(5, 10))

        self.checkbox_update = ctk.CTkCheckBox(frame_choice_update, text='', onvalue=self.CHECK_UPDATE, offvalue=0)
        self.checkbox_update.place(relx=0.93, rely=0.12, anchor='ne')

        ### OK / CANCEL BUTTONS ###
        self.button_ok = ctk.CTkButton(frame_ok_cancel, text="OK", width=80, height=40, font=("Arial", 17, "bold"))
        self.button_ok.pack(side="right", padx=15, pady=(5, 5))

        self.button_cancel = ctk.CTkButton(frame_ok_cancel, text="Cancel", width=80, height=40, font=("Arial", 17, "bold"))
        self.button_cancel.pack(side="left", padx=15, pady=(5, 5))

        # Options Button
        self.options_button = ctk.CTkButton(self, text="Option", width=30, command=self.toggle_options_menu, fg_color="transparent", hover_color="#3b3b3b")
        self.options_button.place(x=10, y=10)

        self.init_option_dropdowm()

    def remove_focus(self, event):
        self.hide_dropdown_option_menu_click_outside(event.widget)
        if not isinstance(event.widget, Entry):
            self.focus_set()

    def init_option_dropdowm(self):
        self.dropdown_frame = ctk.CTkFrame(self, width=120, fg_color="#2b2b2b")
        self.settings_button = ctk.CTkButton(self.dropdown_frame, text="Settings", fg_color="transparent", hover_color="#3b3b3b")
        self.settings_button.pack(fill="x", pady=2)
        self.help_button = ctk.CTkButton(self.dropdown_frame, text="Help", fg_color="transparent", hover_color="#3b3b3b")
        self.help_button.pack(fill="x", pady=2)
        self.dropdown_visible = False

    def toggle_options_menu(self):
        if not self.dropdown_visible:
            self.dropdown_frame.place(x=5, y=30)
            self.dropdown_frame.lift()
            self.dropdown_visible = True
        else:
            self.dropdown_frame.place_forget()
            self.dropdown_visible = False

    def hide_dropdown_option_menu_click_outside(self, widget_parent):
        if self.dropdown_visible:
            allowed_widgets = [self.dropdown_frame, self.options_button]
            while widget_parent:
                if widget_parent in allowed_widgets:
                    return
                widget_parent = widget_parent.master
            self.dropdown_frame.place_forget()
            self.dropdown_visible = False

def center_window(win, width, height):
    screen_width = win.winfo_screenwidth()
    screen_height = win.winfo_screenheight()
    x = (screen_width - width) // 2
    y = (screen_height - height) // 2
    win.geometry(f"{width}x{height}+{x}+{y}")

if __name__ == "__main__":
    app = App()
    app.mainloop()
