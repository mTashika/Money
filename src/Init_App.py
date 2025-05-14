import customtkinter as ctk
from tkinter import Entry,PhotoImage

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("System")  # Modes: "System" (default), "Dark", "Light"
        ctk.set_default_color_theme("blue")  # Thèmes: "blue" (default), "dark-blue", "green"

        # Configurer la fenêtre
        self.title("Money")
        #self.wm_attributes("-toolwindow", 1)
        center_window(self,650, 375)

        self.bind("<Button-3>", self.remove_focus)
        self.bind("<Button-1>", self.remove_focus)

        # Enable window dragging
        self.offset_x = 0
        self.offset_y = 0

        self.CHECK_CREATE = 1
        self.CHECK_EXTRACT = 2
        self.CHECK_UPDATE = 3

        ### FRAME ###
        WIDTH = 300
        # Frame choice
        frame_choice = ctk.CTkFrame(self, fg_color='transparent')
        frame_choice.grid_rowconfigure((0, 1), weight=1)
        frame_choice.grid_columnconfigure(0, weight=1)
        frame_choice.pack(fill='both',pady=(0,20))  # Ensure the frame fills the space
        # Frame ok Cancel
        frame_ok_cancel = ctk.CTkFrame(self, fg_color='transparent')
        frame_ok_cancel.grid_columnconfigure((0,1), weight=1)
        frame_ok_cancel.grid_rowconfigure(0, weight=1)
        frame_ok_cancel.pack()
        # Frame loading
        frame_loading = ctk.CTkFrame(self, fg_color='transparent', width=0, height=0)
        frame_loading.place(relx=0, rely=1, anchor='sw')
        # Frame Top
        frame_top = ctk.CTkFrame(frame_choice, fg_color='#2b2b2b')
        frame_top.grid(row=0, column=0, sticky="ns",pady=(15,15))
        # Frame Bot
        frame_bot = ctk.CTkFrame(frame_choice, fg_color='transparent')
        frame_bot.grid_columnconfigure((0,1), weight=1)
        frame_bot.grid(row=1, column=0, sticky="nsew")
        # Frame Create
        frame_choice_create = frame_top
        # Frame Extract
        frame_choice_extract = ctk.CTkFrame(frame_bot, fg_color='#2b2b2b')
        frame_choice_extract.grid_columnconfigure(0, weight=1)
        frame_choice_extract.grid_rowconfigure((0,1,2), weight=1)
        frame_choice_extract.grid(row=0, column=0, sticky="ens")  # 'ew' ensures horizontal expansion
        # Frame Update
        frame_choice_update = ctk.CTkFrame(frame_bot, fg_color='#2b2b2b')
        frame_choice_update.grid_columnconfigure(0, weight=1)
        frame_choice_update.grid_rowconfigure((0,1), weight=1)
        frame_choice_update.grid(row=0, column=1, sticky="wns",padx=(50,25))

        ### INSIDE FRAME ###
        # Inside frame_choice_create
        frame_choice_create_ym = ctk.CTkFrame(frame_choice_create,fg_color='transparent')
        self.label_create_title = ctk.CTkLabel(frame_choice_create, text="Create",font=("Arial", 30,'bold'))
        self.label_create_title.pack(pady=(5, 0), padx=20,fill='both')
        frame_choice_create_ym.pack(pady=(15, 10),padx=20,fill='both')
        self.label_year_create = ctk.CTkLabel(frame_choice_create_ym, text="Year",font=("Arial", 17))
        self.label_year_create.pack(pady=(10, 0),padx=(0,7), side='left',anchor='n')
        self.entry_year_create = ctk.CTkEntry(frame_choice_create_ym, width=75)
        self.entry_year_create.pack(pady=(10, 0), side='left',anchor='n')
        self.label_month_create = ctk.CTkLabel(frame_choice_create_ym, text="Month",font=("Arial", 17))
        self.label_month_create.pack(pady=(10, 0),padx=(35,7), side='left',anchor='n')
        self.entry_month_create = ctk.CTkComboBox(frame_choice_create_ym, width=120)
        self.entry_month_create.pack(pady=(10, 5), side='left',anchor='n')
        self.checkbox_create = ctk.CTkCheckBox(frame_choice_create, text='',onvalue=self.CHECK_CREATE,offvalue=0)
        self.checkbox_create.place(x=337,y=11,anchor='ne')

        # Inside frame_choice_extract
        self.label_extract_title = ctk.CTkLabel(frame_choice_extract, text="Extraction",font=("Arial", 30,'bold'))
        self.label_extract_title.grid(row=0, column=0, sticky="ew",pady=(10,20),padx=(75,75))
        self.entry_pdf = ctk.CTkEntry(frame_choice_extract,placeholder_text="PDF Path",width=20)
        self.entry_pdf.grid(row=1, column=0, sticky="ew",padx=(20,40),pady=(0,10))
        self.button_browse_pdf = ctk.CTkButton(frame_choice_extract, text="Browse",
                                                width=80,height=40,font=("Arial", 15))
        self.button_browse_pdf.grid(row=2, column=0, sticky="e",padx=(0,20),pady=(0,10))
        self.checkbox_extract = ctk.CTkCheckBox(frame_choice_extract, text='',onvalue=self.CHECK_EXTRACT,offvalue=0)
        self.checkbox_extract.place(x=331,y=16,anchor='ne')
        
        # Inside frame_choice_update
        self.label_update_title = ctk.CTkLabel(frame_choice_update, text="Update",font=("Arial", 30,'bold'))
        self.label_update_title.grid(row=0, column=0, sticky="ew",pady=(2,10),padx=(75,75))
        frame_year_month_update = ctk.CTkFrame(frame_choice_update, fg_color='transparent')
        frame_year_month_update.grid(row=1, column=0, sticky="n")
        frame_year_month_update.grid_columnconfigure((0,1), weight=1)
        frame_year_month_update.grid_rowconfigure((0,1), weight=1)
        self.label_year = ctk.CTkLabel(frame_year_month_update, text="Year",font=("Arial", 17,'bold'))
        self.label_year.grid(row=0, column=0, sticky="ew",padx=(0,12))
        self.entry_year = ctk.CTkEntry(frame_year_month_update, width=75)
        self.entry_year.grid(row=1, column=0, sticky="ns",padx=(0,15),pady=(5,10))
        self.label_month = ctk.CTkLabel(frame_year_month_update, text="Month",font=("Arial", 17,'bold'))
        self.label_month.grid(row=0, column=1, sticky="ew")
        self.entry_month = ctk.CTkComboBox(frame_year_month_update, width=120,values=[''])
        self.entry_month.grid(row=1, column=1, sticky="ns",padx=(15,0),pady=(5,10))
        self.checkbox_update = ctk.CTkCheckBox(frame_choice_update, text='',onvalue=self.CHECK_UPDATE,offvalue=0)
        self.checkbox_update.place(x=286,y=16,anchor='ne')
        
        
        # Inside Boutons OK et Annuler
        self.button_ok = ctk.CTkButton(frame_ok_cancel, text="OK",width=80,height=40,font=("Arial", 17, "bold"))
        self.button_ok.grid(row=0, column=1, sticky="wns", padx=(10,0), pady=(0,0))
        self.button_cancel = ctk.CTkButton(frame_ok_cancel, text="Cancel",width=80,height=40,font=("Arial", 17, "bold"))
        self.button_cancel.grid(row=0, column=0, sticky="ens", padx=(0,10), pady=(0,0))
        
        # Options button
        gear_icon = PhotoImage(file="asset\option.png")
        gear_icon = gear_icon.subsample(15, 15)
        self.options_button = ctk.CTkButton(self, image=gear_icon, text="", width=30, command=self.toggle_options_menu, fg_color="transparent", hover_color="#3b3b3b")
        self.options_button.place(x=0,y=0)
        self.init_option_dropdowm()

        # GIF
        self.loading_label = ctk.CTkLabel(frame_loading,text='')


    def remove_focus(self,event):
        self.hide_dropdown_option_menu_click_outside(event.widget)
        if  not isinstance(event.widget,Entry):
            self.focus_set()

    def init_option_dropdowm(self):
        # Pre-create the dropdown frame and its buttons, but hide them
        self.dropdown_frame = ctk.CTkFrame(self,fg_color="#2b2b2b")
        self.settings_button = ctk.CTkButton(
            self.dropdown_frame,
            text="Settings",
            fg_color="transparent",
            hover_color="#3b3b3b"
        )
        self.settings_button.pack(pady=0,padx=0)
        self.help_button = ctk.CTkButton(
            self.dropdown_frame,
            text="Help",
            fg_color="transparent",
            hover_color="#3b3b3b"
        )
        self.help_button.pack(pady=0,padx=0)
        self.dropdown_visible = False  # track visibility

    def toggle_options_menu(self):
        """Toggles the visibility of the pre-created options dropdown menu."""
        if not self.dropdown_visible:
            self.dropdown_frame.place(x=25, y=25)  # adjust position if neededs
            self.dropdown_frame.lift()
            self.dropdown_visible = True
        else:
            self.dropdown_frame.place_forget()
            self.dropdown_visible = False

    def hide_dropdown_option_menu_click_outside(self,widget_parent):
        """Hide dropdown if click occurs outside dropdown or options button."""
        if self.dropdown_visible:
            # List of widgets to ignore (dropdown + button)
            allowed_widgets = [self.dropdown_frame, self.options_button]
            while widget_parent:
                if widget_parent in allowed_widgets:
                    return  # click inside, do nothing
                widget_parent = widget_parent.master
            self.dropdown_frame.place_forget()
            self.dropdown_visible = False


def center_window(win, width, height):
        # Obtenir la taille de l'écran
        screen_width = win.winfo_screenwidth()
        screen_height = win.winfo_screenheight()

        # Calculer les coordonnées x et y pour centrer la fenêtre
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2

        # Définir la taille et la position de la fenêtre
        win.geometry(f"{width}x{height}+{x}+{y}")

    
if __name__ == "__main__":
    app = App()
    app.mainloop()
