import customtkinter as ctk
from tkinter import Entry

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("System")  # Modes: "System" (default), "Dark", "Light"
        ctk.set_default_color_theme("blue")  # Thèmes: "blue" (default), "dark-blue", "green"

        # Configurer la fenêtre
        self.title("Money")

        # Remove native title bar
        self.overrideredirect(True)
        center_window(self,800, 535)

        # Create custom title bar
        self.create_title_bar()

        self.bind("<Button-3>", self.remove_focus)
        self.bind("<Button-1>", self.remove_focus)

        # Dropdown menu hidden by default
        self.dropdown_frame = None

        # Enable window dragging
        self.offset_x = 0
        self.offset_y = 0

        frame_excel = ctk.CTkFrame(self,fg_color='transparent')
        frame_choice = ctk.CTkFrame(self,fg_color='transparent')
        frame_ok_cancel = ctk.CTkFrame(self,fg_color='transparent')
        frame_excel.pack(pady=(0,15))
        frame_choice.pack()
        frame_ok_cancel.pack(pady=(25,0))
        frame_loading = ctk.CTkFrame(self,fg_color='transparent',width=0, height=0)
        frame_loading.place(relx=0,rely=1,anchor='sw')
        
        frame_left = ctk.CTkFrame(frame_choice)
        frame_choice_extract = ctk.CTkFrame(frame_left,fg_color='transparent')
        frame_choice_create = ctk.CTkFrame(frame_left,fg_color='transparent')
        frame_choice_update = ctk.CTkFrame(frame_choice)
        frame_left.pack(side='left',anchor='center',padx=10)
        frame_choice_update.pack(side='right',anchor='center',padx=10)
        frame_choice_extract.pack()
        frame_choice_create.pack()
        
        WIDTH = 350
        # frame excel
        self.label_excel_dir = ctk.CTkLabel(frame_excel, text="Chemin du dossier Excel:",font=("Arial", 17))
        self.label_excel_dir.pack(pady=(20, 0), padx=20, anchor='w')
        self.entry_excel_dir = ctk.CTkEntry(frame_excel, width=475)
        self.entry_excel_dir.pack(pady=10, padx=20, anchor='w')
        self.button_browse_excel_dir = ctk.CTkButton(frame_excel, text="Browse",
                                                        width=80,height=40,font=("Arial", 17))
        self.button_browse_excel_dir.pack(pady=(5,10), padx=20, anchor='e')
        
        # frame_choice_extract
        self.label_extract_title = ctk.CTkLabel(frame_choice_extract, text="Extraction",font=("Arial", 30,'bold'))
        self.label_extract_title.pack(pady=(10, 0), padx=20, anchor='w',fill='both')
        
        self.entry_pdf = ctk.CTkEntry(frame_choice_extract, width=WIDTH,placeholder_text="PDF Path")
        self.entry_pdf.pack(pady=10, padx=20, anchor='w')
        self.button_browse_pdf = ctk.CTkButton(frame_choice_extract, text="Browse",
                                                width=80,height=40,font=("Arial", 15))
        self.button_browse_pdf.pack(pady=(5,0), padx=20, anchor='e')
        
        self.checkbox_extract = ctk.CTkCheckBox(frame_choice_extract, text='')
        self.checkbox_extract.place(relx=0.96,rely=0.12,anchor='ne')
        
        # frame_choice_create
        frame_choice_create_ym = ctk.CTkFrame(frame_choice_create,fg_color='transparent')
        self.label_create_title = ctk.CTkLabel(frame_choice_create, text="Create",font=("Arial", 30,'bold'))
        self.label_create_title.pack(pady=(15, 0), padx=20,fill='both')
        frame_choice_create_ym.pack(pady=(20, 0),padx=20,fill='both')
        
        self.label_year_create = ctk.CTkLabel(frame_choice_create_ym, text="Year",font=("Arial", 17))
        self.label_year_create.pack(pady=(10, 5),padx=(20,7), side='left',anchor='n')
        self.entry_year_create = ctk.CTkEntry(frame_choice_create_ym, width=WIDTH/5)
        self.entry_year_create.pack(pady=10, side='left',anchor='n')
        self.label_month_create = ctk.CTkLabel(frame_choice_create_ym, text="Month",font=("Arial", 17))
        self.label_month_create.pack(pady=(10, 5),padx=(30,7), side='left',anchor='n')
        self.entry_month_create = ctk.CTkComboBox(frame_choice_create_ym, width=WIDTH/3.3)
        self.entry_month_create.pack(pady=10, side='left',anchor='n')
        
        self.checkbox_create = ctk.CTkCheckBox(frame_choice_create, text='')
        self.checkbox_create.place(relx=0.93,rely=0.19,anchor='ne')
        
        # frame_choice_update
        self.label_update_title = ctk.CTkLabel(frame_choice_update, text="Update",font=("Arial", 30,'bold'))
        self.label_update_title.pack(pady=(20, 0), padx=20, anchor='w',fill='both')
        
        self.label_year = ctk.CTkLabel(frame_choice_update, text="Year",font=("Arial", 17,'bold'))
        self.label_year.pack(pady=(20, 0), padx=20, anchor='w')
        self.entry_year = ctk.CTkEntry(frame_choice_update, width=WIDTH)
        self.entry_year.pack(pady=10, padx=20, anchor='w')
        self.label_month = ctk.CTkLabel(frame_choice_update, text="Month",font=("Arial", 17,'bold'))
        self.label_month.pack(pady=(20, 0), padx=20, anchor='w')
        
        self.entry_month = ctk.CTkComboBox(frame_choice_update, width=WIDTH,values=[''])
        self.entry_month.pack(pady=(10,20), padx=20, anchor='w')
        
        self.checkbox_update = ctk.CTkCheckBox(frame_choice_update, text='')
        self.checkbox_update.place(relx=0.928,rely=0.1,anchor='ne')
        
        
        # Boutons OK et Annuler
        self.button_ok = ctk.CTkButton(frame_ok_cancel, text="OK",width=80,height=40,font=("Arial", 17, "bold"))
        self.button_ok.pack(side="right", padx=20, pady=(25,0))

        self.button_cancel = ctk.CTkButton(frame_ok_cancel, text="Cancel",width=80,height=40,font=("Arial", 17, "bold"))
        self.button_cancel.pack(side="left", padx=10, pady=(25,0))
        
        # GIF
        self.loading_label = ctk.CTkLabel(frame_loading,text='')


    def remove_focus(self,event):
        if  not isinstance(event.widget,Entry):
            self.focus_set()

    def create_title_bar(self):
        """Creates the custom title bar."""
        self.title_bar = ctk.CTkFrame(self, height=30, corner_radius=0, fg_color="#1B1B1B")
        self.title_bar.pack(fill="x", side="top")

        # Application title next to the Options button
        self.title_label = ctk.CTkLabel(self.title_bar, text="Money App", font=("Segoe UI", 12))
        self.title_label.place(relx=0.5, rely=0.5, anchor="center")

        # Options button on the left
        self.options_button = ctk.CTkButton(self.title_bar, text="Option", width=30, command=self.toggle_options_menu, fg_color="transparent", hover_color="#3b3b3b")
        self.options_button.pack(side="left", padx=5)

        self.close_button = ctk.CTkButton(self.title_bar, text="✕", width=45, fg_color="transparent", hover_color="#d32f2f")
        self.close_button.pack(side="right", padx=2)

        # Bind dragging events
        self.title_bar.bind("<ButtonPress-1>", self.start_move)
        self.title_bar.bind("<B1-Motion>", self.do_move)
        self.title_label.bind("<ButtonPress-1>", self.start_move)
        self.title_label.bind("<B1-Motion>", self.do_move)

    def toggle_options_menu(self):
        """Toggles the visibility of the options dropdown menu."""
        if self.dropdown_frame is None:
            self.dropdown_frame = ctk.CTkFrame(self, width=120, fg_color="#2b2b2b")
            # Place under the Options button
            self.dropdown_frame.place(x=5, y=30)

            self.settings_button = ctk.CTkButton(self.dropdown_frame, text="Settings", fg_color="transparent", hover_color="#3b3b3b")
            self.settings_button.pack(fill="x", pady=2)

            self.help_button = ctk.CTkButton(self.dropdown_frame, text="Help", fg_color="transparent", hover_color="#3b3b3b")
            self.help_button.pack(fill="x", pady=2)
        else:
            self.dropdown_frame.destroy()
            self.dropdown_frame = None

    def start_move(self, event):
        """Records the initial absolute mouse position and window position for dragging."""
        self.offset_x = event.x_root
        self.offset_y = event.y_root
        self.start_x = self.winfo_x()
        self.start_y = self.winfo_y()

    def do_move(self, event):
        """Moves the window according to absolute mouse motion."""
        delta_x = event.x_root - self.offset_x
        delta_y = event.y_root - self.offset_y
        new_x = self.start_x + delta_x
        new_y = self.start_y + delta_y
        self.geometry(f'+{new_x}+{new_y}')


    def minimize_window(self):
        """Minimizes the window to the taskbar."""
        self.update_idletasks()
        self.overrideredirect(False)  # Temporarily remove custom titlebar
        self.iconify()

    def open_settings(self):
        """Placeholder for settings functionality."""
        print("Settings selected")

    def open_help(self):
        """Placeholder for help functionality."""
        print("Help selected")

    

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
