import customtkinter as ctk
from tkinter import Entry

# Function to center the window on the screen
def center_window(win, width, height):
    # Get screen dimensions
    screen_width = win.winfo_screenwidth()
    screen_height = win.winfo_screenheight()

    # Calculate x and y coordinates to center the window
    x = (screen_width - width) // 2
    y = (screen_height - height) // 2

    # Set window size and position
    win.geometry(f"{width}x{height}+{x}+{y}")

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Set appearance and default theme
        ctk.set_appearance_mode("System")  # Modes: "System", "Dark", "Light"
        ctk.set_default_color_theme("blue")  # Themes: "blue", "dark-blue", "green"

        # Configure the main window
        self.title("Money")
        center_window(self, 800, 535)

        # Bind mouse button events to remove focus from Entry widgets
        self.bind("<Button-3>", self.remove_focus)
        self.bind("<Button-1>", self.remove_focus)

        # Variables for checkbox values
        self.CHECK_CREATE = 1
        self.CHECK_EXTRACT = 2
        self.CHECK_UPDATE = 3

        ### MAIN FRAMES SETUP ###
        WIDTH = 300  # Base width for some widgets

        # Main container frame which fills the entire root window
        self.main_frame = ctk.CTkFrame(self, fg_color='transparent')
        self.main_frame.pack(fill='both', expand=True, padx=10, pady=10)

        # Configure a grid in main_frame for vertical layout: top (for "Create") and bottom (for "Extract" and "Update")
        self.main_frame.grid_rowconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(1, weight=2)
        self.main_frame.grid_columnconfigure(0, weight=1)

        # Top frame for "Create" section
        frame_top = ctk.CTkFrame(self.main_frame, fg_color='transparent')
        frame_top.grid(row=0, column=0, sticky="nsew", pady=(0, 10))

        # Bottom frame for "Extract" and "Update" sections
        frame_bot = ctk.CTkFrame(self.main_frame, fg_color='transparent')
        frame_bot.grid(row=1, column=0, sticky="nsew")
        frame_bot.grid_columnconfigure(0, weight=1)
        frame_bot.grid_columnconfigure(1, weight=1)
        frame_bot.grid_rowconfigure(0, weight=1)

        # Frame for OK and Cancel buttons at the bottom
        frame_ok_cancel = ctk.CTkFrame(self, fg_color='transparent')
        frame_ok_cancel.pack(fill='x', padx=10, pady=(5,10))

        ### INSIDE TOP FRAME: "CREATE" SECTION ###
        # Frame for "Create" title and its inputs
        frame_choice_create = ctk.CTkFrame(frame_top, fg_color='transparent')
        frame_choice_create.pack(fill='both', expand=True, padx=20, pady=10)

        # Create title label
        self.label_create_title = ctk.CTkLabel(frame_choice_create,
                                               text="Create",
                                               font=("Arial", 30, 'bold'))
        self.label_create_title.pack(pady=(20, 10), fill='x')

        # Sub-frame for Year and Month inputs in Create section (compact horizontally)
        frame_create_inputs = ctk.CTkFrame(frame_choice_create, fg_color='transparent')
        frame_create_inputs.pack(fill='x', padx=20, pady=(10, 20))

        # Year label and entry for Create (compact)
        self.label_year_create = ctk.CTkLabel(frame_create_inputs, text="Year", font=("Arial", 17))
        self.label_year_create.grid(row=0, column=0, sticky="w", padx=(0, 5))
        self.entry_year_create = ctk.CTkEntry(frame_create_inputs, width=60)  # Reduced width
        self.entry_year_create.grid(row=0, column=1, sticky="ew", padx=(0, 20))

        # Month label and combobox for Create (compact)
        self.label_month_create = ctk.CTkLabel(frame_create_inputs, text="Month", font=("Arial", 17))
        self.label_month_create.grid(row=0, column=2, sticky="w", padx=(0, 5))
        self.entry_month_create = ctk.CTkComboBox(frame_create_inputs, width=100)  # Adjusted width
        self.entry_month_create.grid(row=0, column=3, sticky="ew")

        # Checkbox for Create section (positioned at top-right of its container)
        self.checkbox_create = ctk.CTkCheckBox(frame_choice_create, text='',
                                               onvalue=self.CHECK_CREATE, offvalue=0)
        self.checkbox_create.place(relx=0.93, rely=0.2, anchor='ne')

        ### INSIDE BOTTOM FRAME: "EXTRACTION" SECTION (Left) ###
        frame_choice_extract = ctk.CTkFrame(frame_bot, fg_color='transparent')
        frame_choice_extract.grid(row=0, column=0, sticky="nsew", padx=(20, 10), pady=10)
        # Configure grid rows for extraction section
        frame_choice_extract.grid_rowconfigure(0, weight=1)
        frame_choice_extract.grid_rowconfigure(1, weight=1)
        frame_choice_extract.grid_rowconfigure(2, weight=1)
        frame_choice_extract.grid_columnconfigure(0, weight=1)

        # Extraction title
        self.label_extract_title = ctk.CTkLabel(frame_choice_extract,
                                                text="Extraction",
                                                font=("Arial", 30, 'bold'))
        self.label_extract_title.grid(row=0, column=0, sticky="ew", pady=(10, 5))

        # Entry for PDF path in Extraction section
        self.entry_pdf = ctk.CTkEntry(frame_choice_extract, placeholder_text="PDF Path")
        self.entry_pdf.grid(row=1, column=0, sticky="ew", padx=20, pady=(5, 5))

        # Browse button for Extraction
        self.button_browse_pdf = ctk.CTkButton(frame_choice_extract,
                                               text="Browse",
                                               width=80,
                                               height=40,
                                               font=("Arial", 15))
        self.button_browse_pdf.grid(row=2, column=0, sticky="e", padx=20, pady=(5, 10))

        # Checkbox for Extraction section, positioned at top-right
        self.checkbox_extract = ctk.CTkCheckBox(frame_choice_extract, text='',
                                                onvalue=self.CHECK_EXTRACT, offvalue=0)
        self.checkbox_extract.place(relx=0.96, rely=0.12, anchor='ne')

        ### INSIDE BOTTOM FRAME: "UPDATE" SECTION (Right) ###
        frame_choice_update = ctk.CTkFrame(frame_bot, fg_color='transparent')
        frame_choice_update.grid(row=0, column=1, sticky="nsew", padx=(10, 20), pady=10)
        frame_choice_update.grid_columnconfigure(0, weight=1)
        # Configure two rows with different weights (title row and input row)
        frame_choice_update.grid_rowconfigure(0, weight=1)
        frame_choice_update.grid_rowconfigure(1, weight=2)

        # Update title label
        self.label_update_title = ctk.CTkLabel(frame_choice_update,
                                               text="Update",
                                               font=("Arial", 30, 'bold'))
        self.label_update_title.grid(row=0, column=0, sticky="ew", pady=(10, 5))

        # Frame for Year and Month inputs in Update section (reduced size)
        frame_year_month_update = ctk.CTkFrame(frame_choice_update, fg_color='transparent')
        frame_year_month_update.grid(row=1, column=0, sticky="nsew", padx=20, pady=(5, 10))
        frame_year_month_update.grid_columnconfigure(0, weight=1)
        frame_year_month_update.grid_columnconfigure(1, weight=1)
        frame_year_month_update.grid_rowconfigure(0, weight=1)
        frame_year_month_update.grid_rowconfigure(1, weight=1)

        # Year label and entry for Update (smaller width)
        self.label_year = ctk.CTkLabel(frame_year_month_update, text="Year", font=("Arial", 17, 'bold'))
        self.label_year.grid(row=0, column=0, sticky="ew", padx=(0, 5))
        self.entry_year = ctk.CTkEntry(frame_year_month_update, width=60)  # Reduced width
        self.entry_year.grid(row=1, column=0, sticky="ew", padx=(0, 5), pady=(5, 10))

        # Month label and combobox for Update (smaller width)
        self.label_month = ctk.CTkLabel(frame_year_month_update, text="Month", font=("Arial", 17, 'bold'))
        self.label_month.grid(row=0, column=1, sticky="ew", padx=(5, 0))
        self.entry_month = ctk.CTkComboBox(frame_year_month_update, width=80, values=[''])  # Smaller width
        self.entry_month.grid(row=1, column=1, sticky="ew", padx=(5, 0), pady=(5, 10))

        # Checkbox for Update section, positioned at top-right of Update frame
        self.checkbox_update = ctk.CTkCheckBox(frame_choice_update, text='',
                                               onvalue=self.CHECK_UPDATE, offvalue=0)
        self.checkbox_update.place(relx=0.93, rely=0.12, anchor='ne')

        ### BOTTOM BUTTONS: OK and Cancel ###
        self.button_ok = ctk.CTkButton(frame_ok_cancel,
                                       text="OK",
                                       width=80,
                                       height=40,
                                       font=("Arial", 17, "bold"))
        self.button_ok.pack(side="right", padx=15, pady=5)
        self.button_cancel = ctk.CTkButton(frame_ok_cancel,
                                           text="Cancel",
                                           width=80,
                                           height=40,
                                           font=("Arial", 17, "bold"))
        self.button_cancel.pack(side="left", padx=15, pady=5)

        ### OPTIONS BUTTON (using place for overlay) ###
        self.options_button = ctk.CTkButton(self,
                                            text="Option",
                                            width=30,
                                            command=self.toggle_options_menu,
                                            fg_color="transparent",
                                            hover_color="#3b3b3b")
        self.options_button.place(x=10, y=10)

        # Initialize hidden dropdown options
        self.init_option_dropdown()


    def remove_focus(self, event):
        # Hide dropdown if click occurs outside
        self.hide_dropdown_option_menu_click_outside(event.widget)
        # If the event widget is not an Entry, set focus to the main window
        if not isinstance(event.widget, Entry):
            self.focus_set()

    def init_option_dropdown(self):
        # Pre-create the dropdown frame and its buttons, but hide them initially
        self.dropdown_frame = ctk.CTkFrame(self, width=120, fg_color="#2b2b2b")
        self.settings_button = ctk.CTkButton(self.dropdown_frame,
                                             text="Settings",
                                             fg_color="transparent",
                                             hover_color="#3b3b3b")
        self.settings_button.pack(fill="x", pady=2)
        self.help_button = ctk.CTkButton(self.dropdown_frame,
                                         text="Help",
                                         fg_color="transparent",
                                         hover_color="#3b3b3b")
        self.help_button.pack(fill="x", pady=2)
        self.dropdown_visible = False  # Track visibility of the dropdown

    def toggle_options_menu(self):
        """Toggle the visibility of the options dropdown menu."""
        if not self.dropdown_visible:
            self.dropdown_frame.place(x=5, y=30)  # Adjust position if needed
            self.dropdown_frame.lift()
            self.dropdown_visible = True
        else:
            self.dropdown_frame.place_forget()
            self.dropdown_visible = False

    def hide_dropdown_option_menu_click_outside(self, widget_parent):
        """Hide dropdown if click occurs outside the dropdown or options button."""
        if self.dropdown_visible:
            # List of widgets to ignore (dropdown frame and options button)
            allowed_widgets = [self.dropdown_frame, self.options_button]
            while widget_parent:
                if widget_parent in allowed_widgets:
                    return  # Click inside allowed widgets, do nothing
                widget_parent = widget_parent.master
            self.dropdown_frame.place_forget()
            self.dropdown_visible = False

if __name__ == "__main__":
    app = App()
    app.mainloop()
