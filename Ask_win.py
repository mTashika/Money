import customtkinter as ctk
from tkinter import Toplevel
from Tools import center_window

class AskWindow(ctk.CTkToplevel):
    def __init__(self, parent, *args, **kwargs):
        super().__init__(*args, **kwargs)
        
        # Set the parent window
        self.parent = parent
        
        # Initialize the window
        self.title("Choose Action")
        center_window(self, 800, 400)
        self.attributes('-topmost', True)
        
        # Create a frame for the content
        frame = ctk.CTkFrame(self)
        frame.pack(pady=20, padx=20, expand=True, fill="both")
        
        # Add a label asking the question
        label = ctk.CTkLabel(frame, text="Do you want to recreate or update the tab?")
        label.pack(pady=10)
        
        # Create buttons and bind them to the respective functions
        button_recreate = ctk.CTkButton(frame, text="Recreate", command=self.on_recreate)
        button_recreate.pack(side="left", padx=10)
        
        button_update = ctk.CTkButton(frame, text="Update", command=self.on_update)
        button_update.pack(side="left", padx=10)
        
        button_cancel = ctk.CTkButton(frame, text="Cancel", command=self.on_cancel)
        button_cancel.pack(side="left", padx=10)

    def on_recreate(self):
        self.result = 1
        self.destroy()  # Close the window

    def on_update(self):
        self.result = 2
        self.destroy()  # Close the window

    def on_cancel(self):
        self.result = 0
        self.destroy()  # Close the window