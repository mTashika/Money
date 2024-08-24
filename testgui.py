import customtkinter as ctk
from tkinter import Toplevel
from Tools import center_window



# Example of how to use AskWindow in a main application
if __name__ == "__main__":
    app = ctk.CTk()
    app.title("Main Application")
    app.geometry("800x600")

    # Function to open the AskWindow
    def open_ask_window():
        AskWindow(app)  # Pass the main application as the parent

    # Button to open the AskWindow
    open_button = ctk.CTkButton(app, text="Open Ask Window", command=open_ask_window)
    open_button.pack(pady=20)

    app.mainloop()
