import customtkinter as ctk
import os
from tkinter import filedialog
from tkinter import messagebox
from Tools import center_window,normalize_string
from openpyxl import load_workbook
from Ext_PDF_extraction import PDFExtraction
from Manage_excel_file import ManageExcelFile
from Mark_Sheet import SheetMarker
from Ext_Write import WriteExtraction
from Real_Compute import MonthRealisation
from Real_Write import WriteRealisation
from Styles import StyleCell
from Const import TOOL_SHEET_NAME
from Validation import Valid


GREY = "#999999"
WHITE = "white"
class App(ctk.CTk):
    def __init__(self):
        self.wb, self.ws = None, None
        self.excel_file_path = None
        self.year, self.month = None, None
        self.managment = None
        
        super().__init__()
        ctk.set_appearance_mode("System")  # Modes: "System" (default), "Dark", "Light"
        ctk.set_default_color_theme("blue")  # Thèmes: "blue" (default), "dark-blue", "green"

        # Configurer la fenêtre
        self.title("Gestion des Chemins")

        # Centrer la fenêtre sur l'écran
        center_window(self,800, 600)

        # Variables pour stocker les chemins
        self.pdf_path = ctk.StringVar(value="D:/One Drive/OneDrive/Bureau/R1.pdf")
        self.excel_dir_path = ctk.StringVar(value="D:/.Perso_local/Projects/Project_Money/Test_File")
        self.year_var = ctk.StringVar(value="2024")
        self.year_var.trace_add("write", self.update_combo_box)
        self.month_var = ctk.StringVar(value="juin")
        self.checkbox_create_var = ctk.BooleanVar()
        self.checkbox_update_var = ctk.BooleanVar()

        # Relier la fermeture par la croix à la fonction annuler
        self.protocol("WM_DELETE_WINDOW", self.cancel)
        
        frame_excel = ctk.CTkFrame(self)
        frame_choice = ctk.CTkFrame(self)
        frame_excel.pack()
        frame_choice.pack()
        
        frame_choice_create = ctk.CTkFrame(frame_choice)
        frame_choice_update = ctk.CTkFrame(frame_choice)
        frame_choice_create.pack(side='left')
        frame_choice_update.pack(side='left')
        
        # frame excel
        self.label_excel_dir = ctk.CTkLabel(frame_excel, text="Chemin du dossier Excel:")
        self.label_excel_dir.pack(pady=(20, 0), padx=20, anchor='w')
        self.entry_excel_dir = ctk.CTkEntry(frame_excel, textvariable=self.excel_dir_path, width=400)
        self.entry_excel_dir.pack(pady=10, padx=20, anchor='w')
        self.button_browse_excel_dir = ctk.CTkButton(frame_excel, text="Browse", command=self.browse_excel_dir)
        self.button_browse_excel_dir.pack(pady=10, padx=20, anchor='w')
        
        # frame_choice_create
        self.label_create_title = ctk.CTkLabel(frame_choice_create, text="Create")
        self.label_create_title.pack(pady=(20, 0), padx=20, anchor='w')
        
        self.label_pdf = ctk.CTkLabel(frame_choice_create, text="Chemin du fichier PDF:")
        self.label_pdf.pack(pady=(20, 0), padx=20, anchor='w')
        self.entry_pdf = ctk.CTkEntry(frame_choice_create, textvariable=self.pdf_path, width=400)
        self.entry_pdf.pack(pady=10, padx=20, anchor='w')
        self.button_browse_pdf = ctk.CTkButton(frame_choice_create, text="Browse", command=self.browse_pdf)
        self.button_browse_pdf.pack(pady=10, padx=20, anchor='w')
        
        self.checkbox_create = ctk.CTkCheckBox(frame_choice_create, text='', variable=self.checkbox_create_var, onvalue=True, offvalue=False,command=lambda: self.on_checkbox_toggle(self.checkbox_create_var))
        self.checkbox_create.pack(pady=10, padx=20, anchor='w')
        
        # frame_choice_update
        self.label_update_title = ctk.CTkLabel(frame_choice_update, text="Update")
        self.label_update_title.pack(pady=(20, 0), padx=20, anchor='w')
        
        self.label_year = ctk.CTkLabel(frame_choice_update, text="Year")
        self.label_year.pack(pady=(20, 0), padx=20, anchor='w')
        self.entry_year = ctk.CTkEntry(frame_choice_update, textvariable=self.year_var, width=400)
        self.entry_year.pack(pady=10, padx=20, anchor='w')
        self.label_month = ctk.CTkLabel(frame_choice_update, text="Month")
        self.label_month.pack(pady=(20, 0), padx=20, anchor='w')
        
        self.entry_month = ctk.CTkComboBox(frame_choice_update, variable=self.month_var, width=400,values=[''])
        self.entry_month.bind("<Button-1>", self.update_combo_box)
        self.entry_month.pack(pady=10, padx=20, anchor='w')
        
        self.checkbox_update = ctk.CTkCheckBox(frame_choice_update, text='', variable=self.checkbox_update_var, onvalue=True, offvalue=False,command=lambda: self.on_checkbox_toggle(self.checkbox_update_var))
        self.checkbox_update.pack(pady=10, padx=20, anchor='w')
        self.on_checkbox_toggle(self.checkbox_create_var)
        
        self.button_free = ctk.CTkButton(self, text="Free", command=self.free_year_wb)
        self.button_free.pack(side="right", padx=20, pady=20)
        
        # Boutons OK et Annuler
        self.button_ok = ctk.CTkButton(self, text="OK", command=self.ok)
        self.button_ok.pack(side="right", padx=20, pady=20)

        self.button_cancel = ctk.CTkButton(self, text="Annuler", command=self.cancel)
        self.button_cancel.pack(side="right", padx=10, pady=20)
    
    def disable_update(self):
        self.label_update_title.configure(text_color=GREY)
        self.label_year.configure(text_color=GREY)
        self.entry_year.configure(state="disabled",text_color=GREY)
        self.label_month.configure(text_color=GREY)
        self.entry_month.configure(state="disabled",text_color=GREY)
        self.checkbox_update.configure(state="normal")
        self.checkbox_update_var.set(False)
    def enable_update(self):
        self.label_update_title.configure(text_color=WHITE)
        self.label_year.configure(state="normal",text_color=WHITE)
        self.entry_year.configure(state="normal",text_color=WHITE)
        self.label_month.configure(state="normal",text_color=WHITE)
        self.entry_month.configure(state="normal",text_color=WHITE)
        self.checkbox_update.configure(state="normal")
        self.checkbox_update_var.set(True)
    def disable_create(self):
        self.label_create_title.configure(text_color=GREY)
        self.label_pdf.configure(text_color=GREY)
        self.entry_pdf.configure(state="disabled",text_color=GREY)
        self.button_browse_pdf.configure(state="disabled")
        self.checkbox_create.configure(state="normal")
        self.checkbox_create_var.set(False)
    def enable_create(self):
        self.label_create_title.configure(text_color=WHITE)
        self.label_pdf.configure(state="normal",text_color=WHITE)
        self.entry_pdf.configure(state="normal",text_color=WHITE)
        self.button_browse_pdf.configure(state="normal")
        self.checkbox_create.configure(state="normal")
        self.checkbox_create_var.set(True)
    def glob_enable_create(self):
        self.enable_create()
        self.disable_update()
    def glob_enable_update(self):
        self.enable_update()
        self.disable_create()
    
    ########## COMMAND ##########
    def browse_pdf(self):
        file_path = filedialog.askopenfilename(title="Sélectionner un fichier PDF", filetypes=[("PDF Files", "*.pdf")])
        if file_path:
            self.pdf_path.set(file_path)
    
    def browse_excel_dir(self):
        dir_path = filedialog.askdirectory(title="Sélectionner un dossier")
        if dir_path:
            self.excel_dir_path.set(dir_path)
            
    def get_wb_from_year_var(self):
        year = self.year_var.get()
        excel_file_path = os.path.join(self.excel_dir_path.get(), f'{year}.xlsx')
        return load_workbook(filename=excel_file_path)
        
    def get_existing_sheet_from_year(self):
        try:
            wb = self.get_wb_from_year_var()
            sheet_names = [item for item in wb.sheetnames if item!=TOOL_SHEET_NAME]
            return sheet_names
        except Exception:
            return ['']
        
    def free_year_wb(self):
        try:
            wb = self.get_wb_from_year_var()
            wb.close()
            messagebox.showinfo('Info','Workbook closed')
        except Exception:
            pass
        
    def update_combo_box(self, *args):
        sheet_names = self.get_existing_sheet_from_year()
        self.entry_month.configure(values=sheet_names)
        
    def on_checkbox_toggle(self, checkbox_var):
    # Check if the event was triggered by the "Create" checkbox
        if checkbox_var == self.checkbox_create_var:
                self.glob_enable_create()
        # Check if the event was triggered by the "Update" checkbox
        elif checkbox_var == self.checkbox_update_var:
                self.glob_enable_update()
                self.update_combo_box()
                
    def cancel(self):
        self.destroy()

    def ok(self):
        if self.checkbox_update_var.get() == False and self.checkbox_create_var.get() == True: # Create/Extraction
            extraction = PDFExtraction(self.pdf_path.get())
            self.year = normalize_string(extraction.year)
            self.month = normalize_string(extraction.month)
            self.manage_excel_file("Extraction")
            if self.wb and self.ws:
                if not self.managment.ws_creat:
                    if not messagebox.askyesno("Confirmation", f"This sheet ({self.managment.ws_name}) already exist in {self.managment.file_path}, do you want to update it ?"):
                        return
                StyleCell(self.wb)
                mks = SheetMarker(self.ws)
                WriteExtraction(self.ws,extraction,mks)
                Valid(self.ws,mks)
                self.managment.save_wb()
                messagebox.showinfo("Success", "Operation completed successfully!")
                self.wb.close()
            else:
                messagebox.showerror("Error", "ERROR !")
        
        elif self.checkbox_update_var.get() == True and self.checkbox_create_var.get() == False: # Update/Realisation
            self.year = normalize_string(self.year_var.get())
            self.month = normalize_string(self.month_var.get().lower())
            self.manage_excel_file("Realisation")
            if self.wb:
                if self.ws:
                    if messagebox.askyesno("Confirmation", f"This sheet ({self.managment.ws_name}) already exist in {self.managment.file_path}, do you want to update it ?"):
                        StyleCell(self.wb)
                        mks = SheetMarker(self.ws)
                        realisation = MonthRealisation(self.ws,mks)
                        WriteRealisation(self.ws,mks,realisation)
                        Valid(self.ws,mks)
                        self.managment.save_wb()
                        messagebox.showinfo("Success", "Operation completed successfully!")
                        self.wb.close()
                else:
                    messagebox.showinfo("Info", f"Create the sheet {self.managment.ws_name} with the first extraction before")
            else:
                messagebox.showinfo("Info", f"Build the workbook with the first extraction before")
                    
        else:
            print('ERROR clique ok')

    def manage_excel_file(self,man_type):
        self.managment = ManageExcelFile(self.excel_dir_path.get(),self.year,self.month,man_type)
        self.wb,self.ws = self.managment.get_wb_ws()


if __name__ == "__main__":
    app = App()
    app.mainloop()
