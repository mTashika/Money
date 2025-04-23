import customtkinter as ctk
import os
from tkinter import filedialog,messagebox,PhotoImage,Entry
from Tools import center_window,normalize_string,get_current_month,get_current_year,set_all_data_validation,run_excel_path
from openpyxl import load_workbook
from Ext_PDF_extraction import PDFExtraction
from Manage_excel_file import ManageExcelFile
from Mark_Sheet import SheetMarker
from Ext_Write import WriteExtraction
from Real_Compute import MonthRealisation
from Real_Write import WriteRealisation
from Styles import StyleCell
from Const import TOOL_SHEET_NAME,MONTH_FR
from Validation import Valid
import threading
from PIL import Image, ImageTk
import sys

CHECK_CREATE = 1
CHECK_EXTRACT = 2
CHECK_UPDATE = 3
GREY = "#999999"
WHITE = "white"

def ressource_path(rel_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path,rel_path)
        
class App(ctk.CTk):
    def __init__(self):
        self.asset_path = ressource_path('asset')        
        self.wb, self.ws = None, None
        self.excel_file_path = None
        self.year, self.month = None, None
        self.managment = None
        
        super().__init__()
        ctk.set_appearance_mode("System")  # Modes: "System" (default), "Dark", "Light"
        ctk.set_default_color_theme("blue")  # Thèmes: "blue" (default), "dark-blue", "green"

        # Configurer la fenêtre
        self.title("Money")

        # Centrer la fenêtre sur l'écran
        center_window(self,800, 535)

        # Variables pour stocker les chemins
        self.pdf_path = ctk.StringVar(value="")
        self.excel_dir_path = ctk.StringVar(value=r"C:\Users\mcast\OneDrive\Documents\.Perso\[4] Finance\Compte\Fiche")
        self.year_var = ctk.StringVar(value=get_current_year())
        self.year_var.trace_add("write", self.update_combo_box)
        self.month_var = ctk.StringVar(value=get_current_month())
        self.checkbox_extract_var = ctk.IntVar()
        self.checkbox_update_var = ctk.IntVar()
        self.checkbox_create_var = ctk.IntVar()
        self.year_create = ctk.StringVar(value=get_current_year())
        self.month_create = ctk.StringVar(value=get_current_month())
        

        # Relier la fermeture par la croix à la fonction annuler
        self.protocol("WM_DELETE_WINDOW", self.cancel)
        self.bind("<Button-3>", self.remove_focus)
        self.bind("<Button-1>", self.remove_focus)
        
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
        self.entry_excel_dir = ctk.CTkEntry(frame_excel, textvariable=self.excel_dir_path, width=475)
        self.entry_excel_dir.pack(pady=10, padx=20, anchor='w')
        self.button_browse_excel_dir = ctk.CTkButton(frame_excel, text="Browse", command=self.browse_excel_dir,
                                                        width=80,height=40,font=("Arial", 17))
        self.button_browse_excel_dir.pack(pady=(5,10), padx=20, anchor='e')
        
        # frame_choice_extract
        self.label_extract_title = ctk.CTkLabel(frame_choice_extract, text="Extraction",font=("Arial", 30,'bold'))
        self.label_extract_title.pack(pady=(10, 0), padx=20, anchor='w',fill='both')
        
        self.entry_pdf = ctk.CTkEntry(frame_choice_extract, textvariable=self.pdf_path, width=WIDTH,placeholder_text="PDF Path")
        self.entry_pdf.pack(pady=10, padx=20, anchor='w')
        self.button_browse_pdf = ctk.CTkButton(frame_choice_extract, text="Browse", command=self.browse_pdf,
                                                width=80,height=40,font=("Arial", 15))
        self.button_browse_pdf.pack(pady=(5,0), padx=20, anchor='e')
        
        self.checkbox_extract = ctk.CTkCheckBox(frame_choice_extract, text='', variable=self.checkbox_extract_var, onvalue=CHECK_EXTRACT, offvalue=0,command=lambda: self.on_checkbox_toggle(self.checkbox_extract_var))
        self.checkbox_extract.place(relx=0.96,rely=0.12,anchor='ne')
        
        # frame_choice_create
        frame_choice_create_ym = ctk.CTkFrame(frame_choice_create,fg_color='transparent')
        self.label_create_title = ctk.CTkLabel(frame_choice_create, text="Create",font=("Arial", 30,'bold'))
        self.label_create_title.pack(pady=(15, 0), padx=20,fill='both')
        frame_choice_create_ym.pack(pady=(20, 0),padx=20,fill='both')
        
        self.label_year_create = ctk.CTkLabel(frame_choice_create_ym, text="Year",font=("Arial", 17))
        self.label_year_create.pack(pady=(10, 5),padx=(20,7), side='left',anchor='n')
        self.entry_year_create = ctk.CTkEntry(frame_choice_create_ym, textvariable=self.year_create, width=WIDTH/5)
        self.entry_year_create.pack(pady=10, side='left',anchor='n')
        self.label_month_create = ctk.CTkLabel(frame_choice_create_ym, text="Month",font=("Arial", 17))
        self.label_month_create.pack(pady=(10, 5),padx=(30,7), side='left',anchor='n')
        self.entry_month_create = ctk.CTkComboBox(frame_choice_create_ym, variable=self.month_create, width=WIDTH/3.3,values=MONTH_FR)
        self.entry_month_create.pack(pady=10, side='left',anchor='n')
        
        self.checkbox_create = ctk.CTkCheckBox(frame_choice_create, text='', variable=self.checkbox_create_var, onvalue=CHECK_CREATE, offvalue=0,command=lambda: self.on_checkbox_toggle(self.checkbox_create_var))
        self.checkbox_create.place(relx=0.93,rely=0.19,anchor='ne')
        
        # frame_choice_update
        self.label_update_title = ctk.CTkLabel(frame_choice_update, text="Update",font=("Arial", 30,'bold'))
        self.label_update_title.pack(pady=(20, 0), padx=20, anchor='w',fill='both')
        
        self.label_year = ctk.CTkLabel(frame_choice_update, text="Year",font=("Arial", 17,'bold'))
        self.label_year.pack(pady=(20, 0), padx=20, anchor='w')
        self.entry_year = ctk.CTkEntry(frame_choice_update, textvariable=self.year_var, width=WIDTH)
        self.entry_year.pack(pady=10, padx=20, anchor='w')
        self.label_month = ctk.CTkLabel(frame_choice_update, text="Month",font=("Arial", 17,'bold'))
        self.label_month.pack(pady=(20, 0), padx=20, anchor='w')
        
        self.entry_month = ctk.CTkComboBox(frame_choice_update, variable=self.month_var, width=WIDTH,values=[''])
        self.entry_month.pack(pady=(10,20), padx=20, anchor='w')
        
        self.checkbox_update = ctk.CTkCheckBox(frame_choice_update, text='', variable=self.checkbox_update_var, onvalue=CHECK_UPDATE, offvalue=0,command=lambda: self.on_checkbox_toggle(self.checkbox_update_var))
        self.checkbox_update.place(relx=0.928,rely=0.1,anchor='ne')
        self.on_checkbox_toggle(self.checkbox_extract_var)
        
        icon_photo = PhotoImage(file=os.path.join(self.asset_path,'folder.png'))
        icon_photo = icon_photo.subsample(2)  # Diviser par 2, ajustez selon vos besoins
        self.button_free = ctk.CTkButton(self, text="", command=self.free_year_wb,image=icon_photo,width=50,height=40)
        self.button_free.place(relx=0.03,rely=0.03,anchor='nw')
        
        # Boutons OK et Annuler
        self.button_ok = ctk.CTkButton(frame_ok_cancel, text="OK", command=self.ok,width=80,height=40,font=("Arial", 17, "bold"))
        self.button_ok.pack(side="right", padx=20, pady=(25,0))

        self.button_cancel = ctk.CTkButton(frame_ok_cancel, text="Cancel", command=self.cancel,width=80,height=40,font=("Arial", 17, "bold"))
        self.button_cancel.pack(side="left", padx=10, pady=(25,0))
        
        # GIF
        self.gif = Image.open(os.path.join(self.asset_path,'loading_gif.gif'))
        self.init_gif()
        self.loading_label = ctk.CTkLabel(frame_loading,text='')
        
        
    
    def disable_update(self):
        self.label_update_title.configure(text_color=GREY)
        self.label_year.configure(text_color=GREY)
        self.entry_year.configure(state="disabled",text_color=GREY)
        self.label_month.configure(text_color=GREY)
        self.entry_month.configure(state="disabled",text_color=GREY)
        self.checkbox_update.configure(state="normal")
        self.checkbox_update_var.set(0)
    def enable_update(self):
        self.label_update_title.configure(text_color=WHITE)
        self.label_year.configure(state="normal",text_color=WHITE)
        self.entry_year.configure(state="normal",text_color=WHITE)
        self.label_month.configure(state="normal",text_color=WHITE)
        self.entry_month.configure(state="normal",text_color=WHITE)
        self.checkbox_update.configure(state="normal")
        self.checkbox_update_var.set(CHECK_UPDATE)
    def disable_create(self):
        self.label_create_title.configure(text_color=GREY)
        self.label_year_create.configure(text_color=GREY)
        self.entry_year_create.configure(state="disabled",text_color=GREY)
        self.label_month_create.configure(text_color=GREY)
        self.entry_month_create.configure(state="disabled",text_color=GREY)
        self.checkbox_create.configure(state="normal")
        self.checkbox_create_var.set(0)
    def enable_create(self):
        self.label_create_title.configure(text_color=WHITE)
        self.label_year_create.configure(state="normal",text_color=WHITE)
        self.entry_year_create.configure(state="normal",text_color=WHITE)
        self.label_month_create.configure(state="normal",text_color=WHITE)
        self.entry_month_create.configure(state="normal",text_color=WHITE)
        self.checkbox_create.configure(state="normal")
        self.checkbox_create_var.set(CHECK_CREATE)
    
    def disable_extract(self):
        self.label_extract_title.configure(text_color=GREY)
        self.entry_pdf.configure(state="disabled",text_color=GREY)
        self.button_browse_pdf.configure(state="disabled")
        self.checkbox_extract.configure(state="normal")
        self.checkbox_extract_var.set(0)
    def enable_extract(self):
        self.label_extract_title.configure(text_color=WHITE)
        self.entry_pdf.configure(state="normal",text_color=WHITE)
        self.button_browse_pdf.configure(state="normal")
        self.checkbox_extract.configure(state="normal")
        self.checkbox_extract_var.set(CHECK_EXTRACT)
    
    def glob_enable_extract(self):
        self.enable_extract()
        self.disable_create()
        self.disable_update()
        self.update()
    def glob_enable_create(self):
        self.enable_create()
        self.disable_update()
        self.disable_extract()
        self.update()
    def glob_enable_update(self):
        self.enable_update()
        self.disable_extract()
        self.disable_create()
        self.update()

        
    def disable_excel(self):
        self.label_excel_dir.configure(text_color=GREY)
        self.entry_excel_dir.configure(state="disabled",text_color=GREY)
        self.button_browse_excel_dir.configure(state="disabled")
    def enable_excel(self):
        self.label_excel_dir.configure(text_color=WHITE)
        self.entry_excel_dir.configure(state="normal",text_color=WHITE)
        self.button_browse_excel_dir.configure(state="normal")
        
    def init_gif(self):
        self.frames = []
        target_size = (45, 45)  # Définir la taille souhaitée pour le GIF
        try:
            while True:
                # Convertir chaque frame en RGBA pour gérer la transparence
                frame = self.gif.convert("RGBA")
                # Redimensionner la frame
                frame = frame.resize(target_size)
                self.frames.append(ImageTk.PhotoImage(frame))
                self.gif.seek(len(self.frames))  # Passer à la prochaine frame
        except EOFError:
            pass  # Fin du GIF
        self.loading_index = 0  # Initialiser l'index de frame
        self.is_paused = False
    
    
    def remove_focus(self,event):
        if  not isinstance(event.widget,Entry):
            self.focus_set()
    
    ########## COMMAND ##########
    def browse_pdf(self):
        file_paths = filedialog.askopenfilenames(title="Sélectionner un fichier PDF", filetypes=[("PDF Files", "*.pdf")])
        if file_paths:
            self.pdf_path.set("; ".join(file_paths))
    
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
            threading.Thread(target=wb.close).start()
            messagebox.showinfo('Info','Workbook closed')
        except Exception:
            pass
        
    def update_combo_box(self, *args):
        sheet_names = self.get_existing_sheet_from_year()
        self.entry_month.configure(values=sheet_names)
        
    def on_checkbox_toggle(self, checkbox_var):
    # Check if the event was triggered by the "Extract" checkbox
        if checkbox_var == self.checkbox_extract_var:
                self.glob_enable_extract()
    # Check if the event was triggered by the "Create" checkbox
        elif checkbox_var == self.checkbox_create_var:
                self.glob_enable_create()
    # Check if the event was triggered by the "Update" checkbox
        elif checkbox_var == self.checkbox_update_var:
                self.glob_enable_update()
                self.update_combo_box()
                
    def cancel(self):
        self.destroy()

    def ok(self):
        if (self.checkbox_update_var.get() == 0 and self.checkbox_create_var.get() == 0 and
                                    self.checkbox_extract_var.get() == CHECK_EXTRACT): # Extraction (Create)
            self.show_loading()
            threading.Thread(target=self.launch_creation_extraction).start()
        elif (self.checkbox_update_var.get() == 0 and self.checkbox_create_var.get() == CHECK_CREATE and
                                    self.checkbox_extract_var.get() == 0): # Create
            self.show_loading()
            threading.Thread(target=self.launch_creation).start()
        elif (self.checkbox_update_var.get() == CHECK_UPDATE and self.checkbox_create_var.get() == 0 and
                                    self.checkbox_extract_var.get() == 0): # Update/Realisation
            self.show_loading()
            threading.Thread(target=self.launch_realisation).start()
            
    def launch_creation(self):
        self.is_paused = False
        self.year = normalize_string(self.year_create.get())
        self.month = normalize_string(self.month_create.get())
        self.manage_excel_file("Extraction")
        if self.wb and self.ws:
            if not self.managment.ws_creat:
                messagebox.showinfo("Info", f"This sheet ({self.managment.ws_name}) already exist in {self.managment.file_path}")
            else:
                mks = SheetMarker(self.ws)
                set_all_data_validation(self.wb,mks)
                is_save_ok = self.managment.save_wb()
                self.is_paused = True
                messagebox.showinfo("Success", f"Sheet ({self.managment.ws_name}) created here: {self.managment.file_path}") if is_save_ok else None
        else:
            self.is_paused = True
            messagebox.showerror("Error", f"ERROR during Creation for {self.managment.ws_name}!")
        self.after(0, self.hide_loading)
        self.wb.close()
        
    def launch_creation_extraction(self):
        file_paths_list = [path.strip() for path in self.pdf_path.get().split(';')]
        for path_pdf in file_paths_list:
            self.is_paused = False
            extraction = PDFExtraction(path_pdf)
            self.year = normalize_string(extraction.year)
            self.month = normalize_string(extraction.month)
            self.manage_excel_file("Extraction")
            f_update = True
            if self.wb and self.ws:
                if not self.managment.ws_creat:
                    self.is_paused = True
                    if not messagebox.askyesno("Confirmation", f"This sheet ({self.managment.ws_name}) already exist in {self.managment.file_path}, do you want to update it ?"):
                        f_update = False
                if f_update:
                    self.is_paused = False
                    StyleCell(self.wb)
                    mks = SheetMarker(self.ws)
                    WriteExtraction(self.ws,extraction,mks)
                    Valid(self.ws,mks)
                    set_all_data_validation(self.wb,mks)
                    is_save_ok = self.managment.save_wb()
                    self.is_paused = True
                    messagebox.showinfo("Success", f"Extraction for {self.managment.ws_name} completed successfully!") if is_save_ok else None
            else:
                self.is_paused = True
                messagebox.showerror("Error", f"ERROR during Extraction for {self.managment.ws_name}!")
        self.after(0, self.hide_loading)
        self.wb.close()
        run_excel_path(self.file_path)
        return self.ws

    def launch_realisation(self):
        self.is_paused = False
        self.year = normalize_string(self.year_var.get())
        self.month = normalize_string(self.month_var.get().lower())
        self.manage_excel_file("Realisation")
        if self.wb:
            if self.ws:
                self.is_paused = True
                if messagebox.askyesno("Confirmation", f"{self.managment.ws_name} sheet already exist here:\n{self.managment.file_path}\n\nDo you want to update it ?"):
                    self.is_paused = False
                    StyleCell(self.wb)
                    mks = SheetMarker(self.ws)
                    realisation_month = MonthRealisation(self.ws,mks,"Monthly")
                    realisation_pdf = MonthRealisation(self.ws,mks,"Pdf")
                    realisation_intern = MonthRealisation(self.ws,mks,"Intern")
                    WriteRealisation(self.ws,mks,realisation_month,realisation_pdf,realisation_intern)
                    Valid(self.ws,mks)
                    set_all_data_validation(self.wb,mks)
                    is_save_ok = self.managment.save_wb()
                    self.is_paused = True
                    self.after(0, self.hide_loading)
                    messagebox.showinfo("Success", f"Operation completed successfully for {self.managment.ws_name}!") if is_save_ok else None
                    self.wb.close()
                    run_excel_path(self.file_path)
            else:
                self.after(0, self.hide_loading)
                messagebox.showinfo("Info", f"Create the sheet {self.month} before")
        else:
            self.after(0, self.hide_loading)
            messagebox.showinfo("Info", f"Build the workbook {self.year} before")
        
    def manage_excel_file(self,man_type):
        self.managment = ManageExcelFile(self.excel_dir_path.get(),self.year,self.month,man_type)
        self.wb,self.ws,self.file_path = self.managment.get_wb_ws_fpath()

    def show_loading(self):
        self.loading_label.pack(padx=10,pady=10)
        self.loading_label.configure(image= self.frames[self.loading_index])
        self.update()
        self.animate_gif()
        self.disable_excel()
        self.disable_update()
        self.disable_extract()
        
    def hide_loading(self):
        self.loading_label.pack_forget()
        self.enable_excel()
        if self.checkbox_update_var.get() == CHECK_UPDATE:
            self.glob_enable_update()
        elif self.checkbox_extract_var.get() == CHECK_EXTRACT:
            self.glob_enable_extract()
            
    def animate_gif(self):
            if not self.is_paused:
                frame = self.frames[self.loading_index]  # Obtenir la frame courante
                self.loading_label.configure(image=frame)  # Mettre à jour l'image du label
                self.update()
                self.loading_index = (self.loading_index + 1) % len(self.frames)  # Passer à la frame suivante
            # Reprogrammer pour afficher la prochaine frame
            self.after(75, self.animate_gif)


if __name__ == "__main__":
    app = App()
    app.mainloop()
