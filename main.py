import customtkinter as ctk
from tkinter import filedialog
from Tools import center_window

from Extract_from_pdf import PDFExtraction
from Manage_excel_file import ManageExcelFile
from Write_fo_in_excel import WriteFo
from Styles import StyleCell

# Initialiser l'application CustomTkinter
ctk.set_appearance_mode("System")  # Modes: "System" (default), "Dark", "Light"
ctk.set_default_color_theme("blue")  # Thèmes: "blue" (default), "dark-blue", "green"

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Configurer la fenêtre
        self.title("Gestion des Chemins")

        # Centrer la fenêtre sur l'écran
        center_window(self,800, 600)

        # Variables pour stocker les chemins
        self.pdf_path = ctk.StringVar(value="D:/One Drive/OneDrive/Bureau/relevejuin.pdf")
        self.excel_dir_path = ctk.StringVar(value="D:/.Perso_local/Projects/Project_Money")
        self.excel_exe_path = ctk.StringVar()


        # Relier la fermeture par la croix à la fonction annuler
        self.protocol("WM_DELETE_WINDOW", self.cancel)

        # Label et champ pour le PDF
        self.label_pdf = ctk.CTkLabel(self, text="Chemin du fichier PDF:")
        self.label_pdf.pack(pady=(20, 0), padx=20, anchor='w')

        self.entry_pdf = ctk.CTkEntry(self, textvariable=self.pdf_path, width=400)
        self.entry_pdf.pack(pady=10, padx=20, anchor='w')

        self.button_browse_pdf = ctk.CTkButton(self, text="Browse", command=self.browse_pdf)
        self.button_browse_pdf.pack(pady=10, padx=20, anchor='w')

        # Label et champ pour le dossier Excel
        self.label_excel_dir = ctk.CTkLabel(self, text="Chemin du dossier Excel:")
        self.label_excel_dir.pack(pady=(20, 0), padx=20, anchor='w')

        self.entry_excel_dir = ctk.CTkEntry(self, textvariable=self.excel_dir_path, width=400)
        self.entry_excel_dir.pack(pady=10, padx=20, anchor='w')

        self.button_browse_excel_dir = ctk.CTkButton(self, text="Browse", command=self.browse_excel_dir)
        self.button_browse_excel_dir.pack(pady=10, padx=20, anchor='w')

        # Label et champ pour l'exécutable Excel
        self.label_excel_exe = ctk.CTkLabel(self, text="Chemin de l'exécutable Excel:")
        self.label_excel_exe.pack(pady=(20, 0), padx=20, anchor='w')

        self.entry_excel_exe = ctk.CTkEntry(self, textvariable=self.excel_exe_path, width=400)
        self.entry_excel_exe.pack(pady=10, padx=20, anchor='w')
        
        # Boutons OK et Annuler
        self.button_ok = ctk.CTkButton(self, text="OK", command=self.ok)
        self.button_ok.pack(side="right", padx=20, pady=20)

        self.button_cancel = ctk.CTkButton(self, text="Annuler", command=self.cancel)
        self.button_cancel.pack(side="right", padx=10, pady=20)

    # Fonction pour sélectionner un fichier PDF
    def browse_pdf(self):
        file_path = filedialog.askopenfilename(title="Sélectionner un fichier PDF", filetypes=[("PDF Files", "*.pdf")])
        if file_path:
            self.pdf_path.set(file_path)

    # Fonction pour sélectionner un dossier Excel
    def browse_excel_dir(self):
        dir_path = filedialog.askdirectory(title="Sélectionner un dossier")
        if dir_path:
            self.excel_dir_path.set(dir_path)

    # Fonction appelée lorsqu'on clique sur "Annuler" ou ferme la fenêtre
    def cancel(self):
        self.destroy()

    def ok(self):
        print('PDF extraction ..')
        self.extract_pdf_data()
        print('PDF extraction - OK')
        self.manage_excel_file()
        if self.wb and self.ws:
            print('Write in sheet ..')
            self.write_in_sheet()
            print('Write in sheet - OK')
            
            self.wb.save(self.excel_file_path)
        
        
    def extract_pdf_data(self):
        extraction = PDFExtraction(self.pdf_path.get())
        self.fo = extraction.financial_operations
        self.month = extraction.month
        self.year = extraction.year
        
    def write_in_sheet(self):
        S = StyleCell(self.wb)
        WriteFo(self.fo,self.ws,S)
        self.ws['A1'] = self.pdf_path.get()
        
    def manage_excel_file(self):
        man = ManageExcelFile(self.excel_dir_path.get(),self.year,self.month)
        self.wb,self.ws = man.get_wb_ws()
        self.excel_file_path = man.file_path


if __name__ == "__main__":
    app = App()
    app.mainloop()
