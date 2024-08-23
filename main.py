
import customtkinter as ctk
from tkinter import filedialog
from Compute_month import compute_code_month

class AskWin(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.place_center_screen(400,250)
        self.resizable(False,False)
        self.title("Excel Sheet App")
        self.res = False

        frame = ctk.CTkFrame(self,fg_color='transparent')  
        frame.columnconfigure = ((0,1,))
        frame.rowconfigure = ((0,1,2,3,))
        frame.pack(padx=(20,20),pady=(0,0))
        
        self.main_label = ctk.CTkLabel(frame,text='Choose the Excel file',font=(None,20))
        self.main_label.grid(row=0, column=0, columnspan=2,padx=10, pady=10, sticky="nsew")
        
        self.entry_path = ctk.CTkEntry(frame, placeholder_text="Excel Path",font=(None,15))
        self.entry_path.grid(row=1, column=0,sticky="nsew",padx=(20,10), pady=(15,10))

        self.button_browse = ctk.CTkButton(frame, text="Browse", command=self.browse_excel,font=(None,15))
        self.button_browse.grid(row=1, column=1,sticky="nse",padx=(10,20), pady=(15,10))

        self.entry_sheet = ctk.CTkEntry(frame, placeholder_text="Sheet Name",font=(None,15))
        self.entry_sheet.grid(row=2, column=0, sticky="nsew",padx=(20,10), pady=(10,10))

        self.button_run = ctk.CTkButton(frame, text="Run", command=self.results,width = 60,height=35,font=(None,15,'bold'))
        self.button_run.grid(row=3, column=1, sticky="sw",padx=(20,10), pady=(35,10))
        
        self.button_cancel = ctk.CTkButton(frame, text="Cancel", command=self.on_close,width = 65,height=35,font=(None,15,'bold'))
        self.button_cancel.grid(row=3, column=0, sticky="se",padx=(20,10), pady=(35,10))
        
        self.mainloop()
        
    def browse_excel(self):
        filename = filedialog.askopenfilename(filetypes=[("Fichiers Excel", "*.xlsx;*.xls;*.xlsm"), ("Tous les fichiers Excel", "*.xlsx;*.xls;*.xlsm")])
        if filename:
            self.entry_path.delete(0, "end")
            self.entry_path.insert(0,filename)
            print("Fichier sélectionné :", filename)
            
    def results(self):
        self.excel_path = self.entry_path.get()
        self.sheet_name = self.entry_sheet.get()
        if self.excel_path and self.sheet_name:
            self.res = True
            self.destroy()
        
    def on_close(self):
        self.excel_path = None
        self.sheet_name = None
        self.destroy()
            
    def place_center_screen(self,w,h):
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = (screen_width - w) // 2
        y = (screen_height - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

if __name__=='__main__':
    win = AskWin()
    if win.res:
        excel_path = 'exemple.xlsx' if win.excel_path == '*' else win.excel_path
        sheet_name = 'Releve' if win.sheet_name == '*' else win.sheet_name
        
        compute_code_month(excel_path,sheet_name)