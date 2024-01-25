import tkinter as tk
from tkinter import filedialog
from converter import Converter
import ctypes

class ConvertToExcelApp:
    def __init__(self, master):
        self.master = master
        self.master.title("Conversor de Arquivo DB para Excel")
        self.master.tk_setPalette(background='#7B68EE',foreground='White')
        self.definir_icone_barra_tarefas("./resources/transition.ico")
   
        self.master.geometry("500x70")
        self.master.iconbitmap("./resources/transition.ico")
        
        label = tk.Label(self.master, text="Selecione um arquivo .db:")
        label.config(font=('Helvetica',14,'bold'))
        label.pack()
        
        btn_selecionar_arquivo = tk.Button(self.master, text="Selecionar Arquivo", command=self.selecionar_arquivo)
        btn_selecionar_arquivo.pack()
        
    def selecionar_arquivo(self):
        arquivo_db = filedialog.askopenfilename(filetypes=[("Arquivos DB","*.db")])
        if arquivo_db:
            converter = Converter()
            converter.converter_para_excel(arquivo_db)
            
    @staticmethod
    def definir_icone_barra_tarefas(caminho_icone):
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(caminho_icone)
        


root = tk.Tk()
app = ConvertToExcelApp(root)
root.mainloop()
