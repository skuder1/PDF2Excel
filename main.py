import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from aspose.pdf import Document, ExcelSaveOptions
import webbrowser

def abrir_linkedin():
    webbrowser.open("https://www.linkedin.com/in/pablo-passos-2ba525251/")

class PDFtoExcelConverter:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("PDF to Excel Converter")
        self.window.geometry("400x400")
        self.window.config(bg="#f0f0f0")

        titulo = tk.Label(self.window, text="Conversor de PDF para Excel", font=("Helvetica", 16), bg="#f0f0f0", pady=10)
        titulo.pack()

        self.label_infile = tk.Label(self.window, text="Selecione o arquivo PDF:", bg="#f0f0f0")
        self.label_infile.pack(pady=5)
        self.infile_button = tk.Button(self.window, text="Selecionar PDF", command=self.select_infile, bg="#4CAF50", fg="white")
        self.infile_button.pack(pady=5)

        self.label_outfile = tk.Label(self.window, text="Selecione o local para salvar o arquivo Excel:", bg="#f0f0f0")
        self.label_outfile.pack(pady=5)
        self.outfile_button = tk.Button(self.window, text="Salvar como...", command=self.select_outfile, bg="#4CAF50", fg="white")
        self.outfile_button.pack(pady=5)

        self.convert_button = tk.Button(self.window, text="Converter para Excel", command=self.convert_PDF_to_Excel, bg="#4CAF50", fg="white")
        self.convert_button.pack(pady=20)

        self.progress = ttk.Progressbar(self.window, orient="horizontal", length=300, mode='determinate')
        self.progress.pack(pady=10)

        self.infile = ""
        self.outfile = ""

        # Créditos
        creditos_frame = tk.Frame(self.window, bg="#f0f0f0")
        creditos_frame.pack(pady=20)
        creditos_label = tk.Label(creditos_frame, text="Criado por", font=("Helvetica", 10), bg="#f0f0f0", fg="#333333")
        creditos_label.pack(side="left")
        linkedin_label = tk.Label(creditos_frame, text="Pablo Passos", font=("Helvetica", 10, "underline"), bg="#f0f0f0", fg="blue", cursor="hand2")
        linkedin_label.pack(side="left", padx=5)
        linkedin_label.bind("<Button-1>", lambda e: abrir_linkedin())

        self.window.mainloop()

    def select_infile(self):
        self.infile = filedialog.askopenfilename(title="Selecione o arquivo PDF", filetypes=[("PDF Files", "*.pdf")])
        if self.infile:
            self.label_infile.config(text=f"Arquivo selecionado: {self.infile.split('/')[-1]}")

    def select_outfile(self):
        self.outfile = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
        if self.outfile:
            self.label_outfile.config(text=f"Salvar como: {self.outfile.split('/')[-1]}")

    def convert_PDF_to_Excel(self):
        if not self.infile or not self.outfile:
            messagebox.showwarning("Atenção", "Selecione um arquivo PDF e um local para salvar o arquivo Excel.")
            return

        try:
            self.progress['value'] = 0
            self.window.update_idletasks()

            document = Document(self.infile)
            save_option = ExcelSaveOptions()
            document.save(self.outfile, save_option)

            self.progress['value'] = 100
            messagebox.showinfo("Sucesso", f"{self.infile.split('/')[-1]} foi convertido para {self.outfile.split('/')[-1]} com sucesso.")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro: {str(e)}")
            self.progress['value'] = 0

PDFtoExcelConverter()
