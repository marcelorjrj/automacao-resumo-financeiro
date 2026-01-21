import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from fpdf import FPDF
import os
from datetime import datetime

class FinancialAnalyzer:
    def __init__(self, root):
        self.root = root
        self.root.title("Analisador Financeiro")
        self.root.geometry("600x500")
        self.root.resizable(False, False)
        
        self.arquivo_excel = None
        self.total_entradas = 0
        self.total_saidas = 0
        self.saldo = 0
        
        self.criar_interface()
    
    def criar_interface(self):
        # Frame principal
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Título
        titulo = ttk.Label(main_frame, text="Análise de Controle Financeiro", 
                          font=("Arial", 16, "bold"))
        titulo.grid(row=0, column=0, columnspan=2, pady=(0, 20))
        
        # Botão para selecionar arquivo
        btn_selecionar = ttk.Button(main_frame, text="Selecionar Arquivo Excel", 
                                    command=self.selecionar_arquivo, width=30)
        btn_selecionar.grid(row=1, column=0, columnspan=2, pady=10)
        
        # Label para mostrar arquivo selecionado
        self.lbl_arquivo = ttk.Label(main_frame, text="Nenhum arquivo selecionado", 
                                     foreground="gray")
        self.lbl_arquivo.grid(row=2, column=0, columnspan=2, pady=(0, 20))
        
        # Frame para resultados
        resultado_frame = ttk.LabelFrame(main_frame, text="Resultados", padding="15")
        resultado_frame.grid(row=3, column=0, columnspan=2, pady=10, sticky=(tk.W, tk.E))
        
        # Labels de resultados
        ttk.Label(resultado_frame, text="Total Entradas:", 
                 font=("Arial", 10, "bold")).grid(row=0, column=0, sticky=tk.W, pady=5)
        self.lbl_entradas = ttk.Label(resultado_frame, text="R$ 0,00", 
                                      font=("Arial", 10), foreground="green")
        self.lbl_entradas.grid(row=0, column=1, sticky=tk.E, pady=5, padx=(20, 0))
        
        ttk.Label(resultado_frame, text="Total Saídas:", 
                 font=("Arial", 10, "bold")).grid(row=1, column=0, sticky=tk.W, pady=5)
        self.lbl_saidas = ttk.Label(resultado_frame, text="R$ 0,00", 
                                    font=("Arial", 10), foreground="red")
        self.lbl_saidas.grid(row=1, column=1, sticky=tk.E, pady=5, padx=(20, 0))
        
        ttk.Separator(resultado_frame, orient='horizontal').grid(row=2, column=0, 
                                                                 columnspan=2, 
                                                                 sticky=(tk.W, tk.E), 
                                                                 pady=10)
        
        ttk.Label(resultado_frame, text="Saldo Final:", 
                 font=("Arial", 12, "bold")).grid(row=3, column=0, sticky=tk.W, pady=5)
        self.lbl_saldo = ttk.Label(resultado_frame, text="R$ 0,00", 
                                   font=("Arial", 12, "bold"))
        self.lbl_saldo.grid(row=3, column=1, sticky=tk.E, pady=5, padx=(20, 0))
        
        # Configurar colunas do frame de resultados
        resultado_frame.columnconfigure(1, weight=1)
        
        # Botão para salvar PDF
        self.btn_salvar = ttk.Button(main_frame, text="Salvar Relatório em PDF", 
                                     command=self.salvar_pdf, width=30, state="disabled")
        self.btn_salvar.grid(row=4, column=0, columnspan=2, pady=20)
        
        # Barra de status
        self.status_bar = ttk.Label(main_frame, text="Aguardando seleção de arquivo...", 
                                    relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.grid(row=5, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))
    
    def selecionar_arquivo(self):
        arquivo = filedialog.askopenfilename(
            title="Selecione o arquivo Excel",
            filetypes=[("Arquivos Excel", "*.xlsx *.xls"), ("Todos os arquivos", "*.*")]
        )
        
        if arquivo:
            self.arquivo_excel = arquivo
            nome_arquivo = os.path.basename(arquivo)
            self.lbl_arquivo.config(text=f"Arquivo: {nome_arquivo}", foreground="blue")
            self.processar_dados()
    
    def processar_dados(self):
        try:
            self.status_bar.config(text="Processando dados...")
            self.root.update()
            
            # Ler Saídas
            df_saidas = pd.read_excel(self.arquivo_excel, sheet_name="Saídas - Dez 2025")
            df_saidas["Valor"] = pd.to_numeric(df_saidas["Valor"], errors="coerce")
            self.total_saidas = df_saidas["Valor"].sum()
            
            # Ler Entradas
            df_entradas = pd.read_excel(self.arquivo_excel, sheet_name="Entradas - Dez 2025")
            df_entradas["Valor"] = pd.to_numeric(df_entradas["Valor"], errors="coerce")
            self.total_entradas = df_entradas["Valor"].sum()
            
            # Calcular Saldo
            self.saldo = self.total_entradas - self.total_saidas
            
            # Atualizar interface
            self.lbl_entradas.config(text=f"R$ {self.total_entradas:,.2f}")
            self.lbl_saidas.config(text=f"R$ {self.total_saidas:,.2f}")
            self.lbl_saldo.config(text=f"R$ {self.saldo:,.2f}")
            
            # Mudar cor do saldo baseado no valor
            if self.saldo >= 0:
                self.lbl_saldo.config(foreground="green")
            else:
                self.lbl_saldo.config(foreground="red")
            
            # Habilitar botão de salvar
            self.btn_salvar.config(state="normal")
            self.status_bar.config(text="Dados processados com sucesso!")
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao processar arquivo:\n{str(e)}")
            self.status_bar.config(text="Erro ao processar dados.")
    
    def salvar_pdf(self):
        arquivo_pdf = filedialog.asksaveasfilename(
            title="Salvar Relatório PDF",
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf"), ("Todos os arquivos", "*.*")],
            initialfile=f"relatorio_financeiro_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
        )
        
        if arquivo_pdf:
            try:
                self.status_bar.config(text="Gerando PDF...")
                self.root.update()
                
                pdf = FPDF()
                pdf.add_page()
                
                # Título
                pdf.set_font("Arial", "B", 16)
                pdf.cell(0, 10, "Relatório Financeiro", ln=True, align="C")
                pdf.ln(5)
                
                # Data
                pdf.set_font("Arial", "", 10)
                pdf.cell(0, 10, f"Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}", 
                        ln=True, align="C")
                pdf.ln(10)
                
                # Entradas
                pdf.set_font("Arial", "B", 12)
                pdf.cell(100, 10, "Total Entradas:", border=1)
                pdf.set_font("Arial", "", 12)
                pdf.set_text_color(0, 128, 0)
                pdf.cell(0, 10, f"R$ {self.total_entradas:,.2f}", border=1, ln=True, align="R")
                
                # Saídas
                pdf.set_text_color(0, 0, 0)
                pdf.set_font("Arial", "B", 12)
                pdf.cell(100, 10, "Total Saídas:", border=1)
                pdf.set_font("Arial", "", 12)
                pdf.set_text_color(255, 0, 0)
                pdf.cell(0, 10, f"R$ {self.total_saidas:,.2f}", border=1, ln=True, align="R")
                
                pdf.ln(5)
                
                # Saldo
                pdf.set_text_color(0, 0, 0)
                pdf.set_font("Arial", "B", 14)
                pdf.cell(100, 10, "Saldo Final:", border=1)
                pdf.set_font("Arial", "B", 14)
                if self.saldo >= 0:
                    pdf.set_text_color(0, 128, 0)
                else:
                    pdf.set_text_color(255, 0, 0)
                pdf.cell(0, 10, f"R$ {self.saldo:,.2f}", border=1, ln=True, align="R")
                
                # Salvar PDF
                pdf.output(arquivo_pdf)
                
                self.status_bar.config(text=f"PDF salvo com sucesso!")
                messagebox.showinfo("Sucesso", f"Relatório salvo em:\n{arquivo_pdf}")
                
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao gerar PDF:\n{str(e)}")
                self.status_bar.config(text="Erro ao gerar PDF.")

if __name__ == "__main__":
    root = tk.Tk()
    app = FinancialAnalyzer(root)
    root.mainloop()