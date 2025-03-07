import tkinter as tk
from tkinter import ttk, messagebox

def verificar_selecao(app):
    if not arquivos_selecionados(app):
        messagebox.showwarning("Aviso", "Nenhuma planilha selecionada.")
        return
    return True

def arquivos_selecionados(app):
        selecionados = [arquivo_id for var, arquivo_id in app.selecoes if var.get()]
        return selecionados

def criar_barra_progresso(app):
    app.limpar_canvas()
    
    app.label_barra_progresso = tk.Label(app.canvas, text="Trabalhando...", font=(app.fonte, app.tamanho_fonte_cabecalho, "bold"))
    app.label_barra_progresso.pack(pady=10)

    app.progress = ttk.Progressbar(app.canvas, length=250, mode="determinate")
    app.progress.pack(pady=5)

    app.progress["value"] = 0
    app.canvas.update_idletasks()