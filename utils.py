from tkinter import messagebox

def verificar_selecao(app):
    if not arquivos_selecionados(app):
        messagebox.showwarning("Aviso", "Nenhuma planilha selecionada.")
        return
    return True

def arquivos_selecionados(app):
        selecionados = [arquivo_id for var, arquivo_id in app.selecoes if var.get()]
        return selecionados