import os, sys
import tkinter as tk
from tkinter import ttk, BooleanVar
from datetime import datetime
from PIL import Image, ImageTk
from google_drive import conectar_drive
from emails import enviar_emails
from google_sheets import abrir_planilha, ajustar_planilha, baixar_planilha

class App:
    def __init__(self, root):
        # CORES, FONTES E ATRIBUTOS
        self.azul_escuro = "#203464"
        self.azul_claro = "#04acec"
        self.branco = "white"
        self.cinza_claro = "#CCCCCC"
        self.fonte = "Arial"
        self.tamanho_fonte_cabecalho = 13
        self.tamanho_fonte_corpo = 11
        self.cursor = "hand2"

        self.pasta_id_drive = "1pTddHNebIu5Z77Y24xqpe1zug-GLTz8c" # PRODUÇÃO
        #self.pasta_id_drive = "1-z691tcQJdDU8wIwxQpYL151Ep3uEHoU"  # HOMOLOGAÇÃO

        self.arquivos_planilhas =[]
        self.arquivos = []

        base = os.path.abspath(sys.argv[0])
        self.caminhoExe = os.path.dirname(base)
        self.caminhoImagens = os.path.join(self.caminhoExe, "imagens")

        self.root = root
        self.root.title("Controle de Arquivos do Drive")
        root.state('zoomed')
        self.root.configure(bg=self.azul_escuro)
        self.root.iconbitmap(os.path.join(self.caminhoImagens, "Nexus.ico"))

        # LOGO
        caminho_logo = os.path.join(self.caminhoImagens, "Fapec-logo.png")
        imagem_logo = Image.open(caminho_logo).resize((162, 145))
        self.imagem_logo = ImageTk.PhotoImage(imagem_logo)
        tk.Label(root, image=self.imagem_logo, bg=self.azul_escuro).pack(pady=5)

        tk.Label(root, text="Processos Seletivos Internos", font=("Arial", 16, "bold"),
                 bg=self.azul_escuro, fg=self.branco).pack(pady=5)

        # FRAME DE BOTÕES
        self.frame_botoes = tk.Frame(root, bg=self.azul_escuro)
        self.frame_botoes.pack(fill=tk.X, padx=10, pady=10)

        # BOTÃO CONECTAR AO DRIVE
        botao_conectar_normal = Image.open(f'{self.caminhoImagens}/conectar_drive_normal.png').resize((243, 62))
        botao_conectar_normal = ImageTk.PhotoImage(botao_conectar_normal)
        botao_conectar_ativo = Image.open(f'{self.caminhoImagens}/conectar_drive_ativo.png').resize((243, 62))
        botao_conectar_ativo = ImageTk.PhotoImage(botao_conectar_ativo)
        
        self.botao_conectar = tk.Button(self.frame_botoes, image=botao_conectar_normal,
                                        bg=self.azul_escuro, activebackground=self.azul_escuro,
                                        relief="flat", bd=0,
                                        command=lambda: conectar_drive(self))
        self.botao_conectar.image = botao_conectar_normal
        self.botao_conectar.bind("<Enter>", lambda e: self.botao_conectar.config(image=botao_conectar_ativo))
        self.botao_conectar.bind("<Leave>", lambda e: self.botao_conectar.config(image=botao_conectar_normal))
        self.botao_conectar.pack(side=tk.LEFT, padx=5)

        # BOTÃO PESQUISAR
        self.barra_pesquisa = Image.open(f'{self.caminhoImagens}/barra_pesquisa.png').resize((262, 62))
        self.barra_pesquisa = ImageTk.PhotoImage(self.barra_pesquisa)
        self.label_barra_pesquisa = tk.Label(self.frame_botoes, image=self.barra_pesquisa, bg=self.azul_escuro)
        self.label_barra_pesquisa.pack(side=tk.LEFT, expand=True)

        self.texto_pesquisa = tk.Entry(self.frame_botoes, font=(self.fonte, self.tamanho_fonte_corpo), bd=0)
        self.texto_pesquisa.bind("<KeyRelease>", self.filtrar_itens)

        root.bind("<Configure>", lambda e: self.centralizar_pesquisa())
        root.after(100, self.centralizar_pesquisa)

        # BOTÃO ENVIAR EMAILS
        botao_enviar_emails_normal = Image.open(f'{self.caminhoImagens}/enviar_emails_normal.png').resize((201, 62))
        botao_enviar_emails_normal = ImageTk.PhotoImage(botao_enviar_emails_normal)
        botao_enviar_emails_ativo = Image.open(f'{self.caminhoImagens}/enviar_emails_ativo.png').resize((201, 62))
        botao_enviar_emails_ativo = ImageTk.PhotoImage(botao_enviar_emails_ativo)
        
        self.botao_enviar_emails = tk.Button(self.frame_botoes, image=botao_enviar_emails_normal,
                                            bg=self.azul_escuro, activebackground=self.azul_escuro,
                                            relief="flat", bd=0,
                                            command=lambda: enviar_emails(self),
                                            state="disabled")
        self.botao_enviar_emails.image = botao_enviar_emails_normal
        self.botao_enviar_emails.bind("<Enter>", lambda e: self.botao_enviar_emails.config(image=botao_enviar_emails_ativo))
        self.botao_enviar_emails.bind("<Leave>", lambda e: self.botao_enviar_emails.config(image=botao_enviar_emails_normal))
        self.botao_enviar_emails.pack(side=tk.RIGHT, padx=(41, 10))

        # FRAME LISTAGEM
        frame_listagem = tk.Frame(root, bg=self.azul_escuro)
        frame_listagem.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # FRAME CABEÇALHO (COLOCADO EM ORDEM INVERTIDA POR ESTAR USANDO SIDE=RIGHT NOS DOIS ÚLTIMOS)
        self.frame_cabecalho = tk.Frame(frame_listagem, bg=self.azul_claro)
        self.frame_cabecalho.pack(fill=tk.X)

        tk.Label(self.frame_cabecalho, text="Selecionar", font=(self.fonte, self.tamanho_fonte_cabecalho, "bold"),
                bg=self.azul_claro, fg=self.branco, width=9, anchor="w").pack(side=tk.LEFT)

        self.label_nome = tk.Label(self.frame_cabecalho, text="Nome do Arquivo", font=(self.fonte, self.tamanho_fonte_cabecalho, "bold"),
                                    bg=self.azul_claro, fg=self.branco, anchor="w", cursor=self.cursor)
        self.label_nome.pack(side=tk.LEFT, fill=tk.X)
        self.label_nome.bind("<Button-1>", lambda e: self.sort_by("nome"))

        tk.Label(self.frame_cabecalho, text="GPSI", font=(self.fonte, self.tamanho_fonte_cabecalho, "bold"),
                bg=self.azul_claro, fg=self.branco, width=6, anchor="w").pack(side=tk.RIGHT, padx=(0, 5))

        tk.Label(self.frame_cabecalho, text="Download", font=(self.fonte, self.tamanho_fonte_cabecalho, "bold"),
                bg=self.azul_claro, fg=self.branco, width=8, anchor="w").pack(side=tk.RIGHT, padx=(0,15))

        tk.Label(self.frame_cabecalho, text="Abrir", font=(self.fonte, self.tamanho_fonte_cabecalho, "bold"),
                bg=self.azul_claro, fg=self.branco, width=6, anchor="w").pack(side=tk.RIGHT)
        
        self.label_data = tk.Label(self.frame_cabecalho, text="Data de modificação", font=(self.fonte, self.tamanho_fonte_cabecalho, "bold"),
                                    bg=self.azul_claro, fg=self.branco, width=18, anchor="w", cursor=self.cursor)
        self.label_data.pack(side=tk.RIGHT)
        self.label_data.bind("<Button-1>", lambda e: self.sort_by("data"))

        # CANVAS E SCROLLBAR PARA FRAME LISTAGEM
        self.canvas = tk.Canvas(frame_listagem, bg=self.branco)
        self.scroll_y = ttk.Scrollbar(frame_listagem, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.scroll_y.set)
        self.scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.service_drive = None
        self.arquivos_selecionados = []

        # ATIIBUTOS PARA ORDENAÇÃO
        self.sort_column = None
        self.sort_descending = False

    def montar_lista_de_arquivos(self, arquivos=None):
        self.limpar_canvas()

        if arquivos is None:
            arquivos = self.arquivos_planilhas

        if not arquivos:
            self.adicionar_mensagem("📂 Nenhuma planilha encontrada.",)
            self.botao_conectar.config(state="normal")
            return

        self.arquivos = arquivos
        self.selecoes = []

        # FRAME DE LISTA DE ARQUIVOS
        frame_lista = tk.Frame(self.canvas, bg=self.branco)
        frame_lista_id = self.canvas.create_window((0, 0), window=frame_lista, anchor="nw", width=self.canvas.winfo_width())

        # CALLBACK PARA RESPONSIVIDADE DO FRAME_LISTA
        def on_canvas_configure(event):
            self.canvas.itemconfig(frame_lista_id, width=event.width)

        self.canvas.bind("<Configure>", on_canvas_configure)

        for idx, arquivo in enumerate(self.arquivos):
            nome = arquivo.get("name", "Desconhecido")
            arquivo_id = arquivo.get("id", "ID Desconhecida")
            data_dt = arquivo.get("modification_date", datetime.min)
            data_modificacao = data_dt.strftime("%d/%m/%Y %H:%M") if data_dt != datetime.min else "Data desconhecida"

            # COR DAS LINHAS
            cor_fundo = self.cinza_claro if idx % 2 == 0 else self.branco

            # FRAME DE ITENS DA LISTA
            frame_item = tk.Frame(frame_lista, bg=cor_fundo)
            frame_item.pack(fill=tk.X)

            # FUNÇÃO HOVER PARA LINHA
            def on_enter(event, widget=frame_item):
                widget.config(bg=self.azul_claro)
                for filho in widget.winfo_children():
                    filho.config(bg=self.azul_claro)
                    if isinstance(filho, (tk.Label)):
                        filho.config(fg=self.branco)

            def on_leave(event, widget=frame_item, cor_padrao=cor_fundo):
                widget.config(bg=cor_padrao)
                for filho in widget.winfo_children():
                    filho.config(bg=cor_padrao)
                    if isinstance(filho, (tk.Label)):
                        filho.config(fg="#000000")

            frame_item.bind("<Enter>", on_enter)
            frame_item.bind("<Leave>", on_leave)

            var = BooleanVar()
            self.selecoes.append((var, arquivo_id))

            check = tk.Checkbutton(frame_item, variable=var,
                           bg=cor_fundo, fg=self.azul_escuro,
                           activebackground=self.azul_claro, selectcolor=self.branco,
                           width=10, cursor=self.cursor)
            check.pack(side=tk.LEFT)

            label_nome_arquivo = tk.Label(frame_item, text=nome,
                                        font=(self.fonte, self.tamanho_fonte_corpo),
                                        bg=cor_fundo, fg=self.azul_escuro, anchor="w")
            label_nome_arquivo.pack(side=tk.LEFT, fill=tk.X)

            # BOTÃO DOWNLOAD E ABRIR PANILHA (COLOCADO EM ORDEM INVERTIDA POR ESTAR USANDO SIDE=RIGHT)
            botao_gpsi_img = Image.open(f'{self.caminhoImagens}/botao_gpsi.png').resize((48, 27))
            botao_gpsi_img = ImageTk.PhotoImage(botao_gpsi_img)
            botao_gpsi = tk.Button(frame_item, image=botao_gpsi_img, width=58,
                                    bg=cor_fundo, bd=0, activebackground=self.azul_claro,
                                    command=lambda: ajustar_planilha(self), cursor=self.cursor)
            botao_gpsi.image = botao_gpsi_img
            botao_gpsi.pack(side=tk.RIGHT)

            botao_download_img = Image.open(f'{self.caminhoImagens}/download.png').resize((26, 26))
            botao_download_img = ImageTk.PhotoImage(botao_download_img)
            botao_download = tk.Button(frame_item, image=botao_download_img, width=100,
                                    bg=cor_fundo, bd=0, activebackground=self.azul_claro,
                                    command=lambda: baixar_planilha(self), cursor=self.cursor)
            botao_download.image = botao_download_img
            botao_download.pack(side=tk.RIGHT)

            botao_abrir_img = Image.open(f'{self.caminhoImagens}/abrir.png').resize((39, 26))
            botao_abrir_img = ImageTk.PhotoImage(botao_abrir_img)
            botao_abrir = tk.Button(frame_item, image=botao_abrir_img, width=72,
                                    bg=cor_fundo, bd=0, activebackground=self.azul_claro,
                                    command=lambda: abrir_planilha(self), cursor=self.cursor)
            botao_abrir.image = botao_abrir_img
            botao_abrir.pack(side=tk.RIGHT)

            label_data_modificacao = tk.Label(frame_item, text=data_modificacao,
                    font=(self.fonte, self.tamanho_fonte_corpo),
                    bg=cor_fundo, fg=self.azul_escuro, width=18, anchor="w")
            label_data_modificacao.pack(side=tk.RIGHT)
            
            for widget in [check, label_nome_arquivo, botao_gpsi, botao_download, botao_abrir, label_data_modificacao]:
                widget.bind("<Enter>", on_enter)
                widget.bind("<Leave>", on_leave)

        frame_lista.bind("<Configure>", self.atualizar_scroll)
        frame_lista.bind_all("<MouseWheel>", self.rolar_com_bolinha)

    def sort_by(self, coluna):
        if self.arquivos:
            if self.sort_column == coluna:
                self.sort_descending = not self.sort_descending
            else:
                self.sort_column = coluna
                self.sort_descending = False

            if coluna == "nome":
                self.arquivos.sort(key=lambda arq: arq.get("name", "").lower(),
                                            reverse=self.sort_descending)
            elif coluna == "data":
                self.arquivos.sort(key=lambda arq: arq.get("modification_date", datetime.min),
                                            reverse=self.sort_descending)
                
            self.label_nome.config(text="Nome do Arquivo")
            self.label_data.config(text="Data de modificação")
            seta = "▼" if not self.sort_descending else "▲"
            if coluna == "nome":
                self.label_nome.config(text=f"Nome do Arquivo {seta}")
            elif coluna == "data":
                self.label_data.config(text=f"Data de modificação {seta}")

            self.montar_lista_de_arquivos(arquivos=self.arquivos)
        
    def filtrar_itens(self, event=None):
        if self.arquivos_planilhas:
            termo = self.texto_pesquisa.get().lower()

            arquivos_filtrados = [arq for arq in self.arquivos_planilhas if termo in arq["name"].lower()]

            self.montar_lista_de_arquivos(arquivos=arquivos_filtrados)

    def limpar_canvas(self):
        for widget in self.canvas.winfo_children():
            widget.destroy()
    
    def adicionar_mensagem(self, mensagem):
        self.limpar_canvas()

        label = tk.Label(self.canvas, text=mensagem,
                         font=(self.fonte, self.tamanho_fonte_corpo),
                         bg=self.branco, fg=self.azul_escuro)
        label.pack(pady=2)
        self.canvas.update_idletasks()

    def centralizar_pesquisa(self):
        largura_frame = self.frame_botoes.winfo_width()
        largura_entry = 180
        pos_x = (largura_frame - largura_entry - 35) // 2  # POSICIONAMENTO CENTRAL

        self.texto_pesquisa.place(x=pos_x, y=13, width=largura_entry, height=30)

    def atualizar_scroll(self, event=None):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def rolar_com_bolinha(self, event):
        self.canvas.yview_scroll(-1 * (event.delta // 120), "units")

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()