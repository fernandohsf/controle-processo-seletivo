import os
import re
import sys
import smtplib
import threading
from tkinter import messagebox
from dotenv import load_dotenv
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from google_drive import listar_arquivos_drive
from mensagens import mensagens
from utils import arquivos_selecionados, verificar_selecao

def enviar_email(destinatario, assunto, corpo):
    base = os.path.abspath(sys.argv[0])
    caminhoExe = os.path.dirname(base)
    load_dotenv(os.path.join(caminhoExe, ".env"))

    remetente = os.getenv('USUARIONAORESPONDA')
    senha = os.getenv('SENHANAORESPONDA')
    
    mensagem = MIMEMultipart()
    mensagem['From'] = remetente
    mensagem['To'] = destinatario
    mensagem['Subject'] = assunto
    mensagem.attach(MIMEText(corpo, 'html'))

    try:
        with smtplib.SMTP('smtp.gmail.com', 587) as servidor:
            servidor.starttls()
            servidor.login(remetente, senha)
            servidor.sendmail(remetente, destinatario, mensagem.as_string())
        return True
    except Exception as e:
        messagebox.showwarning("Aviso", f"Erro ao enviar email para {destinatario}: {e}")
        return False

def tarefa_envio(app):
    if not verificar_selecao(app):
        return

    ids_arquivos_selecionados = arquivos_selecionados(app)
    
    try:
        app.botao_enviar_emails.config(state="disabled")
        for arquivo_id in ids_arquivos_selecionados:
            try:
                planilha = app.cliente_gspread.open_by_key(arquivo_id)
                sheet = planilha.sheet1
                dados = sheet.get_all_values()
                
                if not dados:
                    messagebox.showwarning("Aviso", "A planilha está vazia.")
                    continue
                
                # MAPEAMENTO DE COLUNAS POR ÍNDICES
                col_map = {
                    26: 25, # Coluna Z (índice 25) Controle de e-mail
                    25: 24, # Coluna Y (índice 24) Motivo de Inaptidão/Eliminação
                    24: 23, # Coluna X (índice 23) Status da inscrição
                    23: 22, # Coluna W (índice 22) Analista RH
                    7: 6,   # Coluna G (índice 6) Nome Completo
                    6: 5,   # Coluna F (índice 5) Vaga
                    5: 4,   # Coluna E (índice 4) Número do Processo Seletivo
                    3: 2,   # Coluna C (índice 2) Endereço de e-mail
                    1: 0,   # Coluna A (índice 0) Número da Inscrição
                }
                
                emails_para_enviar = []
                linhas_para_atualizar = []
                
                # PROCESSAMENTO
                total_emails_enviar = sum(1 for linha in dados[1:] if linha[col_map[26]].strip().lower() == "enviar")

                for i, linha in enumerate(dados[1:], start=2):
                    if linha[col_map[26]].strip().lower() == "enviar":
                        status = linha[col_map[24]].strip().lower()
                        
                        mensagem = mensagens(status, linha, col_map)

                        if not mensagem:
                            continue
                        
                        destinatario = linha[col_map[3]]
                        
                        if not re.match(r"^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$", destinatario):
                            messagebox.showwarning("Aviso", f"Email de {linha[col_map[1]]} - {linha[col_map[7]]} está em formato inválido ({destinatario}).")
                            continue
                        
                        emails_para_enviar.append((destinatario, mensagem, i))
                
                if not emails_para_enviar:
                    messagebox.showwarning("Aviso", "Nenhum email qualificado para envio.")
                    continue
                
                assunto = f"Retorno do processo seletivo {linha[col_map[5]]}"

                quantidade = 0
                for destinatario, mensagem, linha in emails_para_enviar:
                    quantidade +=1
                    app.root.after(0, app.adicionar_mensagem, f"⏩ Enviando e-mail ({quantidade}/{total_emails_enviar}). Destinatário: {destinatario}...")  
                    
                    if enviar_email(destinatario, assunto, mensagem):
                        linhas_para_atualizar.append(linha)
                
                # Atualizar planilha com "Enviado"
                if linhas_para_atualizar:
                    atualizacoes = [
                        {
                            "range": f"Z{linha}",
                            "values": [["Enviado"]]
                        }
                        for linha in linhas_para_atualizar
                    ]
                    sheet.batch_update(atualizacoes)
                
                app.root.after(0, app.adicionar_mensagem, f"✅ Todos os emails da planilha {planilha.title} foram enviados.")
                messagebox.showinfo("Aviso", f"Todos os emails da planilha {planilha.title} foram enviados.")
                listar_arquivos_drive(app)
                
            except Exception as e:
                messagebox.showwarning("Aviso", f"Erro ao processar a planilha {arquivo_id}: {e}\n Se o erro persistir, contate a UNIAE.")
                print(f"Erro ao processar a planilha {arquivo_id}: {e}")
    
    except Exception as e:
        messagebox.showwarning("Aviso", f"Erro inesperado: {e}\n Se o erro persistir, contate a UNIAE.")
    finally:
        app.botao_enviar_emails.config(state="normal")

def enviar_emails(app):
    thread_envio = threading.Thread(target=tarefa_envio, args=(app,))
    thread_envio.start()