import os
import re
import sys
import smtplib
from tkinter import messagebox
from dotenv import load_dotenv
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from google_drive import listar_arquivos_drive
from mensagens import mensagens

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

def enviar_emails(app):
    ids_arquivos_selecionados = list(app.arquivos_selecionados)
    
    if not ids_arquivos_selecionados:
        messagebox.showwarning("Aviso", "Nenhuma planilha selecionada.")
        return
    
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
                
                # Mapear colunas necessárias pelos índices
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
                
                # Processar as linhas
                for i, linha in enumerate(dados[1:], start=2):
                    if linha[col_map[26]].strip().lower() == "enviar":
                        status = linha[col_map[24]].strip().lower()
                        
                        mensagem = mensagens(status, linha, col_map)

                        if not mensagem:
                            continue
                        
                        destinatario = linha[col_map[3]]
                        
                        if not re.match(r"[^@]+@[^@]+\.[^@]+", destinatario):
                            messagebox.showwarning("Aviso", f"Email inválido: {destinatario}")
                            continue
                        
                        emails_para_enviar.append((destinatario, mensagem, i))
                
                if not emails_para_enviar:
                    messagebox.showwarning("Aviso", "Nenhum email qualificado para envio.")
                    continue
                
                assunto = f"Retorno do processo seletivo {linha[col_map[5]]}"

                for email, mensagem, linha in emails_para_enviar:
                    app.adicionar_mensagem(f"⏩ Enviando e-mail para {email}...")
                    if enviar_email(email, assunto, mensagem):
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
                    #sheet.batch_update(atualizacoes)
                
                app.adicionar_mensagem(f"✅ Todos os emails da planilha {planilha.title} foram enviados.")

                messagebox.showwarning("Aviso", f"Todos os emails da planilha {planilha.title} foram enviados.")
                listar_arquivos_drive(app)
                
            except Exception as e:
                messagebox.showwarning("Aviso", f"Erro ao processar a planilha {arquivo_id}: {e}\n Se o erro persistir, contate a UNIAE.")
                print(f"Erro ao processar a planilha {arquivo_id}: {e}")
    
    except Exception as e:
        messagebox.showwarning("Aviso", f"Erro inesperado: {e}\n Se o erro persistir, contate a UNIAE.")
    
    app.botao_enviar_emails.config(state="normal")
