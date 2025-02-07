import os
import re
import sys
import smtplib
from tkinter import messagebox
from dotenv import load_dotenv
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

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
    mensagem.attach(MIMEText(corpo, 'plain'))

    try:
        with smtplib.SMTP('smtp.gmail.com', 587) as servidor:
            servidor.starttls()
            servidor.login(remetente, senha)
            texto = mensagem.as_string()
            servidor.sendmail(remetente, destinatario, texto)
        return True
    except Exception as e:
        messagebox.showwarning("Aviso", f"Erro ao enviar email para {destinatario}: {e}")
        return False

def enviar_emails(api):
    ids_arquivos_selecionados = list(api.arquivos_selecionados)
    
    if not ids_arquivos_selecionados:
        messagebox.showwarning("Aviso", "Nenhuma planilha selecionada.")
        return
    
    try:
        api.botao_enviar_emails.config(state="disabled")
        for arquivo_id in ids_arquivos_selecionados:
            try:
                planilha = api.cliente_gspread.open_by_key(arquivo_id)
                sheet = planilha.sheet1
                dados = sheet.get_all_values()
                
                if not dados:
                    messagebox.showwarning("Aviso", "A planilha está vazia.")
                    continue
                
                # Mapear colunas necessárias pelos índices
                col_map = {
                    20: 19,  # Coluna T (índice 19)
                    19: 18,  # Coluna S (índice 18)
                    18: 17,  # Coluna R (índice 17)
                    5: 4,    # Coluna E (índice 4)
                    3: 2,    # Coluna C (índice 2)
                    1: 0,    # Coluna A (índice 0)
                }
                
                emails_para_enviar = []
                linhas_para_atualizar = []
                
                # Processar as linhas
                for i, linha in enumerate(dados[1:], start=2):
                    if linha[col_map[20]].strip().lower() == "enviar":
                        status = linha[col_map[18]].strip().lower()
                        
                        mensagem = mensagens(status, linha, col_map)
                        
                        destinatario = linha[col_map[3]]
                        
                        if not re.match(r"[^@]+@[^@]+\.[^@]+", destinatario):
                            messagebox.showwarning("Aviso", f"Email inválido: {destinatario}")
                            continue
                        
                        emails_para_enviar.append((destinatario, mensagem, i))
                
                if not emails_para_enviar:
                    messagebox.showwarning("Aviso", "Nenhum email qualificado para envio.")
                    continue
                
                # Extrair números do nome da planilha para o assunto
                numeros_no_nome = re.findall(r'\d+', planilha.title)
                if numeros_no_nome:
                    ultimo_numero = numeros_no_nome[-1]
                    if len(ultimo_numero) >= 4:
                        ultimo_numero = ultimo_numero[:-4] + '/' + ultimo_numero[-4:]
                    numeros_no_nome[-1] = ultimo_numero
                    assunto = f"Retorno do processo seletivo {''.join(numeros_no_nome)}"
                else:
                    assunto = "Retorno do processo seletivo"

                for email, mensagem, linha in emails_para_enviar:
                    api.adicionar_mensagem(f"⏩ Enviando e-mail para {email}...")
                    if enviar_email(email, assunto, mensagem):
                        linhas_para_atualizar.append(linha)
                
                # Atualizar planilha com "Enviado"
                if linhas_para_atualizar:
                    atualizacoes = [
                        {
                            "range": f"T{linha}",
                            "values": [["Enviado"]]
                        }
                        for linha in linhas_para_atualizar
                    ]
                    sheet.batch_update(atualizacoes)
                
                api.adicionar_mensagem(f"✅ Todos os emails da planilha {planilha.title} foram enviados.")

                messagebox.showwarning("Aviso", f"Todos os emails da planilha {planilha.title} foram enviados.")
                api.listar_arquivos_drive()
                
            except Exception as e:
                messagebox.showwarning("Aviso", f"Erro ao processar a planilha {arquivo_id}: {e}\n Se o erro persistir, contate a UNIAE.")
                print(f"Erro ao processar a planilha {arquivo_id}: {e}")
    
    except Exception as e:
        messagebox.showwarning("Aviso", f"Erro inesperado: {e}\n Se o erro persistir, contate a UNIAE.")
    
    api.botao_enviar_emails.config(state="normal")
