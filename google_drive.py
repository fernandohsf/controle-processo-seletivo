from datetime import datetime
from google_connection import autenticar_google_API

def conectar_drive(api):
    api.botao_conectar.config(state="disabled")
    api.adicionar_mensagem("🔃 Conectando ao Google Drive...")
    api.service_drive, api.cliente_gspread = autenticar_google_API()

    if api.service_drive is None:
        api.adicionar_mensagem("❌ Erro ao autenticar no Google Drive.")
        api.botao_conectar.config(state="normal")
    else:
        api.adicionar_mensagem("✅ Conectado ao Google Drive com sucesso!")
        listar_arquivos_drive(api)
        api.botao_enviar_emails.config(state="normal")

def listar_arquivos_drive(api):
    try:
        resultado = api.service_drive.files().list(
            q=f"'{api.pasta_id_drive}' in parents and trashed=false",
            fields="files(id, name, mimeType, modifiedTime)"
        ).execute()

        arquivos = resultado.get('files', [])

        api.arquivos_planilhas = [arq for arq in arquivos if arq.get("mimeType") == "application/vnd.google-apps.spreadsheet"]

        for arq in api.arquivos_planilhas:
            raw_date = arq.get("modifiedTime")
            if raw_date:
                try:
                    arq["modification_date"] = datetime.fromisoformat(raw_date.replace("Z", ""))
                except Exception:
                    arq["modification_date"] = datetime.min
            else:
                arq["modification_date"] = datetime.min

        api.montar_lista_de_arquivos()
    except Exception as e:
        api.adicionar_mensagem(f"⚠️ Erro ao listar arquivos: {e}")
        api.botao_conectar.config(state="normal")