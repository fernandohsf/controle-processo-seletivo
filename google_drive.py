from datetime import datetime
from google_connection import autenticar_google_API

def conectar_drive(app):
    app.botao_conectar.config(state="disabled")
    app.adicionar_mensagem("🔃 Conectando ao Google Drive...")
    app.service_drive, app.cliente_gspread = autenticar_google_API()

    if app.service_drive is None:
        app.adicionar_mensagem("❌ Erro ao autenticar no Google Drive.")
        app.botao_conectar.config(state="normal")
    else:
        app.adicionar_mensagem("✅ Conectado ao Google Drive com sucesso!")
        listar_arquivos_drive(app)
        app.botao_enviar_emails.config(state="normal")

def listar_arquivos_drive(app):
    try:
        resultado = app.service_drive.files().list(
            q=f"'{app.pasta_id_drive}' in parents and trashed=false",
            fields="files(id, name, mimeType, modifiedTime)"
        ).execute()

        arquivos = resultado.get('files', [])

        app.arquivos_planilhas = [arq for arq in arquivos if arq.get("mimeType") == "application/vnd.google-apps.spreadsheet"]

        for arq in app.arquivos_planilhas:
            raw_date = arq.get("modifiedTime")
            if raw_date:
                try:
                    arq["modification_date"] = datetime.fromisoformat(raw_date.replace("Z", ""))
                except Exception:
                    arq["modification_date"] = datetime.min
            else:
                arq["modification_date"] = datetime.min

        app.montar_lista_de_arquivos()
    except Exception as e:
        app.adicionar_mensagem(f"⚠️ Erro ao listar arquivos: {e}")
        app.botao_conectar.config(state="normal")