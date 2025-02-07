import os
import sys
import time
import gspread
from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials

def autenticar_google_API():
    caminhoExe = os.path.dirname(os.path.abspath(sys.argv[0]))
    caminhoCredenciais = os.path.join(caminhoExe, 'credenciais_google.json')
    scope = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'
    ]
    try:
        creds = Credentials.from_service_account_file(caminhoCredenciais, scopes=scope)
        service_drive = build('drive', 'v3', credentials=creds)
        cliente_gspread = gspread.authorize(creds)
        time.sleep(1)
        return service_drive, cliente_gspread
    except:
        return None