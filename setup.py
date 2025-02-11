from cx_Freeze import setup, Executable
import os

caminhoExe = os.path.dirname(os.path.realpath(__file__))

build_exe_options = {
    "packages": [
        "os", "sys", "re", "gspread", "googleapiclient", "smtplib", "ssl", "tkinter", "PIL", "email", "dotenv", "webbrowser"
    ],
    "include_files": [
        f"{caminhoExe}/imagens",
        f"{caminhoExe}/credenciais_google.json",
        f"{caminhoExe}/.env"
    ],
    "excludes": [],
}

setup(
    name="Gerenciador de processos seletivos internos",
    version="1.0",
    description="GPSI FAPEC.",
    options={"build_exe": build_exe_options},
    executables=[Executable("index.py", base="Win32GUI", target_name="GPSI", icon=f"{caminhoExe}/imagens/Nexus.ico", copyright="© 2025 Fundação FAPEC - Todos os direitos reservados.")],
)