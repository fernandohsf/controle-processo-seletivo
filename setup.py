from cx_Freeze import setup, Executable
import os

caminhoExe = os.path.dirname(os.path.realpath(__file__))

build_exe_options = {
    "packages": [
        "os", "sys", "re", "google_connection", "google_sheets", "gspread", "googleapiclient",
        "smtplib", "ssl", "tkinter", "PIL", "email", "dotenv", "webbrowser", "concurrent.futures"
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
    executables=[Executable("index.py", base="Win32GUI", target_name="GPSI")],
)