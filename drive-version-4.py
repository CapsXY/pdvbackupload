import pyodbc
import tkinter as tk
from tkinter import messagebox, Listbox, filedialog
from tkinter.ttk import Progressbar
import os
from datetime import datetime
import threading
import time
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
import logging

logging.basicConfig(level=logging.INFO)

def autenticar_google_drive():
    gauth = GoogleAuth()
    try:
        gauth.LoadCredentialsFile("mycreds.txt")
    except Exception as e:
        logging.error(f"Erro ao acessar mycreds.txt: {e}")

    if not gauth.credentials:
        logging.info("Autenticação local do servidor")
        gauth.LocalWebserverAuth()  # Authenticate if they're not there
    elif gauth.access_token_expired:
        logging.info("Token expirado, tentando renovar")
        gauth.Refresh()  # Refresh them if expired
    else:
        logging.info("Autorizando com credenciais salvas")
        gauth.Authorize()  # Initialize the saved creds

    gauth.SaveCredentialsFile("mycreds.txt")  # Save the current credentials
    logging.info("Credenciais salvas com sucesso")

    return GoogleDrive(gauth)


# Exemplo de uso:
try:
    drive = autenticar_google_drive()
    logging.info("Autenticação concluída com sucesso")
except Exception as e:
    logging.error(f"Erro durante a autenticação: {e}")

def listar_bancos_de_dados(listbox):
    server = '.\\PDVNET'
    username = 'sa'
    password = 'inter#system'

    try:
        with pyodbc.connect(
                f'DRIVER={{SQL Server}};'
                f'SERVER={server};'
                f'UID={username};'
                f'PWD={password}'
        ) as conexao:
            cursor = conexao.cursor()
            cursor.execute("SELECT name FROM sys.databases")

            listbox.delete(0, tk.END)
            for row in cursor.fetchall():
                listbox.insert(tk.END, row[0])

            messagebox.showinfo("Sucesso", "Lista de bancos de dados obtida com sucesso!")

    except Exception as e:
        messagebox.showerror("Erro", str(e))


def fazer_backup(database, progress_bar, root):
    server = '.\\PDVNET'
    username = 'sa'
    password = 'inter#system'

    if not database:
        messagebox.showerror("Erro", "O nome do banco de dados não foi inserido.")
        return

    try:
        with pyodbc.connect(
                f'DRIVER={{SQL Server}};'
                f'SERVER={server};'
                f'DATABASE={database};'
                f'UID={username};'
                f'PWD={password}'
        ) as conexao:
            conexao.autocommit = True
            cursor = conexao.cursor()
            cursor.execute("SELECT 1")

            timestamp = datetime.now().strftime("%d-%m-%Y")
            backup_name = f"{database}-{timestamp}.bak"
            caminho_backup = os.path.join("C:\\SQLBackups", backup_name)
            os.makedirs(os.path.dirname(caminho_backup), exist_ok=True)

            comando_backup = f"BACKUP DATABASE [{database}] TO DISK = N'{caminho_backup}' WITH NOFORMAT, NOINIT, NAME = '{database}-Full Database Backup', SKIP, NOREWIND, NOUNLOAD, STATS = 10"
            cursor.execute(comando_backup)

            progress_bar["value"] = 100
            root.update_idletasks()
            time.sleep(10)
            messagebox.showinfo("Sucesso", "Backup realizado com sucesso!")

    except Exception as e:
        messagebox.showerror("Erro", str(e))

    finally:
        progress_bar["value"] = 0
        root.update_idletasks()


def iniciar_backup(database, progress_bar, root):
    progress_bar["value"] = 0
    root.update_idletasks()
    threading.Thread(target=fazer_backup, args=(database, progress_bar, root)).start()


def autenticar_google_drive():
    gauth = GoogleAuth()
    gauth.LoadClientConfigFile("client_secrets.json")
    gauth.LoadCredentialsFile("mycreds.txt")
    if not gauth.credentials:
        gauth.LocalWebserverAuth()
    elif gauth.access_token_expired:
        gauth.Refresh()
    else:
        gauth.Authorize()
    gauth.SaveCredentialsFile("mycreds.txt")
    return GoogleDrive(gauth)

def fazer_upload_google_drive(file_path):
    drive = autenticar_google_drive()
    arquivo = drive.CreateFile({'title': os.path.basename(file_path)})
    arquivo.SetContentFile(file_path)
    arquivo.Upload()
    messagebox.showinfo("Sucesso", f"Arquivo {os.path.basename(file_path)} enviado com sucesso para o Google Drive!")

def escolher_e_fazer_upload():
    file_path = filedialog.askopenfilename()
    if file_path:
        fazer_upload_google_drive(file_path)

def criar_interface():
    root = tk.Tk()
    root.title("Backup do SQL Server")

    tk.Label(root, text="Nome do banco de dados:").grid(row=0)
    nome_banco = tk.Entry(root)
    nome_banco.grid(row=0, column=1)

    progress_bar = Progressbar(root, orient="horizontal", length=200, mode="determinate")
    progress_bar.grid(row=1, columnspan=2, pady=10)

    tk.Button(root, text="Fazer Backup", command=lambda: iniciar_backup(nome_banco.get(), progress_bar, root)).grid(row=2, columnspan=2)
    listbox = Listbox(root)
    listbox.grid(row=3, columnspan=2, pady=10)
    tk.Button(root, text="Listar bancos de dados", command=lambda: listar_bancos_de_dados(listbox)).grid(row=4, columnspan=2)
    tk.Button(root, text="Fazer Upload para o Google Drive", command=escolher_e_fazer_upload).grid(row=5, columnspan=2, pady=10)

    root.mainloop()

if __name__ == "__main__":
    criar_interface()