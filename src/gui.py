import datetime
import logging
import os
import time
import tkinter as tk
from datetime import datetime
from tkinter import Tk, Canvas, Frame, Label, Button, filedialog, messagebox, Listbox, Entry

import pyodbc
from PIL import Image, ImageTk
from PIL.Image import Resampling
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive


# Função para executar a tela de backup
def run_backup():
    new_frame = Frame(canvas, bg="#E9E9E9", width=500, height=500)
    new_frame.place(x=185, y=0)

    Label(new_frame, text="Nome do banco de dados:", font=("RobotoFlex Bold", 10, "bold"), bg="#E9E9E9",
          fg="#523EAD").place(x=30, y=15)

    entry_db_name = Entry(canvas)
    canvas.create_window(400, 15, anchor="nw", window=entry_db_name, width=150, height=22)

    btn_backup = Button(canvas, text="Fazer Backup", command=lambda: backup(entry_db_name.get(), window))
    canvas.create_window(470, 50, anchor="nw", window=btn_backup)

    btn_bancos = Button(canvas, text="Listar bancos de dados", command=lambda: list_database(listbox))
    canvas.create_window(230, 50, anchor="nw", window=btn_bancos)

    listbox = Listbox(canvas)
    canvas.create_window(220, 90, anchor="nw", window=listbox, width=150, height=400)


# Função para listar os bancos de dados
def list_database(listbox):
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

    except Exception as e:
        messagebox.showerror("Erro", "Não foi foi possível encontrar banco(s) de dados :(")


# Função para realizar backup
def backup(database, root):
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
            caminho_backup = os.path.join("C:\\BASESQL", backup_name)
            os.makedirs(os.path.dirname(caminho_backup), exist_ok=True)

            comando_backup = f"BACKUP DATABASE [{database}] TO DISK = N'{caminho_backup}' WITH NOFORMAT, NOINIT, NAME = '{database}-Full Database Backup', SKIP, NOREWIND, NOUNLOAD, STATS = 10"
            cursor.execute(comando_backup)

            root.update_idletasks()
            time.sleep(10)
            messagebox.showinfo("Sucesso", "Backup realizado com sucesso!")

    except Exception as e:
        messagebox.showerror("Erro", str(e))


# Funções para executar a tela de upload
def run_upload():
    new_frame = Frame(canvas, bg="#E9E9E9", width=500, height=500)
    new_frame.place(x=180, y=0)

    Label(new_frame, text="Selecione um arquivo para realizar o upload", font=("RobotoFlex Bold", 12, "bold"),
          bg="#E9E9E9",
          fg="#523EAD").place(x=40, y=50)
    Button(new_frame, text="Selecione Arquivo", command=select_file).place(x=145, y=90)


# Função para selecionar o arquivo do computador
def select_file():
    file_path = filedialog.askopenfilename()
    if file_path:
        fazer_upload_google_drive(file_path)


# Recurso para verificar e gerar mensagens referente a conexão do Google Drive.
logging.basicConfig(level=logging.INFO)


# Função para autenticar no Google Drive
def autenticar_google_drive():
    gauth = GoogleAuth()
    gauth.LoadClientConfigFile("clientsecrets.json")
    try:
        gauth.LoadCredentialsFile("mycreds.txt")
    except Exception as e:
        logging.error(f"Erro ao acessar mycreds.txt: {e}")

    if not gauth.credentials:
        logging.info("Autenticação local do servidor")
        gauth.LocalWebserverAuth()
    elif gauth.access_token_expired:
        logging.info("Token expirado, tentando renovar")
        gauth.Refresh()
    else:
        logging.info("Autorizando com credenciais salvas")
        gauth.Authorize()

    gauth.SaveCredentialsFile("mycreds.txt")
    logging.info("Credenciais salvas com sucesso")

    return GoogleDrive(gauth)


# Função para fazer o upload no Google Drive
def fazer_upload_google_drive(file_path):
    drive = autenticar_google_drive()
    arquivo = drive.CreateFile({'title': os.path.basename(file_path)})
    arquivo.SetContentFile(file_path)
    arquivo.Upload()

    # Gerar link de compartilhamento público
    arquivo.InsertPermission({
        'type': 'anyone',
        'value': 'anyone',
        'role': 'reader'
    })
    link_publico = arquivo['alternateLink']
    messagebox.showinfo("Sucesso",
                        f"Arquivo {os.path.basename(file_path)} enviado com sucesso para o Google Drive!\nLink público: {link_publico}")


# Função para alterar a cor do texto (clickable_text) ao passar o mouse
def on_enter(event, color="#3B92AC"):
    canvas.itemconfig(event.widget.find_withtag("current"), fill=color)

    # Função para alterar a cor do texto (clickable_text) ao tirar o mouse


# Função para alterar a cor do texto (clickable_text) ao passar o mouse
def on_leave(event, color="#FFFFFF"):
    canvas.itemconfig(event.widget.find_withtag("current"), fill=color)


# Interface principal

window = Tk()
window.geometry("600x500")
window.title("PDVNET - Backup & Upload")
window.configure(bg="#FFFFFF")

canvas = Canvas(
    window,
    bg="#FFFFFF",
    height=500,
    width=600,
    bd=0,
    highlightthickness=0,
    relief="ridge"
)

canvas.place(x=0, y=0)
canvas.create_rectangle(
    0.0,
    0.0,
    600.0,
    500.0,
    fill="#E9E9E9",
    outline="")

canvas.create_rectangle(
    0.0,
    0.0,
    186.0,
    500.0,
    fill="#523EAD",
    outline="")

canvas.create_rectangle(
    0.0,
    0.0,
    186.0,
    74.0,
    fill="#FFFFFF",
    outline="")

canvas.create_text(
    12.0,
    473.0,
    anchor="nw",
    text="ver. 1.0",
    fill="#FFFFFF",
    font=("RobotoFlex Regular", 14 * -1)
)

clickable_text = canvas.create_text(
    12.0,
    145.0,
    anchor="nw",
    text="Sair",
    fill="#FFFFFF",
    font=("RobotoFlex Regular", 20 * -1)
)

canvas.tag_bind(clickable_text, "<Button-1>", lambda e: window.destroy())
canvas.tag_bind(clickable_text, "<Enter>", on_enter)
canvas.tag_bind(clickable_text, "<Leave>", on_leave)

clickable_text = canvas.create_text(
    12.0,
    117.0,
    anchor="nw",
    text="Upload files",
    fill="#FFFFFF",
    font=("RobotoFlex Regular", 20 * -1)
)

canvas.tag_bind(clickable_text, "<Button-1>", lambda e: run_upload())
canvas.tag_bind(clickable_text, "<Enter>", on_enter)
canvas.tag_bind(clickable_text, "<Leave>", on_leave)

clickable_text = canvas.create_text(
    12.0,
    89.0,
    anchor="nw",
    text="Backup",
    fill="#FFFFFF",
    font=("RobotoFlex Regular", 20 * -1)
)

canvas.tag_bind(clickable_text, "<Button-1>", lambda e: run_backup())
canvas.tag_bind(clickable_text, "<Enter>", on_enter)
canvas.tag_bind(clickable_text, "<Leave>", on_leave)

canvas.create_text(
    215.0,
    213.0,
    anchor="nw",
    justify="center",
    text="Seja bem-vindo!\nEscolha uma das opções para iniciar.",
    fill="#523EAD",
    font=("RobotoFlex Bold", 20 * -1, "bold")
)

canvas.create_text(
    25.0,
    45.0,
    anchor="nw",
    text="Backup & Upload",
    fill="#523EAD",
    font=("RobotoFlex Bold", 16 * -1, "bold")
)

canvas.create_rectangle(
    20.0,
    8.0,
    170.0,
    37.0,
    fill="#FFFFFF",
    outline="")

# Icone
icon_path = "logo-short.png"
icon = Image.open(icon_path)
photo = ImageTk.PhotoImage(icon)
window.iconphoto(True, photo)

# Imagem
image_path = "logo.png"
image = Image.open(image_path)
image = image.resize((150, 35), Resampling.LANCZOS)
photo = ImageTk.PhotoImage(image)

canvas.create_image(20, 10, image=photo, anchor="nw")

window.resizable(False, False)
window.mainloop()
