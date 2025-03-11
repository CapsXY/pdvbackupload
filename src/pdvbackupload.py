import datetime
import logging
import os
import sys
import time
import zipfile
from datetime import datetime, timedelta
from tkinter import Canvas, Frame, Label, Button, filedialog, messagebox, Listbox, Entry, END, OptionMenu, StringVar, \
    Tk, ttk

import pyodbc
from PIL import Image, ImageTk
from PIL.Image import Resampling
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
from tkcalendar import DateEntry


# Função para o diretório dos arquivos estarem de acordo com o pyinstaller
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


# Função para executar a tela de backup
def run_backup():
    new_frame = Frame(canvas, bg="#E9E9E9", width=500, height=500)
    new_frame.place(x=185, y=0)

    Label(new_frame, text="Informe o nome do banco de dados para o backup", font=("RobotoFlex Bold", 11, "bold"),
          bg="#E9E9E9",
          fg="#523EAD").place(x=30, y=30)

    Label(new_frame, text="Banco de dados:", bg="#E9E9E9").place(x=25, y=70)

    entry_db_name = Entry(canvas)
    canvas.create_window(325, 68, anchor="nw", window=entry_db_name, width=250, height=22)

    btn_backup = Button(canvas, text="Iniciar backup", command=lambda: backup(entry_db_name.get(), window))
    canvas.create_window(490, 100, anchor="nw", window=btn_backup)

    btn_bancos = Button(canvas, text="Listar bancos de dados", command=lambda: list_database(listbox))
    canvas.create_window(225, 120, anchor="nw", window=btn_bancos)

    listbox = Listbox(canvas)
    canvas.create_window(215, 155, anchor="nw", window=listbox, width=150, height=320)


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

            listbox.delete(0, END)
            for row in cursor.fetchall():
                listbox.insert(END, row[0])

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
          bg="#E9E9E9", fg="#523EAD").place(x=40, y=50)

    Button(new_frame, text="Selecione Arquivo", command=lambda: select_file(new_frame, progress_bar)).place(x=145, y=85)

    # Barra de progresso
    progress_bar = ttk.Progressbar(new_frame, orient="horizontal", length=360, mode="determinate")
    progress_bar.place(x=30, y=125)

    Label(new_frame, text="Dica: Faça o upload do arquivo com apenas o número do chamado.", font=("RobotoFlex", 10,),
          bg="#E9E9E9", fg="#523EAD").place(x=13, y=155)
    Label(new_frame, text="Ex.: 989001.zip", font=("RobotoFlex", 10,),
          bg="#E9E9E9", fg="#523EAD").place(x=13, y=175)


# Função para selecionar o arquivo do computador
def select_file(new_frame, progress_bar):
    file_path = filedialog.askopenfilename()
    if file_path:
        upload_google_drive(file_path, progress_bar, new_frame)


# Recurso para verificar e gerar mensagens referente a conexão do Google Drive.
logging.basicConfig(level=logging.INFO)


# Função para autenticar no Google Drive
def authenticate_google_drive():
    gauth = GoogleAuth()
    client_config_path = resource_path("clientsecrets.json")
    gauth.LoadClientConfigFile(client_config_path)
    try:
        creds_path = resource_path("mycreds.txt")
        gauth.LoadCredentialsFile(creds_path)
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

    gauth.SaveCredentialsFile(creds_path)
    logging.info("Credenciais salvas com sucesso")

    return GoogleDrive(gauth)

# Função para realizar upload no Google Drive
def upload_google_drive(file_path, progress_bar, window):
    try:
        # Inicializa o progresso
        progress_bar["value"] = 0
        window.update_idletasks()

        # Simulação de autenticação e upload
        drive = authenticate_google_drive()
        arquivo = drive.CreateFile({'title': os.path.basename(file_path)})

        arquivo.SetContentFile(file_path)

        # Atualiza a barra de progresso dinamicamente
        for i in range(1, 101):  # Simula um progresso de 1% a cada iteração
            time.sleep(0.05)  # Simula um tempo de delay para demonstrar progresso
            progress_bar["value"] = i
            window.update_idletasks()

        arquivo.Upload()

        # Gerar link de compartilhamento público
        arquivo.InsertPermission({
            'type': 'anyone',
            'value': 'anyone',
            'role': 'reader'
        })
        link_publico = arquivo['alternateLink']
        messagebox.showinfo("Sucesso",
                            f"Arquivo {os.path.basename(file_path)} enviado com sucesso!\nLink: {link_publico}")

    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao fazer upload: {str(e)}")
    finally:
        progress_bar["value"] = 0  # Reseta a barra de progresso ao final


# Módulos
module = {
    "PDVGESTOR": "1",
    "PDVFINAN": "2",
    "PDVANALISE": "3",
    "PDVATACADO": "4",
    "PDVSYNC": "5",
    "ETIQUETA": "6",
    "LINK PCCON - RETAGUARDA": "7",
    "SIF": "8",
    "ESTOQUE": "9",
    "PDVECF/PREVENDA": "10",
    "PDVUP": "11",
    "PDVADM/RETAGUARDA": "12",
    "PDVOS": "13",
    "PDVSYNCPLUS": "20",
    "TODOS": "TODOS"
}


# Função para executar tela de auditoria
def run_audit():
    new_frame = Frame(canvas, bg="#E9E9E9", width=500, height=500)
    new_frame.place(x=185, y=0)

    Label(new_frame, text="Selecione um arquivo para realizar o upload", font=("RobotoFlex Bold", 12, "bold"),
          bg="#E9E9E9",
          fg="#523EAD").place(x=40, y=30)

    def select_origin():
        origin_folder = filedialog.askdirectory()
        entry_origin_folder.delete(0, END)
        entry_origin_folder.insert(0, origin_folder)

    def select_destiny():
        destiny_folder = filedialog.askdirectory()
        entry_destiny_folder.delete(0, END)
        entry_destiny_folder.insert(0, destiny_folder)

    def run_backup_audit():
        initial_date = entry_initial_date.get()
        final_date = entry_final_date.get()
        origin_folder = entry_origin_folder.get()
        destiny_folder = entry_destiny_folder.get()

        if not os.path.exists(destiny_folder):
            os.makedirs(destiny_folder)

        # Obtém o valor selecionado do módulo
        module_name = var_module_name.get()

        # Valida se o módulo foi selecionado corretamente
        if not module_name or module_name not in module:
            # Exibe um popup com mensagem de erro
            messagebox.showerror("Erro",
                                 "Selecione um módulo.")
            return

        option = module[module_name]
        search_files(initial_date, final_date, origin_folder, destiny_folder, option)

    Label(new_frame, text="Período:", bg="#E9E9E9").place(x=20, y=80)
    entry_initial_date = DateEntry(new_frame, date_pattern='dd-mm-yyyy')
    entry_initial_date.place(x=80, y=80)
    Label(new_frame, text="até", bg="#E9E9E9").place(x=175, y=80)
    entry_final_date = DateEntry(new_frame, date_pattern='dd-mm-yyyy')
    entry_final_date.place(x=198, y=80)

    Label(new_frame, text="Origem:", bg="#E9E9E9").place(x=20, y=120)
    entry_origin_folder = Entry(new_frame)
    entry_origin_folder.insert(0, 'C:\\PDV\\AUDIT')
    entry_origin_folder.place(x=80, y=120)
    Button(new_frame, text="Selecionar", bg="#E9E9E9", command=select_origin).place(x=140, y=150)

    Label(new_frame, text="Destino:", bg="#E9E9E9").place(x=210, y=120)
    entry_destiny_folder = Entry(new_frame)
    entry_destiny_folder.insert(0, 'C:\\BASESQL')
    entry_destiny_folder.place(x=270, y=120)
    Button(new_frame, text="Selecionar", bg="#E9E9E9", command=select_destiny).place(x=330, y=150)

    Label(new_frame, text="Módulo:", bg='#E9E9E9').place(x=20, y=205)
    var_module_name = StringVar(new_frame)
    var_module_name.set("Selecione")  # Valor padrão inicial

    # Menu suspenso com as opções de módulo
    OptionMenu(new_frame, var_module_name, *module.keys()).place(x=80, y=200)

    Button(new_frame, text="Iniciar backup", command=run_backup_audit).place(x=310, y=200)


# Função para localizar os arquivos de auditoria
def search_files(initial_date, final_date, origin_folder, destiny_folder, option):
    try:
        initial_date = datetime.strptime(initial_date, '%d-%m-%Y')
        final_date = datetime.strptime(final_date, '%d-%m-%Y') + timedelta(days=1) - timedelta(seconds=1)

        time_now = datetime.now()
        file_zipname = os.path.join(destiny_folder, f'audit-{time_now.strftime("%d-%m-%Y_%H%M%S")}.zip')

        if not os.path.exists(origin_folder):
            raise FileNotFoundError(f"A pasta de origem '{origin_folder}' não foi encontrada.")

        files_found = False
        for root, dirs, files in os.walk(origin_folder):
            for file in files:
                fully_directory = os.path.join(root, file)
                modification_datetime = datetime.fromtimestamp(os.path.getmtime(fully_directory))

                if initial_date <= modification_datetime <= final_date:
                    if option == "TODOS" or file.startswith(option):
                        files_found = True
                        with zipfile.ZipFile(file_zipname, 'a') as zipf:
                            zipf.write(fully_directory, os.path.relpath(fully_directory, origin_folder))

        if files_found:
            messagebox.showinfo("Sucesso", "Arquivos copiados e compactados com sucesso!")
        else:
            messagebox.showinfo("Informação", "Nenhum arquivo encontrado com os critérios selecionados.")

    except FileNotFoundError as e:
        messagebox.showerror("Erro", str(e))
    except Exception as e:
        messagebox.showerror("Erro", str(e))


# Função para alterar a cor do texto (clickable_text) ao passar o mouse
def on_enter(event, color="#3B92AC"):
    canvas.itemconfig(event.widget.find_withtag("current"), fill=color)


# Função para alterar a cor do texto (clickable_text) ao tirar o mouse
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

canvas.create_rectangle(
    20.0,
    8.0,
    170.0,
    37.0,
    fill="#FFFFFF",
    outline="")

clickable_text = canvas.create_text(
    12.0,
    85.0,
    anchor="nw",
    text="Backup",
    fill="#FFFFFF",
    font=("RobotoFlex Regular", 20 * -1)
)

canvas.tag_bind(clickable_text, "<Button-1>", lambda e: run_backup())
canvas.tag_bind(clickable_text, "<Enter>", on_enter)
canvas.tag_bind(clickable_text, "<Leave>", on_leave)

clickable_text = canvas.create_text(
    12.0,
    115.0,
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
    145.0,
    anchor="nw",
    text="Auditoria",
    fill="#FFFFFF",
    font=("RobotoFlex Regular", 20 * -1)
)

canvas.tag_bind(clickable_text, "<Button-1>", lambda e: run_audit())
canvas.tag_bind(clickable_text, "<Enter>", on_enter)
canvas.tag_bind(clickable_text, "<Leave>", on_leave)

clickable_text = canvas.create_text(
    12.0,
    175.0,
    anchor="nw",
    text="Sair",
    fill="#FFFFFF",
    font=("RobotoFlex Regular", 20 * -1)
)

canvas.tag_bind(clickable_text, "<Button-1>", lambda e: window.destroy())
canvas.tag_bind(clickable_text, "<Enter>", on_enter)
canvas.tag_bind(clickable_text, "<Leave>", on_leave)

canvas.create_text(
    25.0,
    45.0,
    anchor="nw",
    text="Backup & Upload",
    fill="#523EAD",
    font=("RobotoFlex Bold", 16 * -1, "bold")
)

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
    12.0,
    473.0,
    anchor="nw",
    text="ver. 1.1",
    fill="#FFFFFF",
    font=("RobotoFlex Regular", 14 * -1)
)

# Icone
icon_path = resource_path("logoshort.png")
icon = Image.open(icon_path)
photo = ImageTk.PhotoImage(icon)
window.iconphoto(True, photo)

# Imagem
image_path = resource_path("logo.png")
image = Image.open(image_path)
image = image.resize((150, 35), Resampling.LANCZOS)
photo = ImageTk.PhotoImage(image)

canvas.create_image(20, 10, image=photo, anchor="nw")

window.resizable(False, False)
window.mainloop()
