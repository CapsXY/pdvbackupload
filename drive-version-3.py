import pyodbc
import tkinter as tk
from tkinter import messagebox, Listbox
from tkinter.ttk import Progressbar
import os
from datetime import datetime
import threading
import time


def listar_bancos_de_dados(listbox):
    # Informações do banco de dados inseridas internamente
    server = '.\\PDVNET'  # Substitua por seu servidor
    username = 'sa'  # Substitua por seu usuário
    password = 'inter#system'  # Substitua por sua senha

    try:
        # Conexão com o SQL Server
        with pyodbc.connect(
                f'DRIVER={{SQL Server}};'
                f'SERVER={server};'
                f'UID={username};'
                f'PWD={password}'
        ) as conexao:

            cursor = conexao.cursor()

            # Consulta para obter os nomes dos bancos de dados
            cursor.execute("SELECT name FROM sys.databases")

            listbox.delete(0, tk.END)
            for row in cursor.fetchall():
                listbox.insert(tk.END, row[0])

            messagebox.showinfo("Sucesso", "Lista de bancos de dados obtida com sucesso!")

    except Exception as e:
        messagebox.showerror("Erro", str(e))


def fazer_backup(database, progress_bar, root):
    # Informações do banco de dados inseridas internamente
    server = '.\\PDVNET'  # Substitua por seu servidor
    username = 'sa'  # Substitua por seu usuário
    password = 'inter#system'  # Substitua por sua senha

    if not database:
        messagebox.showerror("Erro",
                             "O nome do banco de dados não foi inserido.")
        return

    try:
        # Conexão com o SQL Server
        with pyodbc.connect(
                f'DRIVER={{SQL Server}};'
                f'SERVER={server};'
                f'DATABASE={database};'
                f'UID={username};'
                f'PWD={password}'
        ) as conexao:

            conexao.autocommit = True  # Desativa a transação automática

            cursor = conexao.cursor()

            # Confirmação da conexão
            cursor.execute("SELECT 1")

            # Gerar nome do backup com datetime
            timestamp = datetime.now().strftime("%d-%m-%Y")
            backup_name = f"{database}-{timestamp}.bak"

            # Caminho do backup
            caminho_backup = os.path.join("C:\\SQLBackups", backup_name)

            # Verifica se o diretório existe, caso contrário cria-o
            os.makedirs(os.path.dirname(caminho_backup), exist_ok=True)

            # Comando de backup
            comando_backup = f"BACKUP DATABASE [{database}] TO DISK = N'{caminho_backup}' WITH NOFORMAT, NOINIT, NAME = '{database}-Full Database Backup', SKIP, NOREWIND, NOUNLOAD, STATS = 10"
            cursor.execute(comando_backup)

            progress_bar["value"] = 100
            root.update_idletasks()
            time.sleep(10)  # Espera por 10 segundos antes de exibir a mensagem de sucesso
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


def criar_interface():
    root = tk.Tk()
    root.title("Backup do SQL Server")

    # Labels e entrada
    tk.Label(root, text="Nome do banco de dados:").grid(row=0)

    nome_banco = tk.Entry(root)
    nome_banco.grid(row=0, column=1)

    # Barra de progresso
    progress_bar = Progressbar(root, orient="horizontal", length=200, mode="determinate")
    progress_bar.grid(row=1, columnspan=2, pady=10)

    # Botão de backup
    tk.Button(root, text="Fazer Backup", command=lambda: iniciar_backup(nome_banco.get(), progress_bar, root)).grid(
        row=2, columnspan=2)

    # Listbox para exibir bancos de dados
    listbox = Listbox(root)
    listbox.grid(row=3, columnspan=2, pady=10)

    # Botão para listar bancos de dados
    tk.Button(root, text="Listar bancos de dados", command=lambda: listar_bancos_de_dados(listbox)).grid(row=4,
                                                                                                         columnspan=2)

    root.mainloop()


if __name__ == "__main__":
    criar_interface()
