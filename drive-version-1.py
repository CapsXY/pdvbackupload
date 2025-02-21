import pyodbc
import tkinter as tk
from tkinter import messagebox


def fazer_backup(server, database, username, password, backup_name):
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

            # Caminho do backup
            caminho_backup = f"C:\\Program Files\\Microsoft SQL Server\\MSSQL{16}.PDVNET\\MSSQL\\Backup\\{backup_name}.bak"

            # Comando de backup
            comando_backup = f"BACKUP DATABASE [{database}] TO DISK = N'{caminho_backup}' WITH NOFORMAT, NOINIT, NAME = '{database}-Full Database Backup', SKIP, NOREWIND, NOUNLOAD, STATS = 10"
            cursor.execute(comando_backup)

            messagebox.showinfo("Sucesso", "Backup realizado com sucesso!")

    except Exception as e:
        messagebox.showerror("Erro", str(e))


def criar_interface():
    root = tk.Tk()
    root.title("Backup do SQL Server")

    # Labels e entradas
    tk.Label(root, text="Servidor:").grid(row=0)
    tk.Label(root, text="Banco de Dados:").grid(row=1)
    tk.Label(root, text="Usuário:").grid(row=2)
    tk.Label(root, text="Senha:").grid(row=3)
    tk.Label(root, text="Nome do Backup:").grid(row=4)

    servidor = tk.Entry(root)
    banco_de_dados = tk.Entry(root)
    usuario = tk.Entry(root)
    senha = tk.Entry(root, show='*')
    nome_backup = tk.Entry(root)

    servidor.grid(row=0, column=1)
    banco_de_dados.grid(row=1, column=1)
    usuario.grid(row=2, column=1)
    senha.grid(row=3, column=1)
    nome_backup.grid(row=4, column=1)

    # Botão de backup
    tk.Button(root, text="Fazer Backup", command=lambda: fazer_backup(
        servidor.get(),
        banco_de_dados.get(),
        usuario.get(),
        senha.get(),
        nome_backup.get()
    )).grid(row=5, columnspan=2)

    root.mainloop()


if __name__ == "__main__":
    criar_interface()
