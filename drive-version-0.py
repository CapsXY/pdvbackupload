from pathlib import Path
import pyodbc
import os
import tkinter as tk
from datetime import datetime
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google.oauth2 import service_account
from tkinter import ttk, messagebox


# Função para conectar ao SQL Server
def conectar_sql_server(servidor, usuario, senha, banco_de_dados):
    try:
        string_conexao = f"DRIVER={{SQL Server}};SERVER={servidor};DATABASE={banco_de_dados};UID={usuario};PWD={senha}"
        conexao = pyodbc.connect(string_conexao)
        return conexao
    except pyodbc.Error as ex:
        sqlstate = ex.args[0]
        print(f"Erro ao conectar ao SQL Server: {sqlstate}")
        messagebox.showerror("Erro de Conexão",
                             f"Falha ao conectar ao SQL Server: {sqlstate}")  # Mensagem de erro na tela
        return None
    except Exception as e:
        print(f"Erro inesperado na conexão: {e}")
        messagebox.showerror("Erro de Conexão", f"Falha ao conectar ao SQL Server: {e}")  # Mensagem de erro na tela
        return None


# Função para realizar o backup
def realizar_backup(conexao, nome_arquivo):
    try:
        cursor = conexao.cursor()

        # Obtém o nome do banco de dados
        cursor.execute("SELECT DB_NAME()")
        nome_banco_dados = cursor.fetchone()[0]
        print(f"Nome do banco de dados: {nome_banco_dados}")  # Imprime o nome do banco de dados

        nome_arquivo = nome_arquivo + ".bak"
        pasta_backup = Path(r"C:\Backup_Bancos")
        caminho_backup_local = pasta_backup / nome_arquivo
        caminho_backup_temp = Path.cwd() / nome_arquivo

        print(f"Caminho backup temp: {caminho_backup_temp}")  # Imprime o caminho temporário
        print(f"Caminho backup local: {caminho_backup_local}")  # Imprime o caminho local

        # Comando BACKUP DATABASE com parâmetros
        comando_backup = "BACKUP DATABASE ? TO DISK = ?"
        cursor.execute(comando_backup, (nome_banco_dados, str(caminho_backup_temp)))
        conexao.commit()
        print(f"Backup do banco de dados {nome_banco_dados} realizado com sucesso!")
        return str(caminho_backup_local)
    except pyodbc.Error as ex:
        sqlstate = ex.args[0]
        print(f"Erro ao realizar o backup: {sqlstate}")
        messagebox.showerror("Erro de Backup", f"Falha ao realizar o backup: {sqlstate}")
        return None
    except OSError as ex:
        print(f"Erro de sistema ao realizar o backup: {ex}")
        messagebox.showerror("Erro de Backup", f"Falha ao realizar o backup (erro de sistema): {ex}")
        return None
    except Exception as e:
        print(f"Erro inesperado no backup: {e}")
        messagebox.showerror("Erro de Backup", f"Falha ao realizar o backup: {e}")
        return None


# Função para enviar o arquivo de backup para o Google Drive
def enviar_para_drive(caminho_arquivo):
    # Caminho para o arquivo JSON com as credenciais
    CREDENTIALS_FILE = 'credenciais.json'

    # Escopos de acesso necessários
    SCOPES = ['https://www.googleapis.com/auth/drive.file']

    # Autenticação
    credentials = service_account.Credentials.from_service_account_file(
        CREDENTIALS_FILE, scopes=SCOPES)
    service = build('drive', 'v3', credentials=credentials)

    try:
        file_metadata = {'name': os.path.basename(caminho_arquivo)}
        media = MediaFileUpload(caminho_arquivo, mimetype='application/octet-stream')
        file = service.files().create(body=file_metadata, media_body=media, fields='id').execute()
        print(f"Arquivo enviado para o Google Drive com ID: {file.get('id')}")
        messagebox.showinfo("Sucesso", f"Backup enviado para o Google Drive com ID: {file.get('id')}")
    except Exception as e:
        print(f"Erro ao enviar para o Google Drive: {e}")
        messagebox.showerror("Erro no Drive", f"Falha ao enviar o backup para o Google Drive: {e}")


# Função para executar o processo de backup e upload
def executar_backup(servidor, usuario, senha, banco_de_dados, nome_arquivo):
    try:
        conexao = conectar_sql_server(servidor, usuario, senha, banco_de_dados)
        if conexao:
            try:
                caminho_backup = realizar_backup(conexao, nome_arquivo)
                if caminho_backup:
                    enviar_para_drive(caminho_backup)
            except Exception as e:
                print(f"Erro durante o processo de backup: {e}")
                messagebox.showerror("Erro", f"Falha durante o backup: {e}")  # Mensagem de erro na tela
            finally:
                conexao.close()
    except Exception as e:
        print(f"Erro na função executar_backup: {e}")
        messagebox.showerror("Erro", f"Falha na função executar_backup: {e}")  # Mensagem de erro na tela


# Função para abrir a tela de input de dados
def abrir_tela_input():
    def realizar_backup_com_dados():
        # Dados de login fixos
        usuario = "sa"
        senha = "inter#system"
        banco_de_dados = entry_banco_de_dados.get()
        nome_arquivo = entry_nome_arquivo.get()

        # Obtém o nome base do servidor
        nome_base_servidor = entry_nome_servidor.get()
        # Concatena "\PDVNET" ao nome do servidor
        servidor = nome_base_servidor + ""

        # Validação do nome do banco de dados e nome do arquivo
        if not banco_de_dados or not nome_arquivo or not nome_base_servidor:
            messagebox.showerror("Erro", "Todos os campos devem ser preenchidos.")
            return

        executar_backup(servidor, usuario, senha, banco_de_dados, nome_arquivo)
        tela_input.destroy()

    tela_input = tk.Tk()
    tela_input.title("Dados de Conexão")

    # Label e entrada para o nome base do servidor
    label_nome_servidor = ttk.Label(tela_input, text="Nome do Servidor:")
    label_nome_servidor.grid(row=0, column=0, padx=5, pady=5)
    entry_nome_servidor = ttk.Entry(tela_input)
    entry_nome_servidor.grid(row=0, column=1, padx=5, pady=5)

    # Labels e entradas para banco de dados e nome do arquivo
    label_banco_de_dados = ttk.Label(tela_input, text="Banco de Dados:")
    label_banco_de_dados.grid(row=1, column=0, padx=5, pady=5)
    entry_banco_de_dados = ttk.Entry(tela_input)
    entry_banco_de_dados.grid(row=1, column=1, padx=5, pady=5)

    label_nome_arquivo = ttk.Label(tela_input, text="Nome do Arquivo:")
    label_nome_arquivo.grid(row=2, column=0, padx=5, pady=5)
    entry_nome_arquivo = ttk.Entry(tela_input)
    entry_nome_arquivo.grid(row=2, column=1, padx=5, pady=5)

    botao_executar = ttk.Button(tela_input, text="Executar Backup", command=realizar_backup_com_dados)
    botao_executar.grid(row=3, column=0, columnspan=2, padx=5, pady=10)

    tela_input.mainloop()


# Chama a função para abrir a tela de input
abrir_tela_input()