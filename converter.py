import pandas as pd
import sqlite3
import subprocess
import os

class Converter:
    def converter_para_excel(self, arquivo_db):
        with sqlite3.connect(arquivo_db) as conexao:
            cursor=conexao.cursor()
            cursor. execute("SELECT name FROM sqlite_master WHERE type='table'")
            tabelas = cursor.fetchall()
            if tabelas:
                for tabela in tabelas:
                    nome_tabela = tabela[0]
                    query = f"SELECT * FROM {nome_tabela};"
                    dados_do_banco = pd.read_sql_query(query,conexao)
                        
                    pasta_downloads = os.path.join(os.path.expanduser("~"), "Downloads")
                    diretorio_arquivos_convertidos = os.path.join(pasta_downloads,"arquivos_convertidos")
                    os.makedirs(diretorio_arquivos_convertidos, exist_ok=True)
                    arquivo_excel = os.path.join(diretorio_arquivos_convertidos, f"{nome_tabela}.xlsx")
                    dados_do_banco.to_excel(arquivo_excel, index=False)
            subprocess.Popen(["explorer","/select,", diretorio_arquivos_convertidos])
               
        