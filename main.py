import imaplib
from datetime import datetime
import email
import pandas as pd
import csv
import os
import subprocess
import psutil
from email.header import decode_header
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font
from openpyxl.styles.alignment import Alignment
from openpyxl.worksheet.datavalidation import DataValidation

# Função para encerrar o processo do Excel associado ao arquivo
def fechar_arquivo_excel(arquivo):
    for proc in psutil.process_iter():
        try:
            # Verifica se o processo é o Excel e se o arquivo está aberto nele
            if proc.name() == "EXCEL.EXE" and arquivo in proc.open_files():
                proc.kill()  # Encerra o processo do Excel
        except psutil.NoSuchProcess:
            pass

# Configurações da conexão com o servidor de e-mail
imap_server = "outlook.office365.com"  # Endereço do servidor IMAP do Outlook/Hotmail
username = "wendher.santos@outlook.com"  # Seu endereço de e-mail
password = "Wen010112!outlook"  # Sua senha

# Conectando ao servidor IMAP
with imaplib.IMAP4_SSL(imap_server) as mail:
    mail.login(username, password)

    # Selecionando a caixa de entrada
    mail.select("inbox")

    # Pesquisando e-mails na caixa de entrada
    result, data = mail.search(None, f'(TO "{username}")')  # Pesquisar e-mails para o endereço especificado

    # Lista para armazenar as tarefas
    tarefas = []

    # Iterar sobre os IDs dos e-mails encontrados
    for email_id in data[0].split():
        # Obtendo os dados do e-mail
        result, email_data = mail.fetch(email_id, "(RFC822)")

        # Analisando os dados do e-mail
        raw_email = email_data[0][1]
        msg = email.message_from_bytes(raw_email)  # Criar um objeto Message

        # Extraindo informações relevantes do e-mail
        remetente = email.utils.parseaddr(msg["From"])[1]  # Remetente do e-mail
        assunto = msg["Subject"]  # Assunto do e-mail
        data_email = msg["Date"]  # Data do e-mail

        # Convertendo a data para o formato DD/MM/AAAA
        parsed_date = email.utils.parsedate(data_email)
        data_formatada = datetime(*parsed_date[:6]).strftime("%d/%m/%Y")

        # Decodificando a descrição do e-mail
        descricao = " ".join(part.decode(encoding or 'utf-8') if isinstance(part, bytes) else part for part, encoding in decode_header(assunto))

        # Criando a tarefa como um dicionário
        tarefa = {
            "Remetente do E-mail": remetente,
            "Descrição do E-mail": descricao,
            "Data do E-mail": data_formatada,  # Convertido para tipo datetime do pandas
            "ID": email_id.decode("utf-8"),  # ID do e-mail para referência futura
            "Status": "Não Iniciado"  # Status inicial da tarefa
        }

        # Adicionando a tarefa à lista de tarefas
        tarefas.append(tarefa)

# Verificando se o arquivo CSV já existe
if os.path.exists("tarefas.csv"):
    # Encerrando o arquivo CSV caso esteja aberto no Excel
    fechar_arquivo_excel("tarefas.csv")

    # Lendo o arquivo CSV existente com a codificação correta
    df_existente = pd.read_csv("tarefas.csv", encoding="utf-8-sig")

    # Verificando se existem novas tarefas
    ids_existente = df_existente["ID"].tolist()
    novas_tarefas = [tarefa for tarefa in tarefas if tarefa["ID"] not in ids_existente]

    # Verificando se há novas tarefas para adicionar
    if novas_tarefas:
        # Criando um novo DataFrame com as tarefas existentes e as novas tarefas
        df_tarefas = pd.concat([df_existente, pd.DataFrame(novas_tarefas)])

        # Atualizando os status das tarefas existentes
        df_tarefas.update(df_existente)

        # Reindexando o DataFrame
        df_tarefas.reset_index(drop=True, inplace=True)
    else:
        # Não há novas tarefas, mantendo o DataFrame existente
        df_tarefas = df_existente
else:
    # Criando um DataFrame com as tarefas encontradas
    df_tarefas = pd.DataFrame(tarefas)

# Salvando o DataFrame atualizado no arquivo CSV
df_tarefas.to_csv("tarefas.csv", index=False, encoding="utf-8-sig")

# Criando um novo arquivo Excel
wb = Workbook()

# Criando a planilha "Tarefas" e preenchendo com os dados do DataFrame
ws_tarefas = wb.active
ws_tarefas.title = "Tarefas"

# Adicionando os cabeçalhos das colunas
for col_num, col_name in enumerate(df_tarefas.columns, start=1):
    col_letter = chr(64 + col_num)
    cell = ws_tarefas[f"{col_letter}1"]
    cell.value = col_name

# Preenchendo os dados das tarefas
for row in dataframe_to_rows(df_tarefas, index=False, header=True):
    ws_tarefas.append(row)

# Adicionando a lista suspensa de status usando validação de dados no Excel
status_options = ["Não Iniciado", "Em Andamento", "Concluído"]
status_formula = f'"{"|".join(status_options)}"'
status_column = ws_tarefas["E"]
status_column[0].value = "Status"

# Configurando a largura das colunas
column_widths = [20, 50, 12, 10, 12]
for col_num, width in enumerate(column_widths, start=1):
    col_letter = chr(64 + col_num)
    ws_tarefas.column_dimensions[col_letter].width = width

# Aplicando formatação à coluna de status
for cell in status_column[1:]:
    cell.alignment = Alignment(horizontal="center")  # Alinhamento centralizado

# Configurando a validação de dados na coluna de status
dv = DataValidation(type="list", formula1=f"={status_formula}", showDropDown=True)
dv.errorTitle = "Invalido!"
dv.error = "Escolha um valor da lista suspensa."
dv.prompt = "Selecione um valor da lista."
dv.promptTitle = "Status"
dv.add(status_column[1:])  # Ajuste o intervalo conforme necessário
ws_tarefas.add_data_validation(dv)

# Salvando o arquivo Excel
wb.save("tarefas.xlsx")
