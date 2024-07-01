import os
import shutil
import pandas as pd
import streamlit as st
import time
import platform

from win32com import client #biblioteca pip install pywin32
import pythoncom

#---Função para mover arquivos---
def mover_arquivos(df, pasta_origem, pasta_destino):
    start_time = time.time()

    progress_text = "Operação em andamento. Por favor aguarde."
    my_bar = st.progress(0, text=progress_text)
    
    total_arquivos = len(df)
    arquivos_processados = 0

    for idx, arquivo in enumerate(df.iloc[:, 0]):
        caminho_origem = os.path.join(pasta_origem, arquivo)
        caminho_destino = os.path.join(pasta_destino, arquivo)
        
        if os.path.exists(caminho_origem):
            shutil.move(caminho_origem, caminho_destino)
            arquivos_processados += 1
        
        tempo_andamento = time.time() - start_time

        # Atualiza a barra de progresso
        my_bar.progress((idx + 1) / total_arquivos, text=f'{progress_text} Tempo decorrido: {tempo_andamento:.2f} segundos')
        time.sleep(0.5)
    
    my_bar.empty()
    end_time = time.time()
    tempo_total = end_time - start_time

    return tempo_total, arquivos_processados

#---Função para copiar arquivos---
def copiar_arquivos(df, pasta_origem, pasta_destino):
    start_time = time.time()

    progress_text = "Operação em andamento. Por favor aguarde."
    my_bar = st.progress(0, text=progress_text)
    
    total_arquivos = len(df)
    arquivos_processados = 0

    for idx, arquivo in enumerate(df.iloc[:, 0]):
        caminho_origem = os.path.join(pasta_origem, arquivo)
        caminho_destino = os.path.join(pasta_destino, arquivo)
        
        if os.path.exists(caminho_origem):
            shutil.copy(caminho_origem, caminho_destino)
            arquivos_processados += 1
        
        tempo_andamento = time.time() - start_time

        # Atualiza a barra de progresso
        my_bar.progress((idx + 1) / total_arquivos, text=f'{progress_text} Tempo decorrido: {tempo_andamento:.2f} segundos')
        time.sleep(0.5)
    
    my_bar.empty()
    end_time = time.time()
    tempo_total = end_time - start_time

    return tempo_total, arquivos_processados

#---Função para remover arquivos---
def remover_arquivos(df, pasta_origem):
    start_time = time.time()

    progress_text = "Operação em andamento. Por favor aguarde."
    my_bar = st.progress(0, text=progress_text)
    
    total_arquivos = len(df)
    arquivos_processados = 0

    for idx, arquivo in enumerate(df.iloc[:, 0]):
        caminho_origem = os.path.join(pasta_origem, arquivo)
        
        if os.path.exists(caminho_origem):
            os.remove(caminho_origem)
            arquivos_processados += 1
        
        tempo_andamento = time.time() - start_time

        # Atualiza a barra de progresso
        my_bar.progress((idx + 1) / total_arquivos, text=f'{progress_text} Tempo decorrido: {tempo_andamento:.2f} segundos')
        time.sleep(0.5)
    
    my_bar.empty()
    end_time = time.time()
    tempo_total = end_time - start_time

    return tempo_total, arquivos_processados

#---Função para criar lista de arquivos de uma pasta----
def criar_lista_arquivos(pasta_origem, pasta_destino):
    arquivos = os.listdir(pasta_origem)
    nome = 'Lista_arquivos.xlsx'
    df = pd.DataFrame(arquivos, columns=['Lista de arquivos'])
    xlsx_filename = os.path.join(pasta_destino, nome)
    df.to_excel(xlsx_filename, index=False)

#---Função para salvar arquivos .xlsx em .pdf
def salvar_pdf(df, pasta_origem, pasta_destino):
    
    if platform.system() == 'Windows':
        import win32com.client as client
        import pythoncom
    
    if platform.system() != 'Windows':
        st.error("Esta funcionalidade só está disponível no Windows.")
        return
    
    start_time = time.time()

    progress_text = "Operação em andamento. Por favor aguarde."
    my_bar = st.progress(0, text=progress_text)
    
    total_arquivos = len(df)
    arquivos_processados = 0

    pythoncom.CoInitialize()  # Inicializar a biblioteca COM.
    # COM permite que programas escritos em diferentes linguagens de programação interajam uns com os outros. Por exemplo, utilizar Python para controlar e interagir com o Excel, que é escrito em C++.

    try:
        for idx, arquivo in enumerate(df.iloc[:, 0]):
            caminho_origem = os.path.join(pasta_origem, arquivo)
            caminho_destino = os.path.join(pasta_destino, arquivo)

            if os.path.exists(caminho_origem) and arquivo.endswith('.xlsx'):
                # Operação para salvar cada PDF num loop
                app = client.DispatchEx('Excel.Application')
                app.Interactive = False
                app.Visible = False

                # Separando o nome do arquivo da extensão .xlsx e criando o caminho do PDF
                destino, extensao = os.path.splitext(caminho_destino)
                destino_pdf = f"{destino}.pdf"
                
                workbook = app.Workbooks.Open(caminho_origem)
                workbook.ActiveSheet.ExportAsFixedFormat(0, destino_pdf)
                workbook.Close(SaveChanges=False)
                app.Quit()

                arquivos_processados += 1
            
            tempo_andamento = time.time() - start_time

            # Atualiza a barra de progresso
            my_bar.progress((idx + 1) / total_arquivos, text=f'{progress_text} Tempo decorrido: {tempo_andamento:.2f} segundos')
            time.sleep(0.5)
    finally:
        pythoncom.CoUninitialize()  # Desinicializar a biblioteca COM
    
    my_bar.empty()
    end_time = time.time()
    tempo_total = end_time - start_time

    return tempo_total, arquivos_processados

#---Função para organizar arquivos---
def organizar_arquivos(pasta_origem):
    extensions = {
        '.jpg': 'Imagens',
        '.png': 'Imagens',
        '.gif': 'Imagens',
        '.mp4': 'Videos',
        '.avi': 'Videos',
        '.mov': 'Videos',
        '.doc': 'Documentos',
        '.docx': 'Documentos',
        '.pdf': 'Documentos',
        '.txt': 'Documentos',
        '.csv': 'Documentos',
        '.xlsx': 'Documentos',
        '.mp3': 'Audio',
        '.wav': 'Audio',
        '.zip': 'Comprimidos',
        '.7z': 'Comprimidos',
        '.rar': 'Comprimidos',
        '.exe': 'Aplicativos'
    }

    for filename in os.listdir(pasta_origem):
        file_path = os.path.join(pasta_origem, filename)
        
        if os.path.isfile(file_path):
            extension = os.path.splitext(filename)[1].lower()
            
            if extension in extensions:
                folder_name = extensions[extension]

                folder_path = os.path.join(pasta_origem, folder_name)
                os.makedirs(folder_path, exist_ok=True)

                destination_path = os.path.join(folder_path, filename)
                shutil.move(file_path, destination_path)

#---Função para salvar todos os arquivos XLS a PDF de uma pasta indicada pelo usuário---
def salvar_pdf_todos(pasta_origem, pasta_destino):
        
    if platform.system() == 'Windows':
        import win32com.client as client
        import pythoncom
    
    if platform.system() != 'Windows':
        st.error("Esta funcionalidade só está disponível no Windows.")
        return
    
    start_time = time.time()

    #Indicar que o df será a lista gerada dos arquivos na pasta de origem
    arquivos = os.listdir(pasta_origem)
    df = pd.DataFrame(arquivos, columns=['Lista de arquivos'])

    progress_text = "Operação em andamento. Por favor aguarde."
    my_bar = st.progress(0, text=progress_text)
    
    total_arquivos = len(df)
    arquivos_processados = 0

    pythoncom.CoInitialize()  # Inicializar a biblioteca COM.
    # COM permite que programas escritos em diferentes linguagens de programação interajam uns com os outros. Por exemplo, utilizar Python para controlar e interagir com o Excel, que é escrito em C++.

    try:
        for idx, arquivo in enumerate(df.iloc[:, 0]):
            caminho_origem = os.path.join(pasta_origem, arquivo)
            caminho_destino = os.path.join(pasta_destino, arquivo)

            if os.path.exists(caminho_origem) and arquivo.endswith('.xlsx'): #aqui é possível adiional mais uma condicional para verificar se o arquivo tem extensão .xlsx
                # Operação para salvar cada PDF num loop
                app = client.DispatchEx('Excel.Application')
                app.Interactive = False
                app.Visible = False

                # Separando o nome do arquivo da extensão .xlsx e criando o caminho do PDF
                destino, extensao = os.path.splitext(caminho_destino)
                destino_pdf = f"{destino}.pdf"
                
                workbook = app.Workbooks.Open(caminho_origem)
                workbook.ActiveSheet.ExportAsFixedFormat(0, destino_pdf)
                workbook.Close(SaveChanges=False)
                app.Quit()

                arquivos_processados += 1
            
            tempo_andamento = time.time() - start_time

            # Atualiza a barra de progresso
            my_bar.progress((idx + 1) / total_arquivos, text=f'{progress_text} Tempo decorrido: {tempo_andamento:.2f} segundos')
            time.sleep(0.5)
    finally:
        pythoncom.CoUninitialize()  # Desinicializar a biblioteca COM
    
    my_bar.empty()
    end_time = time.time()
    tempo_total = end_time - start_time

    return tempo_total, arquivos_processados
