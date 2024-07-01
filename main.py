import pandas as pd
import streamlit as st
from PIL import Image
import operations as op

#---DEFINIÇÃO DE VARIÁVEIS---
layout = 'centered'
page_title = 'File Management Web App'
page_icon = '📂'
img = Image.open('./assets/files_pic.png')
autor = {'Andrés A. Michelangelli': 'https://andresarturomichelangelli.streamlit.app/'}

#---DEFINIÇÃO DE CONFIGURAÇÕES ADICIONAIS---
#Configurar pagina
st.set_page_config(page_title=page_title,
                    page_icon=page_icon,
                      layout=layout)
#Redimensionar a imagem principal
#img_resize = img.resize((200, 50)) #(width, height)

# Inserindo a imagem na página utilizando os comandos do stremalit
st.image(img, use_column_width=True)

#Mostrar titulo e descrição do aplicativo web
st.write("<div align='center'><h2><i>File Management Web App</i></h2></div>",
        unsafe_allow_html=True)
st.write("")
st.write('Este aplicativo pode manipular arquivos seguindo uma lista definida pelo usuario.')
st.write('Este aplicativo é útil quando é necessário organizar ou gestionar certos arquivos específicos numa pasta com muitos outros arquivos. Para mais informação por favor verifique a aba "Instruções".')
for nome, link in autor.items():
    st.write(f'Este aplicativo web foi desenvolvido por: [{nome}]({link})')
st.write(' ')


#Mostrar input boxes e drag n drop
operacao = st.selectbox('Operação:', ['Copiar', 'Mover', 'Remover', 'Gerar lista .xls de uma pasta', 'Salvar XLS a PDF (a partir de uma lista)', 'Salvar XLS a PDF (todos)', 'Organizar arquivos'])

if operacao in ['Copiar', 'Mover', 'Remover', 'Salvar XLS a PDF (a partir de uma lista)']:
    pasta_origem = st.text_input('Pasta origem:')
    pasta_destino = st.text_input('Pasta destino:')
    arquivo_xlsx = st.file_uploader('Carregar lista a partir de um arquivo .xlsx: ', type='xlsx', accept_multiple_files=False)
elif operacao in ['Gerar lista .xls de uma pasta', 'Salvar XLS a PDF (todos)', 'Organizar arquivos']:
    pasta_origem = st.text_input('Pasta origem:')
    pasta_destino = st.text_input('Pasta destino:')

#Condicionais que mostram advertencias para o usuário
if operacao == 'Remover' and pasta_origem:
    st.warning(f'Atenção: A operação irá remover permanentemente os arquivos da lista carregada, localizados na pasta de origem {pasta_origem} indicada pelo usuário.')
elif operacao == 'Gerar lista .xls de uma pasta':
    st.warning(f'Atenção: A operação irá criar uma lista em formato .xlsx dos arquivos localizados na pasta origem {pasta_origem}. A lista .xlsx será gerada na pasta destino {pasta_destino}')
elif operacao == 'Salvar XLS a PDF (a partir de uma lista)':
    st.warning(f'Atenção: A operação irá salvar os arquivos .xlsx localizados na pasta origem {pasta_origem} na pasta destino {pasta_destino} a partir de uma lista em formato .xlsx definida pelo usuário')
elif operacao == 'Organizar arquivos':
    st.warning(f'Atenção: A operação irá organizar os arquivos apenas na pasta origem {pasta_origem}.')
elif operacao == 'Salvar XLS a PDF (todos)':
    st.warning(f'Atenção: A operação irá salvar em PDF todos os arquivos .xlsx da pasta origem {pasta_origem} na pasta destino {pasta_destino} definida pelo usuário')

#Mostrar mensagem de sucesso desde que os dados input sejam fornecidos corretamente para cada tipo de operação
if pasta_origem and pasta_destino and operacao in ['Copiar', 'Mover', 'Remover', 'Salvar XLS a PDF (a partir de uma lista)'] and arquivo_xlsx:
    st.success('Os dados de entrada foram preenchidos com sucesso.', icon="✅")
    df_xlsx = pd.read_excel(arquivo_xlsx)
    st.write(df_xlsx.head(5))
elif pasta_origem and pasta_destino and operacao in ['Gerar lista .xls de uma pasta', 'Salvar XLS a PDF (todos)', 'Organizar arquivos']:
    st.success('Os dados de entrada foram preenchidos com sucesso.', icon="✅")


#Condicionais para iniciar operação após presionar o botão 'Iniciar operação'--------------------------->

#Operações que precisam de pasta de origem e pasta de destino
if st.button('Iniciar operação') and pasta_origem:
    if operacao == 'Mover' and arquivo_xlsx and pasta_destino:
        tempo_total, arquivos_processados = op.mover_arquivos(df_xlsx, pasta_origem, pasta_destino)
        st.success(f'A operação "{operacao}" foi concluída com sucesso em {tempo_total:.2f} segundos. {arquivos_processados} arquivos movidos.')
    
    elif operacao == 'Copiar' and arquivo_xlsx and pasta_destino:
        tempo_total, arquivos_processados = op.copiar_arquivos(df_xlsx, pasta_origem, pasta_destino)
        st.success(f'A operação "{operacao}" foi concluída com sucesso em {tempo_total:.2f} segundos. {arquivos_processados} arquivos copiados.')
    
    elif operacao == 'Gerar lista .xls de uma pasta' and pasta_destino:
        op.criar_lista_arquivos(pasta_origem, pasta_destino)
        st.success(f'A operação foi concluída com sucesso. Uma lista dos arquivos contidos na pasta {pasta_origem} foi gerada na pasta {pasta_destino}, com o nome "Lista_arquivos".')

    elif operacao == 'Salvar XLS a PDF (a partir de uma lista)' and arquivo_xlsx and pasta_destino:
        tempo_total, arquivos_processados = op.salvar_pdf(df_xlsx, pasta_origem, pasta_destino)
        st.success(f'A operação "{operacao}" foi concluída com sucesso em {tempo_total:.2f} segundos. {arquivos_processados} arquivos XLS salvos em PDF.')
    
    elif operacao == 'Salvar XLS a PDF (todos)' and pasta_destino:
        tempo_total, arquivos_processados = op.salvar_pdf_todos(pasta_origem, pasta_destino)
        st.success(f'A operação "{operacao}" foi concluída com sucesso em {tempo_total:.2f} segundos. {arquivos_processados} arquivos XLS salvos em PDF.')
       
    #Operações que não precisam da pasta destino
    elif operacao == 'Organizar arquivos':
        op.organizar_arquivos(pasta_origem)
        st.success(f'A operação foi concluída com sucesso. Os arquivos da pasta {pasta_origem} foram organizados.')
    
    elif operacao == 'Remover' and arquivo_xlsx:
        tempo_total, arquivos_processados = op.remover_arquivos(df_xlsx, pasta_origem)
        st.success(f'A operação "{operacao}" foi concluída com sucesso em {tempo_total:.2f} segundos. {arquivos_processados} arquivos removidos.')
    