import streamlit as st

st.title('Instruções de uso ⚠️')
st.write(' ')
st.write('Preencher a informação solicitada e assim que a mensagem "Os dados de entrada foram preenchidos com sucesso." aparecer, clicar no botão "Iniciar operação".')
st.write('É possível gestionar os arquivos do seu computador através das seguintes sete (07) operações a seguir:')
st.write('')
st.subheader('1- Copiar')
st.write('Esta operação permite copiar arquivos de uma pasta de origem para uma pasta de destino seguindo a ordem de arquivos em uma lista XLS previamente definida pelo usuário, para selecionar quais arquivos copiar da pasta de origem. Esta operação é muito útil quando os arquivos procurados encontram-se em uma pasta com muitos outros arquivos.')
st.write('Importante: A lista do arquivo XLS fornecida pelo usuario somente será lida na primeira coluna A1. Os nomes dos arquivos na lista devem conter a extensão do arquivo.')
st.write('')

st.subheader('2- Mover')
st.write('Esta operação permite mover arquivos de uma pasta de origem para uma pasta de destino seguindo a ordem de arquivos em uma lista XLS previamente definida pelo usuário, para selecionar quais arquivos mover da pasta de origem. Esta operação é muito útil quando os arquivos procurados encontram-se em uma pasta com muitos outros arquivos.')
st.write('Importante: A lista do arquivo XLS fornecida pelo usuario somente será lida na primeira coluna A1. Os nomes dos arquivos na lista devem conter a extensão do arquivo.')
st.write('')

st.subheader('3- Remover')
st.write('Esta operação permite remover arquivos de uma pasta de origem ,seguindo a ordem de arquivos em uma lista XLS previamente definida pelo usuário, para selecionar quais arquivos remover da pasta de origem. Esta operação é muito útil quando os arquivos procurados encontram-se em uma pasta com muitos outros arquivos.')
st.write('Importante: A lista do arquivo XLS fornecida pelo usuario somente será lida na primeira coluna A1. Os nomes dos arquivos na lista devem conter a extensão do arquivo.')
st.write('')

st.subheader('4- Gerar lista .xls de uma pasta')
st.write('Esta operação permite criar uma lista em formato .xlsx dos arquivos localizados numa pasta origem. A lista .xlsx será gerada na pasta destino indicada pelo usuário. A lista gerada irá incluir os nomes das pastas adicionais que estiverem na pasta origem indicada pelo usuário.')
st.write('')

st.subheader('5- Salvar XLS a PDF (a partir de uma lista)')
st.write('Esta operação irá salvar em PDF os arquivos .xlsx da pasta origem na pasta destino a partir de uma lista em formato .xlsx fornecida pelo usuário.')
st.write('Importante: A lista fornecida pelo usuário somente será lida na primeira coluna A1. Os nomes dos arquivos na lista devem conter a extensão ".xlsx" ao final do nome, pois somente serão processados os arquivos com essa extensão.')
st.write('')

st.subheader('6- Salvar XLS a PDF (todos)')
st.write('Esta operação permite salvar em PDF todos os arquivos .xlsx da pasta origem na pasta destino definida pelo usuário')
st.write('Importante: Somente serão salvos em PDF os arquivos XLS que tenham a extensão ".xlsx"')
st.write('')

st.subheader('7- Organizar arquivos')
st.write('Esta operação permite organizar os arquivos apenas na pasta origem, classificando e movendo os arquivos de acordo com a extensão que possuem em pastas como imagens, videos, audio ou documentos.')
st.write('Importante: Fechar todos os arquivos da pasta que estejam abertos, para evitar erros.')


