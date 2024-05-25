import streamlit as st
import pandas as pd
import requests
import re
from io import BytesIO
from zipfile import ZipFile

def download_imagem(url):
    try:
        response = requests.get(url, stream=True)
        response.raise_for_status()
        return response.content
    except requests.exceptions.RequestException as e:
        st.error(f"Falha ao baixar {url}: {e}")
        return None

def limpar_nome_arquivo(nome_arquivo):
    # Remove caracteres inválidos
    return re.sub(r'[<>:"/\\|?*]', '', nome_arquivo)

def criar_zip(links):
    zip_buffer = BytesIO()
    with ZipFile(zip_buffer, 'w') as zip_file:
        for url in links:
            conteudo_imagem = download_imagem(url)
            if conteudo_imagem:
                nome_arquivo = limpar_nome_arquivo(url.split("/")[-1].split('?')[0])
                zip_file.writestr(nome_arquivo, conteudo_imagem)
    return zip_buffer

st.set_page_config(page_title="Indev Ribas - Aplicativo de Extração e Download de Links de Excel")

st.title("Aplicativo de Extração e Download de Links de Excel - Indev Ribas")
st.sidebar.title("Instruções")

st.sidebar.markdown(
    """
    **Como usar:**

    1. Certifique-se de que os links estejam na coluna **N** do seu arquivo Excel.
    2. Faça o upload do arquivo Excel usando o botão abaixo.
    3. Clique no botão para baixar todas as imagens encontradas nos links.
    """
)

arquivo_upload = st.file_uploader("Escolha um arquivo Excel", type="xlsx")

if arquivo_upload is not None:
    try:
        df = pd.read_excel(arquivo_upload)

        # Assumindo que a coluna N é a 14ª coluna (índice 13)
        nome_coluna = df.columns[13]  # índice 13 corresponde à coluna N

        if nome_coluna:
            links = df[nome_coluna].dropna().astype(str).apply(lambda x: x if x.startswith('http') else None).dropna().tolist()

            if links:
                st.write(f"Encontrados {len(links)} links na coluna '{nome_coluna}':")
                for link in links:
                    st.write(link)

                if st.button("Baixar todas as imagens"):
                    zip_buffer = criar_zip(links)
                    st.download_button(
                        label="Baixar imagens em um arquivo ZIP",
                        data=zip_buffer.getvalue(),
                        file_name="imagens.zip",
                        mime="application/zip"
                    )
            else:
                st.write(f"Nenhum link encontrado na coluna '{nome_coluna}'.")
        else:
            st.error(f"Coluna '{nome_coluna}' não existe no arquivo enviado.")
    except Exception as e:
        st.error(f"Erro ao ler o arquivo: {e}")
