import os
import streamlit as st
import pandas as pd
import requests
import re
from io import BytesIO

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

def salvar_imagens(links):
    pasta_downloads = os.path.join(os.path.expanduser("~"), "Downloads")  # Caminho para a pasta "Downloads" do usuário

    for url in links:
        conteudo_imagem = download_imagem(url)
        if conteudo_imagem:
            # Extrair o nome do arquivo da URL e sanitizá-lo
            nome_arquivo = limpar_nome_arquivo(url.split("/")[-1].split('?')[0])
            caminho_arquivo = os.path.join(pasta_downloads, nome_arquivo)  # Caminho completo do arquivo
            with open(caminho_arquivo, "wb") as arquivo:
                arquivo.write(conteudo_imagem)
            st.success(f"Baixado {nome_arquivo}")

st.set_page_config(page_title="Indev Ribas - Aplicativo de Extração e Download de Links de Excel")

st.title("Aplicativo de Extração e Download de Links de Excel - Indev Ribas")
st.sidebar.title("Instruções")

st.sidebar.markdown(
    """
    **Instruções:**

    Para carregar o arquivo, os links devem estar na coluna N.
    """
)

arquivo_upload = st.file_uploader("Escolha um arquivo Excel", type="xlsx")

if arquivo_upload is not None:
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
                salvar_imagens(links)
        else:
            st.write(f"Nenhum link encontrado na coluna '{nome_coluna}'.")
    else:
        st.error(f"Coluna '{nome_coluna}' não existe no arquivo enviado.")
