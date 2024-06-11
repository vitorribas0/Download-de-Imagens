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
    return re.sub(r'[<>:"/\\|?*]', '', nome_arquivo)

def criar_zip(links_nomes_projetos):
    zip_buffer = BytesIO()
    with ZipFile(zip_buffer, 'w') as zip_file:
        for url, nome_colaborador, nome_projeto in links_nomes_projetos:
            conteudo_imagem = download_imagem(url)
            if conteudo_imagem:
                # Renomear arquivos com o nome do colaborador e o nome do projeto
                extensao = url.split(".")[-1].split('?')[0]  # Obter extensão do arquivo
                nome_arquivo = limpar_nome_arquivo(f"{nome_colaborador}_{nome_projeto}.{extensao}")
                zip_file.writestr(nome_arquivo, conteudo_imagem)
    return zip_buffer

st.set_page_config(page_title="Indev Ribas - Aplicativo de Extração e Download de Links de Excel")

st.title("Aplicativo de Extração e Download de Links de Excel - Indev Ribas")
st.sidebar.title("Instruções")

st.sidebar.markdown(
    """
    **Como usar:**

    1. Certifique-se de que os links estejam na coluna **N** do seu arquivo Excel.
    2. Certifique-se de que o nome do colaborador esteja na coluna **A** e o nome do projeto esteja na coluna **H**, todos na mesma linha do link.
    3. Faça o upload do arquivo Excel usando o botão abaixo.
    4. Clique no botão para baixar todas as imagens encontradas nos links.
    """
)

arquivo_upload = st.file_uploader("Escolha um arquivo Excel", type="xlsx")

if arquivo_upload is not None:
    try:
        df = pd.read_excel(arquivo_upload)

        # Assumindo que a coluna N é a 14ª coluna (índice 13), a coluna A é a primeira coluna (índice 0) e a coluna H é a oitava coluna (índice 7)
        coluna_links = df.columns[13]  # índice 13 corresponde à coluna N
        coluna_nomes = df.columns[0]   # índice 0 corresponde à coluna A
        coluna_projetos = df.columns[1] # índice 7 corresponde à coluna K

        if coluna_links and coluna_nomes and coluna_projetos:
            # Certificar que temos o mesmo número de entradas em todas as colunas relevantes
            df_filtrado = df[[coluna_links, coluna_nomes, coluna_projetos]].dropna()

            links_nomes_projetos = [
                (row[coluna_links], row[coluna_nomes], row[coluna_projetos])
                for idx, row in df_filtrado.iterrows()
                if row[coluna_links].startswith('http')
            ]

            if links_nomes_projetos:
                st.write(f"Encontrados {len(links_nomes_projetos)} links na coluna '{coluna_links}' com respectivos nomes e datas de criação nas colunas '{coluna_nomes}' e '{coluna_projetos}':")
                for link, nome, projeto in links_nomes_projetos:
                    st.write(f"{nome} - {projeto}: {link}")

                if st.button("Baixar todas as imagens"):
                    zip_buffer = criar_zip(links_nomes_projetos)
                    st.download_button(
                        label="Baixar imagens em um arquivo ZIP",
                        data=zip_buffer.getvalue(),
                        file_name="imagens.zip",
                        mime="application/zip"
                    )
            else:
                st.write(f"Nenhum link ou nome correspondente encontrado nas colunas '{coluna_links}', '{coluna_nomes}' e '{coluna_projetos}'.")
        else:
            st.error(f"Coluna '{coluna_links}', '{coluna_nomes}' ou '{coluna_projetos}' não existe no arquivo enviado.")
    except Exception as e:
        st.error(f"Erro ao ler o arquivo: {e}")
