import streamlit as st
import pandas as pd
import os
from datetime import datetime
from PIL import Image
import base64
import time

# Adicionando lógica para atualização automática
if 'last_refresh' not in st.session_state:
    st.session_state['last_refresh'] = time.time()

# Define o intervalo de atualização em segundos (ex: 300 segundos = 5 minutos)
refresh_interval = 300

if time.time() - st.session_state['last_refresh'] > refresh_interval:
    st.experimental_rerun()

# Configurando o layout da página
st.set_page_config(layout="wide")

# Obtendo a data atual
data_atual = datetime.now()
ano_atual = data_atual.year
mes_atual = data_atual.month
dia_atual = data_atual.day

# Construindo o caminho do arquivo
pasta_base = r'\\192.168.3.100\Tuopu\PRODUÇÃO\01-Apontamentos\H-H'
pasta_ano = os.path.join(pasta_base, str(ano_atual))
pasta_mes = os.path.join(pasta_ano, f"{ano_atual}_{mes_atual}")
nome_arquivo = f"Backup H-H_{dia_atual}_{mes_atual}_{ano_atual}.xlsb"
caminho_arquivo = os.path.join(pasta_mes, nome_arquivo)

# Verificando se o arquivo existe
if not os.path.exists(caminho_arquivo):
    st.error(f"Arquivo não encontrado: {caminho_arquivo}")
else:
    # Lendo o arquivo Excel
    df = pd.read_excel(caminho_arquivo, sheet_name='REAL', skiprows=97, engine='pyxlsb')
    
    # Removendo a primeira linha de dados
    df = df.iloc[1:]

    # Selecionando apenas as primeiras 18 colunas
    df = df.iloc[:, :18]

    # Contando o número de linhas na tabela
    qtdLinhas = df.shape[0]

    # Selecionando as colunas relevantes e removendo colunas não nomeadas
    df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
    df = df.set_index(df.columns[0])

    # Filtrando as colunas relevantes
    cols_m = [col for col in df.columns if col.startswith('M')]
    cols_v = [col for col in df.columns if col.startswith('V')]

    # Cálculo da média das colunas ignorando valores vazios
    porcentagens_m = df[cols_m].apply(pd.to_numeric, errors='coerce').mean(skipna=True).values
    porcentagens_v = df[cols_v].apply(pd.to_numeric, errors='coerce').mean(skipna=True).values

    # Função para determinar a cor do card
    def get_color(value):
        if pd.isna(value):
            return 'background-color: grey;'
        elif value <= 0.65:
            return 'background-color: red;'
        elif value <= 0.85:
            return 'background-color: orange;'
        else:
            return 'background-color: green;'

    # Função para criar cards com tooltip estilizado
    def create_card(title, value):
        if pd.isna(value):
            display_value = "Não             programado"
            color = 'background-color: grey;'
            tooltip_text = "Nenhum dado disponível para esta métrica."
        else:
            display_value = f"{value:.2%}"
            color = get_color(value)
            tooltip_text = f"Valor atual: {display_value}"
        
        card_html = f"""
        <style>
            .tooltip {{
                position: relative;
                display: inline-block;
                cursor: pointer;
            }}
            
            .tooltip .tooltiptext {{
                visibility: hidden;
                width: 160px;
                background-color: black;
                color: #fff;
                text-align: center;
                border-radius: 5px;
                padding: 5px 0;
                position: absolute;
                z-index: 1;
                bottom: 125%; /* Position above the card */
                left: 50%;
                margin-left: -80px;
                opacity: 0;
                transition: opacity 0.3s;
            }}
            
            .tooltip:hover .tooltiptext {{
                visibility: visible;
                opacity: 1;
            }}
        </style>
        
        <div class="tooltip" style="{color} padding: 10px; border-radius: 10px; margin: 10px; text-align: center; color: white;">
            <h4>{title}</h4>
            <p style="font-size: 24px;">{display_value}</p>
            <span class="tooltiptext">{tooltip_text}</span>
        </div>
        """
        st.markdown(card_html, unsafe_allow_html=True)

    # Função para converter imagem em base64
    def image_to_base64(image):
        from io import BytesIO
        buffer = BytesIO()
        image.save(buffer, format="PNG")
        img_str = base64.b64encode(buffer.getvalue()).decode("utf-8")
        return img_str

    # Carregando a imagem local
    image = Image.open("logoTuopu.png")
    image_base64 = image_to_base64(image)

    # Exibindo o título e a imagem lado a lado
    st.markdown(
        f"""
        <div style="display: flex; align-items: center;">
            <img src="data:image/png;base64,{image_base64}" alt="Logo" style="width:50px; height:50px; margin-right: 20px;">
            <h1>Overview Tuopu do Brasil</h1>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # Ajustando o layout para os subtítulos
    col_m, col_space, col_v = st.columns([1, 1, 1])

    with col_m:
        st.subheader("Montagem")
    with col_space:
        st.write("")  # Espaço em branco para separar
    with col_v:
        st.subheader("Vulcanização")

    # Ajustando o layout para ter mais colunas
    col1, col2, col3, col4, col5, col6 = st.columns([1, 1, 1, 1, 1, 1])

    # Exibindo os cards M (Montagem)
    with col1:
        for i, val in enumerate(porcentagens_m[:3]):
            create_card(f"M000{i+1}", val)
    with col2:
        for i, val in enumerate(porcentagens_m[3:6]):
            create_card(f"M000{i+4}", val)
    with col3:
        if len(porcentagens_m) > 6:
            create_card(f"M0007", porcentagens_m[6])

    # Exibindo os cards V (Vulcanização)
    with col4:
        for i, val in enumerate(porcentagens_v[:4]):
            create_card(f"V000{i+1}", val)
    with col5:
        for i, val in enumerate(porcentagens_v[4:8]):
            create_card(f"V000{i+5}", val)
    with col6:
        for i, val in enumerate(porcentagens_v[8:]):
            create_card(f"V000{i+9}", val)

    # Exibindo o dataframe lido para depuração
    st.write("Dados lidos do arquivo Excel:")
    st.dataframe(df.head(qtdLinhas))

    # Exibindo o caminho do arquivo para depuração
    #st.write(f"Caminho do arquivo: {caminho_arquivo}")
