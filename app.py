import streamlit as st
import pandas as pd
from io import BytesIO
from streamlit_option_menu import option_menu

# Configuração da página do Streamlit
st.set_page_config(page_title="Análise de Planilha",
                   page_icon=":bar_chart:", layout="wide")

# Função para carregar a planilha XLSX


@st.cache_data
def load_excel(file_path):
    return pd.read_excel(file_path)

# Função para converter DataFrame em Excel e obter bytes


def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    processed_data = output.getvalue()
    return processed_data


# Carregar a planilha
df = load_excel('Bancão.xlsx')

# Converter colunas de data
df['DATA_MANIFESTACAO'] = pd.to_datetime(df['DATA_MANIFESTACAO'])
df['PRAZO_CONCLUSAO_PONTO_DE_RESPOSTA'] = pd.to_datetime(
    df['PRAZO_CONCLUSAO_PONTO_DE_RESPOSTA'])
df['PRAZO_DE_CONCLUSÃO_OUVIDORIA'] = pd.to_datetime(
    df['PRAZO_DE_CONCLUSÃO_OUVIDORIA'])
df['ÚLTIMO_ACOMPANHAMENTO_PONTO_DE_RESPOSTA'] = pd.to_datetime(
    df['ÚLTIMO_ACOMPANHAMENTO_PONTO_DE_RESPOSTA'])

# Menu de navegação
with st.sidebar:
    selected = option_menu(
        menu_title="Menu Principal",  # Nome do menu
        options=["Relatório", "Dados Estatísticos"],  # Páginas
        icons=["file-earmark-spreadsheet", "bar-chart"],  # Ícones
        menu_icon="cast",  # Ícone do menu
        default_index=0,  # Página padrão
    )

# Lista das colunas para filtrar
filter_columns = [
    'OUVIDORIA_ORIGEM', 'MUNICIPIO_OUVIDORIA_ORIGEM', 'UF_OUVIDORIA_ORIGEM',
    'ESFERA_OUVIDORIA_ORIGEM', 'NOME_OUVIDORIA_DESTINO', 'MUNICIPIO_OUVIDORIA_DESTINO',
    'UF_OUVIDORIA_DESTINO', 'ESFERA_DESTINO', 'NIVEL_OUVIDORIA_DESTINO', 'PONTO_DE_RESPOSTA',
    'SITUACAO_ACOMPANHAMENTO_DESTINO', 'STATUS_MANIFESTACAO', 'PRAZO_VENCIDO'
]

# Dicionário para armazenar os filtros selecionados
filters = {}

# Slicer para DATA_MANIFESTACAO
st.sidebar.subheader("Filtrar por DATA_MANIFESTACAO")
start_date = st.sidebar.date_input(
    "Data de Início", df['DATA_MANIFESTACAO'].min().date()
)
end_date = st.sidebar.date_input(
    "Data de Fim", df['DATA_MANIFESTACAO'].max().date()
)
filters['DATA_MANIFESTACAO'] = [start_date, end_date]

# Adicionar filtros na barra lateral
st.sidebar.subheader("Filtros")
for column in filter_columns:
    unique_values = df[column].unique()
    selected_values = st.sidebar.multiselect(
        f"Filtrar por {column}", unique_values)
    filters[column] = selected_values

# Aplicar filtros
filtered_df = df.copy()
for column, selected_values in filters.items():
    if selected_values:
        if column == 'DATA_MANIFESTACAO':
            filtered_df = filtered_df[
                (filtered_df[column] >= pd.to_datetime(selected_values[0])) &
                (filtered_df[column] <= pd.to_datetime(selected_values[1]))
            ]
        else:
            filtered_df = filtered_df[filtered_df[column].isin(
                selected_values)]

# Página Relatório
if selected == "Relatório":
    # Título do aplicativo
    st.title("Extração de Relatórios - Sistema Ouvidor SUS")

    total_linhas = len(filtered_df)
    st.write(f"Total de Manifestações: {total_linhas}")

    # Exibir os dados filtrados
    st.header("Dados da Planilha")
    filtered_df['DATA_MANIFESTACAO'] = filtered_df['DATA_MANIFESTACAO'].dt.strftime(
        '%d/%m/%Y')
    filtered_df['PRAZO_CONCLUSAO_PONTO_DE_RESPOSTA'] = filtered_df['PRAZO_CONCLUSAO_PONTO_DE_RESPOSTA'].dt.strftime(
        '%d/%m/%Y')
    filtered_df['PRAZO_DE_CONCLUSÃO_OUVIDORIA'] = filtered_df['PRAZO_DE_CONCLUSÃO_OUVIDORIA'].dt.strftime(
        '%d/%m/%Y')
    filtered_df['ÚLTIMO_ACOMPANHAMENTO_PONTO_DE_RESPOSTA'] = filtered_df['ÚLTIMO_ACOMPANHAMENTO_PONTO_DE_RESPOSTA'].dt.strftime(
        '%d/%m/%Y')

    st.write(filtered_df)

    # Permitir ao usuário selecionar colunas para visualização
    st.subheader("Monte Seu Relatório")
    selected_columns = []
    for column in filtered_df.columns:
        if st.checkbox(column):
            selected_columns.append(column)

    if selected_columns:
        selected_data = filtered_df[selected_columns]
        st.write(selected_data)

        # Botão para baixar os dados selecionados
        st.subheader("Baixar Dados Selecionados")
        if st.button("Clique aqui para baixar seu relatório"):
            excel_data = to_excel(selected_data)
            st.download_button(
                label="Baixar Dados em XLSX",
                data=excel_data,
                file_name="dados_selecionados.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

# Página Análise Descritiva
elif selected == "Dados Estatísticos":
    st.title("Dados Estatísticos - Sistema Ouvidor SUS")

  # Total de linhas
    total_linhas = len(filtered_df)
    st.markdown(f"<h3 style='text-align: center;'>Total de Manifestações: <strong>{
                total_linhas}</strong></h3>", unsafe_allow_html=True)

    # Top 10 Ouvidorias
    top_10_ouvidorias = filtered_df['NOME_OUVIDORIA_DESTINO'].value_counts().head(
        10).reset_index()
    top_10_ouvidorias.columns = ['NOME_OUVIDORIA_DESTINO', 'COUNT']
    # Adiciona uma coluna de índice começando de 1
    top_10_ouvidorias.index += 1
    top_10_ouvidorias.index.name = 'Ranking'

# Exibir o DataFrame com um estilo que ajusta as larguras das colunas
    st.write("Top 10 Ouvidorias (NOME_OUVIDORIA_DESTINO):")
    st.dataframe(top_10_ouvidorias.style.set_table_attributes(
        'style="width:100%;"'))

    # Total de manifestações por ESFERA_DESTINO
    manifestacoes_por_esfera = filtered_df['ESFERA_DESTINO'].value_counts()
    st.write("Total de Manifestações por ESFERA_DESTINO:")
    st.write(manifestacoes_por_esfera)

    # Total de manifestações por CANAL_DE_ENTRADA
    manifestacoes_por_esfera = filtered_df['CANAL_DE_ENTRADA'].value_counts()
    st.write("Total de Manifestações por CANAL_DE_ENTRADA:")
    st.write(manifestacoes_por_esfera)
