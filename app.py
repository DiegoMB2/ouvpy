import streamlit as st
import pandas as pd
from io import BytesIO


# Função para carregar a planilha XLSX
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
df['PRAZO_DE_CONCLUSÃO_OUVIDORIA'] = pd.to_datetime(df['PRAZO_DE_CONCLUSÃO_OUVIDORIA'])
df['ÚLTIMO_ACOMPANHAMENTO_PONTO_DE_RESPOSTA'] = pd.to_datetime(df['ÚLTIMO_ACOMPANHAMENTO_PONTO_DE_RESPOSTA'])

# Lista das colunas para filtrar
filter_columns = [
    'OUVIDORIA_ORIGEM', 'NOME_OUVIDORIA_DESTINO', 'PONTO_DE_RESPOSTA',
    'SITUACAO_ACOMPANHAMENTO_DESTINO', 'STATUS_MANIFESTACAO'
]

# Dicionário para armazenar os filtros selecionados
filters = {}

# Slicer para DATA_MANIFESTACAO
start_date = st.date_input("Data Inicial", df['DATA_MANIFESTACAO'].min().date())
end_date = st.date_input("Data Final", df['DATA_MANIFESTACAO'].max().date())
filters['DATA_MANIFESTACAO'] = [start_date, end_date]

# Adicionar filtros
for column in filter_columns:
    unique_values = df[column].unique()
    selected_values = st.multiselect(f"Selecione {column}", unique_values, default=list(unique_values))
    filters[column] = selected_values

# Aplicar filtros
filtered_df = df.copy()
for column, selected_values in filters.items():
    if len(selected_values) > 0:
        if column == 'DATA_MANIFESTACAO':
            filtered_df = filtered_df[
                (filtered_df[column] >= pd.to_datetime(selected_values[0])) &
                (filtered_df[column] <= pd.to_datetime(selected_values[1]))
            ]
        else:
            filtered_df = filtered_df[filtered_df[column].isin(selected_values)]

# Total de Manifestações
total_linhas = len(filtered_df)
st.markdown(f"<h3 style='text-align: center;'>Total de Manifestações: <strong>{total_linhas}</strong></h3>", unsafe_allow_html=True)

# Dados da Planilha
filtered_df['DATA_MANIFESTACAO'] = filtered_df['DATA_MANIFESTACAO'].dt.strftime('%d/%m/%Y')
filtered_df['PRAZO_CONCLUSAO_PONTO_DE_RESPOSTA'] = filtered_df['PRAZO_DE_CONCLUSÃO_OUVIDORIA'].dt.strftime('%d/%m/%Y')
filtered_df['ÚLTIMO_ACOMPANHAMENTO_PONTO_DE_RESPOSTA'] = filtered_df['ÚLTIMO_ACOMPANHAMENTO_PONTO_DE_RESPOSTA'].dt.strftime('%d/%m/%Y')

# Exibir os dados filtrados
st.dataframe(filtered_df)

# Monte Seu Relatório
selected_columns = st.multiselect("Selecione as colunas para o relatório", list(filtered_df.columns), default=list(filtered_df.columns))
selected_data = filtered_df[selected_columns]
st.dataframe(selected_data)

# Baixar Dados Selecionados
excel_data = to_excel(selected_data)
st.download_button(label="Baixar Dados Selecionados", data=excel_data, file_name="dados_selecionados.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# Análise Descritiva
st.markdown("### Dados Estatísticos - Sistema Ouvidor SUS")

# Top 10 Ouvidorias
top_10_ouvidorias = filtered_df['NOME_OUVIDORIA_DESTINO'].value_counts().head(10).reset_index()
top_10_ouvidorias.columns = ['NOME_OUVIDORIA_DESTINO', 'COUNT']
top_10_ouvidorias.index += 1
top_10_ouvidorias.index.name = 'Ranking'
st.write("Top 10 Ouvidorias (NOME_OUVIDORIA_DESTINO):")
st.dataframe(top_10_ouvidorias)


# Total de manifestações por CANAL_DE_ENTRADA
manifestacoes_por_canal = filtered_df['CANAL_DE_ENTRADA'].value_counts()
st.write("Total de Manifestações por CANAL_DE_ENTRADA:")
st.dataframe(manifestacoes_por_canal)

