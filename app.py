import streamlit as st
import pandas as pd
import folium
import re
import io
import simplekml
from streamlit_folium import st_folium

# Configurações da página
st.set_page_config(page_title="Mapa de Fiscais", layout="wide")

# Função para extrair números reais
def extrair_numeros(texto):
    return [float(n) for n in re.findall(r'-?\d+\.\d+', str(texto))]

# Carrega a planilha
caminho_excel = 'testar.xlsx'
df = pd.read_excel(caminho_excel)

# Lista de ocupantes
ocupantes = sorted(df['ocupante'].dropna().unique())

# Título centralizado
st.markdown("<h1 style='text-align: center;'>📍 Mapa de Apontamentos dos Fiscais</h1>", unsafe_allow_html=True)

# Filtro centralizado
col_filtro = st.columns([1, 2, 1])[1]
with col_filtro:
    ocupante_selecionado = st.selectbox("Selecione o ocupante:", ["Todos"] + ocupantes)

# Dados filtrados
df_filtrado = df if ocupante_selecionado == "Todos" else df[df['ocupante'] == ocupante_selecionado]

# Cria o mapa
mapa = folium.Map(location=[-8.05, -34.88], zoom_start=13)

# Contador do filtro
pontos_filtrados = 0
for _, row in df_filtrado.iterrows():
    latitudes = extrair_numeros(row['LATITUDE'])
    longitudes = extrair_numeros(row['LONGITUDE'])
    pares = min(len(latitudes), len(longitudes))
    pontos_filtrados += pares

    for lat, lon in zip(latitudes, longitudes):
        popup = f"<b>Ocupante:</b> {row['ocupante']}<br><b>Latitude:</b> {lat}<br><b>Longitude:</b> {lon}"
        folium.Marker(
            location=[lat, lon],
            popup=popup,
            icon=folium.Icon(color='green', icon='check')
        ).add_to(mapa)

valor_filtrado = pontos_filtrados * 11.52

# Contador geral
pontos_geral = 0
for _, row in df.iterrows():
    latitudes = extrair_numeros(row['LATITUDE'])
    longitudes = extrair_numeros(row['LONGITUDE'])
    pontos_geral += min(len(latitudes), len(longitudes))

valor_geral = pontos_geral * 11.52

# Layout
col1, col2 = st.columns([1, 3])

with col1:
    st.subheader("📊 Resumo")
    if ocupante_selecionado == "Todos":
        st.metric("Total de pontos", pontos_geral)
        st.metric("Valor total", f"R$ {valor_geral:,.2f}")
    else:
        st.metric("Pontos do ocupante", pontos_filtrados)
        st.metric("Valor do ocupante", f"R$ {valor_filtrado:,.2f}")
        st.markdown("---")
        st.metric("Total acumulado", pontos_geral)
        st.metric("Valor acumulado", f"R$ {valor_geral:,.2f}")

    # Botão para baixar planilha completa
    buffer_excel = io.BytesIO()
    df.to_excel(buffer_excel, index=False, engine='openpyxl')
    buffer_excel.seek(0)

    st.download_button(
        label="📥 Baixar planilha completa",
        data=buffer_excel,
        file_name="apontamentos_completos.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # Botão para baixar dados filtrados em KMZ
    kml = simplekml.Kml()
    for _, row in df_filtrado.iterrows():
        latitudes = extrair_numeros(row['LATITUDE'])
        longitudes = extrair_numeros(row['LONGITUDE'])
        for lat, lon in zip(latitudes, longitudes):
            kml.newpoint(
                name=row['ocupante'],
                coords=[(lon, lat)],
                description=f"Fiscal: {row['ocupante']}\nLat: {lat}\nLon: {lon}"
            )

    buffer_kmz = io.BytesIO()
    kml.savekmz(buffer_kmz)
    buffer_kmz.seek(0)

    st.download_button(
        label="📍 Baixar dados filtrados em KMZ",
        data=buffer_kmz,
        file_name=f"{ocupante_selecionado}_apontamentos.kmz",
        mime="application/vnd.google-earth.kmz"
    )

with col2:
    st.subheader("🗺️ Mapa de apontamentos")
    st_folium(mapa, width=950, height=600)

