# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from PIL import Image
import io
from datetime import datetime
import os
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import openpyxl
import unicodedata
from wordcloud import WordCloud

# Função para remover acentos de uma string
def remover_acentos(txt):
    return ''.join(c for c in unicodedata.normalize('NFD', txt) if unicodedata.category(c) != 'Mn')

# Função para tokenizar e contar as palavras (para gerar as frequências)
def tokenize_and_count(series):
    words = []
    for text in series.dropna():
        words.extend(text.split())
    return pd.Series(words).value_counts().to_dict()

# Stopwords personalizadas (acrescentamos "a", "há" e "o")
stopwords = set(WordCloud().stopwords)
stopwords.update([
    "de", "da", "do", "que", "e", "com", "em", "as", "os", "à", "ao", "nas", "nos", "um", "uma",
    "para", "dos", "das", "seu", "sua", "seus", "suas", "ele", "ela", "eles", "elas", "esta", "este",
    "estes", "estas", "isto", "aquilo", "aquele", "aquela", "aqueles", "aquelas", "isso", "aquilo",
    "entre", "sobre", "até", "sem", "com", "contra", "por", "perante", "desde", "trás", "sob",
    "durante", "mediante", "exceto", "salvo", "fora", "após", "bem", "como", "mal", "assim",
    "cada", "qual", "quais", "onde", "quando", "quanto", "quantos", "quantas", "tanto", "tantos", "tantas",
    "nenhum", "nenhuma", "nenhuns", "nenhumas", "todo", "toda", "todos", "todas", "muitos", "muitas",
    "poucos", "poucas", "algum", "alguma", "alguns", "algumas", "outro", "outra", "outros", "outras",
    "mesmo", "mesma", "mesmos", "mesmas", "próprio", "própria", "próprios", "próprias", "tal", "tais",
    "se", "mas", "pois", "porque", "portanto", "logo", "então", "nem", "contudo", "todavia", "entretanto",
    "não", "sim", "ainda", "já", "apenas", "somente", "também", "muito", "pouco", "mais", "menos",
    "quem", "cujo", "cuja", "cujos", "cujas",
    "a", "há", "o"
])

# --------------------------------------------------
# Configuração da Página e Cabeçalho
# --------------------------------------------------
st.set_page_config(
    page_title="Relatório de Vistorias Técnicas - COLOG",
    page_icon="🧊",
    layout="wide"
)

# Função para carregar o logo
def load_colog_logo():
    try:
        logo_path = r"D:\GoogleDrive britointeb\IMAGENS\Logo_Colog_Sem_Fundo.png"
        if os.path.exists(logo_path):
            return Image.open(logo_path)
        else:
            st.sidebar.warning(f"Arquivo do logo não encontrado em: {logo_path}. Usando placeholder.")
            raise FileNotFoundError
    except Exception as e:
        st.sidebar.warning(f"Não foi possível carregar o logo. Usando placeholder. Erro: {e}")
        fig, ax = plt.subplots(figsize=(3, 3))
        ax.set_xlim(0, 10)
        ax.set_ylim(0, 10)
        rect = patches.Rectangle((0, 0), 10, 10, linewidth=1, edgecolor='black', facecolor='#FFDB00')
        ax.add_patch(rect)
        rect = patches.Rectangle((1, 7), 8, 2, linewidth=1, edgecolor=None, facecolor='#FF0000')
        ax.add_patch(rect)
        rect = patches.Rectangle((1, 5.5), 8, 1.5, linewidth=1, edgecolor=None, facecolor='#00AEEF')
        ax.add_patch(rect)
        rect = patches.Rectangle((1, 1), 4, 4.5, linewidth=1, edgecolor=None, facecolor='#FF0000')
        ax.add_patch(rect)
        rect = patches.Rectangle((5, 3.5), 4, 2, linewidth=1, edgecolor=None, facecolor='#FFDB00')
        ax.add_patch(rect)
        rect = patches.Rectangle((5, 1), 4, 2.5, linewidth=1, edgecolor=None, facecolor='#CCCCCC')
        ax.add_patch(rect)
        ax.text(5, 8, "COLOG", fontsize=15, ha='center', va='center', color='white', weight='bold')
        ax.axis('off')
        buf = io.BytesIO()
        fig.savefig(buf, format='png')
        buf.seek(0)
        plt.close(fig)
        return Image.open(buf)

# Função para converter valores para float
def converter_valor_para_numero(valor_str):
    if isinstance(valor_str, (int, float)):
        return float(valor_str)
    if isinstance(valor_str, str):
        try:
            valor_limpo = valor_str.replace("R$", "").strip().replace(".", "").replace(",", ".")
            return float(valor_limpo)
        except ValueError:
            return 0.0
    return 0.0

# Função para formatar valores como moeda
def formatar_moeda(valor):
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

# --------------------------------------------------
# Carregamento dos Arquivos Excel
# --------------------------------------------------
caminho_resumo = "TAB_VT_CAMFRIGO_16OM_RESUMO.xlsx"
try:
    df_resumo = pd.read_excel(caminho_resumo, engine='openpyxl')
    st.sidebar.success(f"Arquivo {caminho_resumo} carregado com sucesso!")
except Exception as e:
    st.error(f"Erro ao carregar {caminho_resumo}: {e}")
    df_resumo = pd.DataFrame()

caminho_servicos = "TAB_VT_CAMFRIGO_16OM_SERVICOS.xlsx"
try:
    df_servicos = pd.read_excel(caminho_servicos, engine='openpyxl')
    st.sidebar.success(f"Arquivo {caminho_servicos} carregado com sucesso!")
except Exception as e:
    st.error(f"Erro ao carregar {caminho_servicos}: {e}")
    df_servicos = pd.DataFrame()

caminho_problemas = "TAB_VT_CAMFRIGO_16OM_PRINCIPAIS_PROBLEMAS.xlsx"
try:
    df_problemas = pd.read_excel(caminho_problemas, engine='openpyxl')
    st.sidebar.success(f"Arquivo {caminho_problemas} carregado com sucesso!")
except Exception as e:
    st.error(f"Erro ao carregar {caminho_problemas}: {e}")
    df_problemas = pd.DataFrame()

# Padronizar os nomes das colunas para todos os arquivos
if not df_resumo.empty:
    df_resumo.columns = [remover_acentos(col.strip().lower().replace(" ", "_")) for col in df_resumo.columns]
    if "valor" not in df_resumo.columns:
        if "valor_estimado" in df_resumo.columns:
            df_resumo = df_resumo.rename(columns={"valor_estimado": "valor"})
        else:
            st.error("Coluna de valor não encontrada no arquivo de resumo.")
if not df_servicos.empty:
    df_servicos.columns = [remover_acentos(col.strip().lower().replace(" ", "_")) for col in df_servicos.columns]
    if "valor" not in df_servicos.columns:
        if "valor_estimado" in df_servicos.columns:
            df_servicos = df_servicos.rename(columns={"valor_estimado": "valor"})
        else:
            st.error("Coluna de valor não encontrada no arquivo de serviços.")
if not df_problemas.empty:
    df_problemas.columns = [remover_acentos(col.strip().lower().replace(" ", "_")) for col in df_problemas.columns]
    # As colunas de interesse para as word clouds serão identificadas via busca de substring

# --------------------------------------------------
# Cabeçalho do Dashboard
# --------------------------------------------------
logo = load_colog_logo()
col1_header, col2_header = st.columns([1, 4])
with col1_header:
    if logo:
        st.image(logo, width=120)
with col2_header:
    st.title("RELATÓRIO DE VISTORIAS TÉCNICAS")
    st.subheader("Câmaras Frias de OM da Guarnição de Brasília")
    try:
        import locale
        try:
            locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')
        except locale.Error:
            st.warning("Locale 'pt_BR.UTF-8' não disponível.")
            data_atual_str = datetime.now().strftime("%d/%m/%Y")
        if 'data_atual_str' not in locals():
            data_atual_str = datetime.now().strftime("%d de %B de %Y")
    except ImportError:
        data_atual_str = datetime.now().strftime("%d/%m/%Y")
    st.caption(f"Data do Relatório: {data_atual_str}")

# Função de destaque para a coluna ESTADO GERAL
def highlight_estado(val):
    if val == "Precário":
        return 'background-color: red; color: white'
    elif val == "Bom":
        return 'background-color: green; color: white'
    elif val == "Ruim":
        return 'background-color: yellow; color: black'
    return ''

# ==================================================
# SEÇÃO 1 – DASHBOARD RESUMO (TAB_VT_CAMFRIGO_16OM_RESUMO)
# ==================================================
st.header("Resumo e Análise – 16 OMs (Resumo)")

if df_resumo.empty:
    st.info("Arquivo de resumo não contém dados.")
else:
    df_resumo["valor"] = df_resumo["valor"].apply(converter_valor_para_numero)
    
    total_geral = df_resumo["valor"].sum()
    media_geral = df_resumo["valor"].mean()
    min_geral = df_resumo["valor"].min()
    max_geral = df_resumo["valor"].max()

    if "solucao_proposta" in df_resumo.columns:
        resumo_solucao = df_resumo.groupby("solucao_proposta").agg(
            total=("valor", "sum"),
            media=("valor", "mean"),
            minimo=("valor", "min"),
            maximo=("valor", "max"),
            quantidade=("valor", "count")
        ).reset_index()
        resumo_solucao["total"] = resumo_solucao["total"].apply(formatar_moeda)
        resumo_solucao["media"] = resumo_solucao["media"].apply(formatar_moeda)
        resumo_solucao["minimo"] = resumo_solucao["minimo"].apply(lambda x: f"{formatar_moeda(x)} ⬇️")
        resumo_solucao["maximo"] = resumo_solucao["maximo"].apply(lambda x: f"{formatar_moeda(x)} ⬆️")
    else:
        st.error("Coluna 'solucao_proposta' não encontrada no arquivo de resumo.")
        resumo_solucao = pd.DataFrame()
    
    st.subheader("Indicadores Gerais")
    col1_ind, col2_ind, col3_ind, col4_ind = st.columns(4)
    col1_ind.metric("Total Geral", formatar_moeda(total_geral))
    col2_ind.metric("Média Geral", formatar_moeda(media_geral))
    col3_ind.metric("Mínimo Geral", formatar_moeda(min_geral))
    col4_ind.metric("Máximo Geral", formatar_moeda(max_geral))
    
    st.subheader("Indicadores por Solução")
    st.dataframe(resumo_solucao.rename(columns={
        "solucao_proposta": "Solução",
        "total": "Total",
        "media": "Média",
        "minimo": "Mínimo",
        "maximo": "Máximo",
        "quantidade": "Quantidade"
    }))
    
    st.subheader("Tabela Detalhada dos Dados (Resumo)")
    df_display = df_resumo.copy().rename(columns={
        "om": "OM",
        "estado_geral": "ESTADO GERAL",
        "solucao_proposta": "SOLUÇÃO PROPOSTA",
        "valor": "VALOR ESTIMADO",
        "nr_opus": "NR OPUS"
    })
    st.dataframe(
        df_display.style.applymap(highlight_estado, subset=['ESTADO GERAL']),
        use_container_width=True
    )
    
    col1_pizza, col2_pizza = st.columns(2)
    with col1_pizza:
        st.subheader("Quantidade por OM")
        contagem_om = df_resumo["om"].value_counts().reset_index()
        contagem_om.columns = ["om", "quantidade"]
        if not contagem_om.empty:
            fig_qtd_om = px.pie(contagem_om, values="quantidade", names="om",
                                title="Distribuição de Quantidade por OM", hole=0.3)
            fig_qtd_om.update_traces(textinfo='percent+label', pull=[0.05]*len(contagem_om))
            st.plotly_chart(fig_qtd_om, use_container_width=True)
    with col2_pizza:
        st.subheader("Quantidade por Solução")
        contagem_sol = df_resumo["solucao_proposta"].value_counts().reset_index()
        contagem_sol.columns = ["solucao", "quantidade"]
        if not contagem_sol.empty:
            fig_qtd_sol = px.pie(contagem_sol, values="quantidade", names="solucao",
                                 title="Distribuição de Quantidade por Solução", hole=0.3)
            fig_qtd_sol.update_traces(textinfo='percent+label', pull=[0.05]*len(contagem_sol))
            st.plotly_chart(fig_qtd_sol, use_container_width=True)
    
    col3_pizza, col4_pizza = st.columns(2)
    with col3_pizza:
        st.subheader("Valor por OM")
        valor_om = df_resumo.groupby("om")["valor"].sum().reset_index()
        if not valor_om.empty:
            fig_val_om = px.pie(valor_om, values="valor", names="om",
                                title="Distribuição de Valor por OM", hole=0.3)
            fig_val_om.update_traces(textinfo='percent+label', pull=[0.05]*len(valor_om))
            st.plotly_chart(fig_val_om, use_container_width=True)
    with col4_pizza:
        st.subheader("Valor por Solução")
        valor_sol = df_resumo.groupby("solucao_proposta")["valor"].sum().reset_index()
        if not valor_sol.empty:
            fig_val_sol = px.pie(valor_sol, values="valor", names="solucao_proposta",
                                 title="Distribuição de Valor por Solução", hole=0.3)
            fig_val_sol.update_traces(textinfo='percent+label', pull=[0.05]*len(valor_sol))
            st.plotly_chart(fig_val_sol, use_container_width=True)
    
    st.subheader("Valor por OM (por Solução)")
    df_om_val = df_resumo.groupby(["om", "solucao_proposta"])["valor"].sum().reset_index()
    if not df_om_val.empty:
        df_om_val = df_om_val.sort_values("valor", ascending=True)
        # Gerar dicionário de cores com base nas soluções únicas
        unique_sols = sorted(df_om_val["solucao_proposta"].unique())
        cores_sol = {sol: px.colors.qualitative.Plotly[i % len(px.colors.qualitative.Plotly)]
                     for i, sol in enumerate(unique_sols)}
        df_om_val["cor"] = df_om_val["solucao_proposta"].apply(lambda s: cores_sol.get(s, "#808080"))
        fig_bar = go.Figure()
        fig_bar.add_trace(go.Bar(
            y=df_om_val["om"],
            x=df_om_val["valor"],
            orientation='h',
            marker_color=df_om_val["cor"],
            text=df_om_val["valor"].apply(formatar_moeda),
            textposition="auto",
            hovertemplate="OM: %{y}<br>Valor: %{x:.2f}<extra></extra>"
        ))
        fig_bar.add_vline(x=media_geral, line_dash="dash", line_color="red",
                          annotation_text=f"Média Geral: {formatar_moeda(media_geral)}",
                          annotation_position="bottom right")
        fig_bar.update_layout(
            xaxis_title="Valor (R$)",
            yaxis_title="OM",
            height=max(400, len(df_om_val)*30)
        )
        st.plotly_chart(fig_bar, use_container_width=True)
    else:
        st.info("Não há dados para o gráfico de barras.")

# ==================================================
# SEÇÃO 2 – DASHBOARD SERVIÇOS (TAB_VT_CAMFRIGO_16OM_SERVICOS)
# ==================================================
st.header("Análise de Serviços – 16 OMs (Serviços)")

if df_servicos.empty:
    st.info("Arquivo de serviços não contém dados.")
else:
    # Excluir disciplinas "total" e "total com bdi"
    df_servicos_filtered = df_servicos[~df_servicos["disciplina"].str.strip().str.lower().isin(["total", "total com bdi"])]
    df_servicos_filtered["valor"] = df_servicos_filtered["valor"].apply(converter_valor_para_numero)
    
    st.subheader("Indicadores por Disciplina")
    total_geral_servicos = df_servicos_filtered["valor"].sum()
    disciplinas = df_servicos_filtered["disciplina"].unique()
    for disc in disciplinas:
        st.markdown(f"**{disc}**")
        df_temp = df_servicos_filtered[df_servicos_filtered["disciplina"] == disc]
        total_val = df_temp["valor"].sum()
        perc = (total_val / total_geral_servicos * 100) if total_geral_servicos > 0 else 0
        min_val = df_temp["valor"].min()
        max_val = df_temp["valor"].max()
        if not df_temp.empty:
            idx_min = df_temp["valor"].idxmin()
            idx_max = df_temp["valor"].idxmax()
            om_min = df_temp.loc[idx_min, "om"]
            om_max = df_temp.loc[idx_max, "om"]
        else:
            om_min = ""
            om_max = ""
        colA, colB, colC = st.columns(3)
        colA.metric("Total", f"{formatar_moeda(total_val)} ({perc:.1f}%)")
        colB.metric(f"Mínimo ({om_min})", f"{formatar_moeda(min_val)} ⬇️")
        colC.metric(f"Máximo ({om_max})", f"{formatar_moeda(max_val)} ⬆️")
        st.markdown("---")
    
    st.subheader("Distribuição de Valor por Disciplina")
    valor_disc = df_servicos_filtered.groupby("disciplina")["valor"].sum().reset_index()
    if not valor_disc.empty:
        fig_pizza_serv = px.pie(valor_disc, values="valor", names="disciplina",
                                title="Valor Total por Disciplina", hole=0.3)
        fig_pizza_serv.update_traces(textinfo='percent+label', pull=[0.05]*len(valor_disc))
        st.plotly_chart(fig_pizza_serv, use_container_width=True)
    
    st.subheader("Valor por OM (Segmentado por Disciplina)")
    om_disc = df_servicos_filtered.groupby(["om", "disciplina"])["valor"].sum().reset_index()
    if not om_disc.empty:
        fig_bar_serv = px.bar(om_disc, x="valor", y="om", color="disciplina",
                              orientation="h",
                              title="Valor por OM Segmentado por Disciplina",
                              labels={"valor": "Valor (R$)", "om": "OM", "disciplina": "Disciplina"},
                              barmode="stack")
        fig_bar_serv.update_layout(xaxis_tickformat=',.2f')
        st.plotly_chart(fig_bar_serv, use_container_width=True)
    else:
        st.info("Não há dados para o gráfico de serviços.")

# ==================================================
# NOVA SEÇÃO – ANÁLISE DOS PRINCIPAIS PROBLEMAS
# ==================================================
st.header("Análise dos Principais Problemas")

if df_problemas.empty:
    st.info("Arquivo de principais problemas não contém dados.")
else:
    # Filtro por OM utilizando a coluna "om"
    if "om" in df_problemas.columns:
        oms_problemas = sorted(df_problemas["om"].unique())
        selected_oms_problemas = st.multiselect("Filtrar Problemas por OM", options=oms_problemas, default=oms_problemas)
        df_problemas_filtrado = df_problemas[df_problemas["om"].isin(selected_oms_problemas)]
    else:
        df_problemas_filtrado = df_problemas

    st.subheader("Tabela dos Principais Problemas")
    st.dataframe(df_problemas_filtrado)
    
    # Word Cloud para Frequência de Defeitos usando a coluna que contenha "problema" ou "defeito"
    st.subheader("Word Cloud - Frequência de Defeitos")
    col_def = None
    for col in df_problemas_filtrado.columns:
        if "problema" in col.lower() or "defeito" in col.lower():
            col_def = col
            break
    if col_def:
        freq_defeitos = tokenize_and_count(df_problemas_filtrado[col_def])
        if freq_defeitos:
            wc_defeitos = WordCloud(width=800, height=400, background_color='white', stopwords=stopwords)\
                .generate_from_frequencies(freq_defeitos)
            st.image(wc_defeitos.to_array(), use_column_width=True)
        else:
            st.info("Não há dados para gerar a word cloud de defeitos.")
    else:
        st.info("Nenhuma coluna que contenha 'problema' ou 'defeito' encontrada.")
    
    # Word Cloud para Distribuição de Soluções usando a coluna que contenha "solucao" ou "solucoes"
    st.subheader("Word Cloud - Distribuição de Soluções")
    col_sol = None
    for col in df_problemas_filtrado.columns:
        if "solucao" in col.lower() or "solucoes" in col.lower():
            col_sol = col
            break
    if col_sol:
        freq_solucoes = tokenize_and_count(df_problemas_filtrado[col_sol])
        if freq_solucoes:
            wc_solucoes = WordCloud(width=800, height=400, background_color='white', stopwords=stopwords)\
                .generate_from_frequencies(freq_solucoes)
            st.image(wc_solucoes.to_array(), use_column_width=True)
        else:
            st.info("Não há dados para gerar a word cloud de soluções.")
    else:
        st.info("Nenhuma coluna que contenha 'solucao' encontrada.")

# ==================================================
# Rodapé e Instruções Finais
# ==================================================
st.markdown('---')
st.markdown("<p style='text-align: center; font-size: 12px;'>PRODUZIDO POR: TC BRITO</p>", unsafe_allow_html=True)
st.markdown(f"<p style='text-align: center; font-size: 10px;'>Relatório gerado em {data_atual_str} com base nas vistorias técnicas.</p>", unsafe_allow_html=True)

st.sidebar.markdown('---')
st.sidebar.markdown("""
#### Instruções de Uso
- Este dashboard integra dados dos arquivos de resumo, serviços e principais problemas.
- Na seção “Resumo e Análise – 16 OMs (Resumo)” são exibidos indicadores gerais e por solução, uma tabela detalhada (com nomes originais e coloração do ESTADO GERAL) e gráficos de distribuição (pizza e barras).
- Na seção “Análise de Serviços – 16 OMs (Serviços)” são apresentados indicadores e gráficos relacionados aos serviços por disciplina. Nos indicadores por disciplina, o Total exibe o valor e sua porcentagem do total geral.
- Na seção “Análise dos Principais Problemas”, utilize o filtro por OM para visualizar a tabela dos problemas e as word clouds (a primeira word cloud utiliza uma coluna que contenha 'problema' ou 'defeito', e a segunda uma que contenha 'solucao').
- Todos os arquivos devem estar na mesma pasta do script.
""")
st.sidebar.markdown(f"""
#### Informações sobre os Dados
- Arquivo de Resumo: **{caminho_resumo}**
- Arquivo de Serviços: **{caminho_servicos}**
- Arquivo de Principais Problemas: **{caminho_problemas}**
- Última atualização: {data_atual_str}
""")
