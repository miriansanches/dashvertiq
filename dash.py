import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from datetime import datetime
from io import BytesIO
import openpyxl


# Configuração da página
st.set_page_config(
    page_title="Dashboard Financeiro - Vértiq",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS customizado - Refinado
st.markdown("""
    <style>
    * {
        margin: 0;
        padding: 0;
    }
    
    [data-testid="stAppViewContainer"] {
        background-color: #0d0d0d !important;
    }
    
    [data-testid="stSidebar"] {
        background-color: #0d0d0d !important;
    }
    
    [data-testid="stSidebarContent"] {
        background-color: #0d0d0d !important;
    }
    
    [data-testid="stSidebar"] label,
    [data-testid="stSidebar"] p,
    [data-testid="stSidebar"] span,
    [data-testid="stSidebar"] div {
        color: #FFD700 !important;
    }
    
    h1, h2, h3, h4, h5, h6 {
        color: #FFD700 !important;
    }
    
    [data-testid="stDataFrame"] {
        background-color: #1a1a1a !important;
    }
    
    .stMetric {
        background: linear-gradient(135deg, #1a1a1a 0%, #252525 100%);
        padding: 15px;
        border-radius: 12px;
        border: 1px solid rgba(255, 215, 0, 0.3);
        box-shadow: 0 4px 15px rgba(0,0,0,0.5);
    }
    
    /* Estilo Moderno para Botões */
    .stButton > button {
        background: linear-gradient(135deg, #FFD700 0%, #B8860B 100%) !important;
        color: #0d0d0d !important;
        border: none !important;
        padding: 0.6rem 2rem !important;
        border-radius: 50px !important;
        font-weight: 700 !important;
        text-transform: uppercase !important;
        letter-spacing: 1px !important;
        transition: all 0.3s ease !important;
        box-shadow: 0 4px 15px rgba(255, 215, 0, 0.2) !important;
        width: 100% !important;
    }
    
    .stButton > button:hover {
        transform: translateY(-2px) !important;
        box-shadow: 0 6px 20px rgba(255, 215, 0, 0.4) !important;
        color: #000 !important;
    }
    
    .stTabs [data-baseweb="tab-list"] {
        gap: 15px;
        background-color: transparent !important;
    }
    
    .stTabs [data-baseweb="tab"] {
        background-color: #1a1a1a !important;
        color: #FFD700 !important;
        border: 1px solid rgba(255, 215, 0, 0.3) !important;
        border-radius: 8px 8px 0 0 !important;
        padding: 10px 25px !important;
        font-weight: 600 !important;
    }
    
    .stTabs [aria-selected="true"] {
        background-color: #FFD700 !important;
        color: #0d0d0d !important;
        border: 1px solid #FFD700 !important;
    }
    
    body, p, span, div {
        color: #E0E0E0 !important;
    }
    
    hr {
        border-color: rgba(255, 215, 0, 0.3) !important;
    }
    
    /* Estilização de Dataframe */
    [data-testid="stTable"] {
        color: #E0E0E0 !important;
    }
    </style>
    """, unsafe_allow_html=True)

def format_currency(val):
    """Formata valores para o padrão R$ 1.234,56"""
    try:
        if pd.isna(val) or val == '':
            return "R$ 0,00"
        return f"R$ {float(val):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return str(val)

st.markdown("# 📊 Dashboard Financeiro")

st.sidebar.markdown("## 📋 Menu")

uploaded_file = st.sidebar.file_uploader(
    "📁 Carregar arquivo Excel",
    type=["xlsx", "xls"],
    help="Selecione seu arquivo consorcios.xlsx"
)

@st.cache_data
def load_excel(file):
    try:
        xls = pd.ExcelFile(file)
        sheets = {}
        for sheet in xls.sheet_names:
            try:
                df = pd.read_excel(file, sheet_name=sheet)
                df.columns = df.columns.str.strip()
                sheets[sheet] = df
            except Exception as e:
                st.warning(f"⚠️ Erro ao carregar sheet '{sheet}': {str(e)}")
                continue
        return sheets if sheets else None
    except Exception as e:
        st.error(f"❌ Erro ao carregar arquivo: {str(e)}")
        return None

def find_column(df, search_terms):
    if isinstance(search_terms, str):
        search_terms = [search_terms]
    for term in search_terms:
        matching_cols = [c for c in df.columns if str(term).lower() in str(c).lower()]
        if matching_cols:
            return matching_cols[0]
    return None

def safe_filter_by_column(df, column_name):
    try:
        col = find_column(df, column_name)
        if col is None:
            return None
        return df[(df[col].notna()) & (df[col] != '')]
    except Exception as e:
        st.error(f"❌ Erro ao filtrar por {column_name}: {str(e)}")
        return None

def safe_get_columns(df, column_names):
    existing_columns = []
    for col in column_names:
        found_col = find_column(df, col)
        if found_col and found_col not in existing_columns:
            existing_columns.append(found_col)
    return existing_columns

def display_data_table(df, title, column_names, max_rows=15):
    try:
        if df is None or df.empty:
            st.info(f"ℹ️ Nenhum dado disponível para {title}")
            return
        valid_columns = safe_get_columns(df, column_names)
        if not valid_columns:
            st.warning(f"⚠️ Nenhuma das colunas esperadas encontrada em {title}")
            return
        
        df_filtered = df[valid_columns].copy()
        df_filtered = df_filtered.dropna(how='all').head(max_rows)
        
        # Formatar colunas numéricas como moeda se parecerem valores financeiros
        finance_keywords = ['receita', 'vendido', 'objetivo', 'valor', 'parcela', 'realizado', 'meta']
        for col in df_filtered.columns:
            if any(key in col.lower() for key in finance_keywords):
                df_filtered[col] = df_filtered[col].apply(format_currency)
        
        if df_filtered.empty:
            st.info(f"ℹ️ Nenhum dado encontrado para {title}")
            return
        st.dataframe(df_filtered, use_container_width=True, hide_index=True)
    except Exception as e:
        st.error(f"❌ Erro inesperado em {title}: {str(e)}")

sheets = None
if uploaded_file:
    sheets = load_excel(uploaded_file)
    if sheets:
        st.sidebar.success(f"✅ Arquivo carregado com sucesso!")
    else:
        st.sidebar.error("❌ Erro ao carregar arquivo")
else:
    st.sidebar.info("📌 Faça upload do arquivo Excel para começar.")

st.sidebar.markdown("---")

tab1, tab2, tab3, tab4 = st.tabs(["Visão Geral", "Consórcios", "Seguros", "Advisor"])

with tab1:
    st.markdown("## 📈 Visão Geral")
    st.markdown("---")
    
    if sheets and 'visão geral' in sheets:
        df_vg = sheets['visão geral']
        
        try:
            produto_row_idx = df_vg[df_vg.iloc[:, 0] == 'Produto'].index
            if not produto_row_idx.empty:
                idx = produto_row_idx[0]
                df_prod = df_vg.iloc[idx+1:idx+12, 0:3].copy()
                df_prod.columns = ['Produto', 'Realizado', 'Meta']
                df_prod = df_prod.dropna(subset=['Produto'])
                
                for col in ['Realizado', 'Meta']:
                    df_prod[col] = pd.to_numeric(df_prod[col].replace('-', 0).replace('R$ -   ', 0), errors='coerce').fillna(0)
                
                meta_row = df_vg[df_vg.iloc[:, 0] == 'Meta'].index
                meta_total = 0
                realizado_periodo = 0
                if not meta_row.empty:
                    m_idx = meta_row[0]
                    meta_total = pd.to_numeric(df_vg.iloc[m_idx, 1], errors='coerce') or 0
                    realizado_periodo = pd.to_numeric(df_vg.iloc[m_idx+5, 1], errors='coerce') or 0
                
                total_realizado = df_prod['Realizado'].sum()
                
                # MÉTRICAS PRINCIPAIS
                m1, m2, m3, m4 = st.columns(4)
                with m1:
                    st.metric("Receita Total Realizada", format_currency(total_realizado))
                with m2:
                    st.metric("Meta Total do Mês", format_currency(meta_total))
                with m3:
                    percent_meta = (total_realizado / meta_total * 100) if meta_total > 0 else 0
                    st.metric("% da Meta Atingida", f"{percent_meta:.1f}%")
                with m4:
                    st.metric("Realizado no Período", format_currency(realizado_periodo))
                
                st.markdown("---")
                
                # APENAS GRÁFICO DE PIZZA (Conforme solicitado)
                st.subheader("🎯 Concentração de Receita por Produto")
                df_pizza = df_prod[df_prod['Realizado'] > 0]
                
                fig_pie = go.Figure(data=[go.Pie(
                    labels=df_pizza['Produto'],
                    values=df_pizza['Realizado'],
                    marker=dict(colors=['#FFD700', '#FFC600', '#FFB600', '#FFA600', '#FF9600', '#FF8600', '#B8860B', '#DAA520'],
                               line=dict(color='#0d0d0d', width=2)),
                    textposition='inside',
                    textfont=dict(color='#000000', size=12, weight='bold'),
                    hovertemplate='<b>%{label}</b><br>%{value:,.2f}<br>%{percent}<extra></extra>'
                )])
                fig_pie.update_layout(
                    paper_bgcolor='rgba(0,0,0,0)',
                    plot_bgcolor='rgba(0,0,0,0)',
                    font=dict(color='#FFD700', size=12),
                    height=500,
                    showlegend=True,
                    legend=dict(orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5),
                    margin=dict(t=20, b=100, l=20, r=20)
                )
                st.plotly_chart(fig_pie, use_container_width=True)
                
                st.markdown("---")
                st.subheader("📋 Detalhamento de Receitas")
                
                # Formatar tabela para exibição
                df_prod_display = df_prod.copy()
                df_prod_display['Realizado'] = df_prod_display['Realizado'].apply(format_currency)
                df_prod_display['Meta'] = df_prod_display['Meta'].apply(format_currency)
                st.dataframe(df_prod_display, use_container_width=True, hide_index=True)
                
            else:
                st.warning("⚠️ Estrutura da aba 'visão geral' não reconhecida.")
        except Exception as e:
            st.error(f"❌ Erro ao processar aba Visão Geral: {str(e)}")
            
    elif sheets:
        st.info("💡 Aba 'visão geral' não encontrada no arquivo.")
    else:
        st.info("💡 Carregue um arquivo para ver os dados.")

with tab2:
    st.markdown("## 🏢 Consórcios")
    st.markdown("---")
    if sheets and 'consórcios' in sheets:
        df = sheets['consórcios']
        st.subheader("👥 Dados por Assessor")
        df_assessores = safe_filter_by_column(df, 'Assessor')
        display_data_table(df_assessores, "Assessores Consórcios", ['Assessor', 'Reuniões realizadas', 'Convertidos'])
        
        st.markdown("---")
        st.subheader("📊 Pipeline de Vendas")
        df_pipeline = safe_filter_by_column(df, 'Pipeline')
        display_data_table(df_pipeline, "Pipeline Consórcios", ['Pipeline', 'Vendido', 'Receita Atual', 'Objetivo', 'Receita Projetada'])
    else:
        st.warning("⚠️ Sheet 'consórcios' não encontrada.")

with tab3:
    st.markdown("## 🛡️ Seguros")
    st.markdown("---")
    if sheets and 'seguros' in sheets:
        df = sheets['seguros']
        st.subheader("👥 Dados por Assessor")
        df_assessores = safe_filter_by_column(df, 'Assessor')
        display_data_table(df_assessores, "Assessores Seguros", ['Assessor', 'Reuniões realizadas', 'Convertidos'])
        
        st.markdown("---")
        st.subheader("📊 Pipeline de Seguros")
        df_pipeline = safe_filter_by_column(df, 'Pipeline')
        display_data_table(df_pipeline, "Pipeline Seguros", ['Pipeline', 'Vendido', 'Receita Atual', 'Objetivo', 'Receita Projetada'])
    else:
        st.warning("⚠️ Sheet 'seguros' não encontrada.")

with tab4:
    st.markdown("## 💼 Advisor")
    st.markdown("---")
    if sheets:
        if 'advisor - geral' in sheets:
            df = sheets['advisor - geral']
            st.subheader("📋 Advisor - Geral")
            df_geral = safe_filter_by_column(df, 'Assessor')
            display_data_table(df_geral, "Advisor Geral", ['Assessor', 'Reuniões realizadas', 'Convertidos', 'Produto'])
            
            st.markdown("---")
            st.subheader("📊 Pipeline Advisor")
            df_pipeline = safe_filter_by_column(df, 'Pipeline')
            display_data_table(df_pipeline, "Pipeline Advisor", ['Pipeline', 'Vendido', 'Receita Atual', 'Objetivo', 'Receita Projetada'])
        
        if 'COE -Ouro' in sheets:
            st.markdown("---")
            st.subheader("🏆 COE - Ouro")
            df_ouro = sheets['COE -Ouro'].dropna(axis=1, how='all').dropna(how='all')
            display_data_table(df_ouro, "COE Ouro", df_ouro.columns.tolist())
        
        if 'COE - Prata' in sheets:
            st.markdown("---")
            st.subheader("🥈 COE - Prata")
            df_prata = sheets['COE - Prata'].dropna(axis=1, how='all').dropna(how='all')
            display_data_table(df_prata, "COE Prata", df_prata.columns.tolist())
    else:
        st.warning("⚠️ Nenhum arquivo carregado.")

st.markdown("---")
st.markdown(
    "<p style='text-align: center; color: #FFD700; font-size: 12px;'>Dashboard Financeiro © 2026 | Vértiq Digital</p>",
    unsafe_allow_html=True
)
