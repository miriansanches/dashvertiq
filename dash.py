
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

st.markdown("""
<style>
    * {
        margin: 0;
        padding: 0;
        box-sizing: border-box;
    }

    /* ===== FUNDO PRINCIPAL ===== */
    [data-testid="stAppViewContainer"] {
        background-color: #0d0d0d !important;
        background-image: none !important;
    }

    [data-testid="stSidebar"] {
        background-color: #0d0d0d !important;
        background-image: none !important;
    }

    /* ===== SIDEBAR CONTENT ===== */
    [data-testid="stSidebarContent"] {
        background-color: #0d0d0d !important;
    }

    [data-testid="stSidebar"] label,
    [data-testid="stSidebar"] p,
    [data-testid="stSidebar"] span,
    [data-testid="stSidebar"] div {
        color: #FFD700 !important;
    }

    /* ===== HEADINGS - TUDO AMARELO ===== */
    h1, h2, h3, h4, h5, h6 {
        color: #FFD700 !important;
    }

    /* ===== TEXTO GERAL ===== */
    body, p, span, div {
        color: #E0E0E0 !important;
    }

    hr {
        border-color: rgba(255, 215, 0, 0.3) !important;
    }

    /* ===== METRIC CONTAINER - CARDS PRINCIPAIS (Receita, Forecast, Pace, %, Meta) ===== */
    [data-testid="metric-container"] {
        /* MUDANÇA PRINCIPAL: Fundo cinza chumbo sólido e visível */
        background-color: #2b2b2b !important; 
        
        /* Gradiente sutil para dar acabamento premium */
        background-image: linear-gradient(135deg, #383838 0%, #2b2b2b 100%) !important;
        
        padding: 20px !important;
        border-radius: 12px !important;
        
        /* Borda dourada para definir o limite do card */
        border: 1px solid rgba(255, 215, 0, 0.3) !important;
        
        /* Sombra forte para destacar do fundo preto */
        box-shadow: 0 6px 15px rgba(0, 0, 0, 0.7) !important;
        
        /* Garantir que a cor do texto herde corretamente */
        color: #E0E0E0 !important;
    }

    [data-testid="metric-container"]:hover {
        /* Efeito visual ao passar o mouse */
        border-color: rgba(255, 215, 0, 0.8) !important;
        box-shadow: 0 8px 25px rgba(0, 0, 0, 0.9) !important;
        transform: translateY(-2px) !important;
        background-color: #333333 !important;
    }

    /* Ajuste forçado para o título do card dentro do container */
    [data-testid="metric-container"] > div:nth-child(1) {
        font-size: 14px !important;
        color: #FFD700 !important; /* Dourado */
        font-weight: 600 !important;
    }

    /* Ajuste forçado para o valor do card */
    [data-testid="metric-container"] > div:nth-child(2) {
        color: #FFFFFF !important; /* Branco puro para contraste */
    }

    /* ===== DATAFRAME & TABLE ===== */
    [data-testid="stDataFrame"],
    [data-testid="stTable"] {
        background-color: #1a1a1a !important;
        color: #E0E0E0 !important;
    }

    /* ===== BUTTONS ===== */
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

    /* ===== TABS ===== */
    .stTabs [data-baseweb="tab-list"] {
        gap: 15px !important;
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

    /* ===== INFO CARDS (Missões) ===== */
    .info-card {
        background: linear-gradient(135deg, #0d0d0d 0%, #1a1a1a 100%) !important;
        padding: 20px !important;
        border-radius: 12px !important;
        border: 1px solid rgba(255, 215, 0, 0.5) !important;
        margin-bottom: 20px !important;
        color: #E0E0E0 !important;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.8) !important;
        backdrop-filter: blur(10px) !important;
    }

    .info-card h3 {
        color: #FFD700 !important;
        margin-bottom: 15px !important;
        border-bottom: 1px solid rgba(255, 215, 0, 0.3) !important;
        padding-bottom: 10px !important;
    }

    .info-card ul {
        list-style-type: none !important;
        padding-left: 0 !important;
    }

    .info-card li {
        margin-bottom: 10px !important;
        padding-left: 20px !important;
        position: relative !important;
    }

    /* ===== OBJETIVO CARD ===== */
    .objetivo-card-dark {
        background: linear-gradient(135deg, #1a3050 0%, #0d1f35 100%) !important;
        padding: 25px !important;
        border-radius: 12px !important;
        color: white !important;
        text-align: center !important;
        margin-top: 30px !important;
        border: 1px solid rgba(255, 215, 0, 0.4) !important;
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.8) !important;
        backdrop-filter: blur(10px) !important;
    }

    .objetivo-card-dark h3 {
        color: #FFD700 !important;
        margin-bottom: 15px !important;
    }

    /* ===== PROGRESS BAR ===== */
    .progress-bar {
        background: rgba(255, 255, 255, 0.1) !important;
        height: 8px !important;
        border-radius: 4px !important;
        margin-top: 10px !important;
        overflow: hidden !important;
        border: 1px solid rgba(255, 215, 0, 0.2) !important;
    }

    .progress-fill {
        height: 100% !important;
        background: linear-gradient(90deg, #4a9dd4 0%, #2a7db3 100%) !important;
        border-radius: 4px !important;
        box-shadow: 0 0 10px rgba(74, 157, 212, 0.5) !important;
    }

    /* ===== RANGE CARDS ===== */
    .range-card {
        background: linear-gradient(135deg, #1a3050 0%, #0d1f35 100%) !important;
        padding: 15px !important;
        border-radius: 10px !important;
        color: white !important;
        margin: 5px 0 !important;
        border-left: 4px solid rgba(255, 215, 0, 0.5) !important;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.8) !important;
        backdrop-filter: blur(5px) !important;
    }

    .range-title {
        font-weight: bold !important;
        font-size: 14px !important;
        margin-bottom: 5px !important;
        color: #FFD700 !important;
    }

    .range-classification {
        font-size: 12px !important;
        opacity: 0.9 !important;
        margin-bottom: 8px !important;
        font-weight: 500 !important;
        color: #E0E0E0 !important;
    }

    .range-info {
        display: flex !important;
        justify-content: space-between !important;
        font-size: 12px !important;
        color: #E0E0E0 !important;
    }

    /* ===== CONTAINER PARA CARDS ALINHADOS ===== */
    .metrics-container {
        display: grid !important;
        grid-template-columns: repeat(auto-fit, minmax(280px, 1fr)) !important;
        gap: 20px !important;
        margin-bottom: 30px !important;
        width: 100% !important;
    }

    /* ===== CARDS PRINCIPAIS ===== */
    .metric-card {
        background: linear-gradient(135deg, #0d0d0d 0%, #1a1a1a 100%) !important;
        padding: 25px !important;
        border-radius: 12px !important;
        border: 1px solid rgba(255, 215, 0, 0.5) !important;
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.8) !important;
        color: #E0E0E0 !important;
        backdrop-filter: blur(15px) !important;
        transition: all 0.3s ease !important;
    }

    .metric-card-title {
        font-size: 13px !important;
        color: #FFD700 !important;
        font-weight: 600 !important;
        margin-bottom: 12px !important;
        text-transform: uppercase !important;
        letter-spacing: 0.5px !important;
        opacity: 0.9 !important;
    }

    .metric-card-value {
        font-size: 28px !important;
        font-weight: 700 !important;
        color: #E0E0E0 !important;
        line-height: 1.3 !important;
        margin-bottom: 5px !important;
    }

    .metric-card:hover {
        border-color: rgba(255, 215, 0, 0.8) !important;
        box-shadow: 0 6px 20px rgba(255, 215, 0, 0.2) !important;
        transform: translateY(-2px) !important;
    }

    /* ===== CARDS SECUNDÁRIOS ===== */
    .secondary-metrics-container {
        display: grid !important;
        grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)) !important;
        gap: 20px !important;
        margin-top: 20px !important;
    }

    .secondary-metric-card {
        background: linear-gradient(135deg, #1a3050 0%, #0d1f35 100%) !important;
        padding: 20px !important;
        border-radius: 12px !important;
        border: 1px solid rgba(255, 215, 0, 0.4) !important;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.8) !important;
        backdrop-filter: blur(10px) !important;
        transition: all 0.3s ease !important;
    }

    .secondary-metric-card-title {
        font-size: 12px !important;
        color: #FFD700 !important;
        font-weight: 600 !important;
        margin-bottom: 10px !important;
        text-transform: uppercase !important;
        opacity: 0.9 !important;
    }

    .secondary-metric-card-value {
        font-size: 26px !important;
        font-weight: 700 !important;
        color: #E0E0E0 !important;
    }

    .secondary-metric-card:hover {
        border-color: rgba(255, 215, 0, 0.6) !important;
        box-shadow: 0 5px 15px rgba(255, 215, 0, 0.15) !important;
        transition: all 0.3s ease !important;
    }

    /* ===== RESPONSIVO PARA MOBILE ===== */
    @media (max-width: 768px) {
        [data-testid="metric-container"] > div:nth-child(2) {
            font-size: 20px !important;
        }

        .metric-card-value {
            font-size: 20px !important;
        }

        .secondary-metric-card-value {
            font-size: 18px !important;
        }

        .metrics-container {
            grid-template-columns: 1fr !important;
        }
    }

    /* ===== RESPONSIVO PARA TABLET ===== */
    @media (max-width: 1024px) {
        .metrics-container {
            grid-template-columns: repeat(2, 1fr) !important;
        }
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

# Função que formata os numeros em strings
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
        
        # COLUNAS QUE DEVEM SER FORMATADAS COMO R$
        money_keywords = ['receita', 'vendido', 'objetivo', 'valor', 'parcela', 'realizado', 'meta', 'pipeline', 'whole life','pj 2', 'internacional', 'plano saude', 'câmbio', 'fundos', 'resultado', 'consorcios', 'seguros', 'maxima', 'prem', 'coe', 'pj2', 'rv', 'rf', 'objetivo', 'consórcio', 'total', 'Forecast', 'Volume','Cap Liq', 'necessário para campanha', 'obj. cap. liq', 'obj. cap', 'cap liq', 'consorcio', 'valor venda', 'volume']

        
        # COLUNAS QUE DEVEM FICAR COMO NÚMEROS (NÃO MOEDA)
        number_keywords = ['convertidos', 'reuniões', 'boletas ', 'elegivel' ]
        
        for col in df_filtered.columns:
            col_lower = col.lower().strip()
            
            # Prioridade: Se é número puro, deixa como número
            if any(key in col_lower for key in number_keywords):
                try:
                    df_filtered[col] = pd.to_numeric(
                        df_filtered[col].astype(str).str.replace('R$', '').str.strip(), 
                        errors='coerce'
                    ).fillna(0).astype(int)  # Converte para int para ficar 0, 2, 1 (sem decimais)
                except:
                    pass
            
            # Se é coluna de dinheiro e possivel valor financeir, formata como R$
            elif any(key in col_lower for key in money_keywords):
                try:
                    df_filtered[col] = pd.to_numeric(
                        df_filtered[col].astype(str).str.replace('R$', '').str.replace('-', '0').str.strip(), 
                        errors='coerce'
                    ).fillna(0).apply(format_currency)
                except:
                    pass
        
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

tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs(["Visão Geral", "Consórcios", "Seguros", "Advisor", "Time Comercial", "Comercial - Pipeline", "Captação Liq", 'Campanha - Hot Money'])

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
                
                total_realizado = df_prod['Realizado'].sum()

                # Pega o valor (que internamente é 0.57598050)
                forecast_rows = df_vg[df_vg.iloc[:, 0].astype(str).str.contains('Forecast', case=False, na=False)]
                forecast_val = forecast_rows.iloc[-1, 1] if not forecast_rows.empty else 0

                # Converte para número real (575980.50)
                forecast_val = float(forecast_val)


                # Realizado do período (12/01 a 16/01)
                periodo_row = df_vg[df_vg.iloc[:, 0].astype(str).str.contains('Realizado de 12/01 até 16/01', case=False, na=False)]
                realizado_periodo = periodo_row.iloc[0, 1] if not periodo_row.empty else 0
                
            
                # Média diaria realizada (semana/7)
                media_dia = float(realizado_periodo) / 7 if realizado_periodo else 0
                
                # Média semanal por dias uteis
                meta_dia_row = df_vg[df_vg.iloc[:, 0].astype(str).str.contains('Meta dia útil', case=False, na=False)]
                meta_dia_util = meta_dia_row.iloc[0, 1] if not meta_dia_row.empty else '114%'
                
                # Meta do mês
                meta_rows = df_vg[df_vg.iloc[:, 0].astype(str).str.contains('Meta', case=False, na=False)]
                meta_total = pd.to_numeric(meta_rows.iloc[1, 1], errors='coerce') if len(meta_rows) > 1 else 500000
                
                # Porcetagem da meta total atingida no mês
                percent_meta = (total_realizado / meta_total * 100) if meta_total > 0 else 0

                # Primeiros 4 cards
                # Visão Gera
# PRIMEIRO ROW - 3 cards principais (centralizado)
                m1, m2, m3 = st.columns(3, gap="medium")

                with m1:
                    st.metric("Receita Total Realizada", format_currency(total_realizado))
                    
                with m2:
                    st.metric("Forecast", format_currency(forecast_val))
                    
                with m3:
                    st.metric("Pace", format_currency(meta_dia_util))

                st.markdown("<br>", unsafe_allow_html=True)

                # SEGUNDO ROW - 2 cards (melhor proporção)
                c1, c2 = st.columns(2, gap="medium")

                with c1:
                    st.metric("% da Meta Atingida", f"{percent_meta:.1f}%")
                    
                with c2:
                    st.metric("Meta Total", format_currency(meta_total))

                
                st.markdown("---")
                
                st.subheader("🎯 Concentração de Receita por Produto")
                df_pizza = df_prod[df_prod['Realizado'] > 0]
                
                fig_pie = go.Figure(data=[go.Pie(
                    labels=df_pizza['Produto'],
                    values=df_pizza['Realizado'],
                    marker=dict(colors=["#0008FF", '#FFC600', "#FF001E9A", "#09FF008C", '#FF9600', "#00FF1E9D", '#B8860B', "#2045DA"],
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
                st.subheader("👤 Receita por Assessor")
                
                if 'assessores' in sheets:
                    df_ass_rec = sheets['assessores'].copy()
                    # Encontrar colunas de Assessor e Total
                    col_ass = find_column(df_ass_rec, 'Assessores')
                    col_tot = find_column(df_ass_rec, 'Total')
                    
                    if col_ass and col_tot:
                        # Limpar dados para garantir que são números
                        df_ass_rec[col_tot] = pd.to_numeric(df_ass_rec[col_tot], errors='coerce').fillna(0)
                        
                        # Filtrar quem tem receita > 0 e ordenar do maior para o menor
                        df_ass_rec = df_ass_rec[df_ass_rec[col_tot] > 0].sort_values(by=col_tot, ascending=True)
                        
                        fig_bar = go.Figure(go.Bar(
                            x=df_ass_rec[col_tot],
                            y=df_ass_rec[col_ass],
                            orientation='h',
                            marker=dict(color='#FFD700'), # Cor dourada fixa
                            text=df_ass_rec[col_tot].apply(format_currency), # Valor formatado na barra
                            textposition='auto',
                            hovertemplate='<b>%{y}</b> Receita: %{x:,.2f}<extra></extra>'
                        ))
                        
                        fig_bar.update_layout(
                            paper_bgcolor='rgba(0,0,0,0)',
                            plot_bgcolor='rgba(0,0,0,0)',
                            font=dict(color='#FFD700'),
                            margin=dict(t=20, b=20, l=20, r=20),
                            xaxis=dict(showgrid=False, showticklabels=False),
                            yaxis=dict(showgrid=False),
                            height=max(400, len(df_ass_rec) * 35) # Ajusta altura conforme nº de assessores
                        )
                        st.plotly_chart(fig_bar, use_container_width=True)

                st.markdown("---")
                
                
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
        display_data_table(df, "Pipeline Seguros", ['Whole life', 'Vida', 'Plano Saude', 'Receita Do Mês', 'Receita Acumulada'])
    else:
        st.warning("⚠️ Sheet 'seguros' não encontrada.")

with tab4:
    st.markdown("## 💼 Advisor")
    st.markdown("---")
    st.markdown("---")

    if sheets:
        if 'advisor - geral' in sheets:
            df = sheets['advisor - geral']
            st.subheader("📋 Advisor - Geral")
            df_geral = safe_filter_by_column(df, 'Assessor')
            display_data_table(df_geral, "Advisor Geral", ['Assessor', 'Reuniões realizadas', 'Convertidos', 'Produto', 'Valor Venda'])
            
            st.markdown("---")
            st.subheader("📊 Pipeline Advisor")
            df_pipeline = safe_filter_by_column(df, 'Pipeline')
            display_data_table(df_pipeline, "Pipeline Advisor", ['Pipeline', 'Vendido', 'Receita Atual', 'Objetivo', 'Receita Projetada'])
        
        if 'COE - Ouro' in sheets:
            st.markdown("---")
            st.subheader("🏆 COE - Ouro")
            df_ouro = sheets['COE - Ouro'].dropna(axis=1, how='all').dropna(how='all')
            display_data_table(df_ouro, "COE Ouro", df_ouro.columns.tolist())
        
        if 'PE - Prata' in sheets:
            st.markdown("---")
            st.subheader("🥈 PE - Prata")
            df_prata = sheets['PE - Prata'].dropna(axis=1, how='all').dropna(how='all')
            display_data_table(df_prata, "PE Prata", df_prata.columns.tolist())


    else:
        st.warning("⚠️ Nenhum arquivo carregado.")

    col_m1, col_m2 = st.columns(2)
    
    with col_m1:
        st.markdown("""
        <div class="info-card">
            <h3> Missões 1.0</h3>
            <ul>
                <li><b>Renda Variável:</b> Vol. mín. R$ 250k (Corretagem Bovespa)</li>
                <li><b>Internacional:</b> Remessa mín. USD 30k (Conta Global)</li>
                <li><b>COE:</b> Alocar mín. R$ 100k (Prateleira Janeiro)</li>
                <li style="color: #FFD700;"><b>Premiação:</b> Até R$ 26.000,00 adicionais</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)
        
    with col_m2:
        st.markdown("""
        <div class="info-card">
            <h3> Missão 2.0</h3>
            <ul>
                <li><b>Renda Fixa:</b> Vol. R$ 200k (Prêmio R$ 1k) | R$ 300k (Prêmio R$ 2k)</li>
                <li><b>Fundos Fechados:</b> R$ 300k Balcão / 400k Listados (Prêmio R$ 1k)</li>
                <li><b>Conta PJ:</b> Aporte 350k-500k (Prêmio R$ 1k) | > 600k (Prêmio R$ 2k)</li>
                <li style="color: #FFD700;"><b>Vencimento RF:</b> A partir de Jan/2029</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("---")

    if 'missões' in sheets:
        df = sheets['missões']
        st.subheader("Missões 1.0")
        df_missoes1 = safe_filter_by_column(df, 'Assessor')
        display_data_table(df_missoes1, "Missões", ['Assessor', 'Status', 'Cod Matriz', 'Nome Matriz', 'Núcleo', 'Elegivel RV', 'Elegivel Internacional', 'Elegivel COE', 'Premiação máxima'])
            
    st.markdown("---")

    if 'missões 2.0' in sheets:
        df = sheets['missões 2.0']
        st.subheader("Missões 2.0")
        df_missoes2 = safe_filter_by_column(df, 'Assessor')
        display_data_table(df_missoes2, "Missões 2.0", ['Assessor', 'Status', 'Cod Matriz', 'Nome Matriz', 'Núcleo', 'Elegivel RV', 'Elegivel Fundos', 'Elegivel PJ', 'Prem Max'])
            
    st.markdown("---")

    if 'banco master' in sheets:
        df = sheets['banco master']
        st.subheader("Banco Master")
        df_master = safe_filter_by_column(df, 'Assessores')
        display_data_table(df_master, "Banco Master", ['Assessore', 'Volume FGC', 'Volume Convertido'])

with tab5:

    # 1. Verificar se a aba existe (usando uma busca mais flexível para o nome da aba)
    assessor_sheet_name = next((s for s in sheets.keys() if s.lower().strip() == 'assessores'), None)
    
    if assessor_sheet_name:
        df_assessor_raw = sheets[assessor_sheet_name]
        st.markdown("## 👥 Análise por Assessor")
        st.markdown("---")
        
        # 2. Encontrar a coluna do Assessor (flexível)
        assessor_col = find_column(df_assessor_raw, 'Assessores')
        
        if assessor_col:
            # 3. Criar a lista para o filtro
            lista_assessores = ["Todos"] + sorted([str(x) for x in df_assessor_raw[assessor_col].unique() if x])
            assessor_selecionado = st.selectbox("🔍 Selecione um Assessor para filtrar:", lista_assessores, key="filter_assessor_tab5")
            
            # 4. Filtrar os dados
            if assessor_selecionado != "Todos":
                df_exibir = df_assessor_raw[df_assessor_raw[assessor_col] == assessor_selecionado]
                st.subheader(f"📊 Resultados Detalhados: {assessor_selecionado}")
            else:
                df_exibir = df_assessor_raw
                st.subheader("Visão Geral - Todos os Assessores")

            df_formatado = df_exibir.copy()
            
            # Identifica a coluna de Forecast (ajuste o nome se na planilha for diferente)
                        # Identifica as colunas de Forecast e Pace
            col_forecast = next((c for c in df_formatado.columns if 'forecast' in c.lower()), None)
            col_pace = next((c for c in df_formatado.columns if 'pace' in c.lower()), None)
            
            if col_forecast:
                df_formatado[col_forecast] = pd.to_numeric(
                df_formatado[col_forecast].astype(str).str.replace('R$', '').str.replace('-', '0').str.strip(), 
                    errors='coerce'
                 ).fillna(0).apply(format_currency)

            if col_pace:
                # Converte para numérico e formata como porcentagem (com 1 casa decimal)
                df_formatado[col_pace] = pd.to_numeric(df_formatado[col_pace], errors='coerce')
                df_formatado[col_pace] = df_formatado[col_pace].apply(lambda x: f"{x:.1%}" if pd.notnull(x) else "-")


            # Agora passamos o df_formatado para a sua função
            display_data_table(df_formatado, "Tabela_Assessores", df_formatado.columns.tolist())
            
        else:
            st.warning("⚠️ Não encontramos uma coluna chamada 'Assessor' na aba de dados.")
            st.info("Colunas disponíveis: " + ", ".join(df_assessor_raw.columns))
    else:
        st.error("❌ A aba 'assessor' não foi encontrada no arquivo carregado.")
        st.info(f"Abas disponíveis: {', '.join(sheets.keys())}")

    st.markdown("## 💼 SDR")

    # Primeira tabela - SDR (aba "SDR")
    if 'SDR' in sheets:
        df_sdr_raw = sheets['SDR']
        st.subheader("SDR - Parcial da Semana")
        # Na sua planilha, a coluna na aba 'SDR' chama-se 'SDRS'
        col_name = 'SDRS' if 'SDRS' in df_sdr_raw.columns else 'SDR'
        df_sdr = df_sdr_raw[df_sdr_raw[col_name].notna() & (df_sdr_raw[col_name] != '')].copy()
        cols_to_show = [col_name, 'Agendadas', 'Realizadas', 'Convertidas']
        display_data_table(df_sdr, "SDR_Semanal", cols_to_show)
        
        st.divider()

    # Segunda tabela - SDRS (aba "SDR - Semanal")
    if 'SDR - Semanal' in sheets:
        df_sdrs_raw = sheets['SDR - Semanal']
        st.subheader("SDR - Convertidos no Mês (Janeiro)")
        # Na sua planilha, a coluna na aba 'SDR - Semanal' chama-se 'SDR'
        col_name = 'SDR' if 'SDR' in df_sdrs_raw.columns else 'SDRS'
        df_sdrs = df_sdrs_raw[df_sdrs_raw[col_name].notna() & (df_sdrs_raw[col_name] != '')].copy()
        cols_to_show = [col_name, 'Dez', 'S1', 'S2', 'S3', 'S4']
        display_data_table(df_sdrs, "SDR_Mensal", cols_to_show)



with tab6:
    st.markdown("## 🚀 Assessores - Pipeline")
    st.markdown("---")
    if sheets and 'Pipeline - Assessor' in sheets:
        df_pipe_assessor = sheets['Pipeline - Assessor']
        assessor_col2 = find_column(df_pipe_assessor, 'Assessor')
        if assessor_col2:
            lista_assessores2 = ["Todos"] + sorted(df_pipe_assessor[assessor_col2].unique().tolist())
            assessor_selecionado_pipe = st.selectbox("🔍 Selecione um Assessor para análise detalhada:", lista_assessores2)
            
            if assessor_selecionado_pipe != "Todos":
                df_filtrado2 = df_pipe_assessor[df_pipe_assessor[assessor_col2] == assessor_selecionado_pipe]
                st.subheader(f"📊 Resultados: {assessor_selecionado_pipe}")
                display_data_table(df_filtrado2, "assessor", df_filtrado2.columns.tolist())
            else:
                st.subheader("Visão Geral - Todos os Assessores")
                display_data_table(df_pipe_assessor, "assessor", df_pipe_assessor.columns.tolist())
with tab7:
    st.markdown("## 💰 Captação Líquida")
    st.markdown("---")
    
    if sheets and 'captação liq' in sheets:
        df_captacao = sheets['captação liq'].copy()
        df_captacao.columns = df_captacao.columns.str.strip()
        
        # Remove linhas vazias no início
        df_captacao = df_captacao[df_captacao.iloc[:, 0].notna()].reset_index(drop=True)
        
        try:
            # SEÇÃO 1: RANGES E OBJETIVOS COM CLASSIFICAÇÃO
                        # SEÇÃO 1: RANGES E OBJETIVOS (VALORES FIXOS)
            st.subheader("📊 Ranges de Captação Líquida e Objetivos")
            
            # Definindo os dados fixos para os cards
            ranges_fixos = [
                {"titulo": "0 - 5 MM", "classificacao": "Sales Hunter", "objetivo": 1000000.0},
                {"titulo": "10 - 40 MM", "classificacao": "AAI Pleno", "objetivo": 1500000.0},
                {"titulo": "Acima de 40 MM", "classificacao": "AAI Senior", "objetivo": 2000000.0}
            ]
            
            cols = st.columns(3)
            for idx, item in enumerate(ranges_fixos):
                with cols[idx]:
                    st.markdown(f"""
                    <div class="range-card">
                        <div class="range-title">{item['titulo']}</div>
                        <div class="range-classification">{item['classificacao']}</div>
                        <div class="range-info">
                            <span>Objetivo:</span>
                            <span style="font-weight: bold;">{format_currency(item['objetivo'])}</span>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)

            
            st.markdown("---")
            
            # SEÇÃO 2: FILTRO E TABELA
            st.subheader("📋 Tabela de Captação por Assessor")
            
            # Encontra colunas reais na planilha
            col_assessor = find_column(df_captacao, 'Assessor')
            col_posicao = find_column(df_captacao, 'Posição')
            col_range = find_column(df_captacao, 'Range')
            col_obj = find_column(df_captacao, ['Objetivo Cap Liq'])
            col_capt_liq = find_column(df_captacao, ['Captação Líquida'])
            col_cap_obj = find_column(df_captacao, ['Cap x Objetivo'])
            col_ativacoes = find_column(df_captacao, 'Ativações')
            col_habilitacoes = find_column(df_captacao, 'Habilitações')
            
            colunas_encontradas = [col_assessor, col_obj, col_capt_liq, col_cap_obj, col_ativacoes, col_habilitacoes]
            colunas_encontradas = [c for c in colunas_encontradas if c is not None]
            
            # Remove a linha de totais
            df_display = df_captacao[df_captacao[col_assessor].notna()].copy()
            df_display = df_display[~df_display[col_assessor].astype(str).str.strip().isin(['', 'nan'])].copy()

            # Encontrar a linha "Soma" e parar LÁ (incluindo-a)
            if col_assessor and df_display is not None and len(df_display) > 0:
                soma_indices = df_display[
                df_display[col_assessor].astype(str).str.strip() == 'Soma'
                ].index
    
            if len(soma_indices) > 0:
                soma_idx = soma_indices[0]
                df_display = df_display.loc[:soma_idx].copy()  # Pega até "Soma" (incluindo)

            
            # Remove última linha se for linha de totais
            if len(df_display) > 0 and df_display.iloc[-1][col_assessor] == '':
                df_display = df_display[:-1]
            
            # Filtro de assessores
            assessores_list = df_display[col_assessor].unique().tolist()
            assessores_list = ['Todos'] + [a for a in assessores_list if str(a).strip() != '']
            
            assessor_selecionado = st.selectbox(
                "🔍 Filtrar por Assessor:",
                assessores_list,
                key="assessor_filter_captacao"
            )
            
            # Aplica filtro
            if assessor_selecionado != 'Todos':
                df_display = df_display[df_display[col_assessor] == assessor_selecionado].copy()
            
            # Prepara dataframe para exibição
            df_display_final = df_display[colunas_encontradas].copy()
            
            # Formata colunas
            if col_obj:
                # Remove R$, remove pontos de milhar e troca vírgula por ponto
                df_display_final[col_obj] = df_display_final[col_obj].astype(str).str.replace('R$', '', regex=False).str.strip()
                df_display_final[col_obj] = df_display_final[col_obj].apply(lambda x: x.replace('.', '').replace(',', '.') if ',' in x else x)
                df_display_final[col_obj] = pd.to_numeric(df_display_final[col_obj], errors='coerce').fillna(0).apply(format_currency)

            if col_capt_liq:
                df_display_final[col_capt_liq] = df_display_final[col_capt_liq].astype(str).str.replace('R$', '', regex=False).str.strip()
                df_display_final[col_capt_liq] = df_display_final[col_capt_liq].apply(lambda x: x.replace('.', '').replace(',', '.') if ',' in x else x)
                df_display_final[col_capt_liq] = pd.to_numeric(df_display_final[col_capt_liq], errors='coerce').fillna(0).apply(format_currency)

            
            if col_cap_obj:
                df_display_final[col_cap_obj] = df_display_final[col_cap_obj].apply(
                    lambda x: f"{float(str(x).replace('%', '').strip())*100:.0f}%" if pd.notna(x) and str(x).strip() != '' else "0%"
                )
            
            # Renomeia colunas
            rename_dict = {
                col_obj: 'Obj. Captação',
                col_capt_liq: 'Captação Líquida',
                col_cap_obj: 'Cap. x Obj.',
                col_ativacoes: 'Ativações',
                col_habilitacoes: 'Habilitações'
            }
            df_display_final = df_display_final.rename(columns=rename_dict)
            
            st.dataframe(df_display_final, use_container_width=True, hide_index=True)
            
            st.markdown("---")
            
            # SEÇÃO 3: CARD DE OBJETIVO TOTAL - PUXANDO DA PLANILHA
            st.subheader("🎯 Resumo Geral da Captação")
            
            # Encontra a linha de totais (última linha preenchida)
            df_totais = df_captacao.iloc[-2:].copy()
            
            # Pega valores da linha de totais
            total_row = df_captacao[df_captacao[col_assessor].astype(str).str.strip().isin(['', 'nan', 'NaN'])].iloc[0] if len(df_captacao[df_captacao[col_assessor].astype(str).str.strip().isin(['', 'nan', 'NaN'])]) > 0 else df_captacao.iloc[-1]
            
            # Converte valores
            objetivo_total = float(str(total_row['Objetivo Cap Liq']).replace("R$", "").replace(".", "").replace(",", ".")) if pd.notna(total_row['Objetivo Cap Liq']) else 16000000
            captacao_total = float(str(total_row['Captação Líquida']).replace("R$", "").replace(".", "").replace(",", ".")) if pd.notna(total_row['Captação Líquida']) else 11314555
            percentual_objetivo = float(str(total_row['Cap x Objetivo']).replace("%", "").strip()) if pd.notna(total_row['Cap x Objetivo']) else 71
            
            # Soma ativações e habilitações
            ativacoes_total = pd.to_numeric(df_captacao[col_ativacoes], errors='coerce').sum()
            habilitacoes_total = pd.to_numeric(df_captacao[col_habilitacoes], errors='coerce').sum()
            
            # Cards com cores mais escuras e menos chamativas
            col1, col2, col3 = st.columns([1, 1, 1])
            
            with col1:
                st.markdown(f"""
                <div class="objetivo-card-dark">
                    <div style="font-size: 14px; opacity: 0.95;">Objetivo Total</div>
                    <div style="font-size: 32px; font-weight: bold; margin-top: 10px;">
                        {format_currency(objetivo_total)}
                    </div>
                </div>
                """, unsafe_allow_html=True)
            
            with col2:
                st.markdown(f"""
                <div class="objetivo-card-dark">
                    <div style="font-size: 14px; opacity: 0.95;">Captação Realizada</div>
                    <div style="font-size: 32px; font-weight: bold; margin-top: 10px;">
                        {format_currency(captacao_total)}
                    </div>
                </div>
                """, unsafe_allow_html=True)
            
            with col3:
                st.markdown(f"""
                <div class="objetivo-card-dark">
                    <div style="font-size: 14px; opacity: 0.95;">% do Objetivo</div>
                    <div style="font-size: 32px; font-weight: bold; margin-top: 10px;">
                        {percentual_objetivo:.1f}%
                    </div>
                    <div class="progress-bar">
                        <div class="progress-fill" style="width: {min(percentual_objetivo, 100)}%;"></div>
                    </div>
                </div>
                """, unsafe_allow_html=True)
            
            # Métricas adicionais
            st.markdown("### 📈 Métricas Adicionais")
            col_a, col_b, col_c, col_d = st.columns(4)
            
            with col_a:
                st.metric("Total de Ativações", f"{int(ativacoes_total)}")
            
            with col_b:
                st.metric("Total de Habilitações", f"{int(habilitacoes_total)}")
            
            with col_c:
                assessores_positivos = len(df_captacao[
                    pd.to_numeric(
                        df_captacao[col_capt_liq].astype(str).str.replace('R$', '').str.replace('.', '').str.replace(',', '.'),
                        errors='coerce'
                    ) > 0
                ])
                st.metric("Assessores com Captação Positiva", f"{assessores_positivos}")
            
         
            
            st.markdown("---")
            st.markdown("*Dashboard atualizado dinamicamente a partir da planilha | Vértiq Investimentos*")
            
        except Exception as e:
            st.error(f"❌ Erro ao processar dados de Captação Líquida: {str(e)}")
    else:
        st.warning("⚠️ Aba 'captação liq' não encontrada na planilha!")

# ========== CAMPANHA HOT MONEY ==========
with tab8:
    st.markdown("## 🔥 Campanha Hot Money - Janeiro 2026")
    st.markdown("---")
    
    # CARD EXPLICATIVO DA CAMPANHA (em cima da tabela)
    with st.container(border=True):
        col_info, col_meta = st.columns([2, 1], gap="medium")
        
        with col_info:
            st.markdown("### O que é a Campanha?")
            st.markdown("""
            A "Hot Money" é uma **iniciativa interna** para fecharmos o mês de janeiro com uma **captação líquida mínima de R$ 13,3 milhões** no escritório.
            
            Considerando apenas os assessores ativos, precisamos buscar mais **R$ 4 milhões** para atingir o objetivo, 
            assumindo que os assessores deficitários atinjam ao menos **70% da meta mensal**.
            """)
            
        with col_meta:
            st.markdown("### 🎯 Meta Principal")
            st.metric("Captação Líquida", "R$ 13.300.000", "+R$ 4.000.000")
    
    # CARDS COM ELEGIBILIDADE E PRÊMIO
    card1, card2, card3 = st.columns(3, gap="medium")
    
    with card1:
        st.markdown("### Elegibilidade")
        st.markdown("""
        **Mínimo 70%** da meta mensal de captação líquida
        """)
    
    with card2:
        st.markdown("### Premiação")
        st.markdown("""
        **Almoço especial do time** (local a definir)
        """)
    
    with card3:
        st.markdown("### Participantes")
        st.markdown("""
        Todos os assessores que atingirem **70%+** da meta
        """)
    
    st.markdown("---")
    
    # TABELA DOS DADOS (aqui entra sua tabela original)
    if sheets and 'Hot Money' in sheets:
        df_hot_money = sheets['Hot Money']
        
        display_data_table(
            df_hot_money,
            "Hot Money",
            column_names=['Assessor', 'Posição', 'Obj. Cap. Liq','Meta Campanha (70%)','Cap Liq', 'Necessário para Campanha'],
            max_rows=50
        )
    else:
        st.info("📌 Carregue o arquivo Excel com a sheet 'Campanha - Hot Money'")



st.markdown(
    "<p style='text-align: center; color: #FFD700; font-size: 12px;'>Dashboard Financeiro © 2026 | Vértiq Digital</p>",
    unsafe_allow_html=True)











