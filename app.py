import streamlit as st
from supabase import create_client
import pandas as pd
from PIL import Image
import io
from datetime import date
import time
import plotly.express as px
import requests

# ==========================================
# 1. CONFIGURAÇÃO DA PÁGINA E DESIGN SYSTEM (UX/UI PROFISSIONAL)
# ==========================================
st.set_page_config(
    page_title="Financeiro COMUNA", 
    page_icon="⛪", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilização CSS Clean, Moderna e Sofisticada (Fundo Claro, Estilo Fintech)
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap');

    html, body, [class*="css"] {
        font-family: 'Plus Jakarta Sans', sans-serif;
    }

    /* Fundo geral da aplicação limpo e iluminado */
    .main {
        background-color: #F8FAFC;
        color: #1E293B;
        padding: 2rem 2rem;
    }
    
    /* Sidebar moderna e elegante */
    [data-testid="stSidebar"] {
        background-color: #FFFFFF;
        border-right: 1px solid #E2E8F0;
    }
    
    [data-testid="stSidebar"] .stRadio label {
        font-size: 0.95rem !important;
        font-weight: 500 !important;
        padding: 12px 16px !important;
        border-radius: 10px;
        transition: all 0.2s ease;
        color: #475569 !important;
        margin-bottom: 4px;
    }
    
    [data-testid="stSidebar"] .stRadio label:hover {
        background-color: #F1F5F9;
        color: #0F172A !important;
    }

    /* Cards de Métricas (KPIs) com visual clean */
    div[data-testid="stMetric"] {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        padding: 22px;
        border-radius: 14px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.02), 0 2px 4px -1px rgba(0, 0, 0, 0.01);
    }
    div[data-testid="stMetricValue"] {
        color: #059669;
        font-size: 2rem !important;
        font-weight: 700;
    }
    div[data-testid="stMetricLabel"] {
        color: #64748B !important;
        font-size: 0.9rem !important;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.05em;
    }

    /* Formulários e Containers Brancos */
    div[data-testid="stForm"] {
        background-color: #FFFFFF;
        padding: 35px;
        border-radius: 16px;
        border: 1px solid #E2E8F0;
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.03);
    }

    /* Inputs de formulário altamente refinados */
    .stTextInput input, .stNumberInput input, .stSelectbox select, .stDateInput input {
        background-color: #F8FAFC !important;
        color: #0F172A !important;
        border: 1px solid #CBD5E1 !important;
        border-radius: 10px !important;
        padding: 12px 16px !important;
        font-size: 0.95rem !important;
    }
    
    .stTextInput input:focus, .stNumberInput input:focus {
        border-color: #2563EB !important;
        background-color: #FFFFFF !important;
        box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.15);
    }

    /* Botões de Ação Principais (Grandes, Ergonômicos e Elegantes) */
    .stButton button, div[data-testid="stFormSubmitButton"] button {
        background-color: #2563EB !important;
        color: white !important;
        font-weight: 600 !important;
        font-size: 1rem !important;
        padding: 0.85rem 1.5rem !important;
        border-radius: 10px !important;
        border: none !important;
        width: 100%;
        box-shadow: 0 4px 12px rgba(37, 99, 235, 0.2);
        transition: all 0.2s ease;
    }
    .stButton button:hover, div[data-testid="stFormSubmitButton"] button:hover {
        background-color: #1D4ED8 !important;
        box-shadow: 0 6px 16px rgba(37, 99, 235, 0.3);
    }

    /* Cabeçalhos claros e tipografia profissional */
    h1, h2, h3 {
        color: #0F172A !important;
        font-weight: 700 !important;
    }
    
    p, span, label {
        color: #334155;
    }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 2. CREDENCIAIS E CONEXÃO SEGURA DIRETA
# ==========================================
SUPABASE_URL = st.secrets.get("SUPABASE_URL", "").rstrip("/")
SUPABASE_KEY = st.secrets.get("SUPABASE_KEY", "")

if not SUPABASE_URL or not SUPABASE_KEY:
    st.error("⚠️ Credenciais do Supabase não configuradas nos Secrets.")
    st.stop()

supabase = create_client(SUPABASE_URL, SUPABASE_KEY)

def sb_request(tabela, metodo="GET", payload=None):
    headers = {
        "apikey": SUPABASE_KEY,
        "Authorization": f"Bearer {SUPABASE_KEY}",
        "Content-Type": "application/json",
        "Prefer": "return=representation"
    }
    url = f"{SUPABASE_URL}/rest/v1/{tabela}"
    try:
        if metodo == "GET":
            res = requests.get(url, headers=headers)
            return res.json() if res.status_code == 200 else []
        elif metodo == "POST":
            res = requests.post(url, headers=headers, json=payload)
            return res.json() if res.status_code in [200, 201] else None
    except Exception as e:
        return []

# ==========================================
# 3. FUNÇÕES DE DADOS BLINDADAS
# ==========================================
@st.cache_data(ttl=60)
def carregar_categorias():
    data = sb_request("categorias", "GET")
    if not data or not isinstance(data, list) or len(data) == 0:
        padroes = [
            {"nome": "Dízimos e Ofertas", "tipo": "Entrada", "codigo_contabil": "3.2.10.01"},
            {"nome": "Missões", "tipo": "Saída", "codigo_contabil": "3.2.20.50"},
            {"nome": "Gestão de Pessoas", "tipo": "Saída", "codigo_contabil": "3.2.20.10"},
            {"nome": "Aluguel", "tipo": "Saída", "codigo_contabil": "3.2.20.101"},
            {"nome": "Consumo (Água, Luz)", "tipo": "Saída", "codigo_contabil": "3.2.20.15"},
            {"nome": "Manutenção do Patrimônio", "tipo": "Saída", "codigo_contabil": "3.2.20.30"},
            {"nome": "Eventos", "tipo": "Saída", "codigo_contabil": "3.2.20.40"}
        ]
        sb_request("categorias", "POST", padroes)
        data = sb_request("categorias", "GET")
    return data if isinstance(data, list) else []

@st.cache_data(ttl=60)
def carregar_lancamentos():
    lanc_data = sb_request("lancamentos", "GET")
    if not lanc_data or not isinstance(lanc_data, list):
        return pd.DataFrame()
        
    df = pd.DataFrame(lanc_data)
    if df.empty:
        return df
        
    df['valor'] = pd.to_numeric(df['valor'], errors='coerce').fillna(0.0)
    df['data_competencia'] = pd.to_datetime(df['data_competencia'], errors='coerce')
    df['mes_ano'] = df['data_competencia'].dt.strftime('%Y-%m').fillna('Desconhecido')
    
    cats = carregar_categorias()
    if cats:
        map_cat = {str(c.get('id')): c.get('nome') for c in cats}
        df['categoria_nome'] = df['categoria_id'].astype(str).map(map_cat).fillna('Sem Categoria')
    else:
        df['categoria_nome'] = 'Sem Categoria'
        
    return df

categorias_db = carregar_categorias()
df_lancamentos = carregar_lancamentos()

# ==========================================
# 4. MOTOR DE COMPRESSÃO DE IMAGENS
# ==========================================
def comprimir_e_fazer_upload(arquivo_upload):
    if arquivo_upload is None:
        return None
    nome_arquivo = f"{int(time.time())}_{arquivo_upload.name.replace(' ', '_')}"
    extensao = nome_arquivo.split('.')[-1].lower()
    bytes_data = arquivo_upload.getvalue()
    try:
        if extensao in ['jpg', 'jpeg', 'png']:
            img = Image.open(io.BytesIO(bytes_data))
            if img.mode != 'RGB': img = img.convert('RGB')
            img.thumbnail((1200, 1200), Image.Resampling.LANCZOS)
            output = io.BytesIO()
            img.save(output, format="JPEG", quality=75, optimize=True)
            bytes_data = output.getvalue()
            nome_arquivo = nome_arquivo.rsplit('.', 1)[0] + ".jpg"
            content_type = "image/jpeg"
        else:
            content_type = "application/pdf"
            
        supabase.storage.from_("comprovantes").upload(nome_arquivo, bytes_data, file_options={"content-type": content_type})
        return nome_arquivo
    except Exception as e:
        st.error(f"Erro no anexo: {e}")
        return None

# ==========================================
# 5. MENU LATERAL CLEAN E CORPORATIVO
# ==========================================
st.sidebar.markdown("<h2 style='color: #0F172A; font-weight: 700; padding-top: 10px; margin-bottom: 20px;'>⛪ COMUNA</h2>", unsafe_allow_html=True)

menu = st.sidebar.radio(
    "Menu Principal", 
    [
        "📝 Lançar Movimentação", 
        "📊 Dashboard Congregacional", 
        "📽️ Apresentação Trimestral", 
        "⚙️ Cadastros e Contabilidade"
    ],
    label_visibility="collapsed"
)

st.sidebar.markdown("---")
st.sidebar.caption("Gestão Financeira • v3.0")

# ------------------------------------------
# TELA 1: LANÇAMENTOS
# ------------------------------------------
if menu == "📝 Lançar Movimentação":
    st.title("Novo Lançamento")
    st.markdown("Registre entradas e saídas de forma rápida e segura.")
    st.markdown("")
    
    with st.form("form_lancamento", clear_on_submit=True):
        col1, col2, col3 = st.columns([1, 1, 1])
        with col1:
            tipo_lanc = st.radio("Tipo de Movimentação", ["Entrada", "Saída"], horizontal=True)
        with col2:
            valor = st.number_input("Valor (R$)", min_value=0.0, step=50.0, format="%.2f")
        with col3:
            data_comp = st.date_input("Data de Competência", date.today())
        
        st.markdown("")
        cats_filtradas = [c for c in categorias_db if c.get("tipo") == tipo_lanc]
        opcoes_cats = {c["nome"]: c["id"] for c in cats_filtradas}
        
        col4, col5 = st.columns([2, 1])
        with col4:
            descricao = st.text_input("Descrição (Ex: Conta de Luz, Dízimo Anônimo)")
        with col5:
            categoria_sel = st.selectbox("Categoria", list(opcoes_cats.keys()) if opcoes_cats else ["Nenhuma"])
        
        col6, col7 = st.columns([1, 1])
        with col6:
            tag = st.selectbox("Centro de Custo / Evento", ["Nenhum", "Retiro das Mulheres", "Acampamento Adolescentes", "Construção", "Missões Específicas"])
        with col7:
            arquivo = st.file_uploader("Comprovante / Nota Fiscal", type=['png', 'jpg', 'jpeg', 'pdf'])
        
        st.markdown("")
        submit = st.form_submit_button("💾 Salvar Movimentação no Caixa", use_container_width=True)
        
        if submit:
            if valor <= 0 or not descricao:
                st.warning("⚠️ Preencha a descrição e um valor superior a zero.")
            elif not opcoes_cats:
                st.error("⚠️ Nenhuma categoria cadastrada para este tipo.")
            else:
                with st.spinner("Salvando e otimizando anexo..."):
                    url_anexo = comprimir_e_fazer_upload(arquivo) if arquivo else None
                    dados = {
                        "descricao": descricao, 
                        "tipo": tipo_lanc, 
                        "valor": float(valor),
                        "data_competencia": str(data_comp), 
                        "status": "Concluído",
                        "categoria_id": opcoes_cats[categoria_sel],
                        "centro_custo": None if tag == "Nenhum" else tag,
                        "url_anexo": url_anexo
                    }
                    sb_request("lancamentos", "POST", dados)
                    st.cache_data.clear()
                    st.success("✅ Lançamento registrado com sucesso!")
                    time.sleep(1)
                    st.rerun()

# ------------------------------------------
# TELA 2: DASHBOARD
# ------------------------------------------
elif menu == "📊 Dashboard Congregacional":
    st.title("Visão Geral de Saúde Financeira")
    st.markdown("Acompanhamento consolidado do fluxo de caixa e metas.")
    st.markdown("")
    
    if df_lancamentos.empty:
        st.info("Nenhum lançamento registrado ainda. Comece utilizando a aba lateral de lançamentos.")
    else:
        st.subheader("Filtros de Período")
        col_f1, col_f2 = st.columns(2)
        meses_disp = sorted(df_lancamentos['mes_ano'].unique(), reverse=True)
        with col_f1:
            mes_sel = st.multiselect("Selecionar Mês/Ano", meses_disp, default=meses_disp[:3] if len(meses_disp) >= 3 else meses_disp)
        with col_f2:
            ignorar_eventos = st.checkbox("Ocultar movimentações de Eventos Especiais", value=True)
        
        df_filtrado = df_lancamentos[df_lancamentos['mes_ano'].isin(mes_sel)]
        if ignorar_eventos and 'centro_custo' in df_filtrado.columns:
            df_filtrado = df_filtrado[df_filtrado['centro_custo'].isnull()]
            
        entradas = df_filtrado[df_filtrado['tipo'] == 'Entrada']['valor'].sum()
        saidas = df_filtrado[df_filtrado['tipo'] == 'Saída']['valor'].sum()
        
        st.markdown("")
        col_k1, col_k2, col_k3, col_k4 = st.columns(4)
        col_k1.metric("Total Entradas", f"R$ {entradas:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
        col_k2.metric("Total Saídas", f"R$ {saidas:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
        col_k3.metric("Resultado do Período", f"R$ {entradas - saidas:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
        
        reserva_caixa = 150000.00 
        media_saidas = saidas / len(mes_sel) if len(mes_sel) > 0 else 1
        meses_reserva = reserva_caixa / media_saidas if media_saidas > 0 else 0
        col_k4.metric("Fôlego de Caixa", f"{meses_reserva:.1f} Meses", "Reserva Estimada")

        st.markdown("---")
        st.subheader("Evolução Mensal (Entradas vs Saídas)")
        if not df_filtrado.empty:
            df_agrupado = df_filtrado.groupby(['mes_ano', 'tipo'])['valor'].sum().reset_index()
            fig_bar = px.bar(df_agrupado, x='mes_ano', y='valor', color='tipo', barmode='group',
                             color_discrete_map={'Entrada': '#2563EB', 'Saída': '#EF4444'})
            fig_bar.add_hline(y=100000, line_dash="dot", annotation_text="Meta Entradas", line_color="#2563EB")
            fig_bar.add_hline(y=70000, line_dash="dot", annotation_text="Teto Saídas", line_color="#EF4444")
            fig_bar.update_layout(
                plot_bgcolor='rgba(0,0,0,0)', 
                paper_bgcolor='rgba(0,0,0,0)', 
                font_color='#1E293B',
                legend_title_text=''
            )
            st.plotly_chart(fig_bar, use_container_width=True)

# ------------------------------------------
# TELA 3: APRESENTAÇÃO TRIMESTRAL
# ------------------------------------------
elif menu == "📽️ Apresentação Trimestral":
    st.title("Prestação de Contas Executiva")
    st.markdown("Visão voltada para exibição em assembleias e reuniões de liderança.")
    st.markdown("")
    
    if df_lancamentos.empty:
        st.info("Insira dados de lançamentos para gerar os gráficos executivos.")
    else:
        df_saidas = df_lancamentos[df_lancamentos['tipo'] == 'Saída']
        
        col_g1, col_g2 = st.columns([1, 1])
        with col_g1:
            st.subheader("Destino dos Recursos (Saídas)")
            if not df_saidas.empty:
                df_pizza = df_saidas.groupby('categoria_nome')['valor'].sum().reset_index()
                fig_pie = px.pie(df_pizza, values='valor', names='categoria_nome', hole=0.5,
                                 color_discrete_sequence=px.colors.qualitative.Pastel)
                fig_pie.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', font_color='#1E293B')
                fig_pie.update_traces(textposition='inside', textinfo='percent+label')
                st.plotly_chart(fig_pie, use_container_width=True)
            else:
                st.info("Sem saídas registradas.")
            
        with col_g2:
            st.subheader("Narrativa Ministerial")
            st.markdown("""
            <div style="background-color: #FFFFFF; padding: 25px; border-radius: 14px; border: 1px solid #E2E8F0; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.02);">
                <h4 style="color: #059669; margin-top: 0; font-weight: 700;">🏆 Principais Conquistas</h4>
                <p style="color: #475569; font-size: 0.95rem; line-height: 1.5;">Organização automatizada do fluxo financeiro, corte de despesas redundantes e estabilização do fundo de reserva.</p>
                
                <h4 style="color: #D97706; margin-top: 20px; font-weight: 700;">🎯 Alvos e Próximos Passos</h4>
                <p style="color: #475569; font-size: 0.95rem; line-height: 1.5;">Manutenção da média orçamentária estipulada e expansão dos projetos missionários locais.</p>
            </div>
            """, unsafe_allow_html=True)

# ------------------------------------------
# TELA 4: CADASTROS E CONTABILIDADE
# ------------------------------------------
elif menu == "⚙️ Cadastros e Contabilidade":
    st.title("Plano de Contas & De/Para")
    st.markdown("Associação entre as categorias gerenciais e o escritório contábil.")
    st.markdown("")
    
    if categorias_db:
        df_cats = pd.DataFrame(categorias_db)[['nome', 'tipo', 'codigo_contabil']]
        df_cats.columns = ['Categoria Visível no App', 'Natureza', 'Código Contábil']
        st.dataframe(df_cats, use_container_width=True, hide_index=True)
    else:
        st.info("Nenhuma categoria cadastrada.")
