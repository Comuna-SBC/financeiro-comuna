import streamlit as st
from supabase import create_client
import pandas as pd
from PIL import Image
import io
from datetime import date
import time
import plotly.express as px
import plotly.graph_objects as go

# ==========================================
# 1. CONFIGURAÇÃO DA PÁGINA E IDENTIDADE VISUAL
# ==========================================
st.set_page_config(page_title="Financeiro COMUNA", page_icon="⛪", layout="wide")

st.markdown("""
    <style>
    .main { background-color: #0E1117; color: #FFFFFF; }
    div[data-testid="stMetricValue"] { color: #4CAF50; font-size: 2rem; }
    div[data-testid="stForm"] { background-color: #1A1C23; padding: 20px; border-radius: 10px; border: 1px solid #2D303E;}
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 2. CONEXÃO COM BANCO DE DADOS
# ==========================================
@st.cache_resource
def init_connection():
    url = st.secrets.get("SUPABASE_URL", "")
    key = st.secrets.get("SUPABASE_KEY", "")
    if not url or not key:
        return None
    return create_client(url, key)

supabase = init_connection()

if not supabase:
    st.error("⚠️ Conexão com Supabase não encontrada. Verifique os Secrets.")
    st.stop()

@st.cache_data(ttl=600)
def carregar_categorias():
    try:
        res = supabase.table("categorias").select("*").execute()
        if not res.data:
            padroes = [
                {"nome": "Dízimos e Ofertas", "tipo": "Entrada", "codigo_contabil": "3.2.10.01"},
                {"nome": "Missões", "tipo": "Saída", "codigo_contabil": "3.2.20.50"},
                {"nome": "Gestão de Pessoas", "tipo": "Saída", "codigo_contabil": "3.2.20.10"},
                {"nome": "Aluguel", "tipo": "Saída", "codigo_contabil": "3.2.20.101"},
                {"nome": "Consumo (Água, Luz)", "tipo": "Saída", "codigo_contabil": "3.2.20.15"},
                {"nome": "Manutenção do Patrimônio", "tipo": "Saída", "codigo_contabil": "3.2.20.30"},
                {"nome": "Eventos", "tipo": "Saída", "codigo_contabil": "3.2.20.40"}
            ]
            supabase.table("categorias").insert(padroes).execute()
            res = supabase.table("categorias").select("*").execute()
        return res.data if res.data else []
    except Exception as e:
        st.error(f"Erro crítico ao acessar a tabela 'categorias' no Supabase: {e}")
        return []

categorias_db = carregar_categorias()

def carregar_lancamentos():
    res = supabase.table("lancamentos").select("*, categorias(nome)").execute()
    if not res.data:
        return pd.DataFrame()
    df = pd.DataFrame(res.data)
    df['valor'] = pd.to_numeric(df['valor'])
    df['data_competencia'] = pd.to_datetime(df['data_competencia'])
    df['mes_ano'] = df['data_competencia'].dt.strftime('%Y-%m')
    df['categoria_nome'] = df['categorias'].apply(lambda x: x['nome'] if isinstance(x, dict) else 'Sem Categoria')
    return df

# ==========================================
# 3. MOTOR DE COMPRESSÃO DE IMAGENS
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
# 4. INTERFACE E NAVEGAÇÃO
# ==========================================
st.sidebar.image("https://upload.wikimedia.org/wikipedia/commons/thumb/c/ca/1x1.png/120px-1x1.png", width=50) # Espaço para logo
menu = st.sidebar.radio("Navegação", [
    "📝 Lançar Movimentação", 
    "📊 Dashboard Congregacional", 
    "📽️ Apresentação Trimestral", 
    "⚙️ Cadastros e Contabilidade"
])

df_lancamentos = carregar_lancamentos()

# ------------------------------------------
# TELA 1: LANÇAMENTOS (Simples e Rápida)
# ------------------------------------------
if menu == "📝 Lançar Movimentação":
    st.title("Caixa Diário")
    st.caption("Registro rápido de entradas e despesas. Os relatórios são atualizados automaticamente.")
    
    with st.form("form_lancamento", clear_on_submit=True):
        col1, col2, col3 = st.columns([1, 1, 1])
        tipo_lanc = col1.radio("Tipo", ["Entrada", "Saída"], horizontal=True)
        valor = col2.number_input("Valor (R$)", min_value=0.0, step=50.0, format="%.2f")
        data_comp = col3.date_input("Data", date.today())
        
        cats_filtradas = [c for c in categorias_db if c["tipo"] == tipo_lanc]
        opcoes_cats = {c["nome"]: c["id"] for c in cats_filtradas}
        
        col4, col5 = st.columns([2, 1])
        descricao = col4.text_input("Descrição da Movimentação")
        categoria_sel = col5.selectbox("Categoria Contábil/Gerencial", list(opcoes_cats.keys()))
        
        col6, col7 = st.columns([1, 1])
        tag = col6.selectbox("Projeto / Evento (Centro de Custo)", ["Nenhum", "Retiro das Mulheres", "Acampamento Adolescentes", "Construção", "Missões Específicas"])
        arquivo = col7.file_uploader("Anexar Nota Fiscal / Recibo", type=['png', 'jpg', 'jpeg', 'pdf'])
        
        if st.form_submit_button("💾 Salvar Lançamento", use_container_width=True):
            if valor <= 0 or not descricao:
                st.warning("⚠️ Preencha a descrição e um valor válido.")
            else:
                with st.spinner("Processando..."):
                    url_anexo = comprimir_e_fazer_upload(arquivo) if arquivo else None
                    dados = {
                        "descricao": descricao, "tipo": tipo_lanc, "valor": float(valor),
                        "data_competencia": str(data_comp), "status": "Concluído",
                        "categoria_id": opcoes_cats[categoria_sel],
                        "centro_custo": None if tag == "Nenhum" else tag,
                        "url_anexo": url_anexo
                    }
                    supabase.table("lancamentos").insert(dados).execute()
                    st.success("✅ Registrado com sucesso!")
                    time.sleep(1)
                    st.rerun()

# ------------------------------------------
# TELA 2: DASHBOARD (Filtros e KPIs)
# ------------------------------------------
elif menu == "📊 Dashboard Congregacional":
    st.title("Visão de Saúde Financeira")
    
    if df_lancamentos.empty:
        st.info("Nenhum dado para exibir.")
    else:
        # Filtros Superiores
        st.write("### Filtros")
        col_f1, col_f2 = st.columns(2)
        meses_disp = sorted(df_lancamentos['mes_ano'].unique(), reverse=True)
        mes_sel = col_f1.multiselect("Selecionar Mês/Ano", meses_disp, default=meses_disp[:3] if len(meses_disp) >=3 else meses_disp)
        ignorar_eventos = col_f2.checkbox("Ocultar movimentações de Eventos (Retiro/Acampamento)", value=True)
        
        df_filtrado = df_lancamentos[df_lancamentos['mes_ano'].isin(mes_sel)]
        if ignorar_eventos:
            df_filtrado = df_filtrado[df_filtrado['centro_custo'].isnull()]
            
        # KPIs e Alvos
        entradas = df_filtrado[df_filtrado['tipo'] == 'Entrada']['valor'].sum()
        saidas = df_filtrado[df_filtrado['tipo'] == 'Saída']['valor'].sum()
        
        st.divider()
        col_k1, col_k2, col_k3, col_k4 = st.columns(4)
        col_k1.metric("Total Entradas", f"R$ {entradas:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
        col_k2.metric("Total Saídas", f"R$ {saidas:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
        col_k3.metric("Resultado", f"R$ {entradas - saidas:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
        
        # Simulação de Reserva (Investimento fictício de 150k para o cálculo de fôlego)
        reserva_caixa = 150000.00 
        media_saidas = saidas / len(mes_sel) if len(mes_sel) > 0 else 1
        meses_reserva = reserva_caixa / media_saidas if media_saidas > 0 else 0
        col_k4.metric("Fôlego Financeiro", f"{meses_reserva:.1f} Meses", "Reserva em Aplicação")

        # Gráfico de Entradas vs Saídas com Linha de Alvo
        st.subheader("Entradas vs Saídas (Evolução)")
        df_agrupado = df_filtrado.groupby(['mes_ano', 'tipo'])['valor'].sum().reset_index()
        fig_bar = px.bar(df_agrupado, x='mes_ano', y='valor', color='tipo', barmode='group',
                         color_discrete_map={'Entrada': '#1E88E5', 'Saída': '#FF7043'})
        fig_bar.add_hline(y=100000, line_dash="dot", annotation_text="Alvo Entradas (100k)", line_color="#1E88E5")
        fig_bar.add_hline(y=70000, line_dash="dot", annotation_text="Alvo Saídas (70k)", line_color="#FF7043")
        fig_bar.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', font_color='white')
        st.plotly_chart(fig_bar, use_container_width=True)

# ------------------------------------------
# TELA 3: APRESENTAÇÃO TRIMESTRAL (Substitui o PDF/PPT)
# ------------------------------------------
elif menu == "📽️ Apresentação Trimestral":
    st.title("Reunião de Prestação de Contas")
    st.caption("Projete esta tela durante a assembleia. Os gráficos reproduzem os slides oficiais.")
    
    if df_lancamentos.empty:
        st.info("Necessário lançar dados primeiro.")
    else:
        df_saidas = df_lancamentos[df_lancamentos['tipo'] == 'Saída']
        
        col_g1, col_g2 = st.columns([1, 1])
        with col_g1:
            st.subheader("Destino das Saídas (Pizza)")
            df_pizza = df_saidas.groupby('categoria_nome')['valor'].sum().reset_index()
            fig_pie = px.pie(df_pizza, values='valor', names='categoria_nome', hole=0.4)
            fig_pie.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', font_color='white')
            fig_pie.update_traces(textposition='inside', textinfo='percent+label')
            st.plotly_chart(fig_pie, use_container_width=True)
            
        with col_g2:
            st.subheader("Narrativa do Trimestre")
            st.markdown("### 🏆 Conquistas")
            st.markdown("- Manutenção e Reformas de infra e ar-condicionado\n- Instalação do toldo\n- Biblioteca missional ativa")
            
            st.markdown("### 🎯 Desafios e Alvos")
            st.markdown("- Melhoria de Gestão Orçamentária\n- Alcance da média de R$ 100 mil mensais")
            
            st.markdown("### 📝 Notas Explicativas")
            st.markdown("*1- Antecipação de parte das despesas do próximo mês*\n*2- Valores de eventos estão isolados em seus centros de custo*")

# ------------------------------------------
# TELA 4: CADASTROS E CONTABILIDADE
# ------------------------------------------
elif menu == "⚙️ Cadastros e Contabilidade":
    st.title("De/Para Contábil")
    st.write("Associe as categorias gerenciais (amigáveis) ao Plano de Contas do Contador.")
    
    df_cats = pd.DataFrame(categorias_db)
    if not df_cats.empty:
        df_cats = df_cats[['nome', 'tipo', 'codigo_contabil']]
        df_cats.columns = ['Categoria Visível no App', 'Natureza', 'Código Contábil (Escritório)']
        st.dataframe(df_cats, use_container_width=True, hide_index=True)
