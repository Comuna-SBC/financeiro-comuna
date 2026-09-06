import streamlit as st
from supabase import create_client
import pandas as pd
from PIL import Image
import io
from datetime import date
import time

# ==========================================
# 1. CONFIGURAÇÃO DA PÁGINA E IDENTIDADE VISUAL
# ==========================================
st.set_page_config(page_title="Financeiro COMUNA", page_icon="⛪", layout="wide")

# CSS para forçar um design limpo e focado, parecido com a identidade da apresentação
st.markdown("""
    <style>
    .main { background-color: #0E1117; color: #FFFFFF; }
    div[data-testid="stMetricValue"] { color: #4CAF50; font-size: 2rem; }
    div[data-testid="stForm"] { background-color: #1A1C23; padding: 20px; border-radius: 10px; border: 1px solid #2D303E;}
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 2. CONEXÃO COM BANCO DE DADOS E SETUP
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
    st.error("⚠️ Conexão com Supabase não encontrada. Configure os Secrets no painel do Streamlit.")
    st.stop()

# Função inteligente: se o banco for novo, cria as categorias padrão automaticamente
@st.cache_data(ttl=600)
def carregar_categorias():
    res = supabase.table("categorias").select("*").execute()
    if not res.data: # Se estiver vazio, popula automaticamente
        padroes = [
            {"nome": "Dízimos e Ofertas", "tipo": "Entrada"},
            {"nome": "Missões", "tipo": "Saída"},
            {"nome": "Gestão de Pessoas", "tipo": "Saída"},
            {"nome": "Aluguel", "tipo": "Saída"},
            {"nome": "Consumo (Água, Luz, Net)", "tipo": "Saída"},
            {"nome": "Manutenção / Patrimônio", "tipo": "Saída"},
            {"nome": "Eventos", "tipo": "Saída"}
        ]
        supabase.table("categorias").insert(padroes).execute()
        res = supabase.table("categorias").select("*").execute()
    return res.data

categorias_db = carregar_categorias()

# ==========================================
# 3. MOTOR DE COMPRESSÃO DE RECIBOS (Economiza o Plano Grátis)
# ==========================================
def comprimir_e_fazer_upload(arquivo_upload):
    if arquivo_upload is None:
        return None
    
    nome_arquivo = f"{int(time.time())}_{arquivo_upload.name.replace(' ', '_')}"
    extensao = nome_arquivo.split('.')[-1].lower()
    bytes_data = arquivo_upload.getvalue()
    
    try:
        # Se for imagem pesada de celular, aplica compressão severa sem perder nitidez da letra
        if extensao in ['jpg', 'jpeg', 'png']:
            img = Image.open(io.BytesIO(bytes_data))
            if img.mode != 'RGB':
                img = img.convert('RGB')
            # Limita tamanho máximo para 1200 pixels
            img.thumbnail((1200, 1200), Image.Resampling.LANCZOS)
            output = io.BytesIO()
            img.save(output, format="JPEG", quality=75, optimize=True)
            bytes_data = output.getvalue()
            nome_arquivo = nome_arquivo.rsplit('.', 1)[0] + ".jpg"
            content_type = "image/jpeg"
        else:
            # Se for PDF, sobe direto
            content_type = "application/pdf"
            
        # Faz o upload para o bucket 'comprovantes'
        res = supabase.storage.from_("comprovantes").upload(
            nome_arquivo, 
            bytes_data, 
            file_options={"content-type": content_type}
        )
        return nome_arquivo
    except Exception as e:
        st.error(f"Erro ao processar anexo: {e}")
        return None

# ==========================================
# 4. INTERFACE DO USUÁRIO (Navegação Simplificada)
# ==========================================
menu = st.sidebar.radio("Navegação", ["📝 Lançar Nova Movimentação", "📊 Dashboard e Caixa", "⚙️ Cadastros"])

# ------------------------------------------
# TELA 1: LANÇAMENTOS (Foco em velocidade, substitui a digitação no Excel)
# ------------------------------------------
if menu == "📝 Lançar Nova Movimentação":
    st.title("Novo Lançamento")
    st.caption("Preencha os dados abaixo. O sistema já salva e cruza os dados com o Dashboard automaticamente.")
    
    with st.form("form_lancamento", clear_on_submit=True):
        col1, col2, col3 = st.columns([1, 1, 1])
        tipo_lancamento = col1.radio("Tipo", ["Entrada", "Saída"], horizontal=True)
        valor = col2.number_input("Valor (R$)", min_value=0.0, step=10.0, format="%.2f")
        data_comp = col3.date_input("Data", date.today())
        
        # Filtra categorias baseado se o usuário marcou Entrada ou Saída
        cats_filtradas = [c for c in categorias_db if c["tipo"] == tipo_lancamento]
        opcoes_cats = {c["nome"]: c["id"] for c in cats_filtradas}
        
        col4, col5 = st.columns([2, 1])
        descricao = col4.text_input("Descrição (Ex: Conta de Luz de Março, Dízimo Fulano)")
        categoria_selecionada = col5.selectbox("Categoria", list(opcoes_cats.keys()))
        
        col6, col7 = st.columns([1, 1])
        tag = col6.selectbox("Projeto / Evento (Opcional)", ["Nenhum", "Retiro das Mulheres", "Acampamento Adolescentes", "Construção"])
        arquivo = col7.file_uploader("Anexar Comprovante / NF", type=['png', 'jpg', 'jpeg', 'pdf'])
        
        submit = st.form_submit_button("💾 Salvar Lançamento no Caixa", use_container_width=True)
        
        if submit:
            if valor <= 0 or not descricao:
                st.warning("⚠️ Preencha a descrição e um valor maior que zero.")
            else:
                with st.spinner("Salvando lançamento e otimizando anexo..."):
                    # 1. Sobe o arquivo
                    caminho_anexo = comprimir_e_fazer_upload(arquivo) if arquivo else None
                    
                    # 2. Monta o pacote de dados para o banco
                    dados = {
                        "descricao": descricao,
                        "tipo": tipo_lancamento,
                        "valor": float(valor),
                        "data_competencia": str(data_comp),
                        "status": "Concluído",
                        "categoria_id": opcoes_cats[categoria_selecionada],
                        "centro_custo": None if tag == "Nenhum" else tag,
                        "url_anexo": caminho_anexo
                    }
                    
                    # 3. Insere no Supabase
                    supabase.table("lancamentos").insert(dados).execute()
                    st.success(f"✅ Lançamento de R$ {valor:.2f} registrado com sucesso!")
                    time.sleep(1.5)
                    st.rerun()

# ------------------------------------------
# TELA 2: DASHBOARD (Substitui os slides e o Consolidado)
# ------------------------------------------
elif menu == "📊 Dashboard e Caixa":
    st.title("Visão Geral Congregacional")
    
    # Busca lançamentos
    res_lancamentos = supabase.table("lancamentos").select("*").execute()
    df = pd.DataFrame(res_lancamentos.data)
    
    if df.empty:
        st.info("Nenhum lançamento registrado ainda.")
    else:
        # Converte tipos
        df['valor'] = pd.to_numeric(df['valor'])
        df['data_competencia'] = pd.to_datetime(df['data_competencia'])
        
        # Filtros rápidos no topo
        mes_atual = date.today().month
        entradas_totais = df[(df['tipo'] == 'Entrada') & (df['data_competencia'].dt.month == mes_atual)]['valor'].sum()
        saidas_totais = df[(df['tipo'] == 'Saída') & (df['data_competencia'].dt.month == mes_atual)]['valor'].sum()
        saldo = entradas_totais - saidas_totais
        
        # Cards estilo Apresentação
        col1, col2, col3 = st.columns(3)
        col1.metric("Entradas (Mês Atual)", f"R$ {entradas_totais:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'), "Alvo: R$ 100.000,00")
        col2.metric("Saídas (Mês Atual)", f"R$ {saidas_totais:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'), "Alvo: R$ 70.000,00", delta_color="inverse")
        col3.metric("Resultado do Mês", f"R$ {saldo:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
        
        st.divider()
        
        # Tabela de Extrato Simplificada
        st.subheader("Últimos Lançamentos")
        df_display = df[['data_competencia', 'tipo', 'descricao', 'valor', 'centro_custo']].copy()
        df_display['data_competencia'] = df_display['data_competencia'].dt.strftime('%d/%m/%Y')
        df_display.columns = ['Data', 'Tipo', 'Descrição', 'Valor (R$)', 'Projeto']
        st.dataframe(df_display.sort_values(by='Data', ascending=False).head(15), use_container_width=True, hide_index=True)

# ------------------------------------------
# TELA 3: CADASTROS BÁSICOS
# ------------------------------------------
elif menu == "⚙️ Cadastros":
    st.title("Configurações e Categorias")
    st.write("Aqui você pode adicionar novas contas ao plano de contas contábil/gerencial da Igreja.")
    
    with st.form("nova_categoria"):
        col1, col2 = st.columns(2)
        nova_cat_nome = col1.text_input("Nome da Categoria (Ex: Construção, Equipamentos)")
        nova_cat_tipo = col2.selectbox("Tipo", ["Entrada", "Saída"])
        
        if st.form_submit_button("Salvar Categoria"):
            supabase.table("categorias").insert({"nome": nova_cat_nome, "tipo": nova_cat_tipo}).execute()
            st.success(f"Categoria {nova_cat_nome} criada!")
            st.cache_data.clear() # Limpa o cache para atualizar a lista na tela de lançamento
            time.sleep(1)
            st.rerun()
            
    st.subheader("Categorias Existentes")
    df_cats = pd.DataFrame(categorias_db)
    if not df_cats.empty:
        st.dataframe(df_cats[['nome', 'tipo']], use_container_width=True, hide_index=True)
