import streamlit as st
from supabase import create_client
from PIL import Image
import io

# Configuração da Página (Modo Apresentação / Dark)
st.set_page_config(
    page_title="Financeiro - COMUNA",
    page_icon="⛪",
    layout="wide"
)

# Estilização visual para combinar com o padrão da igreja
st.markdown("""
    <style>
    .main {
        background-color: #121212;
        color: #ffffff;
    }
    </style>
    """, unsafe_allow_html=True)

st.title("⛪ Financeiro COMUNA")
st.write("Sistema integrado de gestão e prestação de contas.")

# Configuração de Conexão com o Supabase (Pegaremos dos Secrets do Streamlit em breve)
SUPABASE_URL = st.secrets.get("SUPABASE_URL", "")
SUPABASE_KEY = st.secrets.get("SUPABASE_KEY", "")

if not SUPABASE_URL or not SUPABASE_KEY:
    st.warning("⚠️ Chaves do Supabase não encontradas. Configure os Secrets no painel do Streamlit.")
else:
    supabase = create_client(SUPABASE_URL, SUPABASE_KEY)
    st.success("Conectado ao Banco de Dados com sucesso!")

# Função para comprimir comprovantes antes de enviar (Protege o limite de 1GB)
def comprimir_e_salvar_arquivo(arquivo_upload):
    try:
        file_extension = arquivo_upload.name.split('.')[-1].lower()
        file_bytes = arquivo_upload.getvalue()
        
        # Se for imagem, comprime para economizar espaço
        if file_extension in ['jpg', 'jpeg', 'png']:
            img = Image.open(io.BytesIO(file_bytes))
            img.thumbnail((1200, 1200)) # Redimensiona limite máximo
            output = io.BytesIO()
            img.save(output, format="JPEG", quality=80) # Aplica compressão de qualidade
            file_bytes = output.getvalue()
            file_extension = 'jpg'
            
        file_path = f"notas/{arquivo_upload.name}"
        
        # Envio para o Supabase Storage 'comprovantes'
        response = supabase.storage.from_comprovantes.upload(
            file_path, 
            file_bytes, 
            file_options={"content-type": f"image/{file_extension}" if file_extension == 'jpg' else "application/pdf"}
        )
        return file_path
    except Exception as e:
        st.error(f"Erro ao enviar arquivo: {e}")
        return None

# Menu Lateral Simples para Navegação
menu = st.sidebar.selectbox("Navegação", ["Dashboard", "Lançamentos (Caixa)", "Cadastros", "Relatórios Trimestrais"])

if menu == "Dashboard":
    st.subheader("Visão Geral do Período")
    col1, col2, col3 = st.columns(3)
    col1.metric("Entradas do Mês", "R$ 0,00", "Alvo: R$ 100k")
    col2.metric("Saídas do Mês", "R$ 0,00", "Alvo: R$ 70k")
    col3.metric("Reserva em Aplicação", "3,2 Meses", "Saudável")

elif menu == "Lançamentos (Caixa)":
    st.subheader("Novo Lançamento Financeiro")
    descricao = st.text_input("Descrição da Despesa ou Entrada")
    valor = st.number_input("Valor (R$)", min_value=0.0, format="%.2f")
    tipo = st.selectbox("Tipo", ["Entrada", "Saída"])
    
    # Campo de anexo com proteção de espaço
    comprovante = st.file_uploader("Anexar Nota Fiscal ou Recibo (PDF/Foto)", type=["pdf", "png", "jpg", "jpeg"])
    
    if st.button("Salvar Lançamento"):
        st.info("Processando lançamento e salvando anexo...")
        # Aqui integraremos a gravação na tabela 'lancamentos'
