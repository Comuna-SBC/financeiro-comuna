import streamlit as st
from supabase import create_client
import pandas as pd
from PIL import Image
import io
from datetime import date
import time
import plotly.express as px
import plotly.graph_objects as go
import requests
import re
import zipfile

# ==========================================
# CONFIGURAÇÃO DA PÁGINA E DESIGN SYSTEM
# ==========================================
st.set_page_config(
    page_title="Gestão Financeira Igreja", 
    page_icon="⛪", 
    layout="wide", 
    initial_sidebar_state="expanded"
)

MESES_PT = [
    "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
    "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
]

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');
    html, body, [class*="css"] { font-family: 'Plus Jakarta Sans', sans-serif; }
    .main { background-color: #F8FAFC; color: #1E293B; padding: 2rem; }
    [data-testid="stSidebar"] { background-color: #FFFFFF; border-right: 1px solid #E2E8F0; }
    [data-testid="stSidebar"] .stButton button {
        width: 100%; text-align: left; justify-content: flex-start;
        border-radius: 10px; padding: 0.65rem 1rem; font-weight: 500;
        margin-bottom: 4px; border: 1px solid transparent; font-size: 0.92rem;
    }
    [data-testid="stSidebar"] button[kind="secondary"] {
        background-color: transparent !important; color: #475569 !important; box-shadow: none !important;
    }
    [data-testid="stSidebar"] button[kind="secondary"]:hover {
        background-color: #F1F5F9 !important; color: #0F172A !important;
    }
    [data-testid="stSidebar"] button[kind="primary"] {
        background-color: #EFF6FF !important; color: #2563EB !important;
        border: 1px solid #BFDBFE !important; box-shadow: none !important; font-weight: 700 !important;
    }
    div[data-testid="stMetric"] {
        background: #FFFFFF; border: 1px solid #E2E8F0; padding: 20px; border-radius: 14px;
        box-shadow: 0 4px 6px -1px rgba(0,0,0,0.02);
    }
    div[data-testid="stMetricValue"] { color: #059669; font-size: 1.8rem !important; font-weight: 700; }
    div[data-testid="stMetricLabel"] { color: #64748B !important; font-size: 0.85rem !important; font-weight: 600; text-transform: uppercase; letter-spacing: 0.04em; }
    div[data-testid="stForm"], div[data-testid="stExpander"] {
        background-color: #FFFFFF; padding: 25px; border-radius: 16px;
        border: 1px solid #E2E8F0; box-shadow: 0 10px 15px -3px rgba(0,0,0,0.03);
    }
    .stTextInput input, .stNumberInput input, .stSelectbox select, .stDateInput input, textarea {
        background-color: #F8FAFC !important; color: #0F172A !important;
        border: 1px solid #CBD5E1 !important; border-radius: 10px !important; padding: 10px 14px !important;
    }
    .stTextInput input:focus, .stNumberInput input:focus {
        border-color: #2563EB !important; background-color: #FFFFFF !important;
        box-shadow: 0 0 0 3px rgba(37,99,235,0.15);
    }
    .main .stButton button, div[data-testid="stFormSubmitButton"] button {
        background-color: #2563EB !important; color: white !important; font-weight: 600 !important;
        padding: 0.8rem 1.4rem !important; border-radius: 10px !important; border: none !important;
        box-shadow: 0 4px 12px rgba(37,99,235,0.2);
    }
    .main .stButton button:hover, div[data-testid="stFormSubmitButton"] button:hover {
        background-color: #1D4ED8 !important;
    }
    h1, h2, h3, h4 { color: #0F172A !important; font-weight: 700 !important; }
    
    .badge-falta-anexo {
        background-color: #FEF2F2;
        color: #DC2626;
        border: 1px solid #FCA5A5;
        padding: 2px 8px;
        border-radius: 6px;
        font-size: 12px;
        font-weight: 600;
        display: inline-block;
    }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# CONEXÃO COM SUPABASE & HELPER DE EXPORTAÇÃO
# ==========================================
SUPABASE_URL = st.secrets.get("SUPABASE_URL", "").rstrip("/")
SUPABASE_KEY = st.secrets.get("SUPABASE_KEY", "")

if not SUPABASE_URL or not SUPABASE_KEY:
    st.error("⚠️ Credenciais do Supabase não configuradas nos Secrets.")
    st.stop()

supabase = create_client(SUPABASE_URL, SUPABASE_KEY)

def sb_request(tabela, metodo="GET", payload=None, filtros=None):
    headers = {
        "apikey": SUPABASE_KEY,
        "Authorization": f"Bearer {SUPABASE_KEY}",
        "Content-Type": "application/json",
        "Prefer": "return=representation"
    }
    url = f"{SUPABASE_URL}/rest/v1/{tabela}"
    try:
        res = None
        if metodo == "GET":
            res = requests.get(url, headers=headers, params=filtros or {})
        elif metodo == "POST":
            res = requests.post(url, headers=headers, json=payload)
        elif metodo == "PATCH":
            res = requests.patch(url, headers=headers, params=filtros or {}, json=payload)
        elif metodo == "DELETE":
            res = requests.delete(url, headers=headers, params=filtros or {})

        if res is not None and not res.ok:
            st.error(f"❌ Erro Supabase ({metodo} {tabela}): Status {res.status_code} - {res.text}")
            return [] if metodo == "GET" else None

        if metodo == "DELETE":
            return True
        return res.json() if res.content else []
    except Exception as e:
        st.error(f"❌ Falha de Conexão Supabase: {e}")
        return [] if metodo == "GET" else None

def to_excel_bytes(dfs_dict):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet_name, df_data in dfs_dict.items():
            df_data.to_excel(writer, sheet_name=sheet_name)
    return output.getvalue()

@st.cache_data(ttl=15)
def carregar(tabela):
    return sb_request(tabela, "GET") or []

def limpar_nome_arquivo(texto):
    if not texto:
        return "geral"
    texto_limpo = re.sub(r'[^a-zA-Z0-9]', '_', texto)
    return re.sub(r'_+', '_', texto_limpo).strip('_')

def processar_e_salvar_anexos(arquivos_upload, lancamento_id, data_lanc, categoria_nome, descricao):
    if not arquivos_upload:
        return
    
    data_str = pd.to_datetime(data_lanc).strftime('%Y%m%d')
    cat_limpa = limpar_nome_arquivo(categoria_nome)
    desc_limpa = limpar_nome_arquivo(descricao)
    
    for idx, arquivo in enumerate(arquivos_upload, start=1):
        extensao = arquivo.name.split('.')[-1].lower()
        nome_padronizado = f"notas/{data_str}.{cat_limpa}.{desc_limpa}.{int(time.time())}.{idx:02d}.{extensao}"
        bytes_data = arquivo.getvalue()
        
        try:
            if extensao in ['jpg', 'jpeg', 'png']:
                img = Image.open(io.BytesIO(bytes_data))
                if img.mode != 'RGB':
                    img = img.convert('RGB')
                img.thumbnail((1200, 1200), Image.Resampling.LANCZOS)
                output = io.BytesIO()
                img.save(output, format="JPEG", quality=75, optimize=True)
                bytes_data = output.getvalue()
                nome_padronizado = nome_padronizado.rsplit('.', 1)[0] + ".jpg"
                content_type = "image/jpeg"
            else:
                content_type = "application/pdf"
            
            supabase.storage.from_("comprovantes").upload(
                nome_padronizado, bytes_data, 
                file_options={"content-type": content_type, "upsert": "true"}
            )
            
            nome_limpo_arquivo = arquivo.name
            sb_request("lancamento_anexos", "POST", {
                "lancamento_id": lancamento_id,
                "url_storage": nome_padronizado,
                "nome_original": nome_limpo_arquivo
            })
        except Exception as e:
            st.error(f"Erro ao salvar o anexo {arquivo.name}: {e}")

@st.cache_data(ttl=15)
def carregar_categorias():
    data = sb_request("categorias", "GET") or []
    nomes_existentes = {c['nome'] for c in data}
    padroes = [
        {"nome": "Dízimos e Ofertas", "tipo": "Entrada", "codigo_contabil": "3.1.10.01"},
        {"nome": "Inscrições de Eventos", "tipo": "Entrada", "codigo_contabil": "3.1.10.05"},
        {"nome": "Missões", "tipo": "Saída", "codigo_contabil": "3.2.20.50"},
        {"nome": "Gestão de Pessoas", "tipo": "Saída", "codigo_contabil": "3.2.20.10"},
        {"nome": "Aluguel", "tipo": "Saída", "codigo_contabil": "3.2.20.101"},
        {"nome": "Consumo (Água, Luz)", "tipo": "Saída", "codigo_contabil": "3.2.20.15"},
        {"nome": "Manutenção do Patrimônio", "tipo": "Saída", "codigo_contabil": "3.2.20.30"},
        {"nome": "Eventos (Despesas)", "tipo": "Saída", "codigo_contabil": "3.2.20.40"}
    ]
    faltantes = [p for p in padroes if p["nome"] not in nomes_existentes]
    if faltantes:
        sb_request("categorias", "POST", faltantes)
        st.cache_data.clear()
        data = sb_request("categorias", "GET") or []
    return data

def carregar_lancamentos_df():
    lanc = carregar("lancamentos")
    if not lanc:
        return pd.DataFrame()
    df = pd.DataFrame(lanc)
    
    for col in ['valor', 'data_competencia', 'status', 'categoria_id', 'conta_bancaria_id', 'centro_custo', 'descricao', 'tipo']:
        if col not in df.columns:
            df[col] = None
            
    df['valor'] = pd.to_numeric(df['valor'], errors='coerce').fillna(0.0)
    df['data_competencia'] = pd.to_datetime(df['data_competencia'], errors='coerce')
    df['data_vencimento'] = pd.to_datetime(df.get('data_vencimento'), errors='coerce') if 'data_vencimento' in df.columns else pd.NaT
    df['conciliado'] = df['conciliado'].fillna(False) if 'conciliado' in df.columns else False
    df['mes_ano'] = df['data_competencia'].dt.strftime('%Y-%m')
    df['mes_num'] = df['data_competencia'].dt.month
    df['ano_num'] = df['data_competencia'].dt.year

    cats = carregar_categorias()
    map_cat = {str(c['id']): c['nome'] for c in cats}
    df['categoria_nome'] = df['categoria_id'].astype(str).map(map_cat).fillna('Sem Categoria')

    contas = carregar("contas_bancarias")
    map_conta = {str(c['id']): c['nome'] for c in contas}
    if 'conta_bancaria_id' in df.columns:
        df['conta_nome'] = df['conta_bancaria_id'].astype(str).map(map_conta).fillna('—')
    else:
        df['conta_nome'] = '—'
    return df

def fmt_moeda(v):
    try:
        return f"R$ {float(v):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    except Exception:
        return "R$ 0,00"

@st.cache_data(ttl=15)
def carregar_usuarios():
    return sb_request("usuarios", "GET") or []

def somente_digitos(txt):
    return "".join(ch for ch in (txt or "") if ch.isdigit())

def cpf_valido(cpf):
    cpf = somente_digitos(cpf)
    if len(cpf) != 11 or cpf == cpf[0] * 11:
        return False
    for i in [9, 10]:
        soma = sum(int(cpf[num]) * ((i + 1) - num) for num in range(0, i))
        digito = ((soma * 10) % 11) % 10
        if digito != int(cpf[i]):
            return False
    return True

def comprimir_e_fazer_upload(arquivo_upload, pasta="notas"):
    if arquivo_upload is None:
        return None
    nome_arquivo = f"{pasta}/{int(time.time())}_{arquivo_upload.name.replace(' ', '_')}"
    extensao = nome_arquivo.split('.')[-1].lower()
    bytes_data = arquivo_upload.getvalue()
    try:
        if extensao in ['jpg', 'jpeg', 'png']:
            img = Image.open(io.BytesIO(bytes_data))
            if img.mode != 'RGB':
                img = img.convert('RGB')
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

def obter_link_arquivo(path):
    if not path:
        return None
    try:
        res = supabase.storage.from_("comprovantes").create_signed_url(path, 3600)
        return res.get("signedURL") or res.get("signed_url")
    except Exception:
        return None

def upsert_meta(ano, mes, meta_entradas, meta_saidas):
    existentes = carregar("metas_mensais")
    match = next((m for m in existentes if int(m["ano"]) == ano and int(m["mes"]) == mes), None)
    payload = {"ano": ano, "mes": mes, "meta_entradas": meta_entradas, "meta_saidas": meta_saidas}
    if match:
        sb_request("metas_mensais", "PATCH", payload, filtros={"id": f"eq.{match['id']}"})
    else:
        sb_request("metas_mensais", "POST", payload)
    st.cache_data.clear()

def upsert_orcamento(ano, categoria_id, valor):
    existentes = carregar("orcamentos_categoria")
    match = next((o for o in existentes if int(o["ano"]) == ano and o["categoria_id"] == categoria_id), None)
    payload = {"ano": ano, "categoria_id": categoria_id, "valor_orcado": valor}
    if match:
        sb_request("orcamentos_categoria", "PATCH", payload, filtros={"id": f"eq.{match['id']}"})
    else:
        sb_request("orcamentos_categoria", "POST", payload)
    st.cache_data.clear()

# ==========================================
# PÁGINA PÚBLICA DE INSCRIÇÃO
# ==========================================
def pagina_inscricao_publica():
    st.markdown("<h1 style='text-align:center;'>⛪ Inscrição em Evento</h1>", unsafe_allow_html=True)
    eventos_abertos = [e for e in carregar("eventos") if e.get("status") == "Aberto"]
    if not eventos_abertos:
        st.info("Não há eventos com inscrições abertas no momento.")
        return

    qp = st.query_params
    evento_id_param = qp.get("evento")
    evento_travado = next((e for e in eventos_abertos if str(e["id"]) == str(evento_id_param)), None) if evento_id_param else None

    if evento_travado:
        evento = evento_travado
        st.info(f"Você está se inscrevendo em: **{evento['nome']}**")
    else:
        nomes = [e["nome"] for e in eventos_abertos]
        evento_sel_nome = st.selectbox("Selecione o Evento", nomes)
        evento = next(e for e in eventos_abertos if e["nome"] == evento_sel_nome)

    parcela_info = f"<p style='margin:0;color:#475569;'>📆 Pagamento em até <b>{evento.get('numero_parcelas',1)}x</b></p>" if evento.get("permite_parcelamento") else ""
    st.markdown(f"""
    <div style="background:#FFFFFF;border:1px solid #E2E8F0;border-radius:14px;padding:20px;margin-bottom:20px;">
        <p style="margin:0;color:#475569;">📅 Data: <b>{evento.get('data_evento','—')}</b></p>
        <p style="margin:0;color:#475569;">💰 Valor Total: <b>{fmt_moeda(evento.get('valor_inscricao'))}</b></p>
        <p style="margin:0;color:#475569;">🔑 Chave Pix: <b>{evento.get('chave_pix','—')}</b></p>
        {parcela_info}
    </div>
    """, unsafe_allow_html=True)

    st.markdown("### 🔎 Informe seu CPF para começar")
    cpf_busca = st.text_input("CPF", key="cpf_busca", placeholder="Somente números")
    cpf_limpo_busca = somente_digitos(cpf_busca)

    inscricao_existente = None
    if len(cpf_limpo_busca) == 11:
        todas_insc = carregar("inscricoes")
        inscricao_existente = next(
            (i for i in todas_insc if i.get("evento_id") == evento["id"] and somente_digitos(i.get("cpf", "")) == cpf_limpo_busca),
            None
        )

    if inscricao_existente:
        _painel_pagamentos_participante(inscricao_existente, evento)
        return

    if cpf_limpo_busca and len(cpf_limpo_busca) == 11:
        st.caption("CPF não encontrado para este evento — preencha os dados abaixo para se inscrever.")

    with st.form("form_inscricao_publica", clear_on_submit=False):
        nome = st.text_input("Nome completo")
        cpf = st.text_input("CPF", value=cpf_busca or "", placeholder="Somente números")
        contato = st.text_input("Telefone / WhatsApp")
        enviar = st.form_submit_button("Criar Inscrição", use_container_width=True)
        if enviar:
            cpf_limpo = somente_digitos(cpf)
            if not nome or not contato or not cpf_valido(cpf_limpo):
                st.warning("Preencha nome, contato e um CPF válido.")
            else:
                todas_insc = carregar("inscricoes")
                duplicado = any(
                    i.get("evento_id") == evento["id"] and somente_digitos(i.get("cpf", "")) == cpf_limpo
                    for i in todas_insc
                )
                if duplicado:
                    st.error("Este CPF já está inscrito neste evento. Digite o CPF no campo acima para carregar sua inscrição e enviar o comprovante.")
                else:
                    nova = sb_request("inscricoes", "POST", {
                        "evento_id": evento["id"],
                        "nome_participante": nome,
                        "contato": contato,
                        "cpf": cpf_limpo,
                        "valor_total": float(evento.get("valor_inscricao") or 0),
                        "valor_pago": 0,
                        "status_pagamento": "Pendente"
                    })
                    if nova is not None:
                        st.cache_data.clear()
                        st.success("✅ Inscrição criada! Agora envie o comprovante do pagamento abaixo.")
                        st.session_state["cpf_busca"] = cpf_limpo
                        st.rerun()
                    else:
                        st.error("Não foi possível criar a inscrição. Verifique as mensagens de erro acima.")

def _painel_pagamentos_participante(inscricao, evento):
    valor_total = float(inscricao.get("valor_total") or 0)
    valor_pago = float(inscricao.get("valor_pago") or 0)
    saldo = max(round(valor_total - valor_pago, 2), 0)

    st.markdown(f"### 👤 {inscricao['nome_participante']}")
    c1, c2, c3 = st.columns(3)
    c1.metric("Valor Total", fmt_moeda(valor_total))
    c2.metric("Já Pago (aprovado)", fmt_moeda(valor_pago))
    c3.metric("Saldo Restante", fmt_moeda(saldo))
    st.progress(min(valor_pago / valor_total, 1.0) if valor_total > 0 else 0.0)

    pagamentos = [p for p in carregar("inscricao_pagamentos") if p.get("inscricao_id") == inscricao["id"]]
    if pagamentos:
        st.markdown("#### Comprovantes enviados")
        for p in sorted(pagamentos, key=lambda x: x.get("numero_parcela", 1)):
            emoji = {"Pendente": "⏳", "Aprovado": "✅", "Rejeitado": "❌"}.get(p["status"], "")
            st.write(f"{emoji} Parcela {p.get('numero_parcela',1)} — {fmt_moeda(p.get('valor'))} — {p['status']}")

    if saldo <= 0.01:
        st.success("🎉 Inscrição totalmente paga. Nenhum novo comprovante é necessário.")
        return

    st.markdown("#### 📎 Enviar novo comprovante")
    proxima_parcela = len(pagamentos) + 1
    parcelas_totais = evento.get("numero_parcelas") or 1
    parcelas_restantes = max(parcelas_totais - len(pagamentos), 1) if evento.get("permite_parcelamento") else 1
    sugestao = min(round(saldo / parcelas_restantes, 2), saldo)

    with st.form("form_novo_comprovante", clear_on_submit=True):
        valor_parcela = st.number_input(
            "Valor pago nesta parcela (R$)", min_value=0.01, max_value=float(saldo),
            value=float(sugestao if sugestao > 0 else saldo), format="%.2f"
        )
        comprovante = st.file_uploader("Comprovante do Pix", type=['png', 'jpg', 'jpeg', 'pdf'])
        enviar_pg = st.form_submit_button("Enviar Comprovante", use_container_width=True)
        if enviar_pg:
            if not comprovante:
                st.warning("Anexe o comprovante.")
            else:
                url_comp = comprimir_e_fazer_upload(comprovante, pasta="eventos")
                sucesso = sb_request("inscricao_pagamentos", "POST", {
                    "inscricao_id": inscricao["id"],
                    "numero_parcela": proxima_parcela,
                    "valor": float(valor_parcela),
                    "comprovante_url": url_comp,
                    "status": "Pendente"
                })
                if sucesso is not None:
                    st.cache_data.clear()
                    st.success("✅ Comprovante enviado! A tesouraria irá validar em breve.")
                    st.rerun()

qp = st.query_params
if qp.get("pagina") == "inscricao":
    st.markdown("<style>[data-testid='stSidebar'], [data-testid='collapsedControl'] {display:none;}</style>", unsafe_allow_html=True)
    pagina_inscricao_publica()
    st.stop()

# ==========================================
# MENU LATERAL INTERNO E CONTROLE DE ACESSO (RBAC)
# ==========================================
categorias_db = carregar_categorias()
eventos_db = carregar("eventos")
contas_bancarias_db = carregar("contas_bancarias")
usuarios_db = carregar_usuarios()

st.sidebar.markdown("<h4 style='margin-top:0px;'>🔑 Acesso ao Sistema</h4>", unsafe_allow_html=True)
if not usuarios_db:
    st.sidebar.warning("Crie o primeiro usuário na aba Gestão de Usuários.")
    usuario_logado = {"nome": "Admin Padrão", "perfil": "Visão Total Tesouraria", "id": None}
else:
    opcoes_login = {u['nome']: u for u in usuarios_db}
    nome_logado = st.sidebar.selectbox("Simular acesso como:", list(opcoes_login.keys()))
    usuario_logado = opcoes_login[nome_logado]

st.session_state["usuario_logado"] = usuario_logado
perfil_ativo = usuario_logado.get("perfil", "Visão Total Tesouraria")

if "page" not in st.session_state:
    st.session_state.page = "Resumo do Dia" if perfil_ativo == "Visão Total Tesouraria" else ("Visão Consolidada" if perfil_ativo == "Visão Conselho" else "Painel de Eventos")

st.markdown("""
    <style>
    [data-testid="stSidebar"] {
        background-color: #F8FAFC !important;
        border-right: 1px solid #E2E8F0 !important;
    }
    [data-testid="stSidebar"] .stButton button {
        width: 100%; text-align: left; justify-content: flex-start;
        border-radius: 8px; padding: 0.5rem 1rem; font-weight: 600;
        margin-bottom: 4px; transition: all 0.2s ease;
    }
    </style>
""", unsafe_allow_html=True)

def secao(nome):
    st.sidebar.markdown(f"<p style='color:#94A3B8;font-size:0.68rem;font-weight:700;letter-spacing:0.08em;margin:10px 0 2px 4px;'>{nome}</p>", unsafe_allow_html=True)

def nav_button(label, icon):
    ativo = st.session_state.page == label
    if st.sidebar.button(f"{icon}  {label}", key=f"nav_{label}", use_container_width=True, type="primary" if ativo else "secondary"):
        st.session_state.page = label
        st.rerun()

st.sidebar.markdown("<hr style='margin: 8px 0; border-color: #E2E8F0;'>", unsafe_allow_html=True)

if perfil_ativo == "Visão Total Tesouraria":
    secao("OPERACIONAL")
    nav_button("Resumo do Dia", "🏠")
    nav_button("Tesouraria", "💰")
    nav_button("Conciliação Bancária", "🏦")
    nav_button("Visão Consolidada", "📋")
    nav_button("Categorias", "🏷️")

    secao("ESTRATÉGICO")
    nav_button("Metas e Orçamentos", "🎯")

    secao("RELATÓRIOS")
    nav_button("Analytics Financeiro", "📊")
    nav_button("Exportar Contabilidade", "📥")

if perfil_ativo in ["Visão Total Tesouraria", "Visão Conselho"]:
    if perfil_ativo == "Visão Conselho":
        secao("VISÃO GERAL")
        nav_button("Visão Consolidada", "📋")
        nav_button("Analytics Financeiro", "📊")

if perfil_ativo in ["Visão Total Tesouraria", "Visão Eventos"]:
    st.sidebar.markdown("<hr style='margin: 8px 0; border-color: #E2E8F0;'>", unsafe_allow_html=True)
    secao("GESTÃO DE EVENTOS")
    nav_button("Painel de Eventos", "🎫")
    nav_button("Inscrições e Comprovantes", "✅")

if perfil_ativo == "Visão Total Tesouraria":
    st.sidebar.markdown("<hr style='margin: 8px 0 4px 0; border-color: #E2E8F0;'>", unsafe_allow_html=True)
    secao("ADMINISTRAÇÃO")
    nav_button("Gestão de Usuários", "👥")

page = st.session_state.page

# ==========================================
# ==========================================
# ==========================================
# RESUMO DO DIA 
# ==========================================
page = st.session_state.page

# ==========================================
# RESUMO DO DIA 
# ==========================================
if page == "Resumo do Dia":
    st.markdown("""
        <style>
        div.stMainBlockContainer, div[data-testid="stVerticalBlock"] {
            padding-top: 0rem !important;
        }
        [data-testid="stHorizontalBlock"] {
            align-items: stretch;
        }
        [data-testid="stHorizontalBlock"] > [data-testid="column"] {
            display: flex;
            flex-direction: column;
        }
        [data-testid="stHorizontalBlock"] > [data-testid="column"] > div {
            display: flex;
            flex-direction: column;
            height: 100%;
        }
        [data-testid="stHorizontalBlock"] > [data-testid="column"] > div > div[data-testid="stVerticalBlock"] {
            display: flex;
            flex-direction: column;
            height: 100%;
        }
        </style>
    """, unsafe_allow_html=True)

    st.title("Resumo do Dia")
    st.markdown(f"Hoje é {date.today().strftime('%d/%m/%Y')}. Aqui está o que precisa da sua atenção.")

    df = carregar_lancamentos_df()
    inscricoes_resumo = carregar("inscricoes")
    pagamentos_resumo = carregar("inscricao_pagamentos")
    eventos_resumo = carregar("eventos")

    hoje = pd.Timestamp(date.today())
    
    saldos_por_conta = []
    saldo_consolidado = 0.0

    for conta in contas_bancarias_db:
        c_id = conta.get("id")
        c_nome = conta.get("nome", "Conta")
        c_saldo_ini = float(conta.get("saldo_inicial") or 0)
        
        if not df.empty and "conta_bancaria_id" in df.columns and "status" in df.columns:
            l_conta = df[(df["conta_bancaria_id"] == c_id) & (df["status"] == "Concluído")]
            ent = l_conta[l_conta["tipo"] == "Entrada"]["valor"].sum()
            sai = l_conta[l_conta["tipo"] == "Saída"]["valor"].sum()
            c_saldo_atual = c_saldo_ini + ent - sai
        else:
            c_saldo_atual = c_saldo_ini

        saldos_por_conta.append({"nome": c_nome, "saldo": c_saldo_atual})
        saldo_consolidado += c_saldo_atual

    contas_pendentes = df[df["status"] == "Pendente"].copy() if not df.empty else pd.DataFrame()
    contas_hoje = contas_pendentes[contas_pendentes["data_vencimento"] == hoje].copy() if not contas_pendentes.empty and "data_vencimento" in contas_pendentes.columns else pd.DataFrame()
    contas_atrasadas = contas_pendentes[contas_pendentes["data_vencimento"] < hoje].copy() if not contas_pendentes.empty and "data_vencimento" in contas_pendentes.columns else pd.DataFrame()

    recorrencias_expirando = pd.DataFrame()
    if not df.empty and 'recorrente' in df.columns and 'data_fim_recorrencia' in df.columns:
        df['data_fim_recorrencia'] = pd.to_datetime(df['data_fim_recorrencia'], errors='coerce')
        mes_atual_num = hoje.month
        ano_atual_num = hoje.year
        recorrencias_expirando = df[
            (df['recorrente'] == True) & 
            (df['data_fim_recorrencia'].dt.month == mes_atual_num) & 
            (df['data_fim_recorrencia'].dt.year == ano_atual_num)
        ]

    pagamentos_pendentes = [p for p in pagamentos_resumo if p.get("status") == "Pendente"]
    inscricoes_por_id = {str(i.get("id")): i for i in inscricoes_resumo}
    eventos_por_id = {str(e.get("id")): e for e in eventos_resumo}

    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        with st.container(border=True):
            st.caption("Saldo Consolidado")
            st.markdown(f"<h2 style='margin:0; font-size: 1.5rem; color: #059669;'>{fmt_moeda(saldo_consolidado)}</h2>", unsafe_allow_html=True)
            st.markdown("<hr style='margin: 8px 0; border: none; border-top: 1px solid #E2E8F0;'>", unsafe_allow_html=True)
            for sc in saldos_por_conta:
                st.markdown(f"<p style='margin: 2px 0; font-size: 0.8rem; color: #475569;'>• {sc['nome']}: <b>{fmt_moeda(sc['saldo'])}</b></p>", unsafe_allow_html=True)
                
    with col2:
        with st.container(border=True):
            st.caption("A Pagar Hoje")
            st.markdown(f"<h2 style='margin:0; font-size: 1.8rem;'>{len(contas_hoje)}</h2>", unsafe_allow_html=True)
            
    with col3:
        with st.container(border=True):
            st.caption("Contas Atrasadas")
            st.markdown(f"<h2 style='margin:0; font-size: 1.8rem;'>{len(contas_atrasadas)}</h2>", unsafe_allow_html=True)
            
    with col4:
        with st.container(border=True):
            st.caption("Recorrências Expirando")
            st.markdown(f"<h2 style='margin:0; font-size: 1.8rem;'>{len(recorrencias_expirando)}</h2>", unsafe_allow_html=True)

    if not recorrencias_expirando.empty:
        st.warning(f"⚠️ Atenção: Existem **{len(recorrencias_expirando)}** lançamentos recorrentes com data final de recorrência programada para este mês de {MESES_PT[hoje.month-1].lower()}. Verifique a necessidade de renovação.")

    st.markdown("---")
    
    # ----------------------------------------------------
    # BLOCO NOVO: APROVAÇÃO DE CRÉDITOS MANUAIS (CENTAVOS)
    # ----------------------------------------------------
    solicitacoes_pendentes_credito = sb_request("creditos_ofx", "GET", filtros={"status": "eq.Pendente Aprovação"}) or []
    if solicitacoes_pendentes_credito:
        st.subheader("🔔 Solicitações de Liberação de Crédito para Eventos")
        st.info("Líderes de eventos informaram que as pessoas abaixo pagaram sem os centavos de identificação. Verifique se o valor entrou em sua conta e aprove para liberar o crédito no Bolsão do evento.")
        for solic in solicitacoes_pendentes_credito:
            ev_nome = eventos_por_id.get(str(solic.get('evento_id')), {}).get('nome', 'Evento Desconhecido')
            with st.container(border=True):
                col_s1, col_s2, col_s3 = st.columns([3, 1, 1])
                col_s1.write(f"**{ev_nome}** — {solic.get('descricao_bancaria')}")
                col_s1.caption(f"Data informada do pagamento: {pd.to_datetime(solic['data']).strftime('%d/%m/%Y')} | Após aprovar aqui, lembre-se de ir na Tesouraria e mudar a categoria do recebimento original para Inscrição de Evento.")
                col_s2.write(f"**{fmt_moeda(solic['valor'])}**")
                
                b_ap, b_rej = col_s3.columns(2)
                if b_ap.button("✅ Aprovar", key=f"aprov_solic_{solic['id']}", help="Liberar no Bolsão"):
                    sb_request("creditos_ofx", "PATCH", {"status": "Disponível"}, filtros={"id": f"eq.{solic['id']}"})
                    st.cache_data.clear()
                    st.success("Aprovado! Crédito liberado no bolsão do evento.")
                    st.rerun()
                if b_rej.button("❌ Rejeitar", key=f"rej_solic_{solic['id']}"):
                    sb_request("creditos_ofx", "PATCH", {"status": "Rejeitado"}, filtros={"id": f"eq.{solic['id']}"})
                    st.cache_data.clear()
                    st.rerun()
        st.markdown("---")
    # ----------------------------------------------------

    coluna_contas, coluna_comprovantes = st.columns(2)

    with coluna_contas:
        st.subheader("💳 Contas para pagar")
        listas_contas = []
        if not contas_atrasadas.empty:
            contas_atrasadas["ordem_resumo"] = 1
            listas_contas.append(contas_atrasadas)
        if not contas_hoje.empty:
            contas_hoje["ordem_resumo"] = 2
            listas_contas.append(contas_hoje)

        contas_para_mostrar = pd.concat(listas_contas, ignore_index=True).sort_values(by=["ordem_resumo", "data_vencimento"]) if listas_contas else pd.DataFrame()

        if contas_para_mostrar.empty:
            st.success("Nenhuma conta atrasada ou com vencimento hoje.")
        else:
            for _, lancamento in contas_para_mostrar.iterrows():
                vencimento = lancamento.get("data_vencimento")
                vencimento_formatado = vencimento.strftime("%d/%m/%Y") if pd.notna(vencimento) else "Sem vencimento"
                situacao = "🔴 Atrasada" if pd.notna(vencimento) and vencimento < hoje else "🟡 Vence hoje"

                linha1, linha2, linha3 = st.columns([3, 1.5, 1.3])
                with linha1:
                    st.write(f"**{lancamento.get('descricao', 'Sem descrição')}**")
                    st.caption(f"{situacao} • {vencimento_formatado}")
                with linha2:
                    st.write(fmt_moeda(lancamento.get("valor")))
                with linha3:
                    if st.button("🔍 Detalhes", key=f"resumo_detalhes_{lancamento['id']}", use_container_width=True):
                        st.session_state.page = "Tesouraria"
                        st.session_state["tesouraria_tab"] = "⏳ Contas a Pagar/Receber"
                        st.rerun()
                st.markdown("<hr style='margin:6px 0;border:none;border-top:1px solid #E2E8F0;'>", unsafe_allow_html=True)

    with coluna_comprovantes:
        st.subheader("🎫 Comprovantes aguardando aprovação")
        if not pagamentos_pendentes:
            st.success("Nenhum comprovante pendente.")
        else:
            pagamentos_ordenados = sorted(pagamentos_pendentes, key=lambda p: p.get("data_envio", ""))
            for pagamento in pagamentos_ordenados[:6]:
                insc = inscricoes_por_id.get(str(pagamento.get("inscricao_id")), {})
                evento = eventos_por_id.get(str(insc.get("evento_id")), {})
                st.write(f"**{insc.get('nome_participante', 'Participante')}** - Parcela {pagamento.get('numero_parcela', 1)}")
                st.caption(f"{evento.get('nome', 'Evento')} • {fmt_moeda(pagamento.get('valor'))}")
                st.markdown("<hr style='margin:6px 0;border:none;border-top:1px solid #E2E8F0;'>", unsafe_allow_html=True)
            if len(pagamentos_pendentes) > 6:
                st.caption(f"Existem mais {len(pagamentos_pendentes) - 6} comprovantes aguardando validação.")
            if st.button("Ir para Inscrições e Comprovantes →", key="resumo_ir_comprovantes", use_container_width=True):
                st.session_state.page = "Inscrições e Comprovantes"
                st.rerun()
# ==========================================
# VISÃO CONSOLIDADA
# ==========================================
elif page == "Visão Consolidada":
    st.title("📋 Visão Consolidada")
    st.markdown("Acompanhe o balanço mensal de Entradas e Saídas consolidado por categoria.")

    df = carregar_lancamentos_df()
    ano_atual = date.today().year
    
    anos_disp = sorted(df['data_competencia'].dt.year.dropna().unique().astype(int), reverse=True) if not df.empty else [ano_atual]
    if ano_atual not in anos_disp:
        anos_disp.append(ano_atual)
        anos_disp = sorted(anos_disp, reverse=True)

    col_a1, col_a2 = st.columns([3, 1])
    ano_sel = col_a1.selectbox("Ano de Referência", anos_disp)

    df_ano = df[df['data_competencia'].dt.year == ano_sel].copy() if not df.empty else pd.DataFrame()

    meses_curtos = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"]

    def fmt_inteiro_moeda(val):
        if pd.isna(val):
            return "R$ 0"
        try:
            v = float(val)
            return f"R$ {int(round(v)):,.0f}".replace(",", "X").replace(".", ",").replace("X", ".")
        except:
            return str(val)

    pivot_ent = pd.DataFrame()
    pivot_sai = pd.DataFrame()
    resumo_geral = pd.DataFrame(index=MESES_PT)

    cats_entrada = [c['nome'] for c in categorias_db if c['tipo'] == 'Entrada']
    if not df_ano.empty:
        df_ent = df_ano[(df_ano['tipo'] == 'Entrada') & (df_ano['status'] == 'Concluído')]
        if not df_ent.empty:
            pivot_ent = pd.pivot_table(df_ent, values='valor', index='categoria_nome', columns='mes_num', aggfunc='sum', fill_value=0.0)

    for m in range(1, 13):
        if m not in pivot_ent.columns:
            pivot_ent[m] = 0.0

    for cat in cats_entrada:
        if cat not in pivot_ent.index:
            pivot_ent.loc[cat] = 0.0

    pivot_ent = pivot_ent[[m for m in range(1, 13)]]
    pivot_ent.columns = meses_curtos
    pivot_ent['TOTAL ANUAL'] = pivot_ent.sum(axis=1)
    pivot_ent.loc['TOTAL'] = pivot_ent.sum(numeric_only=True)

    cats_saida = [c['nome'] for c in categorias_db if c['tipo'] == 'Saída']
    if not df_ano.empty:
        df_sai = df_ano[(df_ano['tipo'] == 'Saída') & (df_ano['status'] == 'Concluído')]
        if not df_sai.empty:
            pivot_sai = pd.pivot_table(df_sai, values='valor', index='categoria_nome', columns='mes_num', aggfunc='sum', fill_value=0.0)

    for m in range(1, 13):
        if m not in pivot_sai.columns:
            pivot_sai[m] = 0.0

    for cat in cats_saida:
        if cat not in pivot_sai.index:
            pivot_sai.loc[cat] = 0.0

    pivot_sai = pivot_sai[[m for m in range(1, 13)]]
    pivot_sai.columns = meses_curtos
    pivot_sai['TOTAL ANUAL'] = pivot_sai.sum(axis=1)
    pivot_sai.loc['TOTAL'] = pivot_sai.sum(numeric_only=True)

    tot_ent = []
    tot_sai = []
    for m_idx, m_nome in enumerate(MESES_PT, 1):
        val_e = df_ano[(df_ano['mes_num'] == m_idx) & (df_ano['tipo'] == 'Entrada') & (df_ano['status'] == 'Concluído')]['valor'].sum() if not df_ano.empty else 0.0
        val_s = df_ano[(df_ano['mes_num'] == m_idx) & (df_ano['tipo'] == 'Saída') & (df_ano['status'] == 'Concluído')]['valor'].sum() if not df_ano.empty else 0.0
        tot_ent.append(val_e)
        tot_sai.append(val_s)
        
    resumo_geral['Total Entradas'] = tot_ent
    resumo_geral['Total Saídas'] = tot_sai
    resumo_geral['Resultado do Mês'] = resumo_geral['Total Entradas'] - resumo_geral['Total Saídas']
    
    saldo_acum = sum(float(c.get("saldo_inicial") or 0) for c in contas_bancarias_db)
    saldos_mes = []
    for res in resumo_geral['Resultado do Mês']:
        saldo_acum += res
        saldos_mes.append(saldo_acum)
    resumo_geral['Saldo em Caixa Acumulado'] = saldos_mes

    with col_a2:
        st.markdown("<div style='height: 28px;'></div>", unsafe_allow_html=True)
        if not df_ano.empty:
            dfs_export = {
                "Entradas": pivot_ent,
                "Saídas": pivot_sai,
                "Resumo Geral": resumo_geral
            }
            excel_completo = to_excel_bytes(dfs_export)
            st.download_button(
                label="📥 Exportar Excel",
                data=excel_completo,
                file_name=f"Consolidado_{ano_sel}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="btn_export_topo"
            )

    tab_ex1, tab_ex2, tab_ex3 = st.tabs(["📥 Entradas por Categoria", "📤 Saídas por Categoria", "📊 Consolidado Bancário e Geral"])

    with tab_ex1:
        pivot_ent_fmt = pivot_ent.copy()
        for col in pivot_ent_fmt.columns:
            pivot_ent_fmt[col] = pivot_ent_fmt[col].apply(fmt_inteiro_moeda)
        
        altura_ent = (len(pivot_ent_fmt) + 1) * 35 + 38
        st.dataframe(pivot_ent_fmt, use_container_width=True, height=altura_ent)

    with tab_ex2:
        pivot_sai_fmt = pivot_sai.copy()
        for col in pivot_sai_fmt.columns:
            pivot_sai_fmt[col] = pivot_sai_fmt[col].apply(fmt_inteiro_moeda)

        altura_sai = (len(pivot_sai_fmt) + 1) * 35 + 38
        st.dataframe(pivot_sai_fmt, use_container_width=True, height=altura_sai)

    with tab_ex3:
        resumo_geral_fmt = resumo_geral.copy()
        for col in resumo_geral_fmt.columns:
            resumo_geral_fmt[col] = resumo_geral_fmt[col].apply(fmt_inteiro_moeda)

        altura_geral = (len(resumo_geral_fmt) + 1) * 35 + 38
        st.dataframe(resumo_geral_fmt, use_container_width=True, height=altura_geral)

    st.markdown("---")
    st.markdown("### 🔍 Detalhar Valores por Mês e Categoria")
    st.markdown("Selecione os filtros abaixo para ver detalhadamente quais itens compõem a soma e os comprovantes anexados (Modo Somente Leitura).")

    st.markdown("""
        <style>
            .drill-row {
                font-size: 13px !important;
                padding: 4px 0px !important;
                white-space: nowrap !important;
                overflow: hidden !important;
                text-overflow: ellipsis !important;
            }
            div.stButton > button {
                padding: 2px 10px !important;
                font-size: 12px !important;
                min-height: 24px !important;
                height: 28px !important;
                line-height: 1 !important;
                margin-top: 2px !important;
            }
        </style>
    """, unsafe_allow_html=True)

    col_d1, col_d2, col_d3 = st.columns(3)
    mes_drill = col_d1.selectbox("Selecione o Mês", ["Todos"] + MESES_PT)
    tipo_drill = col_d2.selectbox("Selecione o Tipo", ["Todos", "Entrada", "Saída"])
    
    opcoes_cat_drill = ["Todas"] + [c['nome'] for c in categorias_db]
    cat_drill = col_d3.selectbox("Selecione a Categoria", opcoes_cat_drill)

    if not df_ano.empty:
        df_detalhe = df_ano[df_ano['status'] == 'Concluído'].copy()
        
        if mes_drill != "Todos":
            df_detalhe = df_detalhe[df_detalhe['mes_num'] == (MESES_PT.index(mes_drill) + 1)]
        if tipo_drill != "Todos":
            df_detalhe = df_detalhe[df_detalhe['tipo'] == tipo_drill]
        if cat_drill != "Todas":
            df_detalhe = df_detalhe[df_detalhe['categoria_nome'] == cat_drill]

        if df_detalhe.empty:
            st.info("Não há transações concluídas para os filtros informados.")
        else:
            st.markdown("#### Lista de Lançamentos e Comprovantes")
            
            todos_anexos = sb_request("lancamento_anexos", "GET")
            mapa_anexos = {}
            if todos_anexos:
                for anexo in todos_anexos:
                    l_id = anexo['lancamento_id']
                    if l_id not in mapa_anexos:
                        mapa_anexos[l_id] = []
                    mapa_anexos[l_id].append(anexo)

            df_detalhe['ordem_tipo'] = df_detalhe['tipo'].map({'Saída': 0, 'Entrada': 1})
            df_detalhe = df_detalhe.sort_values(by=['ordem_tipo', 'data_competencia'], ascending=[True, False])
            
            header_cols = st.columns([1, 1.2, 2.5, 2, 1.3, 1.2, 1.2])
            header_cols[0].markdown("**Tipo**")
            header_cols[1].markdown("**Data**")
            header_cols[2].markdown("**Descrição**")
            header_cols[3].markdown("**Categoria**")
            header_cols[4].markdown("**Valor**")
            header_cols[5].markdown("**Conta**")
            header_cols[6].markdown("**Documento**")
            st.markdown("<hr style='margin:4px 0;border-color:#CBD5E1;'>", unsafe_allow_html=True)

            for _, row in df_detalhe.iterrows():
                row_id = str(row['id'])
                cols = st.columns([1, 1.2, 2.5, 2, 1.3, 1.2, 1.2])
                
                cols[0].markdown(f"<div class='drill-row'>{row['tipo']}</div>", unsafe_allow_html=True)
                cols[1].markdown(f"<div class='drill-row'>{row['data_competencia'].strftime('%d/%m/%Y')}</div>", unsafe_allow_html=True)
                cols[2].markdown(f"<div class='drill-row'>{row['descricao'] or '—'}</div>", unsafe_allow_html=True)
                cols[3].markdown(f"<div class='drill-row'>{row['categoria_nome']}</div>", unsafe_allow_html=True)
                cols[4].markdown(f"<div class='drill-row'>{fmt_moeda(row['valor'])}</div>", unsafe_allow_html=True)
                cols[5].markdown(f"<div class='drill-row'>{row['conta_nome']}</div>", unsafe_allow_html=True)
                
                anexos_deste = mapa_anexos.get(row_id, [])
                if not anexos_deste and row.get('url_anexo') and isinstance(row.get('url_anexo'), str) and row.get('url_anexo').strip():
                    anexos_deste = [{
                        'id': 'legacy',
                        'url_storage': row.get('url_anexo'),
                        'nome_original': row.get('url_anexo').split('/')[-1]
                    }]

                if len(anexos_deste) > 0:
                    if cols[6].button(f"📎 {len(anexos_deste)} anexo(s)", key=f"btn_vis_anexos_{row_id}", help="Visualizar anexos"):
                        st.session_state[f"show_vis_anexo_{row_id}"] = not st.session_state.get(f"show_vis_anexo_{row_id}", False)
                        st.rerun()
                else:
                    if row['tipo'] == 'Saída':
                        cols[6].markdown('<span class="badge-falta-anexo">🔴 Falta anexo</span>', unsafe_allow_html=True)
                    else:
                        cols[6].markdown("<div class='drill-row' style='color: #64748B;'>—</div>", unsafe_allow_html=True)

                if st.session_state.get(f"show_vis_anexo_{row_id}", False):
                    with st.container(border=True):
                        st.markdown(f"**Documentos Anexados - {row['descricao']}**")
                        for anexo in anexos_deste:
                            link_dl = obter_link_arquivo(anexo['url_storage'])
                            nome_ex = anexo.get('nome_original') or anexo['url_storage'].split('/')[-1]
                            
                            if link_dl:
                                st.markdown(f"📄 [{nome_ex}]({link_dl})", unsafe_allow_html=True)
                            else:
                                st.write(nome_ex)
                                
                        st.markdown("")
                        if st.button("Fechar", key=f"close_vis_{row_id}", use_container_width=True):
                            st.session_state[f"show_vis_anexo_{row_id}"] = False
                            st.rerun()
                
                st.markdown("<hr style='margin:2px 0;border-color:#F1F5F9;'>", unsafe_allow_html=True)

# ==========================================
# TESOURARIA
# ==========================================
elif page == "Tesouraria":
    st.title("Tesouraria")
    
    if "tesouraria_tab" in st.session_state:
        st.session_state["tesouraria_active_tab"] = st.session_state.pop("tesouraria_tab")

    tab1, tab2, tab3, tab4 = st.tabs([
        "📝 Novo Lançamento", 
        "⏳ Contas a Pagar/Receber", 
        "📜 Lançamentos", 
        "🔁 Regras Recorrentes"
    ], key="tesouraria_active_tab", on_change="rerun")

    with tab1:
        st.markdown("### Registrar Movimentação")
        
        tipo_lanc = st.radio("Tipo", ["Entrada", "Saída"], horizontal=True, key="novo_lanc_tipo")
        recorrente = st.checkbox("🔁 Este é um lançamento recorrente (despesa ou receita fixa mensal)", key="novo_lanc_rec")

        with st.form("form_novo_lancamento", clear_on_submit=True):
            if recorrente:
                st.info("💡 Modo Recorrente Ativo: Defina o valor base, o dia fixo de vencimento e até quando essa regra se repete.")
                
                col_r1, col_r2, col_r3 = st.columns(3)
                valor = col_r1.number_input("Valor Base (R$)", min_value=0.0, step=50.0, format="%.2f", key="rec_valor")
                dia_vencimento = col_r2.number_input("Dia Fixo de Vencimento", min_value=1, max_value=31, value=10, step=1, key="rec_dia")
                data_fim_rec = col_r3.date_input("Data Final da Recorrência", date.today() + pd.DateOffset(months=12), format="DD.MM.YYYY", key="rec_fim")
                
                cats_filtradas = [c for c in categorias_db if c.get("tipo") == tipo_lanc]
                opcoes_cats = {c["nome"]: c["id"] for c in cats_filtradas}

                col_r4, col_r5 = st.columns([2, 1])
                descricao = col_r4.text_input("Descrição da Recorrência (Ex: Aluguel do Templo)", key="rec_desc")
                categoria_sel = col_r5.selectbox("Categoria", list(opcoes_cats.keys()) if opcoes_cats else ["Cadastre uma categoria"], key="rec_cat")

                contas_opcoes = {"Nenhuma": None} | {c["nome"]: c["id"] for c in contas_bancarias_db}
                conta_sel = st.selectbox("Conta Bancária Principal", list(contas_opcoes.keys()), key="rec_conta")
                tag = st.selectbox("Projeto / Evento", ["Nenhum"] + [e['nome'] for e in eventos_db], key="rec_tag")
                
                data_comp = date.today()
                status_lanc = "Pendente"
                data_venc = date.today()

            else:
                col1, col2, col3 = st.columns(3)
                valor = col1.number_input("Valor (R$)", min_value=0.0, step=50.0, format="%.2f", key="unico_valor")
                
                if tipo_lanc == "Saída":
                    st.caption("📅 Mês de Competência Contábil (Referência MM.YYYY)")
                    cc_m, cc_a = col2.columns(2)
                    mes_comp_sel = cc_m.selectbox("Mês", MESES_PT, index=date.today().month-1, key="unico_mes_comp")
                    ano_comp_sel = cc_a.number_input("Ano", min_value=2020, max_value=2100, value=date.today().year, step=1, key="unico_ano_comp")
                    
                    idx_mes = MESES_PT.index(mes_comp_sel) + 1
                    data_comp = date(int(ano_comp_sel), idx_mes, 1)
                    
                    status_lanc = col3.selectbox("Situação", ["Concluído", "Pendente"], key="unico_status")
                    label_data = "Data de Pagamento" if status_lanc == "Concluído" else "Data de Vencimento"
                    data_venc = st.date_input(label_data, date.today(), format="DD.MM.YYYY", key="unico_venc")
                else:
                    data_comp = col2.date_input("Data de Recebimento", date.today(), format="DD.MM.YYYY", key="unico_data_ent")
                    status_lanc = col3.selectbox("Situação", ["Concluído", "Pendente"], key="unico_status_ent")
                    data_venc = data_comp

                cats_filtradas = [c for c in categorias_db if c.get("tipo") == tipo_lanc]
                opcoes_cats = {c["nome"]: c["id"] for c in cats_filtradas}

                col4, col5 = st.columns([2, 1])
                descricao = col4.text_input("Descrição", key="unico_desc")
                categoria_sel = col5.selectbox("Categoria", list(opcoes_cats.keys()) if opcoes_cats else ["Cadastre uma categoria"], key="unico_cat")

                contas_opcoes = {"Nenhuma": None} | {c["nome"]: c["id"] for c in contas_bancarias_db}
                col6, col7 = st.columns(2)
                conta_sel = col6.selectbox("Conta Bancária", list(contas_opcoes.keys()), key="unico_conta")
                tag = col7.selectbox("Projeto / Evento", ["Nenhum"] + [e['nome'] for e in eventos_db], key="unico_tag")

            arquivos = st.file_uploader(
                "Comprovantes / Notas Fiscais (Permite múltiplos arquivos)", 
                type=['png', 'jpg', 'jpeg', 'pdf'], 
                accept_multiple_files=True, 
                key="lanc_arq_up_multiplo"
            )

            if st.form_submit_button("💾 Salvar Lançamento", use_container_width=True, type="primary"):
                if valor <= 0 or not descricao or not opcoes_cats:
                    st.warning("⚠️ Preencha descrição, valor e categoria corretamente.")
                else:
                    with st.spinner("Salvando lançamento e anexos..."):
                        payload_lanc = {
                            "descricao": descricao,
                            "tipo": tipo_lanc,
                            "valor": float(valor),
                            "data_competencia": str(data_comp),
                            "data_vencimento": str(data_venc),
                            "status": status_lanc,
                            "categoria_id": opcoes_cats[categoria_sel],
                            "conta_bancaria_id": contas_opcoes.get(conta_sel),
                            "centro_custo": None if tag == "Nenhum" else tag,
                            "recorrente": bool(recorrente)
                        }
                        
                        if recorrente:
                            payload_lanc["dia_vencimento_fixo"] = int(dia_vencimento)
                            payload_lanc["data_fim_recorrencia"] = str(data_fim_rec)

                        if status_lanc == "Concluído":
                            payload_lanc["data_pagamento"] = str(data_venc)

                        res_lanc = sb_request("lancamentos", "POST", [payload_lanc])
                        
                        if res_lanc and len(res_lanc) > 0:
                            novo_id = res_lanc[0]['id']
                            if arquivos:
                                processar_e_salvar_anexos(arquivos, novo_id, data_comp, categoria_sel, descricao)
                            
                            st.cache_data.clear()
                            st.success("✅ Lançamento e anexos registrados com sucesso!")
                            time.sleep(1)
                            st.rerun()

    with tab2:
        df = carregar_lancamentos_df()
        pend = df[(df['status'] == 'Pendente') & (df['recorrente'] != True)].copy() if not df.empty else pd.DataFrame()
        
        if pend.empty:
            st.info("Nenhuma conta pendente no momento. 🎉")
        else:
            hoje = pd.Timestamp(date.today())
            
            pend_pagar = pend[pend['tipo'] == 'Saída'].sort_values('data_vencimento') if 'tipo' in pend.columns else pd.DataFrame()
            pend_receber = pend[pend['tipo'] == 'Entrada'].sort_values('data_vencimento') if 'tipo' in pend.columns else pd.DataFrame()

            st.subheader("💳 Contas a Pagar")
            if pend_pagar.empty:
                st.success("Nenhuma conta a pagar pendente. 👍")
            else:
                h_pagar = st.columns([3, 1.4, 1.4, 1.4, 1.2])
                h_pagar[0].markdown("**Descrição / Categoria**")
                h_pagar[1].markdown("**Valor**")
                h_pagar[2].markdown("**Vencimento**")
                h_pagar[3].markdown("**Situação**")
                h_pagar[4].markdown("**Ação**")
                st.markdown("<hr style='margin:4px 0;border-color:#CBD5E1;'>", unsafe_allow_html=True)

                for idx, (_, row) in enumerate(pend_pagar.iterrows()):
                    venc = row['data_vencimento']
                    situacao = "🔴 Atrasado" if pd.notna(venc) and venc < hoje else ("🟡 Vence hoje" if venc == hoje else "🟢 A vencer")
                    c1, c2, c3, c4, c5 = st.columns([3, 1.4, 1.4, 1.4, 1.2])
                    c1.write(f"**{row['descricao']}** — {row['categoria_nome']}")
                    c2.write(fmt_moeda(row['valor']))
                    c3.write(venc.strftime('%d.%m.%Y') if pd.notna(venc) else '—')
                    c4.write(situacao)
                    
                    row_id = str(row['id'])
                    session_key_pay = f"paying_{row_id}_{idx}"
                    is_paying = st.session_state.get(session_key_pay, False)
                    
                    if not is_paying:
                        if c5.button("✅ Pagar", key=f"pagar_tab_{row_id}_{idx}"):
                            st.session_state[session_key_pay] = True
                            st.rerun()
                    else:
                        c5.write("Aguardando...")

                    if st.session_state.get(session_key_pay, False):
                        with st.container(border=True):
                            st.markdown(f"📎 **Anexo Obrigatório (Nota Fiscal/Comprovante) para:** {row['descricao']}")
                            arq_pagar_multiplos = st.file_uploader("Selecione os arquivos (Permite múltiplos)", type=['png', 'jpg', 'jpeg', 'pdf'], accept_multiple_files=True, key=f"file_pagar_{row_id}_{idx}")
                            
                            col_b1, col_b2 = st.columns(2)
                            if col_b1.button("💾 Confirmar Pagamento", key=f"conf_pagar_{row_id}_{idx}", type="primary"):
                                if not arq_pagar_multiplos:
                                    st.error("⚠️ O anexo da nota fiscal/comprovante é **mandatório** para contas a pagar.")
                                else:
                                    with st.spinner("Enviando anexos e registrando pagamento..."):
                                        res = sb_request("lancamentos", "PATCH", {
                                            "status": "Concluído", 
                                            "data_pagamento": str(date.today())
                                        }, filtros={"id": f"eq.{row_id}"})
                                        
                                        if res is not None:
                                            processar_e_salvar_anexos(arq_pagar_multiplos, row_id, row['data_competencia'], row['categoria_nome'], row['descricao'])
                                            
                                            st.session_state[session_key_pay] = False
                                            st.cache_data.clear()
                                            st.success("Pagamento registrado com sucesso!")
                                            time.sleep(1)
                                            st.rerun()
                            if col_b2.button("❌ Cancelar", key=f"canc_pagar_{row_id}_{idx}"):
                                st.session_state[session_key_pay] = False
                                st.rerun()

            st.markdown("---")

            st.subheader("💰 Contas a Receber")
            if pend_receber.empty:
                st.success("Nenhuma conta a receber pendente. 👍")
            else:
                h_receber = st.columns([3, 1.4, 1.4, 1.4, 1.2])
                h_receber[0].markdown("**Descrição / Categoria**")
                h_receber[1].markdown("**Valor**")
                h_receber[2].markdown("**Vencimento**")
                h_receber[3].markdown("**Situação**")
                h_receber[4].markdown("**Ação**")
                st.markdown("<hr style='margin:4px 0;border-color:#CBD5E1;'>", unsafe_allow_html=True)

                for idx, (_, row) in enumerate(pend_receber.iterrows()):
                    venc = row['data_vencimento']
                    situacao = "🔴 Atrasado" if pd.notna(venc) and venc < hoje else ("🟡 Vence hoje" if venc == hoje else "🟢 A vencer")
                    c1, c2, c3, c4, c5 = st.columns([3, 1.4, 1.4, 1.4, 1.2])
                    c1.write(f"**{row['descricao']}** — {row['categoria_nome']}")
                    c2.write(fmt_moeda(row['valor']))
                    c3.write(venc.strftime('%d.%m.%Y') if pd.notna(venc) else '—')
                    c4.write(situacao)
                    
                    row_id = str(row['id'])
                    session_key_rec = f"receiving_{row_id}_{idx}"
                    is_receiving = st.session_state.get(session_key_rec, False)
                    
                    if not is_receiving:
                        if c5.button("✅ Receber", key=f"receber_tab_{row_id}_{idx}"):
                            st.session_state[session_key_rec] = True
                            st.rerun()
                    else:
                        c5.write("Aguardando...")

                    if st.session_state.get(session_key_rec, False):
                        with st.container(border=True):
                            st.markdown(f"📎 **Anexo Opcional para:** {row['descricao']}")
                            arq_receber_multiplos = st.file_uploader("Selecione os arquivos se desejar anexar (Opcional)", type=['png', 'jpg', 'jpeg', 'pdf'], accept_multiple_files=True, key=f"file_receber_{row_id}_{idx}")
                            
                            col_b1, col_b2 = st.columns(2)
                            if col_b1.button("💾 Confirmar Recebimento", key=f"conf_receber_{row_id}_{idx}", type="primary"):
                                with st.spinner("Registrando recebimento..."):
                                    res = sb_request("lancamentos", "PATCH", {
                                        "status": "Concluído", 
                                        "data_pagamento": str(date.today())
                                    }, filtros={"id": f"eq.{row_id}"})
                                    
                                    if res is not None:
                                        if arq_receber_multiplos:
                                            processar_e_salvar_anexos(arq_receber_multiplos, row_id, row['data_competencia'], row['categoria_nome'], row['descricao'])
                                            
                                        st.session_state[session_key_rec] = False
                                        st.cache_data.clear()
                                        st.success("Recebimento registrado com sucesso!")
                                        time.sleep(1)
                                        st.rerun()
                            if col_b2.button("❌ Cancelar", key=f"canc_receber_{row_id}_{idx}"):
                                st.session_state[session_key_rec] = False
                                st.rerun()

    with tab3:
        df = carregar_lancamentos_df()
        if df.empty:
            st.info("Nenhum lançamento registrado.")
        else:
            st.markdown("### Filtros de Lançamentos")
            
            todos_anexos = sb_request("lancamento_anexos", "GET")
            mapa_anexos = {}
            if todos_anexos:
                for anexo in todos_anexos:
                    l_id = anexo['lancamento_id']
                    if l_id not in mapa_anexos:
                        mapa_anexos[l_id] = []
                    mapa_anexos[l_id].append(anexo)

            col_f1, col_f2, col_f3, col_f4, col_f5, col_f6 = st.columns(6)
            start_date = date.today().replace(day=1)
            end_date = (pd.Timestamp.today() + pd.offsets.MonthEnd(1)).date()
            
            data_inicio = col_f1.date_input("De", start_date, format="DD.MM.YYYY", key="hist_dt_ini")
            data_fim = col_f2.date_input("Até", end_date, format="DD.MM.YYYY", key="hist_dt_fim")
            filtro_tipo = col_f3.multiselect("Tipo", ["Entrada", "Saída"], default=["Entrada", "Saída"], key="hist_ft_tipo")
            
            status_unicos = df['status'].unique().tolist() if 'status' in df.columns else []
            filtro_status = col_f4.multiselect("Situação", status_unicos, default=status_unicos, key="hist_ft_status")
            
            categorias_unicas = sorted(df['categoria_nome'].dropna().unique().tolist()) if 'categoria_nome' in df.columns else []
            filtro_categorias = col_f5.multiselect("Categoria", categorias_unicas, default=categorias_unicas, key="hist_ft_cat")
            
            filtro_anexo = col_f6.selectbox("Status do Anexo", ["Todos", "Com Anexo", "Falta Anexo"], key="hist_ft_anexo")
            
            mask_data = (df['data_competencia'].dt.date >= data_inicio) & (df['data_competencia'].dt.date <= data_fim)
            dff = df[
                mask_data & 
                df['tipo'].isin(filtro_tipo) & 
                df['status'].isin(filtro_status) & 
                df['categoria_nome'].isin(filtro_categorias)
            ].copy()

            # Aplicar filtro customizado de anexo
            if filtro_anexo != "Todos":
                ids_filtrados = []
                for _, r in dff.iterrows():
                    r_id = str(r['id'])
                    tem_reg = len(mapa_anexos.get(r_id, [])) > 0 or (isinstance(r.get('url_anexo'), str) and r.get('url_anexo').strip())
                    is_saida = (r['tipo'] == 'Saída')
                    
                    if filtro_anexo == "Com Anexo" and tem_reg:
                        ids_filtrados.append(r['id'])
                    elif filtro_anexo == "Falta Anexo" and is_saida and not tem_reg:
                        ids_filtrados.append(r['id'])
                dff = dff[dff['id'].isin(ids_filtrados)]
            
            if dff.empty:
                st.info("Nenhum lançamento encontrado para os filtros selecionados.")
            else:
                total_entradas_filt = dff[dff['tipo'] == 'Entrada']['valor'].sum()
                total_saidas_filt = dff[dff['tipo'] == 'Saída']['valor'].sum()
                saldo_filtrado = total_entradas_filt - total_saidas_filt

                st.markdown("")
                col_m1, col_m2, col_m3 = st.columns([1.5, 1.5, 3])
                col_m1.metric("Total Filtrado (Líquido)", fmt_moeda(saldo_filtrado))
                col_m2.metric("Qtd. Registros", str(len(dff)))
                st.markdown("---")

                dff['ordem_tipo'] = dff['tipo'].map({'Saída': 0, 'Entrada': 1})
                dff = dff.sort_values(by=['ordem_tipo', 'data_competencia'], ascending=[True, False])
                
                st.markdown("#### Lista de Lançamentos e Comprovantes")

                header_cols = st.columns([1, 1.2, 2.5, 2, 1.3, 1.2, 1.2, 0.8])
                header_cols[0].markdown("**Tipo**")
                header_cols[1].markdown("**Data**")
                header_cols[2].markdown("**Descrição**")
                header_cols[3].markdown("**Categoria**")
                header_cols[4].markdown("**Valor**")
                header_cols[5].markdown("**Conta**")
                header_cols[6].markdown("**Documento**")
                header_cols[7].markdown("**Editar**")
                st.markdown("<hr style='margin:4px 0;border-color:#CBD5E1;'>", unsafe_allow_html=True)

                for _, row in dff.iterrows():
                    row_id = str(row['id'])
                    cols = st.columns([1, 1.2, 2.5, 2, 1.3, 1.2, 1.2, 0.8])
                    
                    cols[0].markdown(f"<div class='drill-row'>{row['tipo']}</div>", unsafe_allow_html=True)
                    cols[1].markdown(f"<div class='drill-row'>{row['data_competencia'].strftime('%d/%m/%Y')}</div>", unsafe_allow_html=True)
                    cols[2].markdown(f"<div class='drill-row'>{row['descricao'] or '—'}</div>", unsafe_allow_html=True)
                    cols[3].markdown(f"<div class='drill-row'>{row['categoria_nome']}</div>", unsafe_allow_html=True)
                    cols[4].markdown(f"<div class='drill-row'>{fmt_moeda(row['valor'])}</div>", unsafe_allow_html=True)
                    cols[5].markdown(f"<div class='drill-row'>{row['conta_nome']}</div>", unsafe_allow_html=True)
                    
                    anexos_deste = mapa_anexos.get(row_id, [])
                    if not anexos_deste and row.get('url_anexo') and isinstance(row.get('url_anexo'), str) and row.get('url_anexo').strip():
                        anexos_deste = [{
                            'id': 'legacy',
                            'url_storage': row.get('url_anexo'),
                            'nome_original': row.get('url_anexo').split('/')[-1]
                        }]

                    if len(anexos_deste) > 0:
                        if cols[6].button(f"📎 {len(anexos_deste)} anexo(s)", key=f"btn_hist_anexos_{row_id}", help="Ver/Gerenciar anexos"):
                            st.session_state[f"show_hist_anexo_{row_id}"] = not st.session_state.get(f"show_hist_anexo_{row_id}", False)
                            st.rerun()
                    else:
                        if row['tipo'] == 'Saída':
                            cols[6].markdown('<span class="badge-falta-anexo">🔴 Falta anexo</span>', unsafe_allow_html=True)
                        else:
                            cols[6].markdown("<div class='drill-row' style='color: #64748B;'>—</div>", unsafe_allow_html=True)

                    if cols[7].button("✏️", key=f"btn_edit_lapis_{row_id}", help="Editar Lançamento"):
                        st.session_state[f"edit_lanc_ativo"] = row_id
                        st.rerun()

                    if st.session_state.get(f"show_hist_anexo_{row_id}", False):
                        with st.container(border=True):
                            st.markdown(f"**Gerenciar Anexos - {row['descricao']}**")
                            for anexo in anexos_deste:
                                col_g1, col_g2 = st.columns([3, 1])
                                link_dl = obter_link_arquivo(anexo['url_storage'])
                                nome_ex = anexo.get('nome_original') or anexo['url_storage'].split('/')[-1]
                                
                                if link_dl:
                                    col_g1.markdown(f"📄 [{nome_ex}]({link_dl})", unsafe_allow_html=True)
                                else:
                                    col_g1.write(nome_ex)
                                    
                                if anexo.get('id') == 'legacy':
                                    if col_g2.button("🗑️ Apagar", key=f"del_leg_hist_{row_id}"):
                                        try:
                                            supabase.storage.from_("comprovantes").remove([anexo['url_storage']])
                                            sb_request("lancamentos", "PATCH", {"url_anexo": None}, filtros={"id": f"eq.{row_id}"})
                                            st.cache_data.clear()
                                            st.success("Anexo excluído!")
                                            st.rerun()
                                        except Exception as e:
                                            st.error(f"Erro: {e}")
                                else:
                                    if col_g2.button("🗑️ Apagar", key=f"del_novo_hist_{anexo['id']}"):
                                        try:
                                            supabase.storage.from_("comprovantes").remove([anexo['url_storage']])
                                            sb_request("lancamento_anexos", "DELETE", filtros={"id": f"eq.{anexo['id']}"})
                                            st.cache_data.clear()
                                            st.success("Anexo excluído!")
                                            st.rerun()
                                        except Exception as e:
                                            st.error(f"Erro: {e}")
                                            
                            if st.button("Fechar Gerenciador", key=f"close_hist_{row_id}"):
                                st.session_state[f"show_hist_anexo_{row_id}"] = False
                                st.rerun()

                    if st.session_state.get(f"edit_lanc_ativo") == row_id:
                        with st.container(border=True):
                            st.markdown(f"#### ✏️ Editando Lançamento: {row.get('descricao', '')}")
                            lanc_raw = next((l for l in carregar("lancamentos") if str(l.get('id')) == row_id), row)
                            
                            with st.form(f"form_ed_lapis_{row_id}"):
                                col_e1, col_e2 = st.columns(2)
                                with col_e1:
                                    n_tipo = st.selectbox("Tipo", ["Entrada", "Saída"], index=0 if lanc_raw.get('tipo', 'Entrada') == "Entrada" else 1, key=f"ed_tipo_{row_id}")
                                    n_desc = st.text_input("Descrição", value=lanc_raw.get('descricao', ''), key=f"ed_desc_{row_id}")
                                    n_valor = st.number_input("Valor (R$)", value=float(lanc_raw.get('valor') or 0), format="%.2f", key=f"ed_val_{row_id}")
                                with col_e2:
                                    raw_data = lanc_raw.get('data_competencia')
                                    try:
                                        default_date = pd.to_datetime(raw_data).date() if raw_data else date.today()
                                    except:
                                        default_date = date.today()
                                    n_data = st.date_input("Data", value=default_date, format="DD/MM/YYYY", key=f"ed_data_{row_id}")
                                    
                                    status_opts = ["Concluído", "Pendente"]
                                    curr_status = lanc_raw.get('status', 'Concluído')
                                    status_idx = status_opts.index(curr_status) if curr_status in status_opts else 0
                                    n_status = st.selectbox("Situação", status_opts, index=status_idx, key=f"ed_status_{row_id}")

                                col_e3, col_e4 = st.columns(2)
                                with col_e3:
                                    cat_nomes = [c.get('nome') for c in categorias_db] if 'categorias_db' in globals() and categorias_db else []
                                    curr_cat = lanc_raw.get('categoria_nome', '')
                                    cat_idx = cat_nomes.index(curr_cat) if curr_cat in cat_nomes else 0
                                    n_cat_nome = st.selectbox("Categoria", cat_nomes if cat_nomes else [curr_cat], index=cat_idx if cat_nomes else 0, key=f"ed_cat_{row_id}")
                                    
                                    selected_cat_obj = next((c for c in categorias_db if c.get('nome') == n_cat_nome), None) if 'categorias_db' in globals() and categorias_db else None
                                    n_cat_id = selected_cat_obj.get('id') if selected_cat_obj else lanc_raw.get('categoria_id')

                                with col_e4:
                                    conta_nomes = [cb.get('nome') for cb in contas_bancarias_db] if 'contas_bancarias_db' in globals() and contas_bancarias_db else []
                                    curr_conta = lanc_raw.get('conta_nome', '')
                                    conta_idx = conta_nomes.index(curr_conta) if curr_conta in conta_nomes else 0
                                    n_conta_nome = st.selectbox("Conta Bancária", conta_nomes if conta_nomes else [curr_conta], index=conta_idx if conta_nomes else 0, key=f"ed_conta_{row_id}")
                                    
                                    selected_conta_obj = next((cb for cb in contas_bancarias_db if cb.get('nome') == n_conta_nome), None) if 'contas_bancarias_db' in globals() and contas_bancarias_db else None
                                    n_conta_id = selected_conta_obj.get('id') if selected_conta_obj else lanc_raw.get('conta_bancaria_id')

                                st.markdown("---")
                                novos_arq_edicao = st.file_uploader("📎 Adicionar Anexos nesta Edição (Opcional)", type=['png', 'jpg', 'jpeg', 'pdf'], accept_multiple_files=True, key=f"ed_up_files_{row_id}")

                                st.markdown("")
                                c_b1, c_b2, c_b3 = st.columns(3)
                                btn_upd = c_b1.form_submit_button("💾 Salvar Alterações", use_container_width=True)
                                btn_del = c_b2.form_submit_button("🗑️ Excluir", use_container_width=True)
                                btn_canc = c_b3.form_submit_button("❌ Fechar", use_container_width=True)
                                
                                if btn_upd:
                                    payload = {
                                        "tipo": n_tipo,
                                        "descricao": n_desc,
                                        "valor": float(n_valor),
                                        "data_competencia": str(n_data),
                                        "status": n_status,
                                        "categoria_id": n_cat_id,
                                        "categoria_nome": n_cat_nome,
                                        "conta_bancaria_id": n_conta_id,
                                        "conta_nome": n_conta_nome
                                    }
                                    sb_request("lancamentos", "PATCH", payload, filtros={"id": f"eq.{row_id}"})
                                    if novos_arq_edicao:
                                        processar_e_salvar_anexos(novos_arq_edicao, row_id, n_data, n_cat_nome, n_desc)
                                        
                                    st.session_state[f"edit_lanc_ativo"] = None
                                    st.cache_data.clear()
                                    st.success("Lançamento atualizado com sucesso!")
                                    time.sleep(1)
                                    st.rerun()
                                if btn_del:
                                    sb_request("lancamentos", "DELETE", filtros={"id": f"eq.{row_id}"})
                                    st.session_state[f"edit_lanc_ativo"] = None
                                    st.cache_data.clear()
                                    st.success("Lançamento excluído com sucesso!")
                                    time.sleep(1)
                                    st.rerun()
                                if btn_canc:
                                    st.session_state[f"edit_lanc_ativo"] = None
                                    st.rerun()
                    
                    st.markdown("<hr style='margin:2px 0;border-color:#F1F5F9;'>", unsafe_allow_html=True)

                st.markdown("")
                exibir_export = dff[['type' if 'type' in dff else 'tipo', 'data_competencia', 'descricao', 'categoria_nome', 'valor', 'conta_nome']].copy()
                exibir_export['data_competencia'] = exibir_export['data_competencia'].dt.strftime('%d/%m/%Y')
                exibir_export.columns = ['Tipo', 'Data', 'Descrição', 'Categoria', 'Valor', 'Conta']
                
                excel_detalhe = to_excel_bytes({"Lancamentos": exibir_export})
                st.download_button(
                    label="💾 Baixar Tabela Selecionada (Excel)",
                    data=excel_detalhe,
                    file_name=f"Lancamentos_Filtro.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="btn_dl_lanc_filtro"
                )

    with tab4:
        st.markdown("### Gerenciamento de Regras Recorrentes (Despesas/Receitas Fixas)")
        df_all = carregar_lancamentos_df()
        recorrencias_ativas = df_all[df_all['recorrente'] == True].copy() if not df_all.empty else pd.DataFrame()
        
        if recorrencias_ativas.empty:
            st.info("Nenhuma regra recorrente cadastrada.")
        else:
            map_cat = {str(c['id']): c['nome'] for c in categorias_db}
            if 'categoria_nome' not in recorrencias_ativas.columns:
                recorrencias_ativas['categoria_nome'] = recorrencias_ativas['categoria_id'].astype(str).map(map_cat)

            recorrencias_ativas['ordem_tipo'] = recorrencias_ativas['tipo'].map({'Saída': 0, 'Entrada': 1})
            recorrencias_ativas['dt_fim_sort'] = pd.to_datetime(recorrencias_ativas['data_fim_recorrencia'], errors='coerce')
            recorrencias_ativas = recorrencias_ativas.sort_values(by=['ordem_tipo', 'dt_fim_sort'], ascending=[True, True])

            st.markdown("---")
            h1, h2, h3, h4, h5, h6, h7 = st.columns([1.2, 2.5, 2, 1.3, 1.3, 1.3, 1.2])
            h1.markdown("**Tipo**")
            h2.markdown("**Descrição**")
            h3.markdown("**Categoria**")
            h4.markdown("**De**")
            h5.markdown("**Até**")
            h6.markdown("**Valor Base**")
            h7.markdown("**Ações**")
            st.markdown("<hr style='margin:4px 0;border:none;border-top:1px solid #CBD5E1;'>", unsafe_allow_html=True)

            for _, rec_row in recorrencias_ativas.iterrows():
                rec_id = str(rec_row['id'])
                dt_inicio = pd.to_datetime(rec_row.get('data_competencia')).strftime('%d.%m.%Y') if pd.notna(rec_row.get('data_competencia')) else '—'
                dt_fim = pd.to_datetime(rec_row.get('data_fim_recorrencia')).strftime('%d.%m.%Y') if pd.notna(rec_row.get('data_fim_recorrencia')) else 'Indeterminado'
                
                rc_tipo, rc_desc, rc_cat, rc_de, rc_ate, rc_val, rc_acoes = st.columns([1.2, 2.5, 2, 1.3, 1.3, 1.3, 1.2])
                
                rc_tipo.write(f"**{rec_row['tipo']}**")
                rc_desc.write(rec_row['descricao'])
                rc_cat.write(rec_row.get('categoria_nome', '—'))
                rc_de.write(dt_inicio)
                rc_ate.write(dt_fim)
                rc_val.write(fmt_moeda(rec_row['valor']))
                
                b_edit, b_del = rc_acoes.columns(2)
                
                if b_edit.button("✏️", key=f"btn_edit_rec_{rec_id}", help="Editar Regra"):
                    st.session_state[f"editing_rec_{rec_id}"] = not st.session_state.get(f"editing_rec_{rec_id}", False)
                    st.rerun()
                    
                if b_del.button("🗑️", key=f"btn_del_rec_{rec_id}", help="Excluir Regra"):
                    sb_request("lancamentos", "DELETE", filtros={"id": f"eq.{rec_id}"})
                    if f"editing_rec_{rec_id}" in st.session_state:
                        del st.session_state[f"editing_rec_{rec_id}"]
                    st.cache_data.clear()
                    st.success("Regra removida!")
                    time.sleep(1)
                    st.rerun()

                if st.session_state.get(f"editing_rec_{rec_id}", False):
                    with st.container():
                        st.markdown(f"---")
                        st.info(f"✏️ **Editando Regra:** {rec_row['descricao']}")
                        with st.form(f"form_edit_rec_{rec_id}"):
                            e_tipo = st.selectbox("Tipo", ["Entrada", "Saída"], index=0 if rec_row['tipo'] == "Entrada" else 1, key=f"ed_rec_tipo_{rec_id}")
                            e_desc = st.text_input("Descrição", value=rec_row['descricao'], key=f"ed_rec_desc_{rec_id}")
                            
                            c_val_dia = st.columns(2)
                            e_valor = c_val_dia[0].number_input("Valor Base (R$)", value=float(rec_row['valor'] or 0), format="%.2f", key=f"ed_rec_val_{rec_id}")
                            e_dia = c_val_dia[1].number_input("Dia Fixo de Vencimento", min_value=1, max_value=31, value=int(rec_row.get('dia_vencimento_fixo') or 10), step=1, key=f"ed_rec_dia_{rec_id}")
                            
                            cats_r_edit = [c for c in categorias_db if c.get("tipo") == e_tipo]
                            opcoes_cats_r = {c["nome"]: c["id"] for c in cats_r_edit}
                            atual_cat_nome = rec_row.get('categoria_nome')
                            idx_c = list(opcoes_cats_r.keys()).index(atual_cat_nome) if atual_cat_nome in opcoes_cats_r else 0
                            
                            e_cat = st.selectbox("Categoria", list(opcoes_cats_r.keys()) if opcoes_cats_r else ["-"], index=idx_c, key=f"ed_rec_cat_{rec_id}")
                            
                            col_d1, col_d2 = st.columns(2)
                            e_dt_ini = col_d1.date_input("Data Inicial (De)", pd.to_datetime(rec_row['data_competencia']).date() if pd.notna(rec_row.get('data_competencia')) else date.today(), format="DD.MM.YYYY", key=f"ed_rec_de_{rec_id}")
                            e_dt_fim = col_d2.date_input("Data Final (Até)", pd.to_datetime(rec_row['data_fim_recorrencia']).date() if pd.notna(rec_row.get('data_fim_recorrencia')) else date.today() + pd.DateOffset(months=12), format="DD.MM.YYYY", key=f"ed_rec_ate_{rec_id}")
                            
                            col_s1, col_s2 = st.columns(2)
                            if col_s1.form_submit_button("💾 Salvar Alterações", use_container_width=True, type="primary"):
                                carga_rec = {
                                    "descricao": e_desc,
                                    "tipo": e_tipo,
                                    "valor": float(e_valor),
                                    "dia_vencimento_fixo": int(e_dia),
                                    "categoria_id": opcoes_cats_r[e_cat] if opcoes_cats_r else None,
                                    "data_competencia": str(e_dt_ini),
                                    "data_vencimento": str(e_dt_ini),
                                    "data_fim_recorrencia": str(e_dt_fim)
                                }
                                res = sb_request("lancamentos", "PATCH", carga_rec, filtros={"id": f"eq.{rec_id}"})
                                if res is not None:
                                    st.session_state[f"editing_rec_{rec_id}"] = False
                                    st.cache_data.clear()
                                    st.success("Regra atualizada com sucesso!")
                                    time.sleep(1)
                                    st.rerun()
                                    
                            if col_s2.form_submit_button("❌ Cancelar", use_container_width=True):
                                st.session_state[f"editing_rec_{rec_id}"] = False
                                st.rerun()
                st.markdown("<hr style='margin:6px 0;border:none;border-top:1px solid #E2E8F0;'>", unsafe_allow_html=True)

# ==========================================
# CONCILIAÇÃO BANCÁRIA
# ==========================================
elif page == "Conciliação Bancária":
    st.title("Conciliação Bancária")
    st.markdown("Confira os lançamentos manualmente ou importe o extrato do banco (OFX) para automatizar o cadastro em lote.")

    col_nova, col_edit = st.columns(2)
    with col_nova:
        with st.expander("➕ Cadastrar Conta Bancária"):
            with st.form("form_conta", clear_on_submit=True):
                nome_conta = st.text_input("Nome da Conta (Ex: Itaú Principal)")
                tipo_conta = st.selectbox("Tipo", ["Corrente", "Poupança", "Investimento"], key="tipo_conta_novo")
                codigo_conta = st.text_input("Código Contábil (Ex: 1.1.10.200.003)")
                saldo_inicial = st.number_input("Saldo Inicial (R$)", min_value=0.0, format="%.2f")
                if st.form_submit_button("Cadastrar Conta", use_container_width=True):
                    if nome_conta:
                        res = sb_request("contas_bancarias", "POST", {"nome": nome_conta, "tipo": tipo_conta, "codigo_contabil": codigo_conta, "saldo_inicial": float(saldo_inicial)})
                        if res is not None:
                            st.cache_data.clear(); st.success("Conta cadastrada!"); time.sleep(1); st.rerun()
                    else:
                        st.warning("Informe o nome da conta.")

    with col_edit:
        with st.expander("✏️ Editar ou Excluir Conta"):
            if contas_bancarias_db:
                conta_map = {c['nome']: c for c in contas_bancarias_db}
                conta_sel_edit = st.selectbox("Selecione a Conta", list(conta_map.keys()), key="select_conta_edicao_unica")
                conta_data = conta_map[conta_sel_edit]
                
                with st.form("form_edit_conta"):
                    n_nome = st.text_input("Nome da Conta", value=conta_data['nome'])
                    n_tipo = st.selectbox("Tipo", ["Corrente", "Poupança", "Investimento"], index=["Corrente", "Poupança", "Investimento"].index(conta_data.get('tipo', 'Corrente')), key="tipo_conta_edit")
                    n_codigo = st.text_input("Código Contábil", value=conta_data.get('codigo_contabil') or "")
                    n_saldo = st.number_input("Saldo Inicial (R$)", value=float(conta_data.get('saldo_inicial') or 0.0), format="%.2f")
                    
                    c_btn1, c_btn2 = st.columns(2)
                    btn_upd = c_btn1.form_submit_button("💾 Atualizar", use_container_width=True)
                    btn_del = c_btn2.form_submit_button("🗑️ Excluir", use_container_width=True)
                    
                    if btn_upd:
                        res = sb_request("contas_bancarias", "PATCH", {"nome": n_nome, "tipo": n_tipo, "codigo_contabil": n_codigo, "saldo_inicial": float(n_saldo)}, filtros={"id": f"eq.{conta_data['id']}"})
                        if res is not None:
                            st.cache_data.clear(); st.success("Atualizada!"); time.sleep(1); st.rerun()
                    if btn_del:
                        res = sb_request("contas_bancarias", "DELETE", filtros={"id": f"eq.{conta_data['id']}"})
                        if res is not None:
                            st.cache_data.clear(); st.success("Excluída!"); time.sleep(1); st.rerun()

    if not contas_bancarias_db:
        st.info("Cadastre ao menos uma conta bancária acima para iniciar a conciliação.")
    else:
        conta_opcoes = {c["nome"]: c["id"] for c in contas_bancarias_db}
        conta_sel = st.selectbox("Selecione a Conta", list(conta_opcoes.keys()), key="select_conta_principal_conciliacao")
        conta_id = conta_opcoes[conta_sel]

        df = carregar_lancamentos_df()
        df_conta = df[(df.get('conta_bancaria_id') == conta_id) & (df['status'] == 'Concluído')].copy() if not df.empty else pd.DataFrame()

        conta_info = next(c for c in contas_bancarias_db if c["id"] == conta_id)
        saldo_calculado = float(conta_info.get('saldo_inicial') or 0)
        if not df_conta.empty:
            saldo_calculado += df_conta[df_conta['tipo'] == 'Entrada']['valor'].sum()
            saldo_calculado -= df_conta[df_conta['tipo'] == 'Saída']['valor'].sum()

        col1, col2 = st.columns(2)
        col1.metric("Saldo Calculado no Sistema", fmt_moeda(saldo_calculado))
        saldo_extrato = col2.number_input("Saldo informado no extrato do banco", min_value=0.0, format="%.2f")

        if saldo_extrato > 0:
            diferenca = saldo_calculado - saldo_extrato
            if abs(diferenca) < 0.01:
                st.success("✅ Saldo conferido! Sistema e banco estão batendo.")
            else:
                st.error(f"⚠️ Diferença encontrada: {fmt_moeda(diferenca)}.")

        st.markdown("---")
        tab_ofx, tab_manual = st.tabs(["⚡ Importação Inteligente (OFX)", "🖐️ Conciliação Manual"])

        with tab_ofx:
            st.markdown("Faça o upload do arquivo **.OFX** gerado pelo seu banco. O sistema cruzará os dados e **cadastrará todas as transações marcadas instantaneamente**. As saídas que precisarem de nota fiscal receberão o aviso 🔴 **Falta anexo** na Conciliação Manual para você anexar depois.")
            arquivo_ofx = st.file_uploader("Selecione o arquivo OFX do banco", type=['ofx', 'txt'], key="up_ofx_novo_v5")

            if arquivo_ofx:
                content = arquivo_ofx.read().decode('latin1', errors='ignore')
                transacoes = []
                
                for bloco in re.split(r'<\s*STMTTRN\s*>', content, flags=re.IGNORECASE)[1:]:
                    dt_match = re.search(r'<DTPOSTED>(\d{8})', bloco)
                    valor_match = re.search(r'<TRNAMT>([-\d\.]+)', bloco)
                    memo_match = re.search(r'<MEMO>(.*?)(?:<|$)', bloco)
                    name_match = re.search(r'<NAME>(.*?)(?:<|$)', bloco)

                    if dt_match and valor_match:
                        dt_str = dt_match.group(1)[:8]
                        try:
                            dt = pd.to_datetime(dt_str, format='%Y%m%d').date()
                        except:
                            continue
                        
                        valor = float(valor_match.group(1))
                        
                        raw_desc = memo_match.group(1).strip() if memo_match else (name_match.group(1).strip() if name_match else "Extrato Bancário")
                        raw_desc = re.sub(r'<[^>]+>', '', raw_desc)
                        
                        desc_upper = raw_desc.upper()
                        nome_extraido = raw_desc
                        
                        for prefixo in ["PIX - RECEBIMENTO:", "PIX -", "PIX TRANSF", "PIX RECEBIDO", "TED -", "DOC -", "TRANSF. COPIE E COLE", "TRANSFERENCIA"]:
                            if desc_upper.startswith(prefixo):
                                nome_extraido = raw_desc[len(prefixo):].strip(" -/")
                                break
                        
                        if not nome_extraido:
                            nome_extraido = raw_desc

                        transacoes.append({
                            "Data": dt,
                            "Valor": abs(valor),
                            "Descrição Bancária": raw_desc[:80],
                            "Nome Identificado": nome_extraido.title() if len(nome_extraido) > 2 else raw_desc,
                            "Tipo": "Entrada" if valor >= 0 else "Saída",
                            "Status": "Não registrado"
                        })

                df_ofx = pd.DataFrame(transacoes)

                if not df_ofx.empty:
                    used_ids = set()
                    for idx, row in df_ofx.iterrows():
                        if not df_conta.empty:
                            mask_tipo = df_conta['tipo'] == row['Tipo']
                            mask_valor = df_conta['valor'] == row['Valor']
                            mask_data = (pd.to_datetime(df_conta['data_competencia']).dt.date - row['Data']).apply(lambda x: abs(x.days)) <= 3
                            mask_used = ~df_conta['id'].isin(used_ids)

                            match = df_conta[mask_tipo & mask_valor & mask_data & mask_used]
                            if not match.empty:
                                match_id = match.iloc[0]['id']
                                df_ofx.at[idx, 'Status'] = 'Já no sistema'
                                used_ids.add(match_id)

                    df_novos = df_ofx[df_ofx['Status'] == 'Não registrado'].copy()

                    if df_novos.empty:
                        st.success("🎉 Todas as transações deste extrato já constam e batem com o sistema!")
                    else:
                        st.info(f"Encontramos **{len(df_novos)} transações** novas no extrato.")
                        
                        cats_entrada = [c['nome'] for c in categorias_db if c['tipo'] == 'Entrada']
                        cats_saida = [c['nome'] for c in categorias_db if c['tipo'] == 'Saída']
                        nomes_eventos = ["Nenhum"] + [e['nome'] for e in eventos_db]
                        map_cat_id = {c['nome']: c['id'] for c in categorias_db}
                        
                        df_entradas = df_novos[df_novos['Tipo'] == 'Entrada'].copy()
                        df_saidas = df_novos[df_novos['Tipo'] == 'Saída'].copy()
                        
                        df_entradas_final = pd.DataFrame()
                        df_saidas_final = pd.DataFrame()

                        if not df_entradas.empty:
                            st.markdown("#### 🟢 Novas Entradas (Recebimentos)")
                            df_entradas.insert(0, 'Cadastrar', True)
                            df_entradas['Descrição para Sistema'] = "Extrato: " + df_entradas['Nome Identificado']
                            df_entradas['Projeto'] = "Nenhum"
                            df_entradas['Categoria'] = cats_entrada[0] if cats_entrada else ""
                            
                            df_entradas_final = st.data_editor(
                                df_entradas[['Cadastrar', 'Data', 'Descrição Bancária', 'Descrição para Sistema', 'Valor', 'Categoria', 'Projeto']],
                                hide_index=True, use_container_width=True, key="ed_entradas_ofx"
                            )

                        if not df_saidas.empty:
                            st.markdown("#### 🔴 Novas Saídas (Pagamentos) - Serão marcadas como 'Falta Anexo' automaticamente")
                            df_saidas.insert(0, 'Cadastrar', True)
                            df_saidas['Descrição para Sistema'] = "Extrato: " + df_saidas['Nome Identificado']
                            df_saidas['Categoria'] = cats_saida[0] if cats_saida else ""
                            df_saidas['Projeto'] = "Nenhum"
                            
                            df_saidas_final = st.data_editor(
                                df_saidas[['Cadastrar', 'Data', 'Descrição Bancária', 'Descrição para Sistema', 'Valor', 'Categoria', 'Projeto']],
                                hide_index=True, use_container_width=True, key="ed_saidas_ofx"
                            )

                        if st.button("💾 Salvar Lote Imediatamente", type="primary"):
                            total_salvos = 0
                            
                            if not df_entradas_final.empty:
                                para_salvar_ent = df_entradas_final[df_entradas_final['Cadastrar'] == True]
                                for _, row in para_salvar_ent.iterrows():
                                    cc = None if row['Projeto'] == "Nenhum" else row['Projeto']
                                    sb_request("lancamentos", "POST", [{
                                        "descricao": row['Descrição para Sistema'], "tipo": "Entrada",
                                        "valor": float(row['Valor']), "data_competencia": str(row['Data']),
                                        "status": "Concluído", "data_pagamento": str(row['Data']),
                                        "categoria_id": map_cat_id.get(row['Categoria']),
                                        "conta_bancaria_id": conta_id, "centro_custo": cc, "conciliado": True
                                    }])
                                    total_salvos += 1

                            if not df_saidas_final.empty:
                                para_salvar_sai = df_saidas_final[df_saidas_final['Cadastrar'] == True]
                                for _, row in para_salvar_sai.iterrows():
                                    cc = None if row['Projeto'] == "Nenhum" else row['Projeto']
                                    sb_request("lancamentos", "POST", [{
                                        "descricao": row['Descrição para Sistema'], "tipo": "Saída",
                                        "valor": float(row['Valor']), "data_competencia": str(row['Data']),
                                        "status": "Concluído", "data_pagamento": str(row['Data']),
                                        "categoria_id": map_cat_id.get(row['Categoria']),
                                        "conta_bancaria_id": conta_id, "centro_custo": cc, "conciliado": True
                                    }])
                                    total_salvos += 1

                            if total_salvos > 0:
                                st.cache_data.clear()
                                st.success(f"✅ {total_salvos} transações salvas com sucesso! Vá para 'Conciliação Manual' para anexar os comprovantes pendentes.")
                                time.sleep(2)
                                st.rerun()
                            else:
                                st.warning("Marque pelo menos um item para salvar.")

        with tab_manual:
            st.subheader("Lista de Lançamentos e Pendências de Anexos")
            
            if df_conta.empty:
                st.info("Nenhum lançamento concluído nesta conta ainda.")
            else:
                col_f1, col_f2 = st.columns(2)
                meses_disponiveis = sorted(df_conta['data_competencia'].dt.strftime('%Y-%m').unique().tolist(), reverse=True)
                filtro_mes = col_f1.selectbox("Filtrar por Mês", ["Todos"] + meses_disponiveis, key="concil_filtro_mes")
                apenas_nao_conciliados = col_f2.checkbox("Mostrar apenas não conciliados", value=True, key="concil_so_pendentes")
                
                dff_manual = df_conta.copy()
                if filtro_mes != "Todos":
                    dff_manual = dff_manual[dff_manual['data_competencia'].dt.strftime('%Y-%m') == filtro_mes]
                if apenas_nao_conciliados:
                    dff_manual = dff_manual[dff_manual['conciliado'] != True]
                
                if dff_manual.empty:
                    st.success("Nenhum lançamento pendente de conciliação para este filtro! 🎉")
                else:
                    todos_anexos_conc = sb_request("lancamento_anexos", "GET")
                    mapa_anexos_conc = {}
                    if todos_anexos_conc:
                        for ax in todos_anexos_conc:
                            l_id = ax['lancamento_id']
                            if l_id not in mapa_anexos_conc:
                                mapa_anexos_conc[l_id] = []
                            mapa_anexos_conc[l_id].append(ax)
                    
                    hc1, hc2, hc3, hc4, hc5, hc6, hc7 = st.columns([1.0, 2.5, 2.0, 1.2, 1.2, 0.8, 1.2])
                    hc1.markdown("**Data**")
                    hc2.markdown("**Descrição**")
                    hc3.markdown("**Categoria**")
                    hc4.markdown("**Valor**")
                    hc5.markdown("**Anexos**")
                    hc6.markdown("**Det.**")
                    hc7.markdown("**Conciliado**")
                    st.markdown("<hr style='margin:4px 0;border-color:#CBD5E1;'>", unsafe_allow_html=True)

                    for _, row in dff_manual.sort_values('data_competencia', ascending=False).iterrows():
                        row_id = str(row['id'])
                        c1, c2, c3, c4, c5, c6, c7 = st.columns([1.0, 2.5, 2.0, 1.2, 1.2, 0.8, 1.2])
                        
                        c1.write(row['data_competencia'].strftime('%d/%m/%Y'))
                        c2.write(row['descricao'] or '—')
                        c3.write(row['categoria_nome'])
                        
                        val_str = fmt_moeda(row['valor']) if row['tipo'] == 'Entrada' else f"-{fmt_moeda(row['valor'])}"
                        c4.write(val_str)
                        
                        anexos_deste = mapa_anexos_conc.get(row_id, [])
                        if not anexos_deste and row.get('url_anexo') and isinstance(row.get('url_anexo'), str) and row.get('url_anexo').strip():
                            anexos_deste = [{'id': 'legacy', 'url_storage': row.get('url_anexo'), 'nome_original': row.get('url_anexo').split('/')[-1]}]
                            
                        if len(anexos_deste) > 0:
                            c5.markdown(f"📎 {len(anexos_deste)} anexo(s)")
                        else:
                            if row['tipo'] == 'Saída':
                                c5.markdown('<span class="badge-falta-anexo">🔴 Falta anexo</span>', unsafe_allow_html=True)
                            else:
                                c5.markdown("—")
                        
                        if c6.button("🔍", key=f"detalhe_btn_{row_id}", help="Ver detalhes e anexos"):
                            st.session_state[f"show_detalhe_{row_id}"] = not st.session_state.get(f"show_detalhe_{row_id}", False)
                            st.rerun()
                        
                        marcado = bool(row.get('conciliado'))
                        novo_valor = c7.checkbox(" ", value=marcado, key=f"conc_{row_id}")
                        
                        if novo_valor != marcado:
                            sb_request("lancamentos", "PATCH", {"conciliado": novo_valor}, filtros={"id": f"eq.{row_id}"})
                            st.cache_data.clear(); st.rerun()

                        if st.session_state.get(f"show_detalhe_{row_id}", False):
                            with st.container(border=True):
                                st.markdown(f"**Conferência e Anexos:** {row['descricao']}")
                                st.write(f"• **Tipo:** {row['tipo']} | **Centro de Custo:** {row.get('centro_custo', 'Nenhum')}")
                                
                                st.markdown("---")
                                if anexos_deste:
                                    st.markdown("**📎 Comprovantes Atuais:**")
                                    for anexo in anexos_deste:
                                        col_a1, col_a2 = st.columns([3, 1])
                                        link_download = obter_link_arquivo(anexo['url_storage'])
                                        nome_ex = anexo.get('nome_original') or anexo['url_storage'].split('/')[-1]
                                        
                                        if link_download:
                                            col_a1.markdown(f"📄 [{nome_ex}]({link_download})", unsafe_allow_html=True)
                                        else:
                                            col_a1.write(nome_ex)
                                            
                                        if col_a2.button("🗑️ Apagar", key=f"del_anexo_conc_{anexo.get('id', 'legacy')}_{row_id}"):
                                            try:
                                                supabase.storage.from_("comprovantes").remove([anexo['url_storage']])
                                                if anexo.get('id') == 'legacy':
                                                    sb_request("lancamentos", "PATCH", {"url_anexo": None}, filtros={"id": f"eq.{row_id}"})
                                                else:
                                                    sb_request("lancamento_anexos", "DELETE", filtros={"id": f"eq.{anexo['id']}"})
                                                st.cache_data.clear(); st.rerun()
                                            except Exception:
                                                pass
                                else:
                                    if row['tipo'] == 'Saída':
                                        st.error("⚠️ Esta saída está exigindo um comprovante para fechamento fiscal.")
                                    else:
                                        st.caption("Nenhum arquivo anexado (Opcional para entradas).")

                                arq_adicionais_conc = st.file_uploader("📎 Enviar Comprovante(s) Pendente(s)", type=['png', 'jpg', 'jpeg', 'pdf'], accept_multiple_files=True, key=f"up_concil_{row_id}")
                                c_btn_up1, c_btn_up2 = st.columns(2)
                                
                                if c_btn_up1.button("💾 Salvar Arquivos", key=f"btn_save_concil_up_{row_id}", type="primary"):
                                    if arq_adicionais_conc:
                                        with st.spinner("Salvando..."):
                                            processar_e_salvar_anexos(arq_adicionais_conc, row_id, row['data_competencia'], row['categoria_nome'], row['descricao'])
                                            st.cache_data.clear()
                                            st.success("Anexos salvos com sucesso!")
                                            time.sleep(1)
                                            st.rerun()
                                    else:
                                        st.warning("Selecione um arquivo para enviar.")
                                        
                                if c_btn_up2.button("❌ Fechar Detalhes", key=f"close_det_{row_id}"):
                                    st.session_state[f"show_detalhe_{row_id}"] = False
                                    st.rerun()

                        st.markdown("<hr style='margin:2px 0;border-color:#F1F5F9;'>", unsafe_allow_html=True)

# ==========================================
# ==========================================
# PAINEL DE EVENTOS
# ==========================================
elif page == "Painel de Eventos":
    st.title("Painel de Eventos e Campanhas")
    st.markdown("Cadastre acampamentos e retiros (com inscrições) ou campanhas de arrecadação (sem participantes).")

    col_novo, col_edit = st.columns(2)
    with col_novo:
        with st.expander("➕ Criar Novo Evento/Campanha", expanded=len(eventos_db) == 0):
            with st.form("form_evento", clear_on_submit=True):
                nome_ev = st.text_input("Nome do Evento / Campanha")
                
                tem_participantes = st.checkbox("Requer inscrição de participantes?", value=True, help="Desmarque para projetos puramente de arrecadação (ex: Reforma, Som, Missões) onde não há lista de pessoas inscritas.")
                
                c_lead1, c_lead2 = st.columns(2)
                lideres_opcoes = {u['nome']: u['id'] for u in usuarios_db if u.get('perfil') == 'Visão Eventos'}
                lider_sel = c_lead1.selectbox("Líder Responsável", ["Nenhum"] + list(lideres_opcoes.keys()))
                codigo_centavos = c_lead2.text_input("Código de Centavos (Ex: 07)", max_chars=2, help="Usado para identificar pagamentos via PIX automaticamente no extrato.")

                codigo_rec_ev = st.text_input("Código Contábil de Receita (Ex: 4.1.10.100.007)")
                data_ev = st.date_input("Data do Evento (ou fim da campanha)", date.today())
                valor_ev = st.number_input("Valor da Inscrição / Alvo (R$)", min_value=0.0, format="%.2f")
                vagas_ev = st.number_input("Total de Vagas (0 para ilimitado)", min_value=0, step=1)
                chave_pix_ev = st.text_input("Chave Pix")
                descricao_ev = st.text_area("Descrição/Orientações")
                
                permite_parc = st.checkbox("Permitir parcelamento")
                num_parc = st.number_input("Máx. Parcelas", min_value=1, max_value=12, value=1)

                if st.form_submit_button("Criar Registro", use_container_width=True):
                    if not nome_ev.strip() or not chave_pix_ev.strip():
                        st.warning("Preencha ao menos o Nome e a Chave Pix.")
                    else:
                        lider_id_val = lideres_opcoes.get(lider_sel) if lider_sel != "Nenhum" else None
                        total_parcelas = int(num_parc) if permite_parc else 1
                        
                        sb_request("eventos", "POST", {
                            "nome": nome_ev.strip(), "descricao": descricao_ev.strip() or None,
                            "codigo_receita_contabil": codigo_rec_ev,
                            "lider_id": lider_id_val,
                            "codigo_centavos": codigo_centavos.strip() if codigo_centavos else None,
                            "data_evento": str(data_ev), "valor_inscricao": float(valor_ev),
                            "vagas_total": int(vagas_ev), "chave_pix": chave_pix_ev.strip(),
                            "centro_custo": nome_ev.strip(), "status": "Aberto",
                            "permite_parcelamento": bool(permite_parc), "numero_parcelas": total_parcelas,
                            "tem_participantes": bool(tem_participantes)
                        })
                        st.cache_data.clear(); st.success("Criado com sucesso!"); time.sleep(1); st.rerun()

    with col_edit:
        with st.expander("✏️ Editar ou Excluir"):
            if eventos_db:
                ev_map = {e['nome']: e for e in eventos_db}
                ev_sel = st.selectbox("Selecione o Evento/Campanha", list(ev_map.keys()))
                ev_data = ev_map[ev_sel]
                
                with st.form("form_edit_evento"):
                    n_nome_ev = st.text_input("Nome", value=ev_data['nome'])
                    n_tem_part = st.checkbox("Requer inscrição de participantes?", value=ev_data.get('tem_participantes', True))
                    n_codigo_rec_ev = st.text_input("Código Contábil", value=ev_data.get('codigo_receita_contabil') or "")
                    
                    raw_data_ev = ev_data.get('data_evento')
                    if pd.isna(raw_data_ev) or not raw_data_ev:
                        default_date = date.today()
                    else:
                        try:
                            default_date = pd.to_datetime(raw_data_ev).date()
                        except Exception:
                            default_date = date.today()

                    n_data_ev = st.date_input("Data", default_date)
                    n_valor_ev = st.number_input("Valor (R$)", value=float(ev_data.get('valor_inscricao') or 0), format="%.2f")
                    n_vagas_ev = st.number_input("Vagas", value=int(ev_data.get('vagas_total') or 0))
                    n_pix = st.text_input("Pix", value=ev_data.get('chave_pix') or "")
                    n_desc = st.text_area("Descrição", value=ev_data.get('descricao') or "")
                    
                    c_btn1, c_btn2 = st.columns(2)
                    b_upd_ev = c_btn1.form_submit_button("💾 Atualizar", use_container_width=True)
                    b_del_ev = c_btn2.form_submit_button("🗑️ Excluir", use_container_width=True)
                    
                    if b_upd_ev:
                        sb_request("eventos", "PATCH", {
                            "nome": n_nome_ev, "codigo_receita_contabil": n_codigo_rec_ev, "data_evento": str(n_data_ev), "valor_inscricao": float(n_valor_ev),
                            "vagas_total": int(n_vagas_ev), "chave_pix": n_pix, "descricao": n_desc, "tem_participantes": n_tem_part
                        }, filtros={"id": f"eq.{ev_data['id']}"})
                        st.cache_data.clear(); st.success("Atualizado!"); time.sleep(1); st.rerun()
                    if b_del_ev:
                        sb_request("eventos", "DELETE", filtros={"id": f"eq.{ev_data['id']}"})
                        st.cache_data.clear(); st.success("Excluído!"); time.sleep(1); st.rerun()
            else:
                st.info("Nenhum evento cadastrado.")

    eventos_atualizados = carregar("eventos")
    st.markdown("---")
    st.subheader("Cadastros Ativos")

    if not eventos_atualizados:
        st.info("Nenhum evento/campanha cadastrado ainda.")
    else:
        inscricoes_all = carregar("inscricoes")
        pagamentos_all = carregar("inscricao_pagamentos")

        for ev in eventos_atualizados:
            tem_part = ev.get('tem_participantes', True)
            status_evento = ev.get("status") or "Aberto"
            cor_status = "#059669" if status_evento == "Aberto" else "#94A3B8"
            
            # Formatação dos centavos para aparecer no cartão
            codigo_centavos = ev.get("codigo_centavos")
            texto_centavos_card = f" &nbsp;•&nbsp; 🪙 Centavos para Pix: <b>,{codigo_centavos}</b>" if codigo_centavos else ""
            
            # Montagem das URLs padrão
            app_url = st.secrets.get("APP_URL", "").rstrip("/")
            if "[" in app_url and "](" in app_url:
                import re
                match = re.search(r'\((https?://[^)]+)\)', app_url)
                app_url = match.group(1) if match else "https://financeiro-comuna-sbc.streamlit.app"
            elif not app_url:
                app_url = "https://financeiro-comuna-sbc.streamlit.app"
                
            complemento_link = f"?pagina=inscricao&evento={ev['id']}"
            link_publico = f"{app_url}/{complemento_link}"

            if tem_part:
                inscricoes_evento = [i for i in inscricoes_all if str(i.get("evento_id")) == str(ev.get("id"))]
                ids_inscricoes_evento = {str(i.get("id")) for i in inscricoes_evento}
                pagamentos_evento = [p for p in pagamentos_all if str(p.get("inscricao_id")) in ids_inscricoes_evento]
                pagamentos_pendentes = [p for p in pagamentos_evento if p.get("status") == "Pendente"]
                inscricoes_quitadas = [i for i in inscricoes_evento if i.get("status_pagamento") == "Completo"]
                inscricoes_parciais = [i for i in inscricoes_evento if i.get("status_pagamento") == "Parcial"]

                total_arrecadado = sum(float(i.get("valor_pago") or 0) for i in inscricoes_evento)
                total_previsto = sum(float(i.get("valor_total") or 0) for i in inscricoes_evento)
                saldo_a_receber = max(total_previsto - total_arrecadado, 0)
                total_inscritos = len(inscricoes_evento)
                vagas_total = int(ev.get("vagas_total") or 0)
                vagas_restantes = max(vagas_total - total_inscritos, 0) if vagas_total > 0 else "Ilimitadas"
                parcelamento_texto = f"Pagamento em até {int(ev.get('numero_parcelas') or 1)}x" if ev.get("permite_parcelamento") else "Pagamento à vista"
                
                info_participantes = f"👥 {total_inscritos} inscritos &nbsp;•&nbsp; {vagas_restantes} vagas restantes &nbsp;•&nbsp; ✅ {len(inscricoes_quitadas)} quitados &nbsp;•&nbsp; 🟡 {len(inscricoes_parciais)} parciais"
                info_alertas = f'<p style="color:#D97706;margin:4px 0;font-weight:600;">⏳ {len(pagamentos_pendentes)} comprovantes aguardando aprovação</p>' if pagamentos_pendentes else ''

                # Geração do texto WhatsApp principal para Evento
                texto_whatsapp = f"Olá! As inscrições para o *{ev.get('nome')}* estão abertas!\n\n*Data:* {ev.get('data_evento') or '—'}\n*Valor:* {fmt_moeda(ev.get('valor_inscricao'))} ({parcelamento_texto})\n\n*Faça sua inscrição pelo link:*\n{link_publico}"
                
                if codigo_centavos:
                    texto_whatsapp += f"\n\n*Atenção:* Ao fazer o pagamento via Pix, adicione nossos centavos (*,{codigo_centavos}*) no valor final. Exemplo: R$ {int(ev.get('valor_inscricao') or 0)},{codigo_centavos}. Isso garante a confirmação automática no sistema!"

            else:
                creditos_ev = sb_request("creditos_ofx", "GET", filtros={"evento_id": f"eq.{ev['id']}"}) or []
                total_arrecadado = sum(float(c.get("valor") or 0) for c in creditos_ev if c.get("status") == "Vinculado")
                meta_alvo = float(ev.get('valor_inscricao') or 0)
                saldo_a_receber = max(meta_alvo - total_arrecadado, 0)
                
                info_participantes = "🎯 Campanha de Arrecadação Pública (Não requer inscritos)"
                info_alertas = ""
                parcelamento_texto = "Doação Espontânea"

                # Geração do texto WhatsApp principal para Campanha
                texto_whatsapp = f"Olá! Nossa campanha *{ev.get('nome')}* está ativa!\n\nNossa meta é arrecadar *{fmt_moeda(ev.get('valor_inscricao'))}* e toda ajuda faz muita diferença!"
                
                if codigo_centavos:
                    texto_whatsapp += f"\n\n*Importante:* Adicione o código (*,{codigo_centavos}*) no final do valor da sua doação. Exemplo: para doar R$ 50 transfira R$ 50,{codigo_centavos}. Isso nos ajuda a identificar sua doação de forma rápida e automática!"

            # Chamada de ação que conecta a primeira mensagem com a segunda (a chave pix isolada)
            texto_whatsapp += "\n\nUse o código PIX abaixo para fazer sua transferência:"

            # Extração da Chave Pix Pura
            chave_pix_pura = ev.get('chave_pix') or '—'

            # O HTML foi envelopado sem quebras de linha com indentação
            card_html = (
                f'<div style="background:#FFFFFF;border:1px solid #E2E8F0;border-radius:14px;padding:20px;margin-bottom:12px;">'
                f'<h4 style="margin-top:0;">{ev.get("nome", "Evento")} <span style="font-size:0.8rem;color:{cor_status};">● {status_evento}</span></h4>'
                f'<p style="color:#475569;margin:4px 0;">📅 {ev.get("data_evento") or "—"} &nbsp;•&nbsp; 💰 {fmt_moeda(ev.get("valor_inscricao"))} {"por pessoa" if tem_part else "(Alvo Global)"}</p>'
                f'<p style="color:#475569;margin:4px 0;">💳 {parcelamento_texto} &nbsp;•&nbsp; 🔑 Pix: {ev.get("chave_pix") or "—"}{texto_centavos_card}</p>'
                f'<p style="color:#475569;margin:4px 0;">{info_participantes}</p>'
                f'{info_alertas}'
                f'<p style="color:#059669;margin:4px 0;font-weight:600;">💵 Arrecadado: {fmt_moeda(total_arrecadado)} &nbsp;•&nbsp; {"A receber" if tem_part else "Faltam"}: {fmt_moeda(saldo_a_receber)}</p>'
                f'</div>'
            )
            st.markdown(card_html, unsafe_allow_html=True)

            col_esq, col_meio, col_dir = st.columns([2.2, 1, 1.2])
            
            with col_esq:
                if tem_part:
                    st.caption("Link de inscrição (Passe o mouse e copie 📋)")
                    st.code(link_publico, language="text")
                else:
                    st.info("💡 Por ser campanha, use o texto do botão ao lado ➡")
            
            with col_meio:
                with st.popover("📱 Divulgar no WhatsApp", use_container_width=True):
                    import urllib.parse
                    st.caption("O WhatsApp não permite enviar duas mensagens separadas com um único clique. Siga os 2 passos abaixo:")
                    
                    st.markdown("**Passo 1: Enviar Instruções**")
                    st.code(texto_whatsapp, language="text")
                    link_wa_texto = f"https://wa.me/?text={urllib.parse.quote(texto_whatsapp)}"
                    st.link_button("💬 Enviar Instruções", link_wa_texto, use_container_width=True)
                    
                    st.markdown("---")
                    
                    st.markdown("**Passo 2: Enviar APENAS a Chave Pix** (para a pessoa conseguir copiar no celular)")
                    st.code(chave_pix_pura, language="text")
                    link_wa_pix = f"https://wa.me/?text={urllib.parse.quote(chave_pix_pura)}"
                    st.link_button("🔑 Enviar só a Chave Pix", link_wa_pix, use_container_width=True)
                    
            with col_dir:
                novo_status = "Encerrado" if status_evento == "Aberto" else "Aberto"
                texto_botao = "🔒 Encerrar " + ("inscrições" if tem_part else "campanha")
                if status_evento != "Aberto":
                    texto_botao = "🔓 Reabrir " + ("inscrições" if tem_part else "campanha")
                if st.button(texto_botao, key=f"alterar_status_ev_{ev['id']}", use_container_width=True):
                    resultado = sb_request("eventos", "PATCH", {"status": novo_status}, filtros={"id": f"eq.{ev['id']}"})
                    if resultado is not None:
                        st.cache_data.clear()
                        st.rerun()
                        
            st.markdown("---")
# ==========================================
# ==========================================
# INSCRIÇÕES E COMPROVANTES
# ==========================================
elif page == "Inscrições e Comprovantes":
    st.title("Validação e Bolsão OFX")
    st.markdown("Valide comprovantes de Pix ou distribua os depósitos bancários identificados por centavos.")

    if not eventos_db:
        st.info("Cadastre um evento ou campanha primeiro no 'Painel de Eventos'.")
    else:
        usuario_atual = st.session_state.get("usuario_logado", {})
        if usuario_atual.get("perfil") == "Visão Eventos":
            eventos_visiveis = [e for e in eventos_db if str(e.get("lider_id")) == str(usuario_atual.get("id"))]
            if not eventos_visiveis:
                st.warning("Você não está associado a nenhum evento no momento.")
                st.stop()
        else:
            eventos_visiveis = eventos_db

        evento_opcoes = {e["nome"]: e["id"] for e in eventos_visiveis}
        evento_sel = st.selectbox("Selecione o Evento / Campanha", list(evento_opcoes.keys()))
        evento_id_sel = evento_opcoes[evento_sel]
        evento_obj_sel = next(e for e in eventos_visiveis if str(e['id']) == str(evento_id_sel))
        ev_tem_participantes = evento_obj_sel.get('tem_participantes', True)

        inscricoes_evento = [i for i in carregar("inscricoes") if str(i.get('evento_id')) == str(evento_id_sel)]
        pagamentos_all = carregar("inscricao_pagamentos")
        creditos_ofx_all = sb_request("creditos_ofx", "GET", filtros={"evento_id": f"eq.{evento_id_sel}"}) or []
        
        map_insc = {i["id"]: i for i in inscricoes_evento}

        tab_bolsao, tab_comprovantes, tab_lista = st.tabs(["💵 Bolsão de Créditos OFX", "⏳ Comprovantes Web", "👥 Participantes/Lista"])

        with tab_bolsao:
            st.markdown("### Créditos do Extrato Bancário Aguardando Vínculo")
            
            if ev_tem_participantes:
                st.markdown("Estes valores entraram na conta com o código de centavos. Você pode vincular o valor total a um participante, ou uma parte dele (o sistema separará o troco automaticamente para o bolsão).")
            else:
                st.markdown("Estes valores entraram via código de centavos da Campanha. Vincule-os para creditá-los ao saldo deste projeto.")

            creditos_disponiveis = [c for c in creditos_ofx_all if c.get("status") == "Disponível"]
            
            if not creditos_disponiveis:
                st.success("Nenhum crédito bancário pendente de distribuição para este evento. 👍")
            else:
                for cred in creditos_disponiveis:
                    c_b1, c_b2, c_b3 = st.columns([2, 2, 2])
                    c_b1.write(f"📅 **{pd.to_datetime(cred['data']).strftime('%d/%m/%Y')}** — {fmt_moeda(cred['valor'])}")
                    c_b2.caption(f"Banco: {cred['descricao_bancaria']}")
                    
                    with c_b3:
                        if ev_tem_participantes:
                            with st.popover("🔗 Vincular a Participante"):
                                if not inscricoes_evento:
                                    st.warning("Nenhum participante inscrito. Cadastre na aba 'Participantes'.")
                                else:
                                    participante_opcoes = {i['nome_participante']: i['id'] for i in inscricoes_evento}
                                    part_sel = st.selectbox("Escolha o Participante", list(participante_opcoes.keys()), key=f"sel_part_{cred['id']}")
                                    
                                    insc_id = participante_opcoes[part_sel]
                                    insc_obj = map_insc[insc_id]
                                    
                                    saldo_devedor = max(float(insc_obj.get('valor_total') or 0) - float(insc_obj.get('valor_pago') or 0), 0.0)
                                    sugestao_vinculo = min(float(cred['valor']), saldo_devedor) if saldo_devedor > 0 else float(cred['valor'])
                                    
                                    valor_a_vincular = st.number_input(
                                        "Valor a vincular (R$)", 
                                        min_value=0.01, 
                                        max_value=float(cred['valor']), 
                                        value=float(sugestao_vinculo), 
                                        format="%.2f", 
                                        key=f"val_vinc_{cred['id']}"
                                    )
                                    
                                    if st.button("Confirmar Vínculo", key=f"btn_vinc_cred_{cred['id']}", type="primary"):
                                        if saldo_devedor <= 0.0:
                                            st.error(f"Atenção: A inscrição de {part_sel} já está 100% quitada.")
                                        elif valor_a_vincular > (saldo_devedor + 0.02):
                                            st.error(f"O limite para este participante é {fmt_moeda(saldo_devedor)} (saldo que falta para quitar). Reduza o valor.")
                                        else:
                                            valor_restante = round(float(cred['valor']) - valor_a_vincular, 2)
                                            
                                            if valor_restante > 0.0:
                                                sb_request("creditos_ofx", "PATCH", {"valor": valor_a_vincular, "status": "Vinculado", "inscricao_id": insc_id}, filtros={"id": f"eq.{cred['id']}"})
                                                sb_request("creditos_ofx", "POST", {
                                                    "evento_id": cred['evento_id'],
                                                    "lancamento_id": cred.get('lancamento_id'),
                                                    "data": cred['data'],
                                                    "valor": valor_restante,
                                                    "descricao_bancaria": cred['descricao_bancaria'] + " (Saldo Restante)",
                                                    "status": "Disponível"
                                                })
                                            else:
                                                sb_request("creditos_ofx", "PATCH", {"status": "Vinculado", "inscricao_id": insc_id}, filtros={"id": f"eq.{cred['id']}"})
                                                
                                            novo_pag = sb_request("inscricao_pagamentos", "POST", {
                                                "inscricao_id": insc_id,
                                                "numero_parcela": len([p for p in pagamentos_all if p.get("inscricao_id") == insc_id]) + 1,
                                                "valor": valor_a_vincular,
                                                "comprovante_url": None,
                                                "status": "Aprovado",
                                                "lancamento_id": cred.get('lancamento_id')
                                            })
                                            
                                            novo_valor_pago = float(insc_obj.get('valor_pago') or 0) + valor_a_vincular
                                            novo_status_p = "Completo" if novo_valor_pago >= float(insc_obj.get('valor_total') or 0) - 0.01 else "Parcial"
                                            sb_request("inscricoes", "PATCH", {"valor_pago": novo_valor_pago, "status_pagamento": novo_status_p}, filtros={"id": f"eq.{insc_id}"})
                                            
                                            st.cache_data.clear()
                                            st.success("Crédito vinculado com sucesso!")
                                            time.sleep(1)
                                            st.rerun()
                        else:
                            with st.popover("🔗 Vincular à Campanha"):
                                desc_campanha = st.text_input("Descrição / Observação (Ex: Doação do Irmão X)", key=f"desc_camp_{cred['id']}")
                                if st.button("Confirmar Associação à Campanha", key=f"btn_vinc_camp_{cred['id']}", type="primary"):
                                    sb_request("creditos_ofx", "PATCH", {"status": "Vinculado"}, filtros={"id": f"eq.{cred['id']}"})
                                    if cred.get('lancamento_id'):
                                        nome_cc = evento_obj_sel.get('centro_custo') or evento_sel
                                        novo_desc = f"{nome_cc} - {desc_campanha}" if desc_campanha else f"Arrecadação - {nome_cc}"
                                        sb_request("lancamentos", "PATCH", {"centro_custo": nome_cc, "descricao": novo_desc}, filtros={"id": f"eq.{cred['lancamento_id']}"})
                                    
                                    st.cache_data.clear()
                                    st.success("Crédito vinculado à campanha com sucesso!")
                                    time.sleep(1)
                                    st.rerun()

                    st.markdown("<hr style='margin:4px 0;border-color:#E2E8F0;'>", unsafe_allow_html=True)

            st.markdown("---")
            with st.expander("❓ Pagaram sem os centavos e não caiu aqui?"):
                st.info("Se alguém pagou e não colocou o código de centavos, preencha os dados abaixo. A Tesouraria será notificada no 'Resumo do Dia' para conferir o extrato e liberar o valor para o Bolsão.")
                with st.form("form_solicitar_credito"):
                    c_s1, c_s2 = st.columns(2)
                    s_nome = c_s1.text_input("Nome de quem pagou")
                    s_valor = c_s2.number_input("Valor exato pago (R$)", min_value=0.01, format="%.2f")
                    
                    c_s3, c_s4 = st.columns(2)
                    s_data = c_s3.date_input("Data em que foi feito o pagamento")
                    s_obs = c_s4.text_input("Observações (Opcional)")
                    
                    if st.form_submit_button("Enviar Solicitação para Tesouraria", type="primary"):
                        if not s_nome:
                            st.warning("Informe o nome do pagador.")
                        else:
                            sb_request("creditos_ofx", "POST", {
                                "evento_id": evento_id_sel,
                                "data": str(s_data),
                                "valor": float(s_valor),
                                "descricao_bancaria": f"Solicitação Manual: {s_nome} - {s_obs}",
                                "status": "Pendente Aprovação"
                            })
                            st.cache_data.clear()
                            st.success("Solicitação enviada ao Resumo do Dia! Assim que a Tesouraria aprovar, o crédito ficará disponível acima.")

        with tab_comprovantes:
            if not ev_tem_participantes:
                st.info("💡 Projetos do tipo Campanha não recebem inscrições e comprovantes via portal web. O controle é feito puramente via extrato e Bolsão.")
            else:
                st.markdown("### Comprovantes enviados pelo Link Web")
                pendentes_pg = [p for p in pagamentos_all if p.get("status") == "Pendente" and p.get("inscricao_id") in map_insc]
                if not pendentes_pg:
                    st.success("Nenhum comprovante pendente via web.")
                else:
                    for p in pendentes_pg:
                        insc = map_insc[p["inscricao_id"]]
                        c1, c2, c3, c4 = st.columns([2, 1, 1.5, 1.3])
                        c1.write(f"**{insc['nome_participante']}** — Parcela {p.get('numero_parcela',1)}")
                        c1.caption(f"CPF: {insc.get('cpf','—')} • {insc.get('contato','')}")
                        c2.write(fmt_moeda(p.get('valor')))
                        link = obter_link_arquivo(p.get('comprovante_url'))
                        c3.markdown(f"[📎 Ver comprovante]({link})" if link else "_Sem comprovante_")
                        col_a, col_b = c4.columns(2)
                        
                        if col_a.button("✅", key=f"aprovar_pg_{p['id']}", help="Aprovar"):
                            cat_evento_id = next((c['id'] for c in categorias_db if c['nome'] == 'Inscrições de Eventos'), None)
                            novo_lanc = sb_request("lancamentos", "POST", [{
                                "descricao": f"Inscrição ({insc['nome_participante']} - parcela {p.get('numero_parcela',1)}) - {evento_sel}",
                                "tipo": "Entrada", "valor": float(p.get('valor') or 0),
                                "data_competencia": str(date.today()), "status": "Concluído",
                                "categoria_id": cat_evento_id, "centro_custo": evento_sel
                            }])
                            
                            if novo_lanc is not None:
                                lanc_id = novo_lanc[0]['id'] if isinstance(novo_lanc, list) and len(novo_lanc)>0 else None
                                sb_request("inscricao_pagamentos", "PATCH", {"status": "Aprovado", "lancamento_id": lanc_id}, filtros={"id": f"eq.{p['id']}"})
                                novo_valor_pago = float(insc.get('valor_pago') or 0) + float(p.get('valor') or 0)
                                novo_status = "Completo" if novo_valor_pago >= float(insc.get('valor_total') or 0) - 0.01 else "Parcial"
                                sb_request("inscricoes", "PATCH", {"valor_pago": novo_valor_pago, "status_pagamento": novo_status}, filtros={"id": f"eq.{insc['id']}"})
                                st.cache_data.clear(); st.rerun()

                        if col_b.button("❌", key=f"rejeitar_pg_{p['id']}", help="Rejeitar"):
                            sb_request("inscricao_pagamentos", "PATCH", {"status": "Rejeitado"}, filtros={"id": f"eq.{p['id']}"})
                            st.cache_data.clear(); st.rerun()
                        st.markdown("<hr style='margin:6px 0;border-color:#E2E8F0;'>", unsafe_allow_html=True)

        with tab_lista:
            if not ev_tem_participantes:
                st.info("💡 Este projeto é do tipo Campanha/Arrecadação e não possui lista nominal de participantes. Acompanhe os totais diretamente na aba Visão Consolidada ou no Analytics.")
            else:
                st.markdown("### Participantes Inscritos")
                
                with st.expander("➕ Cadastrar Participante Manualmente"):
                    with st.form("form_cad_participante"):
                        p_nome = st.text_input("Nome Completo")
                        p_tel = st.text_input("Telefone / WhatsApp")
                        p_cpf = st.text_input("CPF (Opcional)")
                        p_valor = st.number_input("Valor da Inscrição (R$)", value=float(evento_obj_sel.get('valor_inscricao') or 0), format="%.2f")
                        
                        if st.form_submit_button("Cadastrar Inscrição", use_container_width=True):
                            if not p_nome:
                                st.warning("Informe o nome do participante.")
                            else:
                                sb_request("inscricoes", "POST", [{
                                    "evento_id": evento_id_sel,
                                    "nome_participante": p_nome,
                                    "contato": p_tel,
                                    "cpf": p_cpf if p_cpf else None,
                                    "valor_total": float(p_valor),
                                    "valor_pago": 0,
                                    "status_pagamento": "Pendente"
                                }])
                                st.cache_data.clear(); st.success("Inscrição cadastrada!"); time.sleep(1); st.rerun()

                if not inscricoes_evento:
                    st.info("Nenhuma inscrição cadastrada para este evento.")
                else:
                    # Mapeamento auxiliar para trazer a origem e as datas reais das parcelas
                    todos_lancamentos = carregar("lancamentos") or []
                    map_lancamentos = {str(l['id']): l for l in todos_lancamentos}
                    map_cred_por_lanc = {str(c.get('lancamento_id')): c for c in creditos_ofx_all if c.get('lancamento_id')}

                    for insc in inscricoes_evento:
                        pgs = [p for p in pagamentos_all if p.get("inscricao_id") == insc["id"]]
                        emoji_status = {"Pendente": "⏳", "Parcial": "🟡", "Completo": "✅"}.get(insc.get("status_pagamento"), "⏳")
                        
                        st.markdown(f"**{insc['nome_participante']}** ({insc.get('contato','—')}) — {emoji_status} {insc.get('status_pagamento','Pendente')} — Pago: {fmt_moeda(insc.get('valor_pago'))} / Total: {fmt_moeda(insc.get('valor_total'))}")
                        
                        for p in sorted(pgs, key=lambda x: x.get("numero_parcela", 1)):
                            lanc_id = str(p.get('lancamento_id'))
                            l_info = map_lancamentos.get(lanc_id, {})
                            c_info = map_cred_por_lanc.get(lanc_id, {})
                            
                            # Tratamento da Data do Pagamento (Via banco ou lançamento direto)
                            dt_pg = c_info.get('data') or l_info.get('data_competencia') or "N/I"
                            if dt_pg != "N/I":
                                try: dt_pg = pd.to_datetime(dt_pg).strftime('%d/%m/%Y')
                                except: pass
                            
                            # Tratamento da Data de Associação (Quando o líder apertou o botão no sistema)
                            dt_assoc = p.get('created_at', "N/I")
                            if dt_assoc != "N/I":
                                try:
                                    dt_obj = pd.to_datetime(dt_assoc)
                                    dt_assoc = dt_obj.strftime('%d/%m/%Y %H:%M')
                                except: pass
                                
                            # Identificação da pessoa original que enviou o PIX ou upload web
                            desc_origem = c_info.get('descricao_bancaria') or l_info.get('descricao') or "Comprovante anexado via web"
                            
                            st.caption(f" • Parcela {p.get('numero_parcela',1)}: {fmt_moeda(p.get('valor'))} — {p['status']}")
                            st.caption(f"&nbsp;&nbsp;&nbsp;&nbsp;↳ 📅 **Pgto:** {dt_pg} | 🔗 **Vinculado em:** {dt_assoc}")
                            st.caption(f"&nbsp;&nbsp;&nbsp;&nbsp;↳ 📄 **Origem:** {desc_origem}")
                            
                        st.markdown("<hr style='margin:6px 0;border-color:#E2E8F0;'>", unsafe_allow_html=True)

# ==========================================
# METAS E ORÇAMENTOS
# ==========================================
elif page == "Metas e Orçamentos":
    st.title("Metas e Orçamentos")
    tab1, tab2 = st.tabs(["🎯 Metas Mensais", "📦 Orçamento por Categoria"])

    with tab1:
        st.markdown("Defina os alvos de arrecadação e o teto de despesas para cada mês. (Salvar um mês existente irá atualizá-lo).")
        col_f1, col_f2 = st.columns([2, 1])
        
        with col_f1:
            with st.form("form_meta", clear_on_submit=True):
                c1, c2, c3, c4 = st.columns(4)
                ano_meta = c1.number_input("Ano", min_value=2020, max_value=2100, value=date.today().year)
                mes_meta = c2.selectbox("Mês", list(range(1, 13)), format_func=lambda m: MESES_PT[m-1], index=date.today().month-1)
                meta_entradas = c3.number_input("Meta Entradas (R$)", min_value=0.0, format="%.2f")
                meta_saidas = c4.number_input("Teto Saídas (R$)", min_value=0.0, format="%.2f")
                if st.form_submit_button("Salvar Meta", use_container_width=True):
                    upsert_meta(int(ano_meta), int(mes_meta), float(meta_entradas), float(meta_saidas))
                    st.success("Meta salva!"); st.rerun()
                    
        with col_f2:
            metas_lista = carregar("metas_mensais")
            with st.expander("🗑️ Excluir Meta"):
                if metas_lista:
                    opcoes_metas = {f"{m['ano']} - {MESES_PT[int(m['mes'])-1]}": m['id'] for m in metas_lista}
                    meta_del = st.selectbox("Selecione a Meta", list(opcoes_metas.keys()))
                    if st.button("Excluir Meta Selecionado", use_container_width=True):
                        sb_request("metas_mensais", "DELETE", filtros={"id": f"eq.{opcoes_metas[meta_del]}"})
                        st.cache_data.clear(); st.success("Excluída!"); time.sleep(1); st.rerun()

        if metas_lista:
            df_metas_view = pd.DataFrame(metas_lista)
            df_metas_view['Mês'] = df_metas_view['mes'].apply(lambda m: MESES_PT[int(m)-1])
            df_metas_view = df_metas_view[['ano', 'Mês', 'meta_entradas', 'meta_saidas']]
            df_metas_view.columns = ['Ano', 'Mês', 'Meta Entradas', 'Meta Saídas']
            st.dataframe(df_metas_view.sort_values(['Ano', 'Mês']), use_container_width=True, hide_index=True)

    with tab2:
        st.markdown("Defina o teto de gastos anual por categoria.")
        cats_saida = [c for c in categorias_db if c['tipo'] == 'Saída']
        col_o1, col_o2 = st.columns([2, 1])
        
        with col_o1:
            with st.form("form_orcamento", clear_on_submit=True):
                c1, c2, c3 = st.columns(3)
                ano_orc = c1.number_input("Ano", min_value=2020, max_value=2100, value=date.today().year, key="ano_orc")
                cat_orc = c2.selectbox("Categoria", [c['nome'] for c in cats_saida])
                valor_orc = c3.number_input("Valor Orçado (R$)", min_value=0.0, format="%.2f")
                if st.form_submit_button("Salvar Orçamento", use_container_width=True):
                    cat_id = next(c['id'] for c in cats_saida if c['nome'] == cat_orc)
                    upsert_orcamento(int(ano_orc), cat_id, float(valor_orc))
                    st.success("Orçamento salvo!"); st.rerun()

        with col_o2:
            orcamentos_lista = carregar("orcamentos_categoria")
            with st.expander("🗑️ Excluir Orçamento"):
                if orcamentos_lista:
                    map_cat_nome_all = {str(c['id']): c['nome'] for c in categorias_db}
                    opcoes_orc = {f"{o['ano']} - {map_cat_nome_all.get(str(o['categoria_id']))}": o['id'] for o in orcamentos_lista}
                    orc_del = st.selectbox("Selecione o Orçamento", list(opcoes_orc.keys()))
                    if st.button("Excluir Orçamento Selecionado", use_container_width=True):
                        sb_request("orcamentos_categoria", "DELETE", filtros={"id": f"eq.{opcoes_orc[orc_del]}"})
                        st.cache_data.clear(); st.success("Excluído!"); time.sleep(1); st.rerun()

        if orcamentos_lista:
            df_orc_view = pd.DataFrame(orcamentos_lista)
            df_orc_view['Categoria'] = df_orc_view['categoria_id'].astype(str).map(map_cat_nome_all)
            df_orc_view = df_orc_view[['ano', 'Categoria', 'valor_orcado']]
            df_orc_view.columns = ['Ano', 'Categoria', 'Valor Orçado']
            st.dataframe(df_orc_view.sort_values('Ano'), use_container_width=True, hide_index=True)

# ==========================================
# CATEGORIAS
# ==========================================
elif page == "Categorias":
    st.title("Gestão de Categorias")
    st.markdown("Associe as categorias gerenciais ao plano de contas contábil.")

    col_nova, col_edit = st.columns(2)
    with col_nova:
        with st.expander("➕ Nova Categoria", expanded=True):
            with st.form("form_categoria", clear_on_submit=True):
                nome_cat = st.text_input("Nome da Categoria")
                tipo_cat = st.selectbox("Tipo", ["Entrada", "Saída"])
                codigo_cat = st.text_input("Código Contábil (opcional)")
                if st.form_submit_button("Cadastrar Categoria", use_container_width=True):
                    if nome_cat:
                        res = sb_request("categorias", "POST", {"nome": nome_cat, "tipo": tipo_cat, "codigo_contabil": codigo_cat or None})
                        if res is not None:
                            st.cache_data.clear()
                            st.success("Categoria criada!")
                            time.sleep(1); st.rerun()
                    else:
                        st.warning("Informe o nome.")

    with col_edit:
        with st.expander("✏️ Editar ou Excluir Categoria"):
            if categorias_db:
                cat_map = {f"{c['nome']} ({c['tipo']})": c for c in categorias_db}
                cat_sel = st.selectbox("Selecione a Categoria", list(cat_map.keys()))
                cat_data = cat_map[cat_sel]
                
                with st.form("form_edit_cat"):
                    novo_nome = st.text_input("Nome", value=cat_data['nome'])
                    novo_tipo = st.selectbox("Tipo", ["Entrada", "Saída"], index=0 if cat_data['tipo'] == "Entrada" else 1)
                    novo_cod = st.text_input("Código Contábil", value=cat_data.get('codigo_contabil') or "")
                    
                    c_btn1, c_btn2 = st.columns(2)
                    btn_atualizar = c_btn1.form_submit_button("💾 Atualizar", use_container_width=True)
                    btn_excluir = c_btn2.form_submit_button("🗑️ Excluir", use_container_width=True)
                    
                    if btn_atualizar:
                        res = sb_request("categorias", "PATCH", {"nome": novo_nome, "tipo": novo_tipo, "codigo_contabil": novo_cod or None}, filtros={"id": f"eq.{cat_data['id']}"})
                        if res is not None:
                            st.cache_data.clear(); st.success("Atualizado!"); time.sleep(1); st.rerun()
                    if btn_excluir:
                        res = sb_request("categorias", "DELETE", filtros={"id": f"eq.{cat_data['id']}"})
                        if res is not None:
                            st.cache_data.clear(); st.success("Excluída!"); time.sleep(1); st.rerun()
            else:
                st.info("Nenhuma categoria cadastrada.")

    st.markdown("---")
    if categorias_db:
        df_cats = pd.DataFrame(categorias_db)[['nome', 'tipo', 'codigo_contabil']]
        df_cats.columns = ['Categoria', 'Natureza', 'Código Contábil']
        st.dataframe(df_cats, use_container_width=True, hide_index=True)

# ==========================================
# ANALYTICS FINANCEIRO
# ==========================================
elif page == "Analytics Financeiro":
    st.title("Analytics Financeiro")
    df = carregar_lancamentos_df()

    if df.empty:
        st.info("Nenhum lançamento registrado ainda.")
    else:
        st.subheader("Filtros")
        col_f1, col_f2 = st.columns(2)
        meses_disp = sorted(df['mes_ano'].dropna().unique(), reverse=True)
        mes_sel = col_f1.multiselect("Selecionar Mês/Ano", meses_disp, default=meses_disp[:6] if len(meses_disp) >= 6 else meses_disp)
        ignorar_eventos = col_f2.checkbox("Ocultar movimentações de Eventos", value=True)

        df_filtrado = df[df['mes_ano'].isin(mes_sel)]
        if ignorar_eventos:
            df_filtrado = df_filtrado[df_filtrado['centro_custo'].isnull()]

        entradas = df_filtrado[df_filtrado['tipo'] == 'Entrada']['valor'].sum()
        saidas = df_filtrado[df_filtrado['tipo'] == 'Saída']['valor'].sum()
        reserva = sum(float(c.get('saldo_inicial') or 0) for c in contas_bancarias_db if c.get('tipo') == 'Investimento')
        media_saidas = saidas / len(mes_sel) if len(mes_sel) > 0 and saidas > 0 else 1
        meses_reserva = reserva / media_saidas if media_saidas > 0 else 0

        col_k1, col_k2, col_k3, col_k4 = st.columns(4)
        col_k1.metric("Total Entradas", fmt_moeda(entradas))
        col_k2.metric("Total Saídas", fmt_moeda(saidas))
        col_k3.metric("Resultado", fmt_moeda(entradas - saidas))
        col_k4.metric("Fôlego de Caixa", f"{meses_reserva:.1f} Meses")

        st.markdown("---")
        st.subheader("Evolução Mensal (Entradas vs Saídas vs Metas)")
        if not df_filtrado.empty:
            df_agrupado = df_filtrado.groupby(['mes_ano', 'tipo'])['valor'].sum().unstack(fill_value=0)
            df_agrupado = df_agrupado.sort_index()

            if 'Entrada' not in df_agrupado.columns:
                df_agrupado['Entrada'] = 0.0
            if 'Saída' not in df_agrupado.columns:
                df_agrupado['Saída'] = 0.0

            metas_lista = carregar("metas_mensais")
            df_metas = pd.DataFrame(metas_lista)
            metas_map = {}
            if not df_metas.empty:
                df_metas['mes_ano'] = df_metas.apply(lambda r: f"{int(r['ano'])}-{int(r['mes']):02d}", axis=1)
                metas_map = df_metas.set_index('mes_ano').to_dict('index')

            fig = go.Figure()
            fig.add_bar(x=df_agrupado.index, y=df_agrupado['Entrada'], name='Entradas', marker_color='#2563EB')
            fig.add_bar(x=df_agrupado.index, y=df_agrupado['Saída'], name='Saídas', marker_color='#EF4444')
            
            if metas_map:
                fig.add_scatter(x=df_agrupado.index, y=[metas_map.get(m, {}).get('meta_entradas') for m in df_agrupado.index],
                               name='Meta Entradas', mode='lines+markers', line=dict(color='#1E40AF', dash='dot'))
                fig.add_scatter(x=df_agrupado.index, y=[metas_map.get(m, {}).get('meta_saidas') for m in df_agrupado.index],
                               name='Teto Saídas', mode='lines+markers', line=dict(color='#991B1B', dash='dot'))
            fig.update_layout(barmode='group', plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', font_color='#1E293B', legend_title_text='')
            st.plotly_chart(fig, use_container_width=True)
        st.markdown("---")
        col_g1, col_g2 = st.columns(2)

        with col_g1:
            st.subheader("Destino das Saídas")
            df_saidas_pizza = df_filtrado[df_filtrado['tipo'] == 'Saída']
            if not df_saidas_pizza.empty:
                df_pizza = df_saidas_pizza.groupby('categoria_nome')['valor'].sum().reset_index()
                fig_pie = px.pie(df_pizza, values='valor', names='categoria_nome', hole=0.5, color_discrete_sequence=px.colors.qualitative.Pastel)
                fig_pie.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', font_color='#1E293B')
                fig_pie.update_traces(textposition='inside', textinfo='percent+label')
                st.plotly_chart(fig_pie, use_container_width=True)
            else:
                st.info("Sem saídas no período.")

        with col_g2:
            st.subheader("Orçado vs Realizado (Ano)")
            ano_atual = date.today().year
            orcamentos_ano = [o for o in carregar("orcamentos_categoria") if int(o['ano']) == ano_atual]
            if orcamentos_ano:
                map_cat_nome = {str(c['id']): c['nome'] for c in categorias_db}
                for o in orcamentos_ano:
                    cat_nome = map_cat_nome.get(str(o['categoria_id']), 'Desconhecida')
                    realizado = df[(df['categoria_nome'] == cat_nome) & (df['data_competencia'].dt.year == ano_atual) & (df['tipo'] == 'Saída')]['valor'].sum()
                    orcado = float(o['valor_orcado']) or 1
                    pct = min(realizado / orcado, 1.0)
                    st.write(f"**{cat_nome}** — {fmt_moeda(realizado)} / {fmt_moeda(orcado)} ({realizado/orcado*100:.0f}%)")
                    st.progress(pct)
            else:
                st.info("Cadastre orçamentos em 'Metas e Orçamentos' para ver esta análise.")

        st.markdown("---")
        st.subheader("DRE por Evento / Projeto")
        df_eventos_fin = df[df['centro_custo'].notna() & (df['centro_custo'] != '')]
        if not df_eventos_fin.empty:
            resumo = df_eventos_fin.groupby(['centro_custo', 'tipo'])['valor'].sum().unstack(fill_value=0)
            resumo['Resultado'] = resumo.get('Entrada', 0) - resumo.get('Saída', 0)
            st.dataframe(resumo.reset_index().rename(columns={'centro_custo': 'Evento/Projeto'}), use_container_width=True, hide_index=True)
        else:
            st.info("Nenhuma movimentação vinculada a eventos ainda.")

# ==========================================
# EXPORTAR CONTABILIDADE
# ==========================================
elif page == "Exportar Contabilidade":
    st.title("Exportar para Contabilidade")
    st.markdown("Gere o arquivo consolidado do período e baixe um pacote `.zip` contendo os lançamentos em Excel/CSV junto com todos os comprovantes e notas fiscais anexadas.")

    df = carregar_lancamentos_df()
    if df.empty:
        st.info("Nenhum lançamento para exportar.")
    else:
        anos_disp = sorted(df['data_competencia'].dt.year.dropna().unique().astype(int), reverse=True)
        col1, col2 = st.columns(2)
        ano_exp = col1.selectbox("Ano", anos_disp, key="ano_exp_contab")
        mes_exp = col2.selectbox("Mês", ["Todos"] + MESES_PT, key="mes_exp_contab")

        dff = df[df['data_competencia'].dt.year == ano_exp].copy()
        if mes_exp != "Todos":
            idx_mes = MESES_PT.index(mes_exp) + 1
            dff = dff[dff['data_competencia'].dt.month == idx_mes]

        categorias = carregar_categorias()
        contas = carregar("contas_bancarias")
        eventos = carregar("eventos")

        map_cat_cod = {str(c['id']): c.get('codigo_contabil', '') for c in categorias}
        map_conta_nome = {str(c['id']): c['nome'] for c in contas}
        map_conta_cod = {str(c['id']): c.get('codigo_contabil', '') for c in contas}
        map_ev_cod = {e['nome']: e.get('codigo_receita_contabil', '') for e in eventos}

        dff['Cód. Contábil Categoria'] = dff['categoria_id'].astype(str).map(map_cat_cod).fillna('')
        dff['Conta Bancária Interna'] = dff['conta_bancaria_id'].astype(str).map(map_conta_nome).fillna('—')
        dff['Cód. Contábil Conta'] = dff['conta_bancaria_id'].astype(str).map(map_conta_cod).fillna('')
        dff['Cód. Contábil Evento'] = dff['centro_custo'].map(map_ev_cod).fillna('')

        exportar = dff[['data_competencia', 'tipo', 'descricao', 'categoria_nome', 'Cód. Contábil Categoria', 
                        'Conta Bancária Interna', 'Cód. Contábil Conta', 'centro_custo', 'Cód. Contábil Evento', 
                        'valor', 'status']].copy()
        
        exportar['data_competencia'] = exportar['data_competencia'].dt.strftime('%d/%m/%Y')
        exportar.columns = ['Data', 'Tipo', 'Descrição Interna', 'Categoria Interna', 'Cód. Contábil Categoria', 
                            'Conta Bancária Interna', 'Cód. Contábil Conta', 'Evento/Projeto Interno', 'Cód. Contábil Evento', 
                            'Valor (R$)', 'Status']
        
        st.dataframe(exportar, use_container_width=True, hide_index=True)

        col_d1, col_d2, col_d3 = st.columns(3)
        
        with col_d1:
            csv = exportar.to_csv(index=False, sep=';', decimal=',').encode('utf-8-sig')
            st.download_button("📥 Baixar CSV", csv, file_name=f"financeiro_{ano_exp}_{mes_exp}.csv", mime="text/csv", use_container_width=True)
            
        with col_d2:
            excel_data = to_excel_bytes({"Lançamentos": exportar})
            st.download_button("📥 Baixar Excel", excel_data, file_name=f"financeiro_{ano_exp}_{mes_exp}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                               
        with col_d3:
            if st.button("📦 Baixar Pacote ZIP (Excel + Anexos)", use_container_width=True):
                with st.spinner("Empacotando lançamentos e baixando anexos (isso pode demorar dependendo da quantidade de notas)..."):
                    zip_buffer = io.BytesIO()
                    
                    todos_anexos_exp = sb_request("lancamento_anexos", "GET")
                    mapa_anexos_exp = {}
                    if todos_anexos_exp:
                        for ax in todos_anexos_exp:
                            l_id = ax['lancamento_id']
                            if l_id not in mapa_anexos_exp:
                                mapa_anexos_exp[l_id] = []
                            mapa_anexos_exp[l_id].append(ax)

                    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                        zip_file.writestr(f"financeiro_{ano_exp}_{mes_exp}.xlsx", excel_data)
                        
                        anexos_adicionados = 0
                        for _, row_orig in dff.iterrows():
                            row_id = str(row_orig['id'])
                            anexos_row = mapa_anexos_exp.get(row_id, [])
                            
                            if not anexos_row and row_orig.get('url_anexo') and isinstance(row_orig.get('url_anexo'), str) and row_orig.get('url_anexo').strip():
                                anexos_row = [{'url_storage': row_orig.get('url_anexo'), 'nome_original': row_orig.get('url_anexo').split('/')[-1]}]
                            
                            for ax in anexos_row:
                                path_anexo = ax['url_storage']
                                try:
                                    file_res = supabase.storage.from_("comprovantes").download(path_anexo)
                                    if file_res:
                                        nome_original = ax.get('nome_original') or path_anexo.split('/')[-1]
                                        data_str = pd.to_datetime(row_orig['data_competencia']).strftime('%Y%m%d')
                                        nome_no_zip = f"comprovantes/{data_str}_{row_id[:5]}_{nome_original}"
                                        
                                        zip_file.writestr(nome_no_zip, file_res)
                                        anexos_adicionados += 1
                                except Exception:
                                    pass
                                    
                    zip_buffer.seek(0)
                    st.success(f"Pacote gerado com sucesso! ({anexos_adicionados} anexos incluídos).")
                    
                    st.download_button(
                        label="⬇️ Clique aqui para salvar o ZIP",
                        data=zip_buffer.getvalue(),
                        file_name=f"Contabilidade_Pacote_{ano_exp}_{mes_exp}.zip",
                        mime="application/zip",
                        use_container_width=True
                    )

# ==========================================
# GESTÃO DE USUÁRIOS
# ==========================================
elif page == "Gestão de Usuários":
    st.title("Gestão de Usuários e Perfis")
    st.markdown("Crie acessos e defina as permissões para a Tesouraria, Conselho ou Líderes de Eventos.")

    col_nova, col_edit = st.columns(2)
    with col_nova:
        with st.expander("➕ Novo Usuário", expanded=True):
            with st.form("form_usuario", clear_on_submit=True):
                n_nome = st.text_input("Nome Completo *")
                n_email = st.text_input("Email *")
                n_tel = st.text_input("Telefone (WhatsApp) *")
                n_cpf = st.text_input("CPF (Opcional)")
                n_perfil = st.selectbox("Perfil de Acesso *", ["Visão Total Tesouraria", "Visão Conselho", "Visão Eventos"])
                
                if st.form_submit_button("Cadastrar Usuário", use_container_width=True, type="primary"):
                    if not n_nome or not n_email or not n_tel:
                        st.warning("⚠️ Nome, Email e Telefone são obrigatórios.")
                    else:
                        payload = {
                            "nome": n_nome, "email": n_email, "telefone": n_tel, 
                            "cpf": n_cpf if n_cpf else None, "perfil": n_perfil, "status": "Ativo"
                        }
                        res = sb_request("usuarios", "POST", [payload])
                        if res is not None:
                            st.cache_data.clear()
                            st.success("Usuário cadastrado com sucesso!")
                            time.sleep(1)
                            st.rerun()

    with col_edit:
        with st.expander("✏️ Editar ou Excluir Usuário"):
            if usuarios_db:
                user_opcoes = {f"{u['nome']} ({u['perfil']})": u for u in usuarios_db}
                user_sel = st.selectbox("Selecione o Usuário", list(user_opcoes.keys()))
                u_data = user_opcoes[user_sel]

                with st.form("form_edit_user"):
                    e_nome = st.text_input("Nome Completo", value=u_data.get('nome', ''))
                    e_email = st.text_input("Email", value=u_data.get('email', ''))
                    e_tel = st.text_input("Telefone", value=u_data.get('telefone', ''))
                    e_cpf = st.text_input("CPF", value=u_data.get('cpf', ''))
                    
                    perfis_lista = ["Visão Total Tesouraria", "Visão Conselho", "Visão Eventos"]
                    idx_perfil = perfis_lista.index(u_data.get('perfil')) if u_data.get('perfil') in perfis_lista else 0
                    e_perfil = st.selectbox("Perfil de Acesso", perfis_lista, index=idx_perfil)
                    
                    c1, c2 = st.columns(2)
                    btn_upd = c1.form_submit_button("💾 Atualizar", use_container_width=True)
                    btn_del = c2.form_submit_button("🗑️ Excluir", use_container_width=True)

                    if btn_upd:
                        payload = {"nome": e_nome, "email": e_email, "telefone": e_tel, "cpf": e_cpf, "perfil": e_perfil}
                        sb_request("usuarios", "PATCH", payload, filtros={"id": f"eq.{u_data['id']}"})
                        st.cache_data.clear(); st.success("Atualizado!"); time.sleep(1); st.rerun()
                    if btn_del:
                        sb_request("usuarios", "DELETE", filtros={"id": f"eq.{u_data['id']}"})
                        st.cache_data.clear(); st.success("Excluído!"); time.sleep(1); st.rerun()
            else:
                st.info("Nenhum usuário cadastrado.")

    st.markdown("---")
    st.markdown("### Usuários Cadastrados")
    if usuarios_db:
        df_users = pd.DataFrame(usuarios_db)[['nome', 'email', 'telefone', 'perfil']]
        df_users.columns = ['Nome', 'Email', 'Telefone', 'Perfil']
        st.dataframe(df_users, use_container_width=True, hide_index=True)
