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

# ==========================================
# CONFIGURAÇÃO DA PÁGINA E DESIGN SYSTEM
# ==========================================
st.set_page_config(page_title="Financeiro COMUNA", page_icon="⛪", layout="wide", initial_sidebar_state="expanded")

MESES_PT = ["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho",
            "Agosto","Setembro","Outubro","Novembro","Dezembro"]

st.markdown("""
    <style>
    @import url('[fonts.googleapis.com](https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap)');
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
    </style>
""", unsafe_allow_html=True)

# ==========================================
# CONEXÃO COM SUPABASE
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
        if metodo == "GET":
            res = requests.get(url, headers=headers, params=filtros or {})
            return res.json() if res.status_code == 200 else []
        elif metodo == "POST":
            res = requests.post(url, headers=headers, json=payload)
            return res.json() if res.status_code in (200, 201) else None
        elif metodo == "PATCH":
            res = requests.patch(url, headers=headers, params=filtros or {}, json=payload)
            return res.json() if res.status_code in (200, 201, 204) else None
        elif metodo == "DELETE":
            res = requests.delete(url, headers=headers, params=filtros or {})
            return res.status_code in (200, 204)
    except Exception as e:
        return [] if metodo == "GET" else None

@st.cache_data(ttl=30)
def carregar(tabela):
    return sb_request(tabela, "GET") or []

@st.cache_data(ttl=30)
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
        data = sb_request("categorias", "GET") or []
    return data

def carregar_lancamentos_df():
    lanc = carregar("lancamentos")
    if not lanc:
        return pd.DataFrame()
    df = pd.DataFrame(lanc)
    df['valor'] = pd.to_numeric(df['valor'], errors='coerce').fillna(0.0)
    df['data_competencia'] = pd.to_datetime(df['data_competencia'], errors='coerce')
    df['data_vencimento'] = pd.to_datetime(df.get('data_vencimento'), errors='coerce') if 'data_vencimento' in df.columns else pd.NaT
    df['conciliado'] = df['conciliado'].fillna(False) if 'conciliado' in df.columns else False
    df['mes_ano'] = df['data_competencia'].dt.strftime('%Y-%m')

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
# PÁGINA PÚBLICA DE INSCRIÇÃO (sem login)
# ==========================================
def pagina_inscricao_publica():
    st.markdown("<h1 style='text-align:center;'>⛪ Inscrição em Evento</h1>", unsafe_allow_html=True)
    eventos_abertos = [e for e in carregar("eventos") if e.get("status") == "Aberto"]
    if not eventos_abertos:
        st.info("Não há eventos com inscrições abertas no momento.")
        return

    qp = st.query_params
    evento_id_param = qp.get("evento")
    nomes = [e["nome"] for e in eventos_abertos]
    default_idx = 0
    if evento_id_param:
        for i, e in enumerate(eventos_abertos):
            if str(e["id"]) == str(evento_id_param):
                default_idx = i
                break

    evento_sel_nome = st.selectbox("Selecione o Evento", nomes, index=default_idx)
    evento = next(e for e in eventos_abertos if e["nome"] == evento_sel_nome)

    st.markdown(f"""
    <div style="background:#FFFFFF;border:1px solid #E2E8F0;border-radius:14px;padding:20px;margin-bottom:20px;">
        <p style="margin:0;color:#475569;">📅 Data: <b>{evento.get('data_evento','—')}</b></p>
        <p style="margin:0;color:#475569;">💰 Valor da Inscrição: <b>{fmt_moeda(evento.get('valor_inscricao'))}</b></p>
        <p style="margin:0;color:#475569;">🔑 Chave Pix: <b>{evento.get('chave_pix','—')}</b></p>
    </div>
    """, unsafe_allow_html=True)

    with st.form("form_inscricao_publica", clear_on_submit=True):
        nome = st.text_input("Seu nome completo")
        contato = st.text_input("Telefone / WhatsApp")
        comprovante = st.file_uploader("Anexar comprovante do Pix", type=['png','jpg','jpeg','pdf'])
        enviar = st.form_submit_button("Enviar Inscrição", use_container_width=True)
        if enviar:
            if not nome or not contato:
                st.warning("Preencha nome e contato.")
            else:
                url_comp = comprimir_e_fazer_upload(comprovante, pasta="eventos") if comprovante else None
                sb_request("inscricoes", "POST", {
                    "evento_id": evento["id"],
                    "nome_participante": nome,
                    "contato": contato,
                    "valor": float(evento.get("valor_inscricao") or 0),
                    "comprovante_url": url_comp,
                    "status": "Pendente"
                })
                st.success("✅ Inscrição enviada! A equipe irá validar seu pagamento em breve.")

# Verifica se é acesso público (link de inscrição) ANTES de montar o painel interno
qp = st.query_params
if qp.get("pagina") == "inscricao":
    st.markdown("<style>[data-testid='stSidebar'], [data-testid='collapsedControl'] {display:none;}</style>", unsafe_allow_html=True)
    pagina_inscricao_publica()
    st.stop()

# ==========================================
# DADOS GLOBAIS DO PAINEL INTERNO
# ==========================================
categorias_db = carregar_categorias()
eventos_db = carregar("eventos")
contas_bancarias_db = carregar("contas_bancarias")

if "page" not in st.session_state:
    st.session_state.page = "Resumo do Dia"

def secao(nome):
    st.sidebar.markdown(f"<p style='color:#94A3B8;font-size:0.72rem;font-weight:700;letter-spacing:0.08em;margin:18px 0 6px 4px;'>{nome}</p>", unsafe_allow_html=True)

def nav_button(label, icon):
    ativo = st.session_state.page == label
    if st.sidebar.button(f"{icon}  {label}", key=f"nav_{label}", use_container_width=True, type="primary" if ativo else "secondary"):
        st.session_state.page = label
        st.rerun()

st.sidebar.markdown("<h2 style='color:#0F172A;font-weight:800;padding-top:6px;'>⛪ COMUNA</h2>", unsafe_allow_html=True)

secao("OPERACIONAL")
nav_button("Resumo do Dia", "🏠")
nav_button("Tesouraria", "💰")
nav_button("Conciliação Bancária", "🏦")

secao("GESTÃO DE EVENTOS")
nav_button("Painel de Eventos", "🎫")
nav_button("Inscrições e Comprovantes", "✅")

secao("ESTRATÉGICO")
nav_button("Metas e Orçamentos", "🎯")
nav_button("Categorias", "🏷️")

secao("RELATÓRIOS")
nav_button("Analytics Financeiro", "📊")
nav_button("Exportar Contabilidade", "📤")

st.sidebar.markdown("---")
st.sidebar.caption("Gestão Financeira • v4.0")

page = st.session_state.page

# ==========================================
# RESUMO DO DIA
# ==========================================
if page == "Resumo do Dia":
    st.title("Resumo do Dia")
    st.markdown(f"Hoje é {date.today().strftime('%d/%m/%Y')}. Aqui está o que precisa da sua atenção.")

    df = carregar_lancamentos_df()
    hoje = pd.Timestamp(date.today())

    saldo_consolidado = sum(float(c.get('saldo_inicial') or 0) for c in contas_bancarias_db)
    if not df.empty:
        concluidos = df[df['status'] == 'Concluído']
        saldo_consolidado += concluidos[concluidos['tipo'] == 'Entrada']['valor'].sum()
        saldo_consolidado -= concluidos[concluidos['tipo'] == 'Saída']['valor'].sum()

    pendentes = df[df['status'] == 'Pendente'] if not df.empty else pd.DataFrame()
    contas_hoje = pendentes[pendentes['data_vencimento'] == hoje] if not pendentes.empty else pd.DataFrame()
    atrasadas = pendentes[pendentes['data_vencimento'] < hoje] if not pendentes.empty else pd.DataFrame()
    pendentes_insc = [i for i in carregar("inscricoes") if i['status'] == 'Pendente']

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Saldo Consolidado", fmt_moeda(saldo_consolidado))
    c2.metric("A Pagar Hoje", str(len(contas_hoje)))
    c3.metric("Atrasadas", str(len(atrasadas)))
    c4.metric("Inscrições p/ Validar", str(len(pendentes_insc)))

    st.markdown("---")
    col_a, col_b = st.columns(2)

    with col_a:
        st.subheader("💳 Contas para pagar hoje e atrasadas")
        a_mostrar = pd.concat([atrasadas, contas_hoje]) if not atrasadas.empty or not contas_hoje.empty else pd.DataFrame()
        if a_mostrar.empty:
            st.success("Nenhuma pendência para hoje. 🎉")
        else:
            for _, row in a_mostrar.iterrows():
                cc1, cc2, cc3 = st.columns([3, 1.5, 1.3])
                cc1.write(row['descricao'])
                cc2.write(fmt_moeda(row['valor']))
                if cc3.button("✅ Pagar", key=f"resumo_pagar_{row['id']}"):
                    sb_request("lancamentos", "PATCH", {"status": "Concluído", "data_pagamento": str(date.today())}, filtros={"id": f"eq.{row['id']}"})
                    st.cache_data.clear()
                    st.rerun()

    with col_b:
        st.subheader("🎫 Inscrições aguardando aprovação")
        if not pendentes_insc:
            st.success("Nenhuma inscrição pendente.")
        else:
            for i in pendentes_insc[:6]:
                st.write(f"• **{i['nome_participante']}** — {fmt_moeda(i.get('valor'))}")
            if st.button("Ir para Inscrições e Comprovantes →"):
                st.session_state.page = "Inscrições e Comprovantes"
                st.rerun()

# ==========================================
# TESOURARIA
# ==========================================
elif page == "Tesouraria":
    st.title("Tesouraria")
    tab1, tab2, tab3 = st.tabs(["📝 Novo Lançamento", "⏳ Contas a Pagar/Receber", "📜 Histórico Completo"])

    with tab1:
        with st.form("form_lancamento", clear_on_submit=True):
            col1, col2, col3 = st.columns(3)
            tipo_lanc = col1.radio("Tipo", ["Entrada", "Saída"], horizontal=True)
            valor = col2.number_input("Valor (R$)", min_value=0.0, step=50.0, format="%.2f")
            data_comp = col3.date_input("Data", date.today())

            cats_filtradas = [c for c in categorias_db if c.get("tipo") == tipo_lanc]
            opcoes_cats = {c["nome"]: c["id"] for c in cats_filtradas}

            col4, col5 = st.columns([2, 1])
            descricao = col4.text_input("Descrição")
            categoria_sel = col5.selectbox("Categoria", list(opcoes_cats.keys()) if opcoes_cats else ["Cadastre uma categoria"])

            col6, col7, col8 = st.columns(3)
            status_lanc = col6.selectbox("Situação", ["Concluído", "Pendente"])
            data_venc = col7.date_input("Vencimento", data_comp) if status_lanc == "Pendente" else None
            contas_opcoes = {c["nome"]: c["id"] for c in contas_bancarias_db}
            conta_sel = col8.selectbox("Conta Bancária", ["Nenhuma"] + list(contas_opcoes.keys()))

            col9, col10 = st.columns(2)
            tag = col9.selectbox("Projeto / Evento", ["Nenhum"] + [e['nome'] for e in eventos_db])
            arquivo = col10.file_uploader("Comprovante / Nota Fiscal", type=['png', 'jpg', 'jpeg', 'pdf'])

            recorrente = st.checkbox("🔁 Lançamento recorrente (repete nos próximos meses)")
            repeticoes = 1
            if recorrente:
                repeticoes = st.number_input("Repetir por quantos meses (incluindo este)", min_value=2, max_value=24, value=3)

            submit = st.form_submit_button("💾 Salvar Lançamento", use_container_width=True)

            if submit:
                if valor <= 0 or not descricao or not opcoes_cats:
                    st.warning("⚠️ Preencha descrição, valor e categoria corretamente.")
                else:
                    with st.spinner("Salvando..."):
                        url_anexo = comprimir_e_fazer_upload(arquivo, pasta="notas") if arquivo else None
                        dados_base = {
                            "descricao": descricao, "tipo": tipo_lanc, "valor": float(valor),
                            "data_competencia": str(data_comp),
                            "data_vencimento": str(data_venc) if data_venc else None,
                            "status": status_lanc,
                            "categoria_id": opcoes_cats[categoria_sel],
                            "centro_custo": None if tag == "Nenhum" else tag,
                            "conta_bancaria_id": contas_opcoes.get(conta_sel),
                            "url_anexo": url_anexo,
                            "recorrente": recorrente
                        }
                        sb_request("lancamentos", "POST", dados_base)
                        if recorrente and repeticoes > 1:
                            for i in range(1, int(repeticoes)):
                                nova_data = (pd.Timestamp(data_comp) + pd.DateOffset(months=i)).date()
                                dados_rep = dict(dados_base)
                                dados_rep["data_competencia"] = str(nova_data)
                                dados_rep["data_vencimento"] = str(nova_data)
                                dados_rep["status"] = "Pendente"
                                dados_rep["url_anexo"] = None
                                sb_request("lancamentos", "POST", dados_rep)
                        st.cache_data.clear()
                        st.success("✅ Lançamento registrado com sucesso!")
                        time.sleep(1)
                        st.rerun()

    with tab2:
        df = carregar_lancamentos_df()
        pend = df[df['status'] == 'Pendente'].copy() if not df.empty else pd.DataFrame()
        if pend.empty:
            st.info("Nenhuma conta pendente. 🎉")
        else:
            pend = pend.sort_values('data_vencimento')
            hoje = pd.Timestamp(date.today())
            for _, row in pend.iterrows():
                venc = row['data_vencimento']
                if pd.isna(venc):
                    situacao = "⚪ Sem vencimento"
                elif venc < hoje:
                    situacao = "🔴 Atrasado"
                elif venc == hoje:
                    situacao = "🟡 Vence hoje"
                else:
                    situacao = "🟢 A vencer"
                c1, c2, c3, c4, c5 = st.columns([3, 1.4, 1.4, 1.4, 1.2])
                c1.write(f"**{row['descricao']}** — {row['categoria_nome']}")
                c2.write(fmt_moeda(row['valor']))
                c3.write(venc.strftime('%d/%m/%Y') if pd.notna(venc) else '—')
                c4.write(situacao)
                if c5.button("✅ Pagar", key=f"pagar_tab_{row['id']}"):
                    sb_request("lancamentos", "PATCH", {"status": "Concluído", "data_pagamento": str(date.today())}, filtros={"id": f"eq.{row['id']}"})
                    st.cache_data.clear()
                    st.rerun()

    with tab3:
        df = carregar_lancamentos_df()
        if df.empty:
            st.info("Nenhum lançamento registrado.")
        else:
            col1, col2 = st.columns(2)
            filtro_tipo = col1.multiselect("Tipo", ["Entrada", "Saída"], default=["Entrada", "Saída"])
            filtro_status = col2.multiselect("Situação", df['status'].unique().tolist(), default=df['status'].unique().tolist())
            dff = df[df['tipo'].isin(filtro_tipo) & df['status'].isin(filtro_status)].sort_values('data_competencia', ascending=False)
            exibir = dff[['data_competencia', 'tipo', 'descricao', 'categoria_nome', 'valor', 'status', 'centro_custo', 'conta_nome']].copy()
            exibir['data_competencia'] = exibir['data_competencia'].dt.strftime('%d/%m/%Y')
            exibir.columns = ['Data', 'Tipo', 'Descrição', 'Categoria', 'Valor (R$)', 'Situação', 'Projeto', 'Conta']
            st.dataframe(exibir, use_container_width=True, hide_index=True)

# ==========================================
# CONCILIAÇÃO BANCÁRIA
# ==========================================
elif page == "Conciliação Bancária":
    st.title("Conciliação Bancária")
    st.markdown("Confira os lançamentos com o extrato real do banco.")

    with st.expander("➕ Cadastrar nova conta bancária"):
        with st.form("form_conta", clear_on_submit=True):
            cc1, cc2, cc3 = st.columns(3)
            nome_conta = cc1.text_input("Nome da Conta (Ex: Itaú Principal)")
            tipo_conta = cc2.selectbox("Tipo", ["Corrente", "Poupança", "Investimento"])
            saldo_inicial = cc3.number_input("Saldo Inicial (R$)", min_value=0.0, format="%.2f")
            if st.form_submit_button("Cadastrar Conta"):
                if nome_conta:
                    sb_request("contas_bancarias", "POST", {"nome": nome_conta, "tipo": tipo_conta, "saldo_inicial": float(saldo_inicial)})
                    st.cache_data.clear()
                    st.success("Conta cadastrada!")
                    st.rerun()
                else:
                    st.warning("Informe o nome da conta.")

    if not contas_bancarias_db:
        st.info("Cadastre ao menos uma conta bancária acima para iniciar a conciliação.")
    else:
        conta_opcoes = {c["nome"]: c["id"] for c in contas_bancarias_db}
        conta_sel = st.selectbox("Selecione a Conta", list(conta_opcoes.keys()))
        conta_id = conta_opcoes[conta_sel]

        df = carregar_lancamentos_df()
        df_conta = df[(df.get('conta_bancaria_id') == conta_id) & (df['status'] == 'Concluído')] if not df.empty else pd.DataFrame()

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
                st.error(f"⚠️ Diferença encontrada: {fmt_moeda(diferenca)}. Revise os lançamentos abaixo.")

        st.markdown("---")
        st.subheader("Lançamentos desta conta")
        if df_conta.empty:
            st.info("Nenhum lançamento concluído nesta conta ainda.")
        else:
            for _, row in df_conta.sort_values('data_competencia', ascending=False).iterrows():
                c1, c2, c3, c4 = st.columns([3, 1.5, 1.5, 1])
                c1.write(row['descricao'])
                c2.write(row['data_competencia'].strftime('%d/%m/%Y'))
                c3.write(fmt_moeda(row['valor']) if row['tipo'] == 'Entrada' else f"-{fmt_moeda(row['valor'])}")
                marcado = bool(row.get('conciliado'))
                novo_valor = c4.checkbox("Conciliado", value=marcado, key=f"conc_{row['id']}")
                if novo_valor != marcado:
                    sb_request("lancamentos", "PATCH", {"conciliado": novo_valor}, filtros={"id": f"eq.{row['id']}"})
                    st.cache_data.clear()

# ==========================================
# PAINEL DE EVENTOS
# ==========================================
elif page == "Painel de Eventos":
    st.title("Painel de Eventos")
    st.markdown("Cadastre acampamentos, retiros e conferências para abrir inscrições.")

    with st.expander("➕ Criar Novo Evento", expanded=len(eventos_db) == 0):
        with st.form("form_evento", clear_on_submit=True):
            nome_ev = st.text_input("Nome do Evento (Ex: Acampamento de Adolescentes 2026)")
            col1, col2, col3 = st.columns(3)
            data_ev = col1.date_input("Data do Evento", date.today())
            valor_ev = col2.number_input("Valor da Inscrição (R$)", min_value=0.0, format="%.2f")
            vagas_ev = col3.number_input("Total de Vagas", min_value=0, step=1)
            chave_pix_ev = st.text_input("Chave Pix para recebimento")
            if st.form_submit_button("Criar Evento", use_container_width=True):
                if nome_ev:
                    sb_request("eventos", "POST", {
                        "nome": nome_ev, "data_evento": str(data_ev), "valor_inscricao": float(valor_ev),
                        "vagas_total": int(vagas_ev), "chave_pix": chave_pix_ev, "centro_custo": nome_ev,
                        "status": "Aberto"
                    })
                    st.cache_data.clear()
                    st.success("Evento criado! Copie o link de inscrição na lista abaixo.")
                    st.rerun()
                else:
                    st.warning("Informe o nome do evento.")

    st.markdown("---")
    st.subheader("Eventos Cadastrados")

    if not eventos_db:
        st.info("Nenhum evento cadastrado ainda.")
    else:
        inscricoes_all = carregar("inscricoes")
        for ev in eventos_db:
            insc_evento = [i for i in inscricoes_all if i.get('evento_id') == ev['id']]
            aprovadas = [i for i in insc_evento if i['status'] == 'Aprovado']
            pendentes_ev = [i for i in insc_evento if i['status'] == 'Pendente']
            arrecadado = sum(float(i.get('valor') or 0) for i in aprovadas)

            cor_status = '#059669' if ev['status'] == 'Aberto' else '#94A3B8'
            st.markdown(f"""
            <div style="background:#FFFFFF;border:1px solid #E2E8F0;border-radius:14px;padding:20px;margin-bottom:12px;">
                <h4 style="margin-top:0;">{ev['nome']} <span style="font-size:0.8rem;color:{cor_status};">● {ev['status']}</span></h4>
                <p style="color:#475569;margin:4px 0;">📅 {ev.get('data_evento','—')} • 💰 {fmt_moeda(ev.get('valor_inscricao'))} por pessoa</p>
                <p style="color:#475569;margin:4px 0;">👥 {len(aprovadas)} / {ev.get('vagas_total', 0)} vagas preenchidas • ⏳ {len(pendentes_ev)} aguardando aprovação</p>
                <p style="color:#059669;margin:4px 0;font-weight:600;">💵 Total arrecadado: {fmt_moeda(arrecadado)}</p>
            </div>
            """, unsafe_allow_html=True)

            c1, c2 = st.columns([3, 1])
            c1.code(f"?pagina=inscricao&evento={ev['id']}", language=None)
            c1.caption("Cole este trecho ao final do link do seu app e compartilhe com a igreja.")
            novo_status = "Encerrado" if ev['status'] == "Aberto" else "Aberto"
            if c2.button(f"{'🔒 Encerrar' if ev['status']=='Aberto' else '🔓 Reabrir'}", key=f"toggle_{ev['id']}"):
                sb_request("eventos", "PATCH", {"status": novo_status}, filtros={"id": f"eq.{ev['id']}"})
                st.cache_data.clear()
                st.rerun()

# ==========================================
# INSCRIÇÕES E COMPROVANTES
# ==========================================
elif page == "Inscrições e Comprovantes":
    st.title("Inscrições e Comprovantes")
    st.markdown("Valide os comprovantes de Pix enviados pelos participantes.")

    if not eventos_db:
        st.info("Cadastre um evento primeiro em 'Painel de Eventos'.")
    else:
        evento_opcoes = {e["nome"]: e["id"] for e in eventos_db}
        evento_sel = st.selectbox("Evento", list(evento_opcoes.keys()))
        evento_id_sel = evento_opcoes[evento_sel]

        inscricoes_evento = [i for i in carregar("inscricoes") if i.get('evento_id') == evento_id_sel]
        filtro_status = st.radio("Filtrar por situação", ["Pendente", "Aprovado", "Rejeitado"], horizontal=True)
        lista_filtrada = [i for i in inscricoes_evento if i['status'] == filtro_status]

        if not lista_filtrada:
            st.info(f"Nenhuma inscrição com status '{filtro_status}'.")
        else:
            for insc in lista_filtrada:
                c1, c2, c3, c4 = st.columns([2, 1, 1.5, 1.3])
                c1.write(f"**{insc['nome_participante']}**")
                c1.caption(insc.get('contato', ''))
                c2.write(fmt_moeda(insc.get('valor')))
                link = obter_link_arquivo(insc.get('comprovante_url'))
                c3.markdown(f"[📎 Ver comprovante]({link})" if link else "_Sem comprovante_")

                if filtro_status == "Pendente":
                    col_a, col_b = c4.columns(2)
                    if col_a.button("✅", key=f"aprovar_{insc['id']}", help="Aprovar"):
                        cat_evento_id = next((c['id'] for c in categorias_db if c['nome'] == 'Inscrições de Eventos'), None)
                        novo_lanc = sb_request("lancamentos", "POST", {
                            "descricao": f"Inscrição - {insc['nome_participante']} ({evento_sel})",
                            "tipo": "Entrada", "valor": float(insc.get('valor') or 0),
                            "data_competencia": str(date.today()), "status": "Concluído",
                            "categoria_id": cat_evento_id, "centro_custo": evento_sel
                        })
                        lanc_id = novo_lanc[0]['id'] if novo_lanc else None
                        sb_request("inscricoes", "PATCH", {"status": "Aprovado", "lancamento_id": lanc_id}, filtros={"id": f"eq.{insc['id']}"})
                        st.cache_data.clear()
                        st.rerun()
                    if col_b.button("❌", key=f"rejeitar_{insc['id']}", help="Rejeitar"):
                        sb_request("inscricoes", "PATCH", {"status": "Rejeitado"}, filtros={"id": f"eq.{insc['id']}"})
                        st.cache_data.clear()
                        st.rerun()
                st.markdown("<hr style='margin:6px 0;border-color:#E2E8F0;'>", unsafe_allow_html=True)

# ==========================================
# METAS E ORÇAMENTOS
# ==========================================
elif page == "Metas e Orçamentos":
    st.title("Metas e Orçamentos")
    tab1, tab2 = st.tabs(["🎯 Metas Mensais", "📦 Orçamento por Categoria"])

    with tab1:
        st.markdown("Defina os alvos de arrecadação e o teto de despesas para cada mês.")
        with st.form("form_meta", clear_on_submit=True):
            c1, c2, c3, c4 = st.columns(4)
            ano_meta = c1.number_input("Ano", min_value=2020, max_value=2100, value=date.today().year)
            mes_meta = c2.selectbox("Mês", list(range(1, 13)), format_func=lambda m: MESES_PT[m-1], index=date.today().month-1)
            meta_entradas = c3.number_input("Meta de Entradas (R$)", min_value=0.0, format="%.2f")
            meta_saidas = c4.number_input("Teto de Saídas (R$)", min_value=0.0, format="%.2f")
            if st.form_submit_button("Salvar Meta", use_container_width=True):
                upsert_meta(int(ano_meta), int(mes_meta), float(meta_entradas), float(meta_saidas))
                st.success("Meta salva!")
                st.rerun()

        metas_lista = carregar("metas_mensais")
        if metas_lista:
            df_metas_view = pd.DataFrame(metas_lista)
            df_metas_view['Mês'] = df_metas_view['mes'].apply(lambda m: MESES_PT[int(m)-1])
            df_metas_view = df_metas_view[['ano', 'Mês', 'meta_entradas', 'meta_saidas']]
            df_metas_view.columns = ['Ano', 'Mês', 'Meta Entradas', 'Meta Saídas']
            st.dataframe(df_metas_view.sort_values(['Ano', 'Mês']), use_container_width=True, hide_index=True)

    with tab2:
        st.markdown("Defina o teto de gastos anual por categoria (departamento).")
        cats_saida = [c for c in categorias_db if c['tipo'] == 'Saída']
        with st.form("form_orcamento", clear_on_submit=True):
            c1, c2, c3 = st.columns(3)
            ano_orc = c1.number_input("Ano", min_value=2020, max_value=2100, value=date.today().year, key="ano_orc")
            cat_orc = c2.selectbox("Categoria", [c['nome'] for c in cats_saida])
            valor_orc = c3.number_input("Valor Orçado (R$)", min_value=0.0, format="%.2f")
            if st.form_submit_button("Salvar Orçamento", use_container_width=True):
                cat_id = next(c['id'] for c in cats_saida if c['nome'] == cat_orc)
                upsert_orcamento(int(ano_orc), cat_id, float(valor_orc))
                st.success("Orçamento salvo!")
                st.rerun()

        orcamentos_lista = carregar("orcamentos_categoria")
        if orcamentos_lista:
            map_cat_nome_all = {str(c['id']): c['nome'] for c in categorias_db}
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
    st.markdown("Associe as categorias gerenciais ao plano de contas do escritório de contabilidade.")

    with st.expander("➕ Nova Categoria"):
        with st.form("form_categoria", clear_on_submit=True):
            c1, c2, c3 = st.columns(3)
            nome_cat = c1.text_input("Nome da Categoria")
            tipo_cat = c2.selectbox("Tipo", ["Entrada", "Saída"])
            codigo_cat = c3.text_input("Código Contábil (opcional)")
            if st.form_submit_button("Cadastrar Categoria"):
                if nome_cat:
                    sb_request("categorias", "POST", {"nome": nome_cat, "tipo": tipo_cat, "codigo_contabil": codigo_cat or None})
                    st.cache_data.clear()
                    st.success("Categoria criada!")
                    st.rerun()
                else:
                    st.warning("Informe o nome da categoria.")

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

            metas_lista = carregar("metas_mensais")
            df_metas = pd.DataFrame(metas_lista)
            metas_map = {}
            if not df_metas.empty:
                df_metas['mes_ano'] = df_metas.apply(lambda r: f"{int(r['ano'])}-{int(r['mes']):02d}", axis=1)
                metas_map = df_metas.set_index('mes_ano').to_dict('index')

            fig = go.Figure()
            fig.add_bar(x=df_agrupado.index, y=df_agrupado.get('Entrada', 0), name='Entradas', marker_color='#2563EB')
            fig.add_bar(x=df_agrupado.index, y=df_agrupado.get('Saída', 0), name='Saídas', marker_color='#EF4444')
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
    st.markdown("Gere o arquivo consolidado do período para envio ao escritório contábil.")

    df = carregar_lancamentos_df()
    if df.empty:
        st.info("Nenhum lançamento para exportar.")
    else:
        anos_disp = sorted(df['data_competencia'].dt.year.dropna().unique().astype(int), reverse=True)
        col1, col2 = st.columns(2)
        ano_exp = col1.selectbox("Ano", anos_disp)
        mes_exp = col2.selectbox("Mês", ["Todos"] + MESES_PT)

        dff = df[df['data_competencia'].dt.year == ano_exp]
        if mes_exp != "Todos":
            idx_mes = MESES_PT.index(mes_exp) + 1
            dff = dff[dff['data_competencia'].dt.month == idx_mes]

        exportar = dff[['data_competencia', 'tipo', 'descricao', 'categoria_nome', 'valor', 'status', 'centro_custo']].copy()
        exportar['data_competencia'] = exportar['data_competencia'].dt.strftime('%d/%m/%Y')
        exportar.columns = ['Data', 'Tipo', 'Descrição', 'Categoria', 'Valor', 'Status', 'Projeto']

        st.dataframe(exportar, use_container_width=True, hide_index=True)

        col_d1, col_d2 = st.columns(2)
        with col_d1:
            csv = exportar.to_csv(index=False, sep=';', decimal=',').encode('utf-8-sig')
            st.download_button("📥 Baixar CSV", csv, file_name=f"financeiro_{ano_exp}_{mes_exp}.csv", mime="text/csv", use_container_width=True)
        with col_d2:
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                exportar.to_excel(writer, index=False, sheet_name='Lançamentos')
            st.download_button("📥 Baixar Excel", buffer.getvalue(), file_name=f"financeiro_{ano_exp}_{mes_exp}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
