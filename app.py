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
    """Realiza a requisição e IMPRIME O ERRO se falhar."""
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
    """Recebe um dicionário de {NomeDaAba: DataFrame} e retorna bytes de um arquivo Excel"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet_name, df_data in dfs_dict.items():
            df_data.to_excel(writer, sheet_name=sheet_name)
    return output.getvalue()

@st.cache_data(ttl=15)
def carregar(tabela):
    return sb_request(tabela, "GET") or []

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
    
    # Prevenção caso a tabela exista mas esteja vazia e falte colunas
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

def safe_map_moeda(df):
    try:
        return df.map(fmt_moeda)
    except AttributeError:
        return df.applymap(fmt_moeda)

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
# MENU LATERAL INTERNO
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

st.sidebar.markdown("<h2 style='color:#0F172A;font-weight:800;padding-top:6px;'>⛪ Gestão Financeira </h2>", unsafe_allow_html=True)

secao("OPERACIONAL")
nav_button("Resumo do Dia", "🏠")
nav_button("Tesouraria", "💰")
nav_button("Conciliação Bancária", "🏦")
nav_button("Visão Consolidada", "📋")

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
st.sidebar.caption("Gestão Financeira • Final")

page = st.session_state.page

# ==========================================
# ==========================================
# ==========================================
# ==========================================
# RESUMO DO DIA 
# ==========================================
if page == "Resumo do Dia":
    st.title("Resumo do Dia")
    st.markdown(f"Hoje é {date.today().strftime('%d/%m/%Y')}. Aqui está o que precisa da sua atenção.")

    df = carregar_lancamentos_df()
    inscricoes_resumo = carregar("inscricoes")
    pagamentos_resumo = carregar("inscricao_pagamentos")
    eventos_resumo = carregar("eventos")

    hoje = pd.Timestamp(date.today())
    saldo_consolidado = sum(float(conta.get("saldo_inicial") or 0) for conta in contas_bancarias_db)

    if not df.empty:
        lancamentos_concluidos = df[df["status"] == "Concluído"]
        total_entradas = lancamentos_concluidos[lancamentos_concluidos["tipo"] == "Entrada"]["valor"].sum()
        total_saidas = lancamentos_concluidos[lancamentos_concluidos["tipo"] == "Saída"]["valor"].sum()
        saldo_consolidado += total_entradas - total_saidas

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
    col1.metric("Saldo Consolidado", fmt_moeda(saldo_consolidado))
    col2.metric("A Pagar Hoje", str(len(contas_hoje)))
    col3.metric("Contas Atrasadas", str(len(contas_atrasadas)))
    col4.metric("Recorrências Expirando", str(len(recorrencias_expirando)))

    if not recorrencias_expirando.empty:
        st.warning(f"⚠️ Atenção: Existem **{len(recorrencias_expirando)}** lançamentos recorrentes com data final de recorrência programada para este mês de {MESES_PT[hoje.month-1].lower()}. Verifique a necessidade de renovação.")

    st.markdown("---")
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
# ==========================================
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

    col_a1, col_a2 = st.columns([1, 3])
    ano_sel = col_a1.selectbox("Ano de Referência", anos_disp)

    df_ano = df[df['data_competencia'].dt.year == ano_sel].copy() if not df.empty else pd.DataFrame()

    tab_ex1, tab_ex2, tab_ex3 = st.tabs(["📥 Entradas por Categoria", "📤 Saídas por Categoria", "📊 Consolidado Bancário e Geral"])

    # Estruturas para exportação
    pivot_ent = pd.DataFrame()
    pivot_sai = pd.DataFrame()
    resumo_geral = pd.DataFrame(index=MESES_PT)

    with tab_ex1:
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
        pivot_ent.columns = MESES_PT
        pivot_ent['TOTAL ANUAL'] = pivot_ent.sum(axis=1)
        
        st.dataframe(safe_map_moeda(pivot_ent), use_container_width=True)

    with tab_ex2:
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
        pivot_sai.columns = MESES_PT
        pivot_sai['TOTAL ANUAL'] = pivot_sai.sum(axis=1)

        st.dataframe(safe_map_moeda(pivot_sai), use_container_width=True)

    with tab_ex3:
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

        st.dataframe(safe_map_moeda(resumo_geral), use_container_width=True)

    # 1. BOTÃO DE EXPORTAÇÃO GLOBAL
    st.markdown("---")
    st.markdown("### 📥 Exportar Consolidado")
    
    if not df_ano.empty:
        dfs_export = {
            "Entradas": pivot_ent,
            "Saídas": pivot_sai,
            "Resumo Geral": resumo_geral
        }
        excel_completo = to_excel_bytes(dfs_export)
        st.download_button(
            label="💾 Baixar Visão Completa (Excel)",
            data=excel_completo,
            file_name=f"Consolidado_{ano_sel}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    # 2. SEÇÃO DE DETALHAMENTO (DRILL-DOWN COM DOWNLOAD INDIVIDUAL)
    st.markdown("---")
    st.markdown("### 🔍 Detalhar Valores por Mês e Categoria")
    st.markdown("Selecione os filtros abaixo para ver detalhadamente quais itens compõem a soma e baixe os comprovantes de cada despesa diretamente.")

    # Estilo CSS compacto para ajustar a fonte e manter cada item em uma linha só
    st.markdown("""
        <style>
            .drill-row {
                font-size: 13px !important;
                padding: 2px 0px !important;
                white-space: nowrap !important;
                overflow: hidden !important;
                text-overflow: ellipsis !important;
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
            
            # Ordenação solicitada: 1º Saídas em cima e Entradas embaixo, 2º Data mais recente em cima
            df_detalhe['ordem_tipo'] = df_detalhe['tipo'].map({'Saída': 0, 'Entrada': 1})
            df_detalhe = df_detalhe.sort_values(by=['ordem_tipo', 'data_competencia'], ascending=[True, False])
            
            # Cabeçalho customizado invertendo Tipo (coluna 1) e Data (coluna 2)
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
                cols = st.columns([1, 1.2, 2.5, 2, 1.3, 1.2, 1.2])
                
                # Invertido: Coluna 0 = Tipo, Coluna 1 = Data
                cols[0].markdown(f"<div class='drill-row'>{row['tipo']}</div>", unsafe_allow_html=True)
                cols[1].markdown(f"<div class='drill-row'>{row['data_competencia'].strftime('%d/%m/%Y')}</div>", unsafe_allow_html=True)
                cols[2].markdown(f"<div class='drill-row'>{row['descricao'] or '—'}</div>", unsafe_allow_html=True)
                cols[3].markdown(f"<div class='drill-row'>{row['categoria_nome']}</div>", unsafe_allow_html=True)
                cols[4].markdown(f"<div class='drill-row'>{fmt_moeda(row['valor'])}</div>", unsafe_allow_html=True)
                cols[5].markdown(f"<div class='drill-row'>{row['conta_nome']}</div>", unsafe_allow_html=True)
                
                path_anexo = row.get('url_anexo')
                if path_anexo and isinstance(path_anexo, str) and path_anexo.strip():
                    try:
                        link_anexo = supabase.storage.from_("comprovantes").create_signed_url(path_anexo, 3600)
                        url_final = link_anexo.get("signedURL") or link_anexo.get("signed_url")
                        if url_final:
                            cols[6].markdown(f"<div class='drill-row'><a href='{url_final}' target='_blank'>📥 Baixar</a></div>", unsafe_allow_html=True)
                        else:
                            cols[6].markdown("<div class='drill-row'>—</div>", unsafe_allow_html=True)
                    except Exception:
                        cols[6].markdown("<div class='drill-row'>Indisponível</div>", unsafe_allow_html=True)
                else:
                    cols[6].markdown("<div class='drill-row'>Sem anexo</div>", unsafe_allow_html=True)
                    
                st.markdown("<hr style='margin:2px 0;border-color:#F1F5F9;'>", unsafe_allow_html=True)

            st.markdown("")
            exibir_export = df_detalhe[['type' if 'type' in df_detalhe else 'tipo', 'data_competencia', 'descricao', 'categoria_nome', 'valor', 'conta_nome']].copy()
            exibir_export['data_competencia'] = exibir_export['data_competencia'].dt.strftime('%d/%m/%Y')
            exibir_export.columns = ['Tipo', 'Data', 'Descrição', 'Categoria', 'Valor', 'Conta']
            
            excel_detalhe = to_excel_bytes({"Detalhes": exibir_export})
            st.download_button(
                label="💾 Baixar Tabela Selecionada (Excel)",
                data=excel_detalhe,
                file_name=f"Detalhes_{ano_sel}_{mes_drill}_{cat_drill}.xlsx".replace(" ", "_"),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
           
# ==========================================
# ==========================================
# ==========================================
# TESOURARIA (COM ANEXO MANDATÓRIO PARA PAGAR E OPCIONAL PARA RECEBER)
# ==========================================
elif page == "Tesouraria":
    st.title("Tesouraria")
    
    # Sincroniza a aba solicitada externamente (ex: botão Detalhes) com o componente de abas
    if "tesouraria_tab" in st.session_state:
        st.session_state["tesouraria_active_tab"] = st.session_state.pop("tesouraria_tab")

    tab1, tab2, tab3, tab4 = st.tabs([
        "📝 Novo Lançamento", 
        "⏳ Contas a Pagar/Receber", 
        "📜 Histórico Completo", 
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

            arquivo = st.file_uploader("Comprovante / Nota Fiscal (Opcional)", type=['png', 'jpg', 'jpeg', 'pdf'], key="lanc_arq_up")

            if st.form_submit_button("💾 Salvar Lançamento", use_container_width=True, type="primary"):
                if valor <= 0 or not descricao or not opcoes_cats:
                    st.warning("⚠️ Preencha descrição, valor e categoria corretamente.")
                else:
                    with st.spinner("Salvando..."):
                        url_anexo = comprimir_e_fazer_upload(arquivo, pasta="notas") if arquivo else None
                        
                        if recorrente:
                            sb_request("lancamentos", "POST", {
                                "descricao": descricao,
                                "tipo": tipo_lanc,
                                "valor": float(valor),
                                "data_competencia": str(date.today()),
                                "data_vencimento": str(date.today()),
                                "status": "Pendente",
                                "categoria_id": opcoes_cats[categoria_sel],
                                "conta_bancaria_id": contas_opcoes.get(conta_sel),
                                "centro_custo": None if tag == "Nenhum" else tag,
                                "recorrente": True,
                                "dia_vencimento_fixo": int(dia_vencimento),
                                "data_fim_recorrencia": str(data_fim_rec),
                                "url_anexo": url_anexo
                            })
                        else:
                            sb_request("lancamentos", "POST", {
                                "descricao": descricao,
                                "tipo": tipo_lanc,
                                "valor": float(valor),
                                "data_competencia": str(data_comp),
                                "data_vencimento": str(data_venc),
                                "data_pagamento": str(data_venc) if status_lanc == "Concluído" else None,
                                "status": status_lanc,
                                "categoria_id": opcoes_cats[categoria_sel],
                                "conta_bancaria_id": contas_opcoes.get(conta_sel),
                                "centro_custo": None if tag == "Nenhum" else tag,
                                "recorrente": False,
                                "url_anexo": url_anexo
                            })

                        st.cache_data.clear()
                        st.success("✅ Lançamento registrado com sucesso!")
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

            # Bloco Superior: Contas a Pagar
            st.subheader("💳 Contas a Pagar")
            if pend_pagar.empty:
                st.success("Nenhuma conta a pagar pendente. 👍")
            else:
                for _, row in pend_pagar.iterrows():
                    venc = row['data_vencimento']
                    situacao = "🔴 Atrasado" if pd.notna(venc) and venc < hoje else ("🟡 Vence hoje" if venc == hoje else "🟢 A vencer")
                    c1, c2, c3, c4, c5 = st.columns([3, 1.4, 1.4, 1.4, 1.2])
                    c1.write(f"**{row['descricao']}** — {row['categoria_nome']}")
                    c2.write(fmt_moeda(row['valor']))
                    c3.write(venc.strftime('%d.%m.%Y') if pd.notna(venc) else '—')
                    c4.write(situacao)
                    
                    row_id = str(row['id'])
                    is_paying = st.session_state.get(f"paying_{row_id}", False)
                    
                    if not is_paying:
                        if c5.button("✅ Pagar", key=f"pagar_tab_{row_id}"):
                            st.session_state[f"paying_{row_id}"] = True
                            st.rerun()
                    else:
                        c5.write("Aguardando...")

                    if st.session_state.get(f"paying_{row_id}", False):
                        with st.container(border=True):
                            st.markdown(f"📎 **Anexo Obrigatório (Nota Fiscal/Comprovante) para:** {row['descricao']}")
                            arq_pagar = st.file_uploader("Selecione o arquivo (PDF, PNG, JPG)", type=['png', 'jpg', 'jpeg', 'pdf'], key=f"file_pagar_{row_id}")
                            
                            col_b1, col_b2 = st.columns(2)
                            if col_b1.button("💾 Confirmar Pagamento", key=f"conf_pagar_{row_id}", type="primary"):
                                if not arq_pagar:
                                    st.error("⚠️ O anexo da nota fiscal/comprovante é **mandatório** para contas a pagar.")
                                else:
                                    with st.spinner("Enviando anexo e registrando pagamento..."):
                                        url_anexo = comprimir_e_fazer_upload(arq_pagar, pasta="notas")
                                        res = sb_request("lancamentos", "PATCH", {
                                            "status": "Concluído", 
                                            "data_pagamento": str(date.today()),
                                            "url_anexo": url_anexo
                                        }, filtros={"id": f"eq.{row_id}"})
                                        if res is not None:
                                            st.session_state[f"paying_{row_id}"] = False
                                            st.cache_data.clear()
                                            st.success("Pagamento registrado com sucesso!")
                                            time.sleep(1)
                                            st.rerun()
                            if col_b2.button("❌ Cancelar", key=f"canc_pagar_{row_id}"):
                                st.session_state[f"paying_{row_id}"] = False
                                st.rerun()

            st.markdown("---")

            # Bloco Inferior: Contas a Receber
            st.subheader("💰 Contas a Receber")
            if pend_receber.empty:
                st.success("Nenhuma conta a receber pendente. 👍")
            else:
                for _, row in pend_receber.iterrows():
                    venc = row['data_vencimento']
                    situacao = "🔴 Atrasado" if pd.notna(venc) and venc < hoje else ("🟡 Vence hoje" if venc == hoje else "🟢 A vencer")
                    c1, c2, c3, c4, c5 = st.columns([3, 1.4, 1.4, 1.4, 1.2])
                    c1.write(f"**{row['descricao']}** — {row['categoria_nome']}")
                    c2.write(fmt_moeda(row['valor']))
                    c3.write(venc.strftime('%d.%m.%Y') if pd.notna(venc) else '—')
                    c4.write(situacao)
                    
                    row_id = str(row['id'])
                    is_receiving = st.session_state.get(f"receiving_{row_id}", False)
                    
                    if not is_receiving:
                        if c5.button("✅ Receber", key=f"receber_tab_{row_id}"):
                            st.session_state[f"receiving_{row_id}"] = True
                            st.rerun()
                    else:
                        c5.write("Aguardando...")

                    if st.session_state.get(f"receiving_{row_id}", False):
                        with st.container(border=True):
                            st.markdown(f"📎 **Anexo Opcional para:** {row['descricao']}")
                            arq_receber = st.file_uploader("Selecione o arquivo se desejar anexar (Opcional)", type=['png', 'jpg', 'jpeg', 'pdf'], key=f"file_receber_{row_id}")
                            
                            col_b1, col_b2 = st.columns(2)
                            if col_b1.button("💾 Confirmar Recebimento", key=f"conf_receber_{row_id}", type="primary"):
                                with st.spinner("Registrando recebimento..."):
                                    url_anexo = comprimir_e_fazer_upload(arq_receber, pasta="notas") if arq_receber else None
                                    payload = {
                                        "status": "Concluído", 
                                        "data_pagamento": str(date.today())
                                    }
                                    if url_anexo:
                                        payload["url_anexo"] = url_anexo
                                        
                                    res = sb_request("lancamentos", "PATCH", payload, filtros={"id": f"eq.{row_id}"})
                                    if res is not None:
                                        st.session_state[f"receiving_{row_id}"] = False
                                        st.cache_data.clear()
                                        st.success("Recebimento registrado com sucesso!")
                                        time.sleep(1)
                                        st.rerun()
                            if col_b2.button("❌ Cancelar", key=f"canc_receber_{row_id}"):
                                st.session_state[f"receiving_{row_id}"] = False
                                st.rerun()

    with tab3:
        df = carregar_lancamentos_df()
        if df.empty:
            st.info("Nenhum lançamento registrado.")
        else:
            st.markdown("### Filtros do Histórico")
            col_f1, col_f2, col_f3, col_f4 = st.columns(4)
            start_date = date.today().replace(day=1)
            end_date = (pd.Timestamp.today() + pd.offsets.MonthEnd(1)).date()
            
            data_inicio = col_f1.date_input("De", start_date, format="DD.MM.YYYY", key="hist_dt_ini")
            data_fim = col_f2.date_input("Até", end_date, format="DD.MM.YYYY", key="hist_dt_fim")
            filtro_tipo = col_f3.multiselect("Tipo", ["Entrada", "Saída"], default=["Entrada", "Saída"], key="hist_ft_tipo")
            filtro_status = col_f4.multiselect("Situação", df['status'].unique().tolist(), default=df['status'].unique().tolist(), key="hist_ft_status")
            
            mask_data = (df['data_competencia'].dt.date >= data_inicio) & (df['data_competencia'].dt.date <= data_fim)
            dff = df[mask_data & df['tipo'].isin(filtro_tipo) & df['status'].isin(filtro_status)].copy()
            
            if dff.empty:
                st.info("Nenhum lançamento encontrado para os filtros selecionados.")
            else:
                dff['ordem_tipo'] = dff['tipo'].map({'Entrada': 0, 'Saída': 1})
                dff = dff.sort_values(by=['ordem_tipo', 'data_competencia'], ascending=[True, False])
                
                exibir = dff[['id', 'data_competencia', 'tipo', 'descricao', 'categoria_nome', 'valor', 'status', 'centro_custo', 'conta_nome']].copy()
                exibir['Data'] = exibir['data_competencia'].dt.strftime('%d.%m.%Y')
                exibir['Valor (R$)'] = exibir['valor'].apply(fmt_moeda)
                exibir.insert(0, '✏️ Editar', False)
                exibir = exibir[['✏️ Editar', 'Data', 'tipo', 'descricao', 'categoria_nome', 'Valor (R$)', 'status', 'centro_custo', 'conta_nome', 'id']]
                
                editado = st.data_editor(exibir, hide_index=True, use_container_width=True, key="editor_hist_novo", column_config={"id": None})
                selecionados = editado[editado['✏️ Editar'] == True]
                
                if len(selecionados) == 1:
                    id_lanc = str(selecionados.iloc[0]['id'])
                    lanc_raw = next((l for l in carregar("lancamentos") if str(l.get('id')) == id_lanc), None)
                    if lanc_raw:
                        st.markdown("---")
                        st.markdown(f"#### Editando: {lanc_raw.get('descricao', '')}")
                        with st.form(f"form_ed_{id_lanc}"):
                            n_desc = st.text_input("Descrição", value=lanc_raw.get('descricao', ''))
                            n_valor = st.number_input("Valor (R$)", value=float(lanc_raw.get('valor') or 0), format="%.2f")
                            
                            c_b1, c_b2 = st.columns(2)
                            btn_upd = c_b1.form_submit_button("💾 Salvar Alterações", use_container_width=True)
                            btn_del = c_b2.form_submit_button("🗑️ Excluir", use_container_width=True)
                            
                            if btn_upd:
                                sb_request("lancamentos", "PATCH", {"descricao": n_desc, "valor": float(n_valor)}, filtros={"id": f"eq.{id_lanc}"})
                                st.cache_data.clear(); st.success("Atualizado!"); time.sleep(1); st.rerun()
                            if btn_del:
                                sb_request("lancamentos", "DELETE", filtros={"id": f"eq.{id_lanc}"})
                                st.cache_data.clear(); st.success("Excluído!"); time.sleep(1); st.rerun()

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
# CONCILIAÇÃO BANCÁRIA (COM IMPORTAÇÃO INTELIGENTE DE OFX)
# ==========================================
elif page == "Conciliação Bancária":
    st.title("Conciliação Bancária")
    st.markdown("Confira os lançamentos manualmente ou importe o extrato do banco (OFX) para automatizar o cadastro de Dízimos e Ofertas em lote.")

    col_nova, col_edit = st.columns(2)
    with col_nova:
        with st.expander("➕ Cadastrar Conta Bancária"):
            with st.form("form_conta", clear_on_submit=True):
                nome_conta = st.text_input("Nome da Conta (Ex: Itaú Principal)")
                tipo_conta = st.selectbox("Tipo", ["Corrente", "Poupança", "Investimento"], key="tipo_conta_novo")
                saldo_inicial = st.number_input("Saldo Inicial (R$)", min_value=0.0, format="%.2f")
                if st.form_submit_button("Cadastrar Conta", use_container_width=True):
                    if nome_conta:
                        res = sb_request("contas_bancarias", "POST", {"nome": nome_conta, "tipo": tipo_conta, "saldo_inicial": float(saldo_inicial)})
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
                    n_saldo = st.number_input("Saldo Inicial (R$)", value=float(conta_data.get('saldo_inicial') or 0.0), format="%.2f")
                    
                    c_btn1, c_btn2 = st.columns(2)
                    btn_upd = c_btn1.form_submit_button("💾 Atualizar", use_container_width=True)
                    btn_del = c_btn2.form_submit_button("🗑️ Excluir", use_container_width=True)
                    
                    if btn_upd:
                        res = sb_request("contas_bancarias", "PATCH", {"nome": n_nome, "tipo": n_tipo, "saldo_inicial": float(n_saldo)}, filtros={"id": f"eq.{conta_data['id']}"})
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
        # Chave única aplicada para evitar conflito com o selectbox do expander acima
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
            st.markdown("Faça o upload do arquivo **.OFX** gerado pelo seu banco. O sistema cruzará os dados com os eventos/lançamentos já existentes, filtrando apenas os PIX novos para cadastro rápido.")
            arquivo_ofx = st.file_uploader("Selecione o arquivo OFX do banco", type=['ofx', 'txt'])

            if arquivo_ofx:
                import re
                content = arquivo_ofx.read().decode('latin1', errors='ignore')
                transacoes = []
                
                for bloco in re.split(r'<STMTTRN>', content)[1:]:
                    dt_match = re.search(r'<DTPOSTED>(\d{8})', bloco)
                    valor_match = re.search(r'<TRNAMT>([-\d\.]+)', bloco)
                    desc_match = re.search(r'<MEMO>(.*?)(?:<|$)', bloco)

                    if dt_match and valor_match:
                        dt = pd.to_datetime(dt_match.group(1), format='%Y%m%d').date()
                        valor = float(valor_match.group(1))
                        desc = desc_match.group(1).strip() if desc_match else "Extrato Bancário"
                        desc = re.sub(r'<[^>]+>', '', desc)
                        transacoes.append({
                            "Data": dt,
                            "Valor": abs(valor),
                            "Descrição Bancária": desc[:80],
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

                    df_entradas_novas = df_ofx[(df_ofx['Status'] == 'Não registrado') & (df_ofx['Tipo'] == 'Entrada')].copy()

                    if df_entradas_novas.empty:
                        st.success("🎉 Todas as entradas deste extrato já constam e batem com o sistema!")
                    else:
                        st.info(f"Encontramos **{len(df_entradas_novas)} transferências (PIX/Depósitos)** no extrato que ainda não constam no sistema.")
                        df_entradas_novas.insert(0, 'Cadastrar', False)

                        st.write("Marque na tabela abaixo quais transferências você deseja converter em lançamentos:")
                        df_editado = st.data_editor(
                            df_entradas_novas[['Cadastrar', 'Data', 'Descrição Bancária', 'Valor']],
                            hide_index=True, use_container_width=True
                        )

                        col_cat, col_btn = st.columns([2, 1])
                        opcoes_cats_entrada = {c["nome"]: c["id"] for c in categorias_db if c["tipo"] == "Entrada"}
                        
                        cat_padrao = "Dízimos e Ofertas" if "Dízimos e Ofertas" in opcoes_cats_entrada else list(opcoes_cats_entrada.keys())[0]
                        cat_lote = col_cat.selectbox("Classificar os itens marcados como:", list(opcoes_cats_entrada.keys()), index=list(opcoes_cats_entrada.keys()).index(cat_padrao), key="select_cat_lote_ofx")

                        if col_btn.button("💾 Salvar Marcados no Banco", type="primary", use_container_width=True):
                            itens_selecionados = df_editado[df_editado['Cadastrar'] == True]
                            if itens_selecionados.empty:
                                st.warning("Marque pelo menos um item na tabela clicando na caixa 'Cadastrar'.")
                            else:
                                with st.spinner("Registrando lançamentos..."):
                                    for _, row_sel in itens_selecionados.iterrows():
                                        sb_request("lancamentos", "POST", {
                                            "descricao": f"PIX Extrato: {row_sel['Descrição Bancária']}",
                                            "tipo": "Entrada",
                                            "valor": float(row_sel['Valor']),
                                            "data_competencia": str(row_sel['Data']),
                                            "status": "Concluído",
                                            "data_pagamento": str(row_sel['Data']),
                                            "categoria_id": opcoes_cats_entrada[cat_lote],
                                            "conta_bancaria_id": conta_id,
                                            "conciliado": True
                                        })
                                    st.cache_data.clear()
                                    st.success(f"{len(itens_selecionados)} Dízimos/Ofertas registrados com sucesso!")
                                    time.sleep(1.5)
                                    st.rerun()

                    with st.expander("Ver itens cruzados automaticamente (PIX de eventos, etc)"):
                        df_ja_registrado = df_ofx[df_ofx['Status'] == 'Já no sistema']
                        if not df_ja_registrado.empty:
                            st.dataframe(df_ja_registrado[['Data', 'Descrição Bancária', 'Valor', 'Tipo']], use_container_width=True, hide_index=True)
                        else:
                            st.write("Nenhum cruzamento encontrado neste arquivo.")

        with tab_manual:
            st.subheader("Lista de Lançamentos da Conta")
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
                        res = sb_request("lancamentos", "PATCH", {"conciliado": novo_valor}, filtros={"id": f"eq.{row['id']}"})
                        if res is not None:
                            st.cache_data.clear()
# ==========================================
# ==========================================
# PAINEL DE EVENTOS
# ==========================================
elif page == "Painel de Eventos":
    st.title("Painel de Eventos")
    st.markdown("Cadastre acampamentos, retiros e conferências para abrir inscrições.")

    col_novo, col_edit = st.columns(2)
    with col_novo:
        with st.expander("➕ Criar Novo Evento", expanded=len(eventos_db) == 0):
            with st.form("form_evento", clear_on_submit=True):
                nome_ev = st.text_input("Nome do Evento")
                data_ev = st.date_input("Data do Evento", date.today())
                valor_ev = st.number_input("Valor da Inscrição (R$)", min_value=0.0, format="%.2f")
                vagas_ev = st.number_input("Total de Vagas", min_value=0, step=1)
                chave_pix_ev = st.text_input("Chave Pix")
                descricao_ev = st.text_area("Descrição/Orientações")
                permite_parc = st.checkbox("Permitir parcelamento")
                num_parc = st.number_input("Máx. Parcelas", min_value=1, max_value=12, value=1, disabled=not permite_parc)

                if st.form_submit_button("Criar Evento", use_container_width=True):
                    if not nome_ev.strip() or valor_ev <= 0 or vagas_ev <= 0 or not chave_pix_ev.strip():
                        st.warning("Preencha Nome, Valor > 0, Vagas > 0 e Chave Pix.")
                    else:
                        sb_request("eventos", "POST", {
                            "nome": nome_ev.strip(), "descricao": descricao_ev.strip() or None,
                            "data_evento": str(data_ev), "valor_inscricao": float(valor_ev),
                            "vagas_total": int(vagas_ev), "chave_pix": chave_pix_ev.strip(),
                            "centro_custo": nome_ev.strip(), "status": "Aberto",
                            "permite_parcelamento": bool(permite_parc), "numero_parcelas": int(num_parc) if permite_parc else 1
                        })
                        st.cache_data.clear(); st.success("Evento criado!"); time.sleep(1); st.rerun()

    with col_edit:
        with st.expander("✏️ Editar ou Excluir Evento"):
            if eventos_db:
                ev_map = {e['nome']: e for e in eventos_db}
                ev_sel = st.selectbox("Selecione o Evento", list(ev_map.keys()))
                ev_data = ev_map[ev_sel]
                
                with st.form("form_edit_evento"):
                    n_nome_ev = st.text_input("Nome", value=ev_data['nome'])
                    n_data_ev = st.date_input("Data", pd.to_datetime(ev_data.get('data_evento', date.today())).date())
                    n_valor_ev = st.number_input("Valor (R$)", value=float(ev_data.get('valor_inscricao') or 0), format="%.2f")
                    n_vagas_ev = st.number_input("Vagas", value=int(ev_data.get('vagas_total') or 0))
                    n_pix = st.text_input("Pix", value=ev_data.get('chave_pix') or "")
                    n_desc = st.text_area("Descrição", value=ev_data.get('descricao') or "")
                    
                    c_btn1, c_btn2 = st.columns(2)
                    b_upd_ev = c_btn1.form_submit_button("💾 Atualizar", use_container_width=True)
                    b_del_ev = c_btn2.form_submit_button("🗑️ Excluir", use_container_width=True)
                    
                    if b_upd_ev:
                        sb_request("eventos", "PATCH", {
                            "nome": n_nome_ev, "data_evento": str(n_data_ev), "valor_inscricao": float(n_valor_ev),
                            "vagas_total": int(n_vagas_ev), "chave_pix": n_pix, "descricao": n_desc
                        }, filtros={"id": f"eq.{ev_data['id']}"})
                        st.cache_data.clear(); st.success("Atualizado!"); time.sleep(1); st.rerun()
                    if b_del_ev:
                        sb_request("eventos", "DELETE", filtros={"id": f"eq.{ev_data['id']}"})
                        st.cache_data.clear(); st.success("Excluído!"); time.sleep(1); st.rerun()
            else:
                st.info("Nenhum evento cadastrado.")

    eventos_atualizados = carregar("eventos")
    st.markdown("---")
    st.subheader("Eventos Cadastrados")

    if not eventos_atualizados:
        st.info("Nenhum evento cadastrado ainda.")
    else:
        inscricoes_all = carregar("inscricoes")
        pagamentos_all = carregar("inscricao_pagamentos")

        for ev in eventos_atualizados:
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
            vagas_restantes = max(vagas_total - total_inscritos, 0)
            status_evento = ev.get("status") or "Aberto"
            cor_status = "#059669" if status_evento == "Aberto" else "#94A3B8"
            parcelamento_texto = f"Pagamento em até {int(ev.get('numero_parcelas') or 1)}x" if ev.get("permite_parcelamento") else "Pagamento à vista"

            st.markdown(f"""
            <div style="background:#FFFFFF;border:1px solid #E2E8F0;border-radius:14px;padding:20px;margin-bottom:12px;">
                <h4 style="margin-top:0;">{ev.get('nome', 'Evento')} <span style="font-size:0.8rem;color:{cor_status};">● {status_evento}</span></h4>
                <p style="color:#475569;margin:4px 0;">📅 {ev.get('data_evento') or '—'} &nbsp;•&nbsp; 💰 {fmt_moeda(ev.get('valor_inscricao'))} por pessoa</p>
                <p style="color:#475569;margin:4px 0;">💳 {parcelamento_texto} &nbsp;•&nbsp; 🔑 Pix: {ev.get('chave_pix') or '—'}</p>
                <p style="color:#475569;margin:4px 0;">👥 {total_inscritos} inscritos &nbsp;•&nbsp; {vagas_restantes} vagas restantes &nbsp;•&nbsp; ✅ {len(inscricoes_quitadas)} quitados &nbsp;•&nbsp; 🟡 {len(inscricoes_parciais)} parciais</p>
                <p style="color:#D97706;margin:4px 0;font-weight:600;">⏳ {len(pagamentos_pendentes)} comprovantes aguardando aprovação</p>
                <p style="color:#059669;margin:4px 0;font-weight:600;">💵 Arrecadado: {fmt_moeda(total_arrecadado)} &nbsp;•&nbsp; A receber: {fmt_moeda(saldo_a_receber)}</p>
            </div>
            """, unsafe_allow_html=True)

            app_url = st.secrets.get("APP_URL", "").rstrip("/")
            complemento_link = f"?pagina=inscricao&evento={ev['id']}"
            link_publico = f"{app_url}/{complemento_link}" if app_url else complemento_link

            col_link, col_acao = st.columns([3, 1])
            with col_link:
                st.text_input("Link público para inscrição", value=link_publico, key=f"link_evento_{ev['id']}", disabled=True)
            with col_acao:
                novo_status = "Encerrado" if status_evento == "Aberto" else "Aberto"
                texto_botao = "🔒 Encerrar inscrições" if status_evento == "Aberto" else "🔓 Reabrir inscrições"
                if st.button(texto_botao, key=f"alterar_status_evento_{ev['id']}", use_container_width=True):
                    resultado = sb_request("eventos", "PATCH", {"status": novo_status}, filtros={"id": f"eq.{ev['id']}"})
                    if resultado is not None:
                        st.cache_data.clear()
                        st.rerun()
            st.markdown("---")

# ==========================================
# INSCRIÇÕES E COMPROVANTES
# ==========================================
elif page == "Inscrições e Comprovantes":
    st.title("Inscrições e Comprovantes")
    st.markdown("Valide os comprovantes de Pix enviados pelos participantes, parcela por parcela.")

    if not eventos_db:
        st.info("Cadastre um evento primeiro em 'Painel de Eventos'.")
    else:
        evento_opcoes = {e["nome"]: e["id"] for e in eventos_db}
        evento_sel = st.selectbox("Evento", list(evento_opcoes.keys()))
        evento_id_sel = evento_opcoes[evento_sel]

        inscricoes_evento = [i for i in carregar("inscricoes") if i.get('evento_id') == evento_id_sel]
        pagamentos_all = carregar("inscricao_pagamentos")
        map_insc = {i["id"]: i for i in inscricoes_evento}

        filtro = st.radio("Mostrar", ["Comprovantes pendentes", "Todas as inscrições"], horizontal=True)

        if filtro == "Comprovantes pendentes":
            pendentes_pg = [p for p in pagamentos_all if p.get("status") == "Pendente" and p.get("inscricao_id") in map_insc]
            if not pendentes_pg:
                st.success("Nenhum comprovante pendente para este evento.")
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
                        novo_lanc = sb_request("lancamentos", "POST", {
                            "descricao": f"Inscrição ({insc['nome_participante']} - parcela {p.get('numero_parcela',1)}) - {evento_sel}",
                            "tipo": "Entrada", "valor": float(p.get('valor') or 0),
                            "data_competencia": str(date.today()), "status": "Concluído",
                            "categoria_id": cat_evento_id, "centro_custo": evento_sel
                        })
                        
                        if novo_lanc is not None:
                            lanc_id = novo_lanc[0]['id'] if isinstance(novo_lanc, list) and len(novo_lanc)>0 else None
                            sb_request("inscricao_pagamentos", "PATCH", {"status": "Aprovado", "lancamento_id": lanc_id}, filtros={"id": f"eq.{p['id']}"})
                            novo_valor_pago = float(insc.get('valor_pago') or 0) + float(p.get('valor') or 0)
                            novo_status = "Completo" if novo_valor_pago >= float(insc.get('valor_total') or 0) - 0.01 else "Parcial"
                            sb_request("inscricoes", "PATCH", {"valor_pago": novo_valor_pago, "status_pagamento": novo_status}, filtros={"id": f"eq.{insc['id']}"})
                            st.cache_data.clear()
                            st.rerun()

                    if col_b.button("❌", key=f"rejeitar_pg_{p['id']}", help="Rejeitar"):
                        res = sb_request("inscricao_pagamentos", "PATCH", {"status": "Rejeitado"}, filtros={"id": f"eq.{p['id']}"})
                        if res is not None:
                            st.cache_data.clear()
                            st.rerun()
                    st.markdown("<hr style='margin:6px 0;border-color:#E2E8F0;'>", unsafe_allow_html=True)
        else:
            if not inscricoes_evento:
                st.info("Nenhuma inscrição para este evento.")
            else:
                for insc in inscricoes_evento:
                    pgs = [p for p in pagamentos_all if p.get("inscricao_id") == insc["id"]]
                    emoji_status = {"Pendente": "⏳", "Parcial": "🟡", "Completo": "✅"}.get(insc.get("status_pagamento"), "⏳")
                    st.markdown(f"**{insc['nome_participante']}** (CPF {insc.get('cpf','—')}) — {emoji_status} {insc.get('status_pagamento','Pendente')} — {fmt_moeda(insc.get('valor_pago'))} / {fmt_moeda(insc.get('valor_total'))}")
                    for p in sorted(pgs, key=lambda x: x.get("numero_parcela", 1)):
                        st.caption(f" Parcela {p.get('numero_parcela',1)}: {fmt_moeda(p.get('valor'))} — {p['status']}")
                    st.markdown("<hr style='margin:6px 0;border-color:#E2E8F0;'>", unsafe_allow_html=True)

# ==========================================
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
                    if st.button("Excluir Meta Selecionada", use_container_width=True):
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

            # Garante que as colunas existam como listas/Series, mesmo se não houver dados no tipo
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
            # Agora chamamos a coluna diretamente em vez de usar .get()
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
# ==========================================
# ==========================================
# EXPORTAR CONTABILIDADE (COM DOWNLOAD DE ANEXOS EM ZIP)
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

        exportar = dff[['data_competencia', 'tipo', 'descricao', 'categoria_nome', 'valor', 'status', 'centro_custo']].copy()
        exportar['data_competencia'] = exportar['data_competencia'].dt.strftime('%d/%m/%Y')
        exportar.columns = ['Data', 'Tipo', 'Descrição', 'Categoria', 'Valor', 'Status', 'Projeto']

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
            import zipfile
            
            # Botão para baixar Pacote ZIP (Excel + Anexos do Mês)
            if st.button("📦 Baixar Pacote ZIP (Excel + Anexos)", use_container_width=True):
                with st.spinner("Empacotando lançamentos e baixando anexos..."):
                    zip_buffer = io.BytesIO()
                    
                    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                        # 1. Adiciona a planilha Excel dentro do ZIP
                        zip_file.writestr(f"financeiro_{ano_exp}_{mes_exp}.xlsx", excel_data)
                        
                        # 2. Varre os lançamentos filtrados para coletar e baixar os anexos do Storage
                        anexos_adicionados = 0
                        for _, row_orig in dff.iterrows():
                            path_anexo = row_orig.get('url_anexo')
                            if path_anexo and isinstance(path_anexo, str) and path_anexo.strip():
                                try:
                                    # Baixa o arquivo binário diretamente do bucket 'comprovantes' do Supabase
                                    file_res = supabase.storage.from_("comprovantes").download(path_anexo)
                                    if file_res:
                                        # Cria um nome de arquivo limpo para dentro do ZIP
                                        nome_original = path_anexo.split('/')[-1]
                                        data_str = pd.to_datetime(row_orig['data_competencia']).strftime('%Y%m%d')
                                        nome_no_zip = f"comprovantes/{data_str}_{nome_original}"
                                        
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
