import streamlit as st
import pandas as pd
import os
import shutil
from datetime import datetime, date
from pathlib import Path

# ============================================================================
# CONFIGURAÇÕES E INICIALIZAÇÃO
# ============================================================================
st.set_page_config(
    page_title="Sistema de Controle de Diárias",
    page_icon="💰",
    layout="wide",
    initial_sidebar_state="expanded",
)

ARQUIVO_DADOS = "controle_diarias.xlsx"
PASTA_COMPROVANTES = "Comprovantes_Diarias"

# Criar pasta de comprovantes se não existir
os.makedirs(PASTA_COMPROVANTES, exist_ok=True)

# ============================================================================
# CSS PERSONALIZADO
# ============================================================================
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
    }

    .main-title {
        font-size: 2.2rem;
        font-weight: 700;
        color: #1a1a2e;
        text-align: center;
        margin-bottom: 0.5rem;
        letter-spacing: -0.5px;
    }

    .subtitle {
        font-size: 1rem;
        color: #6c757d;
        text-align: center;
        margin-bottom: 2rem;
    }

    .card {
        background: linear-gradient(135deg, #ffffff 0%, #f8f9fa 100%);
        border-radius: 16px;
        padding: 1.2rem;
        box-shadow: 0 4px 20px rgba(0,0,0,0.06);
        border: 1px solid rgba(0,0,0,0.04);
        text-align: center;
        transition: transform 0.2s ease;
    }

    .card:hover {
        transform: translateY(-3px);
    }

    .card-icon {
        font-size: 2rem;
        margin-bottom: 0.5rem;
    }

    .card-value {
        font-size: 1.6rem;
        font-weight: 700;
        color: #1a1a2e;
        margin-bottom: 0.2rem;
    }

    .card-label {
        font-size: 0.85rem;
        color: #6c757d;
        font-weight: 500;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }

    .card-pendente .card-value { color: #dc3545; }
    .card-pago .card-value { color: #198754; }
    .card-total .card-value { color: #0d6efd; }

    .section-title {
        font-size: 1.15rem;
        font-weight: 600;
        color: #1a1a2e;
        margin-top: 1.5rem;
        margin-bottom: 1rem;
        padding-bottom: 0.5rem;
        border-bottom: 2px solid #e9ecef;
    }

    .stButton > button {
        border-radius: 10px !important;
        font-weight: 600 !important;
        padding: 0.5rem 1.2rem !important;
        transition: all 0.2s ease !important;
    }

    .stButton > button:hover {
        transform: translateY(-1px);
        box-shadow: 0 4px 12px rgba(0,0,0,0.12);
    }

    .btn-cadastrar > button {
        background-color: #0d6efd !important;
        color: white !important;
        border: none !important;
    }

    .btn-editar > button {
        background-color: #ffc107 !important;
        color: #1a1a2e !important;
        border: none !important;
    }

    .btn-excluir > button {
        background-color: #dc3545 !important;
        color: white !important;
        border: none !important;
    }

    .btn-limpar > button {
        background-color: #6c757d !important;
        color: white !important;
        border: none !important;
    }

    .btn-exportar > button {
        background-color: #198754 !important;
        color: white !important;
        border: none !important;
    }

    div[data-testid="stDataFrame"] {
        border-radius: 12px;
        overflow: hidden;
        box-shadow: 0 2px 12px rgba(0,0,0,0.05);
    }

    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }

    .stTabs [data-baseweb="tab"] {
        border-radius: 8px 8px 0 0;
        padding: 0.5rem 1.2rem;
        font-weight: 500;
    }

    .badge {
        display: inline-block;
        padding: 0.25em 0.65em;
        font-size: 0.75em;
        font-weight: 600;
        line-height: 1;
        text-align: center;
        white-space: nowrap;
        vertical-align: baseline;
        border-radius: 50rem;
    }

    .badge-pendente {
        background-color: #fff3cd;
        color: #856404;
    }

    .badge-pago {
        background-color: #d1e7dd;
        color: #0f5132;
    }

    .footer {
        text-align: center;
        color: #adb5bd;
        font-size: 0.8rem;
        margin-top: 3rem;
        padding-top: 1rem;
        border-top: 1px solid #e9ecef;
    }
</style>
""", unsafe_allow_html=True)

# ============================================================================
# FUNÇÕES DE DADOS
# ============================================================================
@st.cache_data(ttl=5)
def carregar_dados():
    if os.path.exists(ARQUIVO_DADOS):
        df = pd.read_excel(ARQUIVO_DADOS, dtype=str, keep_default_na=False)
        # Garantir colunas
        for col in ['ID','LOJA','NOME','CPF','DATA_EXECUCAO','QUANTIDADE','VALOR_UNITARIO',
                    'TOTAL','DADOS_BANCARIOS','SUBSTITUICAO','MOTIVO','DATA_PAGAMENTO',
                    'SITUACAO','MES','SEMANA','ANO','COMPROVANTE','OBSERVACOES']:
            if col not in df.columns:
                df[col] = ''
        return df
    else:
        return pd.DataFrame(columns=['ID','LOJA','NOME','CPF','DATA_EXECUCAO','QUANTIDADE',
                                      'VALOR_UNITARIO','TOTAL','DADOS_BANCARIOS','SUBSTITUICAO',
                                      'MOTIVO','DATA_PAGAMENTO','SITUACAO','MES','SEMANA','ANO',
                                      'COMPROVANTE','OBSERVACOES'])

def salvar_dados(df):
    with pd.ExcelWriter(ARQUIVO_DADOS, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Diarias', index=False)

def gerar_id(df):
    if df.empty:
        return 'DIA-0001'
    nums = []
    for val in df['ID']:
        try:
            nums.append(int(str(val).replace('DIA-', '')))
        except:
            pass
    if nums:
        return f'DIA-{max(nums)+1:04d}'
    return 'DIA-0001'

def formatar_moeda(valor):
    try:
        v = float(str(val).replace(',', '.').replace('R$', '').strip())
        return f"R$ {v:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    except:
        return f"R$ {valor}"

def calcular_total(quant, valor):
    try:
        q = float(str(quant).replace(',', '.'))
        v = float(str(valor).replace(',', '.'))
        return int(q * v)
    except:
        return 0

def salvar_comprovante(uploaded_file, id_registro):
    if uploaded_file is not None:
        ext = Path(uploaded_file.name).suffix
        nome_arquivo = f"comprovante_{id_registro}{ext}"
        caminho = os.path.join(PASTA_COMPROVANTES, nome_arquivo)
        with open(caminho, "wb") as f:
            f.write(uploaded_file.getbuffer())
        return caminho
    return ""

# ============================================================================
# INICIALIZAÇÃO DE ESTADO
# ============================================================================
if 'df' not in st.session_state:
    st.session_state.df = carregar_dados()
if 'modo_edicao' not in st.session_state:
    st.session_state.modo_edicao = False
if 'id_edicao' not in st.session_state:
    st.session_state.id_edicao = None
if 'form_limpo' not in st.session_state:
    st.session_state.form_limpo = False
if 'mensagem' not in st.session_state:
    st.session_state.mensagem = ("", "")

def mostrar_mensagem(texto, tipo="info"):
    st.session_state.mensagem = (texto, tipo)

def limpar_mensagem():
    st.session_state.mensagem = ("", "")

# ============================================================================
# CABEÇALHO
# ============================================================================
st.markdown('<div class="main-title">💰 Sistema de Controle de Diárias</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">Gerencie diárias de colaboradores de forma prática e organizada</div>', unsafe_allow_html=True)

# ============================================================================
# SIDEBAR - FILTROS
# ============================================================================
with st.sidebar:
    st.markdown("### 🔍 Filtros Avançados")
    st.markdown("---")

    df_atual = st.session_state.df.copy()

    # Extrair valores únicos para filtros
    lojas_unicas = sorted([x for x in df_atual['LOJA'].unique() if x])
    meses_unicos = sorted([x for x in df_atual['MES'].unique() if x])
    semanas_unicas = sorted([x for x in df_atual['SEMANA'].unique() if x])
    anos_unicos = sorted([x for x in df_atual['ANO'].unique() if x])

    filtro_loja = st.selectbox("🏪 Loja", options=["Todas"] + lojas_unicas, index=0)
    filtro_mes = st.selectbox("📅 Mês", options=["Todos"] + meses_unicos, index=0)
    filtro_semana = st.selectbox("📆 Semana", options=["Todas"] + semanas_unicas, index=0)
    filtro_ano = st.selectbox("📅 Ano", options=["Todos"] + anos_unicos, index=0)
    filtro_situacao = st.selectbox("💳 Situação", options=["Todas", "PENDENTE", "PAGO"], index=0)
    filtro_pesquisa = st.text_input("🔎 Pesquisar (nome, CPF, motivo)", placeholder="Digite para buscar...")

    st.markdown("---")

    if st.button("🔄 Limpar Filtros", use_container_width=True):
        st.rerun()

    st.markdown("---")
    st.markdown("### 📊 Resumo dos Filtros")
    df_filtrado_sidebar = df_atual.copy()
    if filtro_loja != "Todas":
        df_filtrado_sidebar = df_filtrado_sidebar[df_filtrado_sidebar['LOJA'] == filtro_loja]
    if filtro_mes != "Todos":
        df_filtrado_sidebar = df_filtrado_sidebar[df_filtrado_sidebar['MES'] == filtro_mes]
    if filtro_semana != "Todas":
        df_filtrado_sidebar = df_filtrado_sidebar[df_filtrado_sidebar['SEMANA'] == filtro_semana]
    if filtro_ano != "Todos":
        df_filtrado_sidebar = df_filtrado_sidebar[df_filtrado_sidebar['ANO'] == filtro_ano]
    if filtro_situacao != "Todas":
        df_filtrado_sidebar = df_filtrado_sidebar[df_filtrado_sidebar['SITUACAO'] == filtro_situacao]
    if filtro_pesquisa:
        pesq = filtro_pesquisa.lower()
        mask = (
            df_filtrado_sidebar['NOME'].str.lower().str.contains(pesq, na=False) |
            df_filtrado_sidebar['CPF'].str.lower().str.contains(pesq, na=False) |
            df_filtrado_sidebar['MOTIVO'].str.lower().str.contains(pesq, na=False) |
            df_filtrado_sidebar['LOJA'].str.lower().str.contains(pesq, na=False)
        )
        df_filtrado_sidebar = df_filtrado_sidebar[mask]

    st.metric("Registros filtrados", len(df_filtrado_sidebar))

# ============================================================================
# CARDS DE RESUMO
# ============================================================================
df_resumo = df_filtrado_sidebar.copy()

# Calcular valores
total_diarias = len(df_resumo)
valor_total = 0
valor_pendente = 0
valor_pago = 0
for _, row in df_resumo.iterrows():
    try:
        v = float(str(row['TOTAL']).replace(',', '.').replace('R$', '').strip())
        valor_total += v
        if row['SITUACAO'] == 'PENDENTE':
            valor_pendente += v
        elif row['SITUACAO'] == 'PAGO':
            valor_pago += v
    except:
        pass

col1, col2, col3, col4 = st.columns(4)
with col1:
    st.markdown(f'''
    <div class="card card-total">
        <div class="card-icon">📋</div>
        <div class="card-value">{total_diarias}</div>
        <div class="card-label">Total de Diárias</div>
    </div>
    ''', unsafe_allow_html=True)
with col2:
    st.markdown(f'''
    <div class="card card-total">
        <div class="card-icon">💵</div>
        <div class="card-value">R$ {valor_total:,.2f}</div>
        <div class="card-label">Valor Total</div>
    </div>
    ''', unsafe_allow_html=True)
with col3:
    st.markdown(f'''
    <div class="card card-pendente">
        <div class="card-icon">⏳</div>
        <div class="card-value">R$ {valor_pendente:,.2f}</div>
        <div class="card-label">Total Pendente</div>
    </div>
    ''', unsafe_allow_html=True)
with col4:
    st.markdown(f'''
    <div class="card card-pago">
        <div class="card-icon">✅</div>
        <div class="card-value">R$ {valor_pago:,.2f}</div>
        <div class="card-label">Total Pago</div>
    </div>
    ''', unsafe_allow_html=True)

# ============================================================================
# MENSAGENS
# ============================================================================
msg_texto, msg_tipo = st.session_state.mensagem
if msg_texto:
    if msg_tipo == "sucesso":
        st.success(msg_texto)
    elif msg_tipo == "erro":
        st.error(msg_texto)
    elif msg_tipo == "aviso":
        st.warning(msg_texto)
    else:
        st.info(msg_texto)
    limpar_mensagem()

# ============================================================================
# TABELA DE DADOS
# ============================================================================
st.markdown('<div class="section-title">📋 Registros de Diárias</div>', unsafe_allow_html=True)

# Preparar dados para exibição
df_exibir = df_filtrado_sidebar.copy()
if not df_exibir.empty:
    # Ordenar por ID
    try:
        df_exibir['_sort'] = df_exibir['ID'].str.replace('DIA-', '').astype(int)
        df_exibir = df_exibir.sort_values('_sort', ascending=False).drop(columns=['_sort'])
    except:
        pass

    # Selecionar colunas para exibição
    cols_exibir = ['ID','LOJA','NOME','CPF','DATA_EXECUCAO','QUANTIDADE','VALOR_UNITARIO','TOTAL',
                   'DATA_PAGAMENTO','SITUACAO','MES','SEMANA','ANO','SUBSTITUICAO','MOTIVO']
    df_display = df_exibir[[c for c in cols_exibir if c in df_exibir.columns]].copy()

    # Renomear colunas
    df_display.columns = ['ID','Loja','Nome','CPF','Datas Execução','Qtd','Valor Uni.','Total',
                          'Data Pagto.','Situação','Mês','Semana','Ano','Substituição','Motivo']

    st.dataframe(df_display, use_container_width=True, hide_index=True,
                 column_config={
                     "ID": st.column_config.TextColumn("ID", width="small"),
                     "Loja": st.column_config.TextColumn("Loja", width="medium"),
                     "Nome": st.column_config.TextColumn("Nome", width="medium"),
                     "CPF": st.column_config.TextColumn("CPF", width="small"),
                     "Situação": st.column_config.TextColumn("Situação", width="small"),
                 })
else:
    st.info("Nenhum registro encontrado com os filtros selecionados.")

# ============================================================================
# ABAS: CADASTRAR / EDITAR
# ============================================================================
st.markdown('<div class="section-title">📝 Cadastro / Edição de Diárias</div>', unsafe_allow_html=True)

aba_cadastrar, aba_editar = st.tabs(["➕ Cadastrar Nova Diária", "✏️ Editar / Excluir Diária"])

# ---------- ABA CADASTRAR ----------
with aba_cadastrar:
    # Dados pré-carregados se estiver em modo de cópia
    default_loja = ""
    default_nome = ""
    default_cpf = ""
    default_datas = ""
    default_qtd = "1"
    default_valor = "70"
    default_banco = ""
    default_sub = "Não"
    default_motivo = ""
    default_data_pag = date.today()
    default_situacao = "PENDENTE"
    default_mes = ""
    default_semana = ""
    default_ano = str(date.today().year)
    default_obs = ""

    if st.session_state.form_limpo:
        st.session_state.form_limpo = False
        st.rerun()

    col1, col2, col3 = st.columns(3)
    with col1:
        loja = st.text_input("🏪 Loja *", value=default_loja, placeholder="Ex: ASSAÍ ATACADISTA BELÉM")
        nome = st.text_input("👤 Nome Completo do Colaborador *", value=default_nome, placeholder="Nome completo")
        cpf = st.text_input("🆔 CPF *", value=default_cpf, placeholder="000.000.000-00")
    with col2:
        data_exec = st.text_input("📅 Datas de Execução *", value=default_datas,
                                   placeholder="Ex: 06/07, 07/07, 08/07/2026")
        quantidade = st.number_input("📊 Quantidade de Dias *", min_value=1, max_value=31, value=int(default_qtd) if default_qtd.isdigit() else 1)
        valor_unit = st.number_input("💲 Valor Unitário (R$) *", min_value=0.0, value=float(default_valor) if default_valor else 70.0, step=10.0, format="%.2f")
    with col3:
        dados_bancarios = st.text_input("🏦 Dados Bancários / PIX *", value=default_banco,
                                         placeholder="Chave PIX, Agência, Conta...")
        substituicao = st.selectbox("🔄 Substituição?", options=["Não", "Sim"], index=0 if default_sub=="Não" else 1)
        motivo = st.text_input("📝 Motivo da Diária *", value=default_motivo,
                                placeholder="Ex: Reposição de funcionários")

    col4, col5, col6 = st.columns(3)
    with col4:
        data_pagamento = st.date_input("📆 Data de Pagamento", value=default_data_pag)
        situacao = st.selectbox("💳 Situação *", options=["PENDENTE", "PAGO"], index=0 if default_situacao=="PENDENTE" else 1)
    with col5:
        mes_options = ["jan","fev","mar","abr","mai","jun","jul","ago","set","out","nov","dez"]
        mes_idx = mes_options.index(default_mes) if default_mes in mes_options else date.today().month - 1
        mes = st.selectbox("📅 Mês", options=mes_options, index=mes_idx)
        semana_options = ["1º Semana","2º Semana","3º Semana","4º Semana"]
        sem_idx = semana_options.index(default_semana) if default_semana in semana_options else 0
        semana = st.selectbox("📆 Semana", options=semana_options, index=sem_idx)
    with col6:
        ano = st.text_input("📅 Ano", value=default_ano, placeholder="2026")
        observacoes = st.text_area("📝 Observações", value=default_obs, placeholder="Observações adicionais...", height=100)

    st.markdown("---")
    st.markdown("**📎 Anexar Comprovante**")
    comprovante_file = st.file_uploader("Envie o comprovante de pagamento (PDF, JPG, PNG)",
                                         type=["pdf", "jpg", "jpeg", "png"],
                                         key="cadastro_comprovante")

    total_calc = calcular_total(quantidade, valor_unit)
    st.markdown(f"**💰 Total Calculado: R$ {total_calc:,.2f}**")

    col_btn1, col_btn2, col_btn3 = st.columns([1,1,3])
    with col_btn1:
        st.markdown('<div class="btn-cadastrar">', unsafe_allow_html=True)
        if st.button("💾 Cadastrar Diária", use_container_width=True):
            if not loja or not nome or not cpf or not data_exec or not dados_bancarios or not motivo:
                mostrar_mensagem("⚠️ Preencha todos os campos obrigatórios (*)", "erro")
                st.rerun()
            else:
                df = st.session_state.df.copy()
                novo_id = gerar_id(df)
                comp_path = salvar_comprovante(comprovante_file, novo_id)

                novo_registro = {
                    'ID': novo_id,
                    'LOJA': loja.strip().upper(),
                    'NOME': nome.strip().upper(),
                    'CPF': cpf.strip(),
                    'DATA_EXECUCAO': data_exec.strip(),
                    'QUANTIDADE': str(quantidade),
                    'VALOR_UNITARIO': f"{valor_unit:.2f}",
                    'TOTAL': str(total_calc),
                    'DADOS_BANCARIOS': dados_bancarios.strip(),
                    'SUBSTITUICAO': substituicao,
                    'MOTIVO': motivo.strip(),
                    'DATA_PAGAMENTO': str(data_pagamento),
                    'SITUACAO': situacao,
                    'MES': mes,
                    'SEMANA': semana,
                    'ANO': str(ano),
                    'COMPROVANTE': comp_path,
                    'OBSERVACOES': observacoes.strip(),
                }

                df = pd.concat([df, pd.DataFrame([novo_registro])], ignore_index=True)
                salvar_dados(df)
                st.session_state.df = carregar_dados()
                mostrar_mensagem(f"✅ Diária {novo_id} cadastrada com sucesso!", "sucesso")
                st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

    with col_btn2:
        st.markdown('<div class="btn-limpar">', unsafe_allow_html=True)
        if st.button("🧹 Limpar Formulário", use_container_width=True):
            st.session_state.form_limpo = True
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

# ---------- ABA EDITAR / EXCLUIR ----------
with aba_editar:
    df_edit = st.session_state.df.copy()
    if df_edit.empty:
        st.info("Não há registros para editar.")
    else:
        # Combo para selecionar registro
        opcoes = [f"{row['ID']} - {row['NOME']} ({row['LOJA']})" for _, row in df_edit.iterrows()]
        opcoes.insert(0, "Selecione um registro...")
        opcao_selecionada = st.selectbox("Selecione a diária para editar", options=opcoes)

        if opcao_selecionada != "Selecione um registro...":
            id_sel = opcao_selecionada.split(" - ")[0]
            registro = df_edit[df_edit['ID'] == id_sel].iloc[0]

            col1, col2, col3 = st.columns(3)
            with col1:
                ed_loja = st.text_input("🏪 Loja", value=registro['LOJA'], key="ed_loja")
                ed_nome = st.text_input("👤 Nome Completo", value=registro['NOME'], key="ed_nome")
                ed_cpf = st.text_input("🆔 CPF", value=registro['CPF'], key="ed_cpf")
            with col2:
                ed_data_exec = st.text_input("📅 Datas de Execução", value=registro['DATA_EXECUCAO'], key="ed_data_exec")
                try:
                    ed_qtd_default = int(float(str(registro['QUANTIDADE']).replace(',', '.')))
                except:
                    ed_qtd_default = 1
                ed_quantidade = st.number_input("📊 Quantidade", min_value=1, max_value=31, value=ed_qtd_default, key="ed_qtd")
                try:
                    ed_val_default = float(str(registro['VALOR_UNITARIO']).replace(',', '.'))
                except:
                    ed_val_default = 70.0
                ed_valor_unit = st.number_input("💲 Valor Unitário (R$)", min_value=0.0, value=ed_val_default, step=10.0, format="%.2f", key="ed_val")
            with col3:
                ed_banco = st.text_input("🏦 Dados Bancários", value=registro['DADOS_BANCARIOS'], key="ed_banco")
                ed_sub = st.selectbox("🔄 Substituição?", options=["Não", "Sim"],
                                       index=0 if registro['SUBSTITUICAO']=="Não" else 1, key="ed_sub")
                ed_motivo = st.text_input("📝 Motivo", value=registro['MOTIVO'], key="ed_motivo")

            col4, col5, col6 = st.columns(3)
            with col4:
                try:
                    dp = pd.to_datetime(registro['DATA_PAGAMENTO']).date()
                except:
                    dp = date.today()
                ed_data_pag = st.date_input("📆 Data Pagamento", value=dp, key="ed_data_pag")
                ed_situacao = st.selectbox("💳 Situação", options=["PENDENTE", "PAGO"],
                                            index=0 if registro['SITUACAO']=="PENDENTE" else 1, key="ed_situ")
            with col5:
                mes_options = ["jan","fev","mar","abr","mai","jun","jul","ago","set","out","nov","dez"]
                m_idx = mes_options.index(registro['MES']) if registro['MES'] in mes_options else 0
                ed_mes = st.selectbox("📅 Mês", options=mes_options, index=m_idx, key="ed_mes")
                sem_options = ["1º Semana","2º Semana","3º Semana","4º Semana"]
                s_idx = sem_options.index(registro['SEMANA']) if registro['SEMANA'] in sem_options else 0
                ed_semana = st.selectbox("📆 Semana", options=sem_options, index=s_idx, key="ed_sem")
            with col6:
                ed_ano = st.text_input("📅 Ano", value=registro['ANO'], key="ed_ano")
                ed_obs = st.text_area("📝 Observações", value=registro['OBSERVACOES'], height=100, key="ed_obs")

            ed_total = calcular_total(ed_quantidade, ed_valor_unit)
            st.markdown(f"**💰 Total Calculado: R$ {ed_total:,.2f}**")

            # Comprovante atual
            if registro['COMPROVANTE'] and os.path.exists(registro['COMPROVANTE']):
                st.markdown(f"📎 Comprovante atual: `{os.path.basename(registro['COMPROVANTE'])}`")
            else:
                st.markdown("📎 Nenhum comprovante anexado.")

            ed_comprovante = st.file_uploader("Substituir comprovante (deixe vazio para manter)",
                                               type=["pdf", "jpg", "jpeg", "png"],
                                               key="ed_comprovante")

            col_e1, col_e2, col_e3 = st.columns([1,1,1])
            with col_e1:
                st.markdown('<div class="btn-editar">', unsafe_allow_html=True)
                if st.button("💾 Salvar Alterações", use_container_width=True, key="btn_salvar_ed"):
                    df = st.session_state.df.copy()
                    idx = df[df['ID'] == id_sel].index[0]

                    novo_comp = registro['COMPROVANTE']
                    if ed_comprovante is not None:
                        novo_comp = salvar_comprovante(ed_comprovante, id_sel)

                    df.at[idx, 'LOJA'] = ed_loja.strip().upper()
                    df.at[idx, 'NOME'] = ed_nome.strip().upper()
                    df.at[idx, 'CPF'] = ed_cpf.strip()
                    df.at[idx, 'DATA_EXECUCAO'] = ed_data_exec.strip()
                    df.at[idx, 'QUANTIDADE'] = str(ed_quantidade)
                    df.at[idx, 'VALOR_UNITARIO'] = f"{ed_valor_unit:.2f}"
                    df.at[idx, 'TOTAL'] = str(ed_total)
                    df.at[idx, 'DADOS_BANCARIOS'] = ed_banco.strip()
                    df.at[idx, 'SUBSTITUICAO'] = ed_sub
                    df.at[idx, 'MOTIVO'] = ed_motivo.strip()
                    df.at[idx, 'DATA_PAGAMENTO'] = str(ed_data_pag)
                    df.at[idx, 'SITUACAO'] = ed_situacao
                    df.at[idx, 'MES'] = ed_mes
                    df.at[idx, 'SEMANA'] = ed_semana
                    df.at[idx, 'ANO'] = str(ed_ano)
                    df.at[idx, 'COMPROVANTE'] = novo_comp
                    df.at[idx, 'OBSERVACOES'] = ed_obs.strip()

                    salvar_dados(df)
                    st.session_state.df = carregar_dados()
                    mostrar_mensagem(f"✅ Diária {id_sel} atualizada com sucesso!", "sucesso")
                    st.rerun()
                st.markdown('</div>', unsafe_allow_html=True)

            with col_e2:
                st.markdown('<div class="btn-excluir">', unsafe_allow_html=True)
                if st.button("🗑️ Excluir Registro", use_container_width=True, key="btn_excluir"):
                    st.session_state.id_para_excluir = id_sel
                    st.rerun()
                st.markdown('</div>', unsafe_allow_html=True)

            with col_e3:
                st.markdown('<div class="btn-limpar">', unsafe_allow_html=True)
                if st.button("🧹 Cancelar Edição", use_container_width=True, key="btn_cancel_ed"):
                    st.rerun()
                st.markdown('</div>', unsafe_allow_html=True)

# ============================================================================
# CONFIRMAÇÃO DE EXCLUSÃO
# ============================================================================
if 'id_para_excluir' in st.session_state and st.session_state.id_para_excluir:
    id_exc = st.session_state.id_para_excluir
    st.warning(f"⚠️ Tem certeza que deseja excluir a diária **{id_exc}**?")
    c1, c2 = st.columns([1,3])
    with c1:
        if st.button("✅ Sim, excluir", key="conf_excluir"):
            df = st.session_state.df.copy()
            reg_exc = df[df['ID'] == id_exc]
            if not reg_exc.empty and reg_exc.iloc[0]['COMPROVANTE']:
                try:
                    os.remove(reg_exc.iloc[0]['COMPROVANTE'])
                except:
                    pass
            df = df[df['ID'] != id_exc].reset_index(drop=True)
            salvar_dados(df)
            st.session_state.df = carregar_dados()
            st.session_state.id_para_excluir = None
            mostrar_mensagem(f"🗑️ Diária {id_exc} excluída com sucesso!", "sucesso")
            st.rerun()
    with c2:
        if st.button("❌ Não, cancelar", key="canc_excluir"):
            st.session_state.id_para_excluir = None
            st.rerun()

# ============================================================================
# EXPORTAÇÃO E COMPROVANTES
# ============================================================================
st.markdown('<div class="section-title">📤 Exportação e Comprovantes</div>', unsafe_allow_html=True)

col_exp1, col_exp2 = st.columns(2)
with col_exp1:
    st.markdown('<div class="btn-exportar">', unsafe_allow_html=True)
    if st.button("📥 Exportar para Excel", use_container_width=True):
        nome_arquivo = f"relatorio_diarias_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        caminho_export = os.path.join(PASTA_COMPROVANTES, nome_arquivo)
        with pd.ExcelWriter(caminho_export, engine='openpyxl') as writer:
            df_filtrado_sidebar.to_excel(writer, sheet_name='Diarias', index=False)
        with open(caminho_export, "rb") as f:
            st.download_button(
                label="⬇️ Baixar arquivo Excel",
                data=f,
                file_name=nome_arquivo,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
    st.markdown('</div>', unsafe_allow_html=True)

with col_exp2:
    # Listar comprovantes disponíveis
    comps = []
    for _, row in df_filtrado_sidebar.iterrows():
        if row['COMPROVANTE'] and os.path.exists(row['COMPROVANTE']):
            comps.append((row['ID'], row['NOME'], row['COMPROVANTE']))

    if comps:
        comp_opcoes = [f"{c[0]} - {c[1]}" for c in comps]
        comp_sel = st.selectbox("📎 Visualizar comprovante", options=["Selecione..."] + comp_opcoes)
        if comp_sel != "Selecione...":
            id_comp = comp_sel.split(" - ")[0]
            for c in comps:
                if c[0] == id_comp:
                    with open(c[2], "rb") as f:
                        ext = Path(c[2]).suffix.lower()
                        mime = "application/pdf" if ext == ".pdf" else "image/jpeg"
                        st.download_button(
                            label=f"⬇️ Baixar comprovante {id_comp}",
                            data=f,
                            file_name=os.path.basename(c[2]),
                            mime=mime,
                            use_container_width=True
                        )
                    break
    else:
        st.info("Nenhum comprovante anexado nos registros filtrados.")

# ============================================================================
# RODAPÉ
# ============================================================================
st.markdown('<div class="footer">Sistema de Controle de Diárias © 2026 — Desenvolvido para gestão eficiente de pagamentos</div>', unsafe_allow_html=True)
