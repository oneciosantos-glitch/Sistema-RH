import streamlit as st
import pandas as pd
import os
from datetime import datetime, timedelta

# ====================== CONFIGURAÇÕES GERAIS ======================
ARQUIVO = "dados_funcionarios.xlsx"
PASTA_DOCS = "Documentos_Lojas"
os.makedirs(PASTA_DOCS, exist_ok=True)

MESES = ["Todos", "Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"]
MAP_MES = {
    "Jan":1, "Fev":2, "Mar":3, "Abr":4, "Mai":5, "Jun":6,
    "Jul":7, "Ago":8, "Set":9, "Out":10, "Nov":11, "Dez":12
}

SITUACOES = [
    "Ativo", "Pré-cadastro", "Abandono", "Término de Contrato",
    "Demitido S/JC", "Demitido C/JC", "Pedido de Conta",
    "Rescisão Indireta", "Férias", "Doença", "Acidente", "Maternidade"
]

# ====================== BANCO DE DADOS ======================
@st.cache_data(ttl=1)
def carregar_dados():
    try:
        dados = pd.read_excel(ARQUIVO, sheet_name=None, dtype=str, keep_default_na=False)
    except:
        dados = {}
    
    padrao = {
        "Base_Dados": [
            "Matricula","Nome","CPF","RG","PIS","Nascimento","Admissao",
            "Telefone","Endereco","Loja","Cargo","Salario","Situacao",
            "DataAvisoPrevio","DiasAvisoPrevio","DataTerminoAviso",
            "DataFeriasInicio","DiasFerias","DataRetornoFerias",
            "DataPedidoConta","DataRescisao","DataAbandono",
            "DataLicenca","DiasLicenca","DataTerminoLicenca",
            "DataAfastamento","DiasAfastamento","DataRetornoAfastamento"
        ],
        "Historico": [
            "DataEvento","TipoEvento","Matricula","Nome","CPF","RG","PIS",
            "Nascimento","Admissao","Telefone","Endereco","Loja","Cargo",
            "Salario","Situacao","DataAvisoPrevio","DiasAvisoPrevio","DataTerminoAviso",
            "DataFeriasInicio","DiasFerias","DataRetornoFerias",
            "DataPedidoConta","DataRescisao","DataAbandono",
            "DataLicenca","DiasLicenca","DataTerminoLicenca",
            "DataAfastamento","DiasAfastamento","DataRetornoAfastamento","Detalhes"
        ],
        "Auxiliares": ["Loja", "Cargo"],
        "Docs_Lojas": ["Loja","Mes","Ano","NomeArquivo","Caminho","DataAnexado","Responsavel"]
    }
    
    for aba, cols in padrao.items():
        if aba not in dados:
            dados[aba] = pd.DataFrame(columns=cols)
        else:
            for c in cols:
                if c not in dados[aba].columns:
                    dados[aba][c] = ""
            if "Matricula" in dados[aba].columns:
                dados[aba]["Matricula"] = dados[aba]["Matricula"].fillna("").astype(str).str.strip()
    return dados

def salvar_dados(dados):
    try:
        with pd.ExcelWriter(ARQUIVO, engine="openpyxl", mode="w") as f:
            for aba, df in dados.items():
                df.to_excel(f, sheet_name=aba, index=False)
        st.cache_data.clear()
    except PermissionError:
        st.error("❌ ERRO: O arquivo dados_funcionarios.xlsx está ABERTO! Feche o Excel e tente novamente.")
        st.stop()
    except Exception as e:
        st.error(f"❌ Erro ao salvar: {str(e)}")
        st.stop()

def atualizar_combo_lojas():
    dados = carregar_dados()
    lojas_base = sorted(dados["Base_Dados"]["Loja"].dropna().astype(str).unique().tolist())
    lojas_aux = sorted(dados["Auxiliares"]["Loja"].dropna().astype(str).unique().tolist())
    todas = sorted(list(set([l.strip() for l in lojas_base + lojas_aux if str(l).strip() != ""])))
    return todas or ["Sem Loja"]

def atualizar_combo_cargos():
    dados = carregar_dados()
    cargos_base = sorted(dados["Base_Dados"]["Cargo"].dropna().astype(str).unique().tolist())
    cargos_aux = sorted(dados["Auxiliares"]["Cargo"].dropna().astype(str).unique().tolist())
    todas = sorted(list(set([c.strip() for c in cargos_base + cargos_aux if str(c).strip() != ""])))
    return todas or ["Sem Cargo"]

# ====================== CÁLCULOS E SITUAÇÃO AUTOMÁTICA ======================
def calcular_e_atualizar(form):
    if form.get("dt_aviso") and form.get("dias_aviso") and str(form["dias_aviso"]).isdigit():
        try:
            dt = datetime.strptime(form["dt_aviso"], "%d/%m/%Y")
            form["termino_aviso"] = (dt + timedelta(days=int(form["dias_aviso"]))).strftime("%d/%m/%Y")
        except: form["termino_aviso"] = ""
    else: form["termino_aviso"] = ""

    if form.get("dt_lic") and form.get("dias_lic") and str(form["dias_lic"]).isdigit():
        try:
            dt = datetime.strptime(form["dt_lic"], "%d/%m/%Y")
            form["termino_lic"] = (dt + timedelta(days=int(form["dias_lic"]))).strftime("%d/%m/%Y")
        except: form["termino_lic"] = ""
    else: form["termino_lic"] = ""

    if form.get("dt_fer") and form.get("dias_fer") and str(form["dias_fer"]).isdigit():
        try:
            dt = datetime.strptime(form["dt_fer"], "%d/%m/%Y")
            form["retorno_fer"] = (dt + timedelta(days=int(form["dias_fer"]))).strftime("%d/%m/%Y")
            form["situacao"] = "Férias"
        except:
            form["retorno_fer"] = ""
            if form.get("situacao") == "Férias": form["situacao"] = "Ativo"
    else:
        form["retorno_fer"] = ""
        if form.get("situacao") == "Férias": form["situacao"] = "Ativo"

    if form.get("dt_af") and form.get("dias_af") and str(form["dias_af"]).isdigit():
        try:
            dt = datetime.strptime(form["dt_af"], "%d/%m/%Y")
            form["retorno_af"] = (dt + timedelta(days=int(form["dias_af"]))).strftime("%d/%m/%Y")
        except: form["retorno_af"] = ""
    else: form["retorno_af"] = ""

    if form.get("dt_pedido") and form["dt_pedido"].strip():
        form["situacao"] = "Pedido de Conta"
    elif form.get("dt_rescisao") and form["dt_rescisao"].strip():
        form["situacao"] = "Rescisão Indireta"
    elif form.get("dt_abandono") and form["dt_abandono"].strip():
        form["situacao"] = "Abandono"
    elif not any([
        form.get("dt_pedido","").strip(), form.get("dt_rescisao","").strip(),
        form.get("dt_abandono","").strip(), form.get("dt_fer","").strip(),
        form.get("dt_lic","").strip(), form.get("dt_af","").strip()
    ]):
        if form.get("situacao") not in ["Ativo", "Pré-cadastro"]:
            form["situacao"] = "Ativo"
    return form

def add_historico_auto(mat, nome, acao, dados_completos):
    dados = carregar_dados()
    registro = {"DataEvento": datetime.now().strftime("%d/%m/%Y"), "TipoEvento": acao, "Detalhes": "Registro atualizado"}
    registro.update(dados_completos)
    idx = dados["Historico"].index[dados["Historico"]["Matricula"].astype(str).str.strip() == mat].tolist()
    if idx: dados["Historico"].iloc[idx[0]] = registro
    else: dados["Historico"] = pd.concat([dados["Historico"], pd.DataFrame([registro])], ignore_index=True)
    salvar_dados(dados)

# ====================== INTERFACE PRINCIPAL ======================
st.set_page_config(page_title="SISTEMA RH COMPLETO", layout="wide", initial_sidebar_state="collapsed")
st.title("📋 SISTEMA RH COMPLETO")

aba1, aba2, aba3, aba4, aba5, aba6, aba7 = st.tabs([
    "Cadastro", "Painel", "Prazos e Férias", "Histórico", "Relatórios", "📎 Documentos", "⚙️ Lojas e Cargos"
])

# ================ ABA 1 - CADASTRO FUNCIONÁRIOS ================
with aba1:
    dados = carregar_dados()
    busca = st.text_input("🔍 Buscar por Matrícula ou Nome", placeholder="Digite e pressione Enter")
    lista = dados["Base_Dados"].copy()
    lista["Matricula"] = lista["Matricula"].fillna("VAZIO").astype(str).str.strip()
    lista["Nome"] = lista["Nome"].fillna("").astype(str).str.strip()

    if busca.strip():
        lista = lista[
            (lista["Matricula"].str.contains(busca, case=False, na=False)) |
            (lista["Nome"].str.contains(busca, case=False, na=False))
        ]

    st.dataframe(
        lista[["Matricula","Nome","Loja","Situacao","Cargo"]],
        use_container_width=True, hide_index=True
    )

    mat_sel = st.text_input("✏️ Digite a Matrícula para editar / excluir", placeholder="Informe a matrícula")
    reg = pd.DataFrame()
    if mat_sel.strip():
        mat_busca = str(mat_sel).strip()
        reg = dados["Base_Dados"][dados["Base_Dados"]["Matricula"].astype(str).str.strip() == mat_busca]

    if not reg.empty:
        temp = {
            "dt_aviso": reg.iloc[0]["DataAvisoPrevio"], "dias_aviso": reg.iloc[0]["DiasAvisoPrevio"],
            "dt_lic": reg.iloc[0]["DataLicenca"], "dias_lic": reg.iloc[0]["DiasLicenca"],
            "dt_fer": reg.iloc[0]["DataFeriasInicio"], "dias_fer": reg.iloc[0]["DiasFerias"],
            "dt_af": reg.iloc[0]["DataAfastamento"], "dias_af": reg.iloc[0]["DiasAfastamento"],
            "dt_pedido": reg.iloc[0]["DataPedidoConta"], "dt_rescisao": reg.iloc[0]["DataRescisao"],
            "dt_abandono": reg.iloc[0]["DataAbandono"], "situacao": reg.iloc[0]["Situacao"]
        }
        temp = calcular_e_atualizar(temp)
        term_aviso_val, term_lic_val, ret_fer_val, ret_af_val, situacao_val = temp["termino_aviso"], temp["termino_lic"], temp["retorno_fer"], temp["retorno_af"], temp["situacao"]
    else:
        term_aviso_val = term_lic_val = ret_fer_val = ret_af_val = ""
        situacao_val = "Ativo"

    with st.form("form_cadastro", clear_on_submit=False):
        st.subheader("Dados Básicos")
        c1,c2,c3 = st.columns(3)
        with c1:
            matricula = st.text_input("Matrícula *", value=reg.iloc[0]["Matricula"] if not reg.empty else "")
            nome = st.text_input("Nome Completo", value=reg.iloc[0]["Nome"] if not reg.empty else "")
            cpf = st.text_input("CPF", value=reg.iloc[0]["CPF"] if not reg.empty else "")
            rg = st.text_input("RG", value=reg.iloc[0]["RG"] if not reg.empty else "")
            pis = st.text_input("PIS", value=reg.iloc[0]["PIS"] if not reg.empty else "")
        with c2:
            nascimento = st.text_input("Data Nascimento (dd/mm/aaaa)", value=reg.iloc[0]["Nascimento"] if not reg.empty else "")
            admissao = st.text_input("Data Admissão (dd/mm/aaaa)", value=reg.iloc[0]["Admissao"] if not reg.empty else "")
            telefone = st.text_input("Telefone", value=reg.iloc[0]["Telefone"] if not reg.empty else "")
            endereco = st.text_input("Endereço Completo", value=reg.iloc[0]["Endereco"] if not reg.empty else "")
        with c3:
            lojas = atualizar_combo_lojas()
            ind_loja = lojas.index(reg.iloc[0]["Loja"]) if not reg.empty and reg.iloc[0]["Loja"] in lojas else 0
            loja = st.selectbox("Loja", lojas, index=ind_loja)

            cargos = atualizar_combo_cargos()
            ind_cargo = cargos.index(reg.iloc[0]["Cargo"]) if not reg.empty and reg.iloc[0]["Cargo"] in cargos else 0
            cargo = st.selectbox("Cargo", cargos, index=ind_cargo)

            salario = st.text_input("Salário", value=reg.iloc[0]["Salario"] if not reg.empty else "")
            ind_sit = SITUACOES.index(situacao_val) if situacao_val in SITUACOES else 0
            situacao = st.selectbox("Situação", SITUACOES, index=ind_sit)

        st.markdown("---")
        st.subheader("Eventos Trabalhistas")
        av1,av2,av3 = st.columns(3)
        with av1:
            st.markdown("**Aviso Prévio**")
            dt_aviso = st.text_input("Data Aviso", value=reg.iloc[0]["DataAvisoPrevio"] if not reg.empty else "")
            dias_aviso = st.text_input("Dias Aviso", value=reg.iloc[0]["DiasAvisoPrevio"] if not reg.empty else "")
            term_aviso = st.text_input("Término Aviso", value=term_aviso_val, disabled=True)
        with av2:
            st.markdown("**Licença**")
            dt_lic = st.text_input("Data Licença", value=reg.iloc[0]["DataLicenca"] if not reg.empty else "")
            dias_lic = st.text_input("Dias Licença", value=reg.iloc[0]["DiasLicenca"] if not reg.empty else "")
            term_lic = st.text_input("Término Licença", value=term_lic_val, disabled=True)
        with av3:
            st.markdown("**Férias**")
            dt_fer = st.text_input("Início Férias", value=reg.iloc[0]["DataFeriasInicio"] if not reg.empty else "")
            dias_fer = st.text_input("Dias Férias", value=reg.iloc[0]["DiasFerias"] if not reg.empty else "")
            ret_fer = st.text_input("Retorno Férias", value=ret_fer_val, disabled=True)

        af1,af2 = st.columns(2)
        with af1:
            st.markdown("**Afastamento**")
            dt_af = st.text_input("Data Afastamento", value=reg.iloc[0]["DataAfastamento"] if not reg.empty else "")
            dias_af = st.text_input("Dias Afastamento", value=reg.iloc[0]["DiasAfastamento"] if not reg.empty else "")
            ret_af = st.text_input("Retorno Afastamento", value=ret_af_val, disabled=True)
            tipo_af = st.selectbox("Tipo Afastamento", ["Nenhum", "Doença", "Acidente", "Maternidade"])
        with af2:
            st.markdown("**Desligamento**")
            dt_ped = st.text_input("Data Pedido Conta", value=reg.iloc[0]["DataPedidoConta"] if not reg.empty else "")
            dt_res = st.text_input("Data Rescisão", value=reg.iloc[0]["DataRescisao"] if not reg.empty else "")
            dt_aband = st.text_input("Data Abandono", value=reg.iloc[0]["DataAbandono"] if not reg.empty else "")

        btn_salvar = st.form_submit_button("💾 SALVAR CADASTRO", type="primary", use_container_width=True)
        if btn_salvar:
            if not matricula.strip():
                st.error("❌ A MATRÍCULA É OBRIGATÓRIA!")
                st.stop()
            if tipo_af != "Nenhum" and dt_af.strip():
                situacao = tipo_af
            dados_form = calcular_e_atualizar({
                "mat": matricula.strip(), "nome": nome, "cpf": cpf, "rg": rg, "pis": pis,
                "nasc": nascimento, "adm": admissao, "tel": telefone, "end": endereco,
                "loja": loja, "cargo": cargo, "sal": salario, "situacao": situacao,
                "dt_aviso": dt_aviso, "dias_aviso": dias_aviso, "termino_aviso": term_aviso,
                "dt_lic": dt_lic, "dias_lic": dias_lic, "termino_lic": term_lic,
                "dt_fer": dt_fer, "dias_fer": dias_fer, "retorno_fer": ret_fer,
                "dt_af": dt_af, "dias_af": dias_af, "retorno_af": ret_af,
                "dt_pedido": dt_ped, "dt_rescisao": dt_res, "dt_abandono": dt_aband
            })
            registro_final = {
                "Matricula": dados_form["mat"], "Nome": dados_form["nome"], "CPF": dados_form["cpf"],
                "RG": dados_form["rg"], "PIS": dados_form["pis"], "Nascimento": dados_form["nasc"],
                "Admissao": dados_form["adm"], "Telefone": dados_form["tel"], "Endereco": dados_form["end"],
                "Loja": dados_form["loja"], "Cargo": dados_form["cargo"], "Salario": dados_form["sal"],
                "Situacao": dados_form["situacao"], "DataAvisoPrevio": dados_form["dt_aviso"],
                "DiasAvisoPrevio": dados_form["dias_aviso"], "DataTerminoAviso": dados_form["termino_aviso"],
                "DataFeriasInicio": dados_form["dt_fer"], "DiasFerias": dados_form["dias_fer"],
                "DataRetornoFerias": dados_form["retorno_fer"], "DataPedidoConta": dados_form["dt_pedido"],
                "DataRescisao": dados_form["dt_rescisao"], "DataAbandono": dados_form["dt_abandono"],
                "DataLicenca": dados_form["dt_lic"], "DiasLicenca": dados_form["dias_lic"],
                "DataTerminoLicenca": dados_form["termino_lic"], "DataAfastamento": dados_form["dt_af"],
                "DiasAfastamento": dados_form["dias_af"], "DataRetornoAfastamento": dados_form["retorno_af"]
            }
            indice = dados["Base_Dados"].index[dados["Base_Dados"]["Matricula"].astype(str).str.strip() == dados_form["mat"]].tolist()
            acao_hist = "Atualização Cadastral" if indice else "Novo Cadastro"
            if indice: dados["Base_Dados"].iloc[indice[0]] = registro_final
            else: dados["Base_Dados"] = pd.concat([dados["Base_Dados"], pd.DataFrame([registro_final])], ignore_index=True)
            salvar_dados(dados)
            add_historico_auto(dados_form["mat"], dados_form["nome"], acao_hist, registro_final)
            st.success(f"✅ Salvo! Situação: **{dados_form['situacao']}**")
            st.rerun()

    if mat_sel.strip() and st.button("🗑️ EXCLUIR REGISTRO", use_container_width=True, type="secondary"):
        if st.checkbox("⚠️ CONFIRMA EXCLUSÃO PERMANENTE?"):
            indice = dados["Base_Dados"].index[dados["Base_Dados"]["Matricula"].astype(str).str.strip() == mat_sel.strip()].tolist()
            if indice:
                dados_excluir = dados["Base_Dados"].iloc[indice[0]].to_dict()
                dados["Base_Dados"].drop(indice[0], inplace=True)
                salvar_dados(dados)
                add_historico_auto(mat_sel.strip(), dados_excluir["Nome"], "Exclusão de Cadastro", dados_excluir)
                st.success("✅ Registro excluído!")
                st.rerun()

# ================ ABA 7 - CADASTRO DE LOJAS E CARGOS (NOVA!) ================
with aba7:
    st.subheader("⚙️ CADASTRO DE NOVAS LOJAS E CARGOS")
    dados = carregar_dados()

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**🏬 Nova Loja**")
        nova_loja = st.text_input("Nome da Loja", placeholder="Ex: Loja Centro, Filial 2...")
        if st.button("➕ ADICIONAR LOJA", type="primary"):
            if nova_loja.strip():
                existe = dados["Auxiliares"]["Loja"].astype(str).str.strip().eq(nova_loja.strip()).any()
                if not existe:
                    dados["Auxiliares"] = pd.concat([dados["Auxiliares"], pd.DataFrame([{"Loja": nova_loja.strip(), "Cargo": ""}])], ignore_index=True)
                    salvar_dados(dados)
                    st.success(f"✅ Loja **{nova_loja}** cadastrada!")
                    st.rerun()
                else:
                    st.warning("⚠️ Essa loja já está cadastrada!")
            else:
                st.error("❌ Digite o nome da loja!")

    with col2:
        st.markdown("**💼 Novo Cargo**")
        novo_cargo = st.text_input("Nome do Cargo", placeholder="Ex: Vendedor, Caixa...")
        if st.button("➕ ADICIONAR CARGO", type="primary"):
            if novo_cargo.strip():
                existe = dados["Auxiliares"]["Cargo"].astype(str).str.strip().eq(novo_cargo.strip()).any()
                if not existe:
                    dados["Auxiliares"] = pd.concat([dados["Auxiliares"], pd.DataFrame([{"Loja": "", "Cargo": novo_cargo.strip()}])], ignore_index=True)
                    salvar_dados(dados)
                    st.success(f"✅ Cargo **{novo_cargo}** cadastrado!")
                    st.rerun()
                else:
                    st.warning("⚠️ Esse cargo já está cadastrado!")
            else:
                st.error("❌ Digite o nome do cargo!")

    st.markdown("---")
    c1,c2 = st.columns(2)
    with c1:
        st.markdown("**📋 Lojas Cadastradas**")
        lojas_lista = sorted([l.strip() for l in dados["Auxiliares"]["Loja"].dropna().astype(str).unique() if str(l).strip() != ""])
        for loja in lojas_lista:
            cc, ee = st.columns([4,1])
            cc.write(f"• {loja}")
            if ee.button("🗑️", key=f"del_loja_{loja}"):
                dados["Auxiliares"] = dados["Auxiliares"][dados["Auxiliares"]["Loja"].str.strip() != loja]
                salvar_dados(dados)
                st.rerun()

    with c2:
        st.markdown("**📋 Cargos Cadastrados**")
        cargos_lista = sorted([c.strip() for c in dados["Auxiliares"]["Cargo"].dropna().astype(str).unique() if str(c).strip() != ""])
        for cargo in cargos_lista:
            cc, ee = st.columns([4,1])
            cc.write(f"• {cargo}")
            if ee.button("🗑️", key=f"del_cargo_{cargo}"):
                dados["Auxiliares"] = dados["Auxiliares"][dados["Cargo"].str.strip() != cargo]
                salvar_dados(dados)
                st.rerun()

# ================ DEMAIS ABAS (mantidas iguais) ================
with aba2:
    st.subheader("📊 RESUMO GERAL")
    ativos = len(dados["Base_Dados"][dados["Base_Dados"]["Situacao"] == "Ativo"])
    pre_cad = len(dados["Base_Dados"][dados["Base_Dados"]["Situacao"] == "Pré-cadastro"])
    ferias = len(dados["Base_Dados"][dados["Base_Dados"]["Situacao"] == "Férias"])
    licencas = len(dados["Base_Dados"][dados["Base_Dados"]["Situacao"].isin(["Doença","Acidente","Maternidade"])])
    desligados = len(dados["Base_Dados"][~dados["Base_Dados"]["Situacao"].isin(["Ativo","Pré-cadastro","Férias","Doença","Acidente","Maternidade"])])
    c1,c2,c3,c4,c5 = st.columns(5)
    c1.metric("👷 Ativos", ativos)
    c2.metric("📝 Pré-cadastro", pre_cad)
    c3.metric("🏖️ Férias", ferias)
    c4.metric("🏥 Afastados", licencas)
    c5.metric("📤 Desligados", desligados)

with aba3:
    hoje = datetime.now()
    st.subheader("⚠️ PRAZOS DE EXPERIÊNCIA PRÓXIMOS (até 10 dias)")
    tabela_exp = []
    for _, func in dados["Base_Dados"].iterrows():
        if func["Situacao"] not in ["Ativo","Pré-cadastro"]: continue
        try:
            dt_adm = datetime.strptime(str(func["Admissao"]).strip(), "%d/%m/%Y")
            dias = (hoje - dt_adm).days
            for p in [30,45,60,90]:
                if 0 <= p - dias <=10:
                    tabela_exp.append([func["Matricula"], func["Nome"], func["Loja"], f"{p} dias", f"Faltam {p-dias} dias"])
                    break
        except: pass
    st.dataframe(pd.DataFrame(tabela_exp, columns=["Matrícula","Nome","Loja","Prazo","Dias Restantes"]), use_container_width=True, hide_index=True)

    st.subheader("🗓️ FÉRIAS - POR MÊS DE ADMISSÃO")
    filtro_loja = st.selectbox("Loja", ["Todas"] + atualizar_combo_lojas(), key="fl")
    filtro_mes = st.selectbox("Mês", MESES, key="fm")
    tabela_fer = []
    for _, f in dados["Base_Dados"].iterrows():
        if f["Situacao"] not in ["Ativo","Pré-cadastro"]: continue
        if filtro_loja != "Todas" and f["Loja"] != filtro_loja: continue
        try:
            dt = datetime.strptime(str(f["Admissao"]).strip(), "%d/%m/%Y")
            if filtro_mes != "Todos" and dt.month != MAP_MES[filtro_mes]: continue
            meses = (hoje.year - dt.year)*12 + (hoje.month - dt.month) - (1 if hoje.day < dt.day else 0)
            if 23 <= meses < 24:
                tabela_fer.append([f["Matricula"], f["Nome"], f["Loja"], f["Cargo"], f["Admissao"], f"{meses}m", f["DataFeriasInicio"], f["DiasFerias"], f["DataRetornoFerias"]])
        except: pass
    st.dataframe(pd.DataFrame(tabela_fer, columns=["Matrícula","Nome","Loja","Cargo","Admissão","Tempo","Início","Dias","Retorno"]), use_container_width=True, hide_index=True)

with aba4:
    st.subheader("📝 HISTÓRICO")
    st.dataframe(dados["Historico"][["DataEvento","TipoEvento","Matricula","Nome","Situacao","Detalhes"]], use_container_width=True, hide_index=True)
    with st.form("add_ev"):
        t,d,det = st.columns([1,1,3])
        te = t.selectbox("Tipo", ["Reunião","Atestado","Advertência","Elogio","Outros"])
        de = d.text_input("Data", value=datetime.now().strftime("%d/%m/%Y"))
        dee = det.text_input("Detalhes")
        if st.form_submit_button("✅ ADICIONAR") and mat_sel.strip():
            rf = dados["Base_Dados"][dados["Base_Dados"]["Matricula"].str.strip() == mat_sel.strip()]
            if not rf.empty:
                nr = {"DataEvento":de,"TipoEvento":te,"Detalhes":dee}
                nr.update(rf.iloc[0].to_dict())
                ih = dados["Historico"].index[dados["Historico"]["Matricula"].str.strip() == mat_sel.strip()].tolist()
                if ih: dados["Historico"].iloc[ih[0]] = nr
                else: dados["Historico"] = pd.concat([dados["Historico"], pd.DataFrame([nr])], ignore_index=True)
                salvar_dados(dados)
                st.success("Adicionado!")
                st.rerun()

with aba5:
    st.subheader("📄 RELATÓRIOS")
    rel = st.selectbox("Escolha", ["Prazos Experiência","Ativos","Pré-cadastro","Férias","Afastados","Avisos","Histórico","Individual"])
    if rel == "Individual":
        mr = st.text_input("Matrícula")
        if mr.strip():
            fd = dados["Base_Dados"][dados["Base_Dados"]["Matricula"].str.strip() == mr.strip()]
            fh = dados["Historico"][dados["Historico"]["Matricula"].str.strip() == mr.strip()]
            if fd.empty: st.error("Não encontrado")
            elif st.button("GERAR"):
                with pd.ExcelWriter(f"Rel_{mr}_{fd.iloc[0]['Nome'].replace(' ','_')}.xlsx") as arq:
                    fd.to_excel(arq, index=False, sheet_name="Dados")
                    fh.to_excel(arq, index=False, sheet_name="Histórico") if not fh.empty else pd.DataFrame([{"Aviso":"Sem histórico"}]).to_excel(arq, index=False, sheet_name="Histórico")
                with open(f"Rel_{mr}_{fd.iloc[0]['Nome'].replace(' ','_')}.xlsx","rb") as f:
                    st.download_button("⬇️ BAIXAR", f, file_name=f"Rel_{mr}.xlsx")
    elif st.button("GERAR E BAIXAR"):
        if rel == "Prazos Experiência": df = pd.DataFrame(tabela_exp, columns=["Matrícula","Nome","Loja","Prazo","Dias Restantes"])
        elif rel == "Ativos": df = dados["Base_Dados"][dados["Base_Dados"]["Situacao"] == "Ativo"]
        elif rel == "Pré-cadastro": df = dados["Base_Dados"][dados["Base_Dados"]["Situacao"] == "Pré-cadastro"]
        elif rel == "Férias": df = dados["Base_Dados"][dados["Base_Dados"]["Situacao"] == "Férias"]
        elif rel == "Afastados": df = dados["Base_Dados"][dados["Base_Dados"]["Situacao"].isin(["Doença","Acidente","Maternidade"])]
        elif rel == "Avisos": df = dados["Base_Dados"][dados["Base_Dados"]["DataAvisoPrevio"].str.strip()!=""]
        else: df = dados["Historico"]
        with pd.ExcelWriter("rel_temp.xlsx") as arq: df.to_excel(arq, index=False, sheet_name=rel)
        with open("rel_temp.xlsx","rb") as f: st.download_button("⬇️ BAIXAR", f, file_name=f"Rel_{rel.replace(' ','_')}.xlsx")
        os.remove("rel_temp.xlsx")

with aba6:
    st.subheader("📎 DOCUMENTOS DAS LOJAS")
    ls = atualizar_combo_lojas()
    l,m,a = st.columns(3)
    sl = l.selectbox("Loja", ls)
    sm = m.selectbox("Mês", MESES)
    sa = a.selectbox("Ano", [str(x) for x in range(2020, datetime.now().year+2)], index=datetime.now().year-2020)
    st.markdown("---")
    arq = st.file_uploader("Anexar", type=["pdf","doc","docx","xls","xlsx","jpg","png"])
    resp = st.text_input("Responsável")
    if arq and st.button("SALVAR"):
        nome = f"{sl}_{sm}_{sa}_{datetime.now().strftime('%Y%m%d%H%M%S')}_{arq.name}"
        cam = os.path.join(PASTA_DOCS, nome)
        with open(cam,"wb") as f: f.write(arq.read())
        dados["Docs_Lojas"] = pd.concat([dados["Docs_Lojas"], pd.DataFrame([{"Loja":sl,"Mes":sm,"Ano":sa,"NomeArquivo":arq.name,"Caminho":cam,"DataAnexado":datetime.now().strftime("%d/%m/%Y %H:%M"),"Responsavel":resp}])], ignore_index=True)
        salvar_dados(dados)
        st.success("Anexado!")
        st.rerun()
    st.markdown("---")
    filt = dados["Docs_Lojas"].copy()
    if sl != "Todas": filt = filt[filt["Loja"]==sl]
    if sm != "Todos": filt = filt[filt["Mes"]==sm]
    filt = filt[filt["Ano"]==sa]
    if filt.empty: st.info("Nenhum documento")
    else:
        for i,d in filt.iterrows():
            with st.expander(f"📄 {d['NomeArquivo']} | {d['Mes']}/{d['Ano']}"):
                with open(d["Caminho"],"rb") as f: st.download_button("⬇️ BAIXAR", f, file_name=d["NomeArquivo"], key=f"d{i}")
                if st.button("🗑️ EXCLUIR", key=f"x{i}"):
                    os.remove(d["Caminho"])
                    dados["Docs_Lojas"].drop(i,inplace=True)
                    salvar_dados(dados)
                    st.rerun()
