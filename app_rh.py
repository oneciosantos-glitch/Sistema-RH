import streamlit as st
import pandas as pd
import os
from datetime import datetime, timedelta

# ====================== CONFIGURAÇÕES GERAIS ======================
ARQUIVO = os.path.join(os.path.dirname(__file__), "dados_funcionarios.xlsx")
PASTA_DOCS = os.path.join(os.path.dirname(__file__), "Documentos_Lojas")
os.makedirs(PASTA_DOCS, exist_ok=True)
os.makedirs(os.path.dirname(ARQUIVO), exist_ok=True)

MESES = ["Todos", "Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"]
MAP_MES = {
    "Jan":1, "Fev":2, "Mar":3, "Abr":4, "Mai":5, "Jun":6,
    "Jul":7, "Ago":8, "Set":9, "Out":10, "Nov":11, "Dez":12
}

SITUACOES = [
    "Ativo",
    "Pré-cadastro",
    "Abandono",
    "Término de Contrato",
    "Demitido S/JC",
    "Demitido C/JC",
    "Pedido de Conta",
    "Rescisão Indireta",
    "Férias",
    "Doença",
    "Acidente",
    "Maternidade"
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
        "Auxiliares": ["SITUACOES"],
        "Docs_Lojas": ["Loja","Mes","Ano","NomeArquivo","Caminho","DataAnexado","Responsavel"]
    }
    
    for aba, cols in padrao.items():
        if aba not in dados:
            dados[aba] = pd.DataFrame(columns=cols)
        else:
            for c in cols:
                if c not in dados[aba].columns:
                    dados[aba][c] = ""
    return dados

def salvar_dados(dados):
    try:
        with pd.ExcelWriter(ARQUIVO, engine="openpyxl", mode="w") as f:
            for aba, df in dados.items():
                df.to_excel(f, sheet_name=aba, index=False)
        st.cache_data.clear()
    except PermissionError:
        st.error("❌ ERRO: O arquivo dados_funcionarios.xlsx está ABERTO ou sem permissão! Feche o Excel e tente novamente.")
        st.stop()
    except Exception as e:
        st.error(f"❌ Erro ao salvar: {str(e)}")
        st.stop()

def atualizar_combo_lojas():
    dados = carregar_dados()
    lojas = sorted(dados["Base_Dados"]["Loja"].dropna().astype(str).unique().tolist())
    return [l for l in lojas if l.strip() != ""] or ["Todas"]

def atualizar_combo_cargos():
    dados = carregar_dados()
    cargos = sorted(dados["Base_Dados"]["Cargo"].dropna().astype(str).unique().tolist())
    return [c for c in cargos if c.strip() != ""] or ["Sem Cargo"]

# ====================== CÁLCULOS E SITUAÇÃO AUTOMÁTICA ======================
def calcular_e_atualizar(form):
    # Aviso Prévio
    if form.get("dt_aviso") and form.get("dias_aviso") and str(form["dias_aviso"]).isdigit():
        try:
            dt = datetime.strptime(form["dt_aviso"], "%d/%m/%Y")
            form["termino_aviso"] = (dt + timedelta(days=int(form["dias_aviso"]))).strftime("%d/%m/%Y")
        except: form["termino_aviso"] = ""
    else: form["termino_aviso"] = ""

    # Licença
    if form.get("dt_lic") and form.get("dias_lic") and str(form["dias_lic"]).isdigit():
        try:
            dt = datetime.strptime(form["dt_lic"], "%d/%m/%Y")
            form["termino_lic"] = (dt + timedelta(days=int(form["dias_lic"]))).strftime("%d/%m/%Y")
        except: form["termino_lic"] = ""
    else: form["termino_lic"] = ""

    # Férias
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

    # Afastamento
    if form.get("dt_af") and form.get("dias_af") and str(form["dias_af"]).isdigit():
        try:
            dt = datetime.strptime(form["dt_af"], "%d/%m/%Y")
            form["retorno_af"] = (dt + timedelta(days=int(form["dias_af"]))).strftime("%d/%m/%Y")
        except: form["retorno_af"] = ""
    else: form["retorno_af"] = ""

    # Desligamentos
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
    idx = dados["Historico"].index[dados["Historico"]["Matricula"].astype(str) == mat].tolist()
    if idx: dados["Historico"].iloc[idx[0]] = registro
    else: dados["Historico"] = pd.concat([dados["Historico"], pd.DataFrame([registro])], ignore_index=True)
    salvar_dados(dados)

# ====================== INTERFACE PRINCIPAL ======================
st.set_page_config(page_title="SISTEMA RH COMPLETO", layout="wide", initial_sidebar_state="collapsed")
st.title("📋 SISTEMA RH COMPLETO")

aba1, aba2, aba3, aba4, aba5, aba6 = st.tabs([
    "Cadastro", "Painel", "Prazos e Férias", "Histórico", "Relatórios", "📎 Documentos das Lojas"
])

# ================ ABA 1 - CADASTRO ================
with aba1:
    dados = carregar_dados()
    busca = st.text_input("🔍 Buscar por Matrícula ou Nome", placeholder="Digite e pressione Enter")
    lista = dados["Base_Dados"].copy()

    lista["Matricula"] = lista["Matricula"].fillna("").astype(str).str.strip()
    lista["Nome"] = lista["Nome"].fillna("").astype(str).str.strip()

    if busca.strip():
        lista = lista[
            (lista["Matricula"].str.contains(busca, case=False, na=False)) |
            (lista["Nome"].str.contains(busca, case=False, na=False))
        ]

    st.dataframe(
        lista[["Matricula","Nome","Loja","Situacao","Cargo"]],
        use_container_width=True,
        hide_index=True
    )

    mat_sel = st.text_input("✏️ Digite a Matrícula para editar / excluir / gerar relatório", placeholder="Informe a matrícula")
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
        term_aviso_val = temp["termino_aviso"]
        term_lic_val = temp["termino_lic"]
        ret_fer_val = temp["retorno_fer"]
        ret_af_val = temp["retorno_af"]
        situacao_val = temp["situacao"]
    else:
        term_aviso_val = term_lic_val = ret_fer_val = ret_af_val = ""
        situacao_val = "Ativo"

    with st.form("form_cadastro", clear_on_submit=False):
        st.subheader("Dados Básicos")
        c1,c2,c3 = st.columns(3)
        with c1:
            matricula = st.text_input("Matrícula", value=reg.iloc[0]["Matricula"] if not reg.empty else "")
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
            loja = st.selectbox("Loja", lojas, index=lojas.index(reg.iloc[0]["Loja"]) if not reg.empty and reg.iloc[0]["Loja"] in lojas else 0)
            cargos = atualizar_combo_cargos()
            cargo = st.selectbox("Cargo", cargos, index=cargos.index(reg.iloc[0]["Cargo"]) if not reg.empty and reg.iloc[0]["Cargo"] in cargos else 0)
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
                st.error("❌ Informe a Matrícula!")
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
            st.success(f"✅ Salvo! Situação alterada para: **{dados_form['situacao']}**")
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

# ================ ABA 2 - PAINEL ================
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

# ================ ABA 3 - PRAZOS E FÉRIAS ================
with aba3:
    hoje = datetime.now()
    st.subheader("⚠️ PRAZOS DE EXPERIÊNCIA PRÓXIMOS (até 10 dias restantes)")
    tabela_exp = []
    for _, func in dados["Base_Dados"].iterrows():
        if func["Situacao"] not in ["Ativo","Pré-cadastro"]: continue
        try:
            dt_adm = datetime.strptime(str(func["Admissao"]).strip(), "%d/%m/%Y")
            dias_corridos = (hoje - dt_adm).days
            for prazo in [30,45,60,90]:
                faltam = prazo - dias_corridos
                if 0 <= faltam <=10:
                    tabela_exp.append([func["Matricula"], func["Nome"], func["Loja"], f"{prazo} dias", f"Faltam {faltam} dias"])
                    break
        except: pass
    st.dataframe(pd.DataFrame(tabela_exp, columns=["Matrícula","Nome","Loja","Prazo Total","Dias Restantes"]), use_container_width=True, hide_index=True)

    st.subheader("🗓️ FÉRIAS - POR MÊS DE ADMISSÃO")
    filtro_loja = st.selectbox("Filtrar por Loja", ["Todas"] + atualizar_combo_lojas(), key="f_loja")
    filtro_mes = st.selectbox("Mês de Admissão", MESES, key="f_mes")
    tabela_ferias = []
    for _, func in dados["Base_Dados"].iterrows():
        if func["Situacao"] not in ["Ativo","Pré-cadastro"]: continue
        if filtro_loja != "Todas" and str(func["Loja"]) != filtro_loja: continue
        try:
            dt_adm = datetime.strptime(str(func["Admissao"]).strip(), "%d/%m/%Y")
            if filtro_mes != "Todos" and dt_adm.month != MAP_MES[filtro_mes]: continue
            meses_casa = (hoje.year - dt_adm.year)*12 + (hoje.month - dt_adm.month)
            if hoje.day < dt_adm.day: meses_casa -=1
            if meses_casa >=23 and meses_casa !=24:
                tabela_ferias.append([func["Matricula"], func["Nome"], func["Loja"], func["Cargo"], func["Admissao"], f"{meses_casa} meses", func["DataFeriasInicio"], func["DiasFerias"], func["DataRetornoFerias"]])
        except: pass
    st.dataframe(pd.DataFrame(tabela_ferias, columns=["Matrícula","Nome","Loja","Cargo","Admissão","Tempo de Casa","Início Férias","Dias","Retorno"]), use_container_width=True, hide_index=True)

# ================ ABA 4 - HISTÓRICO ================
with aba4:
    st.subheader("📝 HISTÓRICO COMPLETO")
    st.dataframe(dados["Historico"][["DataEvento","TipoEvento","Matricula","Nome","Situacao","Detalhes"]], use_container_width=True, hide_index=True)
    st.markdown("---")
    st.subheader("Adicionar Novo Evento")
    with st.form("add_evento"):
        t,d,det = st.columns([1,1,3])
        tipo_evento = t.selectbox("Tipo de Evento", ["Reunião","Atestado","Advertência","Elogio","Outros"])
        data_evento = d.text_input("Data", value=datetime.now().strftime("%d/%m/%Y"))
        detalhe_evento = det.text_input("Detalhes")
        if st.form_submit_button("✅ ADICIONAR EVENTO") and mat_sel.strip():
            reg_func = dados["Base_Dados"][dados["Base_Dados"]["Matricula"].astype(str).str.strip() == mat_sel.strip()]
            if not reg_func.empty:
                dados_evento = reg_func.iloc[0].to_dict()
                novo_reg = {"DataEvento": data_evento, "TipoEvento": tipo_evento, "Detalhes": detalhe_evento}
                novo_reg.update(dados_evento)
                idx_hist = dados["Historico"].index[dados["Historico"]["Matricula"].astype(str).str.strip() == mat_sel.strip()].tolist()
                if idx_hist: dados["Historico"].iloc[idx_hist[0]] = novo_reg
                else: dados["Historico"] = pd.concat([dados["Historico"], pd.DataFrame([novo_reg])], ignore_index=True)
                salvar_dados(dados)
                st.success("Evento adicionado!")
                st.rerun()

# ================ ABA 5 - RELATÓRIOS ================
with aba5:
    st.subheader("📄 GERAR RELATÓRIOS EM EXCEL")
    tipo_rel = st.selectbox("Escolha o relatório", [
        "Prazos de Experiência", "Funcionários Ativos", "Pré-cadastro", "Férias",
        "Afastados (Doença/Acidente/Maternidade)", "Afastamentos", "Avisos Prévios", 
        "Histórico Completo", "Relatório Individual por Funcionário"
    ])
    if tipo_rel == "Relatório Individual por Funcionário":
        mat_rel = st.text_input("Informe a Matrícula do Funcionário")
        if mat_rel.strip():
            func_dados = dados["Base_Dados"][dados["Base_Dados"]["Matricula"].astype(str).str.strip() == mat_rel.strip()]
            func_hist = dados["Historico"][dados["Historico"]["Matricula"].astype(str).str.strip() == mat_rel.strip()]
            if func_dados.empty: st.error("❌ Matrícula não encontrada!")
            else:
                st.success("✅ Funcionário encontrado!")
                if st.button("📥 GERAR RELATÓRIO", type="primary"):
                    with pd.ExcelWriter(f"Relatorio_{mat_rel}_{func_dados.iloc[0]['Nome'].replace(' ','_')}.xlsx", engine="openpyxl") as arq:
                        func_dados.to_excel(arq, index=False, sheet_name="Dados_Cadastrais")
                        if not func_hist.empty: func_hist.to_excel(arq, index=False, sheet_name="Historico_Eventos")
                        else: pd.DataFrame([{"Aviso":"Sem eventos"}]).to_excel(arq, index=False, sheet_name="Historico_Eventos")
                    with open(f"Relatorio_{mat_rel}_{func_dados.iloc[0]['Nome'].replace(' ','_')}.xlsx", "rb") as arq:
                        st.download_button("⬇️ BAIXAR", arq, file_name=f"Relatorio_{mat_rel}_{func_dados.iloc[0]['Nome'].replace(' ','_')}.xlsx")
    else:
        if st.button("📥 GERAR E BAIXAR EXCEL", type="primary"):
            if tipo_rel == "Prazos de Experiência": df_rel = pd.DataFrame(tabela_exp, columns=["Matrícula","Nome","Loja","Prazo Total","Dias Restantes"])
            elif tipo_rel == "Funcionários Ativos": df_rel = dados["Base_Dados"][dados["Base_Dados"]["Situacao"] == "Ativo"]
            elif tipo_rel == "Pré-cadastro": df_rel = dados["Base_Dados"][dados["Base_Dados"]["Situacao"] == "Pré-cadastro"]
            elif tipo_rel == "Férias": df_rel = dados["Base_Dados"][dados["Base_Dados"]["Situacao"] == "Férias"]
            elif tipo_rel == "Afastados (Doença/Acidente/Maternidade)": df_rel = dados["Base_Dados"][dados["Base_Dados"]["Situacao"].isin(["Doença","Acidente","Maternidade"])]
            elif tipo_rel == "Afastamentos": df_rel = dados["Base_Dados"][dados["Base_Dados"]["DataAfastamento"].str.strip() != ""]
            elif tipo_rel == "Avisos Prévios": df_rel = dados["Base_Dados"][dados["Base_Dados"]["DataAvisoPrevio"].str.strip() != ""]
            else: df_rel = dados["Historico"]
            with pd.ExcelWriter("relatorio_temp.xlsx", engine="openpyxl") as arq: df_rel.to_excel(arq, index=False, sheet_name=tipo_rel)
            with open("relatorio_temp.xlsx", "rb") as arq: st.download_button("⬇️ BAIXAR", arq, file_name=f"Relatorio_{tipo_rel.replace(' ','_')}.xlsx")
            os.remove("relatorio_temp.xlsx")

# ================ ABA 6 - DOCUMENTOS DAS LOJAS ================
with aba6:
    st.subheader("📎 GERENCIADOR DE DOCUMENTOS DAS LOJAS")
    lojas = atualizar_combo_lojas()
    f1,f2,f3 = st.columns(3)
    with f1: loja_sel = st.selectbox("Selecione a Loja", lojas)
    with f2: mes_sel = st.selectbox("Mês", MESES)
    with f3: ano_sel = st.selectbox("Ano", [str(a) for a in range(2020, datetime.now().year+2)], index=datetime.now().year-2020)

    st.markdown("---")
    st.subheader("📤 ANEXAR NOVO DOCUMENTO")
    arq_anexo = st.file_uploader("Escolha o arquivo (PDF, Word, Excel, Imagem)", type=["pdf","doc","docx","xls","xlsx","jpg","jpeg","png"])
    resp = st.text_input("Responsável pelo envio")
    if arq_anexo and loja_sel != "Todas" and st.button("✅ SALVAR ANEXO", type="primary"):
        nome_seguro = f"{loja_sel}_{mes_sel}_{ano_sel}_{datetime.now().strftime('%Y%m%d%H%M%S')}_{arq_anexo.name}"
        caminho_completo = os.path.join(PASTA_DOCS, nome_seguro)
        with open(caminho_completo, "wb") as f: f.write(arq_anexo.read())
        novo_doc = {
            "Loja":loja_sel, "Mes":mes_sel, "Ano":ano_sel, "NomeArquivo":arq_anexo.name,
            "Caminho":caminho_completo, "DataAnexado":datetime.now().strftime("%d/%m/%Y %H:%M"), "Responsavel":resp
        }
        dados["Docs_Lojas"] = pd.concat([dados["Docs_Lojas"], pd.DataFrame([novo_doc])], ignore_index=True)
        salvar_dados(dados)
        st.success("✅ Documento anexado com sucesso!")
        st.rerun()

    st.markdown("---")
    st.subheader("📂 DOCUMENTOS ENCONTRADOS")
    filtro = dados["Docs_Lojas"].copy()
    if loja_sel != "Todas": filtro = filtro[filtro["Loja"] == loja_sel]
    if mes_sel != "Todos": filtro = filtro[filtro["Mes"] == mes_sel]
    filtro = filtro[filtro["Ano"] == ano_sel]

    if filtro.empty: st.info("ℹ️ Nenhum documento encontrado para este filtro.")
    else:
        for idx, doc in filtro.iterrows():
            with st.expander(f"📄 {doc['NomeArquivo']} | {doc['Mes']}/{doc['Ano']} | Enviado em {doc['DataAnexado']} por {doc['Responsavel']}"):
                col1,col2,col3 = st.columns(3)
                with open(doc["Caminho"], "rb") as f:
                    col1.download_button("⬇️ BAIXAR", f, file_name=doc["NomeArquivo"], key=f"dl_{idx}")
                if doc["Caminho"].lower().endswith((".pdf",".jpg",".jpeg",".png")):
                    if col2.button("👁️ VISUALIZAR", key=f"vis_{idx}"):
                        if doc["Caminho"].lower().endswith(".pdf"):
                            with open(doc["Caminho"], "rb") as f: st.download_button("Abrir PDF", f, file_name=doc["NomeArquivo"], key=f"abr_{idx}")
                        else: st.image(doc["Caminho"], caption=doc["NomeArquivo"])
                if col3.button("🗑️ EXCLUIR", type="secondary", key=f"ex_{idx}"):
                    if os.path.exists(doc["Caminho"]): os.remove(doc["Caminho"])
                    dados["Docs_Lojas"].drop(idx, inplace=True)
                    salvar_dados(dados)
                    st.success("✅ Documento excluído!")
                    st.rerun()