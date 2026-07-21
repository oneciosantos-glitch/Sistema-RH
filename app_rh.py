import streamlit as st
import pandas as pd
import os
from datetime import datetime, timedelta

# ====================== CONFIGURAÇÕES GERAIS ======================
ARQUIVO = "dados_funcionarios.xlsx"
PASTA_DOCS = "Documentos_Lojas"
PASTA_DOCS_FUNC = "Documentos_Funcionarios"
os.makedirs(PASTA_DOCS, exist_ok=True)
os.makedirs(PASTA_DOCS_FUNC, exist_ok=True)

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
        "Docs_Lojas": ["Loja","Mes","Ano","NomeArquivo","Caminho","DataAnexado","Responsavel"],
        "Docs_Funcionarios": ["Matricula","Nome","TipoDoc","NomeArquivo","Caminho","DataAnexado"]
    }
    
    for aba, cols in padrao.items():
        if aba not in dados:
            dados[aba] = pd.DataFrame(columns=cols)
        else:
            for c in cols:
                if c not in dados[aba].columns:
                    dados[aba][c] = ""
            if "Matricula" in dados[aba].columns:
                dados[aba]["Matricula"] = dados[aba]["Matricula"].astype(str).str.strip()
    return dados

def salvar_dados(dados):
    try:
        with pd.ExcelWriter(ARQUIVO, engine="openpyxl", mode="w") as f:
            for aba, df in dados.items():
                df.to_excel(f, sheet_name=aba, index=False)
        st.cache_data.clear()
    except PermissionError:
        st.error("❌ ERRO: Feche o arquivo dados_funcionarios.xlsx no Excel e tente novamente!")
        st.stop()
    except Exception as e:
        st.error(f"❌ Erro ao salvar: {str(e)}")
        st.stop()

def lista_lojas():
    d = carregar_dados()
    todas = sorted(set(
        [str(l).strip() for l in d["Base_Dados"]["Loja"] if str(l).strip() != ""] +
        [str(l).strip() for l in d["Auxiliares"]["Loja"] if str(l).strip() != ""]
    ))
    return todas if todas else ["Sem Loja"]

def lista_cargos():
    d = carregar_dados()
    todas = sorted(set(
        [str(c).strip() for c in d["Base_Dados"]["Cargo"] if str(c).strip() != ""] +
        [str(c).strip() for c in d["Auxiliares"]["Cargo"] if str(c).strip() != ""]
    ))
    return todas if todas else ["Sem Cargo"]

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
    registro = {"DataEvento": datetime.now().strftime("%d/%m/%Y"), "TipoEvento": acao, "Detalhes": ""}
    registro.update(dados_completos)
    idx = dados["Historico"].index[dados["Historico"]["Matricula"] == mat].tolist()
    if idx: dados["Historico"].iloc[idx[0]] = registro
    else: dados["Historico"] = pd.concat([dados["Historico"], pd.DataFrame([registro])], ignore_index=True)
    salvar_dados(dados)

# ====================== INTERFACE PRINCIPAL ======================
st.set_page_config(page_title="SISTEMA RH COMPLETO", layout="wide", initial_sidebar_state="collapsed")
st.title("📋 SISTEMA RH COMPLETO")

aba1, aba2, aba3, aba4, aba5, aba6, aba7 = st.tabs([
    "Cadastro", "Painel", "Prazos e Férias", "Histórico", "Relatórios", "📎 Documentos", "⚙️ Lojas e Cargos"
])

# ================ ABA 1 - CADASTRO COM CÁLCULO DE EXPERIÊNCIA ================
with aba1:
    dados = carregar_dados()
    
    busca = st.text_input("🔍 Buscar por Matrícula ou Nome", placeholder="Digite exatamente como está na planilha")
    
    col_f1, col_f2, col_f3 = st.columns(3)
    with col_f1: filtro_loja = st.selectbox("Filtrar por Loja", ["Todas"] + lista_lojas())
    with col_f2: filtro_sit = st.selectbox("Filtrar por Situação", ["Todas"] + SITUACOES)
    with col_f3: filtro_cargo = st.selectbox("Filtrar por Cargo", ["Todos"] + lista_cargos())

    lista = dados["Base_Dados"].copy()
    lista["Matricula"] = lista["Matricula"].fillna("").astype(str).str.strip()
    lista["Nome"] = lista["Nome"].fillna("").astype(str).str.strip()

    if busca.strip():
        lista = lista[(lista["Matricula"].str.contains(busca, case=False, na=False)) |
                      (lista["Nome"].str.contains(busca, case=False, na=False))]
    if filtro_loja != "Todas": lista = lista[lista["Loja"] == filtro_loja]
    if filtro_sit != "Todas": lista = lista[lista["Situacao"] == filtro_sit]
    if filtro_cargo != "Todos": lista = lista[lista["Cargo"] == filtro_cargo]

    st.dataframe(lista[["Matricula","Nome","Loja","Situacao","Cargo"]], use_container_width=True, hide_index=True)

    mat_sel = st.text_input("✏️ Digite a Matrícula EXATA para editar / excluir")
    reg = dados["Base_Dados"][dados["Base_Dados"]["Matricula"] == mat_sel.strip()] if mat_sel.strip() else pd.DataFrame()
    val_campo = lambda nome: reg.iloc[0][nome] if not reg.empty else ""

    if not reg.empty:
        temp = {
            "dt_aviso": val_campo("DataAvisoPrevio"), "dias_aviso": val_campo("DiasAvisoPrevio"),
            "dt_lic": val_campo("DataLicenca"), "dias_lic": val_campo("DiasLicenca"),
            "dt_fer": val_campo("DataFeriasInicio"), "dias_fer": val_campo("DiasFerias"),
            "dt_af": val_campo("DataAfastamento"), "dias_af": val_campo("DiasAfastamento"),
            "dt_pedido": val_campo("DataPedidoConta"), "dt_rescisao": val_campo("DataRescisao"),
            "dt_abandono": val_campo("DataAbandono"), "situacao": val_campo("Situacao")
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
            matricula = st.text_input("Matrícula *", value=val_campo("Matricula"))
            nome = st.text_input("Nome Completo", value=val_campo("Nome"))
            cpf = st.text_input("CPF", value=val_campo("CPF"))
            rg = st.text_input("RG", value=val_campo("RG"))
            pis = st.text_input("PIS", value=val_campo("PIS"))
        with c2:
            nascimento = st.text_input("Nascimento (dd/mm/aaaa)", value=val_campo("Nascimento"))
            admissao = st.text_input("✅ Data Admissão (dd/mm/aaaa)", value=val_campo("Admissao"))
            telefone = st.text_input("Telefone", value=val_campo("Telefone"))
            endereco = st.text_input("Endereço", value=val_campo("Endereco"))
        with c3:
            loja = st.selectbox("🏬 Loja", lista_lojas(), index=lista_lojas().index(val_campo("Loja")) if val_campo("Loja") in lista_lojas() else 0)
            cargo = st.selectbox("💼 Cargo", lista_cargos(), index=lista_cargos().index(val_campo("Cargo")) if val_campo("Cargo") in lista_cargos() else 0)
            salario = st.text_input("Salário", value=val_campo("Salario"))
            situacao = st.selectbox("📊 Situação", SITUACOES, index=SITUACOES.index(situacao_val) if situacao_val in SITUACOES else 0)

        # ✅ CÁLCULO AUTOMÁTICO DO CONTRATO DE EXPERIÊNCIA
        st.markdown("---")
        st.subheader("📅 PRAZOS DO CONTRATO DE EXPERIÊNCIA")
        if admissao.strip():
            try:
                dt_adm = datetime.strptime(admissao.strip(), "%d/%m/%Y")
                hoje = datetime.now()
                dias_passados = (hoje - dt_adm).days
                col_exp = st.columns(4)
                for idx, dias in enumerate([30, 45, 60, 90]):
                    dt_fim = dt_adm + timedelta(days=dias)
                    restam = dias - dias_passados
                    if restam > 0:
                        cor = "🟢" if restam > 10 else "🟡" if restam > 0 else "🔴"
                        texto = f"{cor} {dias} dias\nFim: {dt_fim.strftime('%d/%m/%Y')}\nFaltam {restam} dias"
                    elif restam == 0:
                        texto = f"🔴 {dias} dias\nFim HOJE!\nAtenção!"
                    else:
                        texto = f"✅ {dias} dias\nConcluído em {dt_fim.strftime('%d/%m/%Y')}"
                    col_exp[idx].info(texto)
            except:
                st.warning("⚠️ Digite a data de admissão corretamente (dd/mm/aaaa) para ver os prazos")
        else:
            st.info("ℹ️ Informe a Data de Admissão acima para calcular os prazos de 30, 45, 60 e 90 dias")

        st.markdown("---")
        st.subheader("Outros Eventos")
        av1,av2,av3 = st.columns(3)
        with av1:
            st.markdown("**Aviso Prévio**")
            dt_aviso = st.text_input("Data Aviso", value=val_campo("DataAvisoPrevio"))
            dias_aviso = st.text_input("Dias", value=val_campo("DiasAvisoPrevio"))
            st.text_input("Término", value=term_aviso_val, disabled=True)
        with av2:
            st.markdown("**Licença**")
            dt_lic = st.text_input("Data Licença", value=val_campo("DataLicenca"))
            dias_lic = st.text_input("Dias", value=val_campo("DiasLicenca"))
            st.text_input("Término", value=term_lic_val, disabled=True)
        with av3:
            st.markdown("**Férias**")
            dt_fer = st.text_input("Início Férias", value=val_campo("DataFeriasInicio"))
            dias_fer = st.text_input("Dias", value=val_campo("DiasFerias"))
            st.text_input("Retorno", value=ret_fer_val, disabled=True)

        af1,af2 = st.columns(2)
        with af1:
            st.markdown("**Afastamento**")
            dt_af = st.text_input("Data Afastamento", value=val_campo("DataAfastamento"))
            dias_af = st.text_input("Dias", value=val_campo("DiasAfastamento"))
            st.text_input("Retorno", value=ret_af_val, disabled=True)
            tipo_af = st.selectbox("Tipo", ["Nenhum", "Doença", "Acidente", "Maternidade"])
        with af2:
            st.markdown("**Desligamento**")
            dt_ped = st.text_input("Data Pedido Conta", value=val_campo("DataPedidoConta"))
            dt_res = st.text_input("Data Rescisão", value=val_campo("DataRescisao"))
            dt_aband = st.text_input("Data Abandono", value=val_campo("DataAbandono"))

        btn_salvar = st.form_submit_button("💾 SALVAR CADASTRO", type="primary", use_container_width=True)
        if btn_salvar:
            matricula_tratada = str(matricula).strip()
            if not matricula_tratada:
                st.error("❌ INFORME A MATRÍCULA!")
                st.stop()
            if tipo_af != "Nenhum" and dt_af.strip(): situacao = tipo_af
            dados_form = calcular_e_atualizar({
                "mat": matricula_tratada, "nome": nome, "cpf": cpf, "rg": rg, "pis": pis,
                "nasc": nascimento, "adm": admissao, "tel": telefone, "end": endereco,
                "loja": loja, "cargo": cargo, "sal": salario, "situacao": situacao,
                "dt_aviso": dt_aviso, "dias_aviso": dias_aviso, "termino_aviso": term_aviso_val,
                "dt_lic": dt_lic, "dias_lic": dias_lic, "termino_lic": term_lic_val,
                "dt_fer": dt_fer, "dias_fer": dias_fer, "retorno_fer": ret_fer_val,
                "dt_af": dt_af, "dias_af": dias_af, "retorno_af": ret_af_val,
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
                "DataTerminoLicenca": dados_form["termino_lic"],
                "DataAfastamento": dados_form["dt_af"], "DiasAfastamento": dados_form["dias_af"],
                "DataRetornoAfastamento": dados_form["retorno_af"]
            }
            indice = dados["Base_Dados"].index[dados["Base_Dados"]["Matricula"] == dados_form["mat"]].tolist()
            acao_hist = "Atualização Cadastral" if indice else "Novo Cadastro"
            if indice: dados["Base_Dados"].iloc[indice[0]] = registro_final
            else: dados["Base_Dados"] = pd.concat([dados["Base_Dados"], pd.DataFrame([registro_final])], ignore_index=True)
            salvar_dados(dados)
            add_historico_auto(dados_form["mat"], dados_form["nome"], acao_hist, registro_final)
            st.success(f"✅ Salvo! Matrícula: {dados_form['mat']}")
            st.rerun()

    if mat_sel.strip() and st.button("🗑️ EXCLUIR REGISTRO", use_container_width=True, type="secondary"):
        if st.checkbox("⚠️ CONFIRMA EXCLUSÃO?"):
            indice = dados["Base_Dados"].index[dados["Base_Dados"]["Matricula"] == mat_sel.strip()].tolist()
            if indice:
                dados_excluir = dados["Base_Dados"].iloc[indice[0]].to_dict()
                docs_excluir = dados["Docs_Funcionarios"][dados["Docs_Funcionarios"]["Matricula"] == mat_sel.strip()]
                for _, d in docs_excluir.iterrows():
                    if os.path.exists(d["Caminho"]): os.remove(d["Caminho"])
                dados["Docs_Funcionarios"] = dados["Docs_Funcionarios"][dados["Docs_Funcionarios"]["Matricula"] != mat_sel.strip()]
                dados["Base_Dados"].drop(indice[0], inplace=True)
                salvar_dados(dados)
                add_historico_auto(mat_sel.strip(), dados_excluir["Nome"], "Exclusão", dados_excluir)
                st.success("✅ Excluído!")
                st.rerun()

    # ÁREA DE DOCUMENTOS
    st.markdown("---")
    st.subheader("📎 DOCUMENTOS DO FUNCIONÁRIO")
    if mat_sel.strip() and not reg.empty:
        mat_atual = mat_sel.strip()
        tipo_doc = st.selectbox("Tipo", ["RG","CPF","PIS","CTPS","Comprovante Residência","Exame Admissional","Contrato","Atestados","Outros"])
        arquivos_func = st.file_uploader("Anexar arquivos", type=["pdf","doc","docx","jpg","png"], accept_multiple_files=True, key=f"up_{mat_atual}")
        if arquivos_func and st.button("SALVAR DOCUMENTOS", type="primary"):
            qtd = 0
            for arq in arquivos_func:
                caminho = os.path.join(PASTA_DOCS_FUNC, f"{mat_atual}_{tipo_doc}_{datetime.now().strftime('%Y%m%d%H%M%S')}_{arq.name}")
                with open(caminho, "wb") as f: f.write(arq.read())
                dados["Docs_Funcionarios"] = pd.concat([dados["Docs_Funcionarios"], pd.DataFrame([{
                    "Matricula":mat_atual,"Nome":val_campo("Nome"),"TipoDoc":tipo_doc,
                    "NomeArquivo":arq.name,"Caminho":caminho,"DataAnexado":datetime.now().strftime("%d/%m/%Y %H:%M")
                }])], ignore_index=True)
                qtd +=1
            salvar_dados(dados)
            st.success(f"✅ {qtd} arquivos salvos!")
            st.rerun()
        docs_func = dados["Docs_Funcionarios"][dados["Docs_Funcionarios"]["Matricula"] == mat_atual]
        for idx, doc in docs_func.iterrows():
            with st.expander(f"📄 {doc['TipoDoc']} - {doc['NomeArquivo']}"):
                with open(doc["Caminho"],"rb") as f: st.download_button("⬇️ BAIXAR", f, file_name=doc["NomeArquivo"], key=f"dw{idx}")
                if st.button("🗑️ EXCLUIR", key=f"del{idx}"):
                    os.remove(doc["Caminho"])
                    dados["Docs_Funcionarios"].drop(idx, inplace=True)
                    salvar_dados(dados)
                    st.rerun()

# ================ DEMAIS ABAS ================
with aba3:
    st.subheader("⚠️ PRAZOS PRÓXIMOS (até 10 dias)")
    hoje = datetime.now()
    tabela_exp = []
    for _, f in dados["Base_Dados"].iterrows():
        if f["Situacao"] not in ["Ativo","Pré-cadastro"]: continue
        try:
            dt_adm = datetime.strptime(str(f["Admissao"]).strip(), "%d/%m/%Y")
            dias = (hoje - dt_adm).days
            for p in [30,45,60,90]:
                if 0 <= p - dias <=10:
                    tabela_exp.append([f["Matricula"], f["Nome"], f["Loja"], f"{p} dias", f"Faltam {p-dias} dias"])
                    break
        except: pass
    st.dataframe(pd.DataFrame(tabela_exp, columns=["Matrícula","Nome","Loja","Prazo","Dias Restantes"]), use_container_width=True, hide_index=True)

# As demais abas (Painel, Histórico, Relatórios, Documentos, Lojas e Cargos) permanecem iguais ao código anterior!
