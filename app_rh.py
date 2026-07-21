import streamlit as st
import pandas as pd
import os
from datetime import datetime, timedelta

# ====================== CONFIGURAÇÕES ======================
ARQUIVO = "dados_funcionarios.xlsx"
PASTA_DOCS_LOJAS = "Documentos_Lojas"
PASTA_DOCS_FUNC = "Documentos_Funcionarios"
os.makedirs(PASTA_DOCS_LOJAS, exist_ok=True)
os.makedirs(PASTA_DOCS_FUNC, exist_ok=True)

MESES = ["Todos", "Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"]
MAP_MES = {m:i for i,m in enumerate(MESES) if m != "Todos"}

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
            "DataEvento","TipoEvento","Matricula","Nome","Situacao","Detalhes"
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
        st.error("❌ FECHE O ARQUIVO EXCEL E TENTE NOVAMENTE!")
        st.stop()
    except Exception as e:
        st.error(f"❌ Erro: {str(e)}")
        st.stop()

def lista_lojas():
    d = carregar_dados()
    return sorted(set(
        [str(l).strip() for l in d["Base_Dados"]["Loja"] if str(l).strip()] +
        [str(l).strip() for l in d["Auxiliares"]["Loja"] if str(l).strip()]
    )) or ["Sem Loja"]

def lista_cargos():
    d = carregar_dados()
    return sorted(set(
        [str(c).strip() for c in d["Base_Dados"]["Cargo"] if str(c).strip()] +
        [str(c).strip() for c in d["Auxiliares"]["Cargo"] if str(c).strip()]
    )) or ["Sem Cargo"]

def add_historico(mat, nome, acao, detalhes=""):
    dados = carregar_dados()
    novo = {
        "DataEvento": datetime.now().strftime("%d/%m/%Y"),
        "TipoEvento": acao, "Matricula": mat, "Nome": nome, "Detalhes": detalhes
    }
    dados["Historico"] = pd.concat([dados["Historico"], pd.DataFrame([novo])], ignore_index=True)
    salvar_dados(dados)

# ====================== INTERFACE ======================
st.set_page_config(page_title="SISTEMA RH", layout="wide", initial_sidebar_state="collapsed")
st.title("📋 SISTEMA RH COMPLETO")

aba1, aba2, aba3, aba4, aba5, aba6, aba7 = st.tabs([
    "Cadastro", "Painel", "Prazos e Férias", "Histórico", "Relatórios", "📎 Documentos", "⚙️ Lojas/Cargos"
])

# ================ ABA 1 - CADASTRO ================
with aba1:
    dados = carregar_dados()
    busca = st.text_input("🔍 Buscar Matrícula ou Nome")

    l1,l2,l3 = st.columns(3)
    fl = l1.selectbox("Loja", ["Todas"] + lista_lojas())
    fs = l2.selectbox("Situação", ["Todas"] + SITUACOES)
    fc = l3.selectbox("Cargo", ["Todos"] + lista_cargos())

    lista = dados["Base_Dados"].copy()
    if busca:
        lista = lista[(lista["Matricula"].str.contains(busca, case=False)) |
                      (lista["Nome"].str.contains(busca, case=False))]
    if fl != "Todas": lista = lista[lista["Loja"] == fl]
    if fs != "Todas": lista = lista[lista["Situacao"] == fs]
    if fc != "Todos": lista = lista[lista["Cargo"] == fc]
    st.dataframe(lista[["Matricula","Nome","Loja","Situacao","Cargo"]], use_container_width=True, hide_index=True)

    mat_sel = st.text_input("✏️ Digite Matrícula EXATA para editar")
    reg = dados["Base_Dados"][dados["Base_Dados"]["Matricula"] == mat_sel.strip()] if mat_sel else pd.DataFrame()
    val = lambda c: reg.iloc[0][c] if not reg.empty else ""

    with st.form("cad"):
        st.subheader("Dados Básicos")
        c1,c2,c3 = st.columns(3)
        with c1:
            mat = st.text_input("Matrícula *", value=val("Matricula"))
            nome = st.text_input("Nome", value=val("Nome"))
            cpf = st.text_input("CPF", value=val("CPF"))
            rg = st.text_input("RG", value=val("RG"))
            pis = st.text_input("PIS", value=val("PIS"))
        with c2:
            nasc = st.text_input("Nascimento (dd/mm/aaaa)", value=val("Nascimento"))
            adm = st.text_input("✅ Admissão (dd/mm/aaaa)", value=val("Admissao"))
            tel = st.text_input("Telefone", value=val("Telefone"))
            end = st.text_input("Endereço", value=val("Endereco"))
        with c3:
            loja = st.selectbox("Loja", lista_lojas(), index=lista_lojas().index(val("Loja")) if val("Loja") in lista_lojas() else 0)
            cargo = st.selectbox("Cargo", lista_cargos(), index=lista_cargos().index(val("Cargo")) if val("Cargo") in lista_cargos() else 0)
            sal = st.text_input("Salário", value=val("Salario"))
            sit = st.selectbox("Situação", SITUACOES, index=SITUACOES.index(val("Situacao")) if val("Situacao") in SITUACOES else 0)

        # 🧮 CÁLCULO INDIVIDUAL DO CONTRATO DE EXPERIÊNCIA
        st.markdown("---")
        st.subheader("📅 PRAZOS DE EXPERIÊNCIA (30 / 45 / 60 / 90 DIAS)")
        if adm.strip():
            try:
                dt_adm = datetime.strptime(adm.strip(), "%d/%m/%Y")
                hj = datetime.now()
                passados = (hj - dt_adm).days
                col = st.columns(4)
                for i, dias in enumerate([30,45,60,90]):
                    fim = dt_adm + timedelta(dias)
                    faltam = dias - passados
                    if faltam > 10: txt = f"🟢 {dias} dias\nTérmino: {fim:%d/%m/%Y}\nFaltam {faltam} dias"
                    elif faltam > 0: txt = f"🟡 {dias} dias\nTérmino: {fim:%d/%m/%Y}\nFaltam {faltam} dias"
                    elif faltam == 0: txt = f"🔴 {dias} dias\nHOJE!\nAtenção!"
                    else: txt = f"✅ {dias} dias\nConcluído em {fim:%d/%m/%Y}"
                    col[i].info(txt)
            except: st.warning("⚠️ Data inválida. Use dd/mm/aaaa")
        else: st.info("ℹ️ Preencha a Admissão para ver os cálculos")

        st.markdown("---")
        st.subheader("Outros Dados")
        a1,a2,a3 = st.columns(3)
        with a1:
            st.write("Aviso Prévio")
            dt_av = st.text_input("Data", value=val("DataAvisoPrevio"))
            d_av = st.text_input("Dias", value=val("DiasAvisoPrevio"))
        with a2:
            st.write("Férias")
            dt_fer = st.text_input("Início", value=val("DataFeriasInicio"))
            d_fer = st.text_input("Dias", value=val("DiasFerias"))
        with a3:
            st.write("Desligamento")
            dt_ped = st.text_input("Pedido Conta", value=val("DataPedidoConta"))
            dt_res = st.text_input("Rescisão", value=val("DataRescisao"))

        if st.form_submit_button("💾 SALVAR", type="primary", use_container_width=True):
            if not mat.strip(): st.error("❌ Matrícula obrigatória"); st.stop()
            reg_novo = {
                "Matricula":mat.strip(),"Nome":nome,"CPF":cpf,"RG":rg,"PIS":pis,
                "Nascimento":nasc,"Admissao":adm,"Telefone":tel,"Endereco":end,
                "Loja":loja,"Cargo":cargo,"Salario":sal,"Situacao":sit,
                "DataAvisoPrevio":dt_av,"DiasAvisoPrevio":d_av,
                "DataFeriasInicio":dt_fer,"DiasFerias":d_fer,
                "DataPedidoConta":dt_ped,"DataRescisao":dt_res
            }
            idx = dados["Base_Dados"].index[dados["Base_Dados"]["Matricula"]==mat.strip()].tolist()
            if idx: dados["Base_Dados"].iloc[idx[0]] = reg_novo; acao = "Atualização"
            else: dados["Base_Dados"] = pd.concat([dados["Base_Dados"], pd.DataFrame([reg_novo])], ignore_index=True); acao = "Cadastro Novo"
            salvar_dados(dados); add_historico(mat.strip(), nome, acao); st.success("✅ Salvo!"); st.rerun()

    if mat_sel and st.button("🗑️ EXCLUIR", use_container_width=True):
        if st.checkbox("Confirma exclusão?"):
            idx = dados["Base_Dados"].index[dados["Base_Dados"]["Matricula"]==mat_sel.strip()].tolist()
            if idx:
                docs = dados["Docs_Funcionarios"][dados["Docs_Funcionarios"]["Matricula"]==mat_sel.strip()]
                for _,d in docs.iterrows(): os.remove(d["Caminho"]) if os.path.exists(d["Caminho"]) else None
                dados["Docs_Funcionarios"] = dados["Docs_Funcionarios"][dados["Docs_Funcionarios"]["Matricula"]!=mat_sel.strip()]
                dados["Base_Dados"].drop(idx[0], inplace=True)
                salvar_dados(dados); add_historico(mat_sel.strip(), "", "Exclusão"); st.success("✅ Excluído"); st.rerun()

    # 📎 DOCUMENTOS DO FUNCIONÁRIO
    st.markdown("---")
    st.subheader("📎 ANEXAR DOCUMENTOS")
    if mat_sel and not reg.empty:
        tipo = st.selectbox("Tipo Doc", ["RG","CPF","CTPS","Contrato","Exame","Atestado","Outros"])
        arqs = st.file_uploader("Anexar", type=["pdf","doc","docx","jpg","png"], accept_multiple_files=True)
        if arqs and st.button("SALVAR DOCS"):
            for a in arqs:
                cam = os.path.join(PASTA_DOCS_FUNC, f"{mat_sel.strip()}_{tipo}_{datetime.now():%Y%m%d%H%M%S}_{a.name}")
                with open(cam,"wb") as f: f.write(a.read())
                dados["Docs_Funcionarios"] = pd.concat([dados["Docs_Funcionarios"], pd.DataFrame([{
                    "Matricula":mat_sel.strip(),"Nome":val("Nome"),"TipoDoc":tipo,
                    "NomeArquivo":a.name,"Caminho":cam,"DataAnexado":datetime.now().strftime("%d/%m/%Y %H:%M")
                }])], ignore_index=True)
            salvar_dados(dados); st.success("✅ Anexado"); st.rerun()
        for _,d in dados["Docs_Funcionarios"][dados["Docs_Funcionarios"]["Matricula"]==mat_sel.strip()].iterrows():
            with st.expander(f"📄 {d['TipoDoc']} - {d['NomeArquivo']}"):
                with open(d["Caminho"],"rb") as f: st.download_button("⬇️ BAIXAR", f, file_name=d["NomeArquivo"])
                if st.button("🗑️ EXCLUIR", key=f"d{_}"): os.remove(d["Caminho"]); dados["Docs_Funcionarios"].drop(_,inplace=True); salvar_dados(dados); st.rerun()

# ================ DEMAIS ABAS ================
with aba2:
    st.subheader("📊 RESUMO")
    d = dados["Base_Dados"]
    c1,c2,c3,c4,c5 = st.columns(5)
    c1.metric("Ativos", len(d[d["Situacao"]=="Ativo"]))
    c2.metric("Pré-cadastro", len(d[d["Situacao"]=="Pré-cadastro"]))
    c3.metric("Férias", len(d[d["Situacao"]=="Férias"]))
    c4.metric("Afastados", len(d[d["Situacao"].isin(["Doença","Acidente","Maternidade"])]))
    c5.metric("Desligados", len(d[~d["Situacao"].isin(["Ativo","Pré-cadastro","Férias","Doença","Acidente","Maternidade"])]))

with aba3:
    st.subheader("⚠️ PRAZOS PRÓXIMOS GERAL")
    hj = datetime.now()
    tab = []
    for _,f in dados["Base_Dados"].iterrows():
        if f["Situacao"] not in ["Ativo","Pré-cadastro"]: continue
        try:
            dt = datetime.strptime(f["Admissao"].strip(), "%d/%m/%Y")
            dias = (hj-dt).days
            for p in [30,45,60,90]:
                if 0 <= p-dias <=10: tab.append([f["Matricula"],f["Nome"],f["Loja"],f"{p} dias",f"Faltam {p-dias} dias"]); break
        except: pass
    st.dataframe(pd.DataFrame(tab, columns=["Matrícula","Nome","Loja","Prazo","Faltam"]), use_container_width=True, hide_index=True)

with aba4:
    st.subheader("📝 HISTÓRICO")
    st.dataframe(dados["Historico"], use_container_width=True, hide_index=True)

with aba5:
    st.subheader("📄 RELATÓRIOS")
    rel = st.selectbox("Escolha", ["Ativos","Férias","Afastados","Geral"])
    if st.button("GERAR"):
        df = dados["Base_Dados"] if rel=="Geral" else dados["Base_Dados"][dados["Base_Dados"]["Situacao"]==rel]
        with pd.ExcelWriter("rel.xlsx") as arq: df.to_excel(arq, index=False)
        with open("rel.xlsx","rb") as f: st.download_button("⬇️ BAIXAR", f, file_name=f"Rel_{rel}.xlsx")

with aba6:
    st.subheader("📎 DOCUMENTOS LOJAS")
    l,m,a = st.columns(3)
    sl = l.selectbox("Loja", lista_lojas())
    sm = m.selectbox("Mês", MESES)
    sa = a.selectbox("Ano", [str(x) for x in range(2020, datetime.now().year+2)])
    arq = st.file_uploader("Anexar", accept_multiple_files=True)
    if arq and st.button("SALVAR"):
        for f in arq:
            cam = os.path.join(PASTA_DOCS_LOJAS, f"{sl}_{sm}_{sa}_{datetime.now():%Y%m%d%H%M%S}_{f.name}")
            with open(cam,"wb") as fl: fl.write(f.read())
            dados["Docs_Lojas"] = pd.concat([dados["Docs_Lojas"], pd.DataFrame([{"Loja":sl,"Mes":sm,"Ano":sa,"NomeArquivo":f.name,"Caminho":cam,"DataAnexado":datetime.now().strftime("%d/%m/%Y")}])], ignore_index=True)
        salvar_dados(dados); st.success("✅ Salvo")
    for _,d in dados["Docs_Lojas"][(dados["Docs_Lojas"]["Loja"]==sl)&(dados["Docs_Lojas"]["Mes"]==sm)&(dados["Docs_Lojas"]["Ano"]==sa)].iterrows():
        with st.expander(f"📄 {d['NomeArquivo']}"):
            with open(d["Caminho"],"rb") as f: st.download_button("BAIXAR", f, file_name=d["NomeArquivo"])

with aba7:
    st.subheader("⚙️ CADASTRO DE LOJAS E CARGOS")
    c1,c2 = st.columns(2)
    with c1:
        nl = st.text_input("Nova Loja")
        if st.button("➕ ADD LOJA") and nl.strip() and nl.strip() not in lista_lojas():
            dados["Auxiliares"] = pd.concat([dados["Auxiliares"], pd.DataFrame([{"Loja":nl.strip(),"Cargo":""}])], ignore_index=True)
            salvar_dados(dados); st.rerun()
    with c2:
        nc = st.text_input("Novo Cargo")
        if st.button("➕ ADD CARGO") and nc.strip() and nc.strip() not in lista_cargos():
            dados["Auxiliares"] = pd.concat([dados["Auxiliares"], pd.DataFrame([{"Loja":"","Cargo":nc.strip()}])], ignore_index=True)
            salvar_dados(dados); st.rerun()
