"""
APP RH - Streamlit
Versão Corrigida
"""
import streamlit as st
import pandas as pd
import json
import os
from datetime import datetime, timedelta
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# ==================== CONFIGURAÇÃO ====================
st.set_page_config(page_title="RH App", layout="wide")

BASE_DIR = "dados_rh"
ARQUIVO_DADOS = os.path.join(BASE_DIR, "dados.json")

LOJAS = ["Loja 1", "Loja 2", "Loja 3", "Loja 4", "Loja 5", "Loja 6", "Loja 7", "Loja 8", "Loja 9", "Loja 10"]
CARGOS = ["Gerente", "Subgerente", "Vendedor", "Caixa", "Estoquista", "Auxiliar", "Administrativo", "Outro"]
SITUACOES = ["Ativo", "Desligado", "Férias", "Afastado", "Suspenso", "Licença"]
STATUS_LOJA = ["Ativa", "Inativa", "Fechada", "Em Reforma"]

# ==================== FUNÇÕES UTILITÁRIAS ====================

def safe_rerun():
    try:
        st.rerun()
    except Exception:
        pass

def gerar_id():
    return datetime.now().strftime("%Y%m%d%H%M%S%f")

def init_session():
    if "usuario" not in st.session_state:
        st.session_state["usuario"] = ""
    if "aba" not in st.session_state:
        st.session_state["aba"] = "1 - Cadastro"

init_session()

# ==================== CARREGAR / SALVAR DADOS ====================

@st.cache_data
def carregar_dados():
    os.makedirs(BASE_DIR, exist_ok=True)
    if not os.path.exists(ARQUIVO_DADOS):
        dados = {
            "Base_Dados": pd.DataFrame(columns=[
                "Matricula", "Nome", "CPF", "RG", "DataNascimento", "Sexo", "EstadoCivil", "Email", "Telefone", "Celular",
                "Endereco", "Numero", "Complemento", "Bairro", "Cidade", "Estado", "CEP", "Loja", "Cargo", "DataAdmissao",
                "DataDemissao", "Salario", "Banco", "Agencia", "Conta", "TipoConta", "PIX", "Situacao", "MotivoDemissao",
                "Observacoes", "DataCadastro", "DataUltimaAtualizacao", "UsuarioCadastro", "UsuarioAtualizacao", "HistoricoAcoes"
            ]),
            "Lojas": pd.DataFrame(columns=[
                "Loja", "Endereco", "Cidade", "Estado", "CEP", "Telefone", "Email", "Responsavel", "Status",
                "DataAbertura", "DataFechamento", "MotivoFechamento", "Tipo", "Metragem", "Aluguel", "Condominio",
                "IPTU", "Luz", "Agua", "Internet", "Seguro", "OutrosCustos", "TotalCustos", "FaturamentoMedio",
                "QtdFuncionarios", "QtdVendedores", "MetaMensal", "MetaSemanal", "MetaDiaria", "Observacoes",
                "DataCadastro", "DataUltimaAtualizacao", "UsuarioCadastro", "UsuarioAtualizacao", "HistoricoAcoes",
                "Latitude", "Longitude", "FotoLojaBase64"
            ]),
            "Documentos": pd.DataFrame(columns=[
                "ID", "Matricula", "Nome", "Tipo", "Descricao", "DataValidade", "ArquivoBase64", "DataCadastro",
                "UsuarioCadastro", "Observacoes"
            ]),
            "Historico": pd.DataFrame(columns=[
                "ID", "Matricula", "Nome", "Acao", "Data", "Usuario", "Detalhes", "Observacoes"
            ]),
            "FolhaPagamento": pd.DataFrame(columns=[
                "ID", "Matricula", "Nome", "Loja", "Cargo", "MesAno", "SalarioBase", "HorasExtras", "ValorHorasExtras",
                "Comissao", "Bonus", "Adiantamento", "Descontos", "INSS", "IRRF", "ValeTransporte", "ValeRefeicao",
                "SalarioFamilia", "Insalubridade", "Periculosidade", "OutrosVencimentos", "OutrosDescontos", "SalarioBruto",
                "SalarioLiquido", "DataPagamento", "Status", "Observacoes", "DataCadastro", "UsuarioCadastro"
            ]),
            "Ponto": pd.DataFrame(columns=[
                "ID", "Matricula", "Nome", "Loja", "Data", "Entrada", "Saida", "EntradaAlmoco", "SaidaAlmoco",
                "HorasNormais", "HorasExtras", "Atraso", "Falta", "Justificativa", "Observacoes", "Status", "DataCadastro", "UsuarioCadastro"
            ]),
            "Ferias": pd.DataFrame(columns=[
                "ID", "Matricula", "Nome", "Loja", "Cargo", "PeriodoAquisitivoInicio", "PeriodoAquisitivoFim",
                "DiasDireito", "DiasSolicitados", "DataInicio", "DataFim", "DataRetorno", "Status", "Observacoes",
                "DataCadastro", "UsuarioCadastro"
            ]),
            "Compras": pd.DataFrame(columns=[
                "ID", "Loja", "Tipo", "Descricao", "Fornecedor", "Valor", "Data", "Status", "Observacoes",
                "DataCadastro", "UsuarioCadastro"
            ]),
            "Configuracoes": {}
        }
        salvar_dados(dados)
        return dados
    try:
        with open(ARQUIVO_DADOS, "r", encoding="utf-8") as f:
            dados = json.load(f)
        for chave in ["Base_Dados", "Lojas", "Documentos", "Historico", "FolhaPagamento", "Ponto", "Ferias", "Compras"]:
            if chave in dados:
                dados[chave] = pd.DataFrame(dados[chave])
            else:
                dados[chave] = pd.DataFrame()
        return dados
    except Exception as e:
        st.error(f"Erro ao carregar dados: {e}")
        return {}

def salvar_dados(dados):
    try:
        os.makedirs(BASE_DIR, exist_ok=True)
        dados_salvar = {}
        for chave, valor in dados.items():
            if isinstance(valor, pd.DataFrame):
                dados_salvar[chave] = valor.to_dict(orient="records")
            else:
                dados_salvar[chave] = valor
        with open(ARQUIVO_DADOS, "w", encoding="utf-8") as f:
            json.dump(dados_salvar, f, ensure_ascii=False, indent=2, default=str)
        st.toast("Dados salvos com sucesso!", icon="✅")
    except Exception as e:
        st.error(f"Erro ao salvar dados: {e}")

def add_historico_auto(matricula, nome, acao, detalhes):
    st.cache_data.clear()
    dados = carregar_dados()
    if "Historico" not in dados:
        dados["Historico"] = pd.DataFrame()
    novo = pd.DataFrame([{
        "ID": gerar_id(),
        "Matricula": matricula,
        "Nome": nome,
        "Acao": acao,
        "Data": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        "Usuario": st.session_state.get("usuario", ""),
        "Detalhes": json.dumps(detalhes, ensure_ascii=False, default=str),
        "Observacoes": ""
    }])
    dados["Historico"] = pd.concat([dados["Historico"], novo], ignore_index=True)
    salvar_dados(dados)

def lista_lojas():
    dados = carregar_dados()
    if "Lojas" in dados and not dados["Lojas"].empty:
        return sorted(dados["Lojas"]["Loja"].fillna("").astype(str).str.strip().unique().tolist())
    return LOJAS


# ==================== INTERFACE ====================

st.title("📋 Sistema de RH")

aba = st.sidebar.radio("Menu", [
    "1 - Cadastro", "2 - Consulta / Edição", "3 - Relatórios", "4 - Lojas",
    "5 - Folha de Pagamento", "6 - Ponto", "7 - Férias", "8 - Documentos",
    "9 - Histórico", "10 - Configurações", "11 - Compras"
])


# ==================== ABA 1 - CADASTRO ====================

if aba == "1 - Cadastro":
    st.header("📝 Cadastro de Funcionários")
    dados = carregar_dados()

    st.markdown("### Filtro / Buscar")
    mat_filtro = st.text_input("Matrícula para buscar/editar", key="mat_filtro_cad")
    reg = pd.DataFrame()
    if mat_filtro.strip():
        reg = dados["Base_Dados"][dados["Base_Dados"]["Matricula"].astype(str).str.strip() == mat_filtro.strip()]

    def val_campo(nome):
        return reg.iloc[0][nome] if not reg.empty else ""

    with st.form("form_cadastro"):
        c1, c2, c3 = st.columns(3)
        with c1:
            matricula = st.text_input("Matrícula *", value=val_campo("Matricula"))
            nome = st.text_input("Nome *", value=val_campo("Nome"))
            cpf = st.text_input("CPF", value=val_campo("CPF"))
            rg = st.text_input("RG", value=val_campo("RG"))
            data_nasc = st.text_input("Data Nascimento", value=val_campo("DataNascimento"))
            sexo = st.selectbox("Sexo", ["", "Masculino", "Feminino", "Outro"], index=["", "Masculino", "Feminino", "Outro"].index(val_campo("Sexo")) if val_campo("Sexo") in ["", "Masculino", "Feminino", "Outro"] else 0)
            estado_civil = st.selectbox("Estado Civil", ["", "Solteiro", "Casado", "Divorciado", "Viúvo", "União Estável"], index=["", "Solteiro", "Casado", "Divorciado", "Viúvo", "União Estável"].index(val_campo("EstadoCivil")) if val_campo("EstadoCivil") in ["", "Solteiro", "Casado", "Divorciado", "Viúvo", "União Estável"] else 0)
        with c2:
            email = st.text_input("Email", value=val_campo("Email"))
            telefone = st.text_input("Telefone", value=val_campo("Telefone"))
            celular = st.text_input("Celular", value=val_campo("Celular"))
            endereco = st.text_input("Endereço", value=val_campo("Endereco"))
            numero = st.text_input("Número", value=val_campo("Numero"))
            complemento = st.text_input("Complemento", value=val_campo("Complemento"))
            bairro = st.text_input("Bairro", value=val_campo("Bairro"))
        with c3:
            cidade = st.text_input("Cidade", value=val_campo("Cidade"))
            estado = st.text_input("Estado", value=val_campo("Estado"))
            cep = st.text_input("CEP", value=val_campo("CEP"))
            loja = st.selectbox("Loja *", [""] + lista_lojas(), index=([""] + lista_lojas()).index(val_campo("Loja")) if val_campo("Loja") in ([""] + lista_lojas()) else 0)
            cargo = st.selectbox("Cargo *", [""] + CARGOS, index=([""] + CARGOS).index(val_campo("Cargo")) if val_campo("Cargo") in ([""] + CARGOS) else 0)
            data_adm = st.text_input("Data Admissão", value=val_campo("DataAdmissao"))
            data_dem = st.text_input("Data Demissão", value=val_campo("DataDemissao"))

        st.markdown("---")
        c1, c2, c3 = st.columns(3)
        with c1:
            salario = st.text_input("Salário", value=val_campo("Salario"))
            banco = st.text_input("Banco", value=val_campo("Banco"))
            agencia = st.text_input("Agência", value=val_campo("Agencia"))
            conta = st.text_input("Conta", value=val_campo("Conta"))
        with c2:
            tipo_conta = st.selectbox("Tipo Conta", ["", "Corrente", "Poupança", "Salário"], index=["", "Corrente", "Poupança", "Salário"].index(val_campo("TipoConta")) if val_campo("TipoConta") in ["", "Corrente", "Poupança", "Salário"] else 0)
            pix = st.text_input("PIX", value=val_campo("PIX"))
            situacao = st.selectbox("Situação *", [""] + SITUACOES, index=([""] + SITUACOES).index(val_campo("Situacao")) if val_campo("Situacao") in ([""] + SITUACOES) else 0)
            motivo_dem = st.text_input("Motivo Demissão", value=val_campo("MotivoDemissao"))
        with c3:
            observacoes = st.text_area("Observações", value=val_campo("Observacoes"), height=150)

        btn_salvar = st.form_submit_button("💾 SALVAR", type="primary", use_container_width=True)

        if btn_salvar:
            if not matricula.strip() or not nome.strip() or not loja or not cargo or not situacao:
                st.error("❌ Preencha os campos obrigatórios (*)")
                st.stop()

            registro = {
                "Matricula": matricula.strip(), "Nome": nome.strip(), "CPF": cpf, "RG": rg,
                "DataNascimento": data_nasc, "Sexo": sexo, "EstadoCivil": estado_civil, "Email": email,
                "Telefone": telefone, "Celular": celular, "Endereco": endereco, "Numero": numero,
                "Complemento": complemento, "Bairro": bairro, "Cidade": cidade, "Estado": estado, "CEP": cep,
                "Loja": loja.strip(), "Cargo": cargo, "DataAdmissao": data_adm, "DataDemissao": data_dem,
                "Salario": salario, "Banco": banco, "Agencia": agencia, "Conta": conta, "TipoConta": tipo_conta,
                "PIX": pix, "Situacao": situacao, "MotivoDemissao": motivo_dem, "Observacoes": observacoes,
                "DataCadastro": datetime.now().strftime("%d/%m/%Y %H:%M:%S") if reg.empty else val_campo("DataCadastro"),
                "DataUltimaAtualizacao": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
                "UsuarioCadastro": st.session_state.get("usuario", ""),
                "UsuarioAtualizacao": st.session_state.get("usuario", ""),
                "HistoricoAcoes": ""
            }

            idx = dados["Base_Dados"].index[dados["Base_Dados"]["Matricula"].astype(str).str.strip() == matricula.strip()].tolist()
            if idx:
                for col, val in registro.items():
                    dados["Base_Dados"].at[idx[0], col] = val
                msg = f"✅ Funcionário {matricula.strip()} atualizado!"
                acao_hist = "EDIÇÃO"
            else:
                dados["Base_Dados"] = pd.concat([dados["Base_Dados"], pd.DataFrame([registro])], ignore_index=True)
                msg = f"✅ Funcionário {matricula.strip()} cadastrado!"
                acao_hist = "CADASTRO"

            salvar_dados(dados)
            add_historico_auto(matricula.strip(), nome.strip(), acao_hist, registro)
            st.success(msg)
            safe_rerun()

    st.markdown("---")
    st.markdown("### 📋 Funcionários Cadastrados")
    base = dados["Base_Dados"].copy()
    base["Matricula"] = base["Matricula"].fillna("").astype(str).str.strip()
    base["Nome"] = base["Nome"].fillna("").astype(str).str.strip()
    base["Loja"] = base["Loja"].fillna("").astype(str).str.strip()
    base["Situacao"] = base["Situacao"].fillna("").astype(str).str.strip()
    base["Cargo"] = base["Cargo"].fillna("").astype(str).str.strip()
    st.dataframe(base[["Matricula","Nome","Loja","Cargo","Situacao","DataAdmissao"]], use_container_width=True, hide_index=True)


# ==================== ABA 2 - CONSULTA / EDIÇÃO ====================

if aba == "2 - Consulta / Edição":
    st.header("🔍 Consulta e Edição de Funcionários")
    dados = carregar_dados()
    base = dados["Base_Dados"].copy()
    base["Matricula"] = base["Matricula"].fillna("").astype(str).str.strip()
    base["Nome"] = base["Nome"].fillna("").astype(str).str.strip()
    base["Loja"] = base["Loja"].fillna("").astype(str).str.strip()
    base["Situacao"] = base["Situacao"].fillna("").astype(str).str.strip()
    base["Cargo"] = base["Cargo"].fillna("").astype(str).str.strip()

    st.markdown("### 🔍 Filtros")
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        filtro_nome = st.text_input("Nome", key="filtro_nome_cons")
    with c2:
        filtro_mat = st.text_input("Matrícula", key="filtro_mat_cons")
    with c3:
        filtro_loja = st.selectbox("Loja", ["Todas"] + lista_lojas(), key="filtro_loja_cons")
    with c4:
        filtro_sit = st.selectbox("Situação", ["Todas"] + SITUACOES, key="filtro_sit_cons")

    resultado = base.copy()
    if filtro_nome:
        resultado = resultado[resultado["Nome"].str.contains(filtro_nome, case=False, na=False)]
    if filtro_mat:
        resultado = resultado[resultado["Matricula"] == filtro_mat.strip()]
    if filtro_loja != "Todas":
        resultado = resultado[resultado["Loja"] == filtro_loja.strip()]
    if filtro_sit != "Todas":
        resultado = resultado[resultado["Situacao"] == filtro_sit]

    st.markdown(f"**Resultados:** {len(resultado)}")
    st.dataframe(resultado[["Matricula","Nome","Loja","Cargo","Situacao","DataAdmissao","DataUltimaAtualizacao"]], use_container_width=True, hide_index=True)

    st.markdown("---")
    st.markdown("### ✏️ Edição Rápida")
    mat_edicao = st.selectbox("Selecionar funcionário", [""] + resultado["Matricula"].tolist(), key="mat_edicao")
    if mat_edicao:
        reg_ed = dados["Base_Dados"][dados["Base_Dados"]["Matricula"].astype(str).str.strip() == mat_edicao]
        if not reg_ed.empty:
            st.markdown(f"**Editando:** {reg_ed.iloc[0]['Nome']} (Mat: {mat_edicao})")
            with st.form("form_edicao_rapida"):
                novo_nome = st.text_input("Nome", value=reg_ed.iloc[0]["Nome"])
                novo_cargo = st.selectbox("Cargo", CARGOS, index=CARGOS.index(reg_ed.iloc[0]["Cargo"]) if reg_ed.iloc[0]["Cargo"] in CARGOS else 0)
                novo_salario = st.text_input("Salário", value=str(reg_ed.iloc[0]["Salario"]))
                nova_situacao = st.selectbox("Situação", SITUACOES, index=SITUACOES.index(reg_ed.iloc[0]["Situacao"]) if reg_ed.iloc[0]["Situacao"] in SITUACOES else 0)
                nova_loja = st.selectbox("Loja", lista_lojas(), index=lista_lojas().index(reg_ed.iloc[0]["Loja"]) if reg_ed.iloc[0]["Loja"] in lista_lojas() else 0)
                btn_atualizar = st.form_submit_button("💾 Atualizar", use_container_width=True)
                if btn_atualizar:
                    idx = dados["Base_Dados"].index[dados["Base_Dados"]["Matricula"].astype(str).str.strip() == mat_edicao].tolist()[0]
                    dados["Base_Dados"].at[idx, "Nome"] = novo_nome
                    dados["Base_Dados"].at[idx, "Cargo"] = novo_cargo
                    dados["Base_Dados"].at[idx, "Salario"] = novo_salario
                    dados["Base_Dados"].at[idx, "Situacao"] = nova_situacao
                    dados["Base_Dados"].at[idx, "Loja"] = nova_loja
                    dados["Base_Dados"].at[idx, "DataUltimaAtualizacao"] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                    dados["Base_Dados"].at[idx, "UsuarioAtualizacao"] = st.session_state.get("usuario", "")
                    salvar_dados(dados)
                    add_historico_auto(mat_edicao, novo_nome, "EDIÇÃO RÁPIDA", {})
                    st.success("✅ Atualizado!")
                    safe_rerun()
        else:
            st.warning("Funcionário não encontrado.")

    st.markdown("---")
    st.markdown("### 🗑️ Exclusão")
    mat_excluir = st.selectbox("Selecionar para exclusão", [""] + resultado["Matricula"].tolist(), key="mat_excluir")
    if mat_excluir and st.button("🗑️ EXCLUIR FUNCIONÁRIO", type="primary"):
        idx = dados["Base_Dados"].index[dados["Base_Dados"]["Matricula"].astype(str).str.strip() == mat_excluir].tolist()
        if idx:
            nome_excl = dados["Base_Dados"].at[idx[0], "Nome"]
            dados["Base_Dados"] = dados["Base_Dados"].drop(idx[0]).reset_index(drop=True)
            salvar_dados(dados)
            add_historico_auto(mat_excluir, nome_excl, "EXCLUSÃO", {})
            st.success(f"🗑️ Funcionário {mat_excluir} excluído!")
            safe_rerun()


# ==================== ABA 3 - RELATÓRIOS ====================

if aba == "3 - Relatórios":
    st.header("📊 Relatórios")
    dados = carregar_dados()
    base = dados["Base_Dados"].copy()
    base["Matricula"] = base["Matricula"].fillna("").astype(str).str.strip()
    base["Nome"] = base["Nome"].fillna("").astype(str).str.strip()
    base["Loja"] = base["Loja"].fillna("").astype(str).str.strip()
    base["Situacao"] = base["Situacao"].fillna("").astype(str).str.strip()
    base["Cargo"] = base["Cargo"].fillna("").astype(str).str.strip()

    st.markdown("### 📈 Resumo Geral")
    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        st.metric("Total", len(base))
    with c2:
        st.metric("Ativos", len(base[base["Situacao"] == "Ativo"]))
    with c3:
        st.metric("Desligados", len(base[base["Situacao"] == "Desligado"]))
    with c4:
        st.metric("Férias", len(base[base["Situacao"] == "Férias"]))
    with c5:
        st.metric("Afastados", len(base[base["Situacao"] == "Afastado"]))

    st.markdown("---")
    st.markdown("### 📋 Por Loja")
    rel_loja = base.groupby("Loja").agg({"Matricula": "count", "Situacao": lambda x: (x == "Ativo").sum()}).rename(columns={"Matricula": "Total", "Situacao": "Ativos"}).reset_index()
    st.dataframe(rel_loja, use_container_width=True, hide_index=True)

    st.markdown("---")
    st.markdown("### 📋 Por Cargo")
    rel_cargo = base.groupby("Cargo").agg({"Matricula": "count", "Situacao": lambda x: (x == "Ativo").sum()}).rename(columns={"Matricula": "Total", "Situacao": "Ativos"}).reset_index()
    st.dataframe(rel_cargo, use_container_width=True, hide_index=True)

    st.markdown("---")
    st.markdown("### 📋 Por Situação")
    rel_sit = base.groupby("Situacao").agg({"Matricula": "count"}).rename(columns={"Matricula": "Total"}).reset_index()
    st.dataframe(rel_sit, use_container_width=True, hide_index=True)

    st.markdown("---")
    st.markdown("### 📥 Exportar")
    if st.button("📥 Exportar Excel", use_container_width=True):
        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "Relatório Geral"
            for r in dataframe_to_rows(base, index=False, header=True):
                ws.append(r)
            excel_buffer = BytesIO()
            wb.save(excel_buffer)
            excel_buffer.seek(0)
            st.download_button(label="⬇️ Baixar", data=excel_buffer, file_name=f"relatorio_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"Erro ao exportar: {e}")


# ==================== ABA 4 - LOJAS ====================

if aba == "4 - Lojas":
    st.header("🏪 Cadastro de Lojas")
    dados = carregar_dados()
    lojas_df = dados["Lojas"].copy()
    lojas_df["Loja"] = lojas_df["Loja"].fillna("").astype(str).str.strip()

    st.markdown("### 📋 Lojas Cadastradas")
    st.dataframe(lojas_df[["Loja","Cidade","Estado","Status","Responsavel"]], use_container_width=True, hide_index=True)

    st.markdown("---")
    st.markdown("### 📝 Cadastrar / Editar Loja")
    loja_sel = st.selectbox("Selecionar", ["NOVA"] + lojas_df["Loja"].tolist(), key="loja_sel")
    reg_loja = lojas_df[lojas_df["Loja"] == loja_sel] if loja_sel != "NOVA" else pd.DataFrame()

    def val_loja(nome):
        return reg_loja.iloc[0][nome] if not reg_loja.empty else ""

    with st.form("form_loja"):
        c1, c2, c3 = st.columns(3)
        with c1:
            cod_loja = st.text_input("Código", value=val_loja("Loja"))
            endereco = st.text_input("Endereço", value=val_loja("Endereco"))
            cidade = st.text_input("Cidade", value=val_loja("Cidade"))
            estado = st.text_input("Estado", value=val_loja("Estado"))
        with c2:
            cep = st.text_input("CEP", value=val_loja("CEP"))
            telefone = st.text_input("Telefone", value=val_loja("Telefone"))
            email = st.text_input("Email", value=val_loja("Email"))
            responsavel = st.text_input("Responsável", value=val_loja("Responsavel"))
        with c3:
            status = st.selectbox("Status", STATUS_LOJA, index=STATUS_LOJA.index(val_loja("Status")) if val_loja("Status") in STATUS_LOJA else 0)
            data_abertura = st.text_input("Data Abertura", value=val_loja("DataAbertura"))
            data_fechamento = st.text_input("Data Fechamento", value=val_loja("DataFechamento"))
            motivo_fechamento = st.text_input("Motivo Fechamento", value=val_loja("MotivoFechamento"))

        st.markdown("---")
        c1, c2, c3 = st.columns(3)
        with c1:
            tipo = st.text_input("Tipo", value=val_loja("Tipo"))
            metragem = st.text_input("Metragem", value=val_loja("Metragem"))
            aluguel = st.text_input("Aluguel", value=val_loja("Aluguel"))
            condominio = st.text_input("Condomínio", value=val_loja("Condominio"))
        with c2:
            iptu = st.text_input("IPTU", value=val_loja("IPTU"))
            luz = st.text_input("Luz", value=val_loja("Luz"))
            agua = st.text_input("Água", value=val_loja("Agua"))
            internet = st.text_input("Internet", value=val_loja("Internet"))
        with c3:
            seguro = st.text_input("Seguro", value=val_loja("Seguro"))
            outros_custos = st.text_input("Outros Custos", value=val_loja("OutrosCustos"))
            total_custos = st.text_input("Total Custos", value=val_loja("TotalCustos"))
            faturamento = st.text_input("Faturamento Médio", value=val_loja("FaturamentoMedio"))

        st.markdown("---")
        c1, c2, c3 = st.columns(3)
        with c1:
            qtd_func = st.text_input("Qtd Funcionários", value=val_loja("QtdFuncionarios"))
            qtd_vend = st.text_input("Qtd Vendedores", value=val_loja("QtdVendedores"))
        with c2:
            meta_mensal = st.text_input("Meta Mensal", value=val_loja("MetaMensal"))
            meta_semanal = st.text_input("Meta Semanal", value=val_loja("MetaSemanal"))
            meta_diaria = st.text_input("Meta Diária", value=val_loja("MetaDiaria"))
        with c3:
            observacoes_loja = st.text_area("Observações", value=val_loja("Observacoes"), height=100)

        btn_salvar_loja = st.form_submit_button("💾 SALVAR LOJA", type="primary", use_container_width=True)
        if btn_salvar_loja:
            if not cod_loja.strip():
                st.error("❌ INFORME O CÓDIGO DA LOJA!")
                st.stop()
            registro_loja = {
                "Loja": cod_loja.strip(), "Endereco": endereco, "Cidade": cidade, "Estado": estado, "CEP": cep,
                "Telefone": telefone, "Email": email, "Responsavel": responsavel, "Status": status,
                "DataAbertura": data_abertura, "DataFechamento": data_fechamento, "MotivoFechamento": motivo_fechamento,
                "Tipo": tipo, "Metragem": metragem, "Aluguel": aluguel, "Condominio": condominio,
                "IPTU": iptu, "Luz": luz, "Agua": agua, "Internet": internet, "Seguro": seguro,
                "OutrosCustos": outros_custos, "TotalCustos": total_custos, "FaturamentoMedio": faturamento,
                "QtdFuncionarios": qtd_func, "QtdVendedores": qtd_vend, "MetaMensal": meta_mensal,
                "MetaSemanal": meta_semanal, "MetaDiaria": meta_diaria, "Observacoes": observacoes_loja,
                "DataCadastro": datetime.now().strftime("%d/%m/%Y %H:%M:%S") if reg_loja.empty else val_loja("DataCadastro"),
                "DataUltimaAtualizacao": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
                "UsuarioCadastro": st.session_state.get("usuario", ""),
                "UsuarioAtualizacao": st.session_state.get("usuario", ""),
                "HistoricoAcoes": "", "Latitude": "", "Longitude": "", "FotoLojaBase64": ""
            }
            idx = lojas_df.index[lojas_df["Loja"] == cod_loja.strip()].tolist()
            if idx:
                for col, val in registro_loja.items():
                    dados["Lojas"].at[idx[0], col] = val
            else:
                dados["Lojas"] = pd.concat([dados["Lojas"], pd.DataFrame([registro_loja])], ignore_index=True)
            salvar_dados(dados)
            st.success(f"✅ Loja {cod_loja.strip()} salva!")
            safe_rerun()


# ==================== ABA 5 - FOLHA DE PAGAMENTO ====================

if aba == "5 - Folha de Pagamento":
    st.header("💰 Folha de Pagamento")
    dados = carregar_dados()
    base = dados["Base_Dados"].copy()
    base["Matricula"] = base["Matricula"].fillna("").astype(str).str.strip()
    base["Nome"] = base["Nome"].fillna("").astype(str).str.strip()

    st.markdown("### 📝 Lançar Folha")
    mat_folha = st.selectbox("Funcionário", [""] + base["Matricula"].tolist(), key="mat_folha")
    if mat_folha:
        func = base[base["Matricula"] == mat_folha].iloc[0]
        st.markdown(f"**{func['Nome']}** | {func['Cargo']} | {func['Loja']}")
        with st.form("form_folha"):
            mes_ano = st.text_input("Mês/Ano (MM/YYYY)", key="mes_ano")
            salario_base = st.text_input("Salário Base", value=str(func["Salario"]), key="sal_base")
            horas_extras = st.text_input("Horas Extras", value="0", key="he")
            val_he = st.text_input("Valor HE", value="0", key="val_he")
            comissao = st.text_input("Comissão", value="0", key="comissao")
            bonus = st.text_input("Bônus", value="0", key="bonus")
            adiantamento = st.text_input("Adiantamento", value="0", key="adiant")
            descontos = st.text_input("Descontos", value="0", key="desc")
            inss = st.text_input("INSS", value="0", key="inss")
            irrf = st.text_input("IRRF", value="0", key="irrf")
            vt = st.text_input("Vale Transporte", value="0", key="vt")
            vr = st.text_input("Vale Refeição", value="0", key="vr")
            outros_venc = st.text_input("Outros Vencimentos", value="0", key="out_venc")
            outros_desc = st.text_input("Outros Descontos", value="0", key="out_desc")
            obs_folha = st.text_area("Observações", key="obs_folha")
            btn_lancar = st.form_submit_button("💾 LANÇAR", type="primary", use_container_width=True)
            if btn_lancar:
                try:
                    bruto = float(salario_base or 0) + float(val_he or 0) + float(comissao or 0) + float(bonus or 0) + float(outros_venc or 0)
                    liq = bruto - float(adiantamento or 0) - float(descontos or 0) - float(inss or 0) - float(irrf or 0) - float(vt or 0) - float(vr or 0) - float(outros_desc or 0)
                except:
                    bruto, liq = 0, 0
                novo = pd.DataFrame([{
                    "ID": gerar_id(), "Matricula": mat_folha, "Nome": func["Nome"], "Loja": func["Loja"], "Cargo": func["Cargo"],
                    "MesAno": mes_ano, "SalarioBase": salario_base, "HorasExtras": horas_extras, "ValorHorasExtras": val_he,
                    "Comissao": comissao, "Bonus": bonus, "Adiantamento": adiantamento, "Descontos": descontos,
                    "INSS": inss, "IRRF": irrf, "ValeTransporte": vt, "ValeRefeicao": vr,
                    "OutrosVencimentos": outros_venc, "OutrosDescontos": outros_desc, "SalarioBruto": bruto,
                    "SalarioLiquido": liq, "DataPagamento": "", "Status": "Lançado", "Observacoes": obs_folha,
                    "DataCadastro": datetime.now().strftime("%d/%m/%Y %H:%M:%S"), "UsuarioCadastro": st.session_state.get("usuario", "")
                }])
                dados["FolhaPagamento"] = pd.concat([dados["FolhaPagamento"], novo], ignore_index=True)
                salvar_dados(dados)
                add_historico_auto(mat_folha, func["Nome"], "FOLHA PAGAMENTO", {"MesAno": mes_ano})
                st.success("✅ Folha lançada!")
                safe_rerun()

    st.markdown("---")
    st.markdown("### 📋 Folhas Lançadas")
    folha = dados["FolhaPagamento"].copy()
    if not folha.empty:
        folha["Matricula"] = folha["Matricula"].fillna("").astype(str).str.strip()
        st.dataframe(folha[["Matricula","Nome","Loja","MesAno","SalarioBruto","SalarioLiquido","Status"]], use_container_width=True, hide_index=True)
    else:
        st.info("Nenhuma folha lançada.")


# ==================== ABA 6 - PONTO ====================

if aba == "6 - Ponto":
    st.header("⏰ Controle de Ponto")
    dados = carregar_dados()
    base = dados["Base_Dados"].copy()
    base["Matricula"] = base["Matricula"].fillna("").astype(str).str.strip()
    base["Nome"] = base["Nome"].fillna("").astype(str).str.strip()

    st.markdown("### 📝 Lançar Ponto")
    mat_ponto = st.selectbox("Funcionário", [""] + base["Matricula"].tolist(), key="mat_ponto")
    if mat_ponto:
        func = base[base["Matricula"] == mat_ponto].iloc[0]
        st.markdown(f"**{func['Nome']}** | {func['Cargo']} | {func['Loja']}")
        with st.form("form_ponto"):
            data_ponto = st.date_input("Data", key="data_ponto")
            entrada = st.time_input("Entrada", key="entrada")
            saida = st.time_input("Saída", key="saida")
            entrada_alm = st.time_input("Entrada Almoço", key="entr_alm")
            saida_alm = st.time_input("Saída Almoço", key="sai_alm")
            horas_norm = st.text_input("Horas Normais", value="8", key="hn")
            horas_ext = st.text_input("Horas Extras", value="0", key="he_ponto")
            atraso = st.text_input("Atraso", value="0", key="atraso")
            falta = st.selectbox("Falta", ["Não", "Sim"], key="falta")
            justificativa = st.text_input("Justificativa", key="just")
            obs_ponto = st.text_area("Observações", key="obs_ponto")
            btn_lancar_ponto = st.form_submit_button("💾 LANÇAR PONTO", type="primary", use_container_width=True)
            if btn_lancar_ponto:
                novo = pd.DataFrame([{
                    "ID": gerar_id(), "Matricula": mat_ponto, "Nome": func["Nome"], "Loja": func["Loja"],
                    "Data": data_ponto.strftime("%d/%m/%Y"), "Entrada": entrada.strftime("%H:%M"), "Saida": saida.strftime("%H:%M"),
                    "EntradaAlmoco": entrada_alm.strftime("%H:%M"), "SaidaAlmoco": saida_alm.strftime("%H:%M"),
                    "HorasNormais": horas_norm, "HorasExtras": horas_ext, "Atraso": atraso, "Falta": falta,
                    "Justificativa": justificativa, "Observacoes": obs_ponto, "Status": "Lançado",
                    "DataCadastro": datetime.now().strftime("%d/%m/%Y %H:%M:%S"), "UsuarioCadastro": st.session_state.get("usuario", "")
                }])
                dados["Ponto"] = pd.concat([dados["Ponto"], novo], ignore_index=True)
                salvar_dados(dados)
                add_historico_auto(mat_ponto, func["Nome"], "PONTO", {"Data": data_ponto.strftime("%d/%m/%Y")})
                st.success("✅ Ponto lançado!")
                safe_rerun()

    st.markdown("---")
    st.markdown("### 📋 Registros de Ponto")
    ponto = dados["Ponto"].copy()
    if not ponto.empty:
        ponto["Matricula"] = ponto["Matricula"].fillna("").astype(str).str.strip()
        st.dataframe(ponto[["Matricula","Nome","Loja","Data","Entrada","Saida","HorasNormais","HorasExtras","Falta"]], use_container_width=True, hide_index=True)
    else:
        st.info("Nenhum ponto lançado.")


# ==================== ABA 7 - FÉRIAS ====================

if aba == "7 - Férias":
    st.header("🏖️ Controle de Férias")
    dados = carregar_dados()
    base = dados["Base_Dados"].copy()
    base["Matricula"] = base["Matricula"].fillna("").astype(str).str.strip()
    base["Nome"] = base["Nome"].fillna("").astype(str).str.strip()

    st.markdown("### 📝 Lançar Férias")
    mat_ferias = st.selectbox("Funcionário", [""] + base["Matricula"].tolist(), key="mat_ferias")
    if mat_ferias:
        func = base[base["Matricula"] == mat_ferias].iloc[0]
        st.markdown(f"**{func['Nome']}** | {func['Cargo']} | {func['Loja']}")
        with st.form("form_ferias"):
            per_ini = st.text_input("Período Aquisitivo Início", key="per_ini")
            per_fim = st.text_input("Período Aquisitivo Fim", key="per_fim")
            dias_direito = st.text_input("Dias Direito", value="30", key="dias_dir")
            dias_solic = st.text_input("Dias Solicitados", value="30", key="dias_sol")
            data_ini = st.text_input("Data Início", key="data_ini_fer")
            data_fim = st.text_input("Data Fim", key="data_fim_fer")
            data_ret = st.text_input("Data Retorno", key="data_ret")
            status_ferias = st.selectbox("Status", ["Solicitado", "Aprovado", "Em Andamento", "Concluído", "Cancelado"], key="status_fer")
            obs_ferias = st.text_area("Observações", key="obs_ferias")
            btn_lancar_ferias = st.form_submit_button("💾 LANÇAR FÉRIAS", type="primary", use_container_width=True)
            if btn_lancar_ferias:
                novo = pd.DataFrame([{
                    "ID": gerar_id(), "Matricula": mat_ferias, "Nome": func["Nome"], "Loja": func["Loja"], "Cargo": func["Cargo"],
                    "PeriodoAquisitivoInicio": per_ini, "PeriodoAquisitivoFim": per_fim, "DiasDireito": dias_direito,
                    "DiasSolicitados": dias_solic, "DataInicio": data_ini, "DataFim": data_fim, "DataRetorno": data_ret,
                    "Status": status_ferias, "Observacoes": obs_ferias,
                    "DataCadastro": datetime.now().strftime("%d/%m/%Y %H:%M:%S"), "UsuarioCadastro": st.session_state.get("usuario", "")
                }])
                dados["Ferias"] = pd.concat([dados["Ferias"], novo], ignore_index=True)
                salvar_dados(dados)
                add_historico_auto(mat_ferias, func["Nome"], "FÉRIAS", {"Status": status_ferias})
                st.success("✅ Férias lançadas!")
                safe_rerun()

    st.markdown("---")
    st.markdown("### 📋 Férias Lançadas")
    ferias = dados["Ferias"].copy()
    if not ferias.empty:
        ferias["Matricula"] = ferias["Matricula"].fillna("").astype(str).str.strip()
        st.dataframe(ferias[["Matricula","Nome","Loja","DataInicio","DataFim","DiasSolicitados","Status"]], use_container_width=True, hide_index=True)
    else:
        st.info("Nenhuma férias lançada.")


# ==================== ABA 8 - DOCUMENTOS ====================

if aba == "8 - Documentos":
    st.header("📁 Documentos")
    dados = carregar_dados()
    base = dados["Base_Dados"].copy()
    base["Matricula"] = base["Matricula"].fillna("").astype(str).str.strip()
    base["Nome"] = base["Nome"].fillna("").astype(str).str.strip()

    st.markdown("### 📝 Anexar Documento")
    mat_doc = st.selectbox("Funcionário", [""] + base["Matricula"].tolist(), key="mat_doc")
    if mat_doc:
        func = base[base["Matricula"] == mat_doc].iloc[0]
        st.markdown(f"**{func['Nome']}**")
        with st.form("form_doc"):
            tipo_doc = st.selectbox("Tipo", ["RG", "CPF", "CNH", "CTPS", "Título Eleitor", "Certidão", "Comprovante Residência", "Outro"], key="tipo_doc")
            descricao = st.text_input("Descrição", key="desc_doc")
            data_validade = st.text_input("Data Validade", key="val_doc")
            obs_doc = st.text_area("Observações", key="obs_doc")
            btn_anexar = st.form_submit_button("💾 ANEXAR", type="primary", use_container_width=True)
            if btn_anexar:
                novo = pd.DataFrame([{
                    "ID": gerar_id(), "Matricula": mat_doc, "Nome": func["Nome"], "Tipo": tipo_doc,
                    "Descricao": descricao, "DataValidade": data_validade, "ArquivoBase64": "",
                    "DataCadastro": datetime.now().strftime("%d/%m/%Y %H:%M:%S"), "UsuarioCadastro": st.session_state.get("usuario", ""),
                    "Observacoes": obs_doc
                }])
                dados["Documentos"] = pd.concat([dados["Documentos"], novo], ignore_index=True)
                salvar_dados(dados)
                add_historico_auto(mat_doc, func["Nome"], "DOCUMENTO", {"Tipo": tipo_doc})
                st.success("✅ Documento anexado!")
                safe_rerun()

    st.markdown("---")
    st.markdown("### 📋 Documentos Cadastrados")
    docs = dados["Documentos"].copy()
    if not docs.empty:
        docs["Matricula"] = docs["Matricula"].fillna("").astype(str).str.strip()
        st.dataframe(docs[["Matricula","Nome","Tipo","Descricao","DataValidade"]], use_container_width=True, hide_index=True)
    else:
        st.info("Nenhum documento cadastrado.")


# ==================== ABA 9 - HISTÓRICO ====================

if aba == "9 - Histórico":
    st.header("📜 Histórico de Ações")
    dados = carregar_dados()
    hist = dados["Historico"].copy()
    if not hist.empty:
        hist["Matricula"] = hist["Matricula"].fillna("").astype(str).str.strip()
        hist = hist.sort_values("Data", ascending=False)
        st.dataframe(hist[["Data","Matricula","Nome","Acao","Usuario","Detalhes"]], use_container_width=True, hide_index=True)
    else:
        st.info("Nenhum histórico registrado.")


# ==================== ABA 10 - CONFIGURAÇÕES ====================

if aba == "10 - Configurações":
    st.header("⚙️ Configurações")
    dados = carregar_dados()

    st.markdown("### 👤 Usuário Atual")
    usuario_atual = st.text_input("Nome do Usuário", value=st.session_state.get("usuario", ""), key="usuario_conf")
    if st.button("💾 Salvar Usuário", use_container_width=True):
        st.session_state["usuario"] = usuario_atual
        st.success("✅ Usuário atualizado!")
        safe_rerun()

    st.markdown("---")
    st.markdown("### 💾 Backup / Restore")
    if st.button("📥 Fazer Backup", use_container_width=True):
        try:
            with open(ARQUIVO_DADOS, "r", encoding="utf-8") as f:
                conteudo = f.read()
            st.download_button(label="⬇️ Baixar Backup", data=conteudo, file_name=f"backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json", mime="application/json")
        except Exception as e:
            st.error(f"Erro backup: {e}")

    arquivo_upload = st.file_uploader("Restaurar Backup", type="json", key="upload_backup")
    if arquivo_upload is not None:
        if st.button("🔄 RESTAURAR", type="primary", use_container_width=True):
            try:
                conteudo = arquivo_upload.read().decode("utf-8")
                with open(ARQUIVO_DADOS, "w", encoding="utf-8") as f:
                    f.write(conteudo)
                st.cache_data.clear()
                st.success("✅ Backup restaurado!")
                safe_rerun()
            except Exception as e:
                st.error(f"Erro restore: {e}")

    st.markdown("---")
    st.markdown("### 🗑️ Resetar Dados")
    if st.button("⚠️ RESETAR TODOS OS DADOS", type="primary", use_container_width=True):
        if os.path.exists(ARQUIVO_DADOS):
            os.remove(ARQUIVO_DADOS)
        st.cache_data.clear()
        st.success("🗑️ Dados resetados!")
        safe_rerun()


# ==================== ABA 11 - COMPRAS ====================

if aba == "11 - Compras":
    st.header("🛒 Controle de Compras")
    dados = carregar_dados()

    st.markdown("### 📝 Lançar Compra")
    with st.form("form_compra"):
        loja_compra = st.selectbox("Loja", [""] + lista_lojas(), key="loja_compra")
        tipo_compra = st.selectbox("Tipo", ["Material Escritório", "Material Limpeza", "Material Operacional", "Serviço", "Outro"], key="tipo_compra")
        descricao = st.text_input("Descrição", key="desc_compra")
        fornecedor = st.text_input("Fornecedor", key="forn_compra")
        valor = st.text_input("Valor", key="valor_compra")
        data_compra = st.text_input("Data", key="data_compra")
        status_compra = st.selectbox("Status", ["Solicitado", "Aprovado", "Entregue", "Pago", "Cancelado"], key="status_compra")
        obs_compra = st.text_area("Observações", key="obs_compra")
        btn_lancar_compra = st.form_submit_button("💾 LANÇAR COMPRA", type="primary", use_container_width=True)
        if btn_lancar_compra:
            novo = pd.DataFrame([{
                "ID": gerar_id(), "Loja": loja_compra, "Tipo": tipo_compra, "Descricao": descricao,
                "Fornecedor": fornecedor, "Valor": valor, "Data": data_compra, "Status": status_compra,
                "Observacoes": obs_compra,
                "DataCadastro": datetime.now().strftime("%d/%m/%Y %H:%M:%S"), "UsuarioCadastro": st.session_state.get("usuario", "")
            }])
            dados["Compras"] = pd.concat([dados["Compras"], novo], ignore_index=True)
            salvar_dados(dados)
            add_historico_auto("", "", "COMPRA", {"Loja": loja_compra, "Tipo": tipo_compra, "Valor": valor})
            st.success("✅ Compra lançada!")
            safe_rerun()

    st.markdown("---")
    st.markdown("### 📋 Compras Lançadas")
    compras = dados["Compras"].copy()
    if not compras.empty:
        st.dataframe(compras[["Loja","Tipo","Descricao","Fornecedor","Valor","Data","Status"]], use_container_width=True, hide_index=True)
    else:
        st.info("Nenhuma compra lançada.")
