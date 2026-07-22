import streamlit as st
import pandas as pd
import os
from datetime import datetime, timedelta
from PIL import Image
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

# ====================== CONFIGURAÇÕES GERAIS ======================
ARQUIVO = "dados_funcionarios.xlsx"
ARQUIVO_DIARIAS = "controle_diarias.xlsx"
PASTA_DOCS = "Documentos_Lojas"
PASTA_DOCS_FUNC = "Documentos_Funcionarios"
PASTA_FOTOS = "Fotos_Funcionarios"
PASTA_COMPROVANTES = "Comprovantes_Diarias"
os.makedirs(PASTA_DOCS, exist_ok=True)
os.makedirs(PASTA_DOCS_FUNC, exist_ok=True)
os.makedirs(PASTA_FOTOS, exist_ok=True)
os.makedirs(PASTA_COMPROVANTES, exist_ok=True)

MESES = ["Todos", "jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"]
SEMANAS = ["Todas", "1º Semana", "2º Semana", "3º Semana", "4º Semana"]
SITUACOES_DIARIA = ["Todas", "PENDENTE", "PAGO"]
ANOS = [str(a) for a in range(2020, datetime.now().year + 2)]

SITUACOES = [
    "Ativo", "Pré-cadastro", "Abandono", "Término de Contrato",
    "Demitido S/JC", "Demitido C/JC", "Pedido de Conta",
    "Rescisão Indireta", "Férias", "Doença", "Acidente", "Maternidade"
]

# ====================== BANCO DE DADOS ======================
@st.cache_data(ttl=0, show_spinner=False)
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
            "DataTerminoContrato",
            "DataLicenca","DiasLicenca","DataTerminoLicenca",
            "DataAfastamento","DiasAfastamento","DataRetornoAfastamento",
            "CaminhoFoto"
        ],
        "Historico": [
            "DataEvento","TipoEvento","Matricula","Nome","CPF","RG","PIS",
            "Nascimento","Admissao","Telefone","Endereco","Loja","Cargo",
            "Salario","Situacao","DataAvisoPrevio","DiasAvisoPrevio","DataTerminoAviso",
            "DataFeriasInicio","DiasFerias","DataRetornoFerias",
            "DataPedidoConta","DataRescisao","DataAbandono",
            "DataTerminoContrato",
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
            if "Situacao" in dados[aba].columns:
                dados[aba]["Situacao"] = dados[aba]["Situacao"].astype(str).str.strip()
    return dados

@st.cache_data(ttl=0, show_spinner=False)
def carregar_diarias():
    # COLUNAS EXATAMENTE COMO NA PLANILHA ENVIADA
    cols_originais = [
        "LOJA", "NOME COMPLETO DO COLABORADOR", "CPF", "DATA DA EXECUÇÃO",
        "QUANT.", "VALOR UNI.", "TOTAL", "DADOS BANCÁRIOS",
        "SUBSTITUIÇÃO", "MOTIVO DA DIARIA", "DATA DE PAGAM.",
        "situação", "Mês", "semana"
    ]
    try:
        # PULA A PRIMEIRA LINHA (instruções) E PEGA O CABEÇALHO NA LINHA 2
        df = pd.read_excel(ARQUIVO_DIARIAS, header=1, dtype=str, keep_default_na=False)
    except Exception:
        try:
            df = pd.read_excel(ARQUIVO_DIARIAS, dtype=str, keep_default_na=False)
        except Exception:
            df = pd.DataFrame(columns=cols_originais)

    # GARANTE TODAS AS COLUNAS EXISTEM NA ORDEM CORRETA
    for col in cols_originais:
        if col not in df.columns:
            df[col] = ""
    df = df[cols_originais]

    # LIMPA TEXTOS
    for col in df.columns:
        df[col] = df[col].astype(str).str.strip()

    return df

def salvar_dados(dados):
    try:
        with pd.ExcelWriter(ARQUIVO, engine="openpyxl", mode="w") as f:
            for aba, df in dados.items():
                df.to_excel(f, sheet_name=aba, index=False)
        st.cache_data.clear()
    except PermissionError:
        st.error("❌ ERRO: Arquivo dados_funcionarios.xlsx está ABERTO! Feche o Excel e tente novamente.")
        st.stop()
    except Exception as e:
        st.error(f"❌ Erro ao salvar: {str(e)}")
        st.stop()

def salvar_diarias(df_diarias):
    try:
        with pd.ExcelWriter(ARQUIVO_DIARIAS, engine="openpyxl", mode="w") as f:
            df_diarias.to_excel(f, sheet_name="Diarias", index=False)
        st.cache_data.clear()
    except PermissionError:
        st.error("❌ ERRO: O arquivo controle_diarias.xlsx está ABERTO! Feche o Excel e tente novamente.")
        st.stop()
    except Exception as e:
        st.error(f"❌ Erro ao salvar diárias: {str(e)}")
        st.stop()

def exportar_diarias_formatado(df, caminho):
    """EXPORTA EXATAMENTE COMO A PLANILHA ORIGINAL"""
    df_export = df.copy()
    df_export.to_excel(caminho, index=False, engine="openpyxl")
    wb = load_workbook(caminho)
    ws = wb.active
    ws.title = "Diarias"

    # INSTRUÇÕES NO TOPO (IGUAL ORIGINAL)
    ws.insert_rows(1)
    ws.merge_cells("A1:N1")
    ws["A1"] = "- Pagamento da diária será efetuado em até 5 dias úteis.\n- Os pagamentos de diárias só serão efetuados via transferência bancária.\n- Não é permitido pagamento em conta de terceiros."
    ws["A1"].alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    ws["A1"].font = Font(size=10)

    # ESTILOS
    cor_cabecalho = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    fonte_cabecalho = Font(color="FFFFFF", bold=True, size=10)
    alin_cabecalho = Alignment(horizontal="center", vertical="center", wrap_text=True)
    borda = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin")
    )
    cor_pendente = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    cor_pago = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")

    # LARGURAS IGUAIS AO MODELO
    larguras = {
        "A":16, "B":32, "C":18, "D":22, "E":10, "F":12, "G":12,
        "H":28, "I":14, "J":22, "K":16, "L":12, "M":8, "N":12
    }
    for l, w in larguras.items():
        ws.column_dimensions[l].width = w

    # FORMATA CABEÇALHO (LINHA 2 AGORA)
    for cell in ws[2]:
        cell.fill = cor_cabecalho
        cell.font = fonte_cabecalho
        cell.alignment = alin_cabecalho
        cell.border = borda

    # FORMATA DADOS
    for row in ws.iter_rows(min_row=3, max_row=ws.max_row):
        for cell in row:
            cell.border = borda
            cell.alignment = Alignment(horizontal="left", vertical="center")
            # DESTAQUE SITUAÇÃO
            if ws.cell(row=2, column=cell.column).value == "situação":
                if cell.value == "PENDENTE":
                    cell.fill = cor_pendente
                elif cell.value == "PAGO":
                    cell.fill = cor_pago
            # FORMATAR VALORES
            if ws.cell(row=2, column=cell.column).value in ["VALOR UNI.", "TOTAL"]:
                try:
                    cell.number_format = 'R$ #,##0.00'
                except: pass

    wb.save(caminho)
    wb.close()

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
            if not any([
                form.get("dt_pedido","").strip(), form.get("dt_rescisao","").strip(),
                form.get("dt_abandono","").strip(), form.get("dt_termino_cont","").strip()
            ]):
                form["situacao"] = "Férias"
        except:
            form["retorno_fer"] = ""
    else:
        form["retorno_fer"] = ""

    if form.get("dt_af") and form.get("dias_af") and str(form["dias_af"]).isdigit():
        try:
            dt = datetime.strptime(form["dt_af"], "%d/%m/%Y")
            form["retorno_af"] = (dt + timedelta(days=int(form["dias_af"]))).strftime("%d/%m/%Y")
        except: form["retorno_af"] = ""
    else: form["retorno_af"] = ""

    if form.get("dt_termino_cont") and form.get("dt_termino_cont").strip():
        form["situacao"] = "Término de Contrato"
    elif form.get("dt_pedido") and form.get("dt_pedido").strip():
        form["situacao"] = "Pedido de Conta"
    elif form.get("dt_rescisao") and form.get("dt_rescisao").strip():
        form["situacao"] = "Rescisão Indireta"
    elif form.get("dt_abandono") and form.get("dt_abandono").strip():
        form["situacao"] = "Abandono"

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

aba1, aba2, aba3, aba4, aba5, aba6, aba7, aba8 = st.tabs([
    "Cadastro", "Painel", "Prazos e Férias", "Histórico", "Relatórios", "📎 Documentos", "⚙️ Lojas e Cargos", "💰 CONTROLE DE DIÁRIAS"
])

# ... (TODO O RESTO DAS ABAS 1 A 7 PERMANECE EXATAMENTE IGUAL) ...

# ================ ABA 8 - CONTROLE DE DIÁRIAS (AJUSTADA 100%) ================
with aba8:
    st.subheader("💰 CONTROLE DE DIÁRIAS")
    st.info("ℹ️ Pagamento em até 5 dias úteis, via transferência bancária, não permitido conta de terceiros.")

    st.markdown("---")
    arq_diarias = st.file_uploader("📤 Carregar planilha de Diárias (.xlsx)", type=["xlsx"], key="upload_diarias")
    if arq_diarias is not None:
        with open(ARQUIVO_DIARIAS, "wb") as f:
            f.write(arq_diarias.read())
        st.success("✅ Planilha carregada com sucesso!")
        st.rerun()

    df_diarias = carregar_diarias()
    if df_diarias.empty:
        st.warning("⚠️ Nenhuma diária cadastrada. Faça upload da planilha acima.")

    # RESUMO
    st.markdown("---")
    c1, c2, c3, c4 = st.columns(4)
    with c1: st.metric("👥 Total de Diárias", len(df_diarias))
    with c2:
        try: vtotal = df_diarias["TOTAL"].replace("", "0").astype(float).sum()
        except: vtotal = 0
        st.metric("💰 Valor Total", f"R$ {vtotal:,.2f}")
    with c3:
        try: vp = df_diarias[df_diarias["situação"] == "PENDENTE"]["TOTAL"].replace("", "0").astype(float).sum()
        except: vp = 0
        st.metric("⏳ Pendente", f"R$ {vp:,.2f}")
    with c4:
        try: vpg = df_diarias[df_diarias["situação"] == "PAGO"]["TOTAL"].replace("", "0").astype(float).sum()
        except: vpg = 0
        st.metric("✅ Pago", f"R$ {vpg:,.2f}")
    st.markdown("---")

    # FILTROS
    col_p1, col_p2, col_p3, col_p4 = st.columns(4)
    with col_p1: filtro_loja_d = st.selectbox("Loja", ["Todas"] + lista_lojas(), key="filtro_loja_d")
    with col_p2: filtro_mes_d = st.selectbox("Mês", MESES, key="filtro_mes_d")
    with col_p3: filtro_sem_d = st.selectbox("Semana", SEMANAS, key="filtro_sem_d")
    with col_p4: filtro_sit_d = st.selectbox("Situação", SITUACOES_DIARIA, key="filtro_sit_d")
    busca_d = st.text_input("🔍 Pesquisar por Nome ou CPF")

    df_filtrado = df_diarias.copy()
    if filtro_loja_d != "Todas": df_filtrado = df_filtrado[df_filtrado["LOJA"] == filtro_loja_d]
    if filtro_mes_d != "Todos": df_filtrado = df_filtrado[df_filtrado["Mês"] == filtro_mes_d]
    if filtro_sem_d != "Todas": df_filtrado = df_filtrado[df_filtrado["semana"] == filtro_sem_d]
    if filtro_sit_d != "Todas": df_filtrado = df_filtrado[df_filtrado["situação"] == filtro_sit_d]
    if busca_d.strip():
        df_filtrado = df_filtrado[
            df_filtrado["NOME COMPLETO DO COLABORADOR"].str.contains(busca_d, case=False, na=False) |
            df_filtrado["CPF"].str.contains(busca_d, na=False)
        ]

    st.dataframe(df_filtrado, use_container_width=True, hide_index=True)

    # NOVA DIÁRIA - CAMPOS EXATOS DA PLANILHA
    st.markdown("---")
    st.subheader("➕ CADASTRAR NOVA DIÁRIA")
    with st.form("nova_diaria", clear_on_submit=True):
        c1, c2 = st.columns(2)
        with c1:
            loja_d = st.selectbox("LOJA *", lista_lojas())
            nome_d = st.text_input("NOME COMPLETO DO COLABORADOR *")
            cpf_d = st.text_input("CPF *")
            data_exec = st.text_input("DATA DA EXECUÇÃO * (ex: 06/07, 07/07/2026)")
            qtde_d = st.number_input("QUANT. *", min_value=1, value=1)
            valor_uni_d = st.number_input("VALOR UNI. *", min_value=0.0, format="%.2f")
        with c2:
            dados_banc_d = st.text_input("DADOS BANCÁRIOS * (Pix, Ag/Cc, etc)")
            subst_d = st.selectbox("SUBSTITUIÇÃO", ["Não", "Sim"])
            motivo_d = st.text_input("MOTIVO DA DIARIA *")
            data_pag_d = st.text_input("DATA DE PAGAM. (ex: 7/13/26)")
            situacao_d = st.selectbox("situação", ["PENDENTE", "PAGO"])
            mes_d = st.selectbox("Mês", MESES[1:])
            semana_d = st.selectbox("semana", SEMANAS[1:])

        submitted = st.form_submit_button("💾 SALVAR DIÁRIA", type="primary")
        if submitted:
            erros = []
            if not loja_d.strip(): erros.append("Loja")
            if not nome_d.strip(): erros.append("Nome")
            if not cpf_d.strip(): erros.append("CPF")
            if not data_exec.strip(): erros.append("Data da Execução")
            if not dados_banc_d.strip(): erros.append("Dados Bancários")
            if not motivo_d.strip(): erros.append("Motivo")
            if qtde_d <=0: erros.append("Quantidade > 0")
            if valor_uni_d <=0: erros.append("Valor Unitário > 0")
            
            if erros:
                st.error("❌ Campos obrigatórios: " + ", ".join(erros))
            else:
                total_d = qtde_d * valor_uni_d
                nova_linha = {
                    "LOJA": loja_d,
                    "NOME COMPLETO DO COLABORADOR": nome_d.strip().upper(),
                    "CPF": cpf_d.strip(),
                    "DATA DA EXECUÇÃO": data_exec.strip(),
                    "QUANT.": str(qtde_d),
                    "VALOR UNI.": f"{valor_uni_d:.2f}",
                    "TOTAL": f"{total_d:.2f}",
                    "DADOS BANCÁRIOS": dados_banc_d.strip(),
                    "SUBSTITUIÇÃO": subst_d,
                    "MOTIVO DA DIARIA": motivo_d.strip().upper(),
                    "DATA DE PAGAM.": data_pag_d.strip(),
                    "situação": situacao_d,
                    "Mês": mes_d,
                    "semana": semana_d
                }
                df_diarias = pd.concat([df_diarias, pd.DataFrame([nova_linha])], ignore_index=True)
                salvar_diarias(df_diarias)
                st.success("✅ Diária cadastrada com sucesso!")
                st.rerun()

    # EDITAR / EXCLUIR
    st.markdown("---")
    st.subheader("✏️ EDITAR / EXCLUIR")
    if not df_diarias.empty:
        idx_edit = st.selectbox("Selecione a diária", df_diarias.index,
            format_func=lambda i: f"[{i}] {df_diarias.loc[i,'NOME COMPLETO DO COLABORADOR']} | {df_diarias.loc[i,'LOJA']} | R$ {df_diarias.loc[i,'TOTAL']}")
        
        with st.form("editar_diaria"):
            linha = df_diarias.loc[idx_edit]
            c1, c2 = st.columns(2)
            with c1:
                e_loja = st.selectbox("LOJA", lista_lojas(), index=lista_lojas().index(linha["LOJA"]) if linha["LOJA"] in lista_lojas() else 0)
                e_nome = st.text_input("NOME COMPLETO DO COLABORADOR", linha["NOME COMPLETO DO COLABORADOR"])
                e_cpf = st.text_input("CPF", linha["CPF"])
                e_data_exec = st.text_input("DATA DA EXECUÇÃO", linha["DATA DA EXECUÇÃO"])
                e_qtde = st.number_input("QUANT.", min_value=1, value=int(linha["QUANT."]) if str(linha["QUANT."]).isdigit() else 1)
                e_valor = st.number_input("VALOR UNI.", min_value=0.0, value=float(str(linha["VALOR UNI."]).replace(",",".")), format="%.2f")
            with c2:
                e_dados_banc = st.text_input("DADOS BANCÁRIOS", linha["DADOS BANCÁRIOS"])
                e_subst = st.selectbox("SUBSTITUIÇÃO", ["Não","Sim"], index=["Não","Sim"].index(linha["SUBSTITUIÇÃO"]) if linha["SUBSTITUIÇÃO"] in ["Não","Sim"] else 0)
                e_motivo = st.text_input("MOTIVO DA DIARIA", linha["MOTIVO DA DIARIA"])
                e_data_pag = st.text_input("DATA DE PAGAM.", linha["DATA DE PAGAM."])
                e_sit = st.selectbox("situação", ["PENDENTE","PAGO"], index=0 if linha["situação"]=="PENDENTE" else 1)
                e_mes = st.selectbox("Mês", MESES[1:], index=MESES[1:].index(linha["Mês"]) if linha["Mês"] in MESES[1:] else 0)
                e_sem = st.selectbox("semana", SEMANAS[1:], index=SEMANAS[1:].index(linha["semana"]) if linha["semana"] in SEMANAS[1:] else 0)

            col_salvar, col_excluir = st.columns(2)
            if col_salvar.form_submit_button("💾 SALVAR", type="primary"):
                df_diarias.loc[idx_edit] = {
                    "LOJA": e_loja, "NOME COMPLETO DO COLABORADOR": e_nome.strip().upper(),
                    "CPF": e_cpf.strip(), "DATA DA EXECUÇÃO": e_data_exec.strip(),
                    "QUANT.": str(e_qtde), "VALOR UNI.": f"{e_valor:.2f}",
                    "TOTAL": f"{e_qtde * e_valor:.2f}", "DADOS BANCÁRIOS": e_dados_banc.strip(),
                    "SUBSTITUIÇÃO": e_subst, "MOTIVO DA DIARIA": e_motivo.strip().upper(),
                    "DATA DE PAGAM.": e_data_pag.strip(), "situação": e_sit,
                    "Mês": e_mes, "semana": e_sem
                }
                salvar_diarias(df_diarias)
                st.success("✅ Atualizado!")
                st.rerun()
            if col_excluir.form_submit_button("🗑️ EXCLUIR"):
                df_diarias.drop(idx_edit, inplace=True)
                df_diarias.reset_index(drop=True, inplace=True)
                salvar_diarias(df_diarias)
                st.success("🗑️ Excluído!")
                st.rerun()

    # EXPORTAR - ESTRUTURA EXATA
    st.markdown("---")
    st.subheader("📤 EXPORTAR PLANILHA ORIGINAL")
    if not df_filtrado.empty:
        nome_arq = f"Diarias_{datetime.now().strftime('%d%m%Y_%H%M')}.xlsx"
        exportar_diarias_formatado(df_filtrado, nome_arq)
        with open(nome_arq, "rb") as f:
            st.download_button("⬇️ BAIXAR EXCEL", f, file_name=nome_arq)
        os.remove(nome_arq)
    else:
        st.info("Filtre os dados primeiro.")
