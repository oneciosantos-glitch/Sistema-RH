import streamlit as st
import pandas as pd
import os
from datetime import datetime, timedelta
from PIL import Image
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

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
    cols_padrao = [
        "LOJA","NOME COLABORADOR","CPF","DATA EXECUCAO","QTDE DE DIARIAS","VALOR UNITARIO","VALOR TOTAL",
        "DADOS BANCÁRIOS","SUBSTITUICAO","MOTIVO","DATA PAGAMENTO","SITUACAO","MES","SEMANA","ANO",
        "CARGO","DATA CADASTRO","COMPROVANTE"
    ]
    try:
        # Tenta ler com header na segunda linha (pula texto de instruções)
        df = pd.read_excel(ARQUIVO_DIARIAS, header=1, dtype=str, keep_default_na=False)
        # Renomeia colunas da planilha do usuário para o padrão do app
        rename_map = {
            'NOME COMPLETO DO COLABORADOR': 'NOME COLABORADOR',
            'QUANT.': 'QTDE DE DIARIAS',
            'VALOR UNI.': 'VALOR UNITARIO',
            'TOTAL': 'VALOR TOTAL',
            'DADOS BANCÁRIOS': 'DADOS BANCÁRIOS',
            'MOTIVO DA DIARIA': 'MOTIVO',
            'situação': 'SITUACAO',
            'Mês': 'MES',
            'semana': 'SEMANA',
            'DATA DA EXECUÇÃO': 'DATA EXECUCAO',
            'SUBSTITUIÇÃO': 'SUBSTITUICAO',
            'DATA DE PAGAM.': 'DATA PAGAMENTO'
        }
        df = df.rename(columns=rename_map)
    except Exception:
        # Se falhar, tenta ler normalmente (já no formato do app)
        try:
            df = pd.read_excel(ARQUIVO_DIARIAS, dtype=str, keep_default_na=False)
        except Exception:
            df = pd.DataFrame(columns=cols_padrao)

    # Garante que todas as colunas padrão existam
    for col in cols_padrao:
        if col not in df.columns:
            df[col] = ""

    # Extrai ANO da DATA PAGAMENTO quando estiver vazio
    for i in df.index:
        if str(df.at[i, "ANO"]).strip() == "":
            try:
                dt = pd.to_datetime(str(df.at[i, "DATA PAGAMENTO"]).strip())
                df.at[i, "ANO"] = str(dt.year)
            except Exception:
                df.at[i, "ANO"] = str(datetime.now().year)

    # Limpa strings
    for col in ["NOME COLABORADOR", "CPF", "LOJA", "MOTIVO", "SITUACAO", "MES", "SEMANA"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()

    # Garante ordem correta das colunas
    df = df[[c for c in cols_padrao if c in df.columns]]
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
    """Exporta DataFrame de diárias para Excel com a mesma formatação da planilha padrão."""
    df_export = df.copy()
    df_export.to_excel(caminho, index=False, engine="openpyxl")
    wb = load_workbook(caminho)
    ws = wb.active
    ws.title = "Diarias"

    # Cores e estilos exatos da planilha padrão
    fill_instrucoes = PatternFill(start_color="1B2D4F", end_color="1B2D4F", fill_type="solid")
    font_instrucoes = Font(name="Calibri", size=11, bold=False, color="FFFFFF")
    align_instrucoes = Alignment(horizontal="left", vertical="center", wrap_text=True)

    fill_header = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    align_header = Alignment(horizontal="center", vertical="center")

    borda = Border(
        left=Side(style="thin", color="000000"),
        right=Side(style="thin", color="000000"),
        top=Side(style="thin", color="000000"),
        bottom=Side(style="thin", color="000000")
    )

    fill_pendente = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    font_pendente = Font(name="Calibri", size=11, bold=False, color="9C0006")

    font_data = Font(name="Calibri", size=11, bold=False, color="000000")
    align_data_center = Alignment(horizontal="center", vertical="center")
    align_data_left = Alignment(horizontal="left", vertical="center")

    # Mapear colunas do app para posição
    headers_app = list(df_export.columns)
    num_cols = len(headers_app)

    # Inserir linha de instruções no topo e mesclar
    ws.insert_rows(1)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=num_cols)
    cell_instr = ws.cell(row=1, column=1)
    cell_instr.value = (
        "- Pagamento da diária será efetuado em até 5 dias úteis.\n"
        "- Os pagamentos de diárias só serão efetuados via transferência bancária.\n"
        "- Não é permitido pagamento em conta de terceiros."
    )
    cell_instr.fill = fill_instrucoes
    cell_instr.font = font_instrucoes
    cell_instr.alignment = align_instrucoes
    cell_instr.border = borda
    ws.row_dimensions[1].height = 63

    # Aplicar borda nas células mescladas da linha 1
    for c in range(2, num_cols + 1):
        ws.cell(row=1, column=c).border = borda

    # Formatar cabeçalho (linha 2 após inserção)
    for col_idx, header in enumerate(headers_app, 1):
        cell = ws.cell(row=2, column=col_idx)
        cell.fill = fill_header
        cell.font = font_header
        cell.alignment = align_header
        cell.border = borda

    # Formatar dados (a partir da linha 3)
    for row in ws.iter_rows(min_row=3, max_row=ws.max_row):
        for cell in row:
            cell.border = borda
            cell.font = font_data
            header = ws.cell(row=2, column=cell.column).value
            # Nome alinha à esquerda, resto centralizado
            if header and "NOME" in str(header).upper():
                cell.alignment = align_data_left
            else:
                cell.alignment = align_data_center

            # Destacar PENDENTE na coluna SITUACAO
            if header == "SITUACAO" and cell.value == "PENDENTE":
                cell.fill = fill_pendente
                cell.font = font_pendente

            # Formato de moeda nas colunas de valor
            if header in ["VALOR UNITARIO", "VALOR TOTAL"]:
                try:
                    val = float(str(cell.value).replace(",", "."))
                    cell.number_format = r'_-"R$"\ * #,##0.00_-;\-"R$"\ * #,##0.00_-;_-"R$"\ * "-"??_-;_-@_-'
                    cell.value = val
                except:
                    pass

    # Larguras de coluna
    larguras = {
        "LOJA": 12, "MES": 10, "SEMANA": 12, "ANO": 8,
        "NOME COLABORADOR": 38, "CPF": 16, "CARGO": 18,
        "DADOS BANCÁRIOS": 35,
        "MOTIVO": 28, "QTDE DE DIARIAS": 10, "VALOR UNITARIO": 14,
        "VALOR TOTAL": 14, "SITUACAO": 12, "DATA CADASTRO": 18,
        "COMPROVANTE": 35
    }
    for col_idx, header in enumerate(headers_app, 1):
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = larguras.get(header, 18)

    wb.save(caminho)
    wb.close()

def lista_lojas():
    return [
        "Assaí Atacadista Batista Campos",
        "Assaí Atacadista Almirante Barroso",
        "Assaí Atacadista Castanhal",
        "Assaí Atacadista Ananindeua",
        "Assaí Atacadista Augusto Monte Negro",
        "Assaí Atacadista Boa Vista",
        "Assaí Atacadista Manaus",
        "Assaí Atacadista Macapá",
        "Assaí Atacadista Belém",
        "Smart Fit Shopping Manoa",
        "Smart Fit Shopping Cidade Leste",
        "Smart Fit Macapá Shopping",
        "Smart Fit Shopping Grande Circular",
        "Smart Fit Shopping Via Norte",
        "Smart Fit Cidade Nova",
        "Smart Fit Parque Mosaico",
        "Smart Fit Cachoeirinha",
        "Smart Fit Flores",
        "Smart Fit Ponta Negra",
        "Smart Fit Nova Porto Velho",
        "Smart Fit Porto Velho Flodoaldo",
        "Smart Fit Alvorada",
        "Smart Fit Novo Aleixo",
        "Smart Fit São José do Operário",
        "Smart Fit Santana Macapá",
        "Smart Fit Toequato Tapajós",
        "Self Fit Hiper DB Ponta Negra",
        "Self Fit Manaus Plaza Shopping",
        "Self Fit Vieira Alves",
    ]

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

# ⚠️ LINHA OBRIGATÓRIA: CRIA TODAS AS ABAS ANTES DE USÁ-LAS
aba1, aba2, aba3, aba4, aba5, aba6, aba7, aba8 = st.tabs([
    "Cadastro", "Painel", "Prazos e Férias", "Histórico", "Relatórios", "📎 Documentos", "⚙️ Lojas e Cargos", "💰 CONTROLE DE DIÁRIAS"
])

# ================ ABA 1 - CADASTRO ================
with aba1:
    dados = carregar_dados()
    
    busca = st.text_input("🔍 Buscar por Matrícula ou Nome", placeholder="Digite exatamente como está na planilha")
    
    col_f1, col_f2, col_f3 = st.columns(3)
    with col_f1:
        filtro_loja = st.selectbox("Filtrar por Loja", ["Todas"] + lista_lojas())
    with col_f2:
        filtro_sit = st.selectbox("Filtrar por Situação", ["Todas"] + SITUACOES)
    with col_f3:
        filtro_cargo = st.selectbox("Filtrar por Cargo", ["Todos"] + lista_cargos())

    lista = dados["Base_Dados"].copy()
    lista["Matricula"] = lista["Matricula"].fillna("").astype(str).str.strip()
    lista["Nome"] = lista["Nome"].fillna("").astype(str).str.strip()

    if busca.strip():
        lista = lista[
            (lista["Matricula"].str.contains(busca, case=False, na=False)) |
            (lista["Nome"].str.contains(busca, case=False, na=False))
        ]
    if filtro_loja != "Todas":
        lista = lista[lista["Loja"] == filtro_loja]
    if filtro_sit != "Todas":
        lista = lista[lista["Situacao"] == filtro_sit]
    if filtro_cargo != "Todos":
        lista = lista[lista["Cargo"] == filtro_cargo]

    st.dataframe(
        lista[["Matricula","Nome","Loja","Situacao","Cargo"]],
        use_container_width=True, hide_index=True
    )

    mat_sel = st.text_input("✏️ Digite a Matrícula EXATA para editar / excluir", placeholder="Igual coluna A da planilha")
    reg = pd.DataFrame()
    if mat_sel.strip():
        mat_busca = str(mat_sel).strip()
        reg = dados["Base_Dados"][dados["Base_Dados"]["Matricula"] == mat_busca]

    val_campo = lambda nome: reg.iloc[0][nome] if not reg.empty else ""

    prazos_exp = []
    if not reg.empty and val_campo("Admissao").strip():
        try:
            dt_adm = datetime.strptime(val_campo("Admissao"), "%d/%m/%Y")
            hoje = datetime.now()
            dias_corridos = (hoje - dt_adm).days
            for prazo in [30, 45, 60, 90]:
                rest = prazo - dias_corridos
                if rest > 0:
                    status = f"Faltam {rest} dias"
                elif rest == 0:
                    status = "HOJE"
                else:
                    status = f"Vencido há {abs(rest)} dias"
                prazos_exp.append([f"{prazo} dias", (dt_adm + timedelta(days=prazo)).strftime("%d/%m/%Y"), status])
        except:
            pass

    if not reg.empty:
        temp = {
            "dt_aviso": val_campo("DataAvisoPrevio"), "dias_aviso": val_campo("DiasAvisoPrevio"),
            "dt_lic": val_campo("DataLicenca"), "dias_lic": val_campo("DiasLicenca"),
            "dt_fer": val_campo("DataFeriasInicio"), "dias_fer": val_campo("DiasFerias"),
            "dt_af": val_campo("DataAfastamento"), "dias_af": val_campo("DiasAfastamento"),
            "dt_pedido": val_campo("DataPedidoConta"), "dt_rescisao": val_campo("DataRescisao"),
            "dt_abandono": val_campo("DataAbandono"), "dt_termino_cont": val_campo("DataTerminoContrato"),
            "situacao": val_campo("Situacao"), "caminho_foto": val_campo("CaminhoFoto")
        }
        temp = calcular_e_atualizar(temp)
        term_aviso_val, term_lic_val, ret_fer_val, ret_af_val, situacao_val, caminho_foto_atual = temp["termino_aviso"], temp["termino_lic"], temp["retorno_fer"], temp["retorno_af"], temp["situacao"], temp["caminho_foto"]
    else:
        term_aviso_val = term_lic_val = ret_fer_val = ret_af_val = caminho_foto_atual = ""
        situacao_val = "Ativo"

    if st.button("🗑️ LIMPAR TODOS OS CAMPOS", use_container_width=True, type="secondary"):
        st.rerun()

    with st.form("form_cadastro", clear_on_submit=True):
        st.subheader("Dados Básicos")
        col_foto, col_dados = st.columns([1,3])
        
        with col_foto:
            st.markdown("**Foto do Funcionário**")
            if caminho_foto_atual and os.path.exists(caminho_foto_atual):
                st.image(caminho_foto_atual, width=180, caption="Foto atual")
            else:
                st.info("Sem foto")
            
            nova_foto = st.file_uploader("Enviar/Trocar foto", type=["jpg","jpeg","png"], key=f"foto_{mat_sel}")
            excluir_foto = st.checkbox("🗑️ Excluir foto atual", value=False)

        with col_dados:
            c1,c2,c3 = st.columns(3)
            with c1:
                matricula = st.text_input("Matrícula * (igual planilha)", value=val_campo("Matricula"))
                nome = st.text_input("Nome Completo", value=val_campo("Nome"))
                cpf = st.text_input("CPF", value=val_campo("CPF"))
                rg = st.text_input("RG", value=val_campo("RG"))
                pis = st.text_input("PIS", value=val_campo("PIS"))
            with c2:
                nascimento = st.text_input("Data Nascimento (dd/mm/aaaa)", value=val_campo("Nascimento"))
                admissao = st.text_input("Data Admissão (dd/mm/aaaa)", value=val_campo("Admissao"))
                telefone = st.text_input("Telefone", value=val_campo("Telefone"))
                endereco = st.text_input("Endereço Completo", value=val_campo("Endereco"))
            with c3:
                lojas = lista_lojas()
                idx_loja = lojas.index(val_campo("Loja")) if val_campo("Loja") in lojas else 0
                loja = st.selectbox("🏬 Loja", lojas, index=idx_loja)

                cargos = lista_cargos()
                idx_cargo = cargos.index(val_campo("Cargo")) if val_campo("Cargo") in cargos else 0
                cargo = st.selectbox("💼 Cargo", cargos, index=idx_cargo)

                salario = st.text_input("Salário", value=val_campo("Salario"))

                idx_sit = SITUACOES.index(situacao_val) if situacao_val in SITUACOES else 0
                situacao = st.selectbox("📊 Situação", SITUACOES, index=idx_sit)

        if prazos_exp:
            st.markdown("---")
            st.subheader("⏳ PRAZOS DE EXPERIÊNCIA")
            st.dataframe(
                pd.DataFrame(prazos_exp, columns=["Prazo", "Data Final", "Situação"]),
                use_container_width=True, hide_index=True
            )
        elif not reg.empty:
            st.info("ℹ️ Informe a Data de Admissão para visualizar os prazos.")

        st.markdown("---")
        st.subheader("Eventos Trabalhistas")
        av1,av2,av3 = st.columns(3)
        with av1:
            st.markdown("**Aviso Prévio**")
            dt_aviso = st.text_input("Data Aviso", value=val_campo("DataAvisoPrevio"))
            dias_aviso = st.text_input("Dias Aviso", value=val_campo("DiasAvisoPrevio"))
            term_aviso = st.text_input("Término Aviso", value=term_aviso_val, disabled=True)
        with av2:
            st.markdown("**Licença**")
            dt_lic = st.text_input("Data Licença", value=val_campo("DataLicenca"))
            dias_lic = st.text_input("Dias Licença", value=val_campo("DiasLicenca"))
            term_lic = st.text_input("Término Licença", value=term_lic_val, disabled=True)
        with av3:
            st.markdown("**Férias**")
            dt_fer = st.text_input("Início Férias", value=val_campo("DataFeriasInicio"))
            dias_fer = st.text_input("Dias Férias", value=val_campo("DiasFerias"))
            ret_fer = st.text_input("Retorno Férias", value=ret_fer_val, disabled=True)

        af1,af2 = st.columns(2)
        with af1:
            st.markdown("**Afastamento**")
            dt_af = st.text_input("Data Afastamento", value=val_campo("DataAfastamento"))
            dias_af = st.text_input("Dias Afastamento", value=val_campo("DiasAfastamento"))
            ret_af = st.text_input("Retorno Afastamento", value=ret_af_val, disabled=True)
            tipo_af = st.selectbox("Tipo Afastamento", ["Nenhum", "Doença", "Acidente", "Maternidade"])
        with af2:
            st.markdown("**Desligamento**")
            dt_ped = st.text_input("Data Pedido Conta", value=val_campo("DataPedidoConta"))
            dt_res = st.text_input("Data Rescisão", value=val_campo("DataRescisao"))
            dt_aband = st.text_input("Data Abandono", value=val_campo("DataAbandono"))
            dt_termino_cont = st.text_input("📅 Data Término de Contrato", value=val_campo("DataTerminoContrato"))

        btn_salvar = st.form_submit_button("💾 SALVAR CADASTRO", type="primary", use_container_width=True)
        if btn_salvar:
            matricula_tratada = str(matricula).strip()
            if not matricula_tratada:
                st.error("❌ INFORME A MATRÍCULA!")
                st.stop()
            if tipo_af != "Nenhum" and dt_af.strip():
                situacao = tipo_af

            caminho_final_foto = caminho_foto_atual
            if excluir_foto and caminho_final_foto and os.path.exists(caminho_final_foto):
                os.remove(caminho_final_foto)
                caminho_final_foto = ""
            if nova_foto:
                if caminho_final_foto and os.path.exists(caminho_final_foto):
                    os.remove(caminho_final_foto)
                extensao = os.path.splitext(nova_foto.name)[1].lower()
                nome_foto = f"{matricula_tratada}_foto_{datetime.now().strftime('%Y%m%d%H%M%S')}{extensao}"
                caminho_final_foto = os.path.join(PASTA_FOTOS, nome_foto)
                img = Image.open(nova_foto)
                img.save(caminho_final_foto)

            dados_form = calcular_e_atualizar({
                "mat": matricula_tratada, "nome": nome, "cpf": cpf, "rg": rg, "pis": pis,
                "nasc": nascimento, "adm": admissao, "tel": telefone, "end": endereco,
                "loja": loja, "cargo": cargo, "sal": salario, "situacao": situacao,
                "dt_aviso": dt_aviso, "dias_aviso": dias_aviso, "termino_aviso": term_aviso,
                "dt_lic": dt_lic, "dias_lic": dias_lic, "termino_lic": term_lic,
                "dt_fer": dt_fer, "dias_fer": dias_fer, "retorno_fer": ret_fer,
                "dt_af": dt_af, "dias_af": dias_af, "retorno_af": ret_af,
                "dt_pedido": dt_ped, "dt_rescisao": dt_res, "dt_abandono": dt_aband,
                "dt_termino_cont": dt_termino_cont
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
                "DataTerminoContrato": dados_form["dt_termino_cont"],
                "DataLicenca": dados_form["dt_lic"], "DiasLicenca": dados_form["dias_lic"],
                "DataTerminoLicenca": dados_form["termino_lic"],
                "DataAfastamento": dados_form["dt_af"], "DiasAfastamento": dados_form["dias_af"],
                "DataRetornoAfastamento": dados_form["retorno_af"],
                "CaminhoFoto": caminho_final_foto
            }
            indice = dados["Base_Dados"].index[dados["Base_Dados"]["Matricula"] == dados_form["mat"]].tolist()
            acao_hist = "Atualização Cadastral" if indice else "Novo Cadastro"
            if indice: dados["Base_Dados"].iloc[indice[0]] = registro_final
            else: dados["Base_Dados"] = pd.concat([dados["Base_Dados"], pd.DataFrame([registro_final])], ignore_index=True)
            salvar_dados(dados)
            add_historico_auto(dados_form["mat"], dados_form["nome"], acao_hist, registro_final)
            st.success(f"✅ Salvo! Matrícula: **{dados_form['mat']}**")
            st.rerun()

    if mat_sel.strip() and st.button("🗑️ EXCLUIR REGISTRO", use_container_width=True, type="secondary"):
        if st.checkbox("⚠️ CONFIRMA EXCLUSÃO PERMANENTE?"):
            indice = dados["Base_Dados"].index[dados["Base_Dados"]["Matricula"] == mat_sel.strip()].tolist()
            if indice:
                dados_excluir = dados["Base_Dados"].iloc[indice[0]].to_dict()
                if dados_excluir.get("CaminhoFoto") and os.path.exists(dados_excluir["CaminhoFoto"]):
                    os.remove(dados_excluir["CaminhoFoto"])
                docs_excluir = dados["Docs_Funcionarios"][dados["Docs_Funcionarios"]["Matricula"] == mat_sel.strip()]
                for _, d in docs_excluir.iterrows():
                    if os.path.exists(d["Caminho"]): os.remove(d["Caminho"])
                dados["Docs_Funcionarios"] = dados["Docs_Funcionarios"][dados["Docs_Funcionarios"]["Matricula"] != mat_sel.strip()]
                dados["Base_Dados"].drop(indice[0], inplace=True)
                salvar_dados(dados)
                add_historico_auto(mat_sel.strip(), dados_excluir["Nome"], "Exclusão de Cadastro", dados_excluir)
                st.success("✅ Registro, foto e documentos excluídos!")
                st.rerun()

    st.markdown("---")
    st.subheader("📎 DOCUMENTOS DO FUNCIONÁRIO")
    if mat_sel.strip() and not reg.empty:
        mat_atual = mat_sel.strip()
        nome_atual = val_campo("Nome")
        tipo_doc = st.selectbox("Tipo de Documento", [
            "RG", "CPF", "PIS", "Carteira de Trabalho", "Comprovante Residência",
            "Exame Admissional", "Exame Demissional", "Contrato", "Atestados",
            "Férias", "Rescisão", "Outros"
        ])
        arquivos_func = st.file_uploader("Anexar documentos", type=["pdf","doc","docx","xls","xlsx","jpg","png"], accept_multiple_files=True, key=f"up_{mat_atual}")
        if arquivos_func and st.button("SALVAR DOCUMENTOS", type="primary"):
            qtd = 0
            for arq in arquivos_func:
                nome_arq = f"{mat_atual}_{tipo_doc}_{datetime.now().strftime('%Y%m%d%H%M%S%f')}_{arq.name}"
                caminho = os.path.join(PASTA_DOCS_FUNC, nome_arq)
                with open(caminho, "wb") as f: f.write(arq.read())
                dados["Docs_Funcionarios"] = pd.concat([dados["Docs_Funcionarios"], pd.DataFrame([{
                    "Matricula": mat_atual, "Nome": nome_atual, "TipoDoc": tipo_doc,
                    "NomeArquivo": arq.name, "Caminho": caminho,
                    "DataAnexado": datetime.now().strftime("%d/%m/%Y %H:%M")
                }])], ignore_index=True)
                qtd += 1
            salvar_dados(dados)
            st.success(f"✅ {qtd} documento(s) salvo(s)!")
            st.rerun()
        st.markdown("---")
        docs_func = dados["Docs_Funcionarios"][dados["Docs_Funcionarios"]["Matricula"] == mat_atual]
        if docs_func.empty: st.info("📂 Nenhum documento anexado.")
        else:
            st.markdown(f"**Total: {len(docs_func)} documento(s)**")
            for idx, doc in docs_func.iterrows():
                with st.expander(f"📄 {doc['TipoDoc']} - {doc['NomeArquivo']} | {doc['DataAnexado']}"):
                    col_v, col_b, col_e = st.columns([3,1,1])
                    with col_b:
                        with open(doc["Caminho"], "rb") as f: st.download_button("⬇️ BAIXAR", f, file_name=doc["NomeArquivo"], key=f"dw_{idx}")
                    with col_e:
                        if st.button("🗑️ EXCLUIR", key=f"del_{idx}"):
                            if os.path.exists(doc["Caminho"]): os.remove(doc["Caminho"])
                            dados["Docs_Funcionarios"].drop(idx, inplace=True)
                            salvar_dados(dados)
                            st.rerun()
    else:
        st.info("ℹ️ Digite a Matrícula exata para ver/anexar documentos.")

# ================ ABA 2 - PAINEL ================
with aba2:
    st.subheader("📊 RESUMO GERAL")
    dados_painel = carregar_dados()
    base = dados_painel["Base_Dados"].copy()
    base["Situacao"] = base["Situacao"].fillna("").astype(str).str.strip()

    contagem = {
        "👷 Ativo": len(base[base["Situacao"] == "Ativo"]),
        "📝 Pré-cadastro": len(base[base["Situacao"] == "Pré-cadastro"]),
        "🏖️ Férias": len(base[base["Situacao"] == "Férias"]),
        "🚪 Abandono": len(base[base["Situacao"] == "Abandono"]),
        "⏹️ Término de Contrato": len(base[base["Situacao"] == "Término de Contrato"]),
        "📉 Demitido S/JC": len(base[base["Situacao"] == "Demitido S/JC"]),
        "📉 Demitido C/JC": len(base[base["Situacao"] == "Demitido C/JC"]),
        "🙋 Pedido de Conta": len(base[base["Situacao"] == "Pedido de Conta"]),
        "⚖️ Rescisão Indireta": len(base[base["Situacao"] == "Rescisão Indireta"]),
        "🏥 Doença": len(base[base["Situacao"] == "Doença"]),
        "🚑 Acidente": len(base[base["Situacao"] == "Acidente"]),
        "🤰 Maternidade": len(base[base["Situacao"] == "Maternidade"])
    }

    cols = st.columns(3)
    for i, (rotulo, qtd) in enumerate(contagem.items()):
        cols[i % 3].metric(rotulo, qtd)

    if st.button("🔄 Atualizar Resumo"):
        st.cache_data.clear()
        st.rerun()

    st.markdown("---")
    st.subheader("🔍 Conferência Rápida - Apenas Férias")
    tab_fer = base[base["Situacao"] == "Férias"][["Matricula","Nome","Loja","DataFeriasInicio","DataRetornoFerias"]]
    if tab_fer.empty:
        st.warning("⚠️ Nenhum funcionário com situação marcada como 'Férias' no momento.")
        st.info("💡 Dica: Se a data de férias estiver preenchida mas a situação não for 'Férias', edite o cadastro e confirme se a situação está selecionada corretamente — os dados não são apagados!")
    else:
        st.dataframe(tab_fer, use_container_width=True, hide_index=True)

# ================ ABA 3 - PRAZOS E FÉRIAS ================
with aba3:
    hoje = datetime.now()
    st.subheader("⚠️ PRAZOS DE EXPERIÊNCIA PRÓXIMOS")
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
    filtro_loja = st.selectbox("Loja", ["Todas"] + lista_lojas(), key="fl")
    filtro_mes = st.selectbox("Mês", MESES, key="fm")
    tabela_fer = []
    for _, f in dados["Base_Dados"].iterrows():
        if f["Situacao"] not in ["Ativo","Pré-cadastro"]: continue
        if filtro_loja != "Todas" and f["Loja"] != filtro_loja: continue
        try:
            dt = datetime.strptime(str(f["Admissao"]).strip(), "%d/%m/%Y")
            if filtro_mes != "Todos" and dt.month != [1,2,3,4,5,6,7,8,9,10,11,12][MESES.index(filtro_mes)-1]: continue
            meses = (hoje.year - dt.year)*12 + (hoje.month - dt.month) - (1 if hoje.day < dt.day else 0)
            if 23 <= meses < 24:
                tabela_fer.append([f["Matricula"], f["Nome"], f["Loja"], f["Cargo"], f["Admissao"], f"{meses}m"])
        except: pass
    st.dataframe(pd.DataFrame(tabela_fer, columns=["Matrícula","Nome","Loja","Cargo","Admissão","Tempo"]), use_container_width=True, hide_index=True)

# ================ ABA 4 - HISTÓRICO ================
with aba4:
    st.subheader("📝 HISTÓRICO")
    st.dataframe(dados["Historico"][["DataEvento","TipoEvento","Matricula","Nome","Situacao","Detalhes"]], use_container_width=True, hide_index=True)
    with st.form("add_ev"):
        t,d,det = st.columns([1,1,3])
        te = t.selectbox("Tipo", ["Reunião","Atestado","Advertência","Elogio","Outros"])
        de = d.text_input("Data", value=datetime.now().strftime("%d/%m/%Y"))
        dee = det.text_input("Detalhes")
        if st.form_submit_button("✅ ADICIONAR") and mat_sel.strip():
            rf = dados["Base_Dados"][dados["Base_Dados"]["Matricula"] == mat_sel.strip()]
            if not rf.empty:
                nr = {"DataEvento":de,"TipoEvento":te,"Detalhes":dee}
                nr.update(rf.iloc[0].to_dict())
                ih = dados["Historico"].index[dados["Historico"]["Matricula"] == mat_sel.strip()].tolist()
                if ih: dados["Historico"].iloc[ih[0]] = nr
                else: dados["Historico"] = pd.concat([dados["Historico"], pd.DataFrame([nr])], ignore_index=True)
                salvar_dados(dados)
                st.success("Adicionado!")
                st.rerun()

# ================ ABA 5 - RELATÓRIOS ================
with aba5:
    st.subheader("📄 RELATÓRIOS")
    rel = st.selectbox("Escolha", ["Prazos Experiência","Ativos","Pré-cadastro","Férias","Afastados","Avisos","Histórico","Individual"])
    if rel == "Individual":
        mr = st.text_input("Matrícula")
        if mr.strip():
            fd = dados["Base_Dados"][dados["Base_Dados"]["Matricula"] == mr.strip()]
            fh = dados["Historico"][dados["Historico"]["Matricula"] == mr.strip()]
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

# ================ ABA 6 - DOCUMENTOS DAS LOJAS ================
with aba6:
    st.subheader("📎 DOCUMENTOS DAS LOJAS")
    ls = lista_lojas()
    l,m,a = st.columns(3)
    sl = l.selectbox("Loja", ls)
    sm = m.selectbox("Mês", MESES)
    sa = a.selectbox("Ano", ANOS, index=ANOS.index(str(datetime.now().year)))
    st.markdown("---")
    arquivos = st.file_uploader("Anexar arquivos", type=["pdf","doc","docx","xls","xlsx","jpg","png"], accept_multiple_files=True)
    resp = st.text_input("Responsável")
    if arquivos and st.button("SALVAR TODOS", type="primary"):
        salvos = 0
        for arq in arquivos:
            nome = f"{sl}_{sm}_{sa}_{datetime.now().strftime('%Y%m%d%H%M%S%f')}_{arq.name}"
            cam = os.path.join(PASTA_DOCS, nome)
            with open(cam,"wb") as f: f.write(arq.read())
            dados["Docs_Lojas"] = pd.concat([dados["Docs_Lojas"], pd.DataFrame([{
                "Loja":sl,"Mes":sm,"Ano":sa,"NomeArquivo":arq.name,"Caminho":cam,
                "DataAnexado":datetime.now().strftime("%d/%m/%Y %H:%M"),"Responsavel":resp
            }])], ignore_index=True)
            salvos += 1
        salvar_dados(dados)
        st.success(f"✅ {salvos} arquivo(s) salvo(s)!")
        st.rerun()
    st.markdown("---")
    filt = dados["Docs_Lojas"].copy()
    if sl != "Todas": filt = filt[filt["Loja"]==sl]
    if sm != "Todos": filt = filt[filt["Mes"]==sm]
    filt = filt[filt["Ano"]==sa]
    if filt.empty: st.info("Nenhum documento.")
    else:
        for i,d in filt.iterrows():
            with st.expander(f"📄 {d['NomeArquivo']} | {d['Mes']}/{d['Ano']}"):
                with open(d["Caminho"],"rb") as f: st.download_button("⬇️ BAIXAR", f, file_name=d["NomeArquivo"], key=f"d{i}")
                if st.button("🗑️ EXCLUIR", key=f"x{i}"):
                    os.remove(d["Caminho"])
                    dados["Docs_Lojas"].drop(i,inplace=True)
                    salvar_dados(dados)
                    st.rerun()

# ================ ABA 7 - LOJAS E CARGOS ================
with aba7:
    st.subheader("⚙️ CADASTRO DE LOJAS E CARGOS")
    dados = carregar_dados()
    col1, col2 = st.columns(2)
    with col1:
        nova_loja = st.text_input("Nova Loja")
        if st.button("➕ ADICIONAR LOJA", type="primary") and nova_loja.strip():
            if not dados["Auxiliares"]["Loja"].str.strip().eq(nova_loja.strip()).any():
                dados["Auxiliares"] = pd.concat([dados["Auxiliares"], pd.DataFrame([{"Loja": nova_loja.strip(), "Cargo": ""}])], ignore_index=True)
                salvar_dados(dados)
                st.success("✅ Loja cadastrada!")
                st.rerun()
            else: st.warning("⚠️ Já existe!")
    with col2:
        novo_cargo = st.text_input("Novo Cargo")
        if st.button("➕ ADICIONAR CARGO", type="primary") and novo_cargo.strip():
            if not dados["Auxiliares"]["Cargo"].str.strip().eq(novo_cargo.strip()).any():
                dados["Auxiliares"] = pd.concat([dados["Auxiliares"], pd.DataFrame([{"Loja": "", "Cargo": novo_cargo.strip()}])], ignore_index=True)
                salvar_dados(dados)
                st.success("✅ Cargo cadastrado!")
                st.rerun()
            else: st.warning("⚠️ Já existe!")


# ================ ABA 8 - CONTROLE DE DIÁRIAS ================
with aba8:
    st.subheader("💰 CONTROLE DE DIÁRIAS")
    st.info("ℹ️ Pagamento em até 5 dias úteis, via transferência bancária. Não permitido conta de terceiros.")

    # Upload da planilha de diárias
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

    # ---------- CARDS DE RESUMO ----------
    st.markdown("---")
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.metric("👥 Total de Diárias", len(df_diarias))
    with c2:
        try:
            vtotal = df_diarias["VALOR TOTAL"].replace("", "0").astype(float).sum()
        except:
            vtotal = 0
        st.metric("💰 Valor Total", f"R$ {vtotal:,.2f}")
    with c3:
        try:
            vp = df_diarias[df_diarias["SITUACAO"] == "PENDENTE"]["VALOR TOTAL"].replace("", "0").astype(float).sum()
        except:
            vp = 0
        st.metric("⏳ Pendente", f"R$ {vp:,.2f}")
    with c4:
        try:
            vpg = df_diarias[df_diarias["SITUACAO"] == "PAGO"]["VALOR TOTAL"].replace("", "0").astype(float).sum()
        except:
            vpg = 0
        st.metric("✅ Pago", f"R$ {vpg:,.2f}")
    st.markdown("---")

    # ---------- FILTROS ----------
    col_p1, col_p2, col_p3, col_p4, col_p5 = st.columns(5)
    with col_p1: filtro_loja_d = st.selectbox("Loja", ["Todas"] + lista_lojas(), key="filtro_loja_d")
    with col_p2: filtro_mes_d = st.selectbox("Mês", MESES, key="filtro_mes_d")
    with col_p3: filtro_sem_d = st.selectbox("Semana", SEMANAS, key="filtro_sem_d")
    with col_p4: filtro_ano_d = st.selectbox("Ano", ANOS, index=ANOS.index(str(datetime.now().year)), key="filtro_ano_d")
    with col_p5: filtro_sit_d = st.selectbox("Situação", SITUACOES_DIARIA, key="filtro_sit_d")
    busca_d = st.text_input("🔍 Pesquisar por Nome ou CPF", placeholder="Digite para buscar...")

    df_filtrado = df_diarias.copy()
    if filtro_loja_d != "Todas":
        df_filtrado = df_filtrado[df_filtrado["LOJA"] == filtro_loja_d]
    if filtro_mes_d != "Todos":
        df_filtrado = df_filtrado[df_filtrado["MES"] == filtro_mes_d]
    if filtro_sem_d != "Todas":
        df_filtrado = df_filtrado[df_filtrado["SEMANA"] == filtro_sem_d]
    if filtro_ano_d != "Todos":
        df_filtrado = df_filtrado[df_filtrado["ANO"] == filtro_ano_d]
    if filtro_sit_d != "Todas":
        df_filtrado = df_filtrado[df_filtrado["SITUACAO"] == filtro_sit_d]
    if busca_d.strip():
        df_filtrado = df_filtrado[
            df_filtrado["NOME COLABORADOR"].str.contains(busca_d, case=False, na=False) |
            df_filtrado["CPF"].str.contains(busca_d, case=False, na=False)
        ]

    # ---------- EDITOR INLINE ----------
    if not df_filtrado.empty:
        st.markdown("**📝 Edite os dados diretamente na tabela abaixo e clique em SALVAR ALTERAÇÕES**")
        # Guarda os índices originais para salvar corretamente no DataFrame principal
        idx_original = df_filtrado.index.tolist()
        df_editable = df_filtrado.reset_index(drop=True)

        # Configurar colunas editáveis
        col_config = {
            "LOJA": st.column_config.SelectboxColumn("LOJA", options=lista_lojas(), required=True),
            "MES": st.column_config.SelectboxColumn("MÊS", options=MESES[1:], required=True),
            "SEMANA": st.column_config.SelectboxColumn("SEMANA", options=SEMANAS[1:], required=True),
            "ANO": st.column_config.SelectboxColumn("ANO", options=ANOS, required=True),
            "NOME COLABORADOR": st.column_config.TextColumn("NOME COLABORADOR", required=True),
            "CPF": st.column_config.TextColumn("CPF", required=True),
            "CARGO": st.column_config.TextColumn("CARGO"),
            "DADOS BANCÁRIOS": st.column_config.TextColumn("DADOS BANCÁRIOS"),
            "DATA EXECUCAO": st.column_config.TextColumn("DATA EXECUÇÃO"),
            "DATA PAGAMENTO": st.column_config.TextColumn("DATA PAGAMENTO"),
            "MOTIVO": st.column_config.TextColumn("MOTIVO", required=True),
            "QTDE DE DIARIAS": st.column_config.NumberColumn("QTDE", min_value=1, max_value=30, step=1, required=True),
            "VALOR UNITARIO": st.column_config.NumberColumn("VALOR UNI. (R$)", min_value=0.0, step=0.01, format="%.2f", required=True),
            "SITUACAO": st.column_config.SelectboxColumn("SITUAÇÃO", options=["PENDENTE", "PAGO"], required=True),
            "COMPROVANTE": st.column_config.TextColumn("COMPROVANTE", disabled=True),
            "DATA CADASTRO": st.column_config.TextColumn("DATA CADASTRO", disabled=True),
        }

        edited_df = st.data_editor(
            df_editable,
            column_config=col_config,
            use_container_width=True,
            hide_index=True,
            num_rows="dynamic",
            key="editor_diarias"
        )

        # Calcular VALOR TOTAL automaticamente
        try:
            edited_df["VALOR TOTAL"] = (edited_df["QTDE DE DIARIAS"].astype(float) * edited_df["VALOR UNITARIO"].astype(float)).apply(lambda x: f"{x:.2f}")
        except:
            pass

        col_salvar, col_excluir = st.columns([1, 1])
        with col_salvar:
            if st.button("💾 SALVAR ALTERAÇÕES", type="primary", key="salvar_diarias_editor"):
                for i, idx_orig in enumerate(idx_original):
                    if i < len(edited_df):
                        for col in df_diarias.columns:
                            if col in edited_df.columns:
                                df_diarias.at[idx_orig, col] = str(edited_df.iloc[i][col]) if col not in ["QTDE DE DIARIAS", "VALOR UNITARIO", "VALOR TOTAL"] else str(edited_df.iloc[i][col])
                # Se houver linhas novas (mais que o original)
                if len(edited_df) > len(idx_original):
                    for i in range(len(idx_original), len(edited_df)):
                        nova_linha = {col: "" for col in df_diarias.columns}
                        for col in edited_df.columns:
                            nova_linha[col] = str(edited_df.iloc[i][col])
                        nova_linha["DATA CADASTRO"] = datetime.now().strftime("%d/%m/%Y %H:%M")
                        if not nova_linha.get("VALOR TOTAL"):
                            try:
                                q = float(edited_df.iloc[i]["QTDE DE DIARIAS"])
                                v = float(edited_df.iloc[i]["VALOR UNITARIO"])
                                nova_linha["VALOR TOTAL"] = f"{q * v:.2f}"
                            except:
                                nova_linha["VALOR TOTAL"] = ""
                        df_diarias = pd.concat([df_diarias, pd.DataFrame([nova_linha])], ignore_index=True)
                # Se houver linhas removidas
                if len(edited_df) < len(idx_original):
                    remover = idx_original[len(edited_df):]
                    for idx_rm in remover:
                        comp = str(df_diarias.at[idx_rm, "COMPROVANTE"])
                        if comp and os.path.exists(comp):
                            os.remove(comp)
                    df_diarias.drop(index=remover, inplace=True)
                    df_diarias.reset_index(drop=True, inplace=True)
                salvar_diarias(df_diarias)
                st.success("✅ Alterações salvas com sucesso!")
                st.rerun()

        with col_excluir:
            st.markdown("**🗑️ Excluir linhas selecionadas:** marque a caixa na primeira coluna da tabela acima e depois clique abaixo.")
            if st.button("🗑️ EXCLUIR SELECIONADOS", key="excluir_diarias_editor"):
                st.info("Para excluir, delete as linhas diretamente na tabela usando a tecla Delete ou botão de lixeira do editor, depois clique em SALVAR ALTERAÇÕES.")

    else:
        st.info("Nenhuma diária encontrada com os filtros aplicados.")

    # ---------- NOVA DIÁRIA (CADASTRO RÁPIDO) ----------
    st.markdown("---")
    st.subheader("➕ CADASTRAR NOVA DIÁRIA")
    with st.form("nova_diaria", clear_on_submit=True):
        c1, c2, c3 = st.columns(3)
        with c1:
            loja_d = st.selectbox("Loja *", lista_lojas(), key="nova_loja_d")
            mes_d = st.selectbox("Mês *", MESES[1:], key="nova_mes_d")
            semana_d = st.selectbox("Semana *", SEMANAS[1:], key="nova_sem_d")
            ano_d = st.selectbox("Ano *", ANOS, index=ANOS.index(str(datetime.now().year)), key="nova_ano_d")
        with c2:
            nome_d = st.text_input("Nome do Colaborador *", key="nova_nome_d")
            cpf_d = st.text_input("CPF *", key="nova_cpf_d")
            cargo_d = st.text_input("Cargo", key="nova_cargo_d")
            dados_bancarios_d = st.text_input("Dados Bancários (PIX / Banco / Ag / CC)", key="nova_dados_bancarios_d")
            data_exec_d = st.text_input("Data da Execução (DD/MM/AAAA)", key="nova_data_exec_d")
        with c3:
            data_pag_d = st.text_input("Data de Pagamento (DD/MM/AAAA)", key="nova_data_pag_d")
            motivo_d = st.text_input("Motivo *", key="nova_motivo_d")
            qtde_d = st.number_input("Qtde de Diárias *", min_value=1, max_value=30, value=1, key="nova_qtde_d")
            valor_d = st.number_input("Valor Unitário (R$) *", min_value=0.0, format="%.2f", key="nova_valor_d")
            situacao_d = st.selectbox("Situação *", ["PENDENTE", "PAGO"], key="nova_sit_d")
        submitted = st.form_submit_button("💾 SALVAR DIÁRIA", type="primary")
        if submitted:
            erros = []
            if not loja_d.strip(): erros.append("Loja")
            if not mes_d.strip(): erros.append("Mês")
            if not semana_d.strip(): erros.append("Semana")
            if not ano_d.strip(): erros.append("Ano")
            if not nome_d.strip(): erros.append("Nome do Colaborador")
            if not cpf_d.strip(): erros.append("CPF")
            if not motivo_d.strip(): erros.append("Motivo")
            if qtde_d <= 0: erros.append("Qtde deve ser > 0")
            if valor_d <= 0: erros.append("Valor deve ser > 0")
            if erros:
                st.error("❌ Campos obrigatórios: " + ", ".join(erros))
            else:
                valor_total = qtde_d * valor_d
                nova_linha = {
                    "LOJA": loja_d,
                    "MES": mes_d,
                    "SEMANA": semana_d,
                    "ANO": ano_d,
                    "NOME COLABORADOR": nome_d.strip().upper(),
                    "CPF": cpf_d.strip(),
                    "CARGO": cargo_d.strip().upper(),
                    "DADOS BANCÁRIOS": dados_bancarios_d.strip().upper(),
                    "DATA EXECUCAO": data_exec_d.strip(),
                    "DATA PAGAMENTO": data_pag_d.strip(),
                    "MOTIVO": motivo_d.strip().upper(),
                    "QTDE DE DIARIAS": str(qtde_d),
                    "VALOR UNITARIO": f"{valor_d:.2f}",
                    "VALOR TOTAL": f"{valor_total:.2f}",
                    "SITUACAO": situacao_d,
                    "DATA CADASTRO": datetime.now().strftime("%d/%m/%Y %H:%M"),
                    "COMPROVANTE": ""
                }
                df_diarias = pd.concat([df_diarias, pd.DataFrame([nova_linha])], ignore_index=True)
                salvar_diarias(df_diarias)
                st.success("✅ Diária cadastrada com sucesso!")
                st.rerun()

    # ---------- GERENCIAR COMPROVANTES ----------
    st.markdown("---")
    st.subheader("📎 GERENCIAR COMPROVANTES DE PAGAMENTO")
    if not df_diarias.empty:
        # Dropdown para selecionar diária
        opcoes_diaria = [f"[{i}] {row['NOME COLABORADOR']} | {row['LOJA']} | {row['MES']}/{row['ANO']} | R$ {row['VALOR TOTAL']}" for i, row in df_diarias.iterrows()]
        sel_diaria = st.selectbox("Selecione a diária", options=range(len(opcoes_diaria)), format_func=lambda x: opcoes_diaria[x], key="sel_comp_diaria")
        if sel_diaria is not None:
            idx_comp = df_diarias.index[sel_diaria]
            comp_atual = str(df_diarias.at[idx_comp, "COMPROVANTE"])
            if comp_atual and os.path.exists(comp_atual):
                st.success(f"✅ Comprovante anexado: {os.path.basename(comp_atual)}")
                with open(comp_atual, "rb") as fc:
                    st.download_button("⬇️ Baixar Comprovante", fc, file_name=os.path.basename(comp_atual), key=f"dl_comp_{idx_comp}")
                if st.button("🗑️ Remover Comprovante", key=f"rm_comp_{idx_comp}"):
                    os.remove(comp_atual)
                    df_diarias.at[idx_comp, "COMPROVANTE"] = ""
                    salvar_diarias(df_diarias)
                    st.success("Comprovante removido!")
                    st.rerun()
            else:
                st.info("Nenhum comprovante anexado para esta diária.")
            arq_comp = st.file_uploader("Anexar comprovante (PDF, JPG, PNG)", type=["pdf", "jpg", "png"], key=f"up_comp_{idx_comp}")
            if arq_comp and st.button("📤 ENVIAR COMPROVANTE", type="primary", key=f"btn_comp_{idx_comp}"):
                ext = os.path.splitext(arq_comp.name)[1]
                nome_comp = f"{df_diarias.at[idx_comp, 'CPF']}_{df_diarias.at[idx_comp, 'MES']}_{df_diarias.at[idx_comp, 'ANO']}_{datetime.now().strftime('%Y%m%d%H%M%S')}{ext}"
                cam_comp = os.path.join(PASTA_COMPROVANTES, nome_comp)
                with open(cam_comp, "wb") as f: f.write(arq_comp.read())
                df_diarias.at[idx_comp, "COMPROVANTE"] = cam_comp
                salvar_diarias(df_diarias)
                st.success("✅ Comprovante anexado!")
                st.rerun()
    else:
        st.info("Nenhuma diária cadastrada.")

    # ---------- EXPORTAR ----------
    st.markdown("---")
    st.subheader("📤 EXPORTAR PARA EXCEL")
    if not df_filtrado.empty:
        nome_arq = f"Diarias_{filtro_loja_d}_{filtro_mes_d}_{filtro_ano_d}.xlsx".replace("/", "-").replace(" ", "_")
        exportar_diarias_formatado(df_filtrado, nome_arq)
        with open(nome_arq, "rb") as f:
            st.download_button("⬇️ BAIXAR EXCEL", f, file_name=nome_arq)
        os.remove(nome_arq)
    else:
        st.info("Filtre os dados para exportar.")
