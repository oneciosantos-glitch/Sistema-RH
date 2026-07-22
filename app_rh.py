# ================ ABA 8 - CONTROLE DE DIÁRIAS (COMPLETA E CORRIGIDA) ================
with aba8:
    st.subheader("💰 CONTROLE DE DIÁRIAS")
    st.info("ℹ️ Pagamento em até 5 dias úteis, via transferência bancária, não permitido conta de terceiros.")

    # Filtros de pesquisa
    col_p1, col_p2, col_p3, col_p4, col_p5 = st.columns(5)
    with col_p1: filtro_loja_d = st.selectbox("Loja", ["Todas"] + lista_lojas(), key="filtro_loja_d")
    with col_p2: filtro_mes_d = st.selectbox("Mês", MESES, key="filtro_mes_d")
    with col_p3: filtro_sem_d = st.selectbox("Semana", SEMANAS, key="filtro_sem_d")
    with col_p4: filtro_ano_d = st.selectbox("Ano", ANOS, index=ANOS.index(str(datetime.now().year)), key="filtro_ano_d")
    with col_p5: filtro_sit_d = st.selectbox("Situação", SITUACOES_DIARIA, key="filtro_sit_d")
    
    busca_d = st.text_input("🔍 Pesquisar por Nome ou CPF", placeholder="Digite para buscar...")

    # Carrega e filtra dados
    df_diarias = carregar_diarias()
    df_filtrado = df_diarias.copy()

    if filtro_loja_d != "Todas": df_filtrado = df_filtrado[df_filtrado["LOJA"] == filtro_loja_d]
    if filtro_mes_d != "Todos": df_filtrado = df_filtrado[df_filtrado["MES"] == filtro_mes_d]
    if filtro_sem_d != "Todas": df_filtrado = df_filtrado[df_filtrado["SEMANA"] == filtro_sem_d]
    if filtro_ano_d != "Todos": df_filtrado = df_filtrado[df_filtrado["ANO"] == filtro_ano_d]
    if filtro_sit_d != "Todas": df_filtrado = df_filtrado[df_filtrado["SITUACAO"] == filtro_sit_d]
    if busca_d.strip():
        df_filtrado = df_filtrado[
            df_filtrado["NOME COLABORADOR"].str.contains(busca_d, case=False, na=False) |
            df_filtrado["CPF"].str.contains(busca_d, case=False, na=False)
        ]

    # Mostra tabela
    st.dataframe(
        df_filtrado[["LOJA","NOME COLABORADOR","CPF","DATA EXECUCAO","QUANTIDADE","VALOR UNITARIO","TOTAL","SITUACAO","MES","SEMANA","ANO"]],
        use_container_width=True, hide_index=True
    )

    # Selecionar para editar
    indice_sel = st.text_input("✏️ Digite o ÍNDICE da linha para editar/excluir", placeholder="Número da linha na tabela")
    reg_d = pd.DataFrame()
    # ✅ LINHA CORRIGIDA ABAIXO
    if indice_sel.strip() and indice_sel.isdigit():
        idx = int(indice_sel)
        if 0 <= idx < len(df_diarias):
            reg_d = df_diarias.iloc[idx]

    # Formulário de cadastro/edição
    with st.form("form_diarias", clear_on_submit=True):
        st.subheader("Cadastro / Edição de Diárias")
        col1, col2 = st.columns(2)
        with col1:
            loja_d = st.selectbox("🏬 Loja", lista_lojas(), index=lista_lojas().index(reg_d["LOJA"]) if not reg_d.empty and reg_d["LOJA"] in lista_lojas() else 0)
            nome_d = st.text_input("👤 Nome Colaborador", value=reg_d.get("NOME COLABORADOR", ""))
            cpf_d = st.text_input("📄 CPF", value=reg_d.get("CPF", ""))
            data_d = st.text_input("📅 Data Execução (dd/mm/aaaa)", value=reg_d.get("DATA EXECUCAO", datetime.now().strftime("%d/%m/%Y")))
            qtd_d = st.text_input("🔢 Quantidade", value=reg_d.get("QUANTIDADE", "1"))
            valor_d = st.text_input("💵 Valor Unitário", value=reg_d.get("VALOR UNITARIO", ""))
            total_d = st.text_input("💸 TOTAL", value=reg_d.get("TOTAL", ""), disabled=True)
        with col2:
            dados_banc_d = st.text_input("🏦 Dados Bancários", value=reg_d.get("DADOS BANCARIOS", ""))
            subst_d = st.text_input("🔄 Substituição", value=reg_d.get("SUBSTITUICAO", ""))
            motivo_d = st.text_input("📝 Motivo", value=reg_d.get("MOTIVO", ""))
            data_pag_d = st.text_input("📅 Data Pagamento", value=reg_d.get("DATA PAGAMENTO", ""))
            sit_d = st.selectbox("📊 Situação", SITUACOES_DIARIA[1:], index=SITUACOES_DIARIA.index(reg_d["SITUACAO"])-1 if not reg_d.empty and reg_d["SITUACAO"] in SITUACOES_DIARIA else 0)
            mes_d = st.selectbox("📆 Mês", MESES[1:], index=MESES.index(reg_d["MES"])-1 if not reg_d.empty and reg_d["MES"] in MESES else MESES.index(datetime.now().strftime("%b").lower()))
            sem_d = st.selectbox("📆 Semana", SEMANAS[1:], index=SEMANAS.index(reg_d["SEMANA"])-1 if not reg_d.empty and reg_d["SEMANA"] in SEMANAS else 0)
            ano_d = st.selectbox("📆 Ano", ANOS, index=ANOS.index(reg_d["ANO"]) if not reg_d.empty and reg_d["ANO"] in ANOS else ANOS.index(str(datetime.now().year)))

        comprovante = st.file_uploader("📎 Anexar Comprovante", type=["pdf","jpg","jpeg","png"], key="comp_d")

        col_btn1, col_btn2 = st.columns(2)
        with col_btn1:
            btn_salvar_d = st.form_submit_button("💾 SALVAR DIÁRIA", type="primary", use_container_width=True)
        with col_btn2:
            btn_limpar_d = st.form_submit_button("🗑️ LIMPAR FORMULÁRIO", use_container_width=True)

        if btn_salvar_d:
            if not nome_d.strip() or not cpf_d.strip():
                st.error("❌ Preencha Nome e CPF!")
                st.stop()

            cam_comp = reg_d.get("CAMINHO_COMPROVANTE", "")
            if comprovante:
                if cam_comp and os.path.exists(cam_comp):
                    os.remove(cam_comp)
                nome_arq_comp = f"comp_diaria_{nome_d.replace(' ','_')}_{datetime.now().strftime('%Y%m%d%H%M%S')}{os.path.splitext(comprovante.name)[1].lower()}"
                cam_comp = os.path.join(PASTA_COMPROVANTES, nome_arq_comp)
                with open(cam_comp, "wb") as f:
                    f.write(comprovante.read())

            novo_registro = {
                "LOJA": loja_d, "NOME COLABORADOR": nome_d, "CPF": cpf_d,
                "DATA EXECUCAO": data_d, "QUANTIDADE": qtd_d, "VALOR UNITARIO": valor_d,
                "TOTAL": total_d, "DADOS BANCARIOS": dados_banc_d, "SUBSTITUICAO": subst_d,
                "MOTIVO": motivo_d, "DATA PAGAMENTO": data_pag_d, "SITUACAO": sit_d,
                "MES": mes_d, "SEMANA": sem_d, "ANO": ano_d,
                "COMPROVANTE": comprovante.name if comprovante else reg_d.get("COMPROVANTE", ""),
                "CAMINHO_COMPROVANTE": cam_comp
            }

            if not reg_d.empty:
                df_diarias.iloc[idx] = novo_registro
                st.success("✅ Diária ATUALIZADA com sucesso!")
            else:
                df_diarias = pd.concat([df_diarias, pd.DataFrame([novo_registro])], ignore_index=True)
                st.success("✅ Diária CADASTRADA com sucesso!")
            
            salvar_diarias(df_diarias)
            st.rerun()

    if indice_sel.strip() and indice_sel.isdigit() and st.button("🗑️ EXCLUIR DIÁRIA", use_container_width=True, type="secondary"):
        if st.checkbox("⚠️ CONFIRMA EXCLUSÃO PERMANENTE?"):
            idx_exc = int(indice_sel)
            if 0 <= idx_exc < len(df_diarias):
                cam_exc = df_diarias.iloc[idx_exc].get("CAMINHO_COMPROVANTE", "")
                if cam_exc and os.path.exists(cam_exc):
                    os.remove(cam_exc)
                df_diarias.drop(idx_exc, inplace=True)
                df_diarias.reset_index(drop=True, inplace=True)
                salvar_diarias(df_diarias)
                st.success("✅ Diária excluída!")
                st.rerun()
