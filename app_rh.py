import streamlit as st
import re
try:
    import matplotlib.pyplot as plt
    MATPLOT = True
except ImportError:
    MATPLOT = False
import pandas as pd
import os
import shutil
import time
import io
import json
import requests
from datetime import datetime, timedelta
from PIL import Image
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ====================== GOOGLE SHEETS INTEGRAÇÃO ======================
# Verifica se as credenciais do Google Sheets estão configuradas
GS_ENABLED = False
gc = None
GS_ID_FUNCIONARIOS = None
GS_ID_DIARIAS = None
GS_ID_COMPRAS = None

try:
    import gspread
    from google.oauth2.service_account import Credentials
    
    if "gspread" in st.secrets:
        creds_dict = dict(st.secrets["gspread"])
        creds = Credentials.from_service_account_info(creds_dict, scopes=[
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ])
        gc = gspread.authorize(creds)
        GS_ID_FUNCIONARIOS = st.secrets.get("gsheets", {}).get("id_funcionarios", "")
        GS_ID_DIARIAS = st.secrets.get("gsheets", {}).get("id_diarias", "")
        GS_ID_COMPRAS = st.secrets.get("gsheets", {}).get("id_compras", "")
        if GS_ID_FUNCIONARIOS or GS_ID_DIARIAS or GS_ID_COMPRAS:
            GS_ENABLED = True
except Exception:
    pass


# ====================== MÓDULO DE COMPRAS EMBUTIDO ======================
# -*- coding: utf-8 -*-
"""Módulo Sistema de Compras e Entregas - Streamlit"""
import streamlit as st
import pandas as pd
import uuid
from datetime import datetime, date
import io
import base64

# ========== DADOS ==========
CLIENTES = ["Smart Fit", "Self Fit", "Assaí Atacadista"]

LOJAS_POR_CLIENTE = {
    "Smart Fit": [
        "Smart Fit Shopping Manoa", "Smart Fit Shopping Cidade Leste", "Smart Fit Macapá Shopping",
        "Smart Fit Shopping Grande Circular", "Smart Fit Shopping Via Norte", "Smart Fit Cidade Nova",
        "Smart Fit Parque Mosaico", "Smart Fit Cachoeirinha", "Smart Fit Flores", "Smart Fit Ponta Negra",
        "Smart Fit Nova Porto Velho", "Smart Fit Porto Velho Flodoaldo", "Smart Fit Alvorada",
        "Smart Fit Novo Aleixo", "Smart Fit São José do Operário", "Smart Fit Santana Macapá",
        "Smart Fit Toequato Tapajós"
    ],
    "Self Fit": [
        "Self Fit Hiper DB Ponta Negra", "Self Fit Manaus Plaza Shopping", "Self Fit Vieira Alves"
    ],
    "Assaí Atacadista": [
        "Assaí Atacadista Batista Campos", "Assaí Atacadista Almirante Barroso", "Assaí Atacadista Castanhal",
        "Assaí Atacadista Ananindeua", "Assaí Atacadista Augusto Monte Negro", "Assaí Atacadista Boa Vista",
        "Assaí Atacadista Manaus", "Assaí Atacadista Macapá", "Assaí Atacadista Belém"
    ]
}

MATERIAIS_POR_CLIENTE = {
    "Smart Fit": [
        "ÁGUA SANITÁRIA", "ASPIRADOR SEMI-INDUSTRIAL 23L", "BALDE 15L", "BALDE 6L",
        "BALDE ESPREMEDOR COMPLETO", "CABO DE ALUMINIO SEM ROSCA", "DISCO VERMELHO 510",
        "ENCERADEIRA INDUSTRIAL", "ESCOVA DE MÃO", "ESCOVA SANITÁRIA", "ESPONJA DUPLA FACE",
        "EXTENSÃO DE 30 M", "FIBRAS DE LIMPEZA PESADA", "FLANELAS", "KIT LIMPA VIDRO 2 EM 1 BRALIMPIA",
        "LIMPA TUDO (MINI LOK)", "PÁ DE LIXO COMPLETA", "PANO DE CHÃO ALVEJADO", "PLACAS SINALIZADORA",
        "REFIL MOP ÁGUA", "RODO DE 60cm COMPLETO", "SABÃO EM PÓ", "SACO DE LIXO 100 L",
        "SACO DE LIXO 40 L", "SACO DE LIXO 60 L", "VASSOURA DE NYLON COMPLETA", "VASSOURA DE TETO COMPLETA"
    ],
    "Self Fit": [
        "ASPIRADOR DE PÓ E BATERIA SEM FIO", "BAUDE EXPREMEDOR", "REFIL MOP ÁGUA", "REFIL MOP PÓ",
        "PLACAS SINALIZADORA", "ENCERADEIRA INDUSTRIAL", "KIT LIMPA VIDRO 2 EM 1 BRALIMPIA",
        "PÁ COLETORA DE LIXO", "PULVERIZADOR", "CESTA MULTIUSO"
    ],
    "Assaí Atacadista": [
        "Disco 510 mm - Preto", "Disco 510 mm - Marron para Remoção", "Disco 510 mm - vermelho",
        "Disco 510 mm - Verde", "Disco pelo de porco para Polidora", "Disco champanhe para Polidora",
        "Starlock frange 510mm aço", "Starlock frange 510mm com Velcon", "Starlock frange 510mm escova",
        "Enceradeira industrial 510", "Armação Mop Pó 60 cm", "Armação Mop Pó 1,20 cm", "Refil mop cera",
        "Suporte mop cera", "Lt", "Esponja p/LT", "Rodo madeira 1,20mts", "Rodo 60 cm",
        "Cabeleira Mop Agua", "Mop Pó 60 cm", "Mop Pó 1,20 cm", "Saco Amarelo P/Carrinho",
        "Pa coletora azul pop", "Raspador pesado sem cabo", "Raspador pesado com cabo", "Garra mop Agua",
        "Carro coletor de lixo 240lts", "Cabo de aluminio com rosca", "Cabo de aluminio sem rosca",
        "Lâminas p/ raspdor", "Raspador de mão", "Extenção cabo PP 3x2,5", "Borracha Organizadora Carrinho Funcional",
        "Carrinho funcional kit completo", "Balde expremedor", "Kit manunteção para carrinho", "Vassoura nylon",
        "Vassoura Piassava", "Vassourão gari 60 cm", "Alongador 9mts", "Regador"
    ]
}

EPIS_POR_CLIENTE = {
    "Smart Fit": [
        "Luva látex", "Óculos de proteção", "Luva de Vinil", "Máscara de Proteção",
        "Protetor auricular plug", "Protetor tipo concha", "Luva para jardineiro", "Avental de raspa",
        "Viseira", "Perneira", "Meia Térmica", "Japonha Térmica", "Calça Térmica", "Luvas térmicas",
        "Capuz Térmico", "Avental Térmico", "Bota C.Médio Nº34", "Bota C.Médio Nº35", "Bota C.Médio Nº36",
        "Bota C.Médio Nº37", "Bota C.Médio Nº38", "Bota C.Médio Nº39", "Bota C.Médio Nº40",
        "Bota C.Médio Nº41", "Bota C.Médio Nº42", "Bota C.Médio Nº43", "Bota C.Médio Nº44",
        "Bota C.Médio Nº45", "Bota C.Médio Nº46", "Bota de Couro Nº34", "Bota de Couro Nº35",
        "Bota de Couro Nº36", "Bota de Couro Nº37", "Bota de Couro Nº38", "Bota de Couro Nº39",
        "Bota de Couro Nº40", "Bota de Couro Nº41", "Bota de Couro Nº42", "Bota de Couro Nº43",
        "Bota de Couro Nº44", "Bota de Couro Nº45", "Bota de Couro Nº46", "Sapato Ant-derrapante Nº34",
        "Sapato Ant-derrapante Nº35", "Sapato Ant-derrapante Nº36", "Sapato Ant-derrapante Nº37",
        "Sapato Ant-derrapante Nº38", "Sapato Ant-derrapante Nº39", "Sapato Ant-derrapante Nº40",
        "Sapato Ant-derrapante Nº41", "Sapato Ant-derrapante Nº42", "Sapato Ant-derrapante Nº43",
        "Sapato Ant-derrapante Nº44", "Sapato Ant-derrapante Nº45", "Sapato Ant-derrapante Nº46",
        "Farda C.Feminino (P)", "Farda C.Feminino (M)", "Farda C.Feminino (G)", "Farda C.Feminino (GG)",
        "Farda C.Feminino (XG)", "Farda C.Masculino (P)", "Farda C.Masculino (M)", "Farda C.Masculino (G)",
        "Farda C.Masculino (GG)", "Farda C.Masculino (XG)", "Farda p/ Jardineiro", "Farda p/ Encarregado & Líder",
        "Farda p/ Supervisor", "Camisa Branca", "Cauça", "Chapéu"
    ],
    "Self Fit": [
        "Luva látex", "Óculos de proteção", "Luva de Vinil", "Máscara de Proteção",
        "Protetor auricular plug", "Protetor tipo concha", "Luva para jardineiro", "Avental de raspa",
        "Viseira", "Perneira", "Meia Térmica", "Japonha Térmica", "Calça Térmica", "Luvas térmicas",
        "Capuz Térmico", "Avental Térmico", "Bota C.Médio Nº34", "Bota C.Médio Nº35", "Bota C.Médio Nº36",
        "Bota C.Médio Nº37", "Bota C.Médio Nº38", "Bota C.Médio Nº39", "Bota C.Médio Nº40",
        "Bota C.Médio Nº41", "Bota C.Médio Nº42", "Bota C.Médio Nº43", "Bota C.Médio Nº44",
        "Bota C.Médio Nº45", "Bota C.Médio Nº46", "Bota de Couro Nº34", "Bota de Couro Nº35",
        "Bota de Couro Nº36", "Bota de Couro Nº37", "Bota de Couro Nº38", "Bota de Couro Nº39",
        "Bota de Couro Nº40", "Bota de Couro Nº41", "Bota de Couro Nº42", "Bota de Couro Nº43",
        "Bota de Couro Nº44", "Bota de Couro Nº45", "Bota de Couro Nº46", "Sapato Ant-derrapante Nº34",
        "Sapato Ant-derrapante Nº35", "Sapato Ant-derrapante Nº36", "Sapato Ant-derrapante Nº37",
        "Sapato Ant-derrapante Nº38", "Sapato Ant-derrapante Nº39", "Sapato Ant-derrapante Nº40",
        "Sapato Ant-derrapante Nº41", "Sapato Ant-derrapante Nº42", "Sapato Ant-derrapante Nº43",
        "Sapato Ant-derrapante Nº44", "Sapato Ant-derrapante Nº45", "Sapato Ant-derrapante Nº46",
        "Farda C.Feminino (P)", "Farda C.Feminino (M)", "Farda C.Feminino (G)", "Farda C.Feminino (GG)",
        "Farda C.Feminino (XG)", "Farda C.Masculino (P)", "Farda C.Masculino (M)", "Farda C.Masculino (G)",
        "Farda C.Masculino (GG)", "Farda C.Masculino (XG)", "Farda p/ Jardineiro", "Farda p/ Encarregado & Líder",
        "Farda p/ Supervisor", "Camisa Branca", "Cauça", "Chapéu"
    ],
    "Assaí Atacadista": [
        "Luva látex", "Óculos de proteção", "Luva de Vinil", "Máscara de Proteção",
        "Protetor auricular plug", "Protetor tipo concha", "Luva para jardineiro", "Avental de raspa",
        "Viseira", "Perneira", "Meia Térmica", "Japonha Térmica", "Calça Térmica", "Luvas térmicas",
        "Capuz Térmico", "Avental Térmico", "Bota C.Médio Nº34", "Bota C.Médio Nº35", "Bota C.Médio Nº36",
        "Bota C.Médio Nº37", "Bota C.Médio Nº38", "Bota C.Médio Nº39", "Bota C.Médio Nº40",
        "Bota C.Médio Nº41", "Bota C.Médio Nº42", "Bota C.Médio Nº43", "Bota C.Médio Nº44",
        "Bota C.Médio Nº45", "Bota C.Médio Nº46", "Bota de Couro Nº34", "Bota de Couro Nº35",
        "Bota de Couro Nº36", "Bota de Couro Nº37", "Bota de Couro Nº38", "Bota de Couro Nº39",
        "Bota de Couro Nº40", "Bota de Couro Nº41", "Bota de Couro Nº42", "Bota de Couro Nº43",
        "Bota de Couro Nº44", "Bota de Couro Nº45", "Bota de Couro Nº46", "Sapato Ant-derrapante Nº34",
        "Sapato Ant-derrapante Nº35", "Sapato Ant-derrapante Nº36", "Sapato Ant-derrapante Nº37",
        "Sapato Ant-derrapante Nº38", "Sapato Ant-derrapante Nº39", "Sapato Ant-derrapante Nº40",
        "Sapato Ant-derrapante Nº41", "Sapato Ant-derrapante Nº42", "Sapato Ant-derrapante Nº43",
        "Sapato Ant-derrapante Nº44", "Sapato Ant-derrapante Nº45", "Sapato Ant-derrapante Nº46",
        "Farda C.Feminino (P)", "Farda C.Feminino (M)", "Farda C.Feminino (G)", "Farda C.Feminino (GG)",
        "Farda C.Feminino (XG)", "Farda C.Masculino (P)", "Farda C.Masculino (M)", "Farda C.Masculino (G)",
        "Farda C.Masculino (GG)", "Farda C.Masculino (XG)", "Farda p/ Jardineiro", "Farda p/ Encarregado & Líder",
        "Farda p/ Supervisor", "Camisa Branca", "Cauça", "Chapéu"
    ]
}

STATUS_OPCOES = ["Pendente", "Aprovado", "Em Trânsito", "Entregue", "Cancelado"]
TIPOS_SOLICITACAO = ["Material", "EPI"]
PRIORIDADES = ["Normal", "Urgente", "Baixa"]

TAMANHOS_EPI = ["", "P", "M", "G", "GG", "37", "38", "39", "40", "41", "42", "43", "44"]


def gerar_id():
    return "SOL-" + uuid.uuid4().hex[:8].upper()


def formatar_data_br(d):
    if not d:
        return "-"
    if isinstance(d, str):
        if "T" in d:
            d = d.split("T")[0]
        try:
            y, m, day = d.split("-")
            return f"{day}/{m}/{y}"
        except Exception:
            return d
    if isinstance(d, (datetime, date)):
        return d.strftime("%d/%m/%Y")
    return str(d)


def formatar_moeda(v):
    if not v:
        return "R$ 0,00"
    try:
        return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(v)


def init_session_state():
    if "compras_solicitacoes" not in st.session_state:
        sols, ents = [], []
        # Tenta carregar do Google Sheets
        if GS_ENABLED and GS_ID_COMPRAS:
            try:
                sols, ents = _carregar_compras_gs()
            except Exception:
                sols, ents = [], []
        # Se não conseguiu do GS, tenta carregar do arquivo local
        if not sols and not ents:
            sols, ents = _carregar_compras_local()
        st.session_state["compras_solicitacoes"] = sols
        st.session_state["compras_entregas"] = ents
    if "compras_entregas" not in st.session_state:
        st.session_state["compras_entregas"] = []
    if "compras_page" not in st.session_state:
        st.session_state["compras_page"] = "Dashboard"
    if "compras_edit_id" not in st.session_state:
        st.session_state["compras_edit_id"] = None
    if "compras_nova_itens_material" not in st.session_state:
        st.session_state["compras_nova_itens_material"] = []
    if "compras_nova_itens_epi" not in st.session_state:
        st.session_state["compras_nova_itens_epi"] = []
    if "compras_form_reset" not in st.session_state:
        st.session_state["compras_form_reset"] = False
    if "compras_gs_loaded" not in st.session_state:
        st.session_state["compras_gs_loaded"] = True


def switch_page(page):
    st.session_state["compras_page"] = page
    st.session_state["compras_edit_id"] = None
    st.rerun()


def get_badge_color(status):
    return {
        "Pendente": "🔴",
        "Aprovado": "🟢",
        "Em Trânsito": "🔵",
        "Entregue": "🟢",
        "Cancelado": "⚫"
    }.get(status, "⚪")


# ========== GERAÇÃO DE DOCUMENTOS ==========
def gerar_doc_epi(sol):
    data_geracao = datetime.now().strftime("%d/%m/%Y")
    cliente = sol.get("cliente", "_________________")
    loja = sol.get("loja", "_________________")
    map_itens = {i.get("epi", ""): i for i in sol.get("itens", [])}

    epi_list = [
        "Luva látex", "Luva para jardineiro", "Japona Térmica",
        "Óculos de proteção", "Avental de raspa", "Calça Térmica",
        "Máscara de Proteção", "Viseira", "Luvas térmicas",
        "Protetor auricular plug ( ) ou concha ( )", "Perneira", "Cap", "", "", ""
    ]
    epi_rows = []
    for i in range(0, len(epi_list), 3):
        e1 = epi_list[i] if i < len(epi_list) else ""
        e2 = epi_list[i+1] if i+1 < len(epi_list) else ""
        e3 = epi_list[i+2] if i+2 < len(epi_list) else ""
        q1 = map_itens.get(e1, {}).get("qtd", "&nbsp;") if e1 else "&nbsp;"
        q2 = map_itens.get(e2, {}).get("qtd", "&nbsp;") if e2 else "&nbsp;"
        q3 = map_itens.get(e3, {}).get("qtd", "&nbsp;") if e3 else "&nbsp;"
        epi_rows.append(
            f'<tr><td style="border:1px solid #000;padding:3px 5px;font-size:10px">{e1 or "&nbsp;"}</td>'
            f'<td style="border:1px solid #000;padding:3px 5px;font-size:10px;text-align:center">{q1}</td>'
            f'<td style="border:1px solid #000;padding:3px 5px;font-size:10px">{e2 or "&nbsp;"}</td>'
            f'<td style="border:1px solid #000;padding:3px 5px;font-size:10px;text-align:center">{q2}</td>'
            f'<td style="border:1px solid #000;padding:3px 5px;font-size:10px">{e3 or "&nbsp;"}</td>'
            f'<td style="border:1px solid #000;padding:3px 5px;font-size:10px;text-align:center">{q3}</td></tr>'
        )

    itens_arr = [(i.get("epi", i.get("nome", "Item")) + (f' x{i.get("qtd", "")}' if i.get("qtd") else "")) for i in sol.get("itens", [])]
    botas_rows = []
    total_rows = max(17, len(itens_arr))
    for i in range(total_rows):
        item_obj = sol.get("itens", [])[i] if i < len(sol.get("itens", [])) else None
        nome = item_obj.get("colaborador", sol.get("nomeFuncionario", "&nbsp;")) if item_obj else "&nbsp;"
        loja_val = sol.get("loja", "&nbsp;") if item_obj else "&nbsp;"
        enc = sol.get("encarregado", "&nbsp;") if item_obj else "&nbsp;"
        sup = sol.get("supervisor", "&nbsp;") if item_obj else "&nbsp;"
        item = itens_arr[i] if i < len(itens_arr) else "&nbsp;"
        botas_rows.append(
            f'<tr><td style="border:1px solid #000;padding:3px 5px;font-size:10px">{nome}</td>'
            f'<td style="border:1px solid #000;padding:3px 5px;font-size:10px">{loja_val}</td>'
            f'<td style="border:1px solid #000;padding:3px 5px;font-size:10px">{enc}</td>'
            f'<td style="border:1px solid #000;padding:3px 5px;font-size:10px">{sup}</td>'
            f'<td style="border:1px solid #000;padding:3px 5px;font-size:10px">{item}</td>'
            f'<td style="border:1px solid #000;padding:3px 5px;font-size:10px">&nbsp;</td></tr>'
        )

    states = 'ALAGOAS ( &nbsp;) &nbsp;&nbsp;BAHIA ( &nbsp;) &nbsp;&nbsp;CEARÁ ( &nbsp;) &nbsp;&nbsp;MARANHÃO ( &nbsp;) &nbsp;&nbsp;PARAIBA ( &nbsp;) &nbsp;&nbsp;PARÁ ( &nbsp;) &nbsp;&nbsp;PERNAMBUCO ( &nbsp;) &nbsp;&nbsp;PIAUÍ ( &nbsp;) &nbsp;&nbsp;RIO GRANDE DO NORTE ( &nbsp;) &nbsp;&nbsp;SERGIPE ( &nbsp;) &nbsp;&nbsp;AMAPÁ ( &nbsp;) &nbsp;&nbsp;RORAIMA ( &nbsp;) &nbsp;&nbsp;AMAZONAS ( &nbsp;) &nbsp;&nbsp;RONDÔNIA ( &nbsp;)'

    html = f"""<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
<head><meta charset='utf-8'><title>Solicitacao de EPI - {cliente} - {loja}</title>
<style>@page Section1 {{ size: 841.95pt 595.35pt; margin: 36pt 36pt 36pt 36pt; mso-page-orientation: landscape; }}
div.Section1 {{ page: Section1; }}</style></head>
<body style="font-family:Arial,sans-serif;font-size:10pt;margin:0;padding:0">
<div class="Section1" style="mso-page-orientation: landscape;">
<table style="width:100%;border-collapse:collapse">
<tr><td colspan="6" style="border:1px solid #000;padding:3px 5px;font-size:9px;text-align:center">{states}</td></tr>
<tr><td colspan="6" style="border:1px solid #000;padding:3px 5px;font-size:11px;font-weight:bold;background:#DEEBF6;text-align:center">BOTAS</td></tr>
<tr>
<th style="border:1px solid #000;padding:3px 5px;font-size:10px;font-weight:bold;background:#DEEBF6">Nome</th>
<th style="border:1px solid #000;padding:3px 5px;font-size:10px;font-weight:bold;background:#DEEBF6">Loja</th>
<th style="border:1px solid #000;padding:3px 5px;font-size:10px;font-weight:bold;background:#DEEBF6">Encarregado</th>
<th style="border:1px solid #000;padding:3px 5px;font-size:10px;font-weight:bold;background:#DEEBF6">Supervisor</th>
<th style="border:1px solid #000;padding:3px 5px;font-size:10px;font-weight:bold;background:#DEEBF6">Itens da Solicitação de EPI</th>
<th style="border:1px solid #000;padding:3px 5px;font-size:10px;font-weight:bold;background:#DEEBF6">Situação ( * )</th>
</tr>
{''.join(botas_rows)}
<tr><td colspan="6" style="border:1px solid #000;padding:3px 5px;font-size:11px;font-weight:bold;background:#DEEBF6;text-align:center">EPIS</td></tr>
<tr>
<th style="border:1px solid #000;padding:3px 5px;font-size:10px;font-weight:bold;background:#DEEBF6">Equipamento de Proteção</th>
<th style="border:1px solid #000;padding:3px 5px;font-size:10px;font-weight:bold;background:#DEEBF6">Quantidade</th>
<th style="border:1px solid #000;padding:3px 5px;font-size:10px;font-weight:bold;background:#DEEBF6">Equipamento de Proteção</th>
<th style="border:1px solid #000;padding:3px 5px;font-size:10px;font-weight:bold;background:#DEEBF6">Quantidade</th>
<th style="border:1px solid #000;padding:3px 5px;font-size:10px;font-weight:bold;background:#DEEBF6">Equipamento de Proteção</th>
<th style="border:1px solid #000;padding:3px 5px;font-size:10px;font-weight:bold;background:#DEEBF6">Quantidade</th>
</tr>
{''.join(epi_rows)}
<tr><td colspan="6" style="border:1px solid #000;padding:3px 5px;font-size:9px">Observação: A coluna Situação ( * ), é para o setor SST preencher.</td></tr>
</table>
<p style="font-size:8pt;color:#666;text-align:center;margin-top:6px">Data de geração: {data_geracao}</p>
</div></body></html>"""
    return html


def gerar_xls_material(sol):
    data_geracao = datetime.now().strftime("%d/%m/%Y")
    cliente = sol.get("cliente", "")
    itens = sol.get("itens", [])

    def build_rows_smart_self(items):
        if not items:
            return '<tr><td style="border:1px solid #000;padding:4px 6px;font-size:11px">&nbsp;</td><td style="border:1px solid #000;padding:4px 6px;font-size:11px;text-align:center">&nbsp;</td><td style="border:1px solid #000;padding:4px 6px;font-size:11px;text-align:center">UN.</td></tr>'
        return "".join([
            f'<tr><td style="border:1px solid #000;padding:4px 6px;font-size:11px;white-space:nowrap">{i.get("material","")}</td>'
            f'<td style="border:1px solid #000;padding:4px 6px;font-size:11px;text-align:center;white-space:nowrap">{i.get("qtd","&nbsp;")}</td>'
            f'<td style="border:1px solid #000;padding:4px 6px;font-size:11px;text-align:center;white-space:nowrap">UN.</td></tr>'
            for i in items
        ])

    def build_rows_assai(items):
        if not items:
            return '<tr><td style="border:1px solid #000;padding:4px 6px;font-size:11px;text-align:center">1</td><td style="border:1px solid #000;padding:4px 6px;font-size:11px">&nbsp;</td><td style="border:1px solid #000;padding:4px 6px;font-size:11px">&nbsp;</td><td style="border:1px solid #000;padding:4px 6px;font-size:11px;text-align:center">&nbsp;</td><td style="border:1px solid #000;padding:4px 6px;font-size:11px;text-align:center">peças</td><td style="border:1px solid #000;padding:4px 6px;font-size:11px;text-align:right">&nbsp;</td><td style="border:1px solid #000;padding:4px 6px;font-size:11px;text-align:right">&nbsp;</td></tr>'
        rows = []
        for idx, i in enumerate(items, 1):
            qtd = i.get("qtd", "&nbsp;")
            unit = f'R$ {i["valorUnit"]:.2f}'.replace(".", ",") if i.get("valorUnit") else "&nbsp;"
            total = f'R$ {(i["valorUnit"]*i["qtd"]):.2f}'.replace(".", ",") if i.get("valorUnit") and i.get("qtd") else "&nbsp;"
            rows.append(
                f'<tr><td style="border:1px solid #000;padding:4px 6px;font-size:11px;text-align:center;white-space:nowrap">{idx}</td>'
                f'<td style="border:1px solid #000;padding:4px 6px;font-size:11px;white-space:nowrap">{i.get("material","")}</td>'
                f'<td style="border:1px solid #000;padding:4px 6px;font-size:11px;white-space:nowrap">&nbsp;</td>'
                f'<td style="border:1px solid #000;padding:4px 6px;font-size:11px;text-align:center;white-space:nowrap">{qtd}</td>'
                f'<td style="border:1px solid #000;padding:4px 6px;font-size:11px;text-align:center;white-space:nowrap">peças</td>'
                f'<td style="border:1px solid #000;padding:4px 6px;font-size:11px;text-align:right;white-space:nowrap">{unit}</td>'
                f'<td style="border:1px solid #000;padding:4px 6px;font-size:11px;text-align:right;white-space:nowrap">{total}</td></tr>'
            )
        return "".join(rows)

    if cliente == "Smart Fit":
        rows = build_rows_smart_self(itens)
        html = f"""<html xmlns:x="urn:schemas-microsoft-com:office:excel"><head><meta charset="utf-8"><style>table{{border-collapse:collapse;width:100%}}th,td{{border:1px solid #000}}</style></head><body><table>
<tr><th colspan="3" style="border:1px solid #000;padding:6px;font-size:13px;font-weight:bold;background:#DEEBF6;text-align:center">SMART FIT</th></tr>
<tr><th style="border:1px solid #000;padding:4px 6px;font-size:11px;font-weight:bold;background:#DEEBF6;white-space:nowrap">SMART FIT</th><th style="border:1px solid #000;padding:4px 6px;font-size:11px;font-weight:bold;background:#DEEBF6;white-space:nowrap">QTD</th><th style="border:1px solid #000;padding:4px 6px;font-size:11px;font-weight:bold;background:#DEEBF6;white-space:nowrap">REF.</th></tr>
{rows}
<tr><td colspan="3" style="border:1px solid #000;padding:4px 6px;font-size:10px">Solicitação: {sol.get("id","")} | Loja: {sol.get("loja","")} | Data: {data_geracao}</td></tr>
</table></body></html>"""
        filename = f'Pedido_Material_SmartFit_{sol.get("id","")}.xls'
    elif cliente == "Self Fit":
        rows = build_rows_smart_self(itens)
        html = f"""<html xmlns:x="urn:schemas-microsoft-com:office:excel"><head><meta charset="utf-8"><style>table{{border-collapse:collapse;width:100%}}th,td{{border:1px solid #000}}</style></head><body><table>
<tr><th colspan="3" style="border:1px solid #000;padding:6px;font-size:13px;font-weight:bold;background:#DEEBF6;text-align:center">SELF FIT</th></tr>
<tr><th style="border:1px solid #000;padding:4px 6px;font-size:11px;font-weight:bold;background:#DEEBF6;white-space:nowrap">SELF FIT</th><th style="border:1px solid #000;padding:4px 6px;font-size:11px;font-weight:bold;background:#DEEBF6;white-space:nowrap">QTD</th><th style="border:1px solid #000;padding:4px 6px;font-size:11px;font-weight:bold;background:#DEEBF6;white-space:nowrap">REF.</th></tr>
{rows}
<tr><td style="border:1px solid #000;padding:4px 6px;font-size:11px;font-weight:bold">DATA DO PEDIDO</td><td colspan="2" style="border:1px solid #000;padding:4px 6px;font-size:11px">&nbsp;</td></tr>
<tr><td style="border:1px solid #000;padding:4px 6px;font-size:11px;font-weight:bold">MÊS DE REFERÊNCIA</td><td colspan="2" style="border:1px solid #000;padding:4px 6px;font-size:11px">&nbsp;</td></tr>
<tr><td style="border:1px solid #000;padding:4px 6px;font-size:11px;font-weight:bold">SEPARADO</td><td colspan="2" style="border:1px solid #000;padding:4px 6px;font-size:11px">&nbsp;</td></tr>
<tr><td style="border:1px solid #000;padding:4px 6px;font-size:11px;font-weight:bold">ENVIADO</td><td colspan="2" style="border:1px solid #000;padding:4px 6px;font-size:11px">&nbsp;</td></tr>
<tr><td colspan="3" style="border:1px solid #000;padding:4px 6px;font-size:10px">Solicitação: {sol.get("id","")} | Loja: {sol.get("loja","")} | Data: {data_geracao}</td></tr>
</table></body></html>"""
        filename = f'Pedido_Material_SelfFit_{sol.get("id","")}.xls'
    elif cliente == "Assaí Atacadista":
        item_rows = build_rows_assai(itens)
        html = f"""<html xmlns:x="urn:schemas-microsoft-com:office:excel"><head><meta charset="utf-8"><style>table{{border-collapse:collapse;width:100%}}th,td{{border:1px solid #000}}</style></head><body>
<table style="width:100%;border-collapse:collapse">
<tr><td colspan="7" style="border:1px solid #000;padding:6px;font-size:11px;text-align:center;font-weight:bold">R: Sgto Jeter Augusto Pereira N° 02 e 04 - São Paulo - CEP: 02188-070 - E-mail: vendas@thamesjlara.com.br - Site www.thamesjlara.com.br</td></tr>
<tr><td colspan="7" style="height:10px"></td></tr>
<tr><td style="border:1px solid #000;padding:4px 6px;font-size:11px;font-weight:bold">Orçamento</td><td colspan="2" style="border:1px solid #000;padding:4px 6px;font-size:11px">At.: Sr.(a): Mendonça</td><td colspan="2" style="border:1px solid #000;padding:4px 6px;font-size:11px">Data: {data_geracao}</td><td colspan="2" style="border:1px solid #000;padding:4px 6px;font-size:11px">Vendedor: Hélio</td></tr>
<tr><td colspan="7" style="height:10px"></td></tr>
<tr><td colspan="2" style="border:1px solid #000;padding:4px 6px;font-size:11px;font-weight:bold">Razão Social</td><td colspan="3" style="border:1px solid #000;padding:4px 6px;font-size:11px">FG Services Eireli - ME</td><td style="border:1px solid #000;padding:4px 6px;font-size:11px;font-weight:bold">CNPJ/CPF</td><td style="border:1px solid #000;padding:4px 6px;font-size:11px">23.585.374/0001-11</td></tr>
<tr><td colspan="2" style="border:1px solid #000;padding:4px 6px;font-size:11px;font-weight:bold">Endereço</td><td colspan="2" style="border:1px solid #000;padding:4px 6px;font-size:11px">Av. Barão de Vera Cruz, 586 BR 101 Norte</td><td style="border:1px solid #000;padding:4px 6px;font-size:11px;font-weight:bold">Bairro</td><td style="border:1px solid #000;padding:4px 6px;font-size:11px">Cruz de Rebouças</td><td style="border:1px solid #000;padding:4px 6px;font-size:11px;font-weight:bold">Cidade</td></tr>
<tr><td style="border:1px solid #000;padding:4px 6px;font-size:11px">Igarassu</td><td style="border:1px solid #000;padding:4px 6px;font-size:11px;font-weight:bold">UF</td><td style="border:1px solid #000;padding:4px 6px;font-size:11px">PE</td><td colspan="2" style="border:1px solid #000;padding:4px 6px;font-size:11px;font-weight:bold">CEP</td><td colspan="2" style="border:1px solid #000;padding:4px 6px;font-size:11px">53635-015</td></tr>
<tr><td colspan="7" style="height:10px"></td></tr>
<tr><td style="border:1px solid #000;padding:4px 6px;font-size:11px;font-weight:bold">Telefone</td><td colspan="2" style="border:1px solid #000;padding:4px 6px;font-size:11px">(0xx81) 3545-3990</td><td colspan="2" style="border:1px solid #000;padding:4px 6px;font-size:11px;font-weight:bold">E-mail</td><td colspan="2" style="border:1px solid #000;padding:4px 6px;font-size:11px">&nbsp;</td></tr>
<tr><td colspan="7" style="height:10px"></td></tr>
<tr><th style="border:1px solid #000;padding:4px 6px;font-size:11px;font-weight:bold;background:#DEEBF6;white-space:nowrap">Item</th><th style="border:1px solid #000;padding:4px 6px;font-size:11px;font-weight:bold;background:#DEEBF6;white-space:nowrap">Descrição</th><th style="border:1px solid #000;padding:4px 6px;font-size:11px;font-weight:bold;background:#DEEBF6;white-space:nowrap">Marca</th><th style="border:1px solid #000;padding:4px 6px;font-size:11px;font-weight:bold;background:#DEEBF6;white-space:nowrap">Qtde</th><th style="border:1px solid #000;padding:4px 6px;font-size:11px;font-weight:bold;background:#DEEBF6;white-space:nowrap">Unid</th><th style="border:1px solid #000;padding:4px 6px;font-size:11px;font-weight:bold;background:#DEEBF6;white-space:nowrap">Valor Unit.</th><th style="border:1px solid #000;padding:4px 6px;font-size:11px;font-weight:bold;background:#DEEBF6;white-space:nowrap">Total</th></tr>
{item_rows}
<tr><td colspan="7" style="border:1px solid #000;padding:4px 6px;font-size:10px">Solicitação: {sol.get("id","")} | Loja: {sol.get("loja","")} | Data: {data_geracao}</td></tr>
</table></body></html>"""
        filename = f'Orcamento_Assai_{sol.get("id","")}.xls'
    else:
        return None, None
    return html, filename


# ========== PÁGINAS ==========
def page_dashboard():
    st.markdown("### 📊 Dashboard")
    sols = st.session_state["compras_solicitacoes"]
    ents = st.session_state["compras_entregas"]

    total = len(sols)
    pendentes = sum(1 for s in sols if s.get("status") == "Pendente")
    transito = sum(1 for s in sols if s.get("status") == "Em Trânsito")
    entregues = sum(1 for s in sols if s.get("status") == "Entregue")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Solicitações", total)
    c2.metric("Pendentes", pendentes)
    c3.metric("Em Trânsito", transito)
    c4.metric("Entregues", entregues)

    st.markdown("---")
    st.markdown("#### 📋 Últimas Solicitações")
    if sols:
        recentes = sorted(sols, key=lambda x: x.get("dataCriacao", ""), reverse=True)[:10]
        df = pd.DataFrame([
            {
                "ID": s["id"],
                "Data": formatar_data_br(s.get("data")),
                "Loja": s.get("loja", ""),
                "Cliente": s.get("cliente", ""),
                "Tipo": s.get("tipo", ""),
                "Solicitante": s.get("solicitante", ""),
                "Itens": len(s.get("itens", [])),
                "Valor": formatar_moeda(s.get("valorTotal", 0)),
                "Status": f"{get_badge_color(s.get('status'))} {s.get('status', '')}"
            }
            for s in recentes
        ])
        st.dataframe(df, use_container_width=True, hide_index=True)
    else:
        st.info("Nenhuma solicitação cadastrada.")

    if st.button("➕ Nova Solicitação", type="primary", key="btn_nova_sol_dashboard"):
        switch_page("Nova Solicitação")



def _salvar_compras_local(solicitacoes, entregas):
    """Salva dados de compras em arquivo JSON local."""
    try:
        caminho = os.path.abspath(ARQUIVO_COMPRAS)
        with open(ARQUIVO_COMPRAS, "w", encoding="utf-8") as f:
            json.dump({"solicitacoes": solicitacoes, "entregas": entregas}, f, ensure_ascii=False, indent=2)
        st.toast(f"💾 Compras salvas localmente ({len(solicitacoes)} solicitações, {len(entregas)} entregas) — arquivo: {caminho}")
    except Exception as e:
        st.error(f"❌ Erro ao salvar compras localmente: {e}")


def _carregar_compras_local():
    """Carrega dados de compras do arquivo JSON local."""
    caminho = os.path.abspath(ARQUIVO_COMPRAS)
    if not os.path.exists(ARQUIVO_COMPRAS):
        st.toast(f"📂 Nenhum arquivo de compras encontrado em: {caminho}")
        return [], []
    try:
        with open(ARQUIVO_COMPRAS, "r", encoding="utf-8") as f:
            conteudo = f.read()
        if not conteudo.strip():
            st.warning(f"⚠️ Arquivo de compras vazio: {caminho}")
            return [], []
        dados = json.loads(conteudo)
        sols = dados.get("solicitacoes", [])
        ents = dados.get("entregas", [])
        st.toast(f"📂 Compras carregadas do arquivo ({len(sols)} solicitações, {len(ents)} entregas)")
        return sols, ents
    except Exception as e:
        st.error(f"❌ Erro ao carregar compras do arquivo {caminho}: {e}")
        return [], []


def _salvar_compras_automatico():
    """Salva automaticamente o estado atual do módulo de compras no Google Sheets e localmente."""
    import sys
    sols = st.session_state.get("compras_solicitacoes", [])
    ents = st.session_state.get("compras_entregas", [])
    print(f"[DEBUG] _salvar_compras_automatico chamado com {len(sols)} solicitações, {len(ents)} entregas", file=sys.stderr)
    # Sempre salva localmente para garantir persistência
    _salvar_compras_local(sols, ents)
    if not GS_ENABLED or not GS_ID_COMPRAS:
        return
    try:
        _salvar_compras_gs(sols, ents)
    except Exception as e:
        print(f"[DEBUG] Falha ao salvar no GS: {e}", file=sys.stderr)


# Inicialização Google Sheets: garante que abas existam

def page_solicitacoes():
    st.markdown("### 📋 Todas as Solicitações")

    sols = st.session_state["compras_solicitacoes"]

    # Filtros
    with st.expander("🔍 Filtros", expanded=True):
        c1, c2, c3, c4, c5, c6 = st.columns(6)
        with c1:
            busca = st.text_input("Buscar", placeholder="ID, loja, solicitante...", key="sol_busca")
        with c2:
            fstatus = st.selectbox("Status", ["Todos"] + STATUS_OPCOES, key="sol_fstatus")
        with c3:
            floja = st.selectbox("Loja", ["Todas"] + sorted(list(set(l for c in LOJAS_POR_CLIENTE.values() for l in c))), key="sol_floja")
        with c4:
            fcliente = st.selectbox("Cliente", ["Todos"] + CLIENTES, key="sol_fcliente")
        with c5:
            fdi_txt = st.text_input("Data Início (DD/MM/AAAA)", value="", key="sol_fdi")
            fdi = _parse_data_br(fdi_txt) if fdi_txt.strip() else None
        with c6:
            fdf_txt = st.text_input("Data Fim (DD/MM/AAAA)", value="", key="sol_fdf")
            fdf = _parse_data_br(fdf_txt) if fdf_txt.strip() else None

    filtradas = []
    for s in sols:
        txt = f"{s.get('id','')} {s.get('loja','')} {s.get('solicitante','')}".lower()
        if busca and busca.lower() not in txt:
            continue
        if fstatus != "Todos" and s.get("status") != fstatus:
            continue
        if floja != "Todas" and s.get("loja") != floja:
            continue
        if fcliente != "Todos" and s.get("cliente") != fcliente:
            continue
        if fdi:
            try:
                if s.get("data") and str(s["data"]) < str(fdi):
                    continue
            except Exception:
                pass
        if fdf:
            try:
                if s.get("data") and str(s["data"]) > str(fdf):
                    continue
            except Exception:
                pass
        filtradas.append(s)

    if filtradas:
        for s in filtradas:
            with st.container(border=True):
                col1, col2, col3 = st.columns([4, 2, 2])
                with col1:
                    st.write(f"**{s['id']}** | {formatar_data_br(s.get('data'))} | {s.get('loja','')} | {s.get('cliente','')}")
                    st.caption(f"Tipo: `{s.get('tipo','')}` | Solicitante: {s.get('solicitante','')} | Itens: {len(s.get('itens',[]))} | Valor: {formatar_moeda(s.get('valorTotal',0))}")
                with col2:
                    st.write(f"{get_badge_color(s.get('status'))} **{s.get('status','')}**")
                with col3:
                    c_a, c_b, c_c = st.columns(3)
                    with c_a:
                        if st.button("👁️", key=f"ver_{s['id']}"):
                            st.session_state["compras_ver_id"] = s["id"]
                            st.session_state["compras_page"] = "Detalhes"

                            st.rerun()
                    with c_b:
                        if st.button("✏️", key=f"edit_{s['id']}"):
                            st.session_state["compras_edit_id"] = s["id"]
                            st.session_state["compras_page"] = "Nova Solicitação"

                            st.rerun()
                    with c_c:
                        if st.button("🗑️", key=f"del_{s['id']}"):
                            st.session_state["compras_solicitacoes"] = [x for x in sols if x["id"] != s["id"]]
                            st.session_state["compras_entregas"] = [e for e in st.session_state["compras_entregas"] if e.get("idSolicitacao") != s["id"]]
                            _salvar_compras_automatico()
                            st.success("Excluído!")
                            st.rerun()

        # Exportar CSV
        if filtradas:
            csv_data = []
            for s in filtradas:
                csv_data.append([
                    s["id"], s.get("data",""), s.get("loja",""), s.get("cliente",""),
                    s.get("tipo",""), s.get("solicitante",""), len(s.get("itens",[])),
                    s.get("valorTotal",0), s.get("status","")
                ])
            df_csv = pd.DataFrame(csv_data, columns=["ID","Data","Loja","Cliente","Tipo","Solicitante","Qtd Itens","Valor Total","Status"])
            csv_buffer = io.StringIO()
            df_csv.to_csv(csv_buffer, index=False, sep=";", encoding="utf-8")
            st.download_button("📥 Exportar CSV", data=csv_buffer.getvalue().encode("utf-8"), file_name="solicitacoes.csv", mime="text/csv")
    else:
        st.info("Nenhuma solicitação encontrada.")


def page_nova_solicitacao():
    st.markdown("### ➕ Nova Solicitação de Compra")

    edit_id = st.session_state.get("compras_edit_id")
    edit_sol = None
    if edit_id:
        for s in st.session_state["compras_solicitacoes"]:
            if s["id"] == edit_id:
                edit_sol = s
                break

    # Detectar mudança de contexto (nova solicitação ou edição diferente) e limpar estado do cabeçalho
    last_edit_id = st.session_state.get("compras_last_edit_id")
    if last_edit_id != edit_id:
        for k in ["nova_cliente_val", "nova_loja_val", "nova_tipo_val", "nova_solicitante_val",
                  "nova_data_val", "nova_prioridade_val", "nova_previsao_val", "nova_obs_val",
                  "nova_nome_func_val", "nova_encarregado_val", "nova_supervisor_val", "nova_data_bota_val"]:
            if k in st.session_state:
                del st.session_state[k]
        st.session_state["compras_last_edit_id"] = edit_id

    # Inicializar valores do cabeçalho no session_state
    if "nova_cliente_val" not in st.session_state:
        st.session_state["nova_cliente_val"] = edit_sol["cliente"] if edit_sol and edit_sol.get("cliente") in CLIENTES else CLIENTES[0]
    if "nova_loja_val" not in st.session_state:
        st.session_state["nova_loja_val"] = edit_sol["loja"] if edit_sol and edit_sol.get("loja") else ""
    if "nova_tipo_val" not in st.session_state:
        st.session_state["nova_tipo_val"] = edit_sol["tipo"] if edit_sol and edit_sol.get("tipo") in TIPOS_SOLICITACAO else TIPOS_SOLICITACAO[0]
    if "nova_solicitante_val" not in st.session_state:
        st.session_state["nova_solicitante_val"] = edit_sol.get("solicitante", "") if edit_sol else ""
    if "nova_data_val" not in st.session_state:
        if edit_sol and edit_sol.get("data"):
            st.session_state["nova_data_val"] = _iso_para_br(edit_sol["data"])
        else:
            st.session_state["nova_data_val"] = datetime.now().strftime("%d/%m/%Y")
    if "nova_prioridade_val" not in st.session_state:
        st.session_state["nova_prioridade_val"] = edit_sol["prioridade"] if edit_sol and edit_sol.get("prioridade") in PRIORIDADES else PRIORIDADES[0]
    if "nova_previsao_val" not in st.session_state:
        if edit_sol and edit_sol.get("previsao"):
            st.session_state["nova_previsao_val"] = _iso_para_br(edit_sol["previsao"])
        else:
            st.session_state["nova_previsao_val"] = (datetime.now() + timedelta(days=7)).strftime("%d/%m/%Y")
    if "nova_obs_val" not in st.session_state:
        st.session_state["nova_obs_val"] = edit_sol.get("observacoes", "") if edit_sol else ""
    if "nova_nome_func_val" not in st.session_state:
        st.session_state["nova_nome_func_val"] = edit_sol.get("nomeFuncionario", "") if edit_sol else ""
    if "nova_encarregado_val" not in st.session_state:
        st.session_state["nova_encarregado_val"] = edit_sol.get("encarregado", "") if edit_sol else ""
    if "nova_supervisor_val" not in st.session_state:
        st.session_state["nova_supervisor_val"] = edit_sol.get("supervisor", "") if edit_sol else ""
    if "nova_data_bota_val" not in st.session_state:
        if edit_sol and edit_sol.get("dataUltimaBota"):
            st.session_state["nova_data_bota_val"] = _iso_para_br(edit_sol["dataUltimaBota"])
        else:
            st.session_state["nova_data_bota_val"] = ""

    # ── CABEÇALHO DENTRO DO FORM (evita rerun a cada interação) ──
    with st.form("form_cabecalho", clear_on_submit=False):
        col_sel1, col_sel2, col_sel3 = st.columns(3)
        with col_sel1:
            cliente_sel = st.selectbox(
                "Cliente *",
                CLIENTES,
                index=CLIENTES.index(st.session_state["nova_cliente_val"]) if st.session_state["nova_cliente_val"] in CLIENTES else 0,
                key="form_cliente"
            )
        with col_sel2:
            lojas = LOJAS_POR_CLIENTE.get(cliente_sel, [])
            loja_idx = 0
            if st.session_state["nova_loja_val"] in lojas:
                loja_idx = lojas.index(st.session_state["nova_loja_val"])
            loja_sel = st.selectbox("Loja *", lojas, index=loja_idx, key="form_loja")
        with col_sel3:
            tipo_sel = st.selectbox(
                "Tipo *",
                TIPOS_SOLICITACAO,
                index=TIPOS_SOLICITACAO.index(st.session_state["nova_tipo_val"]) if st.session_state["nova_tipo_val"] in TIPOS_SOLICITACAO else 0,
                key="form_tipo"
            )

        c4, c5, c6 = st.columns(3)
        with c4:
            solicitante_sel = st.text_input("Solicitante *", value=st.session_state["nova_solicitante_val"], key="form_solicitante")
        with c5:
            data_sol_txt = st.text_input("Data (DD/MM/AAAA)", value=st.session_state["nova_data_val"], key="form_data")
        with c6:
            prioridade_sel = st.selectbox(
                "Prioridade",
                PRIORIDADES,
                index=PRIORIDADES.index(st.session_state["nova_prioridade_val"]) if st.session_state["nova_prioridade_val"] in PRIORIDADES else 0,
                key="form_prioridade"
            )

        previsao_txt = st.text_input("Previsão de Entrega (DD/MM/AAAA)", value=st.session_state["nova_previsao_val"], key="form_previsao")
        observacoes_sel = st.text_area("Observações", value=st.session_state["nova_obs_val"], key="form_obs")

        if tipo_sel == "EPI":
            st.markdown("---")
            st.markdown("#### 👷 Informações EPI")
            c7, c8, c9 = st.columns(3)
            with c7:
                nome_func_sel = st.text_input("Nome do Funcionário", value=st.session_state["nova_nome_func_val"], key="form_nome_func")
            with c8:
                encarregado_sel = st.text_input("Encarregado", value=st.session_state["nova_encarregado_val"], key="form_encarregado")
            with c9:
                supervisor_sel = st.text_input("Supervisor", value=st.session_state["nova_supervisor_val"], key="form_supervisor")
            data_bota_txt = st.text_input("Data Última Bota (DD/MM/AAAA)", value=st.session_state["nova_data_bota_val"], key="form_data_bota")

        aplicar = st.form_submit_button("📋 Aplicar Cabeçalho", type="secondary")

    if aplicar:
        st.session_state["nova_cliente_val"] = cliente_sel
        st.session_state["nova_loja_val"] = loja_sel
        st.session_state["nova_tipo_val"] = tipo_sel
        st.session_state["nova_solicitante_val"] = solicitante_sel
        st.session_state["nova_data_val"] = data_sol_txt
        st.session_state["nova_prioridade_val"] = prioridade_sel
        st.session_state["nova_previsao_val"] = previsao_txt
        st.session_state["nova_obs_val"] = observacoes_sel
        if tipo_sel == "EPI":
            st.session_state["nova_nome_func_val"] = nome_func_sel
            st.session_state["nova_encarregado_val"] = encarregado_sel
            st.session_state["nova_supervisor_val"] = supervisor_sel
            st.session_state["nova_data_bota_val"] = data_bota_txt
        st.rerun()

    # ── VALORES FINAIS DO CABEÇALHO ──
    cliente = st.session_state.get("nova_cliente_val", CLIENTES[0])
    loja = st.session_state.get("nova_loja_val", "")
    tipo = st.session_state.get("nova_tipo_val", TIPOS_SOLICITACAO[0])
    solicitante = st.session_state.get("nova_solicitante_val", "")
    data_sol = _parse_data_br(st.session_state.get("nova_data_val", ""))
    prioridade = st.session_state.get("nova_prioridade_val", PRIORIDADES[0])
    previsao = _parse_data_br(st.session_state.get("nova_previsao_val", ""))
    observacoes = st.session_state.get("nova_obs_val", "")
    nome_func = st.session_state.get("nova_nome_func_val", "")
    encarregado = st.session_state.get("nova_encarregado_val", "")
    supervisor = st.session_state.get("nova_supervisor_val", "")
    data_bota = _parse_data_br(st.session_state.get("nova_data_bota_val", "")) if st.session_state.get("nova_data_bota_val") else None

    st.markdown("---")
    st.markdown("#### 📦 Itens")

    if tipo == "Material":
        st.markdown("**Materiais**")
        materiais = MATERIAIS_POR_CLIENTE.get(cliente, [])

        if edit_sol and not st.session_state.get("compras_edit_loaded"):
            st.session_state["compras_nova_itens_material"] = [
                {"material": i.get("material", ""), "qtd": i.get("qtd", 1), "valorUnit": i.get("valorUnit", 0)}
                for i in edit_sol.get("itens", [])
            ]
            st.session_state["compras_edit_loaded"] = True

        itens_mat_default = st.session_state.get("compras_nova_itens_material", [])
        if not itens_mat_default:
            itens_mat_default = [{"material": materiais[0] if materiais else "", "qtd": 1, "valorUnit": 0.0}]

        df_mat = pd.DataFrame(itens_mat_default)

        edited_mat = st.data_editor(
            df_mat,
            num_rows="dynamic",
            use_container_width=True,
            key="editor_mat",
            column_config={
                "material": st.column_config.SelectboxColumn(
                    "Material",
                    options=materiais,
                    required=True,
                ),
                "qtd": st.column_config.NumberColumn(
                    "Qtd",
                    min_value=1,
                    default=1,
                    step=1,
                ),
                "valorUnit": st.column_config.NumberColumn(
                    "Valor Unit.",
                    min_value=0.0,
                    format="%.2f",
                    step=0.01,
                    default=0.0,
                ),
            },
            hide_index=True,
        )

        st.session_state["compras_nova_itens_material"] = edited_mat.to_dict("records")

    elif tipo == "EPI":
        st.markdown("**EPIs**")
        epis = EPIS_POR_CLIENTE.get(cliente, EPIS_POR_CLIENTE.get("Smart Fit", []))

        if edit_sol and not st.session_state.get("compras_edit_loaded"):
            st.session_state["compras_nova_itens_epi"] = [
                {"epi": i.get("epi", ""), "colaborador": i.get("colaborador", ""), "qtd": i.get("qtd", 1), "tamanho": i.get("tamanho", "")}
                for i in edit_sol.get("itens", [])
            ]
            st.session_state["compras_edit_loaded"] = True

        itens_epi_default = st.session_state.get("compras_nova_itens_epi", [])
        if not itens_epi_default:
            itens_epi_default = [{"epi": epis[0] if epis else "", "colaborador": "", "qtd": 1, "tamanho": ""}]

        df_epi = pd.DataFrame(itens_epi_default)

        edited_epi = st.data_editor(
            df_epi,
            num_rows="dynamic",
            use_container_width=True,
            key="editor_epi",
            column_config={
                "epi": st.column_config.SelectboxColumn(
                    "EPI",
                    options=epis,
                    required=True,
                ),
                "colaborador": st.column_config.TextColumn(
                    "Colaborador",
                ),
                "qtd": st.column_config.NumberColumn(
                    "Qtd",
                    min_value=1,
                    default=1,
                    step=1,
                ),
                "tamanho": st.column_config.SelectboxColumn(
                    "Tamanho",
                    options=TAMANHOS_EPI,
                ),
            },
            hide_index=True,
        )

        st.session_state["compras_nova_itens_epi"] = edited_epi.to_dict("records")

    submitted = st.button("💾 Salvar Solicitação", type="primary")

    if submitted:
        itens = []
        valor_total = 0
        if tipo == "Material":
            itens = st.session_state.get("compras_nova_itens_material", [])
            valor_total = sum(i.get("valorUnit", 0) * i.get("qtd", 0) for i in itens)
            if not itens:
                st.error("Adicione pelo menos um material.")
                return
        elif tipo == "EPI":
            itens = st.session_state.get("compras_nova_itens_epi", [])
            if not itens:
                st.error("Adicione pelo menos um EPI.")
                return

        nova_sol = {
            "id": edit_sol["id"] if edit_sol else gerar_id(),
            "data": str(data_sol),
            "cliente": cliente,
            "loja": loja,
            "tipo": tipo,
            "solicitante": solicitante,
            "nomeFuncionario": nome_func if tipo == "EPI" else "",
            "encarregado": encarregado if tipo == "EPI" else "",
            "supervisor": supervisor if tipo == "EPI" else "",
            "dataUltimaBota": str(data_bota) if tipo == "EPI" and data_bota else "",
            "prioridade": prioridade,
            "previsao": str(previsao),
            "observacoes": observacoes,
            "itens": itens,
            "valorTotal": valor_total,
            "anexos": [],
            "status": edit_sol.get("status", "Pendente") if edit_sol else "Pendente",
            "dataCriacao": edit_sol.get("dataCriacao", datetime.now().isoformat()) if edit_sol else datetime.now().isoformat()
        }

        if tipo == "EPI":
            html_doc = gerar_doc_epi(nova_sol)
            nova_sol["anexos"].append({
                "nome": f'Solicitacao_de_EPI_{nova_sol["id"]}_{cliente}_{loja}.doc',
                "conteudo": html_doc,
                "tipo": "doc"
            })
        elif tipo == "Material":
            html_xls, filename_xls = gerar_xls_material(nova_sol)
            if html_xls:
                nova_sol["anexos"].append({
                    "nome": filename_xls,
                    "conteudo": html_xls,
                    "tipo": "xls"
                })

        sols = st.session_state["compras_solicitacoes"]
        if edit_sol:
            sols = [s for s in sols if s["id"] != edit_id]
        sols.append(nova_sol)
        st.session_state["compras_solicitacoes"] = sols

        if not edit_sol:
            st.session_state["compras_entregas"].append({
                "idSolicitacao": nova_sol["id"],
                "loja": loja,
                "tipo": tipo,
                "dataEnvio": "",
                "dataPrevista": str(previsao),
                "dataEntrega": "",
                "transportadora": "",
                "rastreio": "",
                "status": "Pendente",
                "observacoes": ""
            })

        # Limpa estado
        st.session_state["compras_nova_itens_material"] = []
        st.session_state["compras_nova_itens_epi"] = []
        st.session_state["compras_edit_id"] = None
        st.session_state["compras_edit_loaded"] = False
        for k in ["nova_cliente_val", "nova_loja_val", "nova_tipo_val", "nova_solicitante_val",
                  "nova_data_val", "nova_prioridade_val", "nova_previsao_val", "nova_obs_val",
                  "nova_nome_func_val", "nova_encarregado_val", "nova_supervisor_val", "nova_data_bota_val"]:
            if k in st.session_state:
                del st.session_state[k]
        _salvar_compras_automatico()
        st.success("Solicitação salva com sucesso!")
        st.session_state["compras_page"] = "Solicitações"
        st.rerun()
