# -*- coding: utf-8 -*-
"""
Sistema RH — Diárias · Viagens · Compras · Lojas · Eventos Trabalhistas
Banco de dados SQLite embutido: TODOS os dados ficam no próprio sistema.
Leitura automática de PDFs de admissão (Ficha Registro + Contrato Experiência).
"""

import streamlit as st
import sqlite3
import json
import os
import io
import zipfile
import base64
import hashlib
import re
from datetime import datetime, date, timedelta
import pandas as pd

# Leitura de PDF — pdfplumber (melhor) ou pypdf (fallback)
try:
    import pdfplumber
    TEM_LEITOR_PDF = True
    LEITOR_PDF = "pdfplumber"
except ImportError:
    try:
        import pypdf
        TEM_LEITOR_PDF = True
        LEITOR_PDF = "pypdf"
    except ImportError:
        TEM_LEITOR_PDF = False
        LEITOR_PDF = None

# ════════════════════════════════════════════════════════════════
# BANCO DE DADOS SQLITE — TUDO DENTRO DO PRÓPRIO SISTEMA
# ════════════════════════════════════════════════════════════════

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "rh_dados.db")

def _conn():
    con = sqlite3.connect(DB_PATH, timeout=15)
    con.execute("PRAGMA journal_mode=WAL")
    con.execute("PRAGMA foreign_keys=ON")
    con.row_factory = sqlite3.Row
    return con

def _init_db():
    con = _conn()
    c = con.cursor()
    c.executescript("""
    CREATE TABLE IF NOT EXISTS funcionarios (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nome TEXT NOT NULL, cpf TEXT, rg TEXT, data_nascimento TEXT,
        estado_civil TEXT, endereco TEXT, cidade TEXT, estado TEXT,
        cep TEXT, telefone TEXT,
        cargo TEXT, loja TEXT, data_admissao TEXT,
        salario REAL, banco TEXT,
        ctps TEXT, pis TEXT, cbo TEXT,
        naturalidade TEXT, sexo TEXT, raca_cor TEXT,
        grau_instrucao TEXT, filiacao_pai TEXT, filiacao_mae TEXT,
        tipo_contrato TEXT, situacao TEXT DEFAULT 'Ativo',
        data_demissao TEXT, motivo_demissao TEXT,
        inicio_experiencia TEXT, fim_experiencia TEXT,
        inicio_ferias TEXT, dias_ferias INTEGER,
        fim_ferias TEXT, inicio_licenca TEXT, dias_licenca INTEGER,
        fim_licenca TEXT, inicio_afastamento TEXT, dias_afastamento INTEGER,
        fim_afastamento TEXT, inicio_aviso TEXT, dias_aviso INTEGER,
        fim_aviso TEXT, observacoes TEXT,
        criado_em TEXT, atualizado_em TEXT
    );
    CREATE TABLE IF NOT EXISTS documentos (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        funcionario_id INTEGER,
        tipo TEXT, descricao TEXT,
        nome_arquivo TEXT, conteudo BLOB, tamanho INTEGER,
        criado_em TEXT,
        FOREIGN KEY(funcionario_id) REFERENCES funcionarios(id) ON DELETE CASCADE
    );
    CREATE TABLE IF NOT EXISTS fotos_funcionario (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        funcionario_id INTEGER UNIQUE,
        nome_arquivo TEXT, conteudo BLOB, tamanho INTEGER,
        criado_em TEXT,
        FOREIGN KEY(funcionario_id) REFERENCES funcionarios(id) ON DELETE CASCADE
    );
    CREATE TABLE IF NOT EXISTS lojas (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nome TEXT NOT NULL UNIQUE, endereco TEXT, cidade TEXT,
        estado TEXT, telefone TEXT, responsavel TEXT,
        ativa INTEGER DEFAULT 1, observacoes TEXT,
        criado_em TEXT, atualizado_em TEXT
    );
    CREATE TABLE IF NOT EXISTS documentos_loja (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        loja_id INTEGER,
        tipo TEXT, descricao TEXT,
        nome_arquivo TEXT, conteudo BLOB, tamanho INTEGER,
        criado_em TEXT,
        FOREIGN KEY(loja_id) REFERENCES lojas(id) ON DELETE CASCADE
    );
    CREATE TABLE IF NOT EXISTS diarias (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        funcionario_id INTEGER, funcionario_nome TEXT,
        destino TEXT, motivo TEXT,
        data_saida TEXT, data_retorno TEXT,
        valor_diaria REAL, quantidade INTEGER, valor_total REAL,
        observacoes TEXT, status TEXT DEFAULT 'Pendente',
        criado_em TEXT, atualizado_em TEXT,
        FOREIGN KEY(funcionario_id) REFERENCES funcionarios(id)
    );
    CREATE TABLE IF NOT EXISTS comprovantes_diaria (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        diaria_id INTEGER,
        nome_arquivo TEXT, conteudo BLOB, tamanho INTEGER,
        criado_em TEXT,
        FOREIGN KEY(diaria_id) REFERENCES diarias(id) ON DELETE CASCADE
    );
    CREATE TABLE IF NOT EXISTS viagens (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        funcionario_id INTEGER, funcionario_nome TEXT,
        destino TEXT, motivo TEXT,
        data_saida TEXT, data_retorno TEXT,
        transporte TEXT, hospedagem TEXT,
        valor_estimado REAL, observacoes TEXT,
        status TEXT DEFAULT 'Pendente',
        criado_em TEXT, atualizado_em TEXT,
        FOREIGN KEY(funcionario_id) REFERENCES funcionarios(id)
    );
    CREATE TABLE IF NOT EXISTS compras (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        loja_id INTEGER, loja_nome TEXT,
        fornecedor TEXT, descricao TEXT,
        quantidade INTEGER, valor_unitario REAL, valor_total REAL,
        data_compra TEXT, categoria TEXT,
        status TEXT DEFAULT 'Pendente',
        observacoes TEXT, criado_em TEXT, atualizado_em TEXT,
        FOREIGN KEY(loja_id) REFERENCES lojas(id)
    );
    CREATE TABLE IF NOT EXISTS eventos_trabalhistas (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        funcionario_id INTEGER,
        tipo TEXT, data_inicio TEXT, data_fim TEXT,
        dias INTEGER, observacoes TEXT,
        criado_em TEXT
    );
    CREATE TABLE IF NOT EXISTS config (
        chave TEXT PRIMARY KEY, valor TEXT
    );
    """)
    con.commit()

    # ── Migração: adicionar colunas novas se não existirem ──
    _migrar_banco(con)
    con.close()

def _migrar_banco(con):
    """Adiciona colunas novas em bancos já existentes e remove obsoletas."""
    existing = {r[1] for r in con.execute("PRAGMA table_info(funcionarios)").fetchall()}
    novas = {
        "ctps": "TEXT", "pis": "TEXT", "cbo": "TEXT",
        "naturalidade": "TEXT", "sexo": "TEXT", "raca_cor": "TEXT",
        "grau_instrucao": "TEXT", "filiacao_pai": "TEXT", "filiacao_mae": "TEXT",
    }
    for col, tipo in novas.items():
        if col not in existing:
            try:
                con.execute(f"ALTER TABLE funcionarios ADD COLUMN {col} {tipo}")
            except Exception:
                pass
    # Colunas obsoletas que podem existir em bancos antigos
    obsoletas = ["email", "pix", "departamento", "agencia", "conta"]
    for col in obsoletas:
        if col in existing:
            # SQLite não suporta DROP COLUMN em versões antigas;
            # mantemos a coluna mas removemos do formulário
            pass
    con.commit()

_init_db()

# ════════════════════════════════════════════════════════════════
# LEITURA AUTOMÁTICA DE PDFs DE ADMISSÃO
# ════════════════════════════════════════════════════════════════

def _ler_pdf_paginas(arquivo_pdf):
    """Lê um PDF e retorna lista de textos por página (usa pdfplumber ou pypdf)."""
    pdf_bytes = io.BytesIO(arquivo_pdf.read()) if hasattr(arquivo_pdf, "read") else arquivo_pdf
    pdf_bytes.seek(0)
    paginas = []

    if LEITOR_PDF == "pdfplumber":
        try:
            pdf = pdfplumber.open(pdf_bytes)
            for page in pdf.pages:
                t = page.extract_text()
                if t:
                    paginas.append(t)
            pdf.close()
            return paginas
        except Exception:
            pdf_bytes.seek(0)

    # Fallback: pypdf
    if LEITOR_PDF == "pypdf":
        try:
            import pypdf as _pypdf
            reader = _pypdf.PdfReader(pdf_bytes)
            for page in reader.pages:
                t = page.extract_text()
                if t:
                    paginas.append(t)
            return paginas
        except Exception:
            return []

    return []


def _extrair_campos_de_paginas(paginas):
    """
    Recebe uma lista de textos de páginas de UM PDF e retorna dict
    com os campos extraídos da Ficha de Registro e/ou Contrato de Experiência.
    """
    campos = {}
    txt = "\n".join(paginas)

    # ── Detecta tipo de documento e separa texto por seção ──
    eh_ficha = bool(re.search(r"REGISTRO DE EMPREGADO|FICHA DE REGISTRO", txt, re.I))
    eh_contrato = bool(re.search(r"CONTRATO DE EXPERI[ÊE]NCIA", txt, re.I))

    # Texto isolado da ficha (última página que contém REGISTRO DE EMPREGADO)
    txt_ficha = ""
    for p in paginas:
        if re.search(r"REGISTRO DE EMPREGADO|FICHA DE REGISTRO", p, re.I):
            txt_ficha = p

    # Texto isolado do contrato (primeira página com CONTRATO DE EXPERIÊNCIA)
    txt_contrato = ""
    for p in paginas:
        if re.search(r"CONTRATO DE EXPERI[ÊE]NCIA", p, re.I) and not txt_contrato:
            txt_contrato = p

    # ═══ EXTRAÇÃO DA FICHA DE REGISTRO ═══
    if eh_ficha and txt_ficha:
        f = txt_ficha

        # --- Nome ---
        m = re.search(r"^Nome:\s*([A-ZÀÁÂÃÉÊÍÓÔÕÚÜÇ'\s]+?)$", f, re.M)
        if m and m.group(1).strip():
            campos["nome"] = m.group(1).strip()
        else:
            m = re.search(r"(?:Empregado\s+Beneficiários|Empregado)\s*\n\s*([A-ZÀÁÂÃÉÊÍÓÔÕÚÜÇ'\s]+?)\s*\n", f)
            if m and m.group(1).strip():
                campos["nome"] = m.group(1).strip()

        # --- CPF ---
        m = re.search(r"CPF[:\s]*(\d{3}\.?\d{3}\.?\d{3}[-.]?\d{2})", f)
        if m:
            campos["cpf"] = re.sub(r"[^0-9]", "", m.group(1))

        # --- RG ---
        m = re.search(r"RG\s*Número[:\s]*(\d+)", f, re.I)
        if m:
            campos["rg"] = m.group(1).strip()
        else:
            m = re.search(r"C[eé]dula de Identidade[^\n]*\n\s*(\d[\d.-]+)", f, re.I)
            if m:
                campos["rg"] = m.group(1).strip().rstrip(".")

        # --- CTPS + Série ---
        m = re.search(r"CTPS\s*Número[:\s]*(\d[\d/]*)\s*S[eé]rie[:\s]*(\d+)", f, re.I)
        if m:
            campos["ctps"] = f"{m.group(1).strip()}/{m.group(2).strip()}"
        else:
            m = re.search(r"CTPS[\s\n]+S[eé]rie[^\n]*\n\s*(\d{4,})\s+(\d{3,})", f, re.I)
            if m:
                campos["ctps"] = f"{m.group(1).strip()}/{m.group(2).strip()}"

        # --- PIS ---
        m = re.search(r"PIS\s*/?\s*PASEP[:\s]*([\d.\-]+)", f, re.I)
        if m:
            pis_val = m.group(1).strip().rstrip(".")
            if pis_val and not pis_val.replace(".", "").replace("-", "").startswith("000"):
                campos["pis"] = pis_val

        # --- Data Nascimento ---
        m = re.search(r"(?:Data de nascimento|Nascimento)[:\s]*(\d{2}/\d{2}/\d{4})", f, re.I)
        if m:
            campos["data_nascimento"] = _converte_data(m.group(1))

        # --- Naturalidade ---
        m = re.search(r"Naturalidade[:\s]*([A-ZÀÁÂÃÉÊÍÓÔÕÚÜÇ'\s]+?)(?:\s+UF|\n)", f, re.I)
        if m:
            campos["naturalidade"] = m.group(1).strip().rstrip(" -")
        else:
            m = re.search(r"Local do nascimento[^\n]*\n\s*([A-ZÀÁÂÃÉÊÍÓÔÕÚÜÇ'\s]+?)\s*[-–]\s*([A-Z]{2})", f, re.I)
            if m:
                campos["naturalidade"] = m.group(1).strip()

        # --- Estado Civil ---
        m = re.search(r"Estado\s*[Cc]ivil[:\s]*(Solteiro|Casado|Divorciado|Vi[uú]vo|Uni[ãa]o Est[áa]vel)", f, re.I)
        if m:
            campos["estado_civil"] = m.group(1).strip()

        # --- Sexo ---
        m = re.search(r"Sexo[:\s]*(Feminino|Masculino|F|M)\b", f, re.I)
        if m:
            val = m.group(1).strip()
            if val.upper() == "F":
                campos["sexo"] = "Feminino"
            elif val.upper() == "M":
                campos["sexo"] = "Masculino"
            else:
                campos["sexo"] = val

        # --- Raça/Cor ---
        m = re.search(r"(?:Etnia ou Ra[çc]a|Cor)[:\s]*(Branca?|Preta?|Parda?|Ind[ií]gena|Amarela?)\b", f, re.I)
        if m:
            val = m.group(1).strip()
            if val.lower().startswith("pard"):
                campos["raca_cor"] = "Parda"
            elif val.lower().startswith("branc"):
                campos["raca_cor"] = "Branca"
            elif val.lower().startswith("pret"):
                campos["raca_cor"] = "Preta"
            else:
                campos["raca_cor"] = val.capitalize()

        # --- Grau Instrução ---
        m = re.search(r"Instru[çc][aã]o[:\s]*(Ensino[^\n,]+|Analfabeto|Fundamental[^\n]*|M[ée]dio[^\n]*|Superior[^\n]*)", f, re.I)
        if m:
            campos["grau_instrucao"] = m.group(1).strip()

        # --- Filiação Pai ---
        m = re.search(r"Pai[:\s]*([A-ZÀÁÂÃÉÊÍÓÔÕÚÜÇ'\s]+?)(?:\n|Mãe|FILIA)", f)
        if m:
            val = m.group(1).strip()
            if val and val != "FILIAÇÃO":
                campos["filiacao_pai"] = val

        # --- Filiação Mãe ---
        m = re.search(r"M[ãa]e[:\s]*([A-ZÀÁÂÃÉÊÍÓÔÕÚÜÇ'\s]+?)(?:\n|C[eé]dula|CTPS|RG)", f)
        if m:
            val = m.group(1).strip()
            if val:
                campos["filiacao_mae"] = val

        # --- Endereço (Residência) ---
        m = re.search(r"Resid[eê]ncia\s*\n([\s\S]+?)(?:\n\s*\n|Data de nascimento)", f, re.I)
        if m:
            end_bloco = m.group(1).strip()
            m_cep = re.search(r"CEP[:\s]*(\d{2}\.?\d{3}[-]?\d{3})", end_bloco)
            if m_cep:
                campos["cep"] = re.sub(r"[^0-9]", "", m_cep.group(1))
                end_bloco = re.sub(r"CEP[:\s]*\d{2}\.?\d{3}[-]?\d{3}", "", end_bloco).strip()
            linhas = [l.strip() for l in end_bloco.split("\n") if l.strip()]
            for i, linha in enumerate(reversed(linhas)):
                m_cid = re.search(r",\s*([A-ZÀÁÂÃÉÊÍÓÔÕÚÜÇ'\s]+?)[,-]\s*([A-Z]{2})", linha)
                if m_cid and len(m_cid.group(1).strip()) > 2:
                    campos["cidade"] = m_cid.group(1).strip().rstrip(",")
                    campos["estado"] = m_cid.group(2).strip()
                    idx = len(linhas) - 1 - i
                    linhas = linhas[:idx]
                    break
            endereco_limpo = "\n".join(linhas).strip()
            endereco_limpo = re.sub(r"[-,]?\s*$", "", endereco_limpo)
            campos["endereco"] = endereco_limpo.strip().rstrip(",")
        else:
            m = re.search(r"Endereço[:\s]*([^\n]+)", f, re.I)
            if m:
                end_line = m.group(1).strip()
                end_line = re.sub(r"Código\s+Município:.*", "", end_line).strip()
                campos["endereco"] = end_line.rstrip(",")
            m = re.search(r"Cidade[:\s]*([A-ZÀÁÂÃÉÊÍÓÔÕÚÜÇ'\s]+?)(?:\s+UF|\n)", f, re.I)
            if m:
                campos["cidade"] = m.group(1).strip()
            else:
                m = re.search(r"Cidade[:\s]*([A-ZÀÁÂÃÉÊÍÓÔÕÚÜÇ'\s]+)", f, re.I)
                if m:
                    cidade_raw = m.group(1).strip()
                    cidade_raw = re.sub(r"\s*UF:.*", "", cidade_raw).strip()
                    campos["cidade"] = cidade_raw
            m = re.search(r"UF[:\s]*([A-Z]{2})", f)
            if m:
                campos["estado"] = m.group(1).strip()
            m = re.search(r"CEP[:\s]*(\d{2}\.?\d{3}[-]?\d{3})", f)
            if m:
                campos["cep"] = re.sub(r"[^0-9]", "", m.group(1))

        # --- Telefone ---
        m = re.search(r"(?:Fone|Celular)[:\s]*([\d()\s-]+)", f, re.I)
        if m:
            tel = m.group(1).strip()
            if len(re.sub(r"\D", "", tel)) >= 8:
                campos["telefone"] = tel

        # --- Cargo e CBO ---
        m = re.search(r"CBO/Cargo[:\s]*(\d+)\s*[-]\s*([^\n]+)", f, re.I)
        if m:
            campos["cbo"] = m.group(1).strip()
            campos["cargo"] = m.group(2).strip()
        else:
            m = re.search(r"Cargo\s+Fun[çc][aã]o\s+C\.?B\.?O\.?[^\n]*\n\s*(.+?)\s*$", f, re.M)
            if m:
                data_line = m.group(1).strip()
                m_cbo = re.search(r"(\d{4,})\s*$", data_line)
                if m_cbo:
                    campos["cbo"] = m_cbo.group(1).strip()
                    data_line = data_line[:m_cbo.start()].strip()
                words = data_line.split()
                half = len(words) // 2
                if half > 0:
                    first_half = " ".join(words[:half])
                    second_half = " ".join(words[half:])
                    if first_half == second_half:
                        campos["cargo"] = first_half
                    else:
                        campos["cargo"] = data_line
                elif data_line:
                    campos["cargo"] = data_line
            else:
                m = re.search(r"Fun[çc][aã]o[:\s]*([^\n]+)", f, re.I)
                if m:
                    campos["cargo"] = m.group(1).strip()

        # --- Data Admissão ---
        m = re.search(r"(?:Data de )?Admiss[aã]o[:\s]*(\d{2}/\d{2}/\d{4})", f, re.I)
        if m:
            campos["data_admissao"] = _converte_data(m.group(1))

        # --- Salário ---
        m = re.search(r"(?:Sal[aá]rio|Admiss[aã]o)[^\n]*R\$\s*([\d.,]+)", f, re.I)
        if m:
            campos["salario"] = m.group(1).replace(".", "").replace(",", ".").strip()
        else:
            m = re.search(r"R\$\s*([\d.,]+)", f)
            if m and not campos.get("salario"):
                campos["salario"] = m.group(1).replace(".", "").replace(",", ".")

        if not campos.get("tipo_contrato"):
            campos["tipo_contrato"] = "CLT"

    # ═══ EXTRAÇÃO DO CONTRATO DE EXPERIÊNCIA ═══
    if eh_contrato and txt_contrato:
        c = txt_contrato

        if not campos.get("nome"):
            m = re.search(r"Sr\.?\s*\(a\)\s+([A-ZÀÁÂÃÉÊÍÓÔÕÚÜÇ'\s]+?)(?:,|\s+domiciliado|\s+portador)", c, re.I)
            if m:
                campos["nome"] = m.group(1).strip()
            else:
                m = re.search(r"[,\s]+(?:e)\s+([A-ZÀÁÂÃÉÊÍÓÔÕÚÜÇ'\s]+?)\s+(?:portador|domiciliado)", c, re.I)
                if m:
                    campos["nome"] = m.group(1).strip()

        if not campos.get("ctps"):
            m = re.search(r"CTPS[^:]*N[º°:]*\s*(\d[\d/]*)\s*[Ss][eé]rie[:\s]*(\d+)", c, re.I)
            if m:
                campos["ctps"] = f"{m.group(1).strip()}/{m.group(2).strip()}"
            else:
                m = re.search(r"Carteira Profissional\s*No\.?[:\s]*(\d+)[/ ]+S[eé]rie[:\s]*(\d+)", c, re.I)
                if m:
                    campos["ctps"] = f"{m.group(1).strip()}/{m.group(2).strip()}"

        if not campos.get("cargo"):
            m = re.search(r"(?:fun[çc][aã]o|fun[çc][õo]es)\s+de\s+([A-ZÀÁÂÃÉÊÍÓÔÕÚÜÇ\s]+?)(?:\s+e\s+mais|\s*[,.]|\n)", c, re.I)
            if m:
                campos["cargo"] = m.group(1).strip()

        if not campos.get("salario"):
            m = re.search(r"remunera[çc][aã]o\s*(?:de)?\s*:?\s*R\$\s*([\d.,]+)", c, re.I)
            if m:
                campos["salario"] = m.group(1).replace(".", "").replace(",", ".").strip()

        m = re.search(r"in[ií]cio\s*(?:em)?[:\s]*(\d{2}/\d{2}/\d{4})", c, re.I)
        if m:
            campos["inicio_experiencia"] = _converte_data(m.group(1))
            if not campos.get("data_admissao"):
                campos["data_admissao"] = _converte_data(m.group(1))

        m = re.search(r"t[eé]rmino\s*(?:em)?[:\s]*(\d{2}/\d{2}/\d{4})", c, re.I)
        if m:
            campos["fim_experiencia"] = _converte_data(m.group(1))
        else:
            m = re.search(r"(?:prorrogado|vencer|terminar).*?(\d{2}/\d{2}/\d{4})", c, re.I)
            if m:
                campos["fim_experiencia"] = _converte_data(m.group(1))

        if not campos.get("endereco"):
            m = re.search(r"domiciliado\s+(?:na|no)\s+([^,]+?)\s*,", c, re.I)
            if m:
                campos["endereco"] = m.group(1).strip()

        if not campos.get("cidade"):
            m = re.search(r"cidade\s+de\s+([A-ZÀÁÂÃÉÊÍÓÔÕÚÜÇ'\s]+?)[,-]\s*([A-Z]{2})", c, re.I)
            if m:
                campos["cidade"] = m.group(1).strip()
                if not campos.get("estado"):
                    campos["estado"] = m.group(2).strip()

        if not campos.get("telefone"):
            m = re.search(r"(?:Fone|Telefone|Celular)[:\s]*([\d()\s-]+)", c, re.I)
            if m:
                tel = m.group(1).strip()
                if len(re.sub(r"\D", "", tel)) >= 8:
                    campos["telefone"] = tel

        campos["tipo_contrato"] = "Experiência"

    # ═══ EXTRAÇÃO GENÉRICA ═══
    if not eh_ficha and not eh_contrato:
        m = re.search(r"Nome[:\s]+([A-ZÀÁÂÃÉÊÍÓÔÕÚÜÇ'\s]+)", txt, re.I)
        if m and not campos.get("nome"):
            campos["nome"] = m.group(1).strip()
        m = re.search(r"CPF[:\s]*(\d{3}\.?\d{3}\.?\d{3}[-.]?\d{2})", txt)
        if m and not campos.get("cpf"):
            campos["cpf"] = re.sub(r"[^0-9]", "", m.group(1))

    # Limpeza final
    for k, v in list(campos.items()):
        if isinstance(v, str):
            campos[k] = v.strip().strip(".,;: ")
            if not campos[k]:
                del campos[k]

    campos["_tipo_detectado"] = []
    if eh_ficha:
        campos["_tipo_detectado"].append("Ficha de Registro")
    if eh_contrato:
        campos["_tipo_detectado"].append("Contrato de Experiência")
    if not eh_ficha and not eh_contrato:
        campos["_tipo_detectado"].append("Documento genérico")

    return campos


def extrair_campos_pdf(arquivos_pdf):
    """
    Recebe UM arquivo PDF ou uma LISTA de arquivos PDF (UploadedFile ou BytesIO)
    e retorna dict com os campos extraídos da Ficha de Registro e/ou Contrato de Experiência.
    Aceita: 1 PDF só, vários PDFs separados (Ficha + Contrato), ou o formato
    que junta tudo em um único arquivo (DOCS ADMISSIONAIS).
    """
    if not TEM_LEITOR_PDF:
        return {"_erro": "Biblioteca de leitura de PDF não disponível neste servidor."
                " Instale pdfplumber ou pypdf (pip install pdfplumber pypdf)."}

    # Normaliza para lista
    if not isinstance(arquivos_pdf, (list, tuple)):
        arquivos_pdf = [arquivos_pdf]

    # Coleta todas as páginas de todos os PDFs
    todas_paginas = []
    erros = []
    for arq in arquivos_pdf:
        try:
            paginas = _ler_pdf_paginas(arq)
            if paginas:
                todas_paginas.extend(paginas)
            else:
                erros.append("PDF vazio ou ilegível")
        except Exception as e:
            erros.append(str(e))

    if not todas_paginas:
        msg = "Nenhum texto encontrado nos PDFs enviados."
        if erros:
            msg += f" Erros: {'; '.join(erros)}"
        return {"_erro": msg}

    return _extrair_campos_de_paginas(todas_paginas)


def _converte_data(data_br):
    """Converte DD/MM/AAAA para AAAA-MM-DD."""
    if not data_br:
        return ""
    try:
        return datetime.strptime(data_br.strip(), "%d/%m/%Y").strftime("%Y-%m-%d")
    except (ValueError, TypeError):
        return data_br


# ════════════════════════════════════════════════════════════════
# FUNÇÕES AUXILIARES BANCO
# ════════════════════════════════════════════════════════════════

def _agora():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def _hoje():
    return date.today().strftime("%Y-%m-%d")

def _ler_config(chave, padrao=""):
    con = _conn(); r = con.execute("SELECT valor FROM config WHERE chave=?", (chave,)).fetchone(); con.close()
    return r["valor"] if r else padrao

def _gravar_config(chave, valor):
    con = _conn(); con.execute("INSERT OR REPLACE INTO config(chave,valor) VALUES(?,?)", (chave,valor)); con.commit(); con.close()

def _lista_lojas():
    con = _conn()
    rows = con.execute("SELECT nome FROM lojas WHERE ativa=1 ORDER BY nome").fetchall()
    con.close()
    return [r["nome"] for r in rows]

def _lista_funcionarios_ativos():
    con = _conn()
    rows = con.execute("SELECT id, nome FROM funcionarios WHERE situacao='Ativo' ORDER BY nome").fetchall()
    con.close()
    return [(r["id"], r["nome"]) for r in rows]

def _parse_data(val):
    if not val: return None
    try: return datetime.strptime(val[:10], "%Y-%m-%d").date()
    except: return None

def _fmt_data(val):
    if not val: return ""
    try: return datetime.strptime(val[:10], "%Y-%m-%d").strftime("%d/%m/%Y")
    except: return val

# ════════════════════════════════════════════════════════════════
# CONFIGURAÇÃO DA PÁGINA
# ════════════════════════════════════════════════════════════════

st.set_page_config(
    page_title="Sistema RH",
    page_icon="👥",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ════════════════════════════════════════════════════════════════
# SIDEBAR
# ════════════════════════════════════════════════════════════════

with st.sidebar:
    st.markdown("## 🏢 Sistema RH")
    st.success("✅ Banco de dados integrado — dados salvos no próprio sistema")
    st.caption("Nenhum arquivo externo. Nenhuma configuração de pasta.\nTudo fica no banco de dados do próprio sistema.")

    abas_disponiveis = [
        "👥 Cadastro",
        "🏢 Lojas",
        "✈️ Diárias",
        "🧳 Viagens",
        "🛒 Compras",
        "📊 Painel",
        "📎 Documentos",
        "⏱️ Eventos Trabalhistas",
        "⚙️ Configurações",
        "💾 Backup"
    ]
    aba_sel = st.radio("Navegação", abas_disponiveis, index=0)
    st.markdown("---")
    st.caption(f"Versão 3.0 — Leitura automática de PDF — {_agora()}")

# ════════════════════════════════════════════════════════════════
# ABA 1 — CADASTRO DE FUNCIONÁRIOS
# ════════════════════════════════════════════════════════════════

if aba_sel == "👥 Cadastro":
    st.header("👥 Cadastro de Funcionários")

    con = _conn()
    func_rows = con.execute("SELECT * FROM funcionarios ORDER BY nome").fetchall()
    con.close()

    lista_nomes = [f"{r['id']} — {r['nome']}" for r in func_rows]

    col_a, col_b = st.columns([1, 2])
    with col_a:
        selecionado = st.selectbox("Selecione um funcionário", ["➕ Novo"] + lista_nomes)

        if selecionado != "➕ Novo":
            fid = int(selecionado.split(" — ")[0])
            con = _conn()
            f = con.execute("SELECT * FROM funcionarios WHERE id=?", (fid,)).fetchone()
            con.close()
        else:
            f = None

    with col_b:
        # ── UPLOAD DE PDF PARA PREENCHIMENTO AUTOMÁTICO ──
        st.subheader("📄 Importar do PDF de Admissão")
        st.caption("Envie a Ficha de Registro de Empregado e/ou o Contrato de Experiência. "
                   "Pode enviar os dois separados ou tudo junto em um único arquivo.")

        pdf_uploads = st.file_uploader(
            "📎 Enviar PDF(s) de admissão",
            type=["pdf"],
            key="up_pdf_admissao",
            accept_multiple_files=True,
            help="Ficha de Registro, Contrato de Experiência, ou os dois juntos (DOCS ADMISSIONAIS).\n"
                 "Pode enviar mais de um arquivo — os dados serão mesclados automaticamente."
        )

        campos_pdf = {}
        if pdf_uploads:
            with st.spinner(f"Lendo {len(pdf_uploads)} PDF(s)..."):
                campos_pdf = extrair_campos_pdf(pdf_uploads)

            if "_erro" in campos_pdf:
                st.error(campos_pdf["_erro"])
                campos_pdf = {}
            elif campos_pdf:
                tipos = ", ".join(campos_pdf.get("_tipo_detectado", ["PDF"]))
                qtd = len([k for k in campos_pdf if not k.startswith("_")])
                st.success(f"✅ {qtd} campo(s) extraído(s) de {len(pdf_uploads)} arquivo(s) — {tipos}")

                with st.expander("📋 Campos extraídos do(s) PDF(s)", expanded=True):
                    campos_lista = {k: v for k, v in campos_pdf.items() if not k.startswith("_")}
                    for k, v in campos_lista.items():
                        rotulo = k.replace("_", " ").title()
                        st.write(f"**{rotulo}**: {v}")

        # ── FORMULÁRIO DE CADASTRO ──
        with st.form("form_funcionario"):
            st.subheader("Dados Pessoais")
            c1, c2, c3 = st.columns(3)

            nome_val = campos_pdf.get("nome", f["nome"] if f else "")
            cpf_val = campos_pdf.get("cpf", f["cpf"] if f else "")
            rg_val = campos_pdf.get("rg", f["rg"] if f else "")

            nome = c1.text_input("Nome *", value=nome_val)
            cpf = c2.text_input("CPF", value=cpf_val)
            rg = c3.text_input("RG", value=rg_val)

            nasc_val = campos_pdf.get("data_nascimento", f["data_nascimento"] if f else "")
            nat_val = campos_pdf.get("naturalidade", f["naturalidade"] if f and "naturalidade" in f.keys() else "")
            ec_val = campos_pdf.get("estado_civil", f["estado_civil"] if f else "")

            data_nascimento = c1.text_input("Data Nascimento (AAAA-MM-DD)", value=nasc_val)
            naturalidade = c2.text_input("Naturalidade", value=nat_val)

            ec_opcoes = ["Solteiro","Casado","Divorciado","Viúvo","União Estável"]
            ec_idx = ec_opcoes.index(ec_val) if ec_val in ec_opcoes else (ec_opcoes.index(f["estado_civil"]) if f and f["estado_civil"] else 0)
            estado_civil = c3.selectbox("Estado Civil", ec_opcoes, index=ec_idx)

            sexo_val = campos_pdf.get("sexo", f["sexo"] if f and "sexo" in f.keys() else "")
            raca_val = campos_pdf.get("raca_cor", f["raca_cor"] if f and "raca_cor" in f.keys() else "")
            gi_val = campos_pdf.get("grau_instrucao", f["grau_instrucao"] if f and "grau_instrucao" in f.keys() else "")

            sexo_opcoes = ["", "Feminino", "Masculino"]
            sexo_idx = sexo_opcoes.index(sexo_val) if sexo_val in sexo_opcoes else 0
            sexo = c1.selectbox("Sexo", sexo_opcoes, index=sexo_idx)

            raca_opcoes = ["", "Branca", "Preta", "Parda", "Amarela", "Indígena"]
            raca_idx = raca_opcoes.index(raca_val) if raca_val in raca_opcoes else 0
            raca_cor = c2.selectbox("Raça/Cor", raca_opcoes, index=raca_idx)

            gi_opcoes = ["", "Analfabeto", "Ensino Fundamental Incompleto", "Ensino Fundamental Completo",
                         "Ensino Médio Incompleto", "Ensino Médio Completo", "Ensino Superior Incompleto",
                         "Ensino Superior Completo", "Pós-Graduação"]
            gi_idx = gi_opcoes.index(gi_val) if gi_val in gi_opcoes else 0
            grau_instrucao = c3.selectbox("Grau Instrução", gi_opcoes, index=gi_idx)

            end_val = campos_pdf.get("endereco", f["endereco"] if f else "")
            cid_val = campos_pdf.get("cidade", f["cidade"] if f else "")
            est_val = campos_pdf.get("estado", f["estado"] if f else "")
            cep_val = campos_pdf.get("cep", f["cep"] if f else "")
            tel_val = campos_pdf.get("telefone", f["telefone"] if f else "")

            endereco = c1.text_input("Endereço", value=end_val)
            cidade = c2.text_input("Cidade", value=cid_val)
            estado = c3.text_input("UF", value=est_val)
            cep = c1.text_input("CEP", value=cep_val)
            telefone = c2.text_input("Telefone", value=tel_val)

            ctps_val = campos_pdf.get("ctps", f["ctps"] if f and "ctps" in f.keys() else "")
            pis_val = campos_pdf.get("pis", f["pis"] if f and "pis" in f.keys() else "")
            cbo_val = campos_pdf.get("cbo", f["cbo"] if f and "cbo" in f.keys() else "")

            ctps = c3.text_input("CTPS (Nº/Série)", value=ctps_val)
            pis = c1.text_input("PIS/PASEP", value=pis_val)
            cbo = c2.text_input("CBO", value=cbo_val)

            pai_val = campos_pdf.get("filiacao_pai", f["filiacao_pai"] if f and "filiacao_pai" in f.keys() else "")
            mae_val = campos_pdf.get("filiacao_mae", f["filiacao_mae"] if f and "filiacao_mae" in f.keys() else "")

            filiacao_pai = c3.text_input("Filiação (Pai)", value=pai_val)
            filiacao_mae = c1.text_input("Filiação (Mãe)", value=mae_val)

            st.markdown("---")
            st.subheader("Dados Profissionais")
            c1, c2, c3 = st.columns(3)

            cargo_val = campos_pdf.get("cargo", f["cargo"] if f else "")
            adm_val = campos_pdf.get("data_admissao", f["data_admissao"] if f else "")
            sal_val = campos_pdf.get("salario", "")
            if not sal_val and f and f["salario"]:
                sal_val = str(float(f["salario"]))

            cargo = c1.text_input("Cargo", value=cargo_val)
            lojas_lista = _lista_lojas()
            loja_sel = c2.selectbox("Loja", lojas_lista if lojas_lista else ["Sem loja cadastrada"],
                index=lojas_lista.index(f["loja"]) if f and f["loja"] in lojas_lista else 0)
            data_admissao = c3.text_input("Data Admissão (AAAA-MM-DD)", value=adm_val)

            sal_float = 0.0
            if sal_val:
                try:
                    sal_float = float(sal_val)
                except (ValueError, TypeError):
                    pass
            elif f and f["salario"]:
                sal_float = float(f["salario"])

            salario = c1.number_input("Salário (R$)", value=sal_float, min_value=0.0, format="%.2f")

            tc_val = campos_pdf.get("tipo_contrato", f["tipo_contrato"] if f else "")
            tc_opcoes = ["CLT","PJ","Estágio","Temporário","Experiência","Outro"]
            tc_idx = tc_opcoes.index(tc_val) if tc_val in tc_opcoes else 0
            tipo_contrato = c2.selectbox("Tipo Contrato", tc_opcoes, index=tc_idx)

            sit_opcoes = ["Ativo","Demitido C/JC","Demitido S/JC","Pedido de Conta",
                          "Término de Contrato","Rescisão","Abandono","Desistência","Aviso Prévio"]
            sit_val = f["situacao"] if f else "Ativo"
            sit_idx = sit_opcoes.index(sit_val) if sit_val in sit_opcoes else 0
            situacao = c3.selectbox("Situação", sit_opcoes, index=sit_idx)

            st.markdown("---")
            st.subheader("Dados Bancários")
            b1, b2 = st.columns(2)
            banco = b1.text_input("Banco", value=f["banco"] if f else "")

            st.markdown("---")
            st.subheader("Período de Experiência")
            e1, e2, e3 = st.columns(3)
            ie_val = campos_pdf.get("inicio_experiencia", f["inicio_experiencia"] if f else "")
            fe_val = campos_pdf.get("fim_experiencia", f["fim_experiencia"] if f else "")
            inicio_experiencia = e1.text_input("Início Experiência", value=ie_val)
            fim_experiencia = e2.text_input("Fim Experiência", value=fe_val)

            st.markdown("---")
            st.subheader("Outros Eventos / Prazos")
            e1, e2, e3 = st.columns(3)
            inicio_ferias = e1.text_input("Início Férias", value=f["inicio_ferias"] if f else "")
            dias_ferias = e2.number_input("Dias Férias", value=int(f["dias_ferias"]) if f and f["dias_ferias"] else 0, min_value=0)
            inicio_licenca = e1.text_input("Início Licença", value=f["inicio_licenca"] if f else "")
            dias_licenca = e2.number_input("Dias Licença", value=int(f["dias_licenca"]) if f and f["dias_licenca"] else 0, min_value=0)
            inicio_afastamento = e1.text_input("Início Afastamento", value=f["inicio_afastamento"] if f else "")
            dias_afastamento = e2.number_input("Dias Afastamento", value=int(f["dias_afastamento"]) if f and f["dias_afastamento"] else 0, min_value=0)
            inicio_aviso = e1.text_input("Início Aviso Prévio", value=f["inicio_aviso"] if f else "")
            dias_aviso = e2.number_input("Dias Aviso", value=int(f["dias_aviso"]) if f and f["dias_aviso"] else 0, min_value=0)

            obs = st.text_area("Observações", value=f["observacoes"] if f else "")

            col_salvar, col_limpar, col_excluir = st.columns(3)
            salvar = col_salvar.form_submit_button("💾 Salvar", type="primary")
            limpar = col_limpar.form_submit_button("🧹 Limpar")
            excluir = col_excluir.form_submit_button("🗑️ Excluir")

            if limpar:
                st.rerun()

            if salvar and not nome.strip():
                st.error("Nome é obrigatório!")

            if salvar and nome.strip():
                # calcular datas finais
                fim_ferias = ""
                if inicio_ferias and dias_ferias:
                    d = _parse_data(inicio_ferias)
                    if d:
                        fim_ferias = (d + timedelta(days=int(dias_ferias))).strftime("%Y-%m-%d")
                fim_licenca = ""
                if inicio_licenca and dias_licenca:
                    d = _parse_data(inicio_licenca)
                    if d:
                        fim_licenca = (d + timedelta(days=int(dias_licenca))).strftime("%Y-%m-%d")
                fim_afastamento = ""
                if inicio_afastamento and dias_afastamento:
                    d = _parse_data(inicio_afastamento)
                    if d:
                        fim_afastamento = (d + timedelta(days=int(dias_afastamento))).strftime("%Y-%m-%d")
                fim_aviso = ""
                if inicio_aviso and dias_aviso:
                    d = _parse_data(inicio_aviso)
                    if d:
                        fim_aviso = (d + timedelta(days=int(dias_aviso))).strftime("%Y-%m-%d")

                con = _conn()
                if f:  # atualizar
                    con.execute("""UPDATE funcionarios SET
                        nome=?, cpf=?, rg=?, data_nascimento=?, estado_civil=?,
                        endereco=?, cidade=?, estado=?, cep=?, telefone=?,
                        cargo=?, loja=?, data_admissao=?, salario=?,
                        banco=?, ctps=?, pis=?, cbo=?,
                        naturalidade=?, sexo=?, raca_cor=?, grau_instrucao=?,
                        filiacao_pai=?, filiacao_mae=?,
                        tipo_contrato=?, situacao=?,
                        inicio_experiencia=?, fim_experiencia=?,
                        inicio_ferias=?, dias_ferias=?, fim_ferias=?,
                        inicio_licenca=?, dias_licenca=?, fim_licenca=?,
                        inicio_afastamento=?, dias_afastamento=?, fim_afastamento=?,
                        inicio_aviso=?, dias_aviso=?, fim_aviso=?,
                        observacoes=?, atualizado_em=?
                        WHERE id=?""",
                        (nome.strip(), cpf, rg, data_nascimento, estado_civil,
                         endereco, cidade, estado, cep, telefone,
                         cargo, loja_sel, data_admissao, salario,
                         banco, ctps, pis, cbo,
                         naturalidade, sexo, raca_cor, grau_instrucao,
                         filiacao_pai, filiacao_mae,
                         tipo_contrato, situacao,
                         inicio_experiencia, fim_experiencia,
                         inicio_ferias, int(dias_ferias), fim_ferias,
                         inicio_licenca, int(dias_licenca), fim_licenca,
                         inicio_afastamento, int(dias_afastamento), fim_afastamento,
                         inicio_aviso, int(dias_aviso), fim_aviso,
                         obs, _agora(), fid))
                    con.commit(); con.close()
                    st.success(f"✅ {nome.strip()} atualizado com sucesso!")
                else:  # novo
                    con.execute("""INSERT INTO funcionarios (
                        nome, cpf, rg, data_nascimento, estado_civil,
                        endereco, cidade, estado, cep, telefone,
                        cargo, loja, data_admissao, salario,
                        banco, ctps, pis, cbo,
                        naturalidade, sexo, raca_cor, grau_instrucao,
                        filiacao_pai, filiacao_mae,
                        tipo_contrato, situacao,
                        inicio_experiencia, fim_experiencia,
                        inicio_ferias, dias_ferias, fim_ferias,
                        inicio_licenca, dias_licenca, fim_licenca,
                        inicio_afastamento, dias_afastamento, fim_afastamento,
                        inicio_aviso, dias_aviso, fim_aviso,
                        observacoes, criado_em, atualizado_em
                    ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                        (nome.strip(), cpf, rg, data_nascimento, estado_civil,
                         endereco, cidade, estado, cep, telefone,
                         cargo, loja_sel, data_admissao, salario,
                         banco, ctps, pis, cbo,
                         naturalidade, sexo, raca_cor, grau_instrucao,
                         filiacao_pai, filiacao_mae,
                         tipo_contrato, situacao,
                         inicio_experiencia, fim_experiencia,
                         inicio_ferias, int(dias_ferias), fim_ferias,
                         inicio_licenca, int(dias_licenca), fim_licenca,
                         inicio_afastamento, int(dias_afastamento), fim_afastamento,
                         inicio_aviso, int(dias_aviso), fim_aviso,
                         obs, _agora(), _agora()))
                    con.commit(); con.close()
                    st.success(f"✅ {nome.strip()} cadastrado com sucesso!")
                st.rerun()

            if excluir and f:
                con = _conn()
                con.execute("DELETE FROM funcionarios WHERE id=?", (fid,))
                con.commit(); con.close()
                st.warning(f"🗑️ {f['nome']} excluído.")
                st.rerun()

    # TABELA RESUMO
    if func_rows:
        st.markdown("---")
        st.subheader("📋 Funcionários Cadastrados")
        df_func = pd.DataFrame([dict(r) for r in func_rows])
        mostrar = ["id","nome","cpf","cargo","loja","situacao","data_admissao"]
        st.dataframe(df_func[[c for c in mostrar if c in df_func.columns]], use_container_width=True, hide_index=True)

# ════════════════════════════════════════════════════════════════
# ABA 2 — LOJAS
# ════════════════════════════════════════════════════════════════

elif aba_sel == "🏢 Lojas":
    st.header("🏢 Cadastro de Lojas")

    con = _conn()
    lojas_rows = con.execute("SELECT * FROM lojas ORDER BY nome").fetchall()
    con.close()

    lista_lojas = [f"{r['id']} — {r['nome']}" for r in lojas_rows]

    col_a, col_b = st.columns([1, 2])
    with col_a:
        sel = st.selectbox("Selecione uma loja", ["➕ Nova"] + lista_lojas)
        if sel != "➕ Nova":
            lid = int(sel.split(" — ")[0])
            con = _conn()
            loja = con.execute("SELECT * FROM lojas WHERE id=?", (lid,)).fetchone()
            con.close()
        else:
            loja = None

    with col_b:
        with st.form("form_loja"):
            l1, l2, l3 = st.columns(3)
            nome_loja = l1.text_input("Nome da Loja *", value=loja["nome"] if loja else "")
            endereco_loja = l2.text_input("Endereço", value=loja["endereco"] if loja else "")
            cidade_loja = l3.text_input("Cidade", value=loja["cidade"] if loja else "")
            estado_loja = l1.text_input("Estado", value=loja["estado"] if loja else "")
            tel_loja = l2.text_input("Telefone", value=loja["telefone"] if loja else "")
            resp_loja = l3.text_input("Responsável", value=loja["responsavel"] if loja else "")
            ativa = l1.checkbox("Ativa", value=bool(loja["ativa"]) if loja else True)
            obs_loja = st.text_area("Observações", value=loja["observacoes"] if loja else "")

            c_s, c_l, c_e = st.columns(3)
            bs = c_s.form_submit_button("💾 Salvar", type="primary")
            bl = c_l.form_submit_button("🧹 Limpar")
            be = c_e.form_submit_button("🗑️ Excluir")

            if bl:
                st.rerun()

            if bs and not nome_loja.strip():
                st.error("Nome da loja é obrigatório!")

            if bs and nome_loja.strip():
                con = _conn()
                if loja:
                    con.execute("""UPDATE lojas SET nome=?, endereco=?, cidade=?, estado=?,
                        telefone=?, responsavel=?, ativa=?, observacoes=?, atualizado_em=?
                        WHERE id=?""",
                        (nome_loja.strip(), endereco_loja, cidade_loja, estado_loja,
                         tel_loja, resp_loja, 1 if ativa else 0, obs_loja, _agora(), lid))
                else:
                    con.execute("""INSERT INTO lojas (nome, endereco, cidade, estado,
                        telefone, responsavel, ativa, observacoes, criado_em, atualizado_em)
                        VALUES (?,?,?,?,?,?,?,?,?,?)""",
                        (nome_loja.strip(), endereco_loja, cidade_loja, estado_loja,
                         tel_loja, resp_loja, 1 if ativa else 0, obs_loja, _agora(), _agora()))
                con.commit(); con.close()
                st.success(f"✅ Loja '{nome_loja.strip()}' salva!")
                st.rerun()

            if be and loja:
                con = _conn()
                con.execute("DELETE FROM lojas WHERE id=?", (lid,))
                con.commit(); con.close()
                st.warning(f"🗑️ Loja '{loja['nome']}' excluída.")
                st.rerun()

    # Upload de documentos da loja
    if loja:
        st.markdown("---")
        st.subheader(f"📎 Documentos da Loja: {loja['nome']}")

        arq = st.file_uploader("Enviar documento", type=["pdf","jpg","jpeg","png","doc","docx"], key="up_doc_loja")
        if arq:
            conteudo = arq.read()
            con = _conn()
            con.execute("INSERT INTO documentos_loja (loja_id, tipo, descricao, nome_arquivo, conteudo, tamanho, criado_em) VALUES (?,?,?,?,?,?,?)",
                (lid, arq.type, arq.name, arq.name, conteudo, len(conteudo), _agora()))
            con.commit(); con.close()
            st.success(f"✅ '{arq.name}' salvo no sistema!")
            st.rerun()

        con = _conn()
        docs = con.execute("SELECT * FROM documentos_loja WHERE loja_id=? ORDER BY criado_em DESC", (lid,)).fetchall()
        con.close()
        if docs:
            for doc in docs:
                c1, c2 = st.columns([3, 1])
                with c1:
                    st.write(f"📄 **{doc['nome_arquivo']}** — {_fmt_data(doc['criado_em'])}")
                with c2:
                    b64 = base64.b64encode(doc['conteudo']).decode()
                    ext = doc['nome_arquivo'].rsplit('.', 1)[-1].lower()
                    mime = "application/pdf" if ext == "pdf" else "image/jpeg" if ext in ("jpg","jpeg") else "image/png" if ext == "png" else "application/octet-stream"
                    st.markdown(f'<a href="data:{mime};base64,{b64}" download="{doc["nome_arquivo"]}">⬇️ Baixar</a>', unsafe_allow_html=True)
                with c2:
                    if st.button("🗑️", key=f"exc_doc_loja_{doc['id']}"):
                        con = _conn()
                        con.execute("DELETE FROM documentos_loja WHERE id=?", (doc['id'],))
                        con.commit(); con.close()
                        st.rerun()

    # LISTA DE LOJAS
    if lojas_rows:
        st.markdown("---")
        st.subheader("📋 Lojas Cadastradas")
        df_lojas = pd.DataFrame([dict(r) for r in lojas_rows])
        cols = ["id","nome","cidade","estado","telefone","responsavel","ativa"]
        st.dataframe(df_lojas[[c for c in cols if c in df_lojas.columns]], use_container_width=True, hide_index=True)

# ════════════════════════════════════════════════════════════════
# ABA 3 — DIÁRIAS
# ════════════════════════════════════════════════════════════════

elif aba_sel == "✈️ Diárias":
    st.header("✈️ Controle de Diárias")

    func_ativos = _lista_funcionarios_ativos()
    nomes_ativos = [f[1] for f in func_ativos]

    with st.form("form_diaria"):
        d1, d2, d3 = st.columns(3)
        func_nome = d1.selectbox("Funcionário *", nomes_ativos if nomes_ativos else ["Nenhum cadastrado"])
        destino = d2.text_input("Destino *")
        motivo = d3.text_input("Motivo")
        data_saida = d1.text_input("Data Saída (AAAA-MM-DD)")
        data_retorno = d2.text_input("Data Retorno (AAAA-MM-DD)")
        valor_diaria = d3.number_input("Valor Diária (R$)", value=150.0, min_value=0.0, format="%.2f")
        qtd_dias = d1.number_input("Quantidade de Dias", value=1, min_value=1)
        obs_diaria = d2.text_area("Observações")

        fid_diaria = None
        for f in func_ativos:
            if f[1] == func_nome:
                fid_diaria = f[0]
                break

        cd1, cd2 = st.columns(2)
        b_salvar = cd1.form_submit_button("💾 Salvar Diária", type="primary")
        b_limpar = cd2.form_submit_button("🧹 Limpar")

        if b_limpar:
            st.rerun()

        if b_salvar and not destino.strip():
            st.error("Destino é obrigatório!")

        if b_salvar and destino.strip() and fid_diaria:
            valor_total = valor_diaria * qtd_dias
            con = _conn()
            con.execute("""INSERT INTO diarias
                (funcionario_id, funcionario_nome, destino, motivo,
                 data_saida, data_retorno, valor_diaria, quantidade,
                 valor_total, observacoes, status, criado_em, atualizado_em)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (fid_diaria, func_nome, destino.strip(), motivo,
                 data_saida, data_retorno, valor_diaria, int(qtd_dias),
                 valor_total, obs_diaria, "Pendente", _agora(), _agora()))
            con.commit(); con.close()
            st.success(f"✅ Diária para {func_nome} salva! Total: R$ {valor_total:.2f}")
            st.rerun()

    # LISTA DE DIÁRIAS
    st.markdown("---")
    st.subheader("📋 Diárias Registradas")
    con = _conn()
    diarias = con.execute("SELECT * FROM diarias ORDER BY criado_em DESC").fetchall()
    con.close()
    if diarias:
        df_d = pd.DataFrame([dict(r) for r in diarias])
        cols_d = ["id","funcionario_nome","destino","data_saida","data_retorno","valor_diaria","quantidade","valor_total","status"]
        st.dataframe(df_d[[c for c in cols_d if c in df_d.columns]], use_container_width=True, hide_index=True)
    else:
        st.info("Nenhuma diária registrada.")

    # Comprovantes
    st.markdown("---")
    st.subheader("📎 Comprovantes de Diárias")
    if diarias:
        sel_diaria = st.selectbox("Escolha a diária", [f"{d['id']} — {d['funcionario_nome']} → {d['destino']}" for d in diarias])
        did = int(sel_diaria.split(" — ")[0])

        arq_comp = st.file_uploader("Enviar comprovante", type=["pdf","jpg","jpeg","png"], key="up_comp_diaria")
        if arq_comp:
            conteudo = arq_comp.read()
            con = _conn()
            con.execute("INSERT INTO comprovantes_diaria (diaria_id, nome_arquivo, conteudo, tamanho, criado_em) VALUES (?,?,?,?,?)",
                (did, arq_comp.name, conteudo, len(conteudo), _agora()))
            con.commit(); con.close()
            st.success(f"✅ '{arq_comp.name}' salvo!")
            st.rerun()

        con = _conn()
        comps = con.execute("SELECT * FROM comprovantes_diaria WHERE diaria_id=? ORDER BY criado_em DESC", (did,)).fetchall()
        con.close()
        if comps:
            for comp in comps:
                c1, c2 = st.columns([3, 1])
                with c1:
                    st.write(f"📄 **{comp['nome_arquivo']}** — {_fmt_data(comp['criado_em'])}")
                with c2:
                    b64 = base64.b64encode(comp['conteudo']).decode()
                    ext = comp['nome_arquivo'].rsplit('.', 1)[-1].lower()
                    mime = "application/pdf" if ext == "pdf" else "image/jpeg" if ext in ("jpg","jpeg") else "image/png"
                    st.markdown(f'<a href="data:{mime};base64,{b64}" download="{comp["nome_arquivo"]}">⬇️ Baixar</a>', unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════════
# ABA 4 — VIAGENS
# ════════════════════════════════════════════════════════════════

elif aba_sel == "🧳 Viagens":
    st.header("🧳 Registro de Viagens")

    func_ativos = _lista_funcionarios_ativos()
    nomes_ativos = [f[1] for f in func_ativos]

    with st.form("form_viagem"):
        v1, v2, v3 = st.columns(3)
        func_nome_v = v1.selectbox("Funcionário *", nomes_ativos if nomes_ativos else ["Nenhum cadastrado"], key="sel_func_v")
        destino_v = v2.text_input("Destino *")
        motivo_v = v3.text_input("Motivo")
        data_saida_v = v1.text_input("Data Saída (AAAA-MM-DD)")
        data_retorno_v = v2.text_input("Data Retorno (AAAA-MM-DD)")
        transporte_v = v3.selectbox("Transporte", ["Avião","Ônibus","Carro","Outro"])
        hospedagem_v = v1.text_input("Hospedagem")
        valor_v = v2.number_input("Valor Estimado (R$)", value=0.0, min_value=0.0, format="%.2f")
        obs_v = v3.text_area("Observações")

        fid_v = None
        for f in func_ativos:
            if f[1] == func_nome_v:
                fid_v = f[0]
                break

        bv1, bv2 = st.columns(2)
        bsv = bv1.form_submit_button("💾 Salvar Viagem", type="primary")
        blv = bv2.form_submit_button("🧹 Limpar")

        if blv:
            st.rerun()

        if bsv and not destino_v.strip():
            st.error("Destino é obrigatório!")

        if bsv and destino_v.strip() and fid_v:
            con = _conn()
            con.execute("""INSERT INTO viagens
                (funcionario_id, funcionario_nome, destino, motivo,
                 data_saida, data_retorno, transporte, hospedagem,
                 valor_estimado, observacoes, status, criado_em, atualizado_em)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (fid_v, func_nome_v, destino_v.strip(), motivo_v,
                 data_saida_v, data_retorno_v, transporte_v, hospedagem_v,
                 valor_v, obs_v, "Pendente", _agora(), _agora()))
            con.commit(); con.close()
            st.success(f"✅ Viagem de {func_nome_v} registrada!")
            st.rerun()

    st.markdown("---")
    st.subheader("📋 Viagens Registradas")
    con = _conn()
    viagens = con.execute("SELECT * FROM viagens ORDER BY criado_em DESC").fetchall()
    con.close()
    if viagens:
        df_v = pd.DataFrame([dict(r) for r in viagens])
        cols_v = ["id","funcionario_nome","destino","data_saida","data_retorno","transporte","valor_estimado","status"]
        st.dataframe(df_v[[c for c in cols_v if c in df_v.columns]], use_container_width=True, hide_index=True)
    else:
        st.info("Nenhuma viagem registrada.")

# ════════════════════════════════════════════════════════════════
# ABA 5 — COMPRAS
# ════════════════════════════════════════════════════════════════

elif aba_sel == "🛒 Compras":
    st.header("🛒 Controle de Compras")

    lojas_lista = _lista_lojas()

    with st.form("form_compra"):
        co1, co2, co3 = st.columns(3)
        loja_compra = co1.selectbox("Loja *", lojas_lista if lojas_lista else ["Nenhuma cadastrada"])
        fornecedor = co2.text_input("Fornecedor *")
        descricao = co3.text_input("Descrição *")
        qtd = co1.number_input("Quantidade", value=1, min_value=1)
        valor_unit = co2.number_input("Valor Unitário (R$)", value=0.0, min_value=0.0, format="%.2f")
        cat = co3.selectbox("Categoria", ["Material","Serviço","Equipamento","Outro"])
        data_compra = co1.text_input("Data Compra (AAAA-MM-DD)", value=_hoje())
        obs_c = co2.text_area("Observações")

        lid_c = None
        if lojas_lista:
            con = _conn()
            r = con.execute("SELECT id FROM lojas WHERE nome=?", (loja_compra,)).fetchone()
            if r: lid_c = r["id"]
            con.close()

        bc1, bc2 = st.columns(2)
        bsc = bc1.form_submit_button("💾 Salvar Compra", type="primary")
        blc = bc2.form_submit_button("🧹 Limpar")

        if blc:
            st.rerun()

        if bsc and (not fornecedor.strip() or not descricao.strip()):
            st.error("Fornecedor e Descrição são obrigatórios!")

        if bsc and fornecedor.strip() and descricao.strip() and lid_c:
            valor_total = qtd * valor_unit
            con = _conn()
            con.execute("""INSERT INTO compras
                (loja_id, loja_nome, fornecedor, descricao, quantidade,
                 valor_unitario, valor_total, data_compra, categoria,
                 status, observacoes, criado_em, atualizado_em)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (lid_c, loja_compra, fornecedor.strip(), descricao.strip(),
                 int(qtd), valor_unit, valor_total, data_compra, cat,
                 "Pendente", obs_c, _agora(), _agora()))
            con.commit(); con.close()
            st.success(f"✅ Compra salva! Total: R$ {valor_total:.2f}")
            st.rerun()

    st.markdown("---")
    st.subheader("📋 Compras Registradas")
    con = _conn()
    compras_rows = con.execute("SELECT * FROM compras ORDER BY criado_em DESC").fetchall()
    con.close()
    if compras_rows:
        df_c = pd.DataFrame([dict(r) for r in compras_rows])
        cols_c = ["id","loja_nome","fornecedor","descricao","quantidade","valor_unitario","valor_total","data_compra","status"]
        st.dataframe(df_c[[c for c in cols_c if c in df_c.columns]], use_container_width=True, hide_index=True)
    else:
        st.info("Nenhuma compra registrada.")

# ════════════════════════════════════════════════════════════════
# ABA 6 — PAINEL / DASHBOARD
# ════════════════════════════════════════════════════════════════

elif aba_sel == "📊 Painel":
    st.header("📊 Painel de Controle")

    con = _conn()
    total_func = con.execute("SELECT COUNT(*) as t FROM funcionarios").fetchone()["t"]
    ativos = con.execute("SELECT COUNT(*) as t FROM funcionarios WHERE situacao='Ativo'").fetchone()["t"]
    total_lojas = con.execute("SELECT COUNT(*) as t FROM lojas WHERE ativa=1").fetchone()["t"]
    total_diarias = con.execute("SELECT COUNT(*) as t FROM diarias").fetchone()["t"]
    total_viagens = con.execute("SELECT COUNT(*) as t FROM viagens").fetchone()["t"]
    total_compras = con.execute("SELECT COUNT(*) as t FROM compras").fetchone()["t"]
    valor_diarias = con.execute("SELECT COALESCE(SUM(valor_total),0) as t FROM diarias").fetchone()["t"]
    valor_compras = con.execute("SELECT COALESCE(SUM(valor_total),0) as t FROM compras").fetchone()["t"]
    con.close()

    p1, p2, p3 = st.columns(3)
    p1.metric("👥 Funcionários", f"{ativos} ativos / {total_func} total")
    p2.metric("🏢 Lojas Ativas", total_lojas)
    p3.metric("✈️ Diárias", f"{total_diarias} (R$ {valor_diarias:,.2f})")

    p4, p5, p6 = st.columns(3)
    p4.metric("🧳 Viagens", total_viagens)
    p5.metric("🛒 Compras", f"{total_compras} (R$ {valor_compras:,.2f})")
    p6.metric("📎 Documentos", "—")

    st.markdown("---")

    con = _conn()
    funcs_por_loja = con.execute("""
        SELECT loja, COUNT(*) as qtd FROM funcionarios
        WHERE situacao='Ativo' AND loja IS NOT NULL AND loja != ''
        GROUP BY loja ORDER BY qtd DESC
    """).fetchall()
    con.close()

    if funcs_por_loja:
        df_graf = pd.DataFrame([dict(r) for r in funcs_por_loja])
        st.bar_chart(df_graf.set_index("loja")["qtd"], use_container_width=True)
        st.caption("Funcionários por Loja")

    con = _conn()
    sit_por_situacao = con.execute("""
        SELECT situacao, COUNT(*) as qtd FROM funcionarios GROUP BY situacao ORDER BY qtd DESC
    """).fetchall()
    con.close()
    if sit_por_situacao:
        df_sit = pd.DataFrame([dict(r) for r in sit_por_situacao])
        st.dataframe(df_sit, use_container_width=True, hide_index=True)
        st.caption("Situação dos Funcionários")

# ════════════════════════════════════════════════════════════════
# ABA 7 — DOCUMENTOS
# ════════════════════════════════════════════════════════════════

elif aba_sel == "📎 Documentos":
    st.header("📎 Documentos e Anexos")

    tab_func, tab_loja = st.tabs(["👤 Funcionários", "🏢 Lojas"])

    with tab_func:
        func_ativos = _lista_funcionarios_ativos()
        nomes = [f[1] for f in func_ativos]
        if nomes:
            sel_func_doc = st.selectbox("Funcionário", nomes, key="sel_func_doc")
            fid_doc = None
            for f in func_ativos:
                if f[1] == sel_func_doc:
                    fid_doc = f[0]
                    break

            st.subheader("📸 Foto do Funcionário")
            foto_up = st.file_uploader("Enviar foto", type=["jpg","jpeg","png"], key="up_foto")
            if foto_up and fid_doc:
                conteudo = foto_up.read()
                con = _conn()
                con.execute("INSERT OR REPLACE INTO fotos_funcionario (funcionario_id, nome_arquivo, conteudo, tamanho, criado_em) VALUES (?,?,?,?,?)",
                    (fid_doc, foto_up.name, conteudo, len(conteudo), _agora()))
                con.commit(); con.close()
                st.success("✅ Foto salva no sistema!")
                st.rerun()

            con = _conn()
            foto_row = con.execute("SELECT * FROM fotos_funcionario WHERE funcionario_id=?", (fid_doc,)).fetchone()
            con.close()
            if foto_row:
                st.image(io.BytesIO(foto_row["conteudo"]), caption=foto_row["nome_arquivo"], width=200)

            st.subheader("📄 Documentos")
            doc_up = st.file_uploader("Enviar documento", type=["pdf","jpg","jpeg","png","doc","docx"], key="up_doc_func")
            if doc_up and fid_doc:
                conteudo = doc_up.read()
                con = _conn()
                con.execute("INSERT INTO documentos (funcionario_id, tipo, descricao, nome_arquivo, conteudo, tamanho, criado_em) VALUES (?,?,?,?,?,?,?)",
                    (fid_doc, doc_up.type, doc_up.name, doc_up.name, conteudo, len(conteudo), _agora()))
                con.commit(); con.close()
                st.success(f"✅ '{doc_up.name}' salvo no sistema!")
                st.rerun()

            con = _conn()
            docs_func = con.execute("SELECT * FROM documentos WHERE funcionario_id=? ORDER BY criado_em DESC", (fid_doc,)).fetchall()
            con.close()
            if docs_func:
                for doc in docs_func:
                    c1, c2 = st.columns([3, 1])
                    with c1:
                        st.write(f"📄 **{doc['nome_arquivo']}** — {_fmt_data(doc['criado_em'])}")
                    with c2:
                        b64 = base64.b64encode(doc['conteudo']).decode()
                        ext = doc['nome_arquivo'].rsplit('.', 1)[-1].lower()
                        mime = "application/pdf" if ext == "pdf" else "image/jpeg" if ext in ("jpg","jpeg") else "image/png" if ext == "png" else "application/octet-stream"
                        dl_name = doc['nome_arquivo']
                        st.markdown(f'<a href="data:{mime};base64,{b64}" download="{dl_name}">⬇️ Baixar</a>', unsafe_allow_html=True)
                    with c2:
                        if st.button("🗑️", key=f"exc_doc_func_{doc['id']}"):
                            con = _conn()
                            con.execute("DELETE FROM documentos WHERE id=?", (doc['id'],))
                            con.commit(); con.close()
                            st.rerun()
        else:
            st.info("Cadastre funcionários primeiro.")

    with tab_loja:
        lojas_lista = _lista_lojas()
        if lojas_lista:
            sel_loja_doc = st.selectbox("Loja", lojas_lista, key="sel_loja_doc")
            con = _conn()
            r = con.execute("SELECT id FROM lojas WHERE nome=?", (sel_loja_doc,)).fetchone()
            lid_doc = r["id"] if r else None
            con.close()

            if lid_doc:
                doc_up_l = st.file_uploader("Enviar documento da loja", type=["pdf","jpg","jpeg","png","doc","docx"], key="up_doc_loja_v")
                if doc_up_l:
                    conteudo = doc_up_l.read()
                    con = _conn()
                    con.execute("INSERT INTO documentos_loja (loja_id, tipo, descricao, nome_arquivo, conteudo, tamanho, criado_em) VALUES (?,?,?,?,?,?,?)",
                        (lid_doc, doc_up_l.type, doc_up_l.name, doc_up_l.name, conteudo, len(conteudo), _agora()))
                    con.commit(); con.close()
                    st.success(f"✅ '{doc_up_l.name}' salvo no sistema!")
                    st.rerun()

                con = _conn()
                docs_loja = con.execute("SELECT * FROM documentos_loja WHERE loja_id=? ORDER BY criado_em DESC", (lid_doc,)).fetchall()
                con.close()
                if docs_loja:
                    for doc in docs_loja:
                        c1, c2 = st.columns([3, 1])
                        with c1:
                            st.write(f"📄 **{doc['nome_arquivo']}** — {_fmt_data(doc['criado_em'])}")
                        with c2:
                            b64 = base64.b64encode(doc['conteudo']).decode()
                            ext = doc['nome_arquivo'].rsplit('.', 1)[-1].lower()
                            mime = "application/pdf" if ext == "pdf" else "image/jpeg" if ext in ("jpg","jpeg") else "image/png" if ext == "png" else "application/octet-stream"
                            dl_name = doc['nome_arquivo']
                            st.markdown(f'<a href="data:{mime};base64,{b64}" download="{dl_name}">⬇️ Baixar</a>', unsafe_allow_html=True)
                        with c2:
                            if st.button("🗑️", key=f"exc_doc_loja_v_{doc['id']}"):
                                con = _conn()
                                con.execute("DELETE FROM documentos_loja WHERE id=?", (doc['id'],))
                                con.commit(); con.close()
                                st.rerun()
        else:
            st.info("Cadastre lojas primeiro.")

# ════════════════════════════════════════════════════════════════
# ABA 8 — EVENTOS TRABALHISTAS
# ════════════════════════════════════════════════════════════════

elif aba_sel == "⏱️ Eventos Trabalhistas":
    st.header("⏱️ Eventos Trabalhistas")

    func_ativos = _lista_funcionarios_ativos()
    nomes_ativos = [f[1] for f in func_ativos]

    if not nomes_ativos:
        st.info("Cadastre funcionários primeiro.")
    else:
        with st.form("form_evento"):
            ev1, ev2 = st.columns(2)
            func_ev = ev1.selectbox("Funcionário *", nomes_ativos, key="sel_func_ev")
            tipo_ev = ev2.selectbox("Tipo de Evento *",
                ["Férias","Aviso Prévio","Licença","Afastamento","Experiência 30d","Experiência 45d","Experiência 60d","Experiência 90d"])
            data_inicio_ev = ev1.text_input("Data Início (AAAA-MM-DD)", value=_hoje())
            dias_ev = ev2.number_input("Dias", value=30, min_value=1)
            obs_ev = st.text_area("Observações")

            fid_ev = None
            for f in func_ativos:
                if f[1] == func_ev:
                    fid_ev = f[0]
                    break

            be1, be2 = st.columns(2)
            bse = be1.form_submit_button("💾 Salvar Evento", type="primary")
            ble = be2.form_submit_button("🧹 Limpar")

            if ble:
                st.rerun()

            if bse and fid_ev:
                d = _parse_data(data_inicio_ev)
                data_fim_ev = (d + timedelta(days=int(dias_ev))).strftime("%Y-%m-%d") if d else ""

                con = _conn()
                con.execute("INSERT INTO eventos_trabalhistas (funcionario_id, tipo, data_inicio, data_fim, dias, observacoes, criado_em) VALUES (?,?,?,?,?,?,?)",
                    (fid_ev, tipo_ev, data_inicio_ev, data_fim_ev, int(dias_ev), obs_ev, _agora()))

                if tipo_ev == "Férias":
                    con.execute("UPDATE funcionarios SET inicio_ferias=?, dias_ferias=?, fim_ferias=?, atualizado_em=? WHERE id=?",
                        (data_inicio_ev, int(dias_ev), data_fim_ev, _agora(), fid_ev))
                elif tipo_ev == "Aviso Prévio":
                    con.execute("UPDATE funcionarios SET inicio_aviso=?, dias_aviso=?, fim_aviso=?, atualizado_em=? WHERE id=?",
                        (data_inicio_ev, int(dias_ev), data_fim_ev, _agora(), fid_ev))
                elif tipo_ev == "Licença":
                    con.execute("UPDATE funcionarios SET inicio_licenca=?, dias_licenca=?, fim_licenca=?, atualizado_em=? WHERE id=?",
                        (data_inicio_ev, int(dias_ev), data_fim_ev, _agora(), fid_ev))
                elif tipo_ev == "Afastamento":
                    con.execute("UPDATE funcionarios SET inicio_afastamento=?, dias_afastamento=?, fim_afastamento=?, atualizado_em=? WHERE id=?",
                        (data_inicio_ev, int(dias_ev), data_fim_ev, _agora(), fid_ev))
                elif "Experiência" in tipo_ev:
                    dias_map = {"30d":30,"45d":45,"60d":60,"90d":90}
                    dias_exp = dias_map.get(tipo_ev.split()[-1], int(dias_ev))
                    data_fim_exp = (d + timedelta(days=dias_exp)).strftime("%Y-%m-%d") if d else ""
                    con.execute("UPDATE funcionarios SET inicio_experiencia=?, fim_experiencia=?, atualizado_em=? WHERE id=?",
                        (data_inicio_ev, data_fim_exp, _agora(), fid_ev))

                con.commit(); con.close()
                st.success(f"✅ Evento '{tipo_ev}' registrado para {func_ev}!")
                st.rerun()

        st.markdown("---")
        st.subheader("📋 Eventos Registrados")
        con = _conn()
        eventos = con.execute("""
            SELECT e.*, f.nome as func_nome FROM eventos_trabalhistas e
            LEFT JOIN funcionarios f ON e.funcionario_id = f.id
            ORDER BY e.criado_em DESC
        """).fetchall()
        con.close()
        if eventos:
            df_ev = pd.DataFrame([dict(r) for r in eventos])
            cols_ev = ["id","func_nome","tipo","data_inicio","data_fim","dias","observacoes"]
            st.dataframe(df_ev[[c for c in cols_ev if c in df_ev.columns]], use_container_width=True, hide_index=True)
        else:
            st.info("Nenhum evento registrado.")

# ════════════════════════════════════════════════════════════════
# ABA 9 — CONFIGURAÇÕES
# ════════════════════════════════════════════════════════════════

elif aba_sel == "⚙️ Configurações":
    st.header("⚙️ Configurações do Sistema")

    st.success("✅ Todos os dados são salvos automaticamente no banco de dados integrado.")
    st.info("Não é preciso configurar pasta externa, Google Sheets ou qualquer outra coisa.\nO banco de dados fica ao lado do programa e nunca se perde ao fechar e reabrir.")

    st.markdown("---")
    st.subheader("📄 Importação de PDFs")
    if TEM_LEITOR_PDF:
        st.success(f"✅ Biblioteca de leitura de PDF disponível ({LEITOR_PDF}) — a importação automática está funcionando.")
    else:
        st.warning("⚠️ Nenhuma biblioteca de leitura de PDF disponível. A importação automática não funcionará.\n"
                    "Instale pdfplumber ou pypdf: `pip install pdfplumber pypdf`")

    st.markdown("---")
    st.subheader("📋 Informações do Banco")

    con = _conn()
    tabelas = con.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name").fetchall()
    con.close()

    for t in tabelas:
        nome = t["name"]
        con = _conn()
        qtd = con.execute(f"SELECT COUNT(*) as t FROM {nome}").fetchone()["t"]
        con.close()
        st.write(f"**{nome}**: {qtd} registro(s)")

    st.markdown("---")
    st.subheader("📊 Tamanho do Banco")
    try:
        tamanho = os.path.getsize(DB_PATH)
        if tamanho > 1048576:
            st.info(f"Tamanho: {tamanho / 1048576:.1f} MB")
        else:
            st.info(f"Tamanho: {tamanho / 1024:.1f} KB")
    except:
        st.warning("Não foi possível verificar o tamanho.")

# ════════════════════════════════════════════════════════════════
# ABA 10 — BACKUP
# ════════════════════════════════════════════════════════════════

elif aba_sel == "💾 Backup":
    st.header("💾 Backup e Restauração")

    st.success("✅ Seus dados estão seguros no banco de dados integrado.\nMesmo assim, recomendamos baixar uma cópia de vez em quando.")

    st.markdown("---")
    st.subheader("📥 Baixar Cópia Completa")
    st.info("O download inclui o banco de dados inteiro (todas as abas, documentos, fotos, comprovantes).\nGuarde em um pen drive ou envie por e-mail como backup extra.")

    if st.button("💾 GERAR BACKUP COMPLETO", type="primary", use_container_width=True):
        with st.spinner("Preparando cópia..."):
            with open(DB_PATH, "rb") as f:
                dados_db = f.read()
            nome_zip = f"backup_rh_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
            buf = io.BytesIO()
            with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
                zf.writestr("rh_dados.db", dados_db)
            buf.seek(0)
            _tam = len(buf.getvalue())
            st.download_button(
                label=f"⬇️ Baixar {nome_zip} ({_tam/1024:.0f} KB)",
                data=buf.getvalue(),
                file_name=nome_zip,
                mime="application/zip",
                use_container_width=True
            )

    st.markdown("---")
    st.subheader("📤 Restaurar de um Backup")
    st.warning("⚠️ Isso substitui TODOS os dados atuais pelos dados do backup.\nFaça um backup novo antes de restaurar, por segurança.")

    arq_rest = st.file_uploader("Enviar arquivo ZIP de backup", type=["zip"], key="rest_backup")
    if arq_rest:
        try:
            with zipfile.ZipFile(io.BytesIO(arq_rest.read()), "r") as zf:
                nomes = zf.namelist()
                if "rh_dados.db" not in nomes:
                    st.error("❌ Este ZIP não contém o banco de dados. Não é um backup válido.")
                else:
                    if st.button("✅ CONFIRMAR RESTAURAÇÃO", type="primary", use_container_width=True):
                        conteudo_db = zf.read("rh_dados.db")
                        with open(DB_PATH, "wb") as f:
                            f.write(conteudo_db)
                        st.success("✅ Backup restaurado com sucesso! Feche e abra o sistema para ver os dados.")
        except Exception as e_rest:
            st.error(f"❌ Erro ao ler o backup: {e_rest}")

    st.markdown("---")
    st.subheader("📂 Exportar para Excel")
    st.info("Exporta os dados de cada aba como arquivo Excel (sem documentos/fotos, só as tabelas).")

    if st.button("📊 GERAR EXCEL", use_container_width=True):
        con = _conn()
        buf_xlsx = io.BytesIO()
        with pd.ExcelWriter(buf_xlsx, engine="openpyxl") as writer:
            for t in ["funcionarios","lojas","diarias","viagens","compras","eventos_trabalhistas"]:
                try:
                    rows = con.execute(f"SELECT * FROM {t}").fetchall()
                    df_exp = pd.DataFrame([dict(r) for r in rows]) if rows else pd.DataFrame()
                    for col_b in ["conteudo","fotos","anexos"]:
                        if col_b in df_exp.columns:
                            df_exp = df_exp.drop(columns=[col_b])
                    df_exp.to_excel(writer, sheet_name=t, index=False)
                except:
                    pd.DataFrame().to_excel(writer, sheet_name=t, index=False)
        con.close()
        buf_xlsx.seek(0)
        st.download_button(
            label="⬇️ Baixar Excel",
            data=buf_xlsx.getvalue(),
            file_name=f"rh_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
