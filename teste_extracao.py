import re, pdfplumber

# ── Copiar funções do app ──

def _ler_pdf_paginas(arquivo_pdf):
    try:
        pdf = pdfplumber.open(arquivo_pdf)
        paginas = [p.extract_text() or '' for p in pdf.pages]
        pdf.close()
        return paginas
    except Exception:
        try:
            from pypdf import PdfReader
            reader = PdfReader(arquivo_pdf)
            return [p.extract_text() or '' for p in reader.pages]
        except Exception:
            return []

def _normalizar_texto_pdf(txt):
    txt = re.sub(r'(\d)([A-Za-z\u00c0-\u00ff])', r'\1 \2', txt)
    txt = re.sub(r'([a-z\u00c0-\u00ff])(\d)', r'\1 \2', txt, flags=re.I)
    txt = re.sub(r':([A-Za-z\u00c0-\u00ff])', r': \1', txt)
    txt = re.sub(r'([a-z\u00e0-\u00ff])([A-Z])', r'\1 \2', txt)
    txt = re.sub(r'[ \t]+', ' ', txt)
    return txt

def _extrair_campo(txt, padrao_rotulo, padrao_valor, flags=re.I):
    m = re.search(padrao_rotulo + padrao_valor, txt, flags)
    if m:
        return m.group(1)
    m2 = re.search(padrao_valor + r'\s*' + padrao_rotulo.rstrip(':'), txt, flags)
    if m2:
        return m2.group(1)
    return None

def _converte_data(val):
    if not val:
        return ''
    val = val.strip()
    m = re.match(r'(\d{2})/(\d{2})/(\d{4})', val)
    if m:
        return f'{m.group(3)}-{m.group(2)}-{m.group(1)}'
    m = re.match(r'(\d{4})-(\d{2})-(\d{2})', val)
    if m:
        return val
    return val

# ── Copiar _extrair_campos_de_paginas ──
app_code = open('/nfs/104761349/temp/app_rh_corrigido.py').read()
start = app_code.find('def _extrair_campos_de_paginas')
end = app_code.find('\ndef extrair_campos_pdf')
fn_code = app_code[start:end]
exec(fn_code)

# ── Testes ──
pdf_files = [
    ('DOCS ADMISSIONAIS', '/nfs/104761349/uploads/DOCS._ADMISSIONAIS.pdf'),
    ('FICHA ANTIGA', '/nfs/104761349/uploads/Ficha_Registro_de_Empregado.pdf'),
    ('CONTRATO', '/nfs/104761349/uploads/Contrato_de_Experiência.pdf'),
]

for nome, path in pdf_files:
    paginas = _ler_pdf_paginas(path)
    campos = _extrair_campos_de_paginas(paginas)
    print(f'\n=== {nome} ===')
    for k, v in sorted(campos.items()):
        print(f'  {k}: {repr(v)}')
