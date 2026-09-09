import streamlit as st
import traceback
import sys
import os

st.set_page_config(page_title="Diagnóstico RH", layout="wide")

st.title("🔧 Diagnóstico do Sistema RH")
st.markdown("Este app verifica se o ambiente Streamlit Cloud suporta todos os recursos necessários.")

resultados = []
erros = []

# Teste 1: Imports básicos
for mod_name, import_stmt in [
    ("streamlit", "import streamlit as st"),
    ("pandas", "import pandas as pd"),
    ("openpyxl", "from openpyxl import load_workbook, Workbook"),
    ("Pillow", "from PIL import Image"),
    ("requests", "import requests"),
    ("matplotlib", "import matplotlib.pyplot as plt"),
    ("reportlab", "from reportlab.lib.pagesizes import A4"),
    ("pydeck", "import pydeck as pdk"),
    ("gTTS", "from gtts import gTTS"),
    ("pydub", "from pydub import AudioSegment"),
    ("gspread", "import gspread"),
    ("google-auth", "from google.oauth2.service_account import Credentials"),
    ("SpeechRecognition", "import speech_recognition as sr"),
    ("folium", "import folium"),
]:
    try:
        exec(import_stmt)
        resultados.append((mod_name, "✅", ""))
    except Exception as e:
        resultados.append((mod_name, "❌", str(e)))
        erros.append(f"{mod_name}: {e}")

# Teste 2: st.fragment
try:
    @st.fragment
    def _test_fragment():
        pass
    resultados.append(("st.fragment", "✅", ""))
except AttributeError:
    resultados.append(("st.fragment", "⚠️", "Não disponível (fallback definido no código)"))
except Exception as e:
    resultados.append(("st.fragment", "❌", str(e)))
    erros.append(f"st.fragment: {e}")

# Teste 3: st.rerun(scope="fragment")
try:
    st.rerun(scope="fragment")
except TypeError:
    resultados.append(("st.rerun(scope=fragment)", "⚠️", "Não suportado (fallback definido no código)"))
except SystemExit:
    resultados.append(("st.rerun(scope=fragment)", "✅", "Funciona (causou rerun)"))
except Exception as e:
    resultados.append(("st.rerun(scope=fragment)", "❌", str(e)))

# Teste 4: Escrita de arquivo no diretório do app
try:
    base_dir = os.path.dirname(os.path.abspath(__file__)) if '__file__' in dir() else os.getcwd()
    teste_path = os.path.join(base_dir, "_teste_escrita.tmp")
    with open(teste_path, "w") as f:
        f.write("ok")
    os.remove(teste_path)
    resultados.append(("Escrita no dir do app", "✅", f"{base_dir}"))
except Exception as e:
    resultados.append(("Escrita no dir do app", "❌", str(e)))
    erros.append(f"Escrita no dir do app: {e}")

# Teste 5: Escrita em /tmp
try:
    teste_path = "/tmp/_teste_escrita.tmp"
    with open(teste_path, "w") as f:
        f.write("ok")
    os.remove(teste_path)
    resultados.append(("Escrita em /tmp", "✅", ""))
except Exception as e:
    resultados.append(("Escrita em /tmp", "❌", str(e)))
    erros.append(f"Escrita em /tmp: {e}")

# Teste 6: Memória disponível
try:
    import psutil
    mem = psutil.virtual_memory()
    resultados.append(("Memória", "ℹ️", f"Total: {mem.total // (1024**2)}MB, Livre: {mem.available // (1024**2)}MB"))
except ImportError:
    resultados.append(("Memória", "ℹ️", "psutil não instalado"))

# Teste 7: Decodificar dados embarcados
try:
    import gzip, base64, json
    # Lê o app_rh.py e extrai a string base64
    app_path = os.path.join(base_dir, "app_rh.py") if '__file__' in dir() else "app_rh.py"
    if os.path.exists(app_path):
        with open(app_path, "r", encoding="utf-8") as f:
            source = f.read()
        import re
        match = re.search(r'_DADOS_EMBUTIDOS_B64\s*=\s*\((.*?)\)', source, re.DOTALL)
        if match:
            parts = re.findall(r'"([^"]{0,300})"', match.group(1))
            full_b64 = ''.join(parts)
            raw = base64.b64decode(full_b64)
            texto = gzip.decompress(raw).decode("utf-8")
            dados = json.loads(texto)
            resultados.append(("Dados embarcados", "✅", f"{len(dados)} abas, {len(texto)} chars"))
        else:
            resultados.append(("Dados embarcados", "⚠️", "String base64 não encontrada no app_rh.py"))
    else:
        resultados.append(("Dados embarcados", "⚠️", "app_rh.py não encontrado neste diretório"))
except Exception as e:
    resultados.append(("Dados embarcados", "❌", str(e)))
    erros.append(f"Dados embarcados: {e}")

# Teste 8: try import app_rh
try:
    # Don't actually import - just compile check
    if os.path.exists(app_path):
        import py_compile
        py_compile.compile(app_path, doraise=True)
        resultados.append(("app_rh.py compilação", "✅", ""))
except SyntaxError as e:
    resultados.append(("app_rh.py compilação", "❌", str(e)))
    erros.append(f"SyntaxError: {e}")
except Exception as e:
    resultados.append(("app_rh.py compilação", "❌", str(e)))
    erros.append(f"Compile error: {e}")

# Exibir resultados
st.markdown("### Resultados dos Testes")
import pandas as pd
df = pd.DataFrame(resultados, columns=["Teste", "Status", "Detalhe"])
st.dataframe(df, use_container_width=True, hide_index=True)

if erros:
    st.markdown("### ❌ Erros Encontrados")
    for erro in erros:
        st.error(erro)
    st.markdown("### 📋 Traceback Completo")
    st.code(traceback.format_exc() if traceback.format_exc() != 'NoneType: None\n' else 'Nenhum traceback disponível')
else:
    st.success("✅ Todos os testes passaram!")

# Versões dos pacotes
st.markdown("### Versões dos Pacotes")
versoes = []
for mod_name in ["streamlit", "pandas", "openpyxl", "PIL", "requests", "matplotlib", "reportlab", "pydeck"]:
    try:
        mod = __import__(mod_name)
        versoes.append({"Pacote": mod_name, "Versão": getattr(mod, "__version__", "N/A")})
    except ImportError:
        versoes.append({"Pacote": mod_name, "Versão": "NÃO INSTALADO"})
st.dataframe(pd.DataFrame(versoes), use_container_width=True, hide_index=True)

st.markdown("---")
st.markdown("**Se o app_rh.py está dando erro, copie os resultados acima e envie para análise.**")
