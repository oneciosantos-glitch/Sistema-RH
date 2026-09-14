"""Ponto de entrada alternativo para o Streamlit Cloud.

Se o 'Main file path' no Streamlit Cloud estiver configurado como
'app_rh.py', este arquivo redireciona para o modulo principal.
"""
import sys
import os

# Garantir que o diretorio raiz do app esta no path
_raiz = os.path.dirname(os.path.abspath(__file__))
if _raiz not in sys.path:
    sys.path.insert(0, _raiz)

# O streamlit_app.py roda set_page_config + estilo.aplicar() no import.
# main() precisa ser chamada explicitamente porque esta protegida
# por if __name__ == "__main__".
from streamlit_app import main

try:
    main()
except Exception as _e:
    import traceback as _tb
    import streamlit as st
    _msg = ''.join(_tb.format_exception(type(_e), _e, _e.__traceback__))
    st.error("Erro ao iniciar o aplicativo:\n" + _msg)
    st.code(_msg, language="python")
