"""Ponto de entrada alternativo (main.py).

Se o 'Main file path' no Streamlit Cloud estiver configurado como
'main.py', este arquivo redireciona para o modulo principal.
"""
import sys
import os

_raiz = os.path.dirname(os.path.abspath(__file__))
if _raiz not in sys.path:
    sys.path.insert(0, _raiz)

from streamlit_app import main

try:
    main()
except Exception as _e:
    import traceback as _tb
    import streamlit as st
    _msg = ''.join(_tb.format_exception(type(_e), _e, _e.__traceback__))
    st.error("Erro ao iniciar o aplicativo:\n" + _msg)
    st.code(_msg, language="python")
