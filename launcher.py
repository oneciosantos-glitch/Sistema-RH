#!/usr/bin/env python3
"""
Launcher para o Sistema RH - Streamlit
Inicia o servidor e abre o navegador automaticamente.
"""

import sys
import os
import subprocess
import time
import webbrowser
import socket
import threading


def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)


def free_port(start=8501, max_attempts=100):
    for port in range(start, start + max_attempts):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            if s.connect_ex(('127.0.0.1', port)) != 0:
                return port
    return start


def wait_server(port, timeout=120):
    t0 = time.time()
    while time.time() - t0 < timeout:
        try:
            with socket.create_connection(('127.0.0.1', port), timeout=1):
                return True
        except:
            time.sleep(0.5)
    return False


def stream_output(pipe):
    """Le a saida do processo e imprime em tempo real."""
    try:
        for line in iter(pipe.readline, ''):
            print(line, end='')
        pipe.close()
    except Exception:
        pass


def main():
    app_path = resource_path("app_rh.py")
    if not os.path.exists(app_path):
        base = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, 'frozen', False) else __file__))
        app_path = os.path.join(base, "app_rh.py")

    if not os.path.exists(app_path):
        print("[ERRO] Arquivo app_rh.py nao encontrado!")
        print("Caminho tentado:", app_path)
        input("Pressione ENTER para sair...")
        sys.exit(1)

    port = free_port(8501)
    url = f"http://localhost:{port}"

    print("=" * 55)
    print("  SISTEMA RH - Iniciando servidor Streamlit...")
    print("=" * 55)
    print("")
    print(f"Porta: {port}")
    print(f"App:   {app_path}")
    print("")
    print("[INFO] Aguarde o servidor iniciar...")
    print("[INFO] Logs do Streamlit aparecerao abaixo.")
    print("")
    print("-" * 55)

    cmd = [
        sys.executable, "-m", "streamlit", "run", app_path,
        "--server.port", str(port),
        "--server.address", "127.0.0.1",
        "--server.headless", "true",
        "--browser.gatherUsageStats", "false",
        "--server.enableCORS", "false",
        "--server.enableXsrfProtection", "false"
    ]

    # Inicia o processo capturando stdout/stderr
    proc = subprocess.Popen(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        bufsize=1,
        universal_newlines=True
    )

    # Thread para mostrar logs em tempo real
    t = threading.Thread(target=stream_output, args=(proc.stdout,))
    t.daemon = True
    t.start()

    # Aguarda o servidor subir
    print("[INFO] Verificando se o servidor esta pronto...")
    if wait_server(port, timeout=120):
        print("")
        print("[OK] Servidor pronto!")
        print(f"[OK] Abrindo navegador: {url}")
        print("")
        time.sleep(1)
        webbrowser.open(url)
    else:
        print("")
        print("[AVISO] Servidor demorou para responder.")
        print("[AVISO] O navegador pode abrir manualmente em:")
        print(f"        {url}")
        print("")

    print("-" * 55)
    print("Sistema rodando. Pressione CTRL+C para encerrar.")
    print("-" * 55)
    print("")

    # Aguarda o processo terminar
    try:
        proc.wait()
    except KeyboardInterrupt:
        print("\n[INFO] Encerrando servidor...")
        proc.terminate()
        try:
            proc.wait(timeout=5)
        except:
            proc.kill()
        print("[OK] Servidor encerrado.")


if __name__ == "__main__":
    main()
