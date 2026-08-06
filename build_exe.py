#!/usr/bin/env python3
"""
Script de Build - Gera o executavel SistemaRH.exe
Como usar no Windows (depois de instalar o Python):
    python build_exe.py
"""

import subprocess
import sys
import os
import shutil

def install(pkg):
    print(f"[INFO] Instalando {pkg}...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", pkg, "-q"])

def main():
    print("=" * 55)
    print("  BUILD - Sistema RH (Streamlit -> .exe)")
    print("=" * 55)

    # Verifica dependencias
    for mod, pkg in [("streamlit", "streamlit"), ("PIL", "Pillow"), ("PyInstaller", "pyinstaller")]:
        try:
            __import__(mod)
            print(f"[OK] {pkg} ja instalado")
        except ImportError:
            install(pkg)

    # Verifica app_rh.py
    if not os.path.exists("app_rh.py"):
        print("[ERRO] Arquivo app_rh.py nao encontrado!")
        print("       Copie o app_rh.py para esta pasta e tente novamente.")
        input("Pressione ENTER para sair...")
        sys.exit(1)

    # Verifica launcher.py
    if not os.path.exists("launcher.py"):
        print("[ERRO] Arquivo launcher.py nao encontrado!")
        print("       Ele deve estar na mesma pasta que este script.")
        input("Pressione ENTER para sair...")
        sys.exit(1)

    # Converte icone se existir
    icon_path = None
    for fn in ["app_icon.ico", "app_icon.png", "icon.png", "logo.png"]:
        if os.path.exists(fn):
            if fn.endswith(".ico"):
                icon_path = fn
                print(f"[OK] Icone encontrado: {fn}")
            else:
                from PIL import Image
                img = Image.open(fn).convert("RGBA")
                icon_path = "app_icon_temp.ico"
                img.save(icon_path, format="ICO", sizes=[(256,256),(128,128),(64,64),(32,32),(16,16)])
                print(f"[OK] Icone convertido: {fn}")
            break

    # Monta comando PyInstaller
    sep = ";" if sys.platform.startswith("win") else ":"
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onefile", "--noconfirm", "--clean",
        "--name", "SistemaRH",
        "--add-data", f"app_rh.py{sep}.",
        "--hidden-import", "streamlit",
        "--hidden-import", "streamlit.runtime.scriptrunner.magic_funcs",
    ]
    if icon_path and os.path.exists(icon_path):
        cmd.extend(["--icon", icon_path])
    cmd.append("launcher.py")

    print("")
    print("[INFO] Executando PyInstaller...")
    print("Isso pode demorar varios minutos na primeira vez.")
    print("")

    result = subprocess.run(cmd)

    # Limpa arquivos temporarios
    for f in ["launcher.spec", "app_icon_temp.ico"]:
        if os.path.exists(f):
            os.remove(f)
    if os.path.exists("build"):
        shutil.rmtree("build")

    if result.returncode == 0:
        print("")
        print("=" * 55)
        print("  SUCESSO!")
        print(f"  Executavel criado em:  dist/SistemaRH.exe")
        print("=" * 55)
        print("")
        print("Como usar:")
        print("  1. Va na pasta 'dist/'")
        print("  2. Clique duas vezes em 'SistemaRH.exe'")
        print("  3. O navegador abrira automaticamente!")
        print("")
        print("Dica: Clique com o botao direito no .exe,")
        print("      escolha 'Enviar para -> Area de trabalho'")
        print("      para criar um atalho!")
        print("")
    else:
        print("")
        print("[ERRO] PyInstaller falhou.")

    input("Pressione ENTER para sair...")

if __name__ == "__main__":
    main()
