============================================================
   SISTEMA RH - PACOTE PARA WINDOWS
   Abre com icone vermelho (coroa) - SEM JANELA PRETA!
============================================================

ARQUIVOS NESTA PASTA:
  app_rh.py                          -> Seu aplicativo Streamlit
  app_icon.ico                       -> Icone vermelho da coroa
  Iniciar_SistemaRH_Silencioso.vbs   -> INICIA SILENCIOSO (sem CMD!)
  Criar_Atalho_Desktop.vbs           -> Cria o atalho vermelho na Area de Trabalho
  launcher.py                        -> Script Python que inicia o Streamlit
  Iniciar_SistemaRH.bat              -> Versao com tela preta (para debug)
  Iniciar_SistemaRH_Launcher.bat     -> Versao alternativa com tela preta
  build_exe.py                       -> Cria executavel .exe (avancado)
  README.txt                         -> Este arquivo


COMO USAR (MODO SUPER SIMPLES - SEM JANELA PRETA):
----------------------------------------------------
1. Extraia este ZIP para uma pasta (ex: C:\SistemaRH)

2. VA NA PASTA EXTRAIDA e de DOIS CLIQUES em:
       Criar_Atalho_Desktop.vbs
   -> Um aviso vai dizer que o atalho foi criado!

3. Pronto! Agora va na sua AREA DE TRABALHO (Desktop)

4. Voce vera um icone VERMELHO com uma COROA chamado:
       Sistema RH

5. A PARTIR DE AGORA, e so clicar DUAS VEZES nesse icone
   vermelho e o navegador abrira sozinho!

   NADA de janela preta ficando aberta!
   NADA de CMD na tela!
   So o icone bonito e o sistema abrindo no Chrome/Edge!


SE O ATALHO NAO FUNCIONAR:
---------------------------
1. Clique com o botao DIREITO no icone vermelho "Sistema RH"
2. Escolha "Propriedades"
3. Na aba "Atalho", clique em "Localizar destino..."
4. Verifique se a pasta ainda existe


SE PRECISAR VER ERROS (MODO DEBUG):
------------------------------------
Se o icone vermelho clicar e nao abrir nada:
1. Volte na pasta do sistema
2. Clique DUAS VEZES em:
       Iniciar_SistemaRH.bat
3. A janela PRETA vai abrir e mostrar o erro
4. Tire uma foto da tela e me envie!


COMO MUDAR O ICONE DO ATALHO:
------------------------------
Se quiser trocar o icone depois:
1. Clique com o DIREITO no icone "Sistema RH"
2. Propriedades -> Alterar icone...
3. Clique em "Procurar..." e selecione o arquivo app_icon.ico
4. OK -> OK


REQUISITOS:
-----------
- Python 3.9+ instalado (https://python.org)
- Durante a instalacao, marque "Add Python to PATH"
- Navegador: Chrome, Edge ou Firefox


TROUBLESHOOTING:
----------------
1. "Python nao encontrado!"
   -> Reinstale o Python marcando "Add to PATH"

2. A porta 8501 esta em uso:
   -> Abra o CMD e digite:  taskkill /f /im streamlit.exe
   -> Depois tente iniciar novamente

3. Navegador nao abre:
   -> Abra manualmente: http://localhost:8501

4. Quer criar o atalho em outro lugar (nao no Desktop)?
   -> Clique com DIREITO no arquivo .vbs
   -> "Criar atalho"
   -> Mova o atalho para onde quiser
   -> Clique com DIREITO no atalho -> Propriedades
   -> Alterar icone... -> selecione app_icon.ico


DICAS:
------
- O sistema roda em segundo plano. Para fechar, use o Gerenciador
  de Tarefas (Ctrl+Shift+Esc) e finalize "streamlit.exe"
- O icone vermelho na Area de Trabalho e o seu "app" agora!
- Funciona igualzinho um programa instalado!

============================================================
