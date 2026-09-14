# 🏢 Sistema RH - Guia de Instalação

## 📋 Sobre
Sistema completo de Recursos Humanos com 12 módulos:
1. **Cadastro** — Ficha do funcionário, foto, documentos, upload de PDF
2. **Painel** — Dashboard com resumo geral
3. **Prazos e Férias** — Alertas de experiência e férias vencendo
4. **Histórico** — Registro de ocorrências
5. **Relatórios** — Exportação individual e por situação
6. **Diárias** — Controle de diárias por colaborador
7. **Lojas e Cargos** — Cadastro de lojas, cargos e auxiliares
8. **Compras** — Solicitações e entregas
9. **Guia Viagem** — Calculadora de rotas e registro de viagens
10. **Backup/Restauração** — Cópia de segurança e restauração via ZIP
11. **Ferramentas** — Utilitários (IA, TTS, etc.)
12. **Ajuda** — Manual do sistema

## 🚀 Instalação Local

### Pré-requisitos
- Python 3.10 ou superior
- pip (gerenciador de pacotes)

### Passo a passo

1. **Extrair o ZIP** em uma pasta de sua preferência

2. **Criar ambiente virtual** (recomendado):
   ```bash
   python -m venv venv
   venv\Scripts\activate     # Windows
   source venv/bin/activate  # Linux/Mac
   ```

3. **Instalar dependências**:
   ```bash
   pip install -r requirements.txt
   ```

4. **Executar o sistema**:
   ```bash
   streamlit run app_rh.py
   ```

5. **Acesse**: http://localhost:8501

## ☁️ Publicar no Streamlit Cloud

1. Faça upload de todo o conteúdo para um repositório GitHub
2. Acesse https://streamlit.io/cloud
3. Conecte o repositório e configure:
   - **Main file**: app_rh.py
   - **Requirements**: requirements.txt
   - **Packages**: packages.txt

### Configurar Secrets (Google Sheets)
Em Settings → Secrets, cole o conteúdo do arquivo `.streamlit/secrets.toml`
preenchendo com suas credenciais do Google Cloud.

## 📁 Estrutura de Pastas

O sistema cria automaticamente estas pastas na primeira execução:

```
pasta_do_sistema/
├── app_rh.py                    ← Script principal
├── requirements.txt             ← Dependências Python
├── .streamlit/
│   ├── config.toml             ← Configuração do Streamlit
│   └── secrets.toml            ← Credenciais (NÃO compartilhar!)
├── modelos_pdf/                ← Modelos de PDF para referência
├── dados_funcionarios.xlsx     ← Base de dados principal (auto-criado)
├── controle_diarias.xlsx       ← Diárias (auto-criado)
├── registro_viagens.xlsx       ← Viagens (auto-criado)
├── dados_compras.json          ← Compras (auto-criado)
├── Documentos_Lojas/           ← Docs das lojas (auto-criado)
├── Documentos_Funcionarios/   ← Docs dos funcionários (auto-criado)
├── Fotos_Funcionarios/         ← Fotos (auto-criado)
├── Comprovantes_Diarias/       ← Comprovantes (auto-criado)
└── Backups_Automaticos/        ← Backups (auto-criado)
```

## 🔧 Configuração de Pasta Persistente

Para que os dados sobrevivam a reinicializações no Streamlit Cloud,
configure a variável de ambiente `RH_DATA_DIR` apontando para um
disco persistente, ou adicione em `.streamlit/secrets.toml`:

```toml
[dados]
pasta = "/caminho/para/pasta/persistente"
```

## 🔄 Sistema de Espelho (Backup Automático)

O sistema pode copiar automaticamente todos os dados para uma pasta
de backup (ex: unidade D:). Configure:

```toml
[espelho]
pasta = "D:\Backup_RH"
```

Ou a variável de ambiente `RH_ESPELHO_DIR`.

## ⚡ Otimizações de Performance

O sistema já vem otimizado com:
- Cache de dados em session_state (zero recarga redundante)
- Renderização lazy das abas (só a aba ativa executa)
- Cache de lista de lojas/cargos
- Espelhamento adiado para a primeira interação

## 📞 Suporte

Em caso de dúvidas, consulte a aba "Ajuda" dentro do sistema.
