# Guia definitivo: deixar os dados do sistema de RH salvos para sempre (Google Sheets)

> Por que fazer isso: hoje seu sistema roda na internet (Streamlit Cloud) e guarda os
> dados em arquivos temporários do servidor. Toda vez que o sistema reinicia, esses
> arquivos são apagados — foi por isso que "sumiu tudo". Ao conectar o sistema ao
> Google Sheets, todos os cadastros passam a ficar guardados na nuvem do Google,
> de forma permanente e com histórico de versões. Depois disso, nunca mais some.

Você vai fazer isto UMA vez. Leva mais ou menos 30 a 40 minutos. Vá com calma,
seguindo na ordem. Não precisa saber programar.

---

## PARTE 1 — Criar as planilhas no Google (5 min)

1. Entre em https://sheets.google.com com a sua conta Google.
2. Crie uma planilha em branco e dê o nome de **RH_Funcionarios**.
3. Olhe o endereço (link) da planilha na barra do navegador. Ele tem esta cara:
   `https://docs.google.com/spreadsheets/d/`**`1AbCdEfGhIjKlMnOpQrStUvWxYz1234567890`**`/edit`
   A parte em negrito (entre `/d/` e `/edit`) é o **ID da planilha**. Copie e guarde num bloco de notas, escrevendo ao lado: "funcionarios".
4. Repita: crie mais 3 planilhas em branco com os nomes **RH_Diarias**, **RH_Viagens** e **RH_Compras**, e guarde o ID de cada uma (anotando qual é qual).

> Pode deixar as 4 planilhas totalmente vazias. O sistema cria as colunas sozinho na primeira vez que salvar.

---

## PARTE 2 — Criar a "chave de acesso" (conta de serviço) no Google Cloud (15 min)

Essa é a parte que parece técnica, mas é só clicar seguindo os passos.

1. Entre em https://console.cloud.google.com com a MESMA conta Google.
2. No topo, clique no seletor de projeto e depois em **Novo projeto**. Dê o nome **Sistema RH** e clique em **Criar**. Espere criar e selecione esse projeto.
3. Ative as duas APIs necessárias:
   - Vá em https://console.cloud.google.com/apis/library/sheets.googleapis.com e clique em **Ativar**.
   - Vá em https://console.cloud.google.com/apis/library/drive.googleapis.com e clique em **Ativar**.
4. Crie a conta de serviço:
   - Vá em https://console.cloud.google.com/iam-admin/serviceaccounts
   - Clique em **Criar conta de serviço**.
   - Nome: `sistema-rh`. Clique em **Criar e continuar** e depois em **Concluir** (pode pular as permissões opcionais).
5. Copie o e-mail da conta de serviço. Ele aparece na lista e tem esta cara:
   `sistema-rh@sistema-rh-123456.iam.gserviceaccount.com`
   **Guarde esse e-mail** — vamos usar já já.
6. Gere a chave em JSON:
   - Clique na conta de serviço que você criou.
   - Aba **Chaves** (Keys) → **Adicionar chave** → **Criar nova chave** → tipo **JSON** → **Criar**.
   - Um arquivo `.json` será baixado no seu computador. **Esse arquivo é a senha do sistema — não compartilhe com ninguém.** Guarde num lugar seguro.

---

## PARTE 3 — Dar acesso das planilhas à conta de serviço (3 min)

Muito importante, senão o sistema não consegue escrever nas planilhas.

1. Abra cada uma das 4 planilhas (RH_Funcionarios, RH_Diarias, RH_Viagens, RH_Compras).
2. Em cada uma, clique no botão **Compartilhar** (canto superior direito).
3. Cole o **e-mail da conta de serviço** (aquele do passo 5 da Parte 2).
4. Defina a permissão como **Editor** e clique em **Enviar** (pode desmarcar "notificar").
5. Faça isso nas 4 planilhas.

---

## PARTE 4 — Colocar as chaves no Streamlit (10 min)

1. Abra o arquivo `.json` que você baixou (Parte 2, passo 6) com o Bloco de Notas.
   Você vai ver vários campos, tipo `"project_id"`, `"private_key"`, `"client_email"` etc.
2. Entre no painel do seu app em https://share.streamlit.io , abra o seu aplicativo,
   clique nos três pontinhos → **Settings** → aba **Secrets**.
3. Cole o conteúdo abaixo, substituindo cada `COLE_AQUI...` pelos valores do arquivo `.json`
   e os IDs das planilhas pelos que você anotou na Parte 1. Use o arquivo modelo
   `secrets_MODELO.toml` que veio junto com este guia — ele já está no formato certo.

```toml
[gspread]
type = "service_account"
project_id = "COLE_AQUI_project_id"
private_key_id = "COLE_AQUI_private_key_id"
private_key = "-----BEGIN PRIVATE KEY-----\nCOLE_AQUI_A_CHAVE_INTEIRA\n-----END PRIVATE KEY-----\n"
client_email = "COLE_AQUI_client_email"
client_id = "COLE_AQUI_client_id"
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "COLE_AQUI_client_x509_cert_url"

[gsheets]
id_funcionarios = "COLE_AQUI_ID_DA_PLANILHA_RH_Funcionarios"
id_diarias = "COLE_AQUI_ID_DA_PLANILHA_RH_Diarias"
id_viagens = "COLE_AQUI_ID_DA_PLANILHA_RH_Viagens"
id_compras = "COLE_AQUI_ID_DA_PLANILHA_RH_Compras"
somente_nuvem = true
```

> Atenção ao campo `private_key`: no arquivo .json ele já vem com os `\n` no meio.
> Copie exatamente como está, inteiro, entre as aspas. Não apague os `\n`.

4. Clique em **Save**. O app vai reiniciar sozinho.

---

## PARTE 5 — Conferir se deu certo

1. Abra o sistema pelo link normal.
2. Na barra lateral esquerda deve aparecer uma mensagem verde escrito
   **"100% Google Sheets"** — isso confirma que os dados agora ficam na nuvem.
3. Cadastre um funcionário de teste, saia e entre de novo. Se ele continuar lá
   depois de recarregar, está funcionando. Confira também que ele apareceu na
   planilha RH_Funcionarios no Google.

Pronto! A partir de agora todos os dados ficam salvos no Google Sheets, de forma
permanente, e ainda com histórico de versões (dá para voltar no tempo se precisar).

---

## Dica de segurança extra (opcional, mas recomendado)

- Uma vez por semana, dentro do sistema, use o botão de **baixar backup em ZIP** e
  guarde esse arquivo no seu computador ou num drive. É a sua rede de segurança.
- Não compartilhe o arquivo `.json` nem o conteúdo dos "Secrets" com ninguém.

Se travar em algum passo, me diga em qual número você parou e o que apareceu na tela
que eu te ajudo a destravar.
