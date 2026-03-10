# Automação de Preços (TRRs)

Projeto em Python para automação de planilhas de preços com integração em **Google Sheets** e coleta via **Selenium** nos portais de fornecedores.

## Visão geral

Este repositório possui três fluxos principais:

1. **Preparar aba do dia na planilha mensal** (`rodar_tudo.py`)
   - Localiza a planilha do mês (nome no padrão `Preço teste TRRs %m/%y`).
   - Duplica uma aba base anterior.
   - Atualiza cabeçalhos de data nos blocos `FOB - ...`.
   - Move valores de "hoje" para "ontem" com regras seguras.

2. **Coleta da Vibra e gravação direta** (`vibra.py`)
   - Faz login por base na Vibra.
   - Captura preços de S10/S500.
   - Grava diretamente em células específicas da planilha oficial.

3. **Coleta da Ipiranga (modo teste)** (`ipiranga.py`)
   - Abre portal Ipiranga com login manual.
   - Seleciona CNPJ/base de teste.
   - Captura preços (S10/S500/Gasolina).
   - Grava na linha da companhia `IPIRANGA` dentro do bloco alvo.

---

## Estrutura do projeto

- `rodar_tudo.py` → preparação da aba diária na planilha do mês.
- `vibra.py` → robô Selenium para Vibra + escrita em planilha por célula.
- `ipiranga.py` → robô Selenium para Ipiranga + escrita por bloco/coluna dinâmica.
- `dados-google.json` → credencial de conta de serviço do Google (sensível).
- `.env` → variáveis locais (sensível).

---

## Requisitos

- Python **3.10+** (compatível com 3.14, conforme seu ambiente atual).
- Google Chrome instalado.
- Permissão para abrir navegador com Selenium.
- Conta de serviço no Google Cloud com chave JSON válida.

### Dependências Python

Instale no ambiente virtual (recomendado):

```bash
python -m venv .venv
.venv\Scripts\activate
pip install gspread google-auth selenium webdriver-manager oauth2client
```

> Se quiser fixar versões, crie um `requirements.txt` e instale com `pip install -r requirements.txt`.

---

## Configuração do Google Cloud / Google Sheets

1. No Google Cloud, crie (ou use) uma **Conta de serviço**.
2. Gere uma chave **JSON** (aba _Chaves_ da conta de serviço).
3. Salve esse arquivo como `dados-google.json` na raiz do projeto.
4. Ative as APIs:
   - **Google Sheets API**
   - **Google Drive API**
5. Compartilhe a planilha/pasta com o e-mail da conta de serviço (`...iam.gserviceaccount.com`) com permissão de edição.

### Checklist de validação da credencial

No `dados-google.json`, valide:

- `type` = `service_account`
- `client_email` correto
- `private_key` começa com `-----BEGIN PRIVATE KEY-----`
- chave não revogada no Google Cloud

---

## Como executar

## 1) Preparar aba diária (rodar_tudo)

```bash
python rodar_tudo.py
```

### O que esse script faz

- Conecta no Google via `Credentials.from_service_account_file`.
- Busca planilha do mês por nome (padrão configurável em `NOME_ARQUIVO_MODELO`).
- Identifica se a aba de hoje já existe.
- Se não existir, duplica aba base (`Preço dd-mm` anterior).
- Nos blocos `FOB - ...`:
  - Atualiza cabeçalhos de `ontem` e `hoje`.
  - Copia valores da coluna de hoje para ontem.
  - Limpa coluna de hoje **somente se não for fórmula**.
  - Não altera coluna de diferença (`Dif.`).

## 2) Coletar Vibra

```bash
python vibra.py
```

### Fluxo atual do `vibra.py`

- Processa cada base em `BASES_VIBRA`.
- Faz login automático com usuário/senha da própria configuração do script.
- Navega para carrinho/preços.
- Captura itens contendo `S10` e `S500`.
- Grava em células fixas (ex.: `E13`, `I13`, etc.) da aba do dia.

> Observação: Se a aba do dia não existir, ele duplica a última aba da planilha oficial.

## 3) Coletar Ipiranga (teste)

```bash
python ipiranga.py
```

### Fluxo atual do `ipiranga.py`

- Conecta na planilha do mês.
- Abre o portal Ipiranga e aguarda **login manual** (ENTER no terminal).
- Seleciona o cliente/base definido em `BASE_TESTE`.
- Abre "Pedido de Combustível".
- Captura preço de retirada para:
  - Diesel S10
  - Diesel S500
  - Gasolina Comum
- Escreve na linha `IPIRANGA` do bloco configurado em `titulo_bloco_planilha`.

---

## Configurações importantes no código

### `rodar_tudo.py`

- `ARQUIVO_JSON_GOOGLE`
- `NOME_ARQUIVO_MODELO` (ex.: `Preço teste TRRs %m/%y`)
- `INTERVALO_LEITURA`
- `TITULOS_PROIBIDOS`

### `vibra.py`

- `ID_PLANILHA_OFICIAL`
- `NOME_ABA_HOJE`
- `BASES_VIBRA` (usuário/senha/células)

### `ipiranga.py`

- `BASE_TESTE` (CNPJ, local, título do bloco)
- `NOME_ARQUIVO_MODELO`
- `INTERVALO_LEITURA`

---

## Erros comuns e como resolver

### 1) `invalid_grant: Invalid JWT Signature`

Causa: credencial inválida/incompatível (`client_email` e `private_key` não correspondem, chave revogada, JSON incorreto, relógio da máquina fora de sincronia).

Correção:

- Gerar nova chave JSON da conta de serviço.
- Substituir `dados-google.json`.
- Confirmar APIs habilitadas e compartilhamento da planilha.
- Sincronizar data/hora do Windows.

### 2) `SpreadsheetNotFound`

Causa: conta de serviço não tem acesso à planilha ou nome não bate exatamente com o esperado.

Correção:

- Compartilhar planilha com o e-mail da conta de serviço.
- Conferir padrão do nome (`Preço teste TRRs %m/%y`).
- Verificar se está no projeto Google correto.

### 3) `WorksheetNotFound`

Causa: aba do dia ainda não existe.

Correção:

- Rodar `rodar_tudo.py` antes dos coletores.
- Confirmar padrão de nome da aba (`Preço dd-mm`).

### 4) Selenium não encontra elemento

Causa: mudança de layout, pop-up/cookies, timeout curto ou sessão não autenticada.

Correção:

- Revalidar seletores (`XPATH`, `CSS_SELECTOR`).
- Aumentar timeout.
- Tratar pop-ups adicionais.
- Executar com janela visível para depuração.

---

## Ordem recomendada de execução diária

1. `python rodar_tudo.py` (garante a aba do dia pronta)
2. `python vibra.py` (preenche preços Vibra)
3. `python ipiranga.py` (preenche preços Ipiranga no bloco alvo)

---

## Segurança (muito importante)

Este projeto contém dados sensíveis por natureza (credenciais de portal e chave privada Google).

Boas práticas obrigatórias:

- **Nunca versionar** `dados-google.json` em repositórios públicos.
- Não manter usuário/senha hardcoded no código.
- Mover segredos para variáveis de ambiente (`.env`) e usar carregamento seguro.
- Rotacionar chaves e senhas periodicamente.
- Restringir permissões da conta de serviço ao mínimo necessário.

### `.gitignore` sugerido

```gitignore
# Segredos
.env
dados-google.json

# Ambiente Python
.venv/
__pycache__/
*.pyc
```

---

## Melhorias recomendadas (próximos passos)

1. Padronizar autenticação Google (usar `google-auth` em todos scripts; hoje `vibra.py` usa `oauth2client`).
2. Centralizar configuração em `config.py` + `.env`.
3. Criar `requirements.txt` com versões fixas.
4. Implementar logs estruturados em arquivo (`logging`).
5. Adicionar modo de teste (sem escrita em planilha).
6. Implementar abertura por ID da planilha para evitar falhas por nome.

---

## Licença e uso interno

Projeto voltado para automação operacional interna de atualização de preços.
Se for compartilhar externamente, remover segredos e dados de acesso antes.
