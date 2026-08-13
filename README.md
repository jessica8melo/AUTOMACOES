# automacao-conciliacao
Automação dos processos de Conciliação de Pagamentos, Contas a Receber e Conciliação de ISS

## Como rodar (venv)

Pré-requisito: Python 3.9+ instalado.

### 1. Criar e ativar o ambiente virtual

**Windows (PowerShell):**
```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
```

**Windows (cmd):**
```cmd
python -m venv .venv
.venv\Scripts\activate.bat
```

**Mac/Linux:**
```bash
python3 -m venv .venv
source .venv/bin/activate
```

Depois de ativado, o prompt do terminal passa a mostrar `(.venv)` no início da linha.

### 2. Instalar as dependências
```bash
pip install -r requirements.txt
```

### 3. Rodar os scripts, na ordem

O processo completo tem 4 etapas, sempre nesta ordem:

**Passo 1 - Extrair os extratos bancários (PDF -> planilha)**
```bash
python extrato_parser.py --dir caminho/da/pasta/com/pdfs --xlsx saida.xlsx
```
Lê todos os PDFs de extrato do Banco do Brasil de uma pasta e gera `saida.xlsx`,
com uma aba por conta bancária, cada uma já separada em blocos de Pagamentos e
Recebimentos.

**Passo 2 - Separar a planilha de conciliação em Recebimentos/Pagamentos**
```bash
python separar_tipo.py conciliacao.xlsx
```
Em cada aba de `conciliacao.xlsx` (uma aba por dia), separa os lançamentos em
blocos **Recebimentos** e **Pagamentos** (e **Outros**, se houver linhas não
classificadas), com base na coluna "Tipo".

**Passo 3 - Remover OBs com valor zero**
```bash
python remover_zero.py conciliacao.xlsx
```
Remove, de cada aba, as linhas da tabela auxiliar OB/VALOR em que VALOR = 0.
Esse passo precisa ser rodado antes do Passo 4 - `processar_pagamentos.py`
não faz mais essa limpeza sozinho.

**Passo 4 - Conciliar pagamentos com OBs (+ Etapa 3 opcional com extratos)**
```bash
# só concilia pagamento <-> OB dentro de cada aba
python processar_pagamentos.py conciliacao.xlsx

# concilia e também cruza os pagamentos sem OB com os extratos (saida.xlsx)
python processar_pagamentos.py conciliacao.xlsx --extratos saida.xlsx
```
Concilia pagamento(s) com OB(s) em cada aba (Passo 2 da lógica de conciliação).
Se `--extratos` for informado, roda também a **Etapa 3**: cruza os pagamentos
(conciliados ou não) com a planilha de extratos por Data + Valor, preenchendo
Status/Justificativas.

A Etapa 3 só é acionada em uma aba se a **data da própria aba** (ex.: aba
`"07.08.25"`) tiver pelo menos um lançamento naquela mesma data dentro da
planilha de extratos. Se a data da aba não aparecer em nenhum extrato (ou o
nome da aba não for uma data), a Etapa 3 é pulada nessa aba - com um aviso no
log - mas a conciliação normal (pagamento <-> OB) continua rodando
normalmente nela.

### 4. Sair do ambiente virtual (quando terminar)
```bash
deactivate
```

> `.venv/` já está no `.gitignore` — não é versionado, cada máquina cria o seu.
