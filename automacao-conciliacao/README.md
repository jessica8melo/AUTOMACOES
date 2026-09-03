# automacao-conciliacao

Automação dos processos de Conciliação de Pagamentos e Contas a Receber.

## Visão Geral

A automação realiza o processo de conciliação financeira entre planilhas de
conciliação, Ordens Bancárias (OBs) e extratos bancários do Banco do Brasil.

O fluxo foi dividido em etapas independentes que permitem:

- Extrair movimentações de extratos bancários em PDF;
- Organizar lançamentos em Recebimentos e Pagamentos;
- Remover registros auxiliares inconsistentes;
- Conciliar pagamentos com OBs;
- Validar pagamentos utilizando extratos bancários.

## Fluxo do Processo

O processo completo é composto por quatro etapas executadas em sequência.

### 1. Extração de Extratos

Converte extratos bancários em PDF para uma planilha Excel estruturada,
gerando uma aba para cada conta bancária encontrada.

### 2. Separação dos Lançamentos

Classifica os registros da planilha de conciliação em:

- Recebimentos
- Pagamentos
- Outros

A classificação é realizada com base na coluna `Tipo`.

### 3. Limpeza de Dados

Remove registros auxiliares cuja OB esteja associada a valor igual a zero.

### 4. Conciliação

Relaciona pagamentos às respectivas OBs.

Opcionalmente também realiza cruzamento com os extratos bancários,
preenchendo status e justificativas de conciliação.

## Estrutura do Projeto

| Arquivo | Descrição |
|----------|----------|
| `extrato_parser.py` | Extrai movimentações dos extratos PDF |
| `separar_tipo.py` | Organiza os lançamentos em Recebimentos e Pagamentos |
| `remover_zero.py` | Remove registros auxiliares com valor zero |
| `processar_pagamentos.py` | Executa a conciliação financeira |
| `requirements.txt` | Dependências do projeto |

## Pré-requisitos

- Python 3.9 ou superior

## Instalação

### 1. Criar e ativar o ambiente virtual

#### Windows (PowerShell)

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
```

#### Windows (CMD)

```cmd
python -m venv .venv
.venv\Scripts\activate.bat
```

#### Linux / macOS

```bash
python3 -m venv .venv
source .venv/bin/activate
```

Após a ativação, o terminal exibirá `(.venv)` no início da linha.

### 2. Instalar as dependências

```bash
pip install -r requirements.txt
```

## Como Rodar

O processamento completo deve seguir obrigatoriamente a ordem abaixo.

### Passo 1 - Extrair os extratos bancários

```bash
python extrato_parser.py --dir caminho/da/pasta/com/pdfs --xlsx saida.xlsx
```

O script lê todos os PDFs de extrato do Banco do Brasil presentes na pasta
informada e gera uma planilha Excel contendo uma aba para cada conta
bancária.

Cada aba já é organizada em blocos de:

- Pagamentos
- Recebimentos

### Passo 2 - Separar a planilha de conciliação

```bash
python separar_tipo.py conciliacao.xlsx
```

Para cada aba da planilha de conciliação, os lançamentos são separados em:

- Recebimentos
- Pagamentos
- Outros (quando houver registros não classificados)

A classificação é baseada na coluna `Tipo`.

### Passo 3 - Remover OBs com valor zero

```bash
python remover_zero.py conciliacao.xlsx
```

Remove da tabela auxiliar OB/VALOR todas as linhas cujo valor seja igual a
zero.

> Este passo deve ser executado antes do Passo 4. O script
> `processar_pagamentos.py` não realiza mais essa limpeza automaticamente.

### Passo 4 - Conciliar pagamentos

#### Apenas conciliação entre pagamentos e OBs

```bash
python processar_pagamentos.py conciliacao.xlsx
```

#### Conciliação utilizando também os extratos bancários

```bash
python processar_pagamentos.py conciliacao.xlsx --extratos saida.xlsx
```

Neste modo, além da conciliação de pagamentos com OBs, o sistema executa uma
etapa adicional de validação utilizando os extratos bancários.

O cruzamento é realizado utilizando:

- Data
- Valor

Com isso o sistema pode complementar:

- Status
- Justificativas
- Informações de conciliação

## Exemplo de Fluxo Completo

```bash
# 1. Gerar extratos
python extrato_parser.py --dir extratos_pdf --xlsx extratos.xlsx

# 2. Organizar a planilha de conciliação
python separar_tipo.py conciliacao.xlsx

# 3. Limpar registros auxiliares
python remover_zero.py conciliacao.xlsx

# 4. Executar conciliação completa
python processar_pagamentos.py conciliacao.xlsx --extratos extratos.xlsx
```

## Observações

A etapa de validação com extratos só é executada para abas cuja data esteja
presente na planilha de extratos.

Exemplo:

- Aba da conciliação: `07.08.25`
- Existe movimentação em `07.08.25` nos extratos → etapa executada
- Não existe movimentação em `07.08.25` nos extratos → etapa ignorada

Quando a etapa é ignorada, a conciliação principal entre pagamentos e OBs
continua sendo executada normalmente.

## Encerramento

Ao finalizar o trabalho, desative o ambiente virtual:

```bash
deactivate
```

> A pasta `.venv` já está incluída no `.gitignore` e não deve ser versionada.
