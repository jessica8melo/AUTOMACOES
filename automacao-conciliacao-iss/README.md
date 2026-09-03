# automacao-conciliacao-iss

Automação do processo de Conciliação de ISS.

## Visão Geral

A automação realiza a identificação e validação de retenções de ISS por meio
do cruzamento de informações provenientes de diferentes bases operacionais.

O processo utiliza:

- Controle Financeiro;
- Relatório de Pagamentos;
- Relatório de Recebimento Integrado;
- Tabela de referência de localidades.

Ao final da execução, o sistema identifica retenções já pagas, valida os
respectivos recebimentos e sinaliza inconsistências diretamente nas
planilhas analisadas.

## Fluxo do Processo

O fluxo principal é composto por três etapas executadas em sequência pelo
script `main.py`.

### 1. Controle Financeiro

Responsável por localizar retenções de ISS pendentes de identificação.

O script:

- Percorre a planilha de Controle Financeiro;
- Procura linhas cuja coluna **OB** esteja vazia;
- Considera apenas células destacadas em amarelo;
- Extrai o código da localidade presente na descrição;
- Consulta a tabela de localidades para obter o CAT e a ORG correspondentes.

Exemplo:

```text
Descrição: ISS Retido MPI

Localidade: MPI
CAT: correspondente ao cadastro
ORG: correspondente ao cadastro
```

### 2. Relatório de Pagamentos

Para cada CAT identificado na etapa anterior:

- Filtra a planilha de pagamentos pela coluna **Fornecedor**;
- Localiza os pagamentos relacionados à localidade;
- Agrupa os resultados por data;
- Consolida os valores por ORG.

Essa etapa gera a base utilizada para validação dos recebimentos.

### 3. Relatório de Recebimento Integrado

Para cada ORG encontrada na etapa anterior:

- Filtra o Relatório de Recebimento Integrado pela coluna **Organização**;
- Compara os registros de pagamento e recebimento;
- Valida correspondências utilizando os campos:

```text
N.F  == NFF
ISS  == Valor
```

Quando ambos os critérios são satisfeitos:

- A célula `N.F` recebe destaque amarelo;
- O registro é marcado como `OK`.

Quando existe qualquer divergência:

- O registro é marcado como `INCONSISTENTE`;
- O motivo é exibido no resultado.

Durante o processamento também são verificados:

- Conteúdo na coluna Observações;
- Comentários inseridos diretamente no Excel.

Essas informações são exibidas como avisos durante a execução.

## Estrutura do Projeto

| Arquivo | Descrição |
|----------|----------|
| `main.py` | Orquestra todo o fluxo de conciliação |
| `controle_financeiro.py` | Identifica retenções pendentes no controle financeiro |
| `relatorio_pagamentos.py` | Localiza e consolida pagamentos por CAT |
| `relatorio_recebimento.py` | Valida os recebimentos por ORG |
| `tabela_localidade.py` | Base de referência de localidades, CATs e ORGs |
| `requirements.txt` | Dependências do projeto |

## Pré-requisitos

- Python 3.8 ou superior

## Instalação

### 1. Criar o ambiente virtual

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

### Fluxo Completo

Para executar todas as etapas da conciliação:

```bash
python main.py \
    --controle "controle-financeiro-2026-2.xlsx" \
    --pagamento "Relatório de pagamento.xlsx" \
    --recebimento "Relatório Recebimento Integrado - 07-2026- ERP- Demais Filiais- Contas a Pagar.xlsx"
```

Quando os parâmetros não são informados, cada etapa utiliza o arquivo padrão
definido em seu respectivo módulo (`DEFAULT_PATH`).

Também é possível especificar as abas das planilhas:

```bash
--sheet-controle
--sheet-pagamento
--sheet-recebimento
```

### Executar sem alterar o arquivo de recebimento

Por padrão, a automação grava a marcação amarela na coluna `N.F` da planilha
de recebimento.

Para executar apenas a análise, sem salvar alterações:

```bash
python main.py --nao-salvar-recebimento
```

### Processar apenas um CAT específico

Nesse modo, a etapa de Controle Financeiro é ignorada e o processo inicia
diretamente no cruzamento de pagamentos.

```bash
python main.py --cat "Bauru"
```

Fluxo executado:

```text
CAT informado
↓
Relatório de Pagamentos
↓
Relatório de Recebimento Integrado
```

### Executar etapas isoladamente

#### Etapa 1 - Controle Financeiro

```bash
python controle_financeiro.py [caminho_do_arquivo.xlsx] [--sheet NOME_DA_ABA]
```

#### Etapa 2 - Relatório de Pagamentos

```bash
python relatorio_pagamentos.py "Bauru" [caminho_do_arquivo.xlsx] [--sheet NOME_DA_ABA]
```

#### Etapa 3 - Relatório de Recebimento Integrado

```bash
python relatorio_recebimento.py \
    --org "MBA" \
    --pagamentos-json pagamentos.json \
    [--recebimento CAMINHO.xlsx] \
    [--sheet-recebimento NOME_DA_ABA] \
    [--nao-salvar]
```

## Exemplo de Fluxo Completo

```bash
# Executar a conciliação completa

python main.py \
    --controle controle-financeiro.xlsx \
    --pagamento relatorio-pagamentos.xlsx \
    --recebimento recebimento-integrado.xlsx
```

## Exemplo de JSON para a Etapa 3

Ao executar apenas o módulo de Recebimento Integrado, o parâmetro
`--pagamentos-json` deve apontar para um arquivo com a mesma estrutura
gerada pela etapa de pagamentos.

Exemplo:

```json
[
    {
        "NFF": "996.02",
        "Valor": 21603.16
    }
]
```

## Observações

- O arquivo `tabela_localidade.py` deve permanecer acessível aos demais
  módulos, pois é utilizado por `controle_financeiro.py`,
  `relatorio_pagamentos.py` e `main.py`.

- A identificação da cor amarela admite pequenas variações de tonalidade,
  considerando diferenças normalmente encontradas em planilhas reais.

- Como o `openpyxl` reescreve o arquivo ao salvar, valores calculados por
  fórmulas podem não ser atualizados imediatamente. O recálculo ocorre ao
  abrir novamente a planilha no Excel ou LibreOffice.

## Encerramento

Ao finalizar o trabalho, desative o ambiente virtual:

```bash
deactivate
```

> A pasta `.venv` deve permanecer no `.gitignore` e não deve ser versionada.
