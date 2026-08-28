# automacao-conciliacao-iss

Automação do processo de Conciliação de ISS.

O projeto cruza três planilhas (controle financeiro, relatório de
pagamentos e relatório de recebimento integrado) usando uma tabela de
referência de localidades, para identificar quais retenções de ISS já
foram pagas e sinalizar isso diretamente nas planilhas.

## Fluxo do processo

O script `main.py` encadeia três etapas:

1. **Controle financeiro** (`controle_financeiro.py`)
   Varre a planilha de controle financeiro e localiza todas as linhas em
   que a célula da coluna **OB** está pintada de **amarelo** e vazia (sem
   número de OB preenchido). Para cada ocorrência, extrai o código de
   localidade contido no final da **Descrição** (ex: "ISS Retido MPI" →
   código "MPI") e cruza com a `TABELA_LOCALIDADE` para descobrir o
   **CAT** e a **ORG** correspondentes.

2. **Relatório de pagamentos** (`relatorio_pagamentos.py`)
   Para cada CAT distinto encontrado na etapa 1, filtra a planilha de
   pagamentos pela coluna **Fornecedor** (que costuma conter o nome da
   localidade embutido em textos maiores) e soma os valores por **Data
   Pagamento**. Os resultados são agrupados pela ORG do respectivo CAT.

3. **Relatório de Recebimento Integrado** (`relatorio_recebimento.py`)
   Para cada ORG com pagamentos encontrados na etapa 2, filtra a planilha
   de Recebimento Integrado pela coluna **Organização** e confere, linha a
   linha:
   - `N.F` (recebimento) == `NFF` (pagamento)
   - `ISS` (recebimento) == `Valor` (pagamento)

   Quando ambos batem, a célula **N.F** é pintada de amarelo e a linha é
   marcada como `OK`; caso contrário, é marcada como `INCONSISTENTE` com o
   motivo. A cada linha processada, também é verificado se há texto na
   coluna **Observações** ou comentários nativos do Excel, que são
   exibidos como avisos.

## Estrutura do projeto

| Arquivo                     | Descrição                                                        |
|------------------------------|-------------------------------------------------------------------|
| `main.py`                    | Orquestra as três etapas em sequência                            |
| `controle_financeiro.py`     | Etapa 1 — identifica OBs amarelas/vazias                         |
| `relatorio_pagamentos.py`    | Etapa 2 — cruza pagamentos por CAT                                |
| `relatorio_recebimento.py`   | Etapa 3 — confere e sinaliza o recebimento integrado por ORG      |
| `tabela_localidade.py`       | Tabela de referência (escritório, CAT, sigla, ORG, vencimento...) |
| `requirements.txt`           | Dependências do projeto                                          |

## Requisitos

- Python 3.8+
- Instalar as dependências:

```bash
pip install -r requirements.txt
```

## Uso

### Fluxo completo

```bash
python main.py \
    --controle "controle-financeiro-2026-2.xlsx" \
    --pagamento "Relatório de pagamento.xlsx" \
    --recebimento "Relatório Recebimento Integrado - 07-2026- ERP- Demais Filiais- Contas a Pagar.xlsx"
```

Se os caminhos não forem informados, cada etapa usa o arquivo padrão
definido em seu próprio módulo (`DEFAULT_PATH`). Também é possível indicar
o nome da aba de cada planilha com `--sheet-controle`, `--sheet-pagamento`
e `--sheet-recebimento`.

Por padrão, o `main.py` **grava** a cor amarela na coluna `N.F` do
recebimento. Para apenas visualizar o resultado sem alterar o arquivo, use:

```bash
python main.py --nao-salvar-recebimento
```

### Rodar apenas um CAT específico

Pula a Etapa 1 (controle financeiro) e cruza diretamente esse CAT com o
relatório de pagamentos (Etapa 2) e, em seguida, com o relatório de
recebimento integrado da ORG correspondente (Etapa 3):

```bash
python main.py --cat "Bauru"
```

### Rodar cada etapa isoladamente

Cada script também pode ser executado sozinho, via linha de comando:

```bash
# Etapa 1 — controle financeiro
python controle_financeiro.py [caminho_do_arquivo.xlsx] [--sheet NOME_DA_ABA]

# Etapa 2 — relatório de pagamentos, filtrando por um CAT
python relatorio_pagamentos.py "Bauru" [caminho_do_arquivo.xlsx] [--sheet NOME_DA_ABA]

# Etapa 3 — relatório de recebimento integrado, a partir de um JSON de pagamentos
python relatorio_recebimento.py --org "MBA" --pagamentos-json pagamentos.json \
    [--recebimento CAMINHO.xlsx] [--sheet-recebimento NOME] [--nao-salvar]
```

No caso da Etapa 3 isolada, o `--pagamentos-json` deve apontar para um
arquivo `.json` no mesmo formato dos resultados da Etapa 2, por exemplo:

```json
[{"NFF": "996.02", "Valor": 21603.16}]
```

## Observações

- A `tabela_localidade.py` deve estar na mesma pasta dos demais scripts
  (ou no `PYTHONPATH`), pois é importada por `controle_financeiro.py`,
  `relatorio_pagamentos.py` e `main.py`.
- A detecção de "amarelo" tolera pequenas variações de tom, já que
  planilhas reais às vezes usam amarelos ligeiramente diferentes.
- Como o `openpyxl` reescreve o arquivo ao salvar, valores em cache de
  fórmulas (ex.: coluna "Alíquota" no Recebimento Integrado) só são
  recalculados ao reabrir a planilha no Excel/LibreOffice — isso é uma
  limitação do próprio `openpyxl`, não um efeito deste script.