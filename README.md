# automacao-checklist

Automação do processo de Checklist de Ordens de Compra.

O projeto recebe PDFs e planilhas (contratos, aditivos, solicitações de
entrega de bens, planilhas de custos/formação de preços) e extrai
automaticamente um conjunto de campos-chave, além de checar se o documento
tem indícios de assinatura eletrônica (D4Sign, GOV.BR/ICP-Brasil, DocuSign,
Clicksign, Aprovve).

## Como funciona

> **Nota:** por enquanto o projeto foca só na **extração de dados dos
> documentos**. A parte de fluxos/checklists (`fluxos.py`) fica pausada
> para depois — o `main.py` não pergunta mais qual fluxo está sendo
> aplicado, ele identifica o documento e extrai seus campos direto.

- **`main.py`** — ponto de entrada. Recebe o caminho de um arquivo (ou de uma
  pasta), identifica o tipo real do arquivo pelos bytes iniciais (não só
  pela extensão) e qual DOCUMENTO ele é (Contrato, FQ415-075, Nota
  Técnica, Projeto Básico, Solicitação de Entrega — ver `documentos.py`),
  usando `doc_types.py`. Em seguida busca só os campos daquele documento
  e encaminha para o script certo:
  - PDF → `pdfs.py`
  - XLSX/XLSM/XLS → `tabelas.py`
- **`documentos.py`** — lista os documentos reconhecidos e, para cada um,
  os campos que devem ser extraídos. É o único lugar que você precisa
  editar para adicionar/remover/renomear campos de um documento.
- **`doc_types.py`** — identifica qual documento um arquivo representa,
  combinando indícios do nome do arquivo com indícios do texto/conteúdo.
- **`pdfs.py`** — abre o PDF com `pdfplumber`, extrai texto e tabelas, e
  procura os campos definidos em `CAMPOS_PROCURADOS` dentro do arquivo.
  Quando uma página não tem texto extraível (ex.: certificados de assinatura
  que são "print de tela"), cai para OCR via `pytesseract`. Também analisa o
  texto em busca de indícios de assinatura eletrônica.
- **`tabelas.py`** — abre a planilha com `openpyxl`, localiza automaticamente
  a linha de cabeçalho da tabela (comparando com `CAMPOS_PROCURADOS`) e lê os
  valores de cada linha de dados abaixo dela.

Em ambos os scripts, a comparação entre o nome de um campo procurado e o
texto do documento é **aproximada** (tolera acentos, maiúsculas/minúsculas,
pequenas variações de escrita, parênteses), não exige que o texto seja
idêntico.

## Pré-requisitos

- Python 3.9+
- [Tesseract OCR](https://github.com/tesseract-ocr/tesseract) instalado no
  sistema (não é um pacote Python — é um programa à parte, usado pelo
  `pytesseract` para ler páginas que são só imagem).

  **macOS (Homebrew):**
  ```bash
  brew install tesseract
  brew install tesseract-lang   # inclui o pacote de português ("por")
  ```

  **Ubuntu/Debian:**
  ```bash
  sudo apt-get install tesseract-ocr tesseract-ocr-por
  ```

  Sem o Tesseract instalado, o script continua rodando, mas não consegue ler
  texto de páginas 100% imagem (comum em certificados de assinatura) — e
  avisa isso no terminal (`[AVISO] pytesseract não está instalado...` ou
  falha de OCR).

## Instalação

```bash
git clone https://github.com/jessica8melo/automacao-checklist.git
cd automacao-checklist

python3 -m venv venv
source venv/bin/activate        # no Windows: venv\Scripts\activate

pip install -r requirements.txt
```

## Como rodar

**Um arquivo específico:**
```bash
python main.py Anexos/DGCO_Nº_027322025.pdf
python main.py "Anexos/FQ415-075_v_03_-_GMS_-_LOTE 4_-__02732-2025_-_v2_(1).xlsx"
```

**Uma pasta inteira** (processa todos os arquivos dentro dela, em ordem
alfabética, e imprime um resumo no final):
```bash
python main.py Anexos
```

**Rodar um dos scripts direto**, sem passar pelo `main.py` (útil para testar
isoladamente):
```bash
python pdfs.py Anexos/DGCO_Nº_027322025.pdf
python tabelas.py "Anexos/FQ415-075_v_03_-_GMS_-_LOTE 4_-__02732-2025_-_v2_(1).xlsx"
```

### Exemplo de saída (PDF)

```
Documento analisado: Anexos/DGCO_Nº_027322025.pdf

[SUCESSO] Campo 'DGCO nº' encontrado. Valor: 00111/2022
[SUCESSO] Campo 'OC Master nº' encontrado. Valor: 196804
[ERRO] Não foi possível encontrar o campo 'Empresa' no documento.
...

[ASSINATURA] Documento assinado. Sistema: D4Sign
[ASSINATURA] Nº do processo: d614a4f2-fd4e-40cc-91df-821bfaea9655
[ASSINATURA]   - MONICA GUIZZARDI VAILLANT (Assinou como parte) em 18/05/2026 14:13
[ASSINATURA]   - RICARDO LOPES AUGUSTO (Assinou como parte) em 22/05/2026 14:35
...
```

> A verificação de assinatura é uma checagem **textual** (procura o carimbo
> que a plataforma de assinatura grava no PDF), não uma verificação
> criptográfica do arquivo (`ByteRange`/assinatura digital real).

## Customizando os campos procurados

Os campos de cada documento ficam em `documentos.py` (dicionário
`DOCUMENTOS`) — edite livremente para adicionar, remover ou renomear
campos de qualquer um dos 5 documentos reconhecidos (Contrato,
FQ415-075, Nota Técnica, Projeto Básico, Solicitação de Entrega).

`pdfs.py` e `tabelas.py` também têm sua própria lista `CAMPOS_PROCURADOS`
no topo do arquivo, usada só quando você roda um deles isoladamente (sem
passar pelo `main.py`) — ver seção anterior.

Para PDFs, alguns campos que não aparecem como "Rótulo: valor" (como "DGCO
nº", "OC Master nº" e "Data do fornecimento") usam expressões regulares
próprias em `PADROES_ESPECIAIS`, dentro de `pdfs.py`.

## Solução de problemas

**"Nenhum indício de assinatura eletrônica encontrado" mesmo em documento
assinado** — geralmente é o Tesseract não estar instalado/configurado no
ambiente. Um script de diagnóstico separado testa cada dependência
(pytesseract, binário do Tesseract, pacote de idioma "por", conversão de
página em imagem) e aponta exatamente onde travou — peça-o se precisar
depurar de novo (não faz parte do fluxo principal do projeto).

**Campo não encontrado, mas ele existe no documento** — o nome em
`CAMPOS_PROCURADOS` pode estar diferente demais do texto real. Tente deixar
o nome mais parecido com o rótulo exato do documento, ou ajuste o `limite`
da função `parecido()` (padrão: `0.72`; quanto menor, mais tolerante).

## Estrutura do projeto

```
automacao-checklist/
├── main.py             # ponto de entrada, detecta tipo e encaminha
├── pdfs.py              # extração de campos e assinatura em PDFs
├── tabelas.py            # extração de campos em planilhas xlsx
├── requirements.txt       # dependências Python
├── Anexos/                # exemplos de arquivos para teste
└── README.md
```