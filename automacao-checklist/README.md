# automacao-checklist

Automação do processo de Checklist de Ordens de Compra.

## Visão Geral

A automação realiza a análise de documentos utilizados nos processos de
Ordens de Compra, identificando automaticamente o tipo de documento,
extraindo informações relevantes e verificando a existência de indícios de
assinatura eletrônica.

O sistema suporta documentos como:

- Contratos;
- Aditivos;
- Solicitações de Entrega de Bens;
- Planilhas de Custos;
- Planilhas de Formação de Preços;
- Notas Técnicas;
- Projetos Básicos.

Também é capaz de identificar assinaturas eletrônicas emitidas por:

- D4Sign
- GOV.BR / ICP-Brasil
- DocuSign
- Clicksign
- Aprovve

## Fluxo do Processo

A automação possui dois modos principais de execução.

### Modo 1 - Processamento de Documento Individual

Executado através do script `arquivo.py`.

Fluxo:

```text
Arquivo ou Pasta
↓
Identificação do tipo real do arquivo
↓
Identificação do documento
↓
Extração dos campos configurados
↓
Verificação de assinatura eletrônica
↓
Resultado da análise
```

O sistema identifica o tipo do documento utilizando:

- Nome do arquivo;
- Conteúdo interno;
- Assinatura binária do arquivo.

Após a identificação, a extração é encaminhada para:

```text
PDFs     → pdfs.py
Planilhas → tabelas.py
```

### Modo 2 - Fluxo Completo de Checklist

Executado através do script `extracao.py`.

Fluxo:

```text
Checklist (.docx)
↓
Identificação do fluxo
↓
Localização automática dos anexos
↓
Processamento de cada documento
↓
Extração de campos
↓
Validação de assinaturas
↓
Relatório de documentos encontrados, faltantes e excedentes
```

Nesse modo o sistema identifica automaticamente:

- Quais documentos são esperados;
- Quais anexos correspondem a cada documento;
- Quais documentos obrigatórios não foram encontrados;
- Quais arquivos anexados não pertencem ao checklist.

## Estrutura do Projeto

| Arquivo | Descrição |
|----------|----------|
| `arquivo.py` | Processa documentos individuais ou uma pasta de arquivos |
| `extracao.py` | Executa o fluxo completo de checklist |
| `doc_types.py` | Identifica qual documento um arquivo representa |
| `documentos.py` | Define os documentos reconhecidos e os campos extraídos |
| `fluxos.py` | Define os fluxos de checklist suportados |
| `pdfs.py` | Extração de dados e assinaturas em PDFs |
| `tabelas.py` | Extração de dados em planilhas Excel |
| `requirements.txt` | Dependências do projeto |

## Pré-requisitos

- Python 3.9 ou superior
- Tesseract OCR instalado no sistema operacional

O Tesseract é utilizado para leitura de documentos que não possuem texto
extraível, como certificados de assinatura eletrônica salvos como imagem.

### Instalação do Tesseract

#### macOS

```bash
brew install tesseract
brew install tesseract-lang
```

#### Ubuntu / Debian

```bash
sudo apt-get install tesseract-ocr tesseract-ocr-por
```

Sem o Tesseract instalado, o sistema continua funcionando, porém páginas
compostas exclusivamente por imagens poderão não ser interpretadas
corretamente.

## Instalação

### 1. Clonar o repositório

```bash
git clone https://github.com/jessica8melo/automacao-checklist.git
cd automacao-checklist
```

### 2. Criar ambiente virtual

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

### 3. Instalar dependências

```bash
pip install -r requirements.txt
```

## Como Rodar

### Processar um Documento Específico

PDF:

```bash
python arquivo.py Anexos/DGCO_Nº_027322025.pdf
```

Planilha:

```bash
python arquivo.py "Anexos/FQ415-075_v_03_-_GMS_-_LOTE 4_-__02732-2025_-_v2_(1).xlsx"
```

### Processar uma Pasta Inteira

Processa todos os arquivos encontrados na pasta em ordem alfabética.

```bash
python arquivo.py Anexos
```

### Executar Extração de PDF Isoladamente

```bash
python pdfs.py Anexos/DGCO_Nº_027322025.pdf
```

### Executar Extração de Planilha Isoladamente

```bash
python tabelas.py "Anexos/FQ415-075_v_03_-_GMS_-_LOTE 4_-__02732-2025_-_v2_(1).xlsx"
```

### Executar o Fluxo Completo de Checklist

#### Forma recomendada

Passe a pasta contendo o checklist e seus anexos.

```bash
python extracao.py Checklists
```

O sistema localizará automaticamente:

- O checklist;
- Os documentos relacionados;
- A pasta de anexos.

#### Checklist e anexos separados

```bash
python extracao.py \
    "Checklists/FQ415-031_v10 - Checklist OC Padrão – Com Contrato.docx" \
    Anexos
```

#### Informando apenas o código do fluxo

```bash
python extracao.py FQ415-031 Anexos
```

Caso a pasta de anexos não seja informada, será utilizada a pasta padrão:

```text
Anexos
```

## Exemplo de Fluxo Completo

```bash
# Processamento de uma pasta contendo checklist e anexos

python extracao.py Checklists
```

Fluxo executado:

```text
Checklist
↓
Identificação do fluxo
↓
Identificação dos documentos esperados
↓
Localização dos anexos
↓
Extração dos campos
↓
Validação das assinaturas
↓
Relatório final
```

## Exemplo de Saída

```text
Documento analisado: Anexos/DGCO_Nº_027322025.pdf

[SUCESSO] Campo 'DGCO nº' encontrado. Valor: 00111/2022

[SUCESSO] Campo 'OC Master nº' encontrado. Valor: 196804

[ERRO] Não foi possível encontrar o campo 'Empresa' no documento.

[ASSINATURA] Documento assinado. Sistema: D4Sign

[ASSINATURA] Nº do processo:
d614a4f2-fd4e-40cc-91df-821bfaea9655

[ASSINATURA] - MONICA GUIZZARDI VAILLANT
(Assinou como parte) em 18/05/2026 14:13

[ASSINATURA] - RICARDO LOPES AUGUSTO
(Assinou como parte) em 22/05/2026 14:35
```

## Customização

### Configuração dos Campos Extraídos

Todos os documentos reconhecidos e seus respectivos campos ficam definidos
em:

```text
documentos.py
```

Para adicionar, remover ou renomear campos basta alterar a estrutura
`DOCUMENTOS`.

### Configuração dos Fluxos

Os fluxos de checklist suportados ficam definidos em:

```text
fluxos.py
```

Nesse arquivo é possível configurar:

- Documentos obrigatórios;
- Campos esperados;
- Regras de conferência;
- Validação de assinatura.

### Campos Especiais em PDFs

Alguns campos utilizam expressões regulares específicas para extração.

Exemplos:

- DGCO nº
- OC Master nº
- Data do Fornecimento

Essas regras ficam definidas em:

```text
pdfs.py
```

na estrutura:

```text
PADROES_ESPECIAIS
```

## Solução de Problemas

### Assinatura eletrônica não encontrada

Se um documento assinado não for reconhecido como assinado, normalmente a
causa é:

- Tesseract não instalado;
- Tesseract não configurado corretamente;
- Falha de OCR durante a leitura da página.

Isso é comum em certificados de assinatura salvos como imagem.

### Campo existente não encontrado

Quando um campo existe visualmente no documento, mas não é localizado pela
automação, as causas mais comuns são:

- Nome configurado diferente do rótulo real;
- Diferença excessiva de escrita;
- Limite de similaridade muito restritivo.

Nesses casos, recomenda-se:

- Ajustar o nome do campo em `documentos.py`;
- Tornar a correspondência mais tolerante ajustando a função `parecido()`.

## Observações

- A comparação entre campos procurados e textos encontrados é aproximada,
  tolerando diferenças de acentuação, maiúsculas/minúsculas, pontuação e
  pequenas variações de escrita.

- A validação de assinatura é baseada na análise textual dos carimbos
  inseridos pelas plataformas de assinatura eletrônica.

- Não é realizada validação criptográfica da assinatura digital presente no
  PDF.

## Encerramento

Ao finalizar o trabalho, desative o ambiente virtual:

```bash
deactivate
```

> A pasta `.venv` deve permanecer no `.gitignore` e não deve ser versionada.
