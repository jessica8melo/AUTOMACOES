# Validador de Documentos - versão Desktop

App standalone (sem servidor) pra quem resolve a OS validar um conjunto de
documentos na própria máquina.

Diferença em relação ao projeto `ocr-validator` (Azure Function): aqui **não
existe fallback de IA**. Quando um documento tem leitura fraca demais pra
confiar (comprovante bancário borrado, PDF escaneado torto), o programa
marca **"⚠ Precisa verificação humana"**: alguém precisa olhar o arquivo
original. Essa decisão está espalhada pelo código como
`precisa_verificacao_humana` / `needs_review()`; ver seção
[Como o app decide confiar (ou não) numa leitura](#como-o-app-decide-confiar-ou-não-numa-leitura).

Este documento é a referência de arquitetura do projeto: como ele é
organizado em pacotes, por que existem dois fluxos de validação
independentes, e quais decisões de design não são óbvias só lendo o código.
Pra instruções de instalação/uso do dia a dia, veja a seção
[Modo desenvolvimento](#1-modo-desenvolvimento) e
[Montar a pasta pra distribuir](#2-montar-a-pasta-pra-distribuir-pro-usuário-final)
mais abaixo.

---

## Os dois fluxos de validação

O app valida dois tipos de pacote de documentos, escolhidos pelo combobox
"Tipo de fornecedor" na tela principal:

- **COBAN**: Credenciamento de Correspondente Bancário. Fluxo de
  produção original do projeto: Cartão CNPJ, Contrato Social, Consulta do
  Simples Nacional, Formulário de Credenciamento (assinado via D4Sign),
  Parecer COBAN e, opcionalmente, Comprovante Bancário e Procuração.
- **Não COBAN**: Cadastro/Atualização comum de fornecedor, baseado nos
  documentos normativos **PRO415-004** (processo) e **TAB415-004**
  (tabela de insumos exigidos por tipo de fornecedor). Usa o Formulário
  FQ415-064 no lugar do Formulário de Credenciamento, e os insumos exigidos
  mudam conforme o tipo escolhido (Fornecedor Contratado, Atualização de
  Dados Bancários, Cadastro de Filial, Obrigações Judiciais, etc.) e, em
  alguns tipos, conforme o cadastro ser de Pessoa Jurídica ou Pessoa
  Física; isso é detectado automaticamente a partir do próprio FQ415-064
  (`detectar_tipo_pessoa_fq064`), o analista não precisa selecionar.

Os dois fluxos são **independentes**: cada um tem seu próprio catálogo de
tipos de documento, sua própria função de orquestração
(`process_os_folder` / `process_os_folder_nao_coban`), e roda sem depender
do outro. O que eles compartilham (extração de texto, classificação,
comparação de nomes/CNPJ, e os 4 tipos de documento que são literalmente
o mesmo nos dois: CNPJ, Contrato Social, Simples Nacional, Procuração)
foi fatorado para um pacote comum. É essa divisão que organiza o código em
pacotes (próxima seção).

## Arquitetura em pacotes

```
app_desktop.py            - interface (tkinter), ponto de entrada
src/
  shared/                  - usado pelos dois fluxos; não conhece nenhum dos dois
    constants.py            - SUPPORTED_EXTENSIONS, DEFAULT_MIN_OCR_CONFIDENCE
    extractor.py             - extração de texto em camadas (nativo -> OCR)
    classifier.py             - classifica o tipo de documento (nome/conteúdo)
    validation.py              - DocumentValidation, comparação de nomes/CNPJ
    common_validators.py        - validadores IDÊNTICOS nos dois fluxos
                                   (CNPJ, Contrato Social, Simples Nacional,
                                   Procuração) + ASSINANTE_D4SIGN_RE
    report.py                  - DocumentReportRow, OsReport (dataclasses)
    zip_utils.py                - descompactar zip / localizar pastas de OS
  coban/                    - só o fluxo COBAN
    config.py                 - DOCUMENT_TYPES (catálogo de documentos)
    validators.py               - validadores específicos (Formulário de
                                   Credenciamento, Parecer COBAN,
                                   Comprovante Bancário) + VALIDATORS
    orchestrator.py              - process_os_folder / process_zip
  nao_coban/                - só o fluxo Não COBAN
    config.py                 - TIPOS_FORNECEDOR (insumos por tipo/PJ-PF)
    validators.py               - validadores específicos (FQ415-064,
                                   Documento de Identificação, CPF,
                                   Comprovante de Residência, NF/Fatura
                                   BBTS) + VALIDATORS_NAO_COBAN
    orchestrator.py              - process_os_folder_nao_coban / process_zip_nao_coban
```

Regra prática pra saber onde mexer: se a mudança é sobre **um tipo de
documento específico de um dos fluxos**, mexe só no `coban/` ou `nao_coban/`
correspondente. Se afeta **os dois fluxos** (ex.: melhorar a extração de
OCR, ou a lógica de comparação de nomes), mexe no `shared/`.

Um detalhe que gera confusão à primeira vista: **`CNPJ` e `CONTRATO_SOCIAL`
não são os mesmos dicts em `coban/config.py` e `nao_coban/config.py`**,
mesmo com os validadores por trás sendo idênticos (`common_validators.py`).
Os dois catálogos têm rótulos e alguns `filename_patterns`/`content_keywords`
ligeiramente diferentes (ex.: o rótulo do Não COBAN inclui "/ Documento
Equivalente").

## Como o app decide confiar (ou não) numa leitura

Cada extração (`ExtractionResult`, em `src/shared/extractor.py`) carrega um
método (`native_text`, `ocr`, `failed`) e, quando é OCR, uma confiança
estimada (0–100). `needs_review(min_confidence)` decide se aquela leitura é
confiável o bastante:

- `native_text` (texto extraído direto do PDF, sem OCR): sempre confiável.
- `ocr`: só é confiável se a confiança do OCR bater o limiar mínimo
  daquele tipo de documento (`min_ocr_confidence`, definido por tipo em
  `coban/config.py` / `nao_coban/config.py`; documentos mais críticos,
  como o Formulário de Credenciamento, têm limiar mais alto).
- `failed`: nunca confiável.

Quando não é confiável, a linha do relatório fica com
`precisa_verificacao_humana=True` e ganha uma pendência de topo
("Confiança de leitura insuficiente..."). Não existe fallback de IA nessa
decisão: é sempre um humano que resolve.

## Decisões de performance (não óbvias)

- **OCR em paralelo, mas com 1 núcleo por chamada do Tesseract**
  (`OMP_THREAD_LIMIT=1`, em `src/shared/extractor.py`): a partir da v4 o
  Tesseract usa OpenMP internamente e tenta sozinho usar todos os núcleos
  da máquina. Como o app já processa vários arquivos em paralelo
  (`ThreadPoolExecutor`), sem esse limite cada thread dispara um processo
  de OCR brigando pelos mesmos núcleos: o lote fica **mais lento** que
  processar sequencialmente. Limitar a 1 núcleo por chamada deixa o
  `ThreadPoolExecutor` como único responsável por dividir os núcleos.

- **OCR progressivo com parada antecipada** (`extract_progressivo`/
  `extract_ocr_progressivo`): usado para Contrato Social/Ata/Alteração, que
  costuma ser o arquivo mais pesado da OS. Em vez de sempre renderizar e
  fazer OCR em até `OCR_MAX_PAGES` páginas de uma vez, processa página a
  página e para assim que achar CNPJ + nome empresarial (que quase sempre
  estão na 1ª página).

- **OCR de assinatura disparado em paralelo com o lote principal**
  (`futures_assinatura`, nos dois `orchestrator.py`): o Formulário de
  Credenciamento/FQ415-064 às vezes precisa de uma segunda passada de OCR
  focada em achar o carimbo de assinatura (D4Sign/gov.br/certificado
  ICP-Brasil). Em vez de esperar o resto do processamento terminar pra só
  então decidir se precisa desse OCR extra, a chamada já é submetida ao
  `ThreadPoolExecutor` assim que os arquivos são descobertos, e só é
  **consultada** (`future.result()`) mais tarde, se o texto nativo não
  tiver confirmado a assinatura. Se não for necessária, o resultado é
  simplesmente descartado sem custo adicional de espera:
  `executor.shutdown(wait=False)` garante que o programa não fica preso
  esperando um OCR que ninguém vai usar.

- **OCR de assinatura em resolução mais baixa** (`dpi=150` no fluxo Não
  COBAN): esse OCR só precisa **detectar** um carimbo/palavra-chave
  (D4Sign, gov.br, "NOME:CPF"), não ler dado fino; resolução mais baixa
  já basta e é bem mais barata em CPU que os 300 DPI padrão usados pra
  extração de dados.

---

## 1. Modo desenvolvimento

Precisa de Python, venv, Tesseract e Poppler instalados na SUA máquina.

### Tesseract e Poppler (na sua máquina)
- **Tesseract OCR** (com pacote de idioma português):
  https://github.com/UB-Mannheim/tesseract/wiki (instalador Windows)
- **Poppler** (usado pelo `pdf2image` pra converter PDF em imagem):
  https://github.com/oschwartz10612/poppler-windows/releases

### Python
```powershell
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

### Rodar
```powershell
python app_desktop.py
```
Se Tesseract/Poppler não estiverem no PATH do Windows, abre o app e clica
em **Configurações...** pra apontar os caminhos.

---

## 2. Montar a pasta pra distribuir pro usuário final

O usuário final **não** vai ter Tesseract/Poppler instalados. A solução:
você monta uma pasta com tudo dentro (binários portáteis, sem instalador)
e o programa detecta sozinho; não precisa de configuração manual do lado
de quem usa.

### Passo 1: Gerar o `.exe`
```powershell
pip install pyinstaller
pyinstaller validador-cadastro-de-fornecedores-COBAN.spec
```
Isso cria `dist\validador-cadastro-de-fornecedores-COBAN\` com o `.exe` e
as dependências Python já embutidas (pdfplumber, pytesseract, etc). O
`.spec` usa detecção automática de imports do PyInstaller: como todos os
imports do projeto são estáticos (`from src.coban.orchestrator import
...`), não é preciso listar `hiddenimports` manualmente; se um novo módulo
for adicionado ao pacote e o `.exe` gerado reclamar de um import faltando,
é sinal de que esse módulo só é referenciado dinamicamente em algum lugar;
nesse caso, adicione-o em `hiddenimports` no `.spec`.

### Passo 2: Baixar os binários portáteis do Poppler
Baixa o zip da última release em
https://github.com/oschwartz10612/poppler-windows/releases (não precisa de
instalador, já vem pronto). Copia a pasta `Library\bin` de dentro do zip
pra dentro de `dist\validador-cadastro-de-fornecedores-COBAN\poppler\bin\`.

### Passo 3: Conseguir os binários portáteis do Tesseract
O instalador do Tesseract (UB-Mannheim) não é portátil por padrão, mas o
próprio instalador escreve os arquivos numa pasta comum
(`C:\Program Files\Tesseract-OCR`) que É portátil depois de instalado: só
copiar essa pasta inteira funciona em outra máquina, sem precisar rodar o
instalador lá:
1. Instala o Tesseract normalmente na SUA máquina (uma vez só), marcando o
   pacote de idioma **Portuguese** na tela de seleção de componentes.
2. Copia a pasta inteira `C:\Program Files\Tesseract-OCR` pra dentro de
   `dist\validador-cadastro-de-fornecedores-COBAN\tesseract\` (renomeando
   pra só `tesseract`, sem o `-OCR`).
3. Confirma que existe
   `dist\validador-cadastro-de-fornecedores-COBAN\tesseract\tesseract.exe`
   e `...\tesseract\tessdata\por.traineddata`.

### Passo 4: Conferir a estrutura final
```
dist\validador-cadastro-de-fornecedores-COBAN\
    validador-cadastro-de-fornecedores-COBAN.exe
    (arquivos internos do PyInstaller)
    tesseract\
        tesseract.exe
        tessdata\
            por.traineddata
            ...
    poppler\
        bin\
            pdftoppm.exe
            ...
```

### Passo 5: Testar e distribuir
Abre o `.exe`: a tela principal não deve mostrar nenhum aviso de
"⚠ Não encontrado" (Tesseract/Poppler). Se mostrar, confere os caminhos do
Passo 2/3.

Depois de confirmar que funciona, zipa a pasta inteira e manda pra quem for
usar. É só extrair e clicar duas vezes no `.exe`; nada mais precisa ser
instalado.

---

## Sobre o projeto irmão (`ocr-validator`, Azure Function)

Esse projeto standalone foi criado a partir do `ocr-validator` (mantido à
parte, pra quando for possível fazer o deploy real da Function no Azure;
ver histórico de decisões sobre custo de Container Registry/Container Apps
antes de deployar). As regras de negócio (extração, classificação,
validação) são as mesmas nos dois projetos; a diferença é só a "casca":
aqui é um app desktop sem IA, lá é uma API HTTP com fallback de IA via
Power Automate.

Se alguma regra de validação for corrigida num dos dois projetos, replicar
a mudança no outro manualmente (não são a mesma cópia de código).
