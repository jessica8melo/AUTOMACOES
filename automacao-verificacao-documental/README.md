# Validador de Documentos - Desktop

Aplicação desktop para validação documental de fornecedores, executada localmente na máquina do analista, sem necessidade de servidor ou infraestrutura em nuvem.

## Visão Geral

O sistema foi desenvolvido para automatizar a validação de documentos anexados a Ordens de Serviço (OS), realizando classificação documental, extração de informações, validações cadastrais e conferência de assinaturas eletrônicas.

Entre as funcionalidades disponíveis estão:

- Classificação automática de documentos;
- Extração de texto a partir de PDFs e imagens;
- OCR utilizando Tesseract;
- Comparação de nomes, CPFs e CNPJs;
- Validação de documentos obrigatórios;
- Conferência de assinaturas eletrônicas;
- Geração de relatório consolidado por OS;
- Identificação de pendências documentais.

Este projeto foi derivado do sistema `ocr-validator` (Azure Function), porém possui uma diferença arquitetural importante:

> Não existe fallback com Inteligência Artificial.

Quando a leitura de um documento não apresenta confiança suficiente, o sistema não tenta inferir ou completar informações automaticamente.

Nesses casos, o documento é marcado como:

```text
⚠ Precisa verificação humana
```

e deve ser analisado pelo responsável.

Essa decisão é implementada através dos mecanismos:

```text
precisa_verificacao_humana
needs_review()
```

## Fluxos de Validação

O sistema possui dois fluxos independentes de validação.

A seleção é realizada através do campo **Tipo de Fornecedor** na interface principal.

### Fluxo COBAN

Fluxo utilizado para Credenciamento de Correspondente Bancário.

Documentos normalmente esperados:

- Cartão CNPJ;
- Contrato Social;
- Consulta Simples Nacional;
- Formulário de Credenciamento;
- Parecer COBAN;
- Comprovante Bancário (opcional);
- Procuração (opcional).

A execução é orquestrada por:

```text
process_os_folder()
```

### Fluxo Não COBAN

Fluxo utilizado para processos gerais de cadastro e atualização de fornecedores.

Baseado nos normativos:

```text
PRO415-004
TAB415-004
```

Utiliza como documento principal:

```text
FQ415-064
```

Os documentos obrigatórios variam de acordo com:

- Tipo de fornecedor;
- Tipo de solicitação;
- Pessoa Física ou Pessoa Jurídica.

A identificação entre Pessoa Física e Pessoa Jurídica é realizada automaticamente através da análise do próprio FQ415-064 utilizando:

```text
detectar_tipo_pessoa_fq064()
```

O fluxo é orquestrado por:

```text
process_os_folder_nao_coban()
```

## Arquitetura do Projeto

Os dois fluxos foram construídos de forma independente.

Cada fluxo possui:

- Configuração própria;
- Catálogo próprio de documentos;
- Validadores específicos;
- Regras de negócio independentes;
- Processo de orquestração próprio.

Os componentes compartilhados foram centralizados no pacote:

```text
src/shared/
```

onde ficam implementações reutilizadas pelos dois fluxos.

### Como decidir onde alterar o código

Se a alteração afeta apenas um fluxo específico:

```text
src/coban/
```

ou

```text
src/nao_coban/
```

Se a alteração afeta os dois fluxos:

```text
src/shared/
```

### Observação importante

Os documentos:

```text
CNPJ
CONTRATO_SOCIAL
```

existem tanto no fluxo COBAN quanto no fluxo Não COBAN.

Apesar de utilizarem os mesmos validadores internos (`common_validators.py`), os catálogos são independentes e possuem diferenças em:

- Nome exibido;
- Filename patterns;
- Keywords de classificação.

## Como o Sistema Decide Confiar em uma Leitura

Toda operação de extração gera um objeto:

```text
ExtractionResult
```

que registra:

- Método utilizado na leitura;
- Confiança estimada do OCR;
- Necessidade ou não de revisão humana.

### Texto Nativo

```text
native_text
```

Texto extraído diretamente do PDF.

É sempre considerado confiável.

### OCR

```text
ocr
```

Quando o texto é obtido por OCR, a confiança precisa atingir o limite mínimo configurado para aquele tipo de documento.

Cada documento possui sua própria configuração:

```text
min_ocr_confidence
```

Documentos mais críticos possuem limites maiores.

### Falha de Extração

```text
failed
```

Sempre exige revisão humana.

### Revisão Humana

Quando a confiança não atinge o mínimo configurado:

```text
precisa_verificacao_humana = True
```

O relatório recebe uma pendência indicando necessidade de validação manual.

Não existe fallback via IA para resolver a situação.

## Decisões de Performance

### OCR Paralelo com Limitação de Núcleos

O sistema processa documentos em paralelo utilizando:

```text
ThreadPoolExecutor
```

Porém cada execução do Tesseract é limitada a apenas um núcleo:

```text
OMP_THREAD_LIMIT=1
```

Isso evita concorrência excessiva entre múltiplos processos OCR executados simultaneamente.

### OCR Progressivo

Utilizado principalmente para:

- Contratos Sociais;
- Alterações Contratuais;
- Atas.

Em vez de processar todas as páginas de uma só vez, o sistema realiza OCR página a página e interrompe o processamento assim que encontra informações suficientes, normalmente:

- CNPJ;
- Razão Social.

### OCR de Assinatura em Paralelo

Determinados documentos podem exigir uma segunda análise focada exclusivamente na detecção de assinaturas eletrônicas.

Essa análise é submetida ao executor em paralelo ao processamento principal.

Caso a assinatura seja localizada durante a análise principal, o resultado adicional é descartado.

### OCR de Assinatura em Menor Resolução

No fluxo Não COBAN, a busca por assinaturas utiliza:

```text
dpi=150
```

Como o objetivo é apenas detectar carimbos e palavras-chave, essa resolução reduz consumo de CPU sem comprometer os resultados.

## Estrutura do Projeto

```text
app_desktop.py
│
└── src/
    │
    ├── shared/
    │   ├── constants.py
    │   ├── extractor.py
    │   ├── classifier.py
    │   ├── validation.py
    │   ├── common_validators.py
    │   ├── report.py
    │   └── zip_utils.py
    │
    ├── coban/
    │   ├── config.py
    │   ├── validators.py
    │   └── orchestrator.py
    │
    └── nao_coban/
        ├── config.py
        ├── validators.py
        └── orchestrator.py
```

### Componentes Compartilhados

#### constants.py

Configurações globais do sistema.

#### extractor.py

Extração de texto utilizando múltiplas estratégias:

```text
PDF com texto nativo
↓
OCR
↓
Falha
```

#### classifier.py

Identifica automaticamente o tipo do documento.

#### validation.py

Implementa comparações de:

- CPF;
- CNPJ;
- Nome empresarial;
- Dados cadastrais.

#### common_validators.py

Validadores compartilhados entre os dois fluxos:

- CNPJ;
- Contrato Social;
- Simples Nacional;
- Procuração.

#### report.py

Estruturas de relatório:

```text
DocumentReportRow
OsReport
```

#### zip_utils.py

Utilidades para:

- Descompactação de arquivos ZIP;
- Localização automática de pastas de OS.

## Pré-requisitos

### Desenvolvimento

- Python 3.9+
- Ambiente virtual Python (venv)
- Tesseract OCR
- Poppler

### Tesseract OCR

Utilizado para leitura OCR.

Windows:

```text
https://github.com/UB-Mannheim/tesseract/wiki
```

Instale também o idioma português.

### Poppler

Utilizado pelo pacote `pdf2image`.

Download:

```text
https://github.com/oschwartz10612/poppler-windows/releases
```

## Instalação

### Criar ambiente virtual

#### Windows PowerShell

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
```

#### Windows CMD

```cmd
python -m venv .venv
.venv\Scripts\activate.bat
```

### Instalar dependências

```powershell
pip install -r requirements.txt
```

## Como Rodar

Após instalar as dependências:

```powershell
python app_desktop.py
```

Caso Tesseract ou Poppler não estejam disponíveis no PATH, utilize a opção:

```text
Configurações...
```

para informar manualmente seus caminhos.

## Modo Desenvolvimento

Esse modo pressupõe que Python, Tesseract e Poppler estejam instalados na máquina do desenvolvedor.

Fluxo típico:

```powershell
python -m venv .venv

.venv\Scripts\activate

pip install -r requirements.txt

python app_desktop.py
```

## Distribuição para Usuário Final

O usuário final não precisa instalar:

- Python;
- Tesseract;
- Poppler.

Todos os componentes podem ser distribuídos juntamente com o executável.

### Passo 1 - Gerar o Executável

Instalar PyInstaller:

```powershell
pip install pyinstaller
```

Gerar o executável:

```powershell
pyinstaller validador-cadastro-de-fornecedores-COBAN.spec
```

Será criada a pasta:

```text
dist\validador-cadastro-de-fornecedores-COBAN\
```

### Observação sobre Hidden Imports

O projeto utiliza imports estáticos.

Normalmente não é necessário configurar:

```text
hiddenimports
```

Caso um módulo novo seja carregado dinamicamente, ele deverá ser incluído manualmente no arquivo `.spec`.

### Passo 2 - Adicionar Poppler Portátil

Baixar:

```text
https://github.com/oschwartz10612/poppler-windows/releases
```

Copiar:

```text
Library\bin
```

para:

```text
dist\
└── validador-cadastro-de-fornecedores-COBAN\
    └── poppler\
        └── bin\
```

### Passo 3 - Adicionar Tesseract Portátil

Instalar o Tesseract na máquina de desenvolvimento.

Copiar:

```text
C:\Program Files\Tesseract-OCR
```

para:

```text
dist\
└── validador-cadastro-de-fornecedores-COBAN\
    └── tesseract\
```

Confirmar a existência de:

```text
tesseract\tesseract.exe
```

e

```text
tesseract\tessdata\por.traineddata
```

### Passo 4 - Validar Estrutura Final

```text
dist\
└── validador-cadastro-de-fornecedores-COBAN\
    ├── validador-cadastro-de-fornecedores-COBAN.exe
    │
    ├── tesseract\
    │   ├── tesseract.exe
    │   └── tessdata\
    │       └── por.traineddata
    │
    └── poppler\
        └── bin\
            ├── pdftoppm.exe
            └── ...
```

### Passo 5 - Distribuição

Após validar o funcionamento:

1. Compactar toda a pasta;
2. Enviar ao usuário final;
3. O usuário apenas extrai o conteúdo;
4. Executa o `.exe`.

Nenhuma instalação adicional será necessária.

## Projeto Relacionado

Este sistema foi derivado do projeto:

```text
ocr-validator
```

implementado como Azure Function.

As regras de negócio são equivalentes:

- Classificação;
- Extração;
- Validação;
- Conferência documental.

### Validador Desktop

```text
Aplicação local
Sem IA
Sem Azure
Revisão humana obrigatória
```

### OCR Validator (Azure)

```text
API HTTP
Integração com Power Automate
Execução em nuvem
Fallback com IA
```

Como os projetos possuem bases de código independentes, alterações realizadas em um deles devem ser replicadas manualmente no outro.

## Observações

- O sistema prioriza confiabilidade acima de automação total.
- Não existe fallback utilizando IA.
- Leituras de baixa confiança nunca são aprovadas automaticamente.
- Documentos críticos exigem revisão humana quando o OCR não atinge o nível mínimo configurado.
- A validação de documentos ocorre integralmente de forma local.
- A utilização de OCR paralelo e processamento progressivo reduz significativamente o tempo de execução.
- A ausência de IA é uma decisão arquitetural intencional para reduzir riscos em validações documentais críticas.
