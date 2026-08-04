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

### 3. Rodar os scripts
```bash
python separar_tipo.py conciliacao.xlsx
```
Isso separa a planilha em abas **Recebimentos** e **Pagamentos** (e **Outros**, se houver linhas não classificadas).

### 4. Sair do ambiente virtual (quando terminar)
```bash
deactivate
```

> `.venv/` já está no `.gitignore` — não é versionado, cada máquina cria o seu.

## TO DO
- Desenvolvimento do Script de Renomeação e Extração de dados dos arquivos IEDDEB
- Desenvolvimento do Script da PLanilha de Controle de Transações
- Desenvolvimento do Script de Extração de dados dos Extratos Bancários
