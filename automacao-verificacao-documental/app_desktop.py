r"""
Validador de Documentos - versão desktop, standalone.
"""

import json
import logging
import os
import shutil
import subprocess
import sys
import tempfile
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

if sys.platform == "win32":
    _popen_original = subprocess.Popen

    def _popen_sem_janela(*args, **kwargs):
        kwargs["creationflags"] = (
            kwargs.get("creationflags", 0) | subprocess.CREATE_NO_WINDOW
        )
        return _popen_original(*args, **kwargs)

    subprocess.Popen = _popen_sem_janela


sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.coban.orchestrator import process_os_folder
from src.nao_coban.config import TIPOS_FORNECEDOR
from src.nao_coban.orchestrator import process_os_folder_nao_coban


LOG_PATH = os.path.join(os.path.expanduser("~"), ".validador_documentos_log.txt")
logging.basicConfig(
    filename=LOG_PATH,
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    encoding="utf-8",
)
logger = logging.getLogger(__name__)


CONFIG_PATH = os.path.join(os.path.expanduser("~"), ".validador_documentos_config.json")
CAMPOS_INTERNOS = {
    "nomes_assinantes_d4sign",
    "nome_empresarial_referencia",
}

# Campos formatados para exibição na interface.

NOMES_CAMPOS_FORMATADOS = {
    "cnpj": "CNPJ",
    "cpf": "CPF",
    "nome_empresarial": "Razão Social",
    "situacao_cadastral": "Situação cadastral",
    "data_emissao": "Data de emissão",
    "ano_emissao": "Ano de emissão",
    "validade": "Validade",
    "optante_simples": "Optante pelo Simples Nacional",
    "situacao_simples": "Situação no Simples Nacional",
    "responsavel_analise": "Responsável pela análise",
    "data_analise": "Data da análise",
    "tipo_contribuinte": "Tipo de contribuinte",
    "banco": "Banco",
    "agencia": "Agência",
    "conta": "Conta",
    "nome": "Nome",
    "nome_confere_fq064": "Nome confere com o FQ415-064",
    "outorgante": "Outorgante",
    "nomes_socios": "Nome(s) do(s) sócio(s) / quem assina",
    "nome_socio_conferido": "Sócio confere com o Contrato Social",
    "nome_empresarial_conferido": "Razão social confere com o Cartão CNPJ",
    "metodo_assinatura_socio": "Método de assinatura do sócio",
    "socio_assinante": "Sócio que assinou",
    "assinatura_via_procuracao": "Assinatura via procuração",
    "procurador_confirmado": "Procurador confirmado",
    "tipo_pessoa": "Tipo de pessoa (PJ/PF)",
    "assinatura_confirmada": "Assinatura confirmada",
    "metodo_assinatura": "Método de assinatura",
}

VALORES_CAMPOS_FORMATADOS = {
    "True": "Sim",
    "False": "Não",
    "None": "Não identificado",
    "NAO_IDENTIFICADO": "Não identificado",
    "NAO_VERIFICAVEL": "Não verificável",
    "NAO_CONFIRMADO": "Não confirmado",
    "": "(não identificado)",
}

METODOS_LEITURA_FORMATADOS = {
    "native_text": "Texto do PDF",
    "ocr": "OCR (leitura de imagem)",
    "failed": "Falha na leitura",
    "skipped": "Não lido (fora do escopo)",
}


def _nome_campo_amigavel(chave: str) -> str:
    return NOMES_CAMPOS_FORMATADOS.get(chave, chave.replace("_", " ").capitalize())


def _valor_campo_amigavel(valor) -> str:
    texto = "" if valor is None else str(valor)
    return VALORES_CAMPOS_FORMATADOS.get(texto, texto)


def _metodo_leitura_amigavel(metodo: str) -> str:
    return METODOS_LEITURA_FORMATADOS.get(metodo, metodo)


def _pasta_do_executavel():
    """Pasta onde está o .exe (modo empacotado) ou o .py (modo dev)."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def _aquecer_ocr():
    """Pré-aquece Tesseract/Poppler em background assim que o app abre,
    antes do usuário selecionar qualquer pasta. Evita que o custo do
    'cold start'."""
    try:
        import pytesseract

        tess_cmd = os.environ.get("TESSERACT_CMD")
        if tess_cmd:
            pytesseract.pytesseract.tesseract_cmd = tess_cmd
        pytesseract.get_tesseract_version()
    except Exception:
        pass

    try:
        poppler_path = os.environ.get("POPPLER_PATH")
        pdftoppm = (
            os.path.join(poppler_path, "pdftoppm.exe") if poppler_path else "pdftoppm"
        )
        subprocess.run([pdftoppm, "-v"], capture_output=True, timeout=10)
    except Exception:
        pass


def _detectar_binarios_ao_lado():
    base = _pasta_do_executavel()
    achados = {}

    tess_exe = os.path.join(base, "tesseract", "tesseract.exe")
    if os.path.exists(tess_exe):
        achados["tesseract_cmd"] = tess_exe

    poppler_bin = os.path.join(base, "poppler", "bin")
    if os.path.isdir(poppler_bin):
        achados["poppler_path"] = poppler_bin

    return achados


def carregar_config():
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def salvar_config(cfg):
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def aplicar_config_ambiente(cfg):
    """TESSERACT_CMD/POPPLER_PATH precisam estar no ambiente ANTES de
    importar src.extractor.
    """
    if cfg.get("tesseract_cmd"):
        os.environ["TESSERACT_CMD"] = cfg["tesseract_cmd"]
    if cfg.get("poppler_path"):
        os.environ["POPPLER_PATH"] = cfg["poppler_path"]


_cfg_inicial = {}
_cfg_inicial.update(carregar_config())
_cfg_inicial.update(_detectar_binarios_ao_lado())
aplicar_config_ambiente(_cfg_inicial)

tess_ok = bool(os.environ.get("TESSERACT_CMD") or shutil.which("tesseract"))
poppler_ok = bool(os.environ.get("POPPLER_PATH") or shutil.which("pdftoppm"))
logger.info(
    "Startup - Tesseract: %s (%s) | Poppler: %s (%s)",
    "OK" if tess_ok else "NAO ENCONTRADO",
    os.environ.get("TESSERACT_CMD", "PATH do sistema"),
    "OK" if poppler_ok else "NAO ENCONTRADO",
    os.environ.get("POPPLER_PATH", "PATH do sistema"),
)

COR_VALIDO = "#e6f4ea"
COR_INVALIDO = "#fdeaea"
COR_HUMANO = "#fff4e0"


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Validador de Documentos")
        self.geometry("1000x600")
        self.minsize(800, 500)
        # Abre sempre maximizado.
        try:
            self.state("zoomed")
        except tk.TclError:
            try:
                self.attributes("-zoomed", True)
            except tk.TclError:
                pass

        self.pasta_atual = None
        self.linhas_detalhe = {}

        self._opcoes_tipo_fornecedor = [("COBAN", "Correspondente (COBAN)")] + [
            (chave, cfg["label"]) for chave, cfg in TIPOS_FORNECEDOR.items()
        ]
        self.tipo_fornecedor_var = tk.StringVar(
            value=self._opcoes_tipo_fornecedor[0][1]
        )

        self._monta_interface()

        threading.Thread(target=_aquecer_ocr, daemon=True).start()

    def _texto_status_binarios(self):
        tess_ok = bool(os.environ.get("TESSERACT_CMD") or shutil.which("tesseract"))
        poppler_ok = bool(os.environ.get("POPPLER_PATH") or shutil.which("pdftoppm"))
        if tess_ok and poppler_ok:
            return ""
        faltando = []
        if not tess_ok:
            faltando.append("Tesseract")
        if not poppler_ok:
            faltando.append("Poppler")
        return f"⚠ Não encontrado: {', '.join(faltando)}. Veja Configurações..."

    def _atualizar_estado_botao_validar(self, pronto: bool):
        if pronto:
            self.btn_validar.config(
                state="normal", bg="#0078d4", fg="white", activebackground="#005a9e"
            )
        else:
            self.btn_validar.config(state="disabled", bg="#cccccc", fg="#888888")

    # ------------------------------------------------------------------ UI

    def _monta_interface(self):
        # Barra superior: seleção do tipo de fornecedor e botões de ação.
        topo = ttk.Frame(self, padding=10)
        topo.pack(fill="x")

        BUTTON_STYLE = {
            "relief": "flat",
            "padx": 12,
            "pady": 4,
            "bg": "#e0e0e0",
            "fg": "#222222",
            "activebackground": "#d0d0d0",
        }

        linha_tipo = ttk.Frame(topo)
        linha_tipo.pack(fill="x", pady=(0, 10))
        ttk.Label(
            linha_tipo, text="Tipo de fornecedor:", font=("Segoe UI", 9, "bold")
        ).pack(side="left", padx=(0, 8))
        largura_combo = (
            max(len(label) for _chave, label in self._opcoes_tipo_fornecedor) + 2
        )
        combo_tipo = ttk.Combobox(
            linha_tipo,
            textvariable=self.tipo_fornecedor_var,
            state="readonly",
            width=largura_combo,
            values=[label for _chave, label in self._opcoes_tipo_fornecedor],
        )
        combo_tipo.pack(side="left", fill="x", expand=True)

        linha_acoes = ttk.Frame(topo)
        linha_acoes.pack(fill="x")

        tk.Button(
            linha_acoes,
            text="Selecionar arquivos...",
            command=self.selecionar_arquivos,
            **BUTTON_STYLE,
        ).pack(side="left")
        tk.Button(
            linha_acoes,
            text="Selecionar pasta...",
            command=self.selecionar_pasta,
            **BUTTON_STYLE,
        ).pack(side="left", padx=(8, 0))

        self.btn_validar = tk.Button(
            linha_acoes,
            text="Validar documentos",
            command=self.validar,
            state="disabled",
            relief="flat",
            padx=12,
            pady=4,
            bg="#cccccc",
            fg="#888888",
            activebackground="#005a9e",
            activeforeground="white",
        )
        self.btn_validar.pack(side="left", padx=(8, 0))
        self.btn_config = tk.Button(
            linha_acoes,
            text="Configurações...",
            command=self.abrir_configuracoes,
            **BUTTON_STYLE,
        )
        if self._texto_status_binarios():
            self.btn_config.pack(side="right")

        self.label_pasta = ttk.Label(
            self, text="Nenhum documento selecionado.", padding=(10, 0)
        )
        self.label_pasta.pack(fill="x")

        texto_status = self._texto_status_binarios()
        self.label_status_bin = ttk.Label(
            self, text=texto_status, padding=(10, 0), foreground="#666"
        )
        if texto_status:
            self.label_status_bin.pack(fill="x")

        self.progress = ttk.Progressbar(self, mode="indeterminate")

        self.label_resumo = ttk.Label(
            self, text="", padding=(10, 5), font=("Segoe UI", 10, "bold")
        )
        self.label_resumo.pack(fill="x")

        # Painel de justificativa consolidada: esses dados vêm de
        # documentos DIFERENTES (CNPJ/Razão Social do Cartão CNPJ, Nome do
        # Sócio do cruzamento Contrato Social x Formulário, Agência/Banco/
        # Conta do Comprovante Bancário).

        self.frame_resumo = ttk.LabelFrame(
            self, text="Dados da empresa (justificativa)", padding=10
        )
        self.label_justificativa = ttk.Label(
            self.frame_resumo, text="", justify="left", font=("Consolas", 9)
        )
        self.label_justificativa.pack(fill="x", anchor="w")

        colunas = ("arquivo", "tipo", "status", "metodo", "confianca")
        self.tabela = ttk.Treeview(
            self, columns=colunas, show="headings", selectmode="browse"
        )
        self.tabela.heading("arquivo", text="Arquivo")
        self.tabela.heading("tipo", text="Tipo de documento")
        self.tabela.heading("status", text="Status")
        self.tabela.heading("metodo", text="Método de leitura")
        self.tabela.heading("confianca", text="Confiança")
        self.tabela.column("arquivo", width=260)
        self.tabela.column("tipo", width=220)
        self.tabela.column("status", width=170)
        self.tabela.column("metodo", width=120, anchor="center")
        self.tabela.column("confianca", width=90, anchor="center")

        self.tabela.tag_configure("valido", background=COR_VALIDO)
        self.tabela.tag_configure("invalido", background=COR_INVALIDO)
        self.tabela.tag_configure("humano", background=COR_HUMANO)

        self.tabela.pack(fill="both", expand=True, padx=10, pady=(0, 5))
        self.tabela.bind("<Double-1>", self.mostrar_detalhe)

        ttk.Label(
            self,
            text="Dê um duplo clique numa linha pra ver os detalhes (pendências e campos extraídos).",
            padding=(10, 0),
            foreground="#666",
        ).pack(fill="x", pady=(0, 10))

    # ------------------------------------------------------------- Seleção

    def selecionar_pasta(self):
        pasta = filedialog.askdirectory(title="Selecione a pasta com os documentos")
        if pasta:
            self.pasta_atual = pasta
            self.label_pasta.config(text=f"Pasta selecionada: {pasta}")
            self._atualizar_estado_botao_validar(True)

    def selecionar_arquivos(self):
        arquivos = filedialog.askopenfilenames(
            title="Selecione os documentos",
            filetypes=[
                ("Documentos suportados", "*.pdf *.jpg *.jpeg *.png"),
                ("Todos os arquivos", "*.*"),
            ],
        )
        if not arquivos:
            return
        # Copia pra uma pasta temporária pra reaproveitar a mesma lógica de
        # processamento de pasta (process_os_folder).
        pasta_temp = tempfile.mkdtemp(prefix="validador_")
        for origem in arquivos:
            destino = os.path.join(pasta_temp, os.path.basename(origem))
            shutil.copy2(origem, destino)
        self.pasta_atual = pasta_temp
        self.label_pasta.config(text=f"{len(arquivos)} arquivo(s) selecionado(s).")
        self._atualizar_estado_botao_validar(True)

    # ------------------------------------------------------------ Validação

    def validar(self):
        if not self.pasta_atual:
            return
        self.btn_validar.config(state="disabled")
        self.tabela.delete(*self.tabela.get_children())
        self.linhas_detalhe.clear()
        self.label_resumo.config(text="")
        self.progress.pack(fill="x", padx=10, pady=(0, 5))
        self.progress.start(12)

        thread = threading.Thread(target=self._validar_em_background, daemon=True)
        thread.start()

    def _tipo_fornecedor_selecionado(self):
        """Converte o texto exibido no combobox de volta pra chave interna
        ('COBAN', 'FORNECEDOR_CONTRATADO', ...)."""
        label_atual = self.tipo_fornecedor_var.get()
        for chave, label in self._opcoes_tipo_fornecedor:
            if label == label_atual:
                return chave
        return "COBAN"

    def _validar_em_background(self):
        try:
            # Extração fraca vira "precisa de verificação humana".
            tipo = self._tipo_fornecedor_selecionado()
            if tipo == "COBAN":
                report = process_os_folder(self.pasta_atual, os_id="MANUAL")
            else:
                report = process_os_folder_nao_coban(
                    self.pasta_atual, os_id="MANUAL", tipo_fornecedor=tipo
                )

            erro = None
        except Exception as exc:
            report = None
            erro = str(exc)
        self.after(0, self._mostrar_resultado, report, erro, tipo)

    def _mostrar_resultado(self, report, erro, tipo="COBAN"):
        self.progress.stop()
        self.progress.pack_forget()
        self.btn_validar.config(state="normal")

        if erro:
            messagebox.showerror(
                "Erro ao validar",
                "Não foi possível concluir a validação dos documentos.\n\n"
                f"Detalhe técnico: {erro}",
            )
            return

        validos = sum(1 for r in report.rows if r.valido)
        humanos = sum(1 for r in report.rows if r.precisa_verificacao_humana)
        invalidos = len(report.rows) - validos

        resumo = (
            f"{len(report.rows)} documento(s): {validos} válido(s), {invalidos} com pendência "
            f"(processado em {report.tempo_processamento_segundos:.0f}s)"
        )
        if humanos:
            resumo += f" ({humanos} precisam de verificação humana)"
        if report.documentos_faltantes:
            resumo += f" | FALTANDO: {', '.join(report.documentos_faltantes)}"
        self.label_resumo.config(text=resumo)

        if tipo == "COBAN":
            self._montar_justificativa(report)
        else:
            self._montar_justificativa_nao_coban(report, tipo)

        for row in report.rows:
            if row.precisa_verificacao_humana:
                status, tag = "⚠ Verificação humana", "humano"
            elif row.valido:
                status, tag = "✓ Válido", "valido"
            else:
                status, tag = "✗ Pendência", "invalido"

            confianca = (
                f"{row.confianca_ocr:.0f}%" if row.confianca_ocr is not None else "-"
            )
            iid = self.tabela.insert(
                "",
                "end",
                values=(
                    row.filename,
                    row.doc_label,
                    status,
                    _metodo_leitura_amigavel(row.metodo_leitura),
                    confianca,
                ),
                tags=(tag,),
            )
            self.linhas_detalhe[iid] = row

    def _montar_justificativa(self, report):
        """Junta dados que vêm de documentos DIFERENTES num resumo único:
        CNPJ/Razão Social (Cartão CNPJ), Nome do Sócio (cruzamento
        Contrato Social x Formulário), Agência/Banco/Conta (Comprovante
        Bancário), Assinatura (Formulário + confirmação no Contrato)."""
        por_tipo = {r.doc_type: r for r in report.rows}

        cartao = por_tipo.get("CARTAO_CNPJ") or por_tipo.get("CNPJ")
        contrato = por_tipo.get("CONTRATO_SOCIAL")
        comprovante = por_tipo.get("COMPROVANTE_BANCARIO")
        formulario = por_tipo.get("FORMULARIO_CREDENCIAMENTO")

        cnpj = (cartao.campos.get("cnpj") if cartao else "") or "não identificado"
        razao_social = (
            cartao.campos.get("nome_empresarial") if cartao else ""
        ) or "não identificada"

        # "socio_assinante" é o nome (dentre os sócios listados no
        # Formulário) que orchestrator.py identificou como quem realmente
        # assinou; cai para o primeiro nome de "nomes_socios" quando isso
        # não foi determinado (ex.: nome_socio_conferido="NAO_VERIFICAVEL").
        nome_socio_ref = (
            formulario.campos.get("socio_assinante", "") if formulario else ""
        )
        if not nome_socio_ref and formulario:
            primeiro_socio = formulario.campos.get("nomes_socios", "").split(";")[0]
            nome_socio_ref = primeiro_socio.strip()

        if formulario and formulario.campos.get("nome_socio_conferido") == "True":
            nome_socio = nome_socio_ref + " (confirmado no Contrato Social)"
        elif formulario and formulario.campos.get("nome_socio_conferido") == "False":
            nome_socio = nome_socio_ref + " (NÃO encontrado no Contrato Social ⚠)"
        elif (
            formulario
            and formulario.campos.get("nome_socio_conferido") == "NAO_VERIFICAVEL"
        ):
            nome_socio = "não verificável (assinatura provavelmente via carimbo/imagem, não D4Sign)"
        else:
            nome_socio = "não identificado (verificar manualmente)"

        if comprovante:
            partes = []
            if comprovante.campos.get("banco"):
                partes.append(f"Banco {comprovante.campos['banco']}")
            if comprovante.campos.get("agencia"):
                partes.append(f"Agência {comprovante.campos['agencia']}")
            if comprovante.campos.get("conta"):
                partes.append(f"Conta {comprovante.campos['conta']}")
            dados_bancarios = (
                ", ".join(partes)
                if partes
                else "não identificados (verificar manualmente)"
            )
        else:
            dados_bancarios = "Comprovante Bancário não encontrado no pacote"

        if formulario and contrato:
            conferido = formulario.campos.get("nome_socio_conferido")
            if conferido == "True":
                assinou = "Sim (confirmado)"
            elif conferido == "NAO_VERIFICAVEL":
                assinou = "Não verificável (provável assinatura via carimbo/imagem) ⚠"
            else:
                assinou = "NÃO confirmado ⚠"
        elif formulario:
            assinou = "assinatura via D4Sign presente, mas sem Contrato Social pra conferir o nome"
        else:
            assinou = "Formulário de Credenciamento não encontrado no pacote"

        texto = (
            f"CNPJ:                {cnpj}\n"
            f"Razão Social:        {razao_social}\n"
            f"Nome do Sócio:       {nome_socio}\n"
            f"Dados Bancários:     {dados_bancarios}\n"
            f"Sócio assinou o Formulário de Credenciamento?  {assinou}"
        )
        self.label_justificativa.config(text=texto)
        self.frame_resumo.pack(fill="x", padx=10, pady=(0, 5))

    def _montar_justificativa_nao_coban(self, report, tipo_fornecedor_chave):
        """Monta o painel de justificativa do fluxo Não COBAN: os campos
        exibidos variam conforme o tipo de fornecedor e se o cadastro é de
        Pessoa Jurídica ou Pessoa Física."""
        por_tipo = {}
        for r in report.rows:
            por_tipo.setdefault(r.doc_type, r)

        tipo_cfg = TIPOS_FORNECEDOR.get(tipo_fornecedor_chave, {})
        depende_de_pessoa = "insumos" not in tipo_cfg
        formulario = por_tipo.get("FORMULARIO_CADASTRO_FORNECEDOR")
        tipo_pessoa = formulario.campos.get("tipo_pessoa") if formulario else None
        if tipo_pessoa not in ("PJ", "PF"):
            tipo_pessoa = None

        if depende_de_pessoa:
            doc_types_exigidos = (
                tipo_cfg.get("insumos_pf")
                if tipo_pessoa == "PF"
                else tipo_cfg.get("insumos_pj")
            ) or {}
        else:
            doc_types_exigidos = tipo_cfg.get("insumos", {})

        cnpj_row = por_tipo.get("CNPJ")
        contrato_row = por_tipo.get("CONTRATO_SOCIAL")
        doc_id_row = por_tipo.get("DOCUMENTO_IDENTIFICACAO")
        cpf_row = por_tipo.get("CPF")
        comprovante_row = por_tipo.get("COMPROVANTE_RESIDENCIA")

        linhas = []

        pf_ativo = (
            tipo_pessoa == "PF"
            or "DOCUMENTO_IDENTIFICACAO" in doc_types_exigidos
            or "CPF" in doc_types_exigidos
        )
        if pf_ativo:
            nome_pf = (doc_id_row.campos.get("nome") if doc_id_row else "") or (
                formulario.campos.get("nomes_socios", "") if formulario else ""
            )
            cpf_pf = (doc_id_row.campos.get("cpf") if doc_id_row else "") or (
                cpf_row.campos.get("cpf") if cpf_row else ""
            )
            linhas.append(("Nome", nome_pf or "não identificado"))
            linhas.append(("CPF", cpf_pf or "não identificado"))
        elif "CNPJ" in doc_types_exigidos or cnpj_row or contrato_row:
            cnpj = (cnpj_row.campos.get("cnpj") if cnpj_row else "") or (
                contrato_row.campos.get("cnpj") if contrato_row else ""
            )
            razao_social = (
                (cnpj_row.campos.get("nome_empresarial") if cnpj_row else "")
                or (contrato_row.campos.get("nome_empresarial") if contrato_row else "")
                or (formulario.campos.get("nome_empresarial") if formulario else "")
            )
            linhas.append(("CNPJ", cnpj or "não identificado"))
            linhas.append(("Razão Social", razao_social or "não identificada"))

        if formulario:
            nome_assinante = (
                formulario.campos.get("nomes_socios", "") or "não identificado"
            )
            if "CONTRATO_SOCIAL" in doc_types_exigidos:
                if formulario.campos.get("assinatura_via_procuracao") == "True":
                    situacao = (
                        f"{nome_assinante} (procurador confirmado via Procuração)"
                    )
                elif formulario.campos.get("nome_socio_conferido") == "True":
                    situacao = f"{nome_assinante} (confirmado no Contrato Social)"
                elif formulario.campos.get("nome_socio_conferido") == "False":
                    situacao = f"{nome_assinante} (NÃO encontrado no Contrato Social nem na Procuração ⚠)"
                else:
                    situacao = f"{nome_assinante} (verificar manualmente)"
                linhas.append(("Quem assina o FQ415-064", situacao))
            else:
                # Nos tipos sem Contrato Social exigido, quem assina não
                # precisa ser sócio/representante legal (pode ser o gestor
                # da área solicitante)
                linhas.append(("Quem assina o FQ415-064", nome_assinante))

            assinatura = formulario.campos.get("assinatura_confirmada")
            if assinatura == "True":
                metodo = formulario.campos.get("metodo_assinatura", "")
                texto_assinatura = "Sim" + (f" ({metodo})" if metodo else "")
            elif assinatura == "False":
                texto_assinatura = "NÃO confirmada ⚠"
            else:
                texto_assinatura = "não verificável"
            linhas.append(("Assinatura do FQ415-064 confirmada?", texto_assinatura))

            optante = formulario.campos.get("optante_simples")
            if optante in ("True", "False"):
                linhas.append(
                    (
                        "Optante pelo Simples Nacional",
                        "Sim" if optante == "True" else "Não",
                    )
                )
        else:
            linhas.append(("FQ415-064", "não encontrado no pacote"))

        if pf_ativo and doc_id_row:
            confere = doc_id_row.campos.get("nome_confere_fq064")
            if confere == "True":
                linhas.append(("Nome confere com o documento de identificação", "Sim"))
            elif confere == "False":
                linhas.append(
                    ("Nome confere com o documento de identificação", "NÃO ⚠")
                )

        if "COMPROVANTE_RESIDENCIA" in doc_types_exigidos:
            if comprovante_row:
                data_emissao = comprovante_row.campos.get("data_emissao", "")
                valor = (
                    f"emitido em {data_emissao}"
                    if data_emissao
                    else "data de emissão não identificada (verificar manualmente)"
                )
            else:
                valor = "não encontrado no pacote"
            linhas.append(("Comprovante de Residência", valor))

        if not linhas:
            self.frame_resumo.pack_forget()
            return

        largura_rotulo = max(len(rotulo) for rotulo, _valor in linhas) + 2
        texto = "\n".join(
            f"{(rotulo + ':').ljust(largura_rotulo)} {valor}"
            for rotulo, valor in linhas
        )
        self.label_justificativa.config(text=texto)
        self.frame_resumo.pack(fill="x", padx=10, pady=(0, 5))

    def mostrar_detalhe(self, _event):
        selecionado = self.tabela.selection()
        if not selecionado:
            return
        row = self.linhas_detalhe.get(selecionado[0])
        if not row:
            return

        janela = tk.Toplevel(self)
        janela.title(row.filename)
        janela.geometry("560x420")

        texto = tk.Text(janela, wrap="word", padx=10, pady=10)
        texto.pack(fill="both", expand=True)

        texto.insert("end", f"Arquivo: {row.filename}\n")
        texto.insert("end", f"Tipo: {row.doc_label}\n")
        texto.insert(
            "end",
            f"Método de leitura: {_metodo_leitura_amigavel(row.metodo_leitura)}\n",
        )
        if row.confianca_ocr is not None:
            texto.insert("end", f"Confiança: {row.confianca_ocr:.0f}%\n")
        texto.insert("end", "\n")

        if row.precisa_verificacao_humana:
            texto.insert("end", "⚠ PRECISA VERIFICAÇÃO HUMANA\n\n")

        if row.pendencias:
            texto.tag_configure("pendencia", foreground="#c62828")
            texto.insert("end", "Pendências:\n")
            for p in row.pendencias:
                texto.insert("end", f"  - {p}\n", "pendencia")
            texto.insert("end", "\n")

        campos_visiveis = {
            k: v for k, v in row.campos.items() if k not in CAMPOS_INTERNOS
        }
        if campos_visiveis:
            texto.insert("end", "Campos identificados:\n")
            for chave, valor in campos_visiveis.items():
                if chave == "outorgados":
                    nomes = [n for n in str(valor).split(";") if n]
                    rotulo = (
                        "Outorgado" if len(nomes) <= 1 else "Outorgados (procuradores)"
                    )
                    valor_exibido = (
                        "; ".join(nomes) if nomes else VALORES_CAMPOS_FORMATADOS[""]
                    )
                    texto.insert("end", f"  {rotulo}: {valor_exibido}\n")
                    continue
                texto.insert(
                    "end",
                    f"  {_nome_campo_amigavel(chave)}: {_valor_campo_amigavel(valor)}\n",
                )

        texto.config(state="disabled")

    # --------------------------------------------------------- Configurações

    def abrir_configuracoes(self):
        cfg = carregar_config()
        janela = tk.Toplevel(self)
        janela.title("Configurações")
        janela.geometry("560x280")
        janela.resizable(False, False)

        ttk.Label(
            janela,
            text="Só precisa preencher se o programa não achar Tesseract/Poppler\n"
            "sozinho (pasta 'tesseract\\' e 'poppler\\bin\\' ao lado do .exe,\n"
            "ou já instalados no PATH do Windows).",
            padding=10,
        ).pack(fill="x")

        frame = ttk.Frame(janela, padding=10)
        frame.pack(fill="x")

        ttk.Label(frame, text="Caminho do tesseract.exe:").grid(
            row=0, column=0, sticky="w"
        )
        var_tess = tk.StringVar(value=cfg.get("tesseract_cmd", ""))
        ttk.Entry(frame, textvariable=var_tess, width=50).grid(
            row=1, column=0, sticky="we"
        )
        ttk.Button(
            frame,
            text="...",
            command=lambda: var_tess.set(
                filedialog.askopenfilename(title="tesseract.exe") or var_tess.get()
            ),
        ).grid(row=1, column=1, padx=(4, 0))

        ttk.Label(frame, text="Pasta do Poppler (bin):").grid(
            row=2, column=0, sticky="w", pady=(10, 0)
        )
        var_poppler = tk.StringVar(value=cfg.get("poppler_path", ""))
        ttk.Entry(frame, textvariable=var_poppler, width=50).grid(
            row=3, column=0, sticky="we"
        )
        ttk.Button(
            frame,
            text="...",
            command=lambda: var_poppler.set(
                filedialog.askdirectory(title="Pasta bin do Poppler")
                or var_poppler.get()
            ),
        ).grid(row=3, column=1, padx=(4, 0))

        def salvar_e_fechar():
            nova_cfg = {
                "tesseract_cmd": var_tess.get().strip(),
                "poppler_path": var_poppler.get().strip(),
            }
            salvar_config(nova_cfg)
            aplicar_config_ambiente(nova_cfg)
            texto_status = self._texto_status_binarios()
            self.label_status_bin.config(text=texto_status)
            if texto_status:
                self.label_status_bin.pack(fill="x")
                self.btn_config.pack(side="right")
            else:
                self.label_status_bin.pack_forget()
                self.btn_config.pack_forget()
            messagebox.showinfo(
                "Configurações salvas",
                "Salvo. Se o programa já estava aberto, reinicie pra garantir que valha pra todos os arquivos.",
            )
            janela.destroy()

        ttk.Button(janela, text="Salvar", command=salvar_e_fechar).pack(pady=10)


if __name__ == "__main__":
    App().mainloop()
