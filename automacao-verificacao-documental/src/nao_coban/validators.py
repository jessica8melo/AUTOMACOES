"""
Regras de validação específicas do fluxo Não COBAN (Cadastro/Atualização
comum de fornecedor, PRO415-004 + TAB415-004). Cada validador recebe o
texto já extraído (src.shared.extractor) e devolve um DocumentValidation
com:
  - valido: passou nas regras mínimas?
  - campos: dados estruturados extraídos (CNPJ, situação, datas...)
  - pendencias: lista de problemas encontrados (para o relatório e e-mail)

CNPJ, Contrato Social, Simples Nacional e Procuração são validados pelas
mesmas funções do fluxo COBAN; ver src.shared.common_validators.
"""

import re
from datetime import datetime
from typing import Optional

from src.shared.common_validators import (
    ASSINANTE_D4SIGN_RE,
    validate_cnpj_doc,
    validate_contrato_social,
    validate_procuracao,
    validate_simples_nacional,
)
from src.shared.validation import DATE_RE, DocumentValidation, find_cnpj

RAZAO_SOCIAL_FQ064_RE = re.compile(r"Raz[aã]o\s+Social:\s*([^\n\r]+)", re.I)

ASSINATURA_REPRESENTANTE_RE = re.compile(
    # [ \t]* (não \s*) antes do nome: se o campo "Nome:" estiver mesmo em
    # branco no formulário, isso evita "pular a linha" e capturar por engano
    # o rótulo do campo seguinte ("Cargo: Assinatura:") como se fosse nome.
    r"Assinatura\s+do\s+representante\s+legal.{0,400}?Nome:[ \t]*([^\n\r]+)",
    re.I | re.S,
)

OPTANTE_SIMPLES_FQ064_RE = re.compile(
    r"Optante\s+pelo\s+Simples\s+Nacional\?\s*\(\s*([xX ]?)\s*\)\s*SIM\s*\(\s*([xX ]?)\s*\)\s*N[AÃ]O",
    re.I,
)

# O FQ415-064 às vezes vem assinado com certificado digital nativo do PDF
# (ICP-Brasil) em vez de D4Sign; carimbo "Assinado de forma digital por NOME".
ASSINADO_DIGITALMENTE_POR_RE = re.compile(
    r"[Aa]ssinado\s+de\s+forma\s+digital\s+por\s*:?\s*([^\n\r]{2,80})", re.I
)

# Terceiro formato de assinatura visto na prática: plataformas tipo
# NetLex/Unico anexam uma página de autenticação separada (não é D4Sign nem
# certificado ICP-Brasil embutido) com "1º ASSINANTE - Própria NOME" +
# CPF mascarado + "Assinado em:". Nesses casos o campo "Nome:" do próprio
# formulário costuma ficar em branco (quem assina não digita ali, a
# plataforma é que registra a identidade), então esse é o único jeito de
# achar o nome de quem realmente assinou.
ASSINANTE_NETLEX_RE = re.compile(
    r"\d+º?\s*ASSINANTE\s*-\s*Pr[oó]pria\s*\n?\s*([A-ZÀ-Ý][A-ZÀ-Ý\s]{2,60}?)\n"
)

CNPJ_CAMPO_FQ064_RE = re.compile(r"CNPJ\s*N[ºo°]?\s*:?\s*([\d./\-]{14,20})", re.I)
CPF_CAMPO_FQ064_RE = re.compile(r"CPF\s*N[ºo°]?\s*:?\s*([\d.\-]{11,14})", re.I)


def detectar_tipo_pessoa_fq064(text: str) -> Optional[str]:
    """O FQ415-064 tem uma seção 'exclusiva PJ' e uma 'exclusiva PF'. Detecta
    qual delas veio preenchida de fato (tem número de CNPJ/CPF), pra decidir
    automaticamente qual conjunto de insumos da TAB415-004 exigir. O
    analista não precisa selecionar isso manualmente."""
    m_cnpj = CNPJ_CAMPO_FQ064_RE.search(text or "")
    if m_cnpj and re.sub(r"\D", "", m_cnpj.group(1)):
        return "PJ"
    m_cpf = CPF_CAMPO_FQ064_RE.search(text or "")
    if m_cpf and re.sub(r"\D", "", m_cpf.group(1)):
        return "PF"
    return None


def validate_formulario_cadastro_fornecedor(
    text: str, filename: str
) -> DocumentValidation:
    """FQ415-064  usado no fluxo Não COBAN no lugar do Formulário de
    Credenciamento (que é específico do COBAN). Estrutura diferente (não
    tem 'Nome Sócio N:'), então o nome de quem assina vem do bloco
    'Assinatura do representante legal (...) / Nome: XXXX'.

    IMPORTANTE: o próprio formulário permite que quem assine seja "o gestor
    da área solicitante" em vez do representante legal do fornecedor, nos
    casos de Cadastro de Não Fornecedor previstos na TAB415-004. Por isso o
    orchestrator só faz a checagem cruzada de sócio x Contrato Social no
    tipo FORNECEDOR_CONTRATADO / OBRIGACOES_JUDICIAIS (PJ), não nos outros.

    campos["nomes_socios"] usa o mesmo formato do Formulário de
    Credenciamento (nome único, sem separador ';') pra reaproveitar a
    lógica de checagem cruzada do orchestrator.
    """
    pendencias = []
    cnpj = find_cnpj(text)

    tipo_pessoa = detectar_tipo_pessoa_fq064(text)

    razao_social_match = RAZAO_SOCIAL_FQ064_RE.search(text)
    razao_social = razao_social_match.group(1).strip() if razao_social_match else ""

    if tipo_pessoa == "PJ" and not cnpj:
        pendencias.append("CNPJ não localizado no Formulário FQ415-064.")

    assinantes_certificado = [
        m.group(1).strip() for m in ASSINADO_DIGITALMENTE_POR_RE.finditer(text)
    ]
    assinante_netlex_match = ASSINANTE_NETLEX_RE.search(text)
    assinado_d4sign = bool(re.search(r"d4sign", text, re.I))
    assinado_certificado = bool(assinantes_certificado)
    assinado_netlex = bool(assinante_netlex_match)
    if not assinado_d4sign and not assinado_certificado and not assinado_netlex:
        pendencias.append(
            "Não foi possível confirmar assinatura (D4Sign ou certificado digital) no documento."
        )

    assinatura_match = ASSINATURA_REPRESENTANTE_RE.search(text)
    nome_assinante = assinatura_match.group(1).strip() if assinatura_match else ""
    if not nome_assinante and assinante_netlex_match:
        nome_assinante = assinante_netlex_match.group(1).strip()
    if not nome_assinante:
        pendencias.append(
            "Nome de quem assina não localizado no FQ415-064. Verificar manualmente."
        )

    optante_match = OPTANTE_SIMPLES_FQ064_RE.search(text)
    optante_simples = None
    if optante_match:
        marcado_sim, marcado_nao = (
            optante_match.group(1).strip(),
            optante_match.group(2).strip(),
        )
        if marcado_sim.lower() == "x":
            optante_simples = True
        elif marcado_nao.lower() == "x":
            optante_simples = False

    nomes_assinantes_d4sign = [
        m.group(1).strip()
        for m in ASSINANTE_D4SIGN_RE.finditer(text)
        if not m.group(2).lower().endswith("@bbts.com.br")
    ] + assinantes_certificado
    if assinante_netlex_match:
        nomes_assinantes_d4sign.append(assinante_netlex_match.group(1).strip())

    return DocumentValidation(
        doc_type="FORMULARIO_CADASTRO_FORNECEDOR",
        filename=filename,
        valido=not pendencias,
        campos={
            "cnpj": cnpj or "",
            "nome_empresarial": razao_social,
            "tipo_pessoa": tipo_pessoa or "NAO_IDENTIFICADO",
            "nomes_socios": nome_assinante,
            "optante_simples": str(optante_simples)
            if optante_simples is not None
            else "NAO_IDENTIFICADO",
            "nomes_assinantes_d4sign": ";".join(nomes_assinantes_d4sign),
        },
        pendencias=pendencias,
    )


def _extrair_data_validade(texto_proximo: str) -> str:
    """CNH/RG têm layout de cartão com colunas; o OCR às vezes emenda a
    data de emissão e a de validade uma do lado da outra ('4b VALIDADE
    ACC 02/08/2023 30/04/2031', onde a primeira é a emissão que vazou da
    coluna ao lado). Quando isso acontece, pega a data mais distante no
    futuro: validade nunca vem antes da emissão."""
    datas_str = re.findall(r"\d{2}/\d{2}/\d{4}", texto_proximo)
    datas_validas = []
    for d in datas_str:
        try:
            dia, mes, ano = (int(p) for p in d.split("/"))
            datas_validas.append((datetime(ano, mes, dia), d))
        except ValueError:
            continue
    if not datas_validas:
        return ""
    return max(datas_validas, key=lambda t: t[0])[1]


CPF_RE = re.compile(r"\d{3}\.?\d{3}\.?\d{3}-?\d{2}")


def validate_documento_identificacao(text: str, filename: str) -> DocumentValidation:
    """TAB415-004]: precisa ter foto e estar dentro da validade.
    CPF é extraído daqui. É dispensada a cópia do
    CPF quando o número já consta neste documento.

    O NOME não é extraído aqui: o layout de RG/CNH varia entre
    modelos e o OCR embaralha a ordem dos campos do cartão. O nome já é
    conhecido de forma confiável pelo campo "Nome:" do FQ415-064; o
    orchestrator só confere se esse nome aparece neste documento.
    """
    pendencias = []
    validade = ""
    if re.search(r"validade", text or "", re.I):
        validade = _extrair_data_validade(text or "")
    if validade:
        try:
            dia, mes, ano = (int(p) for p in validade.split("/"))
            if datetime(ano, mes, dia) < datetime.now():
                pendencias.append(
                    f"Documento de identificação com validade vencida ({validade})."
                )
        except ValueError:
            pendencias.append(
                "Data de validade do documento de identificação não pôde ser interpretada. Verificar manualmente."
            )

    cpf_match = CPF_RE.search(text or "")
    cpf = re.sub(r"\D", "", cpf_match.group()) if cpf_match else ""
    if not cpf:
        pendencias.append(
            "CPF não localizado no documento de identificação. Verificar manualmente."
        )

    return DocumentValidation(
        doc_type="DOCUMENTO_IDENTIFICACAO",
        filename=filename,
        valido=not pendencias,
        campos={"validade": validade, "cpf": cpf},
        pendencias=pendencias,
    )


def validate_cpf_doc(text: str, filename: str) -> DocumentValidation:
    pendencias = []
    cpf_match = CPF_RE.search(text or "")
    cpf = re.sub(r"\D", "", cpf_match.group()) if cpf_match else ""
    if not cpf:
        pendencias.append("CPF não localizado no documento.")
    return DocumentValidation(
        doc_type="CPF",
        filename=filename,
        valido=not pendencias,
        campos={"cpf": cpf},
        pendencias=pendencias,
    )


def validate_comprovante_residencia(text: str, filename: str) -> DocumentValidation:
    """TAB415-004: precisa ter sido emitido há menos de 90 dias
    corridos."""
    pendencias = []
    data_match = DATE_RE.search(text or "")
    data_emissao = data_match.group() if data_match else ""
    if not data_emissao:
        pendencias.append(
            "Data de emissão não localizada no comprovante de residência. Verificar manualmente."
        )
    else:
        try:
            dia, mes, ano = (int(p) for p in data_emissao.split("/"))
            dias_corridos = (datetime.now() - datetime(ano, mes, dia)).days
            if dias_corridos > 90:
                pendencias.append(
                    f"Comprovante de residência emitido há {dias_corridos} dias (> 90 dias). Solicitar comprovante mais recente."
                )
        except ValueError:
            pendencias.append(
                "Data de emissão do comprovante de residência não pôde ser interpretada. Verificar manualmente."
            )
    return DocumentValidation(
        doc_type="COMPROVANTE_RESIDENCIA",
        filename=filename,
        valido=not pendencias,
        campos={"data_emissao": data_emissao},
        pendencias=pendencias,
    )


def validate_nota_fiscal_fatura_bbts(text: str, filename: str) -> DocumentValidation:
    """TAB415-004: NF/fatura emitida à BBTS. Checagem só de
    presença/indício textual."""
    pendencias = []
    if not re.search(r"bb\s*tecnologia|bbts", text or "", re.I):
        pendencias.append(
            "Documento não menciona explicitamente a BB Tecnologia e Serviços. Verificar manualmente se é a NF/fatura correta."
        )
    return DocumentValidation(
        doc_type="NOTA_FISCAL_FATURA_BBTS",
        filename=filename,
        valido=not pendencias,
        campos={},
        pendencias=pendencias,
    )


def validate_fora_de_escopo(_text: str, filename: str) -> DocumentValidation:
    """Documento reconhecido (pra não colidir por engano com outro tipo),
    mas fora do escopo desta validação. Não é insumo da TAB415-004 (ex:
    Relatório de Enquadramento Fisco-Tributário, Especificações Técnicas,
    Projeto Básico). Sempre válido, sem pendências."""
    return DocumentValidation(
        doc_type="DOCUMENTO_FORA_DE_ESCOPO",
        filename=filename,
        valido=True,
        campos={},
        pendencias=[],
    )


VALIDATORS_NAO_COBAN = {
    "CNPJ": validate_cnpj_doc,
    "CONTRATO_SOCIAL": validate_contrato_social,
    "PROCURACAO": validate_procuracao,
    "SIMPLES_NACIONAL": validate_simples_nacional,
    "FORMULARIO_CADASTRO_FORNECEDOR": validate_formulario_cadastro_fornecedor,
    "DOCUMENTO_IDENTIFICACAO": validate_documento_identificacao,
    "CPF": validate_cpf_doc,
    "COMPROVANTE_RESIDENCIA": validate_comprovante_residencia,
    "NOTA_FISCAL_FATURA_BBTS": validate_nota_fiscal_fatura_bbts,
    "DOCUMENTO_FORA_DE_ESCOPO": validate_fora_de_escopo,
}


def validate_document_nao_coban(
    doc_type: str, text: str, filename: str
) -> DocumentValidation:
    validator = VALIDATORS_NAO_COBAN.get(doc_type)
    if not validator:
        return DocumentValidation(
            doc_type=doc_type or "DESCONHECIDO",
            filename=filename,
            valido=False,
            pendencias=[
                "Tipo de documento não reconhecido. Revisão manual necessária."
            ],
        )
    return validator(text, filename)
