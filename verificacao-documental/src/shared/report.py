"""
Estruturas de dados do relatório de uma OS. Compartilhadas pelos dois
fluxos (COBAN e Não COBAN): cada um preenche as mesmas duas dataclasses,
só muda quem gera as linhas (src.coban.orchestrator /
src.nao_coban.orchestrator).
"""

from dataclasses import dataclass, field
from typing import Optional


@dataclass
class DocumentReportRow:
    filename: str
    doc_type: str
    doc_label: str
    metodo_leitura: str
    confianca_ocr: float
    valido: bool
    pendencias: list[str] = field(default_factory=list)
    campos: dict = field(default_factory=dict)
    erro_extracao: Optional[str] = None
    precisa_verificacao_humana: bool = False  # extração fraca demais pra confiar


@dataclass
class OsReport:
    os_id: str
    rows: list[DocumentReportRow] = field(default_factory=list)
    pendencias_gerais: list[str] = field(default_factory=list)
    documentos_faltantes: list[str] = field(default_factory=list)
    tempo_processamento_segundos: float = 0.0

    @property
    def aprovado(self) -> bool:
        sem_pendencia_geral = not self.pendencias_gerais
        sem_doc_faltante = not self.documentos_faltantes
        todos_validos = all(r.valido for r in self.rows)
        return sem_pendencia_geral and sem_doc_faltante and todos_validos
