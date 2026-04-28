# -*- coding: utf-8 -*-

from dataclasses import dataclass
from decimal import Decimal
from datetime import date
from typing import Optional

@dataclass
class PurchaseOrder:
    pr_id: str
    requester: str
    total_amount: Decimal = Decimal("0.00")
    planned_date: date = None
    creation_date: date = None
    project_id: str = None
    supplier_id: str = None
    forecast_scoped: bool = False
    description: str = ""
    cost_group: str = ""
    issue_id: str = ""
    invoice_delivery_status: str = ""
    po_number:  Optional[str]  = None
    po_date:    Optional[date] = None

    def __post_init__(self):
        self._validate()
        self._normalize()

    def _validate(self):
        if not self.pr_id.strip():
            raise ValueError("Número da PR não pode ser vazio")
        if not self.requester.strip():
            raise ValueError("Requisitante não pode ser vazio")
        if self.total_amount <= 0:
            raise ValueError("Valor total deve ser maior que zero")
        if self.planned_date is None:
            raise ValueError("Data planejada nao pode ser vazia")
        if self.creation_date is None:
            raise ValueError("Data de criacao nao pode ser vazia")
        if not isinstance(self.creation_date, date):
            raise ValueError("Data de criacao deve ser uma data valida")
        if not self.project_id.strip():
            raise ValueError("Projeto nao pode ser vazio")
        if self.supplier_id is not None and not self.supplier_id.strip():
            raise ValueError("Fornecedor nao pode ser vazio")
        if  self.forecast_scoped is None:
            raise ValueError("Previsto em forecast nao pode ser vazio")
        if not self.description.strip():
            raise ValueError("Descrição nao pode ser vazia")
        if not self.cost_group.strip():
            raise ValueError("Grupo de custo nao pode ser vazio")
        if not self.issue_id.strip():
            raise ValueError("ID do chamado nao pode ser vazio")
        if not self.invoice_delivery_status.strip():
            raise ValueError("Status de entrega da nota fiscal nao pode ser vazio")

    def _normalize(self):
        self.pr_id = self.pr_id.strip().upper()
        self.requester = self.requester.strip()
        self.project_id = self.project_id.strip()
        self.description = self.description.strip()
        self.cost_group = self.cost_group.strip()
        self.issue_id = self.issue_id.strip()
        self.invoice_delivery_status = self.invoice_delivery_status.strip()

    def receive_purchase_order(self, po_number: str, po_date: date) -> None:
        """
        Transicao explicita: PR vira PO.
        Equivalente ao momento que voce preenchia o num_po no formulario VBA.
        """
        if not po_number.strip():
            raise ValueError("Numero da PO nao pode ser vazio")
        if po_date is None:
            raise ValueError("Data da PO e obrigatoria ao receber uma PO")

        self.po_number = po_number.strip().upper()
        self.po_date   = po_date

    
    