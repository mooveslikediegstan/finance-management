# -*- coding: utf-8 -*-

from dataclasses import dataclass
from decimal import Decimal
from datetime import date

@dataclass
class PurchaseOrder:
    pr_id: str
    requester: str
    total_amount: Decimal = Decimal("0.00")
    planned_date: date = None
    project_id: str = None
    supplier_id: str = None

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
        if not self.project_id.strip():
            raise ValueError("Projeto nao pode ser vazio")
        if self.supplier_id is not None and not self.supplier_id.strip():
            raise ValueError("Fornecedor nao pode ser vazio")
        
        
    def _normalize(self):
        self.pr_id = self.pr_id.strip().upper()
        self.requester = self.requester.strip()
        self.project_id = self.project_id.strip()