# -*- coding: utf-8 -*-
from dataclasses import dataclass
from typing import Optional

VALID_UF = {
    "AC", "AL", "AP", "AM", "BA", "CE", "DF", "ES", "GO",
    "MA", "MT", "MS", "MG", "PA", "PB", "PR", "PE", "PI",
    "RJ", "RN", "RS", "RO", "RR", "SC", "SP", "SE", "TO",
    "EX"
}

@dataclass
class Supplier:
    supplier_id:  str
    name:         str
    fantasy_name: str
    cnpj:         str
    city:         str
    uf:           str
    contact:      Optional[str] = None

    def __post_init__(self):
        self._normalize()
        self._validate()

    def _normalize(self):
        self.supplier_id  = self.supplier_id.strip().upper()
        self.name         = self.name.strip()
        self.fantasy_name = self.fantasy_name.strip()
        self.city         = self.city.strip()
        self.uf           = self.uf.strip().upper()
        self.cnpj         = self._format_cnpj(self._strip_cnpj(self.cnpj))
        if self.contact:
            self.contact  = self.contact.strip()

    def _validate(self):
        if not self.supplier_id:
            raise ValueError("ID do fornecedor nao pode ser vazio")
        if not self.name:
            raise ValueError("Razao social nao pode ser vazia")
        if not self.fantasy_name:
            raise ValueError("Nome fantasia nao pode ser vazio")
        if not self.city:
            raise ValueError("Cidade nao pode ser vazia")
        if self.uf not in VALID_UF:
            raise ValueError("UF invalida")
        if not self._is_valid_cnpj(self._strip_cnpj(self.cnpj)):
            raise ValueError("CNPJ invalido")

    def _strip_cnpj(self, cnpj: str) -> str:
        """Remove formatacao — equivalente ao seu NormalizeCNPJNumber() do VBA."""
        return cnpj.strip().replace(".", "").replace("/", "").replace("-", "")

    def _format_cnpj(self, cnpj: str) -> str:
        """Aplica mascara XX.XXX.XXX/XXXX-XX."""
        if len(cnpj) != 14:
            return cnpj  # deixa passar para o _validate acusar
        return f"{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:12]}-{cnpj[12:]}"

    def _is_valid_cnpj(self, cnpj: str) -> bool:
        if len(cnpj) != 14 or not cnpj.isdigit():
            return False
        if len(set(cnpj)) == 1:
            return False

        def _calc_digit(cnpj: str, weights: list[int]) -> int:
            total     = sum(int(d) * w for d, w in zip(cnpj, weights))
            remainder = total % 11
            return 0 if remainder < 2 else 11 - remainder

        weights_1 = [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]
        weights_2 = [6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]
        first     = _calc_digit(cnpj[:12], weights_1)
        second    = _calc_digit(cnpj[:13], weights_2)

        return cnpj[12] == str(first) and cnpj[13] == str(second)