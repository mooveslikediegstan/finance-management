# -*- coding: utf-8 -*-
from decimal import Decimal
import pytest

from backend.app.models.supplier import Supplier

@pytest.fixture
def valid_supplier():
    return Supplier(
        supplier_id="SUPP-001",
        name="Empresa Exemplo Ltda",
        fantasy_name="Exemplo",
        cnpj="07.526.557/0001-00",
        city="Sao Paulo",
        uf="SP",
        contact="joao@exemplo.com.br"
    )

def make_supplier(valid_supplier, **overrides):
    return Supplier(**{**valid_supplier.__dict__, **overrides})

# --- criacao basica ---

def test_supplier_created_with_valid_fields(valid_supplier):
    assert valid_supplier.supplier_id == "SUPP-001"
    assert valid_supplier.name == "Empresa Exemplo Ltda"
    assert valid_supplier.fantasy_name == "Exemplo"
    assert valid_supplier.cnpj == "07.526.557/0001-00"
    assert valid_supplier.city == "Sao Paulo"
    assert valid_supplier.uf == "SP"
    assert valid_supplier.contact == "joao@exemplo.com.br"

# --- cnpj ---

def test_cnpj_accepted_with_formatting(valid_supplier):
    assert valid_supplier.cnpj == "07.526.557/0001-00"

def test_cnpj_accepted_without_formatting(valid_supplier):
    supplier = make_supplier(valid_supplier, cnpj="07526557000100")
    assert supplier.cnpj == "07.526.557/0001-00"

def test_cnpj_invalid_digits_should_fail(valid_supplier):
    with pytest.raises(ValueError, match="CNPJ invalido"):
        make_supplier(valid_supplier, cnpj="11.111.111/1111-11")

def test_cnpj_wrong_format_should_fail(valid_supplier):
    with pytest.raises(ValueError, match="CNPJ invalido"):
        make_supplier(valid_supplier, cnpj="123")

def test_cnpj_empty_should_fail(valid_supplier):
    with pytest.raises(ValueError, match="CNPJ invalido"):
        make_supplier(valid_supplier, cnpj="")

# --- uf ---

def test_valid_uf_accepted(valid_supplier):
    assert valid_supplier.uf == "SP"

def test_exterior_uf_accepted(valid_supplier):
    supplier = make_supplier(valid_supplier, uf="EX")
    assert supplier.uf == "EX"

def test_invalid_uf_should_fail(valid_supplier):
    with pytest.raises(ValueError, match="UF invalida"):
        make_supplier(valid_supplier, uf="XX")

def test_lowercase_uf_should_be_normalized(valid_supplier):
    supplier = make_supplier(valid_supplier, uf="sp")
    assert supplier.uf == "SP"

# --- contact opcional ---

def test_supplier_created_without_contact(valid_supplier):
    supplier = make_supplier(valid_supplier, contact=None)
    assert supplier.contact is None

# --- campos obrigatorios ---

def test_empty_supplier_id_should_fail(valid_supplier):
    with pytest.raises(ValueError, match="ID do fornecedor nao pode ser vazio"):
        make_supplier(valid_supplier, supplier_id="")

def test_empty_name_should_fail(valid_supplier):
    with pytest.raises(ValueError, match="Razao social nao pode ser vazia"):
        make_supplier(valid_supplier, name="")

def test_empty_fantasy_name_should_fail(valid_supplier):
    with pytest.raises(ValueError, match="Nome fantasia nao pode ser vazio"):
        make_supplier(valid_supplier, fantasy_name="")

def test_empty_city_should_fail(valid_supplier):
    with pytest.raises(ValueError, match="Cidade nao pode ser vazia"):
        make_supplier(valid_supplier, city="")