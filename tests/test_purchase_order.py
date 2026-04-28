# -*- coding: utf-8 -*-
from datetime import date
from decimal import Decimal
import pytest

from backend.app.models.purchase_order import PurchaseOrder

@pytest.fixture
def valid_order():
    return PurchaseOrder(
        pr_id="PR-001",
        requester="Joao Silva",
        total_amount=Decimal("99999.99"),
        planned_date=date(2026, 4, 26),
        creation_date=date(2026, 4, 26),
        project_id="PROJ-001",
        supplier_id="SUPP-001",
        forecast_scoped=True,
        description="Testando criacao de pedido de compra com descricao valida",
        cost_group="Consumables",
        issue_id="ISSUE-001",
        invoice_delivery_status="Delivered"
    )

# --- helpers ---

def make_order(valid_order, **overrides):
    """
    Cria um novo PurchaseOrder a partir da fixture, sobrescrevendo campos especificos.
    Evita repetir **{**valid_order.__dict__, ...} em cada teste.
    """
    return PurchaseOrder(**{**valid_order.__dict__, **overrides})

# --- criacao basica ---

def test_purchase_order_created_with_basic_fields(valid_order):
    assert valid_order.pr_id == "PR-001"
    assert valid_order.requester == "Joao Silva"

def test_purchase_order_with_valid_amount(valid_order):
    assert valid_order.total_amount == Decimal("99999.99")

def test_purchase_order_with_valid_planned_date(valid_order):
    assert valid_order.planned_date == date(2026, 4, 26)

def test_past_planned_date_should_be_accepted(valid_order):
    order = make_order(valid_order, planned_date=date(2020, 1, 1))
    assert order.planned_date == date(2020, 1, 1)

def test_purchase_order_created_with_valid_project_id(valid_order):
    assert valid_order.project_id == "PROJ-001"

def test_purchase_order_created_with_valid_supplier_id(valid_order):
    assert valid_order.supplier_id == "SUPP-001"

def test_purchase_order_with_valid_creation_date(valid_order):
    assert valid_order.creation_date == date(2026, 4, 26)

def test_purchase_order_with_valid_forecast_scoped(valid_order):
    assert valid_order.forecast_scoped == True

def test_purchase_order_with_valid_description(valid_order):
    assert valid_order.description == "Testando criacao de pedido de compra com descricao valida"

def test_purchase_order_with_valid_cost_group(valid_order):
    assert valid_order.cost_group == "Consumables"

def test_purchase_order_with_valid_issue_id(valid_order):
    assert valid_order.issue_id == "ISSUE-001"

def test_purchase_order_with_valid_invoice_delivery_status(valid_order):
    assert valid_order.invoice_delivery_status == "Delivered"

# --- validacao de campos ---

def test_pr_id_empty_should_fail(valid_order):
    with pytest.raises(ValueError, match="Número da PR não pode ser vazio"):
        make_order(valid_order, pr_id="")

def test_pr_id_only_spaces_should_fail(valid_order):
    with pytest.raises(ValueError, match="Número da PR não pode ser vazio"):
        make_order(valid_order, pr_id="   ")

def test_zero_amount_should_fail(valid_order):
    with pytest.raises(ValueError, match="Valor total deve ser maior que zero"):
        make_order(valid_order, total_amount=Decimal("0"))

def test_negative_amount_should_fail(valid_order):
    with pytest.raises(ValueError, match="Valor total deve ser maior que zero"):
        make_order(valid_order, total_amount=Decimal("-100"))

def test_none_planned_date_should_fail(valid_order):
    with pytest.raises(ValueError, match="Data planejada nao pode ser vazia"):
        make_order(valid_order, planned_date=None)

def test_empty_project_id_should_fail(valid_order):
    with pytest.raises(ValueError, match="Projeto nao pode ser vazio"):
        make_order(valid_order, project_id="")

def test_empty_supplier_id_should_fail(valid_order):
    with pytest.raises(ValueError, match="Fornecedor nao pode ser vazio"):
        make_order(valid_order, supplier_id="")

def test_creation_date_none_should_fail(valid_order):
    with pytest.raises(ValueError, match="Data de criacao nao pode ser vazia"):
        make_order(valid_order, creation_date=None)

def test_creation_date_invalid_type_should_fail(valid_order):
    with pytest.raises(ValueError, match="Data de criacao deve ser uma data valida"):
        make_order(valid_order, creation_date="2026-04-26")

def test_forecast_scoped_none_should_fail(valid_order):
    with pytest.raises(ValueError, match="Previsto em forecast nao pode ser vazio"):
        make_order(valid_order, forecast_scoped=None)

def test_empty_description_should_fail(valid_order):
    with pytest.raises(ValueError, match="Descrição nao pode ser vazia"):
        make_order(valid_order, description="")

def test_description_with_only_spaces_should_fail(valid_order):
    with pytest.raises(ValueError, match="Descrição nao pode ser vazia"):
        make_order(valid_order, description="     ")

def test_cost_group_empty_should_fail(valid_order):
    with pytest.raises(ValueError, match="Grupo de custo nao pode ser vazio"):
        make_order(valid_order, cost_group="")

def test_cost_group_with_only_spaces_should_fail(valid_order):
    with pytest.raises(ValueError, match="Grupo de custo nao pode ser vazio"):
        make_order(valid_order, cost_group="     ")

def test_issue_id_empty_should_fail(valid_order):
    with pytest.raises(ValueError, match="ID do chamado nao pode ser vazio"):
        make_order(valid_order, issue_id="")

def test_issue_id_with_only_spaces_should_fail(valid_order):
    with pytest.raises(ValueError, match="ID do chamado nao pode ser vazio"):
        make_order(valid_order, issue_id="     ")

def test_invoice_delivery_status_empty_should_fail(valid_order):
    with pytest.raises(ValueError, match="Status de entrega da nota fiscal nao pode ser vazio"):
        make_order(valid_order, invoice_delivery_status="")

def test_invoice_delivery_status_with_only_spaces_should_fail(valid_order):
    with pytest.raises(ValueError, match="Status de entrega da nota fiscal nao pode ser vazio"):
        make_order(valid_order, invoice_delivery_status="     ")

# --- transicao PR -> PO ---

def test_receive_po_updates_fields(valid_order):
    valid_order.receive_purchase_order(po_number="PO-999", po_date=date(2025, 7, 1))
    assert valid_order.po_number == "PO-999"
    assert valid_order.po_date   == date(2025, 7, 1)

def test_receive_po_without_date_should_fail(valid_order):
    with pytest.raises(ValueError, match="Data da PO e obrigatoria"):
        valid_order.receive_purchase_order(po_number="PO-999", po_date=None)

def test_receive_po_with_empty_number_should_fail(valid_order):
    with pytest.raises(ValueError, match="Numero da PO nao pode ser vazio"):
        valid_order.receive_purchase_order(po_number="", po_date=date(2025, 7, 1))

def test_receive_po_with_spaces_should_fail(valid_order):
    with pytest.raises(ValueError, match="Numero da PO nao pode ser vazio"):
        valid_order.receive_purchase_order(po_number="        ", po_date=date(2025, 7, 1))