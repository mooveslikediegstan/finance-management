# -*- coding: utf-8 -*-
import pytest

from backend.app.models.purchase_order import PurchaseOrder

def test_pr_created_with_basic_fields():
    """
    Dado um numero de PR e um requisitante,
    quando eu crio um PurchaseOrder,
    entao esses valores devem estar acessiveis no objeto.
    """
    
    pedido = PurchaseOrder(
        pr_id="PR-001",
        requester="Joao Silva",
        total_amount=Decimal("1500.50"),
        planned_date=date(2024, 12, 31),
        project_id="PROJ-001",
        supplier_id="SUPP-001"
    )

    assert pedido.pr_id == "PR-001"
    assert pedido.requester == "Joao Silva"

def test_pr_id_empty_should_fail():
    """
    Dado um pr_id= vazio,
    quando tento criar um PurchaseOrder,
    então deve lançar um erro com mensagem clara.
    """
    with pytest.raises(ValueError, match="Número da PR não pode ser vazio"):
        PurchaseOrder(
            pr_id="",
            requester="João Silva",
            total_amount=Decimal("1500.50"),
            planned_date=date(2024, 12, 31),
            project_id="PROJ-001",
            supplier_id="SUPP-001"
        )

def test_pr_id_only_spaces_should_fail():
    """
    Espaços em branco são tão inválidos quanto string vazia.
    Você fazia Trim() no VBA — aqui testamos o mesmo comportamento.
    """
    with pytest.raises(ValueError, match="Número da PR não pode ser vazio"):
        PurchaseOrder(
            pr_id="   ",
            requester="João Silva",
            total_amount=Decimal("1500.50"),
            planned_date=date(2024, 12, 31),
            project_id="PROJ-001",
            supplier_id="SUPP-001"
        )

from decimal import Decimal

def test_pr_with_valid_amount():

    """
    Dado um valor positivo,
    quando crio um PurchaseOrder,
    entao total_brl deve estar acessivel.
    """
    pedido = PurchaseOrder(
        pr_id="PR-001",
        requester="Joao",
        total_amount=Decimal("1500.50"),
        planned_date=date(2024, 12, 31),
        project_id="PROJ-001",
        supplier_id="SUPP-001"
    )

    assert pedido.total_amount == Decimal("1500.50")

def test_zero_amount_should_fail():
    """
    Valor zero nao faz sentido num pedido de compra.
    """
    with pytest.raises(ValueError, match="Valor total deve ser maior que zero"):
        PurchaseOrder(
            pr_id="PR-001",
            requester="Joao",
            total_amount=Decimal("0"),
            planned_date=date(2024, 12, 31),
            project_id="PROJ-001",
            supplier_id="SUPP-001"
        )

def test_negative_amount_should_fail():
    with pytest.raises(ValueError, match="Valor total deve ser maior que zero"):
        PurchaseOrder(
            pr_id="PR-001",
            requester="Joao",
            total_amount=Decimal("-100"),
            planned_date=date(2024, 12, 31),
            project_id="PROJ-001",
            supplier_id="SUPP-001"
        )

from datetime import date

def test_pr_with_valid_planned_date():
    """
    A solicitacao de compra deve conter uma data planejada de entrega.
    """
    pl_date = date(2026, 12, 31)

    order = PurchaseOrder(
        pr_id="PR-001",
        requester="Joao",
        total_amount=Decimal("1500.50"),
        planned_date = pl_date,
        project_id="PROJ-001",
        supplier_id="SUPP-001"
    )

    assert order.planned_date == date(2026, 12, 31)

def test_past_planned_date_should_be_accepted():
    """
    Pedidos adicionados postumamente sao validos.
    """
    # Arrange
    past_date = date(2020, 1, 1)

    # Act
    order = PurchaseOrder(
        pr_id="PR-001",
        requester="Joao Silva",
        total_amount=Decimal("1500.50"),
        planned_date=past_date,
        project_id="PROJ-001",
        supplier_id="SUPP-001"
    )

    # Assert
    assert order.planned_date == date(2020, 1, 1)

def test_none_planned_date_should_fail():
    with pytest.raises(ValueError, match="Data planejada nao pode ser vazia"):
        PurchaseOrder(
            pr_id="PR-001",
            requester="Joao Silva",
            total_amount=Decimal("1500.50"),
            planned_date=None,
            project_id="PROJ-001",
            supplier_id="SUPP-001"
        )

def test_purchase_order_created_with_valid_project_id():
    # Arrange
    order = PurchaseOrder(
        pr_id="PR-001",
        requester="Joao Silva",
        total_amount=Decimal("1500.50"),
        planned_date=date(2025, 6, 1),
        project_id="PROJ-001",
        supplier_id="SUPP-001"
    )

    # Assert
    assert order.project_id == "PROJ-001"

def test_empty_project_id_should_fail():
    with pytest.raises(ValueError, match="Projeto nao pode ser vazio"):
        PurchaseOrder(
            pr_id="PR-001",
            requester="Joao Silva",
            total_amount=Decimal("1500.50"),
            planned_date=date(2025, 6, 1),
            project_id="",
            supplier_id="SUPP-001"
        )

def test_purchase_order_created_with_valid_supplier_id():
    # Arrange
    order = PurchaseOrder(
        pr_id="PR-001",
        requester="Joao Silva",
        total_amount=Decimal("1500.50"),
        planned_date=date(2025, 6, 1),
        project_id="PROJ-001",
        supplier_id="SUPP-001"
    )

    # Assert
    assert order.supplier_id == "SUPP-001"

def test_empty_supplier_id_should_fail():
    with pytest.raises(ValueError, match="Fornecedor nao pode ser vazio"):
        PurchaseOrder(
            pr_id="PR-001",
            requester="Joao Silva",
            total_amount=Decimal("1500.50"),
            planned_date=date(2025, 6, 1),
            project_id="PROJ-001",
            supplier_id=""
        )