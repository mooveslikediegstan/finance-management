
from app.backend.models.purchase_order import PedidoCompra  # ainda não existe — tudo bem

def test_pedido_criado_com_campos_basicos():
    """
    Dado um número de PR e um requisitante,
    quando eu crio um PedidoCompra,
    então esses valores devem estar acessíveis no objeto.
    """
    pedido = PedidoCompra(num_pr="PR-001", requisitante="João Silva")

    assert pedido.num_pr == "PR-001"
    assert pedido.requisitante == "João Silva"