# Finance Management — Controle de Pedidos de Compra
> Arquivo de contexto para agentes de IA. Gerado em 26/04/2026.
> Itens marcados com `[ TODO ]` ainda não foram definidos e devem ser preenchidos nas próximas sessões.

---

## 1. Visão Geral

Sistema administrativo e financeiro de **Controle de Pedidos de Compra**, desenvolvido para substituir uma solução legada em Excel/VBA (`.xlsm`). O sistema replica e expande as funcionalidades do sistema VBA original, que incluía CRUDs de Purchase Orders, Suppliers, Projects e Customers, com formulários de busca e registro.

O desenvolvedor é iniciante em Python, com sólida experiência em VBA/Excel e SQL. A metodologia de desenvolvimento adotada é **TDD (Test-Driven Development)** com ciclo Red → Green → Refactor.

---

## 2. Stack Tecnológico

| Camada | Tecnologia | Observação |
|---|---|---|
| Backend / API | FastAPI | Framework principal |
| Servidor ASGI | Uvicorn | Roda o FastAPI |
| Frontend | Streamlit | Fase inicial; migrar para React futuramente |
| Banco de dados | PostgreSQL 16 | Produção |
| ORM | SQLAlchemy 2.x | Com padrão Repository |
| Migrações | Alembic | Controle de versão do schema |
| Validação | Pydantic v2 | Schemas de entrada/saída da API |
| Testes | pytest | TDD — testes escritos antes do código |
| Containerização | Docker + Docker Compose | 3 containers: frontend, backend, banco |
| Linguagem | Python 3.13 | |
| Controle de versão | Git + GitHub | Repositório: `finance-management` |
| Sistema operacional (dev) | Windows 11 | Terminal: PowerShell no VS Code |

### Comando padrão para rodar testes (Windows)
```powershell
python -m pytest tests/ -v
```

---

## 3. Arquitetura

O sistema segue arquitetura em três camadas com separação explícita de responsabilidades:

```
Frontend (Streamlit)
      ↓ HTTP/JSON
Backend (FastAPI)
    ├── Routers      → endpoints HTTP (equivalente aos botões dos UserForms VBA)
    ├── Schemas      → validação Pydantic de entrada/saída
    ├── Services     → lógica de negócio
    └── Repository   → acesso ao banco via SQLAlchemy
      ↓ SQL
Banco de Dados (PostgreSQL)
```

### Padrão de camadas adotado
- **Models** (`app/backend/models/`) — objetos de negócio puros (`@dataclass`), sem conhecimento do banco
- **ORM Models** (`app/backend/database/orm_models.py`) — mapeamento SQLAlchemy das tabelas [ TODO ]
- **Repository** (`app/backend/database/repository.py`) — operações CRUD [ TODO ]
- **Services** (`app/backend/services/`) — lógica de negócio e orquestração [ TODO ]
- **Routers** (`app/backend/routers/`) — rotas FastAPI [ TODO ]
- **Schemas** (`app/backend/schemas/`) — Pydantic [ TODO ]

---

## 4. Estrutura de Diretórios

```
finance-management/
├── docker-compose.yml
├── docker-compose.prod.yml          [ TODO ]
├── PROJECT_CONTEXT.md               ← este arquivo
│
├── backend/
│   ├── Dockerfile                   [ TODO ]
│   ├── requirements.txt             [ TODO ]
│   └── app/
│       ├── main.py                  [ TODO ]
│       ├── models/
│       │   ├── purchase_order.py    ← em desenvolvimento (ver seção 6)
│       │   ├── supplier.py          [ TODO ]
│       │   ├── project.py           [ TODO ]
│       │   └── customer.py          [ TODO ]
│       ├── database/
│       │   ├── session.py           [ TODO ]
│       │   ├── orm_models.py        [ TODO ]
│       │   └── repository.py        [ TODO ]
│       ├── services/                [ TODO ]
│       ├── routers/                 [ TODO ]
│       └── schemas/                 [ TODO ]
│
├── frontend/
│   ├── Dockerfile                   [ TODO ]
│   ├── requirements.txt             [ TODO ]
│   └── app/
│       └── main.py                  [ TODO ]
│
└── tests/
    └── test_purchase_order.py       ← em desenvolvimento (ver seção 7)
```

---

## 5. Variáveis de Ambiente

### Backend
| Variável | Exemplo | Descrição |
|---|---|---|
| `DATABASE_URL` | `postgresql://admin:senha@banco:5432/compras_db` | Connection string do PostgreSQL |
| `SECRET_KEY` | `sua_chave_secreta_jwt` | Chave para assinatura JWT [ TODO ] |

### Banco de Dados (container)
| Variável | Exemplo | Descrição |
|---|---|---|
| `POSTGRES_DB` | `compras_db` | Nome do banco |
| `POSTGRES_USER` | `admin` | Usuário |
| `POSTGRES_PASSWORD` | `senha_segura` | Senha |

### Frontend [ TODO ]
| Variável | Exemplo | Descrição |
|---|---|---|
| `API_URL` | `http://backend:8000` | URL do backend dentro da rede Docker |

---

## 6. Models Definidos

### `PurchaseOrder` — `app/backend/models/purchase_order.py`
Objeto de negócio puro (`@dataclass`). Não conhece o banco de dados.

| Campo | Tipo | Obrigatório | Regras |
|---|---|---|---|
| `pr_id` | `str` | Sim | Não pode ser vazio; normalizado para maiúsculas e sem espaços |
| `requester` | `str` | Sim | Não pode ser vazio; espaços removidos |
| `total_amount` | `Decimal` | Sim | Deve ser maior que zero |
| `planned_date` | `date` | Sim | Qualquer data válida aceita (inclusive passado — pedidos postumamente) |
| `project_id` | `str` | Sim | Não pode ser vazio; sem validação de existência no banco |
| `supplier_id` | `str` | Sim | Não pode ser vazio; sem validação de existência no banco |

**Campos ainda a definir para esta classe:**
- `[ TODO ]` `input_date` — data de input no sistema
- `[ TODO ]` `description` — descrição do serviço
- `[ TODO ]` `num_po` — número da Purchase Order
- `[ TODO ]` `po_date` — data de recebimento da PO (só válida se `num_po` preenchido)
- `[ TODO ]` `num_nf` — número da Nota Fiscal
- `[ TODO ]` `balance_brl` — saldo da SC
- `[ TODO ]` `cost_allocation` — alocação de recurso/grupo de custo
- `[ TODO ]` `in_forecast` — previsto em forecast (booleano)

### `Supplier` — [ TODO ]
Baseado em `Supplier.cls` do VBA. Campos identificados no código original:
`supplier_id`, `name`, `fantasy_name`, `cnpj`, `city`, `uf`, `contact`

### `Project` — [ TODO ]
Baseado em `Project.cls` do VBA. Campos identificados:
`project_id`, `customer_id`

### `Customer` — [ TODO ]
Não estava nos arquivos enviados. Campos inferidos pelo código:
`customer_id`, `customer_name`, `customer_short_name`

### `PurchaseOrderPlanning` — [ TODO ]
Baseado em `PurchOrderPlanning.cls` do VBA. Campos identificados:
`purchase_requisition_id`, `planned_date`, `amount`

---

## 7. Testes

### Convenções adotadas
- Padrão de nomenclatura: `test_context_with_situation` / `test_situation_should_fail`
- Estrutura interna: **AAA — Arrange, Act, Assert**
- Uma asserção por teste
- Agrupamento por contexto usando classes (ex: `TestPurchaseOrderValidation`)
- Arquivo de testes tende a ser maior que o arquivo da classe — isso é esperado

### `tests/test_purchase_order.py` — status atual
| Teste | Status |
|---|---|
| `test_purchase_order_created_with_valid_fields` | Passando |
| `test_empty_pr_id_should_fail` | Passando |
| `test_whitespace_pr_id_should_fail` | Passando |
| `test_pr_with_valid_amount` | Passando |
| `test_zero_amount_should_fail` | Passando |
| `test_negative_amount_should_fail` | Passando |
| `test_purchase_order_created_with_valid_planned_date` | Passando |
| `test_past_planned_date_should_be_accepted` | Passando |
| `test_none_planned_date_should_fail` | Passando |
| `test_purchase_order_created_with_valid_project_id` | Passando |
| `test_empty_project_id_should_fail` | Passando |
| `test_purchase_order_created_with_valid_supplier_id` | Passando |
| `test_empty_supplier_id_should_fail` | Passando |

---

## 8. Decisões de Design Registradas

| Decisão | Motivo |
|---|---|
| `planned_date` aceita datas passadas | Pedidos são frequentemente inseridos postumamente no sistema |
| Objetos de negócio separados dos ORM models | Melhor separação de responsabilidades; models não conhecem o banco |
| Nomes de código em inglês | Consistência e compatibilidade com ferramentas; mensagens de erro em português |
| SQLite para dev, PostgreSQL para produção | [ TODO ] — ambiente de dev local ainda não configurado com Docker |
| Streamlit como frontend inicial | Curva de aprendizado menor; migrar para React quando o sistema estiver maduro |
| Arquivos sempre em UTF-8 | Evitar erros de encoding no Windows com caracteres especiais |

---

## 9. Próximos Passos Sugeridos

1. Completar os campos restantes de `PurchaseOrder` com TDD
2. Criar models de `Supplier`, `Project` e `Customer` com TDD
3. Criar ORM models (`orm_models.py`) e configurar SQLAlchemy session
4. Implementar `Repository` para `PurchaseOrder`
5. Implementar `PurchaseOrderService` com mocks nos testes
6. Criar primeiras rotas FastAPI
7. Configurar Docker Compose com os três containers
8. Criar tela inicial no Streamlit

---

*Última atualização: 26/04/2026 — sessão de onboarding TDD e arquitetura Python*
