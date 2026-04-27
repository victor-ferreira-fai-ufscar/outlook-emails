"""
Template para implementação seguindo TDD.
Use este arquivo como referência após escrever seus testes.

Estrutura:
1. Imports necessários
2. Router do FastAPI com prefix
3. Função helper privada (se needed)
4. Endpoint público

Este arquivo implementa o "GREEN" do Red-Green-Refactor.
"""

from typing import Any
from fastapi import APIRouter, HTTPException, Request

from app.utils import get_local_access_token

# Criar router com prefixo apropriado
router = APIRouter(prefix="/new", tags=["features"])


# ============================================================================
# ENDPOINTS (GREEN Phase)
# ============================================================================


@router.get("/endpoint")
def new_endpoint(request: Request) -> dict[str, Any]:
    """
    Novo endpoint que retorna dados processados.

    Requer autenticação via session cookie local_session_id.

    Returns:
        dict: Resposta com campos field_one, field_two, e status.

    Raises:
        HTTPException 401: Se não autenticado.
    """
    # Obter token - lança 401 se não autenticado
    access_token = get_local_access_token(request)

    # Implementar lógica minimalista necessária para passar testes
    return {"field_one": "value1", "field_two": "value2", "status": "success"}


# ============================================================================
# HELPER FUNCTIONS (use conforme necessário)
# ============================================================================


def _fetch_external_data(access_token: str) -> dict[str, Any]:
    """
    Helper privada para buscar dados de API externa.

    Use para manter endpoint limpo e lógica isolada.
    Facilita testing via mock.
    """
    import requests
    from app.config import GRAPH_BASE_URL

    response = requests.get(
        f"{GRAPH_BASE_URL}/endpoint",
        headers={"Authorization": f"Bearer {access_token}"},
        timeout=30,
    )

    if response.status_code >= 400:
        from fastapi import HTTPException

        raise HTTPException(status_code=response.status_code, detail=response.text)

    return response.json()


def _process_data(raw_data: dict) -> dict[str, Any]:
    """
    Helper privada para processar dados brutos.

    Use para separar lógica de processamento da chamada de API.
    """
    return {
        "field_one": raw_data.get("field_one"),
        "field_two": raw_data.get("field_two"),
        "status": "processed",
    }


# ============================================================================
# EXEMPLO COMPLETO (uncomment quando implementar versão final)
# ============================================================================

# @router.get("/endpoint/advanced")
# def new_endpoint_advanced(request: Request) -> dict[str, Any]:
#     """Versão completa do endpoint com erro handling."""
#     access_token = get_local_access_token(request)
#
#     # Buscar dados externos
#     external_data = _fetch_external_data(access_token)
#
#     # Processar
#     result = _process_data(external_data)
#
#     return result


# ============================================================================
# REGISTRAR ROUTER EM app/main.py
# ============================================================================

"""
Após criar este arquivo em app/routes/seu_arquivo.py:

1. Atualize app/main.py:
   
   from app.routes import seu_arquivo
   
   app.include_router(seu_arquivo.router)

2. Seu endpoint estará disponível em: /new/endpoint

3. Execute testes:
   pytest tests/test_seu_arquivo.py -v

4. Refatore conforme necessário
"""
