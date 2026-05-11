from abc import ABC, abstractmethod
from typing import Dict, Any, List

class BaseProvider(ABC):
    @abstractmethod
    async def gerar_resumo(self, conteudo_email: str, anexos: List[Dict[str, str]] = None) -> Dict[str, Any]:
        """
        Recebe o conteúdo do e-mail e uma lista de anexos (nome e tipo), 
        e retorna um dicionário com o resumo, prioridade e ação.
        
        Args:
            conteudo_email (str): O corpo de texto do e-mail.
            anexos (List[Dict[str, str]]): Lista de anexos contendo 'nome' e 'tipo'.
            
        Returns:
            Dict[str, Any]: Dicionário com chaves como 'resumo', 'prioridade', 'acao', etc.
        """
        pass
