import httpx
import json
from typing import Dict, Any, List
from .base import BaseProvider
from ..Prompt.email_prompts import PROMPT_RESUMO_EMAIL

class OllamaProvider(BaseProvider):
    def __init__(self, base_url: str, model: str):
        self.base_url = base_url
        self.model = model

    async def gerar_resumo(self, conteudo_email: str, anexos: List[Dict[str, str]] = None) -> Dict[str, Any]:
        prompt_completo = PROMPT_RESUMO_EMAIL
        
        info_anexos = ""
        if anexos:
            info_anexos = "Anexos presentes no e-mail:\n"
            for anexo in anexos:
                info_anexos += f"- {anexo.get('nome', 'Desconhecido')} ({anexo.get('tipo', 'Desconhecido')})\n"
            info_anexos += "\n"
            
        prompt_completo = prompt_completo.replace("{{CONTEUDO_EMAIL}}", info_anexos + conteudo_email)

        payload = {
            "model": self.model,
            "prompt": prompt_completo,
            "stream": False,
            "format": "json" # Força a saída em formato JSON
        }

        try:
            async with httpx.AsyncClient() as client:
                response = await client.post(f"{self.base_url}/api/generate", json=payload, timeout=60.0)
                response.raise_for_status()
                
                data = response.json()
                resposta_texto = data.get("response", "{}")
                
                try:
                    resultado_json = json.loads(resposta_texto)
                    return resultado_json
                except json.JSONDecodeError:
                    return {
                        "resumo": "Erro ao interpretar a resposta da IA. O formato não é um JSON válido.",
                        "prioridade": "Indefinida",
                        "acao": "Nenhuma",
                        "raw_response": resposta_texto
                    }
        except Exception as e:
            return {
                "resumo": f"Erro na comunicação com o Ollama: {str(e)}",
                "prioridade": "Erro",
                "acao": "Verificar servidor Ollama"
            }
