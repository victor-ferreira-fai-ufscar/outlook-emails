from .Provedores.base import BaseProvider
from .Provedores.ollama_provider import OllamaProvider
import os

def get_llm_provider() -> BaseProvider:
    provider_name = os.getenv("LLM_PROVIDER", "ollama").lower()
    
    if provider_name == "ollama":
        # Pega a URL do Ollama do .env, ou usa localhost por padrão
        base_url = os.getenv("OLLAMA_BASE_URL", "http://localhost:11434")
        model = os.getenv("OLLAMA_MODEL", "llama3") # ou mistral, gemma, etc.
        return OllamaProvider(base_url=base_url, model=model)
    
    # Futuramente, implementar OpenAI e Gemini aqui
    raise ValueError(f"Provedor LLM '{provider_name}' não suportado.")
