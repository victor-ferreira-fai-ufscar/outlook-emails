import json
from cryptography.fernet import Fernet
from app.config import ENCRYPTION_KEY

def get_cipher():
    """Retorna uma instância de Fernet se a chave estiver configurada."""
    if not ENCRYPTION_KEY:
        return None
    try:
        return Fernet(ENCRYPTION_KEY.encode())
    except Exception:
        return None

def encrypt_data(data: dict | str) -> str:
    """
    Criptografa um dicionário ou string usando AES-256 (Fernet).
    Retorna o texto plano se a chave não estiver configurada (fallback de transição).
    """
    if isinstance(data, dict):
        plain_text = json.dumps(data)
    else:
        plain_text = data

    cipher = get_cipher()
    if not cipher:
        return plain_text
    
    return cipher.encrypt(plain_text.encode()).decode()

def decrypt_data(encrypted_data: str) -> dict | str:
    """
    Descriptografa uma string usando AES-256 (Fernet).
    Tenta retornar um dicionário se o conteúdo for JSON.
    """
    cipher = get_cipher()
    if not cipher:
        # Tenta decodificar como JSON direto se não houver cifra (sessões antigas)
        try:
            return json.loads(encrypted_data)
        except Exception:
            return encrypted_data
    
    try:
        decrypted_text = cipher.decrypt(encrypted_data.encode()).decode()
        try:
            return json.loads(decrypted_text)
        except Exception:
            return decrypted_text
    except Exception:
        # Se falhar a descriptografia, pode ser uma sessão antiga em texto plano
        try:
            return json.loads(encrypted_data)
        except Exception:
            return encrypted_data
