import asyncio
import logging
from datetime import datetime, timezone

from app.supabase_client import get_supabase_client
from app.config import SUPABASE_ENABLED, WHATSAPP_ALLOWED_GROUP_ID, WHATSAPP_ALLOWED_NUMBERS
from app.routes.whatsapp import _extract_sender, _extract_text, _command_response, _build_summary_for_number
from app.utils import send_whatsapp_via_evolution_api

logger = logging.getLogger(__name__)

async def process_inbound_webhooks():
    """Worker em background que processa mensagens inseridas pelo webhook do Supabase."""
    if not SUPABASE_ENABLED:
        logger.warning("Supabase não está habilitado. Worker de webhooks inativo.")
        return

    logger.info("Iniciando worker de webhooks do Supabase...")
    
    while True:
        try:
            client = get_supabase_client()
            
            # Busca mensagens pendentes (limitando a 10 por vez)
            response = (
                client.table("whatsapp_inbound")
                .select("*")
                .eq("status", "pending")
                .order("created_at")
                .limit(10)
                .execute()
            )
            
            messages = response.data or []
            
            for msg_record in messages:
                record_id = msg_record["id"]
                payload = msg_record.get("payload", {})
                
                try:
                    # Lógica similar à do endpoint /webhook
                    if payload.get("event") != "messages.upsert":
                        _delete_record(client, record_id)
                        continue

                    data = payload.get("data") or {}
                    sender_number = _extract_sender(payload)
                    message = data.get("message") or {}
                    text = _extract_text(message).strip()

                    if not sender_number or not text:
                        _delete_record(client, record_id)
                        continue

                    # Verifica se o remetente é o grupo permitido ou um número autorizado
                    is_allowed_group = WHATSAPP_ALLOWED_GROUP_ID and sender_number == WHATSAPP_ALLOWED_GROUP_ID
                    is_allowed_number = sender_number in WHATSAPP_ALLOWED_NUMBERS
                    
                    if not is_allowed_group and not is_allowed_number:
                        _delete_record(client, record_id)
                        continue

                    # Processa o comando
                    response_text = _command_response(text, sender_number)
                    if response_text is None:
                        _delete_record(client, record_id)
                        continue

                    if response_text == "__SEND_SUMMARY__":
                        response_text = _build_summary_for_number(sender_number)

                    # Envia a resposta via Evolution API
                    send_whatsapp_via_evolution_api(message=response_text, number=sender_number)
                    
                    # Marca como processado com sucesso
                    _mark_as_processed(client, record_id, status="processed")
                    
                except Exception as e:
                    logger.error(f"Erro ao processar mensagem {record_id}: {e}")
                    _mark_as_processed(client, record_id, status="error", error=str(e))
                    
        except Exception as e:
            logger.error(f"Erro no loop do worker do Supabase: {e}")
            
        # Espera 5 segundos antes da próxima checagem
        await asyncio.sleep(5)

def _mark_as_processed(client, record_id: str, status: str, error: str = None):
    try:
        update_data = {
            "status": status,
            "processed_at": datetime.now(timezone.utc).isoformat()
        }
        if error:
            update_data["error"] = error
            
        client.table("whatsapp_inbound").update(update_data).eq("id", record_id).execute()
    except Exception as e:
        logger.error(f"Erro ao atualizar status do registro {record_id}: {e}")

def _delete_record(client, record_id: str):
    """Remove o registro da tabela para evitar acumulo de lixo."""
    try:
        client.table("whatsapp_inbound").delete().eq("id", record_id).execute()
    except Exception as e:
        logger.error(f"Erro ao deletar registro {record_id}: {e}")
