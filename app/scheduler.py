import asyncio
import logging
from datetime import datetime, timezone
from apscheduler.schedulers.asyncio import AsyncIOScheduler

from app.supabase_client import get_supabase_client, get_user_settings
from app.routes.whatsapp import _build_summary_for_number
from app.utils import send_whatsapp_via_evolution_api
from app.config import SUPABASE_ENABLED

logger = logging.getLogger(__name__)

class EmailScheduler:
    def __init__(self):
        self.scheduler = AsyncIOScheduler()
        self._is_running = False

    async def start(self):
        if self._is_running:
            return
        
        # Adiciona tarefa para rodar a cada minuto
        self.scheduler.add_job(self.check_and_send_summaries, 'cron', minute='*')
        self.scheduler.start()
        self._is_running = True
        logger.info("Scheduler de e-mails iniciado.")

    async def stop(self):
        if not self._is_running:
            return
        self.scheduler.shutdown()
        self._is_running = False
        logger.info("Scheduler de e-mails encerrado.")

    async def check_and_send_summaries(self):
        """Verifica quais usuários devem receber resumo no minuto atual."""
        if not SUPABASE_ENABLED:
            return

        now = datetime.now()
        current_time = now.strftime("%H:%M")
        logger.debug(f"Verificando agendamentos para {current_time}...")

        try:
            client = get_supabase_client()
            
            # Busca todos os vínculos de WhatsApp
            # Em uma escala maior, isso precisaria de filtros ou paginação
            response = client.table("whatsapp_links").select("user_id, whatsapp_number").execute()
            links = response.data or []

            for link in links:
                user_id = link["user_id"]
                whatsapp_number = link["whatsapp_number"]

                try:
                    settings = get_user_settings(user_id)
                    schedule = settings.get("summary_schedule", "08:00")

                    if schedule == current_time:
                        logger.info(f"Disparando resumo agendado para {whatsapp_number} (User: {user_id})")
                        
                        # Gera o resumo
                        summary_text = _build_summary_for_number(whatsapp_number)
                        
                        # Envia
                        send_whatsapp_via_evolution_api(message=summary_text, number=whatsapp_number)
                        
                except Exception as e:
                    logger.error(f"Erro ao processar agendamento para o usuário {user_id}: {e}")

        except Exception as e:
            logger.error(f"Erro no loop do scheduler: {e}")

# Instância global
email_scheduler = EmailScheduler()
