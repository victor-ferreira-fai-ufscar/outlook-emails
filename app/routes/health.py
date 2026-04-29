"""
Rota de health check da aplicacao.
"""

from fastapi import APIRouter
from fastapi.responses import HTMLResponse

from app.config import BOT_LOGIN_URL, EVOLUTION_DEFAULT_NUMBER

router = APIRouter()


@router.get("/")
def root() -> HTMLResponse:
    """Pagina principal com onboarding pratico para uso diario."""
    login_url = BOT_LOGIN_URL or "/auth/login"
    default_number = (EVOLUTION_DEFAULT_NUMBER or "").strip()
    wa_link = (
        f"https://wa.me/{default_number}?text=login"
        if default_number
        else "https://wa.me/SEU_NUMERO?text=login"
    )

    html = f"""
<!doctype html>
<html lang="pt-BR">
    <head>
        <meta charset="utf-8" />
        <meta name="viewport" content="width=device-width, initial-scale=1" />
        <title>Onboarding | Outlook Emails</title>
        <style>
            :root {{
                --bg: #f5efe7;
                --panel: #fffaf2;
                --ink: #1f2a34;
                --muted: #4f6372;
                --accent: #1f8a70;
                --accent-2: #f26b38;
                --line: #dfd3c3;
            }}
            * {{ box-sizing: border-box; }}
            body {{
                margin: 0;
                font-family: "Segoe UI", Tahoma, sans-serif;
                color: var(--ink);
                background:
                    radial-gradient(circle at 12% 10%, #ffe4c5 0%, transparent 30%),
                    radial-gradient(circle at 88% 15%, #d7f9f1 0%, transparent 32%),
                    var(--bg);
            }}
            .container {{
                max-width: 960px;
                margin: 24px auto;
                padding: 0 16px 32px;
            }}
            .hero {{
                background: linear-gradient(140deg, #fff6e9, #f2fff7);
                border: 1px solid var(--line);
                border-radius: 18px;
                padding: 24px;
            }}
            h1 {{ margin: 0 0 8px; font-size: 1.9rem; }}
            p {{ margin: 0; color: var(--muted); line-height: 1.45; }}
            .actions {{ margin-top: 16px; display: flex; flex-wrap: wrap; gap: 10px; }}
            .btn {{
                display: inline-block;
                text-decoration: none;
                color: white;
                background: var(--accent);
                border-radius: 10px;
                padding: 10px 14px;
                font-weight: 600;
            }}
            .btn.alt {{ background: var(--accent-2); }}
            .grid {{
                display: grid;
                grid-template-columns: repeat(auto-fit, minmax(260px, 1fr));
                gap: 14px;
                margin-top: 14px;
            }}
            .card {{
                background: var(--panel);
                border: 1px solid var(--line);
                border-radius: 14px;
                padding: 16px;
            }}
            h2 {{ margin: 0 0 8px; font-size: 1.1rem; }}
            ol, ul {{ margin: 8px 0 0; padding-left: 18px; }}
            li {{ margin-bottom: 6px; }}
            code {{
                background: #f0e5d8;
                border: 1px solid #e1d3c4;
                border-radius: 6px;
                padding: 1px 6px;
            }}
            .footer {{ margin-top: 10px; font-size: .95rem; color: var(--muted); }}
            .profile {{
                margin-top: 14px;
                background: rgba(255, 255, 255, 0.72);
                border: 1px solid var(--line);
                border-radius: 12px;
                padding: 12px;
                display: none;
                align-items: center;
                gap: 10px;
            }}
            .profile img {{
                width: 56px;
                height: 56px;
                border-radius: 50%;
                object-fit: cover;
                border: 2px solid #f1e2d2;
                background: #fff;
            }}
            .profile .meta {{ line-height: 1.35; }}
            .profile .meta strong {{ display: block; }}
        </style>
    </head>
    <body>
        <main class="container">
            <section class="hero">
                <h1>Onboarding do Outlook + WhatsApp</h1>
                <p>Use esta pagina como guia rapido para testar e operar o fluxo completo sem se perder.</p>
                <div class="actions">
                    <a class="btn" href="{login_url}">Conectar Outlook (/auth/login)</a>
                    <a class="btn alt" href="{wa_link}" target="_blank" rel="noreferrer">Abrir WhatsApp (login)</a>
                    <a class="btn" href="/docs">Ver Swagger (/docs)</a>
                </div>
                <div id="profile-box" class="profile" aria-live="polite">
                    <img id="profile-photo" src="" alt="Foto de perfil" />
                    <div class="meta">
                        <strong id="profile-name">Perfil autenticado</strong>
                        <span id="profile-email">Carregando dados do Outlook...</span>
                    </div>
                </div>
            </section>

            <section class="grid">
                <article class="card">
                    <h2>Como comecar</h2>
                    <ol>
                        <li>Abra o WhatsApp no link acima e envie <code>login</code>.</li>
                        <li>Abra o link de autenticacao recebido e conclua o login Outlook.</li>
                        <li>Volte no chat e envie <code>status</code> para validar o vinculo.</li>
                        <li>Envie <code>resumo agora</code> para receber o resumo de e-mails.</li>
                    </ol>
                </article>

                <article class="card">
                    <h2>Comandos no WhatsApp</h2>
                    <ul>
                        <li><code>ajuda</code></li>
                        <li><code>login</code></li>
                        <li><code>status</code></li>
                        <li><code>perfil</code></li>
                        <li><code>ultimo-email</code></li>
                        <li><code>resumo agora</code></li>
                    </ul>
                </article>

                <article class="card">
                    <h2>Links uteis</h2>
                    <ul>
                        <li><a href="/auth/login">/auth/login</a></li>
                        <li><a href="/notifications/settings">/notifications/settings</a></li>
                        <li><a href="/notifications/daily-summary">/notifications/daily-summary</a></li>
                        <li><a href="/whatsapp/webhook">/whatsapp/webhook</a></li>
                        <li><a href="/docs">/docs</a></li>
                    </ul>
                    <p class="footer">Se algo falhar, valide primeiro se a instância da Evolution esta com estado <code>open</code>.</p>
                </article>
            </section>
        </main>
        <script>
            (async () => {{
                const params = new URLSearchParams(window.location.search);
                const box = document.getElementById("profile-box");
                const nameEl = document.getElementById("profile-name");
                const emailEl = document.getElementById("profile-email");
                const photoEl = document.getElementById("profile-photo");

                try {{
                    const profileResp = await fetch("/profile", {{ credentials: "include" }});
                    if (!profileResp.ok) return;

                    const profile = await profileResp.json();
                    const name = profile.displayName || "Usuario autenticado";
                    const email = profile.mail || profile.userPrincipalName || "sem-email";
                    nameEl.textContent = name;
                    emailEl.textContent = email;
                    box.style.display = "flex";

                    const photoResp = await fetch("/profile/photo", {{ credentials: "include" }});
                    if (photoResp.ok) {{
                        const blob = await photoResp.blob();
                        photoEl.src = URL.createObjectURL(blob);
                    }} else {{
                        photoEl.style.display = "none";
                    }}

                    if (params.get("welcome") === "1") {{
                        emailEl.textContent = `${{email}} · login concluido com sucesso`;
                    }}
                }} catch (_error) {{
                    return;
                }}
            }})();
        </script>
    </body>
</html>
"""

    return HTMLResponse(content=html)
