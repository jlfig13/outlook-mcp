"""
Autenticação delegada com Microsoft Graph via MSAL (device code flow).

Na primeira execução, o terminal mostra um código e uma URL
(https://microsoft.com/devicelogin) para você autorizar no navegador,
usando sua conta Microsoft/Outlook normal. Depois disso o token fica
em cache local (token_cache.bin) e é renovado automaticamente.
"""

import os
import atexit
import msal

CLIENT_ID = os.environ.get("OUTLOOK_MCP_CLIENT_ID", "")
TENANT_ID = os.environ.get("OUTLOOK_MCP_TENANT_ID", "consumers")  # 'consumers' = contas @outlook/@hotmail pessoais
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"

SCOPES = [
    "Mail.Read",
    "Mail.ReadWrite",
    "Mail.Send",  # remova se não precisar enviar e-mails
]

# As ferramentas de regras (list_rules/create_rule/delete_rule) exigem
# MailboxSettings.ReadWrite — um escopo separado de Mail.*, que dá acesso a
# TODAS as configurações da caixa (incluindo respostas automáticas e fuso).
# Fica desligado por padrão: quem não usa regras não concede esse acesso.
# Para ligar: registre MailboxSettings.ReadWrite no app do Entra ID, defina
# OUTLOOK_MCP_ENABLE_RULES=1 e refaça o login (o consentimento é pedido de novo).
ENABLE_RULES = os.environ.get("OUTLOOK_MCP_ENABLE_RULES", "").strip().lower() in ("1", "true", "yes")
if ENABLE_RULES:
    SCOPES.append("MailboxSettings.ReadWrite")

# Raiz do projeto (um nivel acima do pacote), para o cache continuar
# no mesmo lugar dos entrypoints e do volume do Docker.
_PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CACHE_FILE = os.environ.get(
    "OUTLOOK_MCP_TOKEN_CACHE",
    os.path.join(_PROJECT_ROOT, "token_cache.bin"),
)


def _load_cache() -> msal.SerializableTokenCache:
    cache = msal.SerializableTokenCache()
    if os.path.exists(CACHE_FILE):
        with open(CACHE_FILE, "r") as f:
            cache.deserialize(f.read())

    def _save():
        if cache.has_state_changed:
            with open(CACHE_FILE, "w") as f:
                f.write(cache.serialize())

    atexit.register(_save)
    return cache


def _app() -> msal.PublicClientApplication:
    if not CLIENT_ID:
        raise RuntimeError(
            "Defina a variável de ambiente OUTLOOK_MCP_CLIENT_ID com o Application (client) ID "
            "do app registrado no Entra ID (Azure AD). Veja o README.md."
        )
    return msal.PublicClientApplication(
        client_id=CLIENT_ID,
        authority=AUTHORITY,
        token_cache=_load_cache(),
    )


def get_access_token() -> str:
    app = _app()
    accounts = app.get_accounts()

    result = None
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])

    if not result:
        flow = app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            raise RuntimeError(f"Falha ao iniciar device flow: {flow}")
        print(flow["message"])  # instrução para o usuário autorizar no navegador
        result = app.acquire_token_by_device_flow(flow)

    if "access_token" not in result:
        raise RuntimeError(f"Falha na autenticação: {result.get('error_description', result)}")

    return result["access_token"]
