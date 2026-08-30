"""
Roda o mesmo servidor MCP (definido em server.py) via HTTP, escutando na
rede local (Wi-Fi de casa) em vez de stdio. Feito para rodar continuamente
num aparelho dedicado (ex: celular Android antigo via Termux) e ser
acessado pelo Claude Desktop de outro dispositivo na mesma rede.

Segurança (duas camadas, use as duas):
  1. Rede: nada disso deve ser exposto na internet — sem port-forward no
     roteador, sem UPnP. Só funciona dentro do Wi-Fi de casa.
  2. Aplicação: token compartilhado via variável de ambiente
     OUTLOOK_MCP_AUTH_TOKEN. Se não for definido, o servidor sobe SEM
     autenticação (avisa no log) — não recomendado.

Variáveis de ambiente:
  OUTLOOK_MCP_HOST          padrão: 0.0.0.0 (escuta em todas interfaces)
  OUTLOOK_MCP_PORT          padrão: 8787
  OUTLOOK_MCP_AUTH_TOKEN    token secreto que os clientes devem enviar
                            como header 'Authorization: Bearer <token>'
  OUTLOOK_MCP_ALLOWED_HOST  host:porta esperado no header Host (proteção
                            contra DNS rebinding). Ex: 192.168.1.50:8787
"""

import os
import secrets

import uvicorn
from starlette.middleware.base import BaseHTTPMiddleware
from starlette.responses import JSONResponse

from .server import mcp

HOST = os.environ.get("OUTLOOK_MCP_HOST", "0.0.0.0")
PORT = int(os.environ.get("OUTLOOK_MCP_PORT", "8787"))
AUTH_TOKEN = os.environ.get("OUTLOOK_MCP_AUTH_TOKEN", "").strip()
ALLOWED_HOST = os.environ.get("OUTLOOK_MCP_ALLOWED_HOST", "").strip()

mcp.settings.host = HOST
mcp.settings.port = PORT

# Protege contra DNS rebinding: só aceita requisições cujo header Host bate
# com o esperado (mais o localhost, útil para testes no próprio aparelho).
allowed_hosts = ["127.0.0.1:*", "localhost:*"]
if ALLOWED_HOST:
    allowed_hosts.append(ALLOWED_HOST)
else:
    # Sem host configurado explicitamente, libera geral dentro da rede local.
    # Prefira sempre configurar OUTLOOK_MCP_ALLOWED_HOST com o IP:porta do aparelho.
    allowed_hosts.append("*")
mcp.settings.transport_security.allowed_hosts = allowed_hosts
mcp.settings.transport_security.allowed_origins = ["*"]


class BearerAuthMiddleware(BaseHTTPMiddleware):
    async def dispatch(self, request, call_next):
        if not AUTH_TOKEN:
            return await call_next(request)
        provided = request.headers.get("authorization", "")
        expected = f"Bearer {AUTH_TOKEN}"
        if not secrets.compare_digest(provided, expected):
            return JSONResponse({"error": "unauthorized"}, status_code=401)
        return await call_next(request)


def build_app():
    app = mcp.streamable_http_app()
    app.add_middleware(BearerAuthMiddleware)
    return app


app = build_app()


def main() -> None:
    if not AUTH_TOKEN:
        print(
            "AVISO: OUTLOOK_MCP_AUTH_TOKEN não definido. O servidor vai subir SEM "
            "autenticação — qualquer aparelho na sua rede Wi-Fi vai conseguir usar "
            "suas ferramentas de e-mail. Defina a variável antes de rodar em produção."
        )
    if not ALLOWED_HOST:
        print(
            "AVISO: OUTLOOK_MCP_ALLOWED_HOST não definido — proteção de DNS rebinding "
            "está permissiva ('*'). Configure com o IP local do aparelho, ex: "
            "192.168.1.50:8787."
        )
    print(f"Servidor MCP (HTTP) escutando em http://{HOST}:{PORT}{mcp.settings.streamable_http_path}")
    uvicorn.run(app, host=HOST, port=PORT, log_level="info")


if __name__ == "__main__":
    main()
