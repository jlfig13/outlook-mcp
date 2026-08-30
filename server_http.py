"""Entrypoint HTTP: roda o servidor MCP na rede local (ex: Termux/Docker)."""

from outlook_mcp.http_app import app, main

if __name__ == "__main__":
    main()
