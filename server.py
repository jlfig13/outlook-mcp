"""Entrypoint stdio: roda o servidor MCP para o Claude Desktop local."""

from outlook_mcp.server import mcp

if __name__ == "__main__":
    mcp.run(transport="stdio")
