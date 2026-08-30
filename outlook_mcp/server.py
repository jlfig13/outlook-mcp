"""
Servidor MCP local para organizar e-mails do Outlook via Microsoft Graph API.

Ferramentas expostas:
  - list_folders
  - list_recent_emails
  - search_emails
  - get_email_content
  - move_email
  - mark_as_read
  - flag_email

Autenticação: MSAL (device code flow), delegada, com cache de token local.
"""

import os
import json
from typing import Optional

import httpx
from mcp.server.fastmcp import FastMCP

from .auth import get_access_token

mcp = FastMCP("outlook-organizer")

GRAPH_BASE = "https://graph.microsoft.com/v1.0"

def _headers() -> dict:
    token = get_access_token()
    return {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}


def _graph_get(path: str, params: Optional[dict] = None) -> dict:
    with httpx.Client(timeout=30) as client:
        resp = client.get(f"{GRAPH_BASE}{path}", headers=_headers(), params=params)
        resp.raise_for_status()
        return resp.json()


def _graph_patch(path: str, body: dict) -> dict:
    with httpx.Client(timeout=30) as client:
        resp = client.patch(f"{GRAPH_BASE}{path}", headers=_headers(), json=body)
        resp.raise_for_status()
        return resp.json() if resp.content else {}


def _graph_post(path: str, body: dict) -> dict:
    with httpx.Client(timeout=30) as client:
        resp = client.post(f"{GRAPH_BASE}{path}", headers=_headers(), json=body)
        resp.raise_for_status()
        return resp.json() if resp.content else {}


def _summarize_message(m: dict) -> dict:
    return {
        "id": m.get("id"),
        "assunto": m.get("subject"),
        "de": (m.get("from") or {}).get("emailAddress", {}).get("address"),
        "recebido_em": m.get("receivedDateTime"),
        "lido": m.get("isRead"),
        "preview": m.get("bodyPreview"),
        "pasta_id": m.get("parentFolderId"),
    }


@mcp.tool()
def list_folders() -> str:
    """Lista as pastas de e-mail do usuário (Inbox, Arquivo, pastas customizadas etc.)."""
    data = _graph_get("/me/mailFolders", params={"$top": 100})
    folders = [{"id": f["id"], "nome": f["displayName"], "nao_lidos": f.get("unreadItemCount")}
               for f in data.get("value", [])]
    return json.dumps(folders, ensure_ascii=False, indent=2)


@mcp.tool()
def list_recent_emails(folder: str = "inbox", top: int = 20) -> str:
    """
    Lista os e-mails mais recentes de uma pasta.

    Args:
        folder: nome ou id da pasta (padrão: inbox)
        top: quantidade máxima de e-mails a retornar (padrão: 20)
    """
    data = _graph_get(
        f"/me/mailFolders/{folder}/messages",
        params={
            "$top": top,
            "$orderby": "receivedDateTime desc",
            "$select": "id,subject,from,receivedDateTime,isRead,bodyPreview,parentFolderId",
        },
    )
    emails = [_summarize_message(m) for m in data.get("value", [])]
    return json.dumps(emails, ensure_ascii=False, indent=2)


@mcp.tool()
def search_emails(query: str, top: int = 20) -> str:
    """
    Busca e-mails em toda a caixa de correio usando busca do Graph ($search).

    Args:
        query: termo de busca (assunto, remetente, corpo etc.)
        top: quantidade máxima de resultados
    """
    with httpx.Client(timeout=30) as client:
        resp = client.get(
            f"{GRAPH_BASE}/me/messages",
            headers={**_headers(), "ConsistencyLevel": "eventual"},
            params={
                "$search": f'"{query}"',
                "$top": top,
                "$select": "id,subject,from,receivedDateTime,isRead,bodyPreview,parentFolderId",
            },
        )
        resp.raise_for_status()
        data = resp.json()
    emails = [_summarize_message(m) for m in data.get("value", [])]
    return json.dumps(emails, ensure_ascii=False, indent=2)


@mcp.tool()
def get_email_content(email_id: str) -> str:
    """Retorna o conteúdo completo (corpo em texto/HTML) de um e-mail específico pelo ID."""
    data = _graph_get(f"/me/messages/{email_id}", params={"$select": "subject,from,receivedDateTime,body"})
    return json.dumps({
        "assunto": data.get("subject"),
        "de": (data.get("from") or {}).get("emailAddress", {}).get("address"),
        "recebido_em": data.get("receivedDateTime"),
        "corpo": (data.get("body") or {}).get("content"),
        "tipo_corpo": (data.get("body") or {}).get("contentType"),
    }, ensure_ascii=False, indent=2)


@mcp.tool()
def move_email(email_id: str, target_folder: str) -> str:
    """
    Move um e-mail para outra pasta.

    Args:
        email_id: id do e-mail
        target_folder: nome bem-conhecido (ex: 'archive', 'deleteditems') ou id da pasta destino
    """
    result = _graph_post(f"/me/messages/{email_id}/move", {"destinationId": target_folder})
    return json.dumps({"movido": True, "novo_id": result.get("id")}, ensure_ascii=False)


@mcp.tool()
def mark_as_read(email_id: str, read: bool = True) -> str:
    """Marca um e-mail como lido (ou não lido)."""
    _graph_patch(f"/me/messages/{email_id}", {"isRead": read})
    return json.dumps({"ok": True, "isRead": read})


@mcp.tool()
def flag_email(email_id: str, status: str = "flagged") -> str:
    """
    Sinaliza (flag) um e-mail para acompanhamento.

    Args:
        email_id: id do e-mail
        status: 'flagged', 'complete' ou 'notFlagged'
    """
    _graph_patch(f"/me/messages/{email_id}", {"flag": {"flagStatus": status}})
    return json.dumps({"ok": True, "status": status})
