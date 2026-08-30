"""
Servidor MCP local para organizar e-mails do Outlook via Microsoft Graph API.

Ferramentas expostas:
  - list_folders
  - list_recent_emails
  - search_emails
  - get_email_content
  - move_email
  - move_emails_batch
  - list_rules
  - create_rule
  - delete_rule
  - mark_as_read
  - flag_email

Autenticação: MSAL (device code flow), delegada, com cache de token local.
"""

import os
import json
import time
import functools
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


BATCH_LIMIT = 20  # limite do Graph: 20 operações por requisição /$batch


def _graph_batch(requests_: list[dict]) -> list[dict]:
    """
    Executa até 20 sub-requisições numa única chamada a POST /$batch.

    O Graph responde 200 mesmo quando sub-requisições individuais falham —
    o status real de cada uma vem dentro de 'responses'. A ordem da resposta
    não é garantida, por isso cada sub-requisição leva um 'id' próprio.
    """
    with httpx.Client(timeout=60) as client:
        resp = client.post(f"{GRAPH_BASE}/$batch", headers=_headers(), json={"requests": requests_})
        resp.raise_for_status()
        return resp.json().get("responses", [])


_ERRO_ESCOPO_REGRAS = {
    "erro": "acesso negado às regras (403)",
    "causa": "as regras exigem o escopo MailboxSettings.ReadWrite, que não está no token atual "
             "(Mail.Read/ReadWrite/Send não servem para este endpoint).",
    "como_resolver": [
        "1. No Entra ID → seu app → Permissões de APIs → Microsoft Graph → Permissões delegadas, "
        "adicione MailboxSettings.ReadWrite",
        "2. Defina a variável de ambiente OUTLOOK_MCP_ENABLE_RULES=1",
        "3. Apague token_cache.bin e rode o login de novo para consentir o novo escopo",
    ],
}


def _regras_guard(fn):
    """Traduz o 403 de escopo faltando numa mensagem acionável em vez de um traceback."""
    @functools.wraps(fn)
    def wrapper(*a, **kw):
        try:
            return fn(*a, **kw)
        except httpx.HTTPStatusError as e:
            if e.response.status_code == 403:
                return json.dumps(_ERRO_ESCOPO_REGRAS, ensure_ascii=False, indent=2)
            raise
    return wrapper


def _resolve_folder_id(nome_ou_id: str) -> str:
    """
    Regras do Graph exigem o ID da pasta em moveToFolder — nomes bem-conhecidos
    como 'archive' não são aceitos ali (ao contrário do endpoint /move).
    Aceita id cru, nome bem-conhecido ou nome de exibição (sem diferenciar caixa).
    """
    pastas = _graph_get("/me/mailFolders", params={"$top": 100}).get("value", [])

    for f in pastas:
        if f["id"] == nome_ou_id:
            return f["id"]

    alvo = nome_ou_id.strip().lower()
    for f in pastas:
        if f["displayName"].strip().lower() == alvo:
            return f["id"]

    # Nome bem-conhecido (archive, inbox, deleteditems...): o Graph resolve no GET.
    try:
        return _graph_get(f"/me/mailFolders/{nome_ou_id}")["id"]
    except httpx.HTTPStatusError:
        pass

    disponiveis = ", ".join(f["displayName"] for f in pastas)
    raise ValueError(f"pasta '{nome_ou_id}' não encontrada. Disponíveis: {disponiveis}")


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
def move_emails_batch(email_ids: list[str], target_folder: str) -> str:
    """
    Move vários e-mails de uma vez, agrupando em lotes de 20 numa única
    requisição ao Graph (POST /$batch). Para centenas de e-mails isso é
    ~20x menos chamadas de rede do que chamar move_email um por um.

    E-mails que falharem não interrompem o lote: a resposta traz a lista
    de sucessos e a de falhas, cada uma com o motivo.

    Args:
        email_ids: lista de ids de e-mail a mover
        target_folder: nome bem-conhecido (ex: 'archive', 'deleteditems') ou id da pasta destino
    """
    if not email_ids:
        return json.dumps({"erro": "email_ids está vazio"}, ensure_ascii=False)

    movidos: list[dict] = []
    falhas: list[dict] = []
    pendentes = list(email_ids)
    ja_tentou_retry = False

    while pendentes:
        lote, pendentes = pendentes[:BATCH_LIMIT], pendentes[BATCH_LIMIT:]

        # O id da sub-requisição é o índice no lote; guardamos o mapa para
        # reassociar cada resposta ao e-mail certo (a ordem não é garantida).
        por_id = {str(i): eid for i, eid in enumerate(lote)}
        requests_ = [
            {
                "id": str(i),
                "method": "POST",
                "url": f"/me/messages/{eid}/move",
                "headers": {"Content-Type": "application/json"},
                "body": {"destinationId": target_folder},
            }
            for i, eid in enumerate(lote)
        ]

        respostas = _graph_batch(requests_)

        reenfileirar: list[str] = []
        espera = 0
        for r in respostas:
            eid = por_id.get(str(r.get("id")), "?")
            status = r.get("status", 0)
            corpo = r.get("body") or {}

            if 200 <= status < 300:
                movidos.append({"id_antigo": eid, "id_novo": corpo.get("id")})
            elif status == 429 and not ja_tentou_retry:
                # Throttling: o Graph diz em quantos segundos podemos voltar.
                reenfileirar.append(eid)
                espera = max(espera, int((r.get("headers") or {}).get("Retry-After", 5)))
            else:
                erro = corpo.get("error") or {}
                falhas.append({
                    "id": eid,
                    "status": status,
                    "motivo": erro.get("message") or erro.get("code") or "erro desconhecido",
                })

        if reenfileirar:
            ja_tentou_retry = True
            time.sleep(espera)
            pendentes = reenfileirar + pendentes

    return json.dumps({
        "total_pedido": len(email_ids),
        "movidos": len(movidos),
        "falharam": len(falhas),
        "lotes": -(-len(email_ids) // BATCH_LIMIT),
        "detalhes_movidos": movidos,
        "detalhes_falhas": falhas,
    }, ensure_ascii=False, indent=2)


@mcp.tool()
@_regras_guard
def list_rules() -> str:
    """Lista as regras de caixa de entrada já existentes (nome, sequência, condições e ações)."""
    data = _graph_get("/me/mailFolders/inbox/messageRules")
    regras = [{
        "id": r.get("id"),
        "nome": r.get("displayName"),
        "sequencia": r.get("sequence"),
        "ativa": r.get("isEnabled"),
        "condicoes": r.get("conditions"),
        "acoes": r.get("actions"),
    } for r in data.get("value", [])]
    return json.dumps({"total": len(regras), "regras": regras}, ensure_ascii=False, indent=2)


@mcp.tool()
@_regras_guard
def create_rule(
    name: str,
    target_folder: str,
    sender_contains: Optional[list[str]] = None,
    subject_contains: Optional[list[str]] = None,
    body_contains: Optional[list[str]] = None,
    mark_as_read: bool = False,
    stop_processing: bool = True,
    enabled: bool = True,
) -> str:
    """
    Cria uma regra de caixa de entrada que move e-mails para uma pasta.

    Só expõe ações seguras: mover, marcar como lido e parar o processamento.
    Encaminhamento automático (forwardTo/redirectTo) e exclusão permanente
    NÃO são expostos de propósito — regra de encaminhamento é o vetor
    clássico de vazamento de e-mail.

    É preciso informar ao menos uma condição.

    Args:
        name: nome da regra (aparece na interface do Outlook)
        target_folder: pasta destino — id, nome bem-conhecido ('archive') ou nome de exibição ('Arquivo Morto')
        sender_contains: casa se o remetente contiver qualquer um destes textos
        subject_contains: casa se o assunto contiver qualquer um destes textos
        body_contains: casa se o corpo contiver qualquer um destes textos
        mark_as_read: também marcar como lido ao aplicar
        stop_processing: parar de avaliar as regras seguintes quando esta casar (padrão: sim)
        enabled: criar já ativa (padrão: sim)
    """
    condicoes = {}
    if sender_contains:
        condicoes["senderContains"] = sender_contains
    if subject_contains:
        condicoes["subjectContains"] = subject_contains
    if body_contains:
        condicoes["bodyContains"] = body_contains

    if not condicoes:
        return json.dumps(
            {"erro": "informe ao menos uma condição (sender_contains, subject_contains ou body_contains)"},
            ensure_ascii=False,
        )

    try:
        folder_id = _resolve_folder_id(target_folder)
    except ValueError as e:
        return json.dumps({"erro": str(e)}, ensure_ascii=False)

    existentes = _graph_get("/me/mailFolders/inbox/messageRules").get("value", [])
    if any((r.get("displayName") or "").strip().lower() == name.strip().lower() for r in existentes):
        return json.dumps(
            {"erro": f"já existe uma regra chamada '{name}'. Use outro nome ou remova a antiga."},
            ensure_ascii=False,
        )

    acoes = {"moveToFolder": folder_id, "stopProcessingRules": stop_processing}
    if mark_as_read:
        acoes["markAsRead"] = True

    corpo = {
        "displayName": name,
        "sequence": max([r.get("sequence", 0) for r in existentes], default=0) + 1,
        "isEnabled": enabled,
        "conditions": condicoes,
        "actions": acoes,
    }

    r = _graph_post("/me/mailFolders/inbox/messageRules", corpo)
    return json.dumps({
        "criada": True,
        "id": r.get("id"),
        "nome": r.get("displayName"),
        "sequencia": r.get("sequence"),
        "ativa": r.get("isEnabled"),
        "condicoes": r.get("conditions"),
        "acoes": r.get("actions"),
        "aviso": "Regras só valem para e-mails que chegarem a partir de agora; "
                 "o backlog já na caixa precisa ser movido com move_emails_batch.",
    }, ensure_ascii=False, indent=2)


@mcp.tool()
@_regras_guard
def delete_rule(rule_id: str) -> str:
    """
    Remove uma regra de caixa de entrada pelo id (veja os ids em list_rules).
    Apaga só a regra — nenhum e-mail é afetado.
    """
    with httpx.Client(timeout=30) as client:
        resp = client.delete(f"{GRAPH_BASE}/me/mailFolders/inbox/messageRules/{rule_id}", headers=_headers())
        resp.raise_for_status()
    return json.dumps({"removida": True, "id": rule_id}, ensure_ascii=False)


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
