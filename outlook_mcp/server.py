"""
Servidor MCP local para organizar e-mails do Outlook via Microsoft Graph API.

Ferramentas expostas:
  - list_folders
  - list_recent_emails
  - sender_stats
  - search_emails
  - get_email_content
  - move_email
  - move_emails_batch
  - move_by_sender
  - preview_plan
  - mark_as_read_batch
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


# Campo devolvido -> campo do Graph. Pedir menos campos encolhe tanto a
# resposta do Graph quanto o texto que o modelo precisa ler: um bodyPreview
# são ~250 caracteres por e-mail, um id do Graph são ~150.
CAMPOS = {
    "id": "id",
    "assunto": "subject",
    "de": "from",
    "recebido_em": "receivedDateTime",
    "lido": "isRead",
    "preview": "bodyPreview",
    "pasta_id": "parentFolderId",
}
CAMPOS_PADRAO = list(CAMPOS)


def _select_para(campos: list[str]) -> str:
    return ",".join(CAMPOS[c] for c in campos)


def _summarize_message(m: dict, campos: Optional[list[str]] = None) -> dict:
    campos = campos or CAMPOS_PADRAO
    fonte = {
        "id": m.get("id"),
        "assunto": m.get("subject"),
        "de": (m.get("from") or {}).get("emailAddress", {}).get("address"),
        "recebido_em": m.get("receivedDateTime"),
        "lido": m.get("isRead"),
        "preview": m.get("bodyPreview"),
        "pasta_id": m.get("parentFolderId"),
    }
    return {c: fonte[c] for c in campos}


def _validar_campos(fields: Optional[list[str]]) -> list[str]:
    if not fields:
        return CAMPOS_PADRAO
    invalidos = [f for f in fields if f not in CAMPOS]
    if invalidos:
        raise ValueError(f"campos inválidos: {invalidos}. Válidos: {list(CAMPOS)}")
    return fields


def _iter_messages(folder: str = "inbox", unread_only: bool = False,
                   campos: Optional[list[str]] = None, max_items: int = 200,
                   page_size: int = 100, skip: int = 0):
    """
    Percorre mensagens seguindo o @odata.nextLink do Graph, que pagina em
    blocos (o $top é o tamanho da página, não um total). Para além de ~1000
    mensagens é a única forma de alcançar o backlog antigo.

    skip só se aplica à primeira página — a partir daí o nextLink já carrega
    a posição. Serve para retomar uma varredura em fatias (0-400, 400-800...)
    quando um único max_scan grande estourasse o tempo de uma chamada.
    """
    campos = campos or CAMPOS_PADRAO
    params = {
        "$top": min(page_size, max_items),
        "$orderby": "receivedDateTime desc",
        "$select": _select_para(campos),
    }
    if skip:
        params["$skip"] = skip
    if unread_only:
        params["$filter"] = "isRead eq false"

    url = f"{GRAPH_BASE}/me/mailFolders/{folder}/messages"
    lidos = 0
    with httpx.Client(timeout=60) as client:
        while url and lidos < max_items:
            resp = client.get(url, headers=_headers(), params=params)
            resp.raise_for_status()
            data = resp.json()
            for m in data.get("value", []):
                yield m
                lidos += 1
                if lidos >= max_items:
                    return
            url = data.get("@odata.nextLink")
            params = None  # o nextLink já carrega todos os parâmetros


@mcp.tool()
def list_folders() -> str:
    """Lista as pastas de e-mail do usuário (Inbox, Arquivo, pastas customizadas etc.)."""
    data = _graph_get("/me/mailFolders", params={"$top": 100})
    folders = [{"id": f["id"], "nome": f["displayName"], "nao_lidos": f.get("unreadItemCount")}
               for f in data.get("value", [])]
    return json.dumps(folders, ensure_ascii=False, indent=2)


@mcp.tool()
def list_recent_emails(
    folder: str = "inbox",
    top: int = 20,
    skip: int = 0,
    unread_only: bool = False,
    fields: Optional[list[str]] = None,
) -> str:
    """
    Lista os e-mails mais recentes de uma pasta.

    Args:
        folder: nome ou id da pasta (padrão: inbox)
        top: quantidade máxima de e-mails a retornar (padrão: 20)
        skip: quantos e-mails pular antes de começar — use para paginar o backlog
              (ex: skip=100 pega do 101º em diante)
        unread_only: retornar apenas não lidos
        fields: subconjunto de campos a devolver, para economizar contexto.
                Válidos: id, assunto, de, recebido_em, lido, preview, pasta_id.
                Ex: ["id","de","assunto"] corta o preview de ~250 caracteres por e-mail.
    """
    try:
        campos = _validar_campos(fields)
    except ValueError as e:
        return json.dumps({"erro": str(e)}, ensure_ascii=False)

    params = {
        "$top": top,
        "$orderby": "receivedDateTime desc",
        "$select": _select_para(campos),
    }
    if skip:
        params["$skip"] = skip
    if unread_only:
        params["$filter"] = "isRead eq false"

    data = _graph_get(f"/me/mailFolders/{folder}/messages", params=params)
    emails = [_summarize_message(m, campos) for m in data.get("value", [])]
    return json.dumps({
        "total_retornado": len(emails),
        "skip_usado": skip,
        "proximo_skip": skip + len(emails) if len(emails) == top else None,
        "emails": emails,
    }, ensure_ascii=False, indent=2)


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
def sender_stats(folder: str = "inbox", max_scan: int = 400, skip: int = 0,
                 unread_only: bool = False, top_remetentes: int = 50) -> str:
    """
    Devolve o mapa da caixa agrupado por remetente, em vez da lista de e-mails.

    Feito para ser a PRIMEIRA chamada ao organizar: para classificar e desenhar
    regras o que importa é quem manda e quanto, não o conteúdo de cada mensagem.
    Varre as mensagens pedindo só remetente/assunto/isRead e agrega aqui no
    servidor — a resposta sai em poucos KB, contra centenas de KB de uma
    listagem equivalente.

    Para caixas grandes, prefira várias chamadas com max_scan~400 e skip
    crescente (0, 400, 800...) a uma única chamada com max_scan alto — o
    total de itens percorridos entra na duração da chamada.

    Args:
        folder: pasta a varrer (padrão: inbox)
        max_scan: teto de mensagens a percorrer nesta chamada (padrão: 400)
        skip: quantas mensagens já varridas em chamadas anteriores pular
        unread_only: contar apenas não lidos
        top_remetentes: quantos remetentes devolver, do maior volume para o menor
    """
    campos = ["de", "assunto", "lido"]
    agg: dict[str, dict] = {}
    total = 0

    for m in _iter_messages(folder=folder, unread_only=unread_only, campos=campos,
                            max_items=max_scan, skip=skip):
        total += 1
        end = (m.get("from") or {}).get("emailAddress", {}).get("address") or "(sem remetente)"
        end = end.lower()
        e = agg.setdefault(end, {
            "remetente": end,
            "dominio": end.split("@")[-1] if "@" in end else "",
            "total": 0,
            "nao_lidos": 0,
            "assuntos_exemplo": [],
        })
        e["total"] += 1
        if not m.get("isRead"):
            e["nao_lidos"] += 1
        assunto = m.get("subject")
        # Até 3 exemplos distintos: um só esconde remetentes mistos (ex: o
        # mesmo endereço mandando fatura e cupom promocional).
        if assunto and len(e["assuntos_exemplo"]) < 3 and assunto not in e["assuntos_exemplo"]:
            e["assuntos_exemplo"].append(assunto)

    ranking = sorted(agg.values(), key=lambda x: x["total"], reverse=True)

    dominios: dict[str, int] = {}
    for e in agg.values():
        if e["dominio"]:
            dominios[e["dominio"]] = dominios.get(e["dominio"], 0) + e["total"]
    top_dom = sorted(dominios.items(), key=lambda kv: kv[1], reverse=True)[:15]

    atingiu_limite = total >= max_scan
    return json.dumps({
        "mensagens_varridas": total,
        "skip_usado": skip,
        "atingiu_limite": atingiu_limite,
        "proximo_skip": skip + total if atingiu_limite else None,
        "remetentes_distintos": len(agg),
        "top_dominios": [{"dominio": d, "total": n} for d, n in top_dom],
        "remetentes": ranking[:top_remetentes],
    }, ensure_ascii=False, indent=2)


def _compilar_regra(sender_contains: str, subject_contains: Optional[list[str]],
                    subject_not_contains: Optional[list[str]]):
    """Devolve uma função (mensagem -> bool) para o critério de uma regra de plano."""
    alvo = sender_contains.strip().lower()
    inclui = [s.lower() for s in (subject_contains or [])]
    exclui = [s.lower() for s in (subject_not_contains or [])]

    def casa(m: dict) -> bool:
        end = ((m.get("from") or {}).get("emailAddress", {}).get("address") or "").lower()
        if alvo and alvo not in end:
            return False
        assunto_lower = (m.get("subject") or "").lower()
        if inclui and not any(s in assunto_lower for s in inclui):
            return False
        if exclui and any(s in assunto_lower for s in exclui):
            return False
        return True

    return casa


@mcp.tool()
def preview_plan(rules: list[dict], folder: str = "inbox", max_scan: int = 400,
                 skip: int = 0, unread_only: bool = False, amostra_por_regra: int = 3,
                 max_nao_classificados: int = 30) -> str:
    """
    Avalia várias regras de classificação numa única varredura, em vez de uma
    varredura por regra. Pensado para o passo de planejamento antes de mover
    de fato: classificar N remetentes com move_by_sender custa N varreduras
    completas da mesma janela de mensagens; aqui custa uma.

    Cada mensagem é testada contra as regras NA ORDEM DADA e cai na primeira
    que casar (mesma semântica de stopProcessingRules do create_rule) — por
    isso a soma dos 'encontrados' bate com 'mensagens_varridas' menos
    'nao_classificados', sem precisar somar na mão para conferir sobreposição.
    O que nenhuma regra pegar aparece em 'nao_classificados', com amostra —
    é o que sobra sem destino nesta janela.

    Não move nada — é sempre um dry_run. Para executar, aplique move_by_sender
    (ou move_emails_batch) regra por regra, usando os mesmos filtros aqui
    validados.

    Args:
        rules: lista de {sender_contains, target_folder, subject_contains?, subject_not_contains?}.
               target_folder só rotula a saída aqui — nenhuma pasta precisa existir ainda.
        folder: pasta a varrer (padrão: inbox)
        max_scan: teto de mensagens a percorrer nesta chamada (padrão: 400)
        skip: quantas mensagens já varridas em chamadas anteriores pular
        unread_only: considerar apenas não lidos
        amostra_por_regra: quantos exemplos mostrar por regra que casou
        max_nao_classificados: quantos exemplos mostrar do que sobrou sem destino
    """
    if not rules:
        return json.dumps({"erro": "rules está vazio"}, ensure_ascii=False)

    compiladas = []
    for i, r in enumerate(rules):
        if not r.get("sender_contains") or not r.get("target_folder"):
            return json.dumps(
                {"erro": f"rules[{i}] precisa de sender_contains e target_folder"},
                ensure_ascii=False,
            )
        compiladas.append({
            "regra": r,
            "casa": _compilar_regra(r["sender_contains"], r.get("subject_contains"),
                                    r.get("subject_not_contains")),
            "encontrados": 0,
            "amostra": [],
        })

    nao_classificados_total = 0
    nao_classificados_amostra = []
    varridas = 0

    for m in _iter_messages(folder=folder, unread_only=unread_only,
                            campos=["id", "de", "assunto"], max_items=max_scan, skip=skip):
        varridas += 1
        end = ((m.get("from") or {}).get("emailAddress", {}).get("address") or "").lower()

        for c in compiladas:
            if c["casa"](m):
                c["encontrados"] += 1
                if len(c["amostra"]) < amostra_por_regra:
                    c["amostra"].append({"de": end, "assunto": m.get("subject")})
                break
        else:
            nao_classificados_total += 1
            if len(nao_classificados_amostra) < max_nao_classificados:
                nao_classificados_amostra.append({"de": end, "assunto": m.get("subject")})

    atingiu_limite = varridas >= max_scan
    return json.dumps({
        "mensagens_varridas": varridas,
        "skip_usado": skip,
        "atingiu_limite_varredura": atingiu_limite,
        "proximo_skip": skip + varridas if atingiu_limite else None,
        "regras": [{
            "sender_contains": c["regra"]["sender_contains"],
            "subject_contains": c["regra"].get("subject_contains"),
            "subject_not_contains": c["regra"].get("subject_not_contains"),
            "target_folder": c["regra"]["target_folder"],
            "encontrados": c["encontrados"],
            "amostra": c["amostra"],
        } for c in compiladas],
        "nao_classificados": {
            "total": nao_classificados_total,
            "amostra": nao_classificados_amostra,
        },
    }, ensure_ascii=False, indent=2)


@mcp.tool()
def move_by_sender(
    sender_contains: str,
    target_folder: str,
    folder: str = "inbox",
    dry_run: bool = True,
    max_scan: int = 400,
    skip: int = 0,
    unread_only: bool = False,
    subject_contains: Optional[list[str]] = None,
    subject_not_contains: Optional[list[str]] = None,
) -> str:
    """
    Move todos os e-mails de um remetente sem precisar transportar os IDs.

    O filtro acontece aqui no servidor: os ids nunca precisam trafegar até quem
    chamou (100 ids do Graph são ~15 KB só de identificadores).

    Por padrão roda em dry_run: apenas conta e mostra uma amostra do que SERIA
    movido. Para executar de fato é preciso passar dry_run=False explicitamente
    — assim ninguém move "tudo do remetente X, seja quanto for" por engano.

    Para caixas grandes, prefira várias chamadas com max_scan~400 e skip
    crescente a uma única chamada com max_scan alto (veja atingiu_limite_varredura
    e proximo_skip na resposta). Um remetente pode misturar tipos de e-mail (ex:
    fatura e cupom promocional do mesmo endereço) — nesse caso use
    subject_contains/subject_not_contains para separar em vez de mover tudo.

    Args:
        sender_contains: trecho do endereço do remetente (ex: '99freelas', '@newsletter.com')
        target_folder: pasta destino (nome bem-conhecido, nome de exibição ou id)
        folder: pasta de origem (padrão: inbox)
        dry_run: se True (padrão), só simula e devolve a contagem
        max_scan: teto de mensagens a percorrer nesta chamada (padrão: 400)
        skip: quantas mensagens já varridas em chamadas anteriores pular
        unread_only: considerar apenas não lidos
        subject_contains: só move se o assunto contiver algum destes textos
        subject_not_contains: não move se o assunto contiver algum destes textos
                              (avaliado depois de subject_contains)
    """
    alvo = sender_contains.strip().lower()
    if not alvo:
        return json.dumps({"erro": "sender_contains está vazio"}, ensure_ascii=False)

    casa = _compilar_regra(sender_contains, subject_contains, subject_not_contains)

    campos = ["id", "de", "assunto"]
    casaram, amostra = [], []
    varridas = 0
    for m in _iter_messages(folder=folder, unread_only=unread_only, campos=campos,
                            max_items=max_scan, skip=skip):
        varridas += 1
        if not casa(m):
            continue
        end = ((m.get("from") or {}).get("emailAddress", {}).get("address") or "").lower()
        casaram.append(m["id"])
        if len(amostra) < 5:
            amostra.append({"de": end, "assunto": m.get("subject")})

    atingiu_limite = varridas >= max_scan
    base = {
        "sender_contains": sender_contains,
        "subject_contains": subject_contains,
        "subject_not_contains": subject_not_contains,
        "pasta_origem": folder,
        "pasta_destino": target_folder,
        "mensagens_varridas": varridas,
        "skip_usado": skip,
        "atingiu_limite_varredura": atingiu_limite,
        "proximo_skip": skip + varridas if atingiu_limite else None,
        "encontrados": len(casaram),
        "amostra": amostra,
    }
    if atingiu_limite:
        base["aviso"] = ("A varredura bateu o teto desta chamada — pode haver mais mensagens "
                         "além de mensagens_varridas. Repita com skip=proximo_skip para continuar.")

    if dry_run:
        base["dry_run"] = True
        base["proximo_passo"] = ("Nada foi movido. Para executar, chame de novo com dry_run=False "
                                 "— confirme antes o número em 'encontrados'.")
        return json.dumps(base, ensure_ascii=False, indent=2)

    if not casaram:
        base["dry_run"] = False
        base["movidos"] = 0
        return json.dumps(base, ensure_ascii=False, indent=2)

    resultado = json.loads(move_emails_batch(casaram, target_folder))
    base["dry_run"] = False
    base["movidos"] = resultado["movidos"]
    base["falharam"] = resultado["falharam"]
    base["detalhes_falhas"] = resultado["detalhes_falhas"]
    return json.dumps(base, ensure_ascii=False, indent=2)


@mcp.tool()
def mark_as_read_batch(email_ids: list[str], read: bool = True) -> str:
    """
    Marca vários e-mails como lidos (ou não lidos) em lotes de 20 via POST /$batch.

    Ao contrário do move, o PATCH não troca o id do e-mail — os mesmos ids
    continuam válidos depois.

    Args:
        email_ids: lista de ids
        read: True para marcar como lido, False para não lido
    """
    if not email_ids:
        return json.dumps({"erro": "email_ids está vazio"}, ensure_ascii=False)

    ok, falhas = 0, []
    pendentes = list(email_ids)
    ja_tentou_retry = False

    while pendentes:
        lote, pendentes = pendentes[:BATCH_LIMIT], pendentes[BATCH_LIMIT:]
        por_id = {str(i): eid for i, eid in enumerate(lote)}
        reqs = [{
            "id": str(i),
            "method": "PATCH",
            "url": f"/me/messages/{eid}",
            "headers": {"Content-Type": "application/json"},
            "body": {"isRead": read},
        } for i, eid in enumerate(lote)]

        reenfileirar, espera = [], 0
        for r in _graph_batch(reqs):
            eid = por_id.get(str(r.get("id")), "?")
            status = r.get("status", 0)
            if 200 <= status < 300:
                ok += 1
            elif status == 429 and not ja_tentou_retry:
                reenfileirar.append(eid)
                espera = max(espera, int((r.get("headers") or {}).get("Retry-After", 5)))
            else:
                err = (r.get("body") or {}).get("error") or {}
                falhas.append({"id": eid, "status": status,
                               "motivo": err.get("message") or err.get("code") or "erro desconhecido"})

        if reenfileirar:
            ja_tentou_retry = True
            time.sleep(espera)
            pendentes = reenfileirar + pendentes

    return json.dumps({
        "total_pedido": len(email_ids),
        "atualizados": ok,
        "falharam": len(falhas),
        "isRead": read,
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
