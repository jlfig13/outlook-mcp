# outlook-mcp

Servidor MCP local que conecta o Claude Desktop à sua caixa de e-mail do Outlook
(via Microsoft Graph API), para listar, buscar e organizar e-mails direto
pela conversa com o Claude.

## Estrutura do projeto

```
MCP Outlook/
├── server.py            # entrypoint stdio (Claude Desktop local)
├── server_http.py       # entrypoint HTTP (rede local: Termux, Docker)
├── outlook_mcp/         # código do servidor
│   ├── __init__.py
│   ├── auth.py          # login MSAL (device code) + cache de token
│   ├── server.py        # definição do MCP e das 14 ferramentas
│   └── http_app.py      # app Starlette + middleware de auth Bearer
├── docker/
│   └── Dockerfile
├── docs/
│   ├── entra-id.md      # registrar o app no Azure AD (passo 1)
│   └── rede-local.md    # rodar via Termux/Android/Docker na Wi-Fi
├── requirements.txt
├── .env.example         # modelo das variáveis de ambiente
└── token_cache.bin      # gerado no 1º login — NUNCA versionar
```

Os dois entrypoints da raiz são atalhos finos: `python server.py` e
`python server_http.py` funcionam como antes.

## 1. Registrar um app no Entra ID (Azure AD)

O Graph API exige um app registrado, mesmo para uso 100% pessoal e gratuito.
Passo a passo completo em **[docs/entra-id.md](docs/entra-id.md)**.
No fim você terá o `OUTLOOK_MCP_CLIENT_ID`.

## 2. Instalar dependências

```bash
cd outlook-mcp
python3 -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

## 3. Configurar variáveis de ambiente

```bash
export OUTLOOK_MCP_CLIENT_ID="cole-o-client-id-aqui"

# "consumers" = contas pessoais @outlook/@hotmail (padrão, não precisa mudar)
export OUTLOOK_MCP_TENANT_ID="consumers"
```

No Windows (PowerShell): `$env:OUTLOOK_MCP_CLIENT_ID = "..."`.

Há um modelo com todas as variáveis em `.env.example` — copie para `.env`
(já ignorado pelo git) e preencha, ou use como referência.
Para não repetir isso toda vez, você pode salvar essas variáveis direto no
JSON de configuração do Claude Desktop (passo 5).

## 4. Primeira execução (autorizar a conta)

```bash
python server.py
```

Na primeira vez, o terminal vai mostrar algo como:

```
To sign in, use a web browser to open the page https://microsoft.com/devicelogin
and enter the code ABCD-1234 to authenticate.
```

Abra o link, cole o código, faça login com sua conta Outlook normal.
O token fica salvo em `token_cache.bin` (local, não sobe pro git) e é renovado
automaticamente nas próximas execuções.

## 5. Conectar ao Claude Desktop

Edite o arquivo de configuração do Claude Desktop:

- Windows: `%APPDATA%\Claude\claude_desktop_config.json`
- macOS: `~/Library/Application Support/Claude/claude_desktop_config.json`

Adicione:

```json
{
  "mcpServers": {
    "outlook-organizer": {
      "command": "C:\\caminho\\completo\\outlook-mcp\\.venv\\Scripts\\python.exe",
      "args": ["C:\\caminho\\completo\\outlook-mcp\\server.py"],
      "env": {
        "OUTLOOK_MCP_CLIENT_ID": "cole-o-client-id-aqui",
        "OUTLOOK_MCP_TENANT_ID": "consumers"
      }
    }
  }
}
```

Reinicie o Claude Desktop. As ferramentas devem aparecer disponíveis na conversa.

## Ferramentas disponíveis

| Ferramenta | O que faz |
|---|---|
| `list_folders` | Lista as pastas de e-mail |
| `list_recent_emails` | Lista e-mails de uma pasta (com `skip`, `unread_only` e escolha de campos) |
| `sender_stats` | Mapa da caixa agrupado por remetente, com contagens |
| `search_emails` | Busca por termo em toda a caixa |
| `get_email_content` | Retorna o corpo completo de um e-mail |
| `move_email` | Move um e-mail para outra pasta |
| `move_emails_batch` | Move vários e-mails de uma vez (lotes de 20 via `POST /$batch`) |
| `move_by_sender` | Move tudo de um remetente, com `dry_run` por padrão |
| `mark_as_read_batch` | Marca vários como lido/não lido em lote |
| `list_rules` | Lista as regras de caixa de entrada |
| `create_rule` | Cria regra que move e-mails para uma pasta |
| `delete_rule` | Remove uma regra pelo id |
| `mark_as_read` | Marca como lido/não lido |
| `flag_email` | Sinaliza um e-mail para acompanhamento |

## Rodando na rede local (Termux / Android / Docker)

Dá pra rodar o servidor 24/7 num celular Android antigo via Termux, fechado à
sua Wi-Fi, e conectar o Claude Desktop de outra máquina. Guia completo —
incluindo Docker, token de auth e proteção contra DNS rebinding — em
**[docs/rede-local.md](docs/rede-local.md)**.

## Trabalhando com volume alto

Para caixas com centenas de e-mails, a ordem que gasta menos contexto:

1. **`sender_stats`** primeiro — devolve o mapa da caixa (quem manda, quanto,
   quantos não lidos) em poucos KB. Uma listagem equivalente custaria ~7x mais.
2. **`move_by_sender`** com `dry_run=True` (o padrão) — confira a contagem
   antes de executar. Os IDs são filtrados no servidor e nunca trafegam.
3. Só então `list_recent_emails` com `fields=["id","de","assunto"]` e `skip`
   para o que sobrou. Cortar o `bodyPreview` reduz a resposta em ~75%.

`move_emails_batch` e `mark_as_read_batch` agrupam em lotes de 20 numa única
requisição. Atenção: **o move troca o ID do e-mail** — use os `id_novo` que
voltam em `detalhes_movidos` para qualquer passo seguinte. O `PATCH` do
`mark_as_read_batch` preserva os ids.

## Regras de caixa de entrada (opcional)

As ferramentas `list_rules` / `create_rule` / `delete_rule` exigem o escopo
`MailboxSettings.ReadWrite`, separado de `Mail.*`. Esse escopo dá acesso a
**todas** as configurações da caixa (respostas automáticas, fuso horário etc.),
por isso vem desligado por padrão. Para habilitar:

1. No Entra ID → seu app → **Permissões de APIs** → Microsoft Graph →
   **Permissões delegadas** → adicione `MailboxSettings.ReadWrite`
2. Defina `OUTLOOK_MCP_ENABLE_RULES=1` (no ambiente ou no `env` do
   `claude_desktop_config.json`)
3. Apague `token_cache.bin` e refaça o login para consentir o novo escopo

Sem isso, as três ferramentas devolvem um erro explicando o que falta — o
restante do servidor funciona normalmente.

Por decisão de projeto, `create_rule` **não** expõe as ações `forwardTo`,
`redirectTo` nem `permanentDelete` do Graph: regra de encaminhamento automático
é o vetor clássico de vazamento de e-mail, e exclusão permanente destrói
mensagem sem passar pela lixeira.

## Segurança

- `token_cache.bin` contém tokens de acesso à sua conta — não compartilhe nem
  suba para repositórios públicos.
- O app só tem os escopos que você concedeu (Mail.Read/ReadWrite/Send) — não
  tem acesso a outros dados do Microsoft 365.
- O `.gitignore` do projeto já cobre `token_cache.bin`, `.env` e `.venv/`.
  Confira com `git status` antes do primeiro commit.
