# Rodar na rede local (Termux / Android / Docker)

Modo HTTP: o servidor roda num aparelho dedicado (ex: celular Android antigo)
e o Claude Desktop de outra máquina se conecta pela Wi-Fi de casa.
Para o modo local normal (stdio), veja o [README](../README.md).

**Sobre "Docker no Android": não recomendo.** Docker de verdade precisa de um
daemon com privilégios de root e cgroups que o Android normal não expõe pra
apps sem root. Existem gambiarras (UserLAnd, chroot + qemu), mas são frágeis
e vão consumir bem mais bateria/RAM que o necessário num aparelho antigo. O
**Termux** entrega o que você quer de verdade — ambiente isolado, self-contained,
rodando 24/7 num aparelho dedicado — sem a dor de cabeça do Docker em Android.
Se seu aparelho for rooteado e você realmente quiser Docker, dá pra usar o
mesmo `server_http.py` dentro de um container Alpine/Debian normal; o
Dockerfile abaixo serve pra isso.

### Opção A — Termux (recomendada, sem root)

1. Instale o **Termux** pela F-Droid (a versão da Play Store está descontinuada):
   https://f-droid.org/packages/com.termux/
2. No Termux:
   ```bash
   pkg update && pkg upgrade
   pkg install python git
   termux-setup-storage   # opcional
   ```
3. Copie os arquivos deste projeto pro celular (via `git clone`, `termux-share`,
   ou um cabo/`adb push`) e instale as dependências:
   ```bash
   cd outlook-mcp
   pip install -r requirements.txt
   ```
4. Descubra o IP local do celular: `ifconfig wlan0` (ou Ajustes → Wi-Fi →
   detalhes da rede). Anote algo como `192.168.1.50`.
5. Configure as variáveis e suba o servidor:
   ```bash
   export OUTLOOK_MCP_CLIENT_ID="seu-client-id"
   export OUTLOOK_MCP_TENANT_ID="consumers"
   export OUTLOOK_MCP_AUTH_TOKEN="$(python3 -c 'import secrets;print(secrets.token_urlsafe(32))')"
   export OUTLOOK_MCP_ALLOWED_HOST="192.168.1.50:8787"
   echo "Guarde esse token: $OUTLOOK_MCP_AUTH_TOKEN"

   python server_http.py
   ```
   Na primeira execução, ele vai pedir a autorização via device code
   (mesmo fluxo do modo local — abra o link no navegador do celular ou de
   outro aparelho e autorize).
6. Pra manter rodando 24/7:
   - `termux-wake-lock` (evita o Android suspender o processo)
   - instale **Termux:Boot** (F-Droid) pra reiniciar o servidor automaticamente
     depois de reboot/queda de energia
   - ou use `pkg install tmux` e rode dentro de uma sessão `tmux` persistente

### Trancando pra rede local

- **Não** faça port-forward nem UPnP no roteador — isso é o que garante que só
  dispositivos dentro do seu Wi-Fi alcancem o servidor.
- `OUTLOOK_MCP_AUTH_TOKEN` garante que, mesmo dentro da rede, só quem tem o
  token consegue usar as ferramentas (protege contra outros dispositivos na
  mesma Wi-Fi, tipo IoT comprometido).
- `OUTLOOK_MCP_ALLOWED_HOST` trava contra ataques de DNS rebinding (sites
  maliciosos tentando usar seu navegador pra bater no servidor local).
- Se seu roteador suportar, isole o celular numa VLAN/rede de convidados
  separada da rede principal — camada extra opcional.

### Conectando o Claude Desktop ao servidor remoto

O `claude_desktop_config.json` só sabe rodar comandos locais (stdio), então
usamos a ponte `mcp-remote` (via `npx`, precisa de Node.js instalado na
máquina que roda o Claude Desktop):

```json
{
  "mcpServers": {
    "outlook-organizer": {
      "command": "npx",
      "args": [
        "-y", "mcp-remote",
        "http://192.168.1.50:8787/mcp",
        "--header", "Authorization:Bearer SEU_TOKEN_AQUI"
      ]
    }
  }
}
```

Troque `192.168.1.50` pelo IP do celular e `SEU_TOKEN_AQUI` pelo token gerado
no passo 5. Isso só funciona com o computador na mesma rede Wi-Fi do celular
(exatamente o comportamento "fechado à rede local").

### Opção B — Docker (só se o celular for rooteado)

O Dockerfile fica em [`docker/Dockerfile`](../docker/Dockerfile). O contexto de
build é a **raiz do projeto** (por causa do `outlook_mcp/`), então rode de lá
com `-f`:

```bash
docker build -f docker/Dockerfile -t outlook-mcp .
```

O `token_cache.bin` precisa existir antes de montar como volume (senão o Docker
cria um diretório no lugar do arquivo):

```bash
touch token_cache.bin
docker run -d --name outlook-mcp   -p 8787:8787   -e OUTLOOK_MCP_CLIENT_ID="seu-client-id"   -e OUTLOOK_MCP_AUTH_TOKEN="seu-token"   -e OUTLOOK_MCP_ALLOWED_HOST="192.168.1.50:8787"   -v "$(pwd)/token_cache.bin:/app/token_cache.bin"   outlook-mcp
```

Na primeira vez, rode sem `-d` (ou veja `docker logs -f outlook-mcp`) para pegar
o código do device login e autorizar a conta.
