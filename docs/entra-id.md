# Registrar o app no Entra ID (Azure AD)

O Graph API exige um app registrado, mesmo para uso 100% pessoal e gratuito.
Faça isso **uma vez**, antes de instalar as dependências.

1. Acesse https://entra.microsoft.com → **Aplicativos** → **Registros de aplicativo** → **Novo registro**.
2. Nome: `outlook-mcp-local` (ou o que preferir).
3. Tipos de conta com suporte: escolha **"Contas em qualquer diretório organizacional e contas Microsoft pessoais"**
   (é a opção que cobre `@outlook.com` / `@hotmail.com`).
4. Em **URI de redirecionamento**, escolha tipo **"Cliente público/nativo (móvel e desktop)"** e use:
   `http://localhost`
5. Após criar, copie o **Application (client) ID** — vai virar `OUTLOOK_MCP_CLIENT_ID`.
6. Em **Autenticação**, marque **"Permitir fluxos de cliente público"** = Sim.
7. Em **Permissões de API** → **Adicionar uma permissão** → **Microsoft Graph** → **Permissões delegadas**, adicione:
   - `Mail.Read`
   - `Mail.ReadWrite`
   - `Mail.Send` (opcional, remova de `outlook_mcp/auth.py` se não for usar)
   Permissões delegadas de Mail numa conta pessoal não exigem consentimento de administrador — você mesmo autoriza no primeiro login.

Depois de criar, volte para o [README](../README.md#2-instalar-dependências).
