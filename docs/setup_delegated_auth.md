# Configuração passo a passo para autenticação delegada (usuário)

Este guia configura o acesso ao SharePoint via Microsoft Graph API usando o **fluxo de código de autorização OAuth 2.0** (autenticação delegada), onde o aplicativo age em nome de um usuário autenticado.

> **Diferença em relação à autenticação por credenciais do cliente:**
>
> | | Credenciais do cliente (app-only) | Autenticação delegada (usuário) |
> |---|---|---|
> | Quem executa | O próprio aplicativo | Um usuário autenticado |
> | Login necessário | Não — execução sem supervisão | Sim — a cada execução |
> | Trilha de auditoria | Identidade do aplicativo | Identidade do usuário individual |
> | Permissões necessárias | Permissões de aplicativo (`Role`) | Permissões delegadas (`Scope`) |
> | Etapa 4 (inscrever site) | Obrigatória | Obrigatória |
>
> Use autenticação delegada quando seus requisitos de governança ou conformidade exigirem trilha de auditoria vinculando cada ação do SharePoint a um usuário nomeado.

Para configuração com credenciais do cliente (sem login de usuário), veja [setup_portal.md](setup_portal.md) ou [setup_cli.md](setup_cli.md).

---

## Quem faz o quê — visão geral

| Etapa | Ação | 🔧 Desenvolvimento | 🔑 Administrador Entra | 🛡️ Administrador SP |
|---|---|:---:|:---:|:---:|
| 1 | Criar registro de aplicativo com URI de redirecionamento | ✅ | | |
| 2 | Adicionar `Sites.Selected` como permissão delegada | ✅ | | |
| 3 | Conceder consentimento de administrador | | ✅ | |
| 4 | Inscrever o site via `POST /sites/{id}/permissions` | | | ✅ |
| 5 | Adaptar `auth.py` para o fluxo interativo | ✅ | | |
| 6 | Configurar `.env` e executar | ✅ | | |

> **Nota:** Neste projeto, a autenticação delegada também é sempre restringida a um site concreto. O usuário autenticado precisa ter acesso ao site e a aplicação precisa ter `Sites.Selected` consentido e uma concessão explícita no site. Não são usadas permissões amplas como `Sites.Read.All` ou `Sites.ReadWrite.All`.

---

## Etapa 1 — Criar o registro de aplicativo com URI de redirecionamento

> 🔧 **Equipe de desenvolvimento**

1. Abra o [portal do Azure](https://portal.azure.com) → **Microsoft Entra ID** → **Registros de app** → **Novo registro**.
2. **Nome**: escolha um nome descritivo, ex: `MSGraphTest-Delegated`.
3. **Tipos de conta suportados**: selecione **Contas neste diretório organizacional apenas**.
4. Em **URI de Redirecionamento**, selecione o tipo **Web** e informe `http://localhost:8000` (para execução local).
5. Clique em **Registrar** e anote:
   - **ID do Aplicativo (cliente)** → `AZURE_CLIENT_ID`
   - **ID do Diretório (tenant)** → `AZURE_TENANT_ID`
6. Vá para **Certificados e segredos** → **Segredos do cliente** → **Novo segredo do cliente**.
7. Defina uma descrição e uma expiração alinhada com sua política de rotação (máximo 24 meses).
8. Clique em **Adicionar** e **copie imediatamente** o valor do segredo → `AZURE_CLIENT_SECRET`.

> ⚠️ O valor do segredo é mostrado apenas uma vez. Armazene-o com segurança. Se você navegar para longe sem copiá-lo, delete e recrie-o.

---

## Etapa 2 — Adicionar `Sites.Selected` como permissão delegada

> 🔧 **Equipe de desenvolvimento**

As permissões delegadas controlam o que o aplicativo pode fazer **em nome do usuário autenticado**. Neste projeto, a permissão delegada usada é `Sites.Selected`, para que o aplicativo só consiga operar em sites explicitamente inscritos.

1. No registro de aplicativo, vá para **Permissões de API** → **Adicionar uma permissão** → **Microsoft Graph** → **Permissões delegadas**.
2. Adicione **`Sites.Selected`** e, se necessário para evitar novo login frequente, também **`offline_access`**.

| Permissão delegada | Necessária para |
|---|---|
| `Sites.Selected` | Restringir o aplicativo aos sites explicitamente inscritos |
| `offline_access` | Obter refresh token para renovar a sessão sem novo login |

> `Sites.Selected` não concede acesso a nenhum site por si só. A aplicação só conseguirá ler ou escrever no site depois da inscrição da Etapa 4 e sempre limitada ao que o usuário autenticado já puder fazer nesse mesmo site.

3. Clique em **Adicionar permissões**.

---

## Etapa 3 — Consentimento de administrador

> 🔑 **Administrador Entra**

Para que o aplicativo use `Sites.Selected` de forma controlada em todo o tenant, conceda consentimento administrativo ao escopo delegado:

1. No registro de aplicativo, vá para **Permissões de API**.
2. Clique em **Conceder consentimento de administrador para \<tenant\>** e confirme.

> Este consentimento não concede acesso a todos os sites. Ele apenas autoriza o uso do escopo `Sites.Selected`; o acesso aos dados continua dependente da inscrição do site na Etapa 4.

---

## Etapa 4 — Inscrever o site para o aplicativo

> 🛡️ **Administrador SP** (administrador do tenant do SharePoint)

Siga a mesma inscrição por site descrita na Etapa 4 de [setup_portal.md](setup_portal.md) ou [setup_cli.md](setup_cli.md):

1. Obtenha o `SHAREPOINT_SITE_ID` do site que a aplicação poderá acessar.
2. Execute `POST /sites/{site-id}/permissions` com o `AZURE_CLIENT_ID` da aplicação e o papel `read` ou `write`.
3. Guarde o `site_id` para configuração local e para os scripts de automação em lote.

> O acesso efetivo em runtime será a interseção entre: 1) o papel atribuído ao aplicativo nesse site e 2) as permissões do usuário autenticado no próprio site.

---

## Etapa 5 — Habilitar o modo delegado no pacote

> 🔧 **Equipe de desenvolvimento**

Com a configuração do tenant concluída nas etapas anteriores, configure o modo
delegado via `.env`:

```env
GRAPH_AUTH_MODE=delegated
AZURE_REDIRECT_URI=http://localhost
GRAPH_DELEGATED_LOGIN_MODE=interactive
# Opcional
# GRAPH_DELEGATED_SCOPES=https://graph.microsoft.com/Sites.Selected offline_access openid profile
```

Depois, execute um exemplo usando a API class-based:

```bash
uv run examples/example_delegated_site_contents.py
```

Você também pode usar o modo `device_code` (útil quando o redirect local não é
viável):

```env
GRAPH_DELEGATED_LOGIN_MODE=device_code
```

Em modo delegado, o `GraphClient` e o `GraphAuthenticator` funcionam com a mesma
superfície pública já usada no modo `client_credentials`; muda apenas a forma de
obtenção do token.

Exemplos adicionais ainda válidos:

```bash
uv run examples/example_drive_list.py
uv run examples/example_list_get.py
uv run notebooks/graph_auth_site_attributes.ipynb
```
