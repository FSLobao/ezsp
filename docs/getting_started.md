# Guia de Início

Este guia é dividido em duas partes:

1. [**Permissões — conceitos, funções e escopo de acesso**](#permissões--conceitos-funções-e-escopo-de-acesso) — quais são as permissões disponíveis, quem é responsável por cada uma e que conteúdo elas concedem acesso.
2. [**Configuração passo a passo para `Sites.Selected`**](#configuração-passo-a-passo-para-sitesselected) — as ações concretas que cada função deve executar para configurar e executar o projeto.

---

## Permissões: conceitos, funções e escopo de acesso

### Funções envolvidas

Três funções distintas participam da configuração de acesso de aplicativo ao SharePoint
através da Microsoft Graph API. As etapas na Parte 2 são marcadas com o distintivo da
função responsável.

| Distintivo | Função | Descrição |
|---|---|---|
| 🔧 **Equipe de desenvolvimento** | Desenvolvedor de aplicativo | Cria o registro de aplicativo e escreve o código. |
| 🔑 **Administrador Entra** | Administrador do Azure AD / Entra ID | Tem a função **Administrador de Aplicativos** (ou Administrador Global). Concede consentimento de administrador para permissões de API no portal do Azure. |
| 🛡️ **Administrador SP** | Administrador do tenant (locatário) do SharePoint | Tem a função **Administrador do SharePoint** no Centro de Administração Microsoft 365. **Isso não é o mesmo que um administrador de coleção de sites (proprietário de site).** Um proprietário de site gerencia usuários dentro de um site, mas não pode conceder acesso à API do Graph — isso requer direitos de administrador do SharePoint no nível do tenant. |

### Permissões disponíveis da API Microsoft Graph para SharePoint

A Microsoft Graph API oferece dois níveis de permissões de aplicativo para
acesso ao SharePoint:

| Permissão | Escopo | Observações |
|---|---|---|
| `Sites.Read.All` | Todos os sites no tenant | Somente leitura, sem restrição por site |
| `Sites.ReadWrite.All` | Todos os sites no tenant | Leitura e escrita, sem restrição por site |
| `Sites.Selected` | Apenas sites explicitamente inscritos | Privilégio mínimo; recomendado para todas as implantações não triviais |

Permissões em todo o tenant (`Sites.Read.All`, `Sites.ReadWrite.All`) são simples de
configurar, mas concedem ao aplicativo acesso a cada site do SharePoint na
organização. Uma credencial comprometida exporia todo conteúdo em todo o tenant.
Elas **não são usadas neste projeto** e são mencionadas aqui apenas para
completude.

**`Sites.Selected` é a abordagem usada em todo este guia.**

### O que `Sites.Selected` concede acesso

`Sites.Selected` por si só não concede ao aplicativo acesso a nada. O
administrador do tenant (locatário) do SharePoint deve inscrever separadamente cada site que o aplicativo tem permissão de acessar, escolhendo nível `read` (leitura) ou `write` (escrita).

Depois que um site é inscrito, a concessão cobre **todo conteúdo dentro dessa
coleção de sites**:

| Tipo de conteúdo | Concessão de leitura | Concessão de escrita |
|---|:---:|:---:|
| Bibliotecas de documentos — enumerar pastas e arquivos | ✅ | ✅ |
| Bibliotecas de documentos — baixar conteúdo de arquivo | ✅ | ✅ |
| Bibliotecas de documentos — enviar ou sobrescrever arquivos | ❌ | ✅ |
| Metadados de arquivo (nome, tamanho, timestamps, URL) | ✅ | ✅ |
| Listas — ler itens e valores de campo | ✅ | ✅ |
| Listas — criar ou atualizar itens | ❌ | ✅ |

> **Limite de granularidade:** `Sites.Selected` é a permissão de aplicativo mais refinada disponível
> na Microsoft Graph API. Não é possível restringir o acesso a uma única
> biblioteca de documentos ou lista *dentro* de um site via permissões do Graph isoladamente.
> Se isolamento de subsite for necessário, aplique-o na camada de aplicativo validando
> o ID da drive ou ID da lista antes de agir.

Os papéis `read` (leitura) e `write` (escrita) em `Sites.Selected` mapeiam para o mesmo
nível de acesso subjacente que `Sites.Read.All` e `Sites.ReadWrite.All` — a única
diferença é escopo: o acesso é restrito a sites explicitamente inscritos.

Conceda `read` quando o aplicativo precisar apenas ler. Conceda `write` apenas a
sites que o requeiram. Use registros de aplicativo separados se diferentes sites precisarem
de níveis de acesso diferentes.

### Fluxo de autenticação: credenciais do cliente vs. autenticação delegada (usuário)

Este projeto usa o fluxo OAuth 2.0 de **credenciais do cliente (client credentials)**, onde o aplicativo
se autentica como ele mesmo (o registro de aplicativo) sem contexto de usuário:

- ✅ **Execução sem supervisão** — não é necessário login do usuário; adequado para trabalhos em lote, serviços em background ou tarefas agendadas.
- ✅ **Configuração simples** — um único ID do cliente e segredo.
- ❌ **Trilha de auditoria** — logs de atividade do SharePoint registram ações pela identidade do aplicativo, não pelo usuário executando o código.

**Alternativa: autenticação delegada** — se você precisar rastrear ações pelo
usuário executando o aplicativo, você pode mudar para o **fluxo de código de autorização**, onde
um usuário faz login interativamente:

- ✅ **Trilha de auditoria do usuário** — logs do SharePoint identificam o usuário específico que
  executou cada ação.
- ❌ **Requer interação do usuário** — o aplicativo não pode executar sem supervisão; alguém deve
  fazer login cada vez.

Ambos os fluxos usam o mesmo modelo de permissão `Sites.Selected` e concedem o mesmo
acesso ao conteúdo do SharePoint. A diferença é **quem o aplicativo representa**
(ele mesmo vs. um usuário) e como essa identidade aparece nos logs de auditoria.

Mudar para autenticação delegada requer:

1. Mudança da aquisição de token em `src/msgraphtest/auth.py` do endpoint de
   credenciais do cliente para o fluxo de código de autorização (usando uma biblioteca como
   [MSAL for Python](https://github.com/AzureAD/microsoft-authentication-library-for-python)
   com `acquire_token_interactive()`).
2. Adição de uma **URI de Redirecionamento** no registro de aplicativo (ex:
   `http://localhost:8000`) para tratar o callback OAuth.
3. Usuários fazendo login quando o aplicativo executa.

Para a maioria dos cenários, **credenciais do cliente (abordagem atual) é mais simples**. Use
autenticação delegada apenas se seus requisitos de governança ou conformidade mandatarem uma
trilha de auditoria vinculando cada ação do SharePoint a um usuário nomeado.

## Configuração passo a passo para `Sites.Selected`

### Quem faz o quê — visão geral

| Etapa | Ação | 🔧 Desenvolvimento | 🔑 Administrador Entra | 🛡️ Administrador SP |
|---|---|:---:|:---:|:---:|
| 1 | Criar registro de aplicativo e segredo do cliente | ✅ | | |
| 2 | Adicionar permissão `Sites.Selected` | ✅ | | |
| 2 | Conceder consentimento de administrador | | ✅ | |
| 3 | Descobrir o ID do site do SharePoint | ✅ | | |
| 4 | Inscrever o site para o aplicativo | | | ✅ |
| 5 | Descobrir ID da drive e ID da lista | ✅ | | |
| 6 | Configurar `.env` e executar | ✅ | | |

---

### Etapa 1 — Criar o registro de aplicativo e segredo do cliente

> 🔧 **Equipe de desenvolvimento**

1. Abra o [portal do Azure](https://portal.azure.com) → **Microsoft Entra ID** → **Registros de app** → **Novo registro**.
2. **Nome**: escolha um nome descritivo, ex: `MSGraphTest-SharePoint`.
3. **Tipos de conta suportados**: selecione **Contas neste diretório organizacional apenas**.
4. Deixe **URI de Redirecionamento** em branco (fluxo de credenciais do cliente — sem login do usuário).
5. Clique em **Registrar** e anote:
   - **ID do Aplicativo (cliente)** → `AZURE_CLIENT_ID`
   - **ID do Diretório (tenant)** → `AZURE_TENANT_ID`
6. Vá para **Certificados e segredos** → **Segredos do cliente** → **Novo segredo do cliente**.
7. Defina uma descrição e uma expiração alinhada com sua política de rotação (máximo 24 meses).
8. Clique em **Adicionar** e **copie imediatamente** o valor do segredo → `AZURE_CLIENT_SECRET`.

> ⚠️ O valor do segredo é mostrado apenas uma vez. Armazene-o com segurança (ex: Azure Key Vault). Se você navegar para longe sem copiá-lo, delete e recrie-o.

---

### Etapa 2 — Adicionar `Sites.Selected` e conceder consentimento de administrador

> 🔧 **Equipe de desenvolvimento** adiciona a permissão. 🔑 **Administrador Entra** concede consentimento.

1. No registro de aplicativo vá para **Permissões de API** → **Adicionar uma permissão** → **Microsoft Graph** → **Permissões de aplicativo**.
2. Pesquise e adicione **`Sites.Selected`** apenas.
3. Clique em **Conceder consentimento de administrador para \<tenant\>** (requer Administrador Entra) e confirme.

> `Sites.Selected` não concede acesso a site ainda — isso acontece na Etapa 4.

---

### Etapa 3 — Descobrir o ID do site do SharePoint

> 🔧 **Equipe de desenvolvimento**

Esta etapa usa um **token delegado de curta duração** (sua conta de usuário pessoal)
apenas para procurar a string do ID do site. Este token é independente das
credenciais do aplicativo: não concede nada ao registro de aplicativo, não deixa
estado duradouro e não é usado em lugar nenhum depois desta etapa. O próprio aplicativo
sempre autentica com um token de credenciais do cliente separado (identidade do app) adquirido
na Etapa 5.

Obtenha o token delegado usando a Azure CLI:

```bash
az login
TOKEN=$(az account get-access-token --resource https://graph.microsoft.com --query accessToken -o tsv)
```

Consulte o site:

```bash
# Substitua <hostname> e <site-path> com seus valores
# Exemplo: contoso.sharepoint.com e sites/ProjectAlpha
curl -s -H "Authorization: Bearer $TOKEN" \
  "https://graph.microsoft.com/v1.0/sites/<hostname>:/sites/<site-path>" \
  | python -m json.tool
```

Copie o campo `id` da resposta → `SHAREPOINT_SITE_ID`.  
Formato: `contoso.sharepoint.com,xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx,yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy`

Compartilhe este valor com o Administrador SP antes de prosseguir para a Etapa 4.

#### Alternativas para obter o token (sem Azure CLI)

Se você não tem Azure CLI instalado ou prefere outra abordagem, aqui estão as alternativas:

##### Opção 1: Microsoft Graph Explorer (mais fácil, browser)

1. Abra [Microsoft Graph Explorer](https://developer.microsoft.com/graph/graph-explorer)
2. Clique em **Sign in** (Entrar) e autentique com sua conta do Microsoft 365
3. Cole a URL de consulta no campo de entrada:
   ```
   https://graph.microsoft.com/v1.0/sites/<hostname>:/sites/<site-path>
   ```
4. Clique em **Run query** (Executar consulta)
5. O token já é adquirido automaticamente — apenas copie o campo `id` da resposta

**Vantagens:** Sem instalação necessária, interface visual, ideal para teste rápido.

##### Opção 2: PowerShell com Microsoft Graph SDK

```powershell
# Instale o módulo (primeira vez apenas)
Install-Module Microsoft.Graph -Scope CurrentUser

# Conecte com sua conta
Connect-MgGraph -Scopes "Sites.Read.All"

# Consulte o site
Get-MgSite -Search "<site-path>" | Select-Object Id
```

**Vantagens:** Nativo do Windows, script reutilizável.

##### Opção 3: Python com MSAL (programático)

Salve este script como `get_site_id.py`:

```python
import msal
import requests

# Configuração
TENANT_ID = "<seu-tenant-id>"
# Você precisará fazer login interativamente
CLIENT_ID = "04b07795-8ddb-461a-bbee-02f9e1bf7b46"  # ID público do Azure CLI

app = msal.PublicClientApplication(
    client_id=CLIENT_ID,
    authority=f"https://login.microsoftonline.com/{TENANT_ID}"
)

# Adquire token interativamente
result = app.acquire_token_interactive(
    scopes=["https://graph.microsoft.com/.default"]
)

if "access_token" in result:
    token = result["access_token"]
    
    # Consulte o site
    hostname = "contoso.sharepoint.com"
    site_path = "sites/ProjectAlpha"
    
    response = requests.get(
        f"https://graph.microsoft.com/v1.0/sites/{hostname}:/sites/{site_path}",
        headers={"Authorization": f"Bearer {token}"}
    )
    
    print(response.json()["id"])
else:
    print("Falha ao obter token:", result.get("error_description"))
```

Execute:
```bash
python get_site_id.py
```

**Vantagens:** Controle total, pode ser integrado em scripts Python.

##### Opção 4: Postman (cliente HTTP visual)

1. Baixe [Postman](https://www.postman.com/downloads/)
2. Crie uma nova requisição GET
3. URL: `https://graph.microsoft.com/v1.0/sites/<hostname>:/sites/<site-path>`
4. Abra a aba **Authorization** → Tipo: **OAuth 2.0**
5. Clique em **Get New Access Token** e autentique
6. Clique em **Send** e copie o campo `id` da resposta

**Vantagens:** Interface visual, útil se você trabalha com APIs frequentemente.

##### Instalando Azure CLI (se preferir usar)

**Windows:**
```bash
choco install azure-cli
# ou download em https://aka.ms/installazurecliwindows
```

**macOS:**
```bash
brew install azure-cli
```

**Linux (Ubuntu/Debian):**
```bash
curl -sL https://aka.ms/InstallAzureCLIDeb | sudo bash
```

**Verificar instalação:**
```bash
az version
```

---

### Etapa 4 — Inscrever o site para o aplicativo

> 🛡️ **Administrador SP** (administrador do tenant do SharePoint)

> ⚠️ A conta executando esta etapa deve ter a função **Administrador do SharePoint**
> ou **Administrador Global** no Microsoft 365. Um administrador de coleção de sites /
> proprietário de site não pode chamar `POST /sites/{site-id}/permissions`.

Obtenha um token de administrador:

```bash
az login  # faça login com a conta do Administrador SP
TOKEN=$(az account get-access-token --resource https://graph.microsoft.com --query accessToken -o tsv)
SITE_ID="<SHAREPOINT_SITE_ID da Etapa 3>"
APP_ID="<AZURE_CLIENT_ID da Etapa 1>"
```

**Conceder acesso somente leitura:**

```bash
curl -s -X POST \
  -H "Authorization: Bearer $TOKEN" \
  -H "Content-Type: application/json" \
  "https://graph.microsoft.com/v1.0/sites/${SITE_ID}/permissions" \
  -d '{
    "roles": ["read"],
    "grantedToIdentities": [{
      "application": {
        "id": "'"${APP_ID}"'",
        "displayName": "MSGraphTest-SharePoint"
      }
    }]
  }'
```

**Conceder acesso leitura + escrita** (substitua `"read"` com `"write"` — escrita implicitamente inclui leitura):

```bash
curl -s -X POST \
  -H "Authorization: Bearer $TOKEN" \
  -H "Content-Type: application/json" \
  "https://graph.microsoft.com/v1.0/sites/${SITE_ID}/permissions" \
  -d '{
    "roles": ["write"],
    "grantedToIdentities": [{
      "application": {
        "id": "'"${APP_ID}"'",
        "displayName": "MSGraphTest-SharePoint"
      }
    }]
  }'
```

**Verificar a concessão:**

```bash
curl -s -H "Authorization: Bearer $TOKEN" \
  "https://graph.microsoft.com/v1.0/sites/${SITE_ID}/permissions" \
  | python -m json.tool
```

**Revogar uma concessão** (se necessário — use o `id` retornado pela chamada de verificação acima):

```bash
curl -s -X DELETE \
  -H "Authorization: Bearer $TOKEN" \
  "https://graph.microsoft.com/v1.0/sites/${SITE_ID}/permissions/<permission-id>"
```

---

### Etapa 5 — Descobrir o ID da drive e ID da lista

> 🔧 **Equipe de desenvolvimento** — usa o token de credenciais do cliente do próprio aplicativo.

```bash
TOKEN=$(curl -s -X POST \
  "https://login.microsoftonline.com/${AZURE_TENANT_ID}/oauth2/v2.0/token" \
  -d "grant_type=client_credentials" \
  -d "client_id=${AZURE_CLIENT_ID}" \
  -d "client_secret=${AZURE_CLIENT_SECRET}" \
  -d "scope=https://graph.microsoft.com/.default" \
  | python -c "import sys,json; print(json.load(sys.stdin)['access_token'])")
```

**Encontrar o ID da drive:**

```bash
curl -s -H "Authorization: Bearer $TOKEN" \
  "https://graph.microsoft.com/v1.0/sites/${SHAREPOINT_SITE_ID}/drives" \
  | python -m json.tool
```

Escolha a entrada cujo `name` corresponde à sua biblioteca de documentos e copie seu `id` → `SHAREPOINT_DRIVE_ID`.

**Encontrar o ID da lista:**

```bash
curl -s -H "Authorization: Bearer $TOKEN" \
  "https://graph.microsoft.com/v1.0/sites/${SHAREPOINT_SITE_ID}/lists" \
  | python -m json.tool
```

Escolha a lista com a qual você deseja trabalhar e copie seu `id` → `SHAREPOINT_LIST_ID`.

---

### Etapa 6 — Configurar `.env` e executar

> 🔧 **Equipe de desenvolvimento**

```bash
uv sync
cp .env.example .env
```

Edite `.env` com todos os valores coletados acima:

```ini
AZURE_TENANT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
AZURE_CLIENT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
AZURE_CLIENT_SECRET=<your-secret-value>
SHAREPOINT_SITE_ID=contoso.sharepoint.com,<site-guid>,<web-guid>
SHAREPOINT_DRIVE_ID=b!<drive-id>
SHAREPOINT_LIST_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
```

Verificar conectividade:

```bash
uv run examples/example_drive_list.py
uv run examples/example_list_get.py
```
