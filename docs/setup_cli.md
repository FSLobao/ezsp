# Configuração passo a passo para `Sites.Selected` — via Azure CLI e PowerShell

Este guia usa exclusivamente ferramentas de linha de comando: **Azure CLI** e **PowerShell com o módulo Microsoft Graph SDK**.

> Para instruções equivalentes usando o portal do Azure e o navegador, veja [setup_portal.md](setup_portal.md).

## Pré-requisitos

### Azure CLI — necessário para as Etapas 1–4

**Windows:**
```bash
winget install Microsoft.AzureCLI
# ou download direto em https://aka.ms/installazurecliwindows
```

**macOS:**
```bash
brew install azure-cli
```

**Linux (Ubuntu/Debian):**
```bash
curl -sL https://aka.ms/InstallAzureCLIDeb | sudo bash
```

Verificar instalação:
```bash
az version
```

### Microsoft Graph PowerShell SDK — alternativa para as Etapas 1–3

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```

---

## Quem faz o quê — visão geral

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

## Etapa 1 — Criar o registro de aplicativo e segredo do cliente

> 🔧 **Equipe de desenvolvimento**

### Azure CLI

```bash
az login

# Criar registro de aplicativo
APP_OBJECT_ID=$(az ad app create \
  --display-name "MSGraphClient-SharePoint" \
  --sign-in-audience "AzureADMyOrg" \
  --query id -o tsv)

AZURE_CLIENT_ID=$(az ad app show --id $APP_OBJECT_ID --query appId -o tsv)
AZURE_TENANT_ID=$(az account show --query tenantId -o tsv)

echo "AZURE_CLIENT_ID=$AZURE_CLIENT_ID"
echo "AZURE_TENANT_ID=$AZURE_TENANT_ID"

# Criar segredo do cliente (válido por 2 anos)
AZURE_CLIENT_SECRET=$(az ad app credential reset \
  --id $APP_OBJECT_ID \
  --display-name "MSGraphClient-secret" \
  --years 2 \
  --query password -o tsv)

echo "AZURE_CLIENT_SECRET=$AZURE_CLIENT_SECRET"
```

### PowerShell

```powershell
Connect-MgGraph -Scopes "Application.ReadWrite.All"

# Criar registro de aplicativo
$app = New-MgApplication -DisplayName "MSGraphClient-SharePoint" `
  -SignInAudience "AzureADMyOrg"

$AZURE_CLIENT_ID = $app.AppId
$AZURE_TENANT_ID = (Get-MgOrganization).Id

Write-Host "AZURE_CLIENT_ID=$AZURE_CLIENT_ID"
Write-Host "AZURE_TENANT_ID=$AZURE_TENANT_ID"

# Criar segredo do cliente (válido por 2 anos)
$secret = Add-MgApplicationPassword -ApplicationId $app.Id `
  -PasswordCredential @{
    DisplayName = "MSGraphClient-secret"
    EndDateTime = (Get-Date).AddYears(2)
  }

$AZURE_CLIENT_SECRET = $secret.SecretText
Write-Host "AZURE_CLIENT_SECRET=$AZURE_CLIENT_SECRET"
```

> ⚠️ O segredo é exibido apenas neste momento. Copie e armazene com segurança (ex: Azure Key Vault).

---

## Etapa 2 — Adicionar `Sites.Selected` e conceder consentimento de administrador

> 🔧 **Equipe de desenvolvimento** adiciona a permissão. 🔑 **Administrador Entra** concede consentimento.

> ⚠️ O consentimento requer **Administrador global** ou **Administrador de função com privilégios** (*Privileged Role Administrator*) — a função Administrador de Aplicativos não é suficiente para permissões de aplicativo do Microsoft Graph.

> Use `Sites.Selected` tanto para apps `app_only` quanto para apps `delegated`. O que muda é o tipo do consentimento: `Role` para `app_only` e `Scope` para `delegated`. Não são usadas permissões amplas como `Sites.Read.All` ou `Sites.ReadWrite.All`.

### Azure CLI

```bash
# Resolver dinamicamente o ID de Sites.Selected no Microsoft Graph
APP_ONLY_SITES_SELECTED_ID=$(az ad sp show \
  --id 00000003-0000-0000-c000-000000000000 \
  --query "appRoles[?value=='Sites.Selected' && contains(allowedMemberTypes, 'Application')] | [0].id" \
  -o tsv)

DELEGATED_SITES_SELECTED_ID=$(az ad sp show \
  --id 00000003-0000-0000-c000-000000000000 \
  --query "oauth2PermissionScopes[?value=='Sites.Selected'] | [0].id" \
  -o tsv)

# Para app_only (client credentials)
az ad app permission add \
  --id $APP_OBJECT_ID \
  --api 00000003-0000-0000-c000-000000000000 \
  --api-permissions "${APP_ONLY_SITES_SELECTED_ID}=Role"

# Para delegated (authorization code)
az ad app permission add \
  --id $APP_OBJECT_ID \
  --api 00000003-0000-0000-c000-000000000000 \
  --api-permissions "${DELEGATED_SITES_SELECTED_ID}=Scope"

# Conceder consentimento de administrador
az ad app permission admin-consent --id $APP_OBJECT_ID
```

### PowerShell

```powershell
# Obter service principal do Microsoft Graph
$graphSP = Get-MgServicePrincipal -Filter "AppId eq '00000003-0000-0000-c000-000000000000'"
$sitesSelectedAppPermission = $graphSP.AppRoles | Where-Object {
  $_.Value -eq "Sites.Selected" -and $_.AllowedMemberTypes -contains "Application"
} | Select-Object -First 1
$sitesSelectedDelegatedPermission = $graphSP.Oauth2PermissionScopes | Where-Object {
  $_.Value -eq "Sites.Selected"
} | Select-Object -First 1

# Adicionar permissão ao registro de aplicativo para app_only
$requiredAccess = @{
  ResourceAppId  = "00000003-0000-0000-c000-000000000000"
  ResourceAccess = @(
    @{ Id = $sitesSelectedAppPermission.Id; Type = "Role" }
  )
}
Update-MgApplication -ApplicationId $app.Id `
  -RequiredResourceAccess @($requiredAccess)

# Adicionar permissão ao registro de aplicativo para delegated
$requiredAccessDelegated = @{
  ResourceAppId  = "00000003-0000-0000-c000-000000000000"
  ResourceAccess = @(
    @{ Id = $sitesSelectedDelegatedPermission.Id; Type = "Scope" }
  )
}
Update-MgApplication -ApplicationId $app.Id `
  -RequiredResourceAccess @($requiredAccessDelegated)

# Criar service principal para o aplicativo (necessário para o consentimento)
$sp = New-MgServicePrincipal -AppId $AZURE_CLIENT_ID

# Conceder consentimento de administrador para app_only
New-MgServicePrincipalAppRoleAssignment `
  -ServicePrincipalId $sp.Id `
  -PrincipalId        $sp.Id `
  -ResourceId         $graphSP.Id `
  -AppRoleId          $sitesSelectedAppPermission.Id

# Conceder consentimento de administrador para delegated
New-MgOauth2PermissionGrant `
  -ClientId    $sp.Id `
  -ConsentType "AllPrincipals" `
  -ResourceId  $graphSP.Id `
  -Scope       "Sites.Selected"
```

> `Sites.Selected` não concede acesso a nenhum site ainda — isso acontece na Etapa 4.

---

## Etapa 3 — Descobrir o ID do site do SharePoint

> 🔧 **Equipe de desenvolvimento**

Esta etapa usa um token delegado da sua conta de usuário pessoal apenas para consultar o ID do site. Este token não é usado pelo aplicativo e não concede nada ao registro de aplicativo.

### Azure CLI

```bash
# Obter token delegado (usa sua conta de usuário, não as credenciais do aplicativo)
az login
TOKEN=$(az account get-access-token \
  --resource https://graph.microsoft.com \
  --query accessToken -o tsv)

# Consultar o site — substitua <hostname> e <site-path>
# Exemplo: contoso.sharepoint.com e sites/ProjectAlpha
curl -s -H "Authorization: Bearer $TOKEN" \
  "https://graph.microsoft.com/v1.0/sites/<hostname>:/sites/<site-path>" \
  | python -m json.tool
```

Copie o campo `id` da resposta → `SHAREPOINT_SITE_ID`.

### PowerShell

```powershell
Connect-MgGraph -Scopes "Sites.Read.All"

# Substitua <site-path> pelo caminho do site
$site = Get-MgSite -Search "<site-path>"
$SHAREPOINT_SITE_ID = $site.Id
Write-Host "SHAREPOINT_SITE_ID=$SHAREPOINT_SITE_ID"
```

Formato esperado: `contoso.sharepoint.com,xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx,yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy`

> Compartilhe este valor com o Administrador SP antes de prosseguir para a Etapa 4.

---

## Etapa 4 — Inscrever o site para o aplicativo

> 🛡️ **Administrador SP** (administrador do tenant do SharePoint)

> ⚠️ A conta executando esta etapa deve ter a função **Administrador do SharePoint** ou **Administrador Global**. Um administrador de coleção de sites / proprietário de site não pode chamar `POST /sites/{site-id}/permissions`.

```bash
az login  # faça login com a conta do Administrador SP
TOKEN=$(az account get-access-token \
  --resource https://graph.microsoft.com \
  --query accessToken -o tsv)
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
        "displayName": "MSGraphClient-SharePoint"
      }
    }]
  }'
```

**Conceder acesso leitura + escrita** (escrita implicitamente inclui leitura):

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
        "displayName": "MSGraphClient-SharePoint"
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

**Revogar uma concessão** (use o `id` retornado pela verificação acima):

```bash
curl -s -X DELETE \
  -H "Authorization: Bearer $TOKEN" \
  "https://graph.microsoft.com/v1.0/sites/${SITE_ID}/permissions/<permission-id>"
```

---

## Etapa 5 — Descobrir o ID da drive e ID da lista

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

Localize a entrada cujo `name` corresponde à sua biblioteca de documentos e copie seu `id` → `SHAREPOINT_DRIVE_ID`.

**Encontrar o ID da lista:**

```bash
curl -s -H "Authorization: Bearer $TOKEN" \
  "https://graph.microsoft.com/v1.0/sites/${SHAREPOINT_SITE_ID}/lists" \
  | python -m json.tool
```

Localize a lista desejada e copie seu `id` → `SHAREPOINT_LIST_ID`.

---

## Etapa 6 — Configurar `.env` e executar

> 🔧 **Equipe de desenvolvimento**

```bash
uv sync
cp .env.example .env
```

Edite `.env` com todos os valores coletados:

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
uv run examples/example_drive_folder_operations.py
uv run examples/example_list_get.py
```
