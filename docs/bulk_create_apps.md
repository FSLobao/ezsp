# Bulk Create — Automação para Criar Múltiplas Aplicações Azure AD

Automação para criar múltiplas aplicações Azure AD em lote, gerando automaticamente segredos, adicionando permissões e concedendo consentimento de administrador.

---

## Seção 1: Objetivo e Alternativas

### Objetivo

Este utilitário automatiza a criação em lote de aplicações Azure AD para acesso a SharePoint via Microsoft Graph API com permissão `Sites.Selected`. Para cada aplicação, o script:

1. ✅ Cria o registro de aplicativo no Azure AD
2. ✅ Gera um segredo do cliente com data de expiração explícita
3. ✅ Adiciona a permissão `Sites.Selected`
4. ✅ Concede consentimento de administrador
5. ✅ Inscreve o site informado com papel `read` ou `write`
6. ✅ Retorna JSON com credenciais (`client_id`, `client_secret`, `tenant_id`)

**Entrada**: JSON array com configurações de aplicações
**Saída**: Mesmo JSON com credenciais adicionadas

---

### Alternativa 1: PowerShell (Recomendado) ⭐

**Comando:**
```powershell
.\src\bulkCreate\Bulk-CreateApps.ps1 -InputPath "apps.json"
```

| Aspecto | Vantagem | Desvantagem |
|--------|---------|-----------|
| **Instalação** | ✅ Nativo do Windows (apenas atualizar) | ❌ Requer PS 7+ |
| **Dependências** | ✅ Só 1 módulo PowerShell | ❌ Nenhum |
| **Ferramentas extras** | ✅ Nenhuma | ❌ Nenhuma |
| **Segurança** | ✅ Política de execução do Windows | ❌ Requer habilitação |
| **Sessão reutilizável** | ✅ Suporta `-SkipLogin` | ❌ Sim |
| **Performance** | ✅ Nativa do SO | ❌ ~5-10% mais rápido |

✅ **Quando usar**: Ambiente Windows, máxima simplicidade, mínimas dependências  
❌ **Quando evitar**: Ambiente Linux/macOS (use Python em vez disso)

---

### Alternativa 2: Python + Azure CLI

**Comando:**
```bash
python -m bulkCreate.bulk_create_apps apps.json --method cli
```

| Aspecto | Vantagem | Desvantagem |
|--------|---------|-----------|
| **Instalação** | ✅ Multiplataforma (Windows/Mac/Linux) | ❌ Requer Python + Azure CLI |
| **Dependências** | ✅ Azure CLI bem documentado | ❌ 2+ ferramentas externas |
| **Ferramentas extras** | ✅ Azure CLI multipropósito | ❌ Instalação separada |
| **Segurança** | ✅ Sem políticas de execução | ❌ Menos integrado |
| **Sessão reutilizável** | ✅ Suporta `--skip-login` | ❌ Sim |
| **Performance** | ✅ Bom desempenho | ❌ Subprocess overhead |

✅ **Quando usar**: Ambiente Linux/macOS, já tem Python + Azure CLI instalados  
❌ **Quando evitar**: Ambiente Windows (use PowerShell em vez disso)

---

## Seção 2: Formato do Arquivo JSON de Entrada

### Estrutura Básica

O arquivo JSON deve ser um **array de objetos**, cada um descrevendo uma aplicação a ser criada:

```json
[
  {
    "name": "app-identifier",
    "display_name": "Display Name da Aplicação",
    "auth_type": "app_only",
    "sign_in_audience": "AzureADMyOrg",
    "site_id": "contoso.sharepoint.com,site-guid,web-guid",
    "access_type": "leitura",
    "secret_display_name": "app-secret",
    "secret_expiration_date": "11/05/2028"
  },
  {
    "name": "another-app",
    "display_name": "Another Application",
    "auth_type": "delegated",
    "sign_in_audience": "AzureADMyOrg",
    "site_id": "contoso.sharepoint.com,site-guid,web-guid",
    "access_type": "read",
    "redirect_uri": "http://localhost:8000",
    "secret_display_name": "another-secret",
    "secret_expiration_date": "10/05/2027"
  }
]
```

### Campos Disponíveis

| Campo | Obrigatório | Tipo | Padrão | Descrição |
|-------|:-----------:|------|--------|-----------|
| `name` | ✅* | string | — | Identificador único (ex: `app-sharepoint-sales`) |
| `display_name` | ✅* | string | — | Nome exibido no Azure AD (ex: `Sales Team SharePoint`) |
| `auth_type` | ❌ | string | `"app_only"` | Tipo de autenticação: `app_only` para client credentials ou `delegated` para fluxo com usuário |
| `sign_in_audience` | ❌ | string | `"AzureADMyOrg"` | Tipo de conta: `AzureADMyOrg`, `AzureADMultipleOrgs`, `AzureADandPersonalMicrosoftAccount` |
| `site_id` | ✅ | string | — | ID do site do SharePoint que receberá a concessão (ex: `contoso.sharepoint.com,<site-guid>,<web-guid>`) |
| `access_type` | ✅ | string | — | Tipo de acesso para a aplicação no site: `leitura` ou `escrita` |
| `redirect_uri` | ✅** | string | — | URI de redirecionamento obrigatória para apps `delegated` (ex: `http://localhost:8000`) |
| `secret_display_name` | ❌ | string | `"{display_name}-secret"` | Nome do segredo no Azure (ex: `sales-app-secret`) |
| `secret_expiration_date` | ✅ | string | — | Data de expiração do segredo (formato: **dd/mm/aaaa**, ex: `11/05/2028`). Máximo: 730 dias a partir de hoje |

**\*** Forneça **`name`** OU **`display_name`** (ambos aceitos, um é obrigatório)

**\*\*** Obrigatório quando `auth_type` for `delegated`

Para `auth_type: "app_only"`, o script adiciona `Sites.Selected` como permissão de aplicativo (`Role`), concede consentimento administrativo e cria a permissão no site informado em `site_id` com o papel definido em `access_type`.

Para `auth_type: "delegated"`, o script cria o app com a `redirect_uri` informada, adiciona `Sites.Selected` como permissão delegada (`Scope`), concede consentimento administrativo e cria a mesma concessão por site. Em tempo de execução, o acesso fica limitado ao site inscrito e à interseção com as permissões do usuário autenticado.

O utilitário não adiciona permissões amplas como `Sites.Read.All` ou `Sites.ReadWrite.All`. Não são incluídas autorizações de dados no nível do tenant.

### Exemplo Completo

Veja [`examples/bulk_create_example.json`](../examples/bulk_create_example.json) para um exemplo pronto para usar:

```json
[
  {
    "name": "app-sharepoint-sales",
    "display_name": "Sales Team SharePoint Access",
    "auth_type": "app_only",
    "sign_in_audience": "AzureADMyOrg",
    "site_id": "contoso.sharepoint.com,11111111-1111-1111-1111-111111111111,22222222-2222-2222-2222-222222222222",
    "access_type": "leitura",
    "secret_display_name": "sales-app-secret",
    "secret_expiration_date": "11/05/2028"
  },
  {
    "name": "app-sharepoint-hr",
    "display_name": "HR Department SharePoint Access",
    "auth_type": "app_only",
    "sign_in_audience": "AzureADMyOrg",
    "site_id": "contoso.sharepoint.com,33333333-3333-3333-3333-333333333333,44444444-4444-4444-4444-444444444444",
    "access_type": "escrita",
    "secret_display_name": "hr-app-secret",
    "secret_expiration_date": "10/05/2027"
  },
  {
    "name": "app-sharepoint-portal",
    "display_name": "Portal Delegated Access",
    "auth_type": "delegated",
    "sign_in_audience": "AzureADMyOrg",
    "site_id": "contoso.sharepoint.com,55555555-5555-5555-5555-555555555555,66666666-6666-6666-6666-666666666666",
    "access_type": "leitura",
    "redirect_uri": "http://localhost:8000",
    "secret_display_name": "portal-app-secret",
    "secret_expiration_date": "15/06/2028"
  }
]
```

---

## Seção 3: Guia Passo a Passo — PowerShell (Recomendado)

### Pré-requisitos

- Windows 10/11
- Permissões de administrador (para instalar software e habilitar scripts)
- Acesso a conta Azure AD com direitos de **Administrador Global** ou **Privileged Role Administrator**

### Passo 1: Atualizar/Instalar PowerShell 7

PowerShell 7 é recomendado (5.1 requer modificações para o operador `??`).

```powershell
# Verificar versão atual
$PSVersionTable.PSVersion

# Se versão < 7, instalar PowerShell 7
winget install Microsoft.PowerShell
```

**Após instalação:**
- Abra "PowerShell 7" no menu Iniciar (não é o "Windows PowerShell")
- Feche a aba atual e reabra para usar a nova versão

```powershell
# Confirmar versão (deve ser 7.x)
$PSVersionTable.PSVersion
```

### Passo 2: Habilitar Execução de Scripts (Política de Segurança)

Por padrão, PowerShell bloqueia execução de scripts não assinados. Você precisa ajustar isso:

```powershell
# Verificar política atual
Get-ExecutionPolicy

NÃO é necessário fazer nada se a resposta é: `Bypass`, `Unrestricted` ou `RemoteSigned`

# Se retornar "Restricted" ou "AllSigned", habilitar para usuário atual
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force

# Confirmar (deve retornar "RemoteSigned")
Get-ExecutionPolicy
```

**O que significa `RemoteSigned`:**
- ✅ Scripts locais podem executar
- ✅ Scripts da internet precisam ser assinados
- ⚠️ Segurança equilibrada

### Passo 3: Instalar Microsoft Graph PowerShell SDK

```powershell
# Instalar módulo (requer internet)
Install-Module Microsoft.Graph -Scope CurrentUser -Force

# Confirmar instalação
Get-Module Microsoft.Graph -ListAvailable
```

**Nota sobre "Untrusted repository":**
Você pode receber um aviso de repositório não confiável. Isso é normal e seguro — PSGallery é o repositório oficial da Microsoft. O `-Force` pula esse aviso automaticamente.

### Passo 4: Autenticar com Azure AD

```powershell
# Conectar ao Microsoft Graph
Connect-MgGraph -Scopes "Application.ReadWrite.All", "Directory.Read.All", "Sites.FullControl.All"

# Navegador abrirá — entre com conta de Administrador Global
# Após autenticação, terminal confirmará "Welcome To Microsoft Graph!"
```

### Passo 5: Preparar Arquivo JSON de Entrada

**Opção A: Usar modelo existente**
```powershell
# Copiar exemplo
Copy-Item examples/bulk_create_example.json apps.json

# Editar conforme necessário
notepad apps.json
```

**Opção B: Criar arquivo novo**
```powershell
# Criar arquivo com 1 aplicação
@(
    @{
        "name" = "app-sharepoint-test"
        "display_name" = "Test SharePoint Application"
        "auth_type" = "app_only"
        "sign_in_audience" = "AzureADMyOrg"
        "site_id" = "contoso.sharepoint.com,11111111-1111-1111-1111-111111111111,22222222-2222-2222-2222-222222222222"
        "access_type" = "leitura"
        "secret_display_name" = "test-app-secret"
        "secret_expiration_date" = "11/05/2028"
    }
) | ConvertTo-Json | Set-Content apps.json
```

### Passo 6: Executar o Script Bulk Create

```powershell
# Navegar ao diretório do projeto
cd C:\GitHub\MSGraphTest

# Executar script
.\src\bulkCreate\Bulk-CreateApps.ps1 -InputPath "apps.json"
```

**Saída esperada:**
```
ℹ️  Creating 1 application(s)...

[1/1] Test SharePoint Application
  Creating app registration: Test SharePoint Application
    ✓ App created
    ✓ Secret created (expires: 11/05/2028)
    ✓ Sites.Selected permission added
    ✓ Admin consent granted
    ✓ Site access granted (read)
✅ Created successfully

✅ Results saved to: apps-output.json

📊 Summary: 1/1 apps created successfully
```

### Passo 7: Verificar Arquivo de Saída

```powershell
# Visualizar resultado em formato legível
Get-Content apps-output.json | ConvertFrom-Json | Format-Table

# Ou exportar para CSV
Get-Content apps-output.json | ConvertFrom-Json | Export-Csv credentials.csv
```

**Arquivo `apps-output.json` conterá:**
```json
[
  {
    "name": "app-sharepoint-test",
    "display_name": "Test SharePoint Application",
    "auth_type": "app_only",
    "sign_in_audience": "AzureADMyOrg",
    "site_id": "contoso.sharepoint.com,11111111-1111-1111-1111-111111111111,22222222-2222-2222-2222-222222222222",
    "access_type": "leitura",
    "secret_display_name": "test-app-secret",
    "secret_expiration_date": "11/05/2028",
    "app_object_id": "12345678-...",
    "tenant_id": "87654321-...",
    "client_id": "aaaaaaaa-...",
    "client_secret": "Abc~defGhIjk...",
    "status": "success",
    "created_at": "2025-05-11T14:30:45.123456+00:00"
  }
]
```

### Passo 8: Salvar Credenciais no `.env`

```powershell
# Extrair credenciais
$app = (Get-Content apps-output.json | ConvertFrom-Json)[0]

# Copiar template
Copy-Item .env.example .env

# Editar e adicionar:
notepad .env
```

**Conteúdo do `.env`:**
```bash
AZURE_TENANT_ID=87654321-4321-4321-4321-abc123456789
AZURE_CLIENT_ID=aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee
AZURE_CLIENT_SECRET=Abc~defGhIjklMnOpQrStUvWxYz1234567890
```

### Passo 9 (Opcional): Criar Múltiplas Apps sem Reauthenticar

Se você já está autenticado (Passo 4), pode reutilizar a sessão:

```powershell
# Criar lotes adicionais
.\src\bulkCreate\Bulk-CreateApps.ps1 -InputPath "batch2.json" -OutputPath "results2.json" -SkipLogin
.\src\bulkCreate\Bulk-CreateApps.ps1 -InputPath "batch3.json" -OutputPath "results3.json" -SkipLogin
```

### Passo 10: Desconectar (Limpeza)

```powershell
# Finalizar sessão
Disconnect-MgGraph
```

---

## Seção 4: Guia Passo a Passo — Python + Azure CLI

### Pré-requisitos

- Python 3.11+ instalado
- Permissões de administrador (para instalar Azure CLI)
- Acesso a conta Azure AD com direitos de **Administrador Global** ou **Privileged Role Administrator**

### Passo 1: Instalar/Atualizar Azure CLI

**Windows:**
```bash
winget install Microsoft.AzureCLI
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

### Passo 2: Autenticar com Azure AD

```bash
# Fazer login
az login

# Navegador abrirá — entre com conta de Administrador Global
# Terminal exibirá: "You have logged in"
```

### Passo 3: Instalar Dependências Python

```bash
# Navegar ao diretório do projeto
cd C:\GitHub\MSGraphTest

# Instalar ambiente e dependências
uv sync
```

### Passo 4: Preparar Arquivo JSON de Entrada

**Opção A: Usar modelo existente**
```bash
# Copiar exemplo
cp examples/bulk_create_example.json apps.json

# Editar conforme necessário
nano apps.json  # ou seu editor favorito
```

**Opção B: Criar arquivo novo**
```bash
cat > apps.json << 'EOF'
[
  {
    "name": "app-sharepoint-test",
    "display_name": "Test SharePoint Application",
    "auth_type": "app_only",
    "sign_in_audience": "AzureADMyOrg",
    "site_id": "contoso.sharepoint.com,11111111-1111-1111-1111-111111111111,22222222-2222-2222-2222-222222222222",
    "access_type": "leitura",
    "secret_display_name": "test-app-secret",
    "secret_expiration_date": "11/05/2028"
  }
]
EOF
```

### Passo 5: Executar o Script Bulk Create

```bash
# Executar com Azure CLI (padrão, mais rápido)
python -m bulkCreate.bulk_create_apps apps.json
```

**Saída esperada:**
```
🔐 Logging in to Azure (cli)...
✅ Logged in

📱 Creating 1 application(s)...

[1/1] Creating: Test SharePoint Application
  ✅ Created successfully
     Client ID: aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee
     Tenant ID: 87654321-4321-4321-4321-abc123456789

✅ Results saved to: apps-output.json

📊 Summary: 1/1 apps created successfully
```

### Passo 6: Verificar Arquivo de Saída

```bash
# Visualizar resultado
cat apps-output.json | python -m json.tool

# Ou extrair campos específicos
python -c "
import json
with open('apps-output.json') as f:
    apps = json.load(f)
for app in apps:
    print(f'{app[\"display_name\"]}: {app[\"client_id\"]}')"
```

**Arquivo `apps-output.json` conterá:**
```json
[
  {
    "name": "app-sharepoint-test",
    "display_name": "Test SharePoint Application",
    "auth_type": "app_only",
    "sign_in_audience": "AzureADMyOrg",
    "site_id": "contoso.sharepoint.com,11111111-1111-1111-1111-111111111111,22222222-2222-2222-2222-222222222222",
    "access_type": "leitura",
    "secret_display_name": "test-app-secret",
    "secret_expiration_date": "11/05/2028",
    "app_object_id": "12345678-...",
    "tenant_id": "87654321-...",
    "client_id": "aaaaaaaa-...",
    "client_secret": "Abc~defGhIjk...",
    "status": "success",
    "created_at": "2025-05-11T14:30:45.123456+00:00"
  }
]
```

### Passo 7: Salvar Credenciais no `.env`

```bash
# Copiar template
cp .env.example .env

# Extrair e adicionar credenciais
python << 'EOF'
import json

with open('apps-output.json') as f:
    apps = json.load(f)

app = apps[0]  # primeira app
print(f"AZURE_TENANT_ID={app['tenant_id']}")
print(f"AZURE_CLIENT_ID={app['client_id']}")
print(f"AZURE_CLIENT_SECRET={app['client_secret']}")
EOF

# Editar arquivo manualmente
nano .env  # ou seu editor
```

**Conteúdo do `.env`:**
```bash
AZURE_TENANT_ID=87654321-4321-4321-4321-abc123456789
AZURE_CLIENT_ID=aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee
AZURE_CLIENT_SECRET=Abc~defGhIjklMnOpQrStUvWxYz1234567890
```

### Passo 8 (Opcional): Criar Múltiplas Apps sem Reauthenticar

Se você já está autenticado (Passo 2), pode reutilizar a sessão:

```bash
# Criar lotes adicionais sem fazer login novamente
python -m bulkCreate.bulk_create_apps batch2.json --output results2.json --skip-login
python -m bulkCreate.bulk_create_apps batch3.json --output results3.json --skip-login
```

### Passo 9 (Opcional): Usar PowerShell SDK em vez de Azure CLI

Se o Azure CLI der problemas, use Microsoft Graph PowerShell SDK:

```bash
# Instalar módulo PowerShell
pwsh -Command "Install-Module Microsoft.Graph -Scope CurrentUser"

# Executar com PowerShell SDK
python -m bulkCreate.bulk_create_apps apps.json --method powershell
```

### Passo 10: Desconectar (Limpeza)

```bash
# Logout do Azure CLI
az logout
```

---

## Segurança e Próximos Passos

### ⚠️ Tratamento de Secrets

O arquivo `*-output.json` **contém credenciais reais**:

- ✅ **Armazene com segurança**: Azure Key Vault, secrets manager, etc.
- ✅ **Nunca commite**: Adicione `*-output.json` ao `.gitignore`
- ✅ **Delete após usar**: Não deixe credenciais em disco
- ✅ **Rotacione regularmente**: Renove secrets a cada 6-12 meses
- ✅ **Não compartilhe**: Evite email, Slack, repositórios públicos

```bash
# Adicionar ao .gitignore
echo "*-output.json" >> .gitignore
```

### Próximas Etapas Após Criar Apps

1. **Registrar sites no SharePoint** (requer administrador SP)
   - Veja [setup_cli.md](setup_cli.md) — Etapa 4

2. **Descobrir IDs de site/drive/list**
   - Usar scripts em `examples/`
   - Ou [Microsoft Graph Explorer](https://developer.microsoft.com/en-us/graph/graph-explorer)

3. **Testar com exemplos Python**
   ```bash
   uv run examples/example_drive_list.py
   uv run examples/example_list_get.py
   ```

4. **Validar acesso ao SharePoint**
   - Confirmar que apps conseguem ler/escrever conforme esperado

---

## Referências

- [Microsoft Graph PowerShell SDK](https://github.com/microsoftgraph/msgraph-sdk-powershell)
- [Azure CLI Documentation](https://docs.microsoft.com/en-us/cli/azure/)
- [Sites.Selected Permission](https://docs.microsoft.com/en-us/graph/sites-selected-permission)
- [PowerShell Execution Policies](https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_execution_policies)
