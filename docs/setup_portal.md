# Configuração passo a passo para `Sites.Selected` — via Portal e Navegador

Este guia usa exclusivamente interfaces web: **portal do Azure** e **Microsoft Graph Explorer**. Não requer instalação de ferramentas de linha de comando.

> Para instruções equivalentes usando Azure CLI e PowerShell, veja [setup_cli.md](setup_cli.md).

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

1. Abra o [portal do Azure](https://portal.azure.com) → **Microsoft Entra ID** → **Registros de app** → **Novo registro**.
2. **Nome**: escolha um nome descritivo, ex: `MSGraphClient-SharePoint`.
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

## Etapa 2 — Adicionar `Sites.Selected` e conceder consentimento de administrador

> 🔧 **Equipe de desenvolvimento** adiciona a permissão. 🔑 **Administrador Entra** concede consentimento.

1. No registro de aplicativo vá para **Permissões de API** → **Adicionar uma permissão** → **Microsoft Graph**.
2. Para apps `app_only`, escolha **Permissões de aplicativo** e adicione **`Sites.Selected`**.
3. Para apps `delegated`, escolha **Permissões delegadas** e adicione **`Sites.Selected`**.
4. Não adicione `Sites.Read.All` nem `Sites.ReadWrite.All`; este repositório não usa autorizações de dados no nível do tenant.
5. Clique em **Conceder consentimento de administrador para <tenant>** e confirme.

> ⚠️ Esta ação requer **Administrador global** ou **Administrador de função com privilégios** (*Privileged Role Administrator*) — a função Administrador de Aplicativos não é suficiente para permissões de aplicativo do Microsoft Graph.

> `Sites.Selected` não concede acesso a nenhum site ainda — isso acontece na Etapa 4 para ambos os fluxos. No caso delegado, o acesso final ainda será a interseção entre a concessão do site e as permissões do usuário autenticado.

---

## Etapa 3 — Descobrir o ID do site do SharePoint

> 🔧 **Equipe de desenvolvimento**

Esta etapa usa um token delegado da sua conta de usuário pessoal apenas para consultar o ID do site. Este token não é usado pelo aplicativo e não concede nada ao registro de aplicativo.

1. Abra o [Microsoft Graph Explorer](https://developer.microsoft.com/graph/graph-explorer).
2. Clique em **Sign in** e autentique com sua conta do Microsoft 365.
3. Cole a URL no campo de entrada, substituindo `<hostname>` e `<site-path>` pelos valores do seu tenant:
   ```
   https://graph.microsoft.com/v1.0/sites/<hostname>:/sites/<site-path>
   ```
   Exemplo: `https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/sites/ProjectAlpha`
4. Clique em **Run query**.
5. Copie o campo `id` da resposta → `SHAREPOINT_SITE_ID`.

Formato esperado: `contoso.sharepoint.com,xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx,yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy`

> Compartilhe este valor com o Administrador SP antes de prosseguir para a Etapa 4.

---

## Etapa 4 — Inscrever o site para o aplicativo

> 🛡️ **Administrador SP** (administrador do tenant do SharePoint)

> ⚠️ A conta executando esta etapa deve ter a função **Administrador do SharePoint** ou **Administrador Global**. Um administrador de coleção de sites / proprietário de site não pode chamar `POST /sites/{site-id}/permissions`.

1. Abra o [Microsoft Graph Explorer](https://developer.microsoft.com/graph/graph-explorer) e autentique com a conta do **Administrador SP**.
2. Clique em **Modify permissions** (ícone de cadeado) e conceda o escopo **`Sites.FullControl.All`** — permissão delegada necessária para gerenciar permissões de site.
3. Configure a requisição:
   - **Método HTTP**: `POST`
   - **URL**: `https://graph.microsoft.com/v1.0/sites/<SHAREPOINT_SITE_ID>/permissions`
4. Clique na aba **Request body**, selecione o tipo `application/json` e cole um dos corpos abaixo.

**Conceder acesso somente leitura** (substitua `<AZURE_CLIENT_ID>`):

```json
{
  "roles": ["read"],
  "grantedToIdentities": [{
    "application": {
      "id": "<AZURE_CLIENT_ID>",
      "displayName": "MSGraphClient-SharePoint"
    }
  }]
}
```

**Conceder acesso leitura + escrita** — substitua `"read"` por `"write"` (escrita implicitamente inclui leitura):

```json
{
  "roles": ["write"],
  "grantedToIdentities": [{
    "application": {
      "id": "<AZURE_CLIENT_ID>",
      "displayName": "MSGraphClient-SharePoint"
    }
  }]
}
```

5. Clique em **Run query**. A resposta deve retornar `201 Created`.

**Verificar a concessão:**
- Método: `GET`
- URL: `https://graph.microsoft.com/v1.0/sites/<SHAREPOINT_SITE_ID>/permissions`

**Revogar uma concessão** (use o `id` retornado pela verificação acima):
- Método: `DELETE`
- URL: `https://graph.microsoft.com/v1.0/sites/<SHAREPOINT_SITE_ID>/permissions/<permission-id>`

---

## Replicar esta configuração para outros aplicativos

Para criar um segundo app com a mesma configuração (mesmas permissões, sem URI de redirecionamento, fluxo de credenciais do cliente), use o **Manifest JSON** — sem precisar repetir cada passo manualmente.

> 🔧 **Equipe de desenvolvimento** executa as etapas A–C. 🔑 **Administrador Entra** executa D. 🛡️ **Administrador SP** executa E.

**A — Exportar o manifest do app existente**
1. Abra o registro do app original no portal do Azure.
2. Clique em **Manifest** no menu lateral.
3. Clique em **Download** e salve o arquivo JSON.

**B — Criar o novo registro de aplicativo**
1. **Registros de app** → **Novo registro**.
2. Informe apenas o **Nome** do novo app. Deixe os demais campos em branco por enquanto.
3. Clique em **Registrar** e anote o novo `AZURE_CLIENT_ID`.

**C — Importar o manifest e criar novo segredo**
1. Abra o novo registro → **Manifest** → **Upload**.
2. Selecione o arquivo JSON exportado no passo A.
3. O portal importará permissões e configurações. O `appId` e `displayName` do arquivo exportado são ignorados — o Azure AD usa os do novo registro.
4. Vá para **Certificados e segredos** → **Novo segredo do cliente**.
5. Copie imediatamente o valor gerado → novo `AZURE_CLIENT_SECRET`.
   > ⚠️ Segredos **não são exportados** pelo manifest e devem ser criados individualmente para cada app.

**D — Conceder consentimento de administrador** *(obrigatório — não é copiado pelo manifest)*

O consentimento é vinculado ao `appId` específico. Deve ser concedido individualmente para cada novo app:
- **Portal**: Registro do app → **Permissões de API** → **Conceder consentimento de administrador para \<tenant\>**.

**E — Inscrever o site para o novo app** *(obrigatório — não é copiado pelo manifest)*

A inscrição de site é uma permissão no nível do SharePoint, independente do Azure AD. Deve ser repetida para cada novo app:
```
POST https://graph.microsoft.com/v1.0/sites/<SHAREPOINT_SITE_ID>/permissions
```
Corpo: substitua `<AZURE_CLIENT_ID>` pelo ID do novo app (veja a Etapa 4 deste guia para o corpo completo).

**Resumo do que é e não é copiado pelo manifest:**

| Configuração | Copiado pelo manifest? |
|---|:---:|
| Permissões de API (`Sites.Selected` de aplicativo) | ✅ |
| Tipos de conta suportados | ✅ |
| Segredo de cliente | ❌ (deve ser criado por app) |
| Consentimento de administrador | ❌ (deve ser concedido por app) |
| Inscrição de site no SharePoint | ❌ (deve ser repetida por app) |

> Para criar **múltiplos apps em lote**, veja [bulk_create_apps.md](bulk_create_apps.md) — automatiza as etapas C, D e E de uma só vez via PowerShell ou Python.

---

## Etapa 5 — Descobrir o ID da drive e ID da lista

> 🔧 **Equipe de desenvolvimento**

1. Abra o [Microsoft Graph Explorer](https://developer.microsoft.com/graph/graph-explorer) autenticado com sua conta de desenvolvedor.

**Encontrar o ID da drive:**
- Método: `GET`
- URL: `https://graph.microsoft.com/v1.0/sites/<SHAREPOINT_SITE_ID>/drives`

Localize na resposta a entrada cujo `name` corresponde à sua biblioteca de documentos e copie seu `id` → `SHAREPOINT_DRIVE_ID`.

**Encontrar o ID da lista:**
- Método: `GET`
- URL: `https://graph.microsoft.com/v1.0/sites/<SHAREPOINT_SITE_ID>/lists`

Localize a lista desejada e copie seu `id` → `SHAREPOINT_LIST_ID`.

---

## Etapa 6 — Configurar `.env` e executar

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
uv run examples/example_drive_folder_operations.py
uv run examples/example_list_get.py
```
