# Copilot Context — MSGraphTest

Este arquivo resume o contexto atual do repositório para manutenção de código e documentação.
Atualize-o sempre que o contrato funcional do projeto mudar.

---

## Visão Geral

**Objetivo**: biblioteca Python e documentação de referência para acessar SharePoint via Microsoft Graph com privilégio mínimo, sempre restringindo o acesso a sites explicitamente inscritos.

**Modelos suportados**:
- `app_only` (client credentials)
- `delegated` (authorization code / usuário autenticado)

**Princípio central**:
- O projeto usa `Sites.Selected` em ambos os modelos.
- Não usa `Sites.Read.All` nem `Sites.ReadWrite.All` para acesso de dados.
- Não são incluídas autorizações de dados no nível do tenant.
- Para qualquer app, o acesso só existe depois de uma concessão explícita em um `site_id` concreto.

No fluxo `delegated`, o acesso efetivo em runtime é a interseção entre:
1. a concessão do aplicativo no site (`read` ou `write`)
2. as permissões que o usuário autenticado já tem nesse mesmo site

---

## Stack

- Python 3.11+
- MSAL
- Microsoft Graph API
- Azure CLI
- Microsoft Graph PowerShell SDK
- pytest + pytest-mock

---

## Estrutura

```text
MSGraphTest/
├── src/
│   ├── bulkCreate/
│   │   ├── bulk_create_apps.py
│   │   └── Bulk-CreateApps.ps1
│   └── msgraphtest/
│       ├── __init__.py
│       ├── auth.py
│       ├── graph_client.py
│       ├── drive.py
│       └── lists.py
├── docs/
│   ├── getting_started.md
│   ├── setup_cli.md
│   ├── setup_portal.md
│   ├── setup_delegated_auth.md
│   └── bulk_create_apps.md
├── examples/
│   └── bulk_create_example.json
├── tests/
└── pyproject.toml
```

---

## Permissões e Constantes

**Microsoft Graph App ID**:

```text
00000003-0000-0000-c000-000000000000
```

**Permissão usada para SharePoint**:
- Nome: `Sites.Selected`
- Tipo para `app_only`: `Role`
- Tipo para `delegated`: `Scope`

**Importante**:
- Não hardcode IDs de `Sites.Selected` em documentação nova se puder resolvê-los dinamicamente.
- No código, prefira descobrir o ID correto a partir do service principal do Microsoft Graph.

---

## Bulk Create

### Arquivos principais

- `src/bulkCreate/bulk_create_apps.py`
- `src/bulkCreate/Bulk-CreateApps.ps1`

### Responsabilidade

Criar apps em lote a partir de JSON, incluindo:
- criação do app registration
- criação de segredo
- adição de `Sites.Selected`
- consentimento administrativo
- inscrição do `site_id` informado com `read` ou `write`

### Contrato atual do JSON

Campos relevantes:
- `name` ou `display_name`
- `auth_type`: `app_only` ou `delegated`
- `site_id`: obrigatório para todos os apps
- `access_type`: obrigatório para todos os apps; aceita `leitura`, `escrita`, `read`, `write`
- `redirect_uri`: obrigatório quando `auth_type = delegated`
- `secret_expiration_date`: obrigatório no formato `dd/mm/aaaa`, limitado a 730 dias

Exemplo mínimo:

```json
[
  {
    "display_name": "Portal Delegated Access",
    "auth_type": "delegated",
    "site_id": "contoso.sharepoint.com,site-guid,web-guid",
    "access_type": "leitura",
    "redirect_uri": "http://localhost:8000",
    "secret_expiration_date": "15/06/2028"
  }
]
```

### Comportamento esperado

- `app_only`: adiciona `Sites.Selected` como `Role`
- `delegated`: adiciona `Sites.Selected` como `Scope`
- ambos os fluxos concedem consentimento administrativo e inscrevem o site
- ambos os fluxos devem falhar se `site_id` ou `access_type` não forem fornecidos

---

## Documentação Fonte de Verdade

- `docs/getting_started.md`: visão geral, papéis administrativos, modelo de privilégio mínimo
- `docs/setup_cli.md`: configuração por linha de comando
- `docs/setup_portal.md`: configuração manual pelo portal e Graph Explorer
- `docs/setup_delegated_auth.md`: fluxo delegado com `Sites.Selected` e grant por site
- `docs/bulk_create_apps.md`: automação em lote e schema JSON

Ao alterar o comportamento do código, mantenha esses documentos alinhados.

---

## Papéis Administrativos

- **Desenvolvimento**: cria o app registration, configura código, descobre IDs de site/drive/list
- **Administrador Entra**: concede consentimento administrativo de `Sites.Selected`
- **Administrador SharePoint**: inscreve sites via `POST /sites/{site-id}/permissions`

Observações:
- `Application Administrator` e `Cloud Application Administrator` não são suficientes para este consentimento do Microsoft Graph.
- A inscrição do site exige papel de administrador do SharePoint no tenant, não apenas proprietário do site.

---

## Testes e Validação

Comandos preferenciais:

```bash
uv run pytest
```

Para mudanças locais em scripts:
- Python: checagem estática e `py_compile` quando houver interpretador disponível
- PowerShell: parser/diagnóstico do editor e execução controlada quando o ambiente permitir

---

## Segurança

Nunca versionar:
- `.env`
- arquivos de saída contendo `client_secret`
- caches de token como `.msal_token_cache.json`

Preferências do projeto:
- privilégios mínimos
- grants por site
- segredos com expiração explícita e curta

---

## Última Atualização

- Data: 11/05/2026
- Alteração principal: `Sites.Selected` passou a ser tratado como obrigatório para `app_only` e `delegated`, sempre com restrição a um `site_id` concreto e sem autorizações de dados no nível do tenant
