# Copilot Context — MSGraphClient

Este arquivo resume o estado atual do repositório para manutenção de código e documentação.
Atualize-o sempre que a superfície pública, dependências ou estrutura do projeto mudarem.

---

## Visão Geral

**Objetivo**: biblioteca Python para abstrair autenticação com MSAL e simplificar acesso ao SharePoint Online via Microsoft Graph API, com foco em privilégio mínimo (`Sites.Selected`).

**Fluxos suportados**:
- `client_credentials` (app-only)
- `delegated` (`interactive` ou `device_code`)

**Componentes públicos principais**:
- `GraphAuthenticator`: aquisição de token e validação de credenciais/configuração.
- `GraphClient`: sessão HTTP para Graph, helpers de request e normalização de erros HTTP.
- `GraphDrive`: operações de biblioteca de documentos (navegação, upload/download, leitura/escrita).
- `GraphList`: operações de listas (schema, views, leitura paginada, validação e save).

No fluxo `delegated`, o acesso efetivo em runtime continua sendo a interseção entre:
- concessão do aplicativo no site (`Sites.Selected`)
- permissões reais do usuário autenticado nesse mesmo site.

---

## Stack

- Python `>=3.11` (README indica testes com 3.14)
- `uv` (ambiente/dependências)
- `pytest`, `pytest-cov`, `pytest-mock`
- `requests`
- `msal`
- `python-dotenv`
- `pandas`
- `python-dateutil`

---

## Estrutura Atual

```text
MSGraphClient/
├── docs/
│   ├── bulk_create_apps.md
│   ├── getting_started.md
│   ├── setup_cli.md
│   ├── setup_delegated_auth.md
│   └── setup_portal.md
├── examples/
│   ├── example_drive_download.py
│   ├── example_drive_folder_operations.py
│   ├── example_drive_list.py
│   ├── example_drive_read_write.py
│   ├── example_drive_upload.py
│   ├── example_list_create.py
│   ├── example_list_get.py
│   ├── example_list_update.py
│   ├── example_site_contents.py
│   ├── list_value_generation.py
│   └── downloads/
├── notebooks/
│   ├── graph_auth_site_attributes.ipynb
│   └── downloads/
├── src/
│   ├── bulkCreate/
│   └── msgraphclient/
│       ├── __init__.py
│       ├── auth.py
│       ├── client.py
│       ├── drive.py
│       ├── lists.py
│       ├── messages.py
│       ├── settings.py
│       └── locales/
├── tests/
│   ├── test_auth.py
│   ├── test_drive.py
│   ├── test_graph_client.py
│   ├── test_lists.py
│   ├── test_list_value_generation.py
│   ├── test_settings.py
│   └── test_site.py
├── pyproject.toml
└── README.md
```

---

## Superfície de API (Atual)

### `src/msgraphclient/client.py`
- `GraphAuthorizationError` (especializa falhas 401/403).
- `GraphClient.format_http_error(error)`
- `GraphClient.get(path, **kwargs)`
- `GraphClient.post(path, json, **kwargs)`
- `GraphClient.patch(path, json, **kwargs)`
- `GraphClient.put_bytes(path, data, content_type=..., **kwargs)`
- `GraphClient.get_raw(path, **kwargs)`
- `GraphClient.get_raw_with_encoding(path, **kwargs)`

Também popula atributos de site (`site_graph_id`, `site_name`, `site_display_name`, `site_web_url`, `site_drives`, `site_lists`) quando `SHAREPOINT_SITE_ID` está disponível.

### `src/msgraphclient/auth.py`
- `GraphAuthenticator` com resolução por `GraphSettings` e suporte a `client_credentials` e `delegated`.
- Reexporta `GraphClient` e `GraphAuthorizationError`.
- Usa cache de token em memória para fluxo delegado.

### `src/msgraphclient/drive.py`
- `GraphDrive.pwd()`
- `GraphDrive.cd(path)`
- `GraphDrive.ls(path=None)`
- `GraphDrive.download(item_id, local_path)`
- `GraphDrive.upload(local_path, remote_folder="root", remote_name=None)`
- `GraphDrive.read(item_id, encoding=None)`
- `GraphDrive.write(item_id, content, encoding=None)`

`read()` detecta charset da resposta HTTP quando possível e persiste em `last_encoding`, usado por `write()` quando `encoding` não é informado.

### `src/msgraphclient/lists.py`
- `GraphList.get_views()` (com fallback seguro)
- `GraphList.get_view_columns(view_id)`
- `GraphList.get_columns(names=None)`
- `GraphList.get_schema()`
- `GraphList.get_field_types()`
- `GraphList.validate_item(data)`
- `GraphList.get_items(select=None, include_id=True)`
- `GraphList.get_item_template(include_optional=True)`
- `GraphList.get_items_dataframe(select=None, include_id=True)`
- `GraphList.save_dataframe(dataframe)`
- `GraphList.save_item(data)`
- `GraphList.save_items(items)`

---

## Documentação e Exemplos

- README concentra onboarding, autenticação (`client_credentials` e `delegated`), execução de exemplos e testes.
- Notebook principal de validação end-to-end: `notebooks/graph_auth_site_attributes.ipynb`.
- Guias operacionais estão em `docs/`, incluindo setup por CLI/Portal e criação em lote de apps.

---

## Testes e Validação

Comando recomendado:

```bash
uv run pytest tests/
```

Observação desta revisão:
- A execução local não foi concluída nesta sessão devido a erro de ambiente ao invocar `uv run` (`Failed to canonicalize script path`).
- Não afirmar contagem de testes aprovados sem nova execução local ou CI.

---

## Segurança

Nunca versionar:
- `.env`
- caches de token
- outputs contendo segredos/tokens

Diretrizes:
- manter privilégio mínimo (`Sites.Selected`)
- restringir grants por site
- usar segredos com rotação/expiração explícita

---

## Última Atualização

- Data: 09/06/2026
- Alteração principal: contexto sincronizado com estrutura atual (`src/msgraphclient`, exemplos e testes), superfície pública revisada e seção de validação ajustada para refletir o estado real da sessão.
