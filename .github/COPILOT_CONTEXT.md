# Copilot Context — MSGraphClient

Este arquivo resume o estado atual do repositório para manutenção de código e documentação.
Atualize-o sempre que a superfície pública ou a estrutura do projeto mudar.

---

## Visão Geral

**Objetivo**: biblioteca Python e exemplos para acessar SharePoint via Microsoft Graph com privilégio mínimo, fornecendo uma camada de abstração sobre MSAL para integração em apps desktop e mobile.

**Modelos suportados**:
- `GraphClient` como ponto de entrada principal para Microsoft Graph
- `GraphAuthenticator` para autenticação e descoberta do site
- `GraphDrive` para operações de biblioteca de documentos
- `GraphList` para operações de listas do SharePoint

**Princípios centrais**:
- `GraphClient` possui `GraphAuthenticator`.
- `GraphList` deriva `site_id` do `client.authenticator` quando possível.
- Não existem mais wrappers de compatibilidade em nível de módulo para drive/list/site.
- As consultas de listas usam seleção no momento da requisição, evitando filtragem tardia desnecessária.
- `get_views()` retorna `[]` quando a lista não expõe views.

No fluxo `delegated`, o acesso efetivo em runtime continua sendo a interseção entre a concessão do aplicativo no site e as permissões do usuário autenticado nesse site.

---

## Stack

- Python 3.14.0
- `uv`
- `pytest`
- `requests`
- `msal`
- `pandas`
- `python-dotenv`

---

## Estrutura Atual

```text
MSGraphClient/
├── docs/
├── downloads/
├── examples/
│   ├── example_drive_download.py
│   ├── example_drive_list.py
│   ├── example_drive_read_write.py
│   ├── example_drive_upload.py
│   ├── example_list_create.py
│   ├── example_list_get.py
│   ├── example_list_update.py
│   ├── example_site_contents.py
├── notebooks/
│   └── graph_auth_site_attributes.ipynb
├── src/
│   ├── bulkCreate/
│   └── python/
│       ├── __init__.py
│       ├── auth.py
│       ├── drive.py
│       └── lists.py
├── tests/
├── pyproject.toml
└── README.md
```

---

## Graph Surface

### `src/msgraphclient/auth.py`
- `GraphClient` encapsula a sessão HTTP autenticada e a formatação de erros Graph.
- `GraphAuthenticator` valida config, obtém token e expõe metadados do site.
- A descoberta do site agora vive em `GraphAuthenticator`.

### `src/msgraphclient/drive.py`
- `GraphDrive.ls(path=None)`
- `GraphDrive.pwd()`
- `GraphDrive.cd(path)`
- `GraphDrive.download(item_id, local_path)`
- `GraphDrive.upload(local_path, remote_folder="root", remote_name=None)`
- `GraphDrive.read(item_id, encoding=None)` — charset auto-detected from HTTP response; stored in `last_encoding`
- `GraphDrive.write(item_id, content, encoding=None)` — uses `last_encoding` for round-trip fidelity

### `src/msgraphclient/lists.py`
- `GraphList.get_views()` com fallback seguro para listas sem views
- `GraphList.get_view_columns(view_id)`
- `GraphList.get_columns(names=None)` com filtro de metadados via Graph
- `GraphList.get_schema()` e `GraphList.get_field_types()`
- `GraphList.get_items(select=None, include_id=True)`
- `GraphList.get_items_dataframe(select=None, include_id=True)`
- `GraphList.get_item_template(include_optional=True)`
- `GraphList.validate_item(data)`
- `GraphList.save_item(data)` e `GraphList.save_items(items)`
- `GraphList.save_dataframe(dataframe)`

---

## Documentação e Exemplos

- Os exemplos de Graph foram migrados para uso direto das classes.
- O notebook `notebooks/graph_auth_site_attributes.ipynb` é a validação end-to-end principal do fluxo SharePoint.
- O README foi alinhado com a superfície atual e não deve mencionar wrappers de compatibilidade removidos.

---

## Testes e Validação

Comando preferencial:

```bash
uv run pytest tests/
```

Estado validado no último sweep: 44 testes passaram.

---

## Segurança

Nunca versionar:
- `.env`
- caches de token
- arquivos de saída contendo segredos

Preferências do projeto:
- privilégios mínimos
- grants por site
- segredos com expiração explícita

---

## Última Atualização

- Data: 26/05/2026
- Alteração principal: remoção dos wrappers legados, consolidação da API em classes e atualização dos exemplos/notebooks.
