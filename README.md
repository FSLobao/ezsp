# MSGraphTest — SharePoint via Microsoft Graph API

Um projeto de teste em Python demonstrando como acessar o SharePoint através da
**Microsoft Graph API** usando MSAL para autenticação. As operações abordadas
incluem gerenciamento de biblioteca de documentos (drive) e manipulação de listas do SharePoint.

O projeto adota **privilégio mínimo** como regra: o acesso ao SharePoint é feito com
`Sites.Selected`, sempre restrito a sites explicitamente inscritos. Isso vale tanto
para autenticação `app_only` quanto para autenticação `delegated`.

Licenciado sob a [GNU General Public License v3.0](LICENSE).

---

## Estrutura do projeto

```
MSGraphTest/
├── src/
│   ├── bulkCreate/
│   │   ├── bulk_create_apps.py        # utilitário para criar múltiplas apps em lote (Python)
│   │   └── Bulk-CreateApps.ps1        # idem, versão PowerShell
│   └── msgraphtest/
│       ├── __init__.py                # ponto de entrada do pacote
│       ├── auth.py                    # auxiliar de token client-credentials com MSAL
│       ├── graph_client.py            # wrapper HTTP fino para chamadas REST do Graph
│       ├── drive.py                   # operações de biblioteca de documentos
│       └── lists.py                   # operações de lista do SharePoint
├── tests/
│   ├── test_auth.py
│   ├── test_drive.py
│   └── test_lists.py
├── examples/
│   ├── example_drive_list.py          # listar conteúdo raiz da drive
│   ├── example_drive_download.py      # baixar arquivo para pasta local
│   ├── example_drive_upload.py        # enviar arquivo local
│   ├── example_drive_read_write.py    # ler e atualizar conteúdo de texto do arquivo
│   ├── example_list_get.py            # recuperar todos os itens de lista
│   ├── example_list_create.py         # criar item de lista
│   ├── example_list_update.py         # atualizar item de lista
│   └── bulk_create_example.json       # modelo de entrada para bulk_create_apps
├── docs/
│   ├── getting_started.md             # guia de início rápido
│   ├── setup_cli.md                   # setup com Azure CLI / PowerShell
│   ├── setup_portal.md                # setup com Azure Portal
│   ├── setup_delegated_auth.md        # setup com autenticação delegada (usuário)
│   └── bulk_create_apps.md            # documentação de criação em lote de apps
├── downloads/                 # (ignorado por git) destino de download local
├── .env.example               # copie para .env e preencha as credenciais
├── pyproject.toml
└── LICENSE
```

---

## Pré-requisitos

| Requisito | Observações |
|---|---|
| Python ≥ 3.11 | Testado com 3.11+ |
| [UV](https://docs.astral.sh/uv/) | Gerenciador de pacotes e ambiente virtual |
| Registro de aplicativo no Microsoft Entra ID | Configure `Sites.Selected` e inscreva os sites necessários |

> Este repositório **não usa** permissões amplas como `Sites.Read.All` ou `Sites.ReadWrite.All` para acesso a dados no SharePoint.

---

## Início rápido

### 1. Clonar e instalar dependências

```bash
git clone <repo-url>
cd MSGraphTest
uv sync
```

### 2. Configurar credenciais

```bash
cp .env.example .env
# edite .env com seus detalhes do Azure AD e SharePoint
```

Variáveis obrigatórias em `.env`:

| Variável | Descrição |
|---|---|
| `AZURE_TENANT_ID` | ID do tenant (locatário) do Azure AD |
| `AZURE_CLIENT_ID` | ID do cliente do registro de aplicativo |
| `AZURE_CLIENT_SECRET` | Segredo do cliente do registro de aplicativo |
| `SHAREPOINT_SITE_ID` | ID do site do Graph (ex: `contoso.sharepoint.com,guid,guid`) |
| `SHAREPOINT_DRIVE_ID` | ID da drive (unidade) da biblioteca de documentos |
| `SHAREPOINT_LIST_ID` | ID da lista para operações de lista |

> **Encontrando IDs** — veja [docs/getting_started.md](docs/getting_started.md).

### 3. Escolher o modelo de autenticação

- **`app_only`**: indicado para automação sem interação do usuário.
- **`delegated`**: indicado quando é necessário associar as ações a um usuário autenticado.

Nos dois casos, o projeto usa `Sites.Selected` e exige inscrição explícita do site.
No fluxo `delegated`, o acesso efetivo é a interseção entre a concessão do aplicativo
no site e as permissões que o usuário já possui nesse mesmo site.

### 4. Executar um exemplo

```bash
uv run examples/example_drive_list.py
uv run examples/example_list_get.py
```

---

## Executando testes

```bash
uv run pytest
```

Relatório de cobertura é impresso automaticamente. Os testes usam mocking e **não**
requerem credenciais reais.

---

## Criação em lote de aplicações Azure AD

Para criar múltiplas aplicações Azure AD com credenciais automaticamente:

- **PowerShell (recomendado)**: `.\src/bulkCreate/Bulk-CreateApps.ps1 -InputPath config.json`
- **Python**: `python -m bulkCreate.bulk_create_apps config.json`

Autentique uma vez e execute múltiplas vezes com `-SkipLogin` (PowerShell) ou `--skip-login` (Python).
Use [examples/bulk_create_example.json](examples/bulk_create_example.json) como modelo e
veja [docs/bulk_create_apps.md](docs/bulk_create_apps.md) para documentação completa.

O utilitário em lote aplica o mesmo modelo de segurança do restante do projeto:

- exige `site_id` e `access_type` para toda aplicação
- usa `Sites.Selected` como `Role` para `app_only`
- usa `Sites.Selected` como `Scope` para `delegated`
- não adiciona autorizações de dados no nível do tenant

---

## Visão geral dos módulos

### `auth.py`
Obtém um token Bearer para a Microsoft Graph API usando o fluxo OAuth 2.0 de
**credenciais do cliente** via [MSAL](https://github.com/AzureAD/microsoft-authentication-library-for-python).

Para o fluxo delegado, veja [docs/setup_delegated_auth.md](docs/setup_delegated_auth.md).

### `graph_client.py`
`GraphClient` é um wrapper fino sobre `requests.Session` que injeta o token
Bearer e expõe os helpers `get`, `post`, `patch`, `put_bytes` e `get_raw`.

### `drive.py`
Operações de biblioteca de documentos:

| Função | Descrição |
|---|---|
| `list_drive_items(folder_path)` | Lista os filhos de uma pasta |
| `download_file(item_id, local_path)` | Baixa um arquivo para disco |
| `upload_file(local_path, remote_folder)` | Envia um arquivo local (≤ 4 MB) |
| `read_file_content(item_id)` | Retorna o conteúdo textual do arquivo |
| `write_file_content(item_id, content)` | Sobrescreve o conteúdo textual do arquivo |

### `lists.py`
Operações de listas do SharePoint:

| Função | Descrição |
|---|---|
| `get_list_items(select)` | Recupera todos os itens, opcionalmente selecionando campos |
| `create_list_item(fields)` | Cria um novo item |
| `update_list_item(item_id, fields)` | Atualiza campos de um item existente |

---

## Documentação adicional

- [docs/getting_started.md](docs/getting_started.md) — visão geral, papéis administrativos e permissões
- [docs/setup_portal.md](docs/setup_portal.md) — configuração manual pelo portal
- [docs/setup_cli.md](docs/setup_cli.md) — configuração via Azure CLI e PowerShell
- [docs/setup_delegated_auth.md](docs/setup_delegated_auth.md) — fluxo delegado com login interativo
- [docs/bulk_create_apps.md](docs/bulk_create_apps.md) — criação em lote de aplicações

---

## Licença

Este projeto é licenciado sob a **GNU General Public License v3.0**.
Consulte [LICENSE](LICENSE) para o texto completo.
