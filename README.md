<a id="topo"></a>

# MSGraphTest — SharePoint via Microsoft Graph API

Um projeto de teste em Python demonstrando como acessar o SharePoint através da
**Microsoft Graph API** usando MSAL para autenticação. As operações abordadas
incluem gerenciamento de biblioteca de documentos (drive) e manipulação de listas do SharePoint.

O projeto adota **privilégio mínimo** como regra: o acesso ao SharePoint é feito com
`Sites.Selected`, sempre restrito a sites explicitamente inscritos. Isso vale tanto
para autenticação `app_only` quanto para autenticação `delegated`.

Licenciado sob a [GNU General Public License v3.0](LICENSE).

---

<details>
	<summary><strong>Índice</strong></summary>
	<ul>
		<li><a href="#sec-estrutura-do-projeto">Estrutura do projeto</a></li>
		<li><a href="#sec-pre-requisitos">Pre-requisitos</a></li>
		<li><a href="#sec-inicio-rapido">Inicio rapido</a></li>
		<li><a href="#sec-notebooks-interativos">Notebooks interativos</a></li>
		<li><a href="#sec-executando-testes">Executando testes</a></li>
		<li><a href="#sec-criacao-em-lote">Criacao em lote de aplicacoes Azure AD</a></li>
		<li><a href="#sec-visao-geral-modulos">Visao geral dos modulos</a></li>
		<li><a href="#sec-documentacao-adicional">Documentacao adicional</a></li>
		<li><a href="#sec-licenca">Licenca</a></li>
	</ul>
</details>

---

<a id="sec-estrutura-do-projeto"></a>

## Estrutura do projeto

Este repositório apresenta a seguinte organização:

<details>
	<summary><strong>src/</strong>: codigo-fonte principal do pacote e scripts utilitarios</summary>
	<ul>
		<li><strong>bulkCreate/</strong>: automacao de criacao em lote de aplicacoes no Entra ID
			<ul>
				<li><code>bulk_create_apps.py</code>: script Python para provisionar apps e permissoes de forma padronizada</li>
				<li><code>Bulk-CreateApps.ps1</code>: alternativa PowerShell para execucao operacional em ambientes Windows</li>
			</ul>
		</li>
		<li><strong>msgraphtest/</strong>: implementacao do cliente Graph e operacoes de SharePoint
			<ul>
				<li><code>__init__.py</code>: ponto de entrada do pacote para importacoes publicas</li>
				<li><code>auth.py</code>: autenticacao (app_only/delegated), sessao HTTP e descoberta de metadados do site</li>
				<li><code>drive.py</code>: listagem, upload, download e leitura/escrita de conteudo em bibliotecas de documentos</li>
				<li><code>lists.py</code>: consulta de colunas/views e operacoes de create/update em itens de lista</li>
			</ul>
		</li>
	</ul>
</details>

<details>
	<summary><strong>tests/</strong>: testes automatizados de unidade e comportamento</summary>
	<ul>
		<li><code>test_auth.py</code>: validacao dos fluxos de autenticacao e descoberta de site</li>
		<li><code>test_drive.py</code>: cobertura das operacoes de arquivos e bibliotecas</li>
		<li><code>test_lists.py</code>: cobertura das operacoes em listas e itens</li>
	</ul>
</details>

<details>
	<summary><strong>examples/</strong>: scripts de referencia para execucao rapida por caso de uso</summary>
	<ul>
		<li><code>example_drive_list.py</code>: demonstra listagem de itens no drive do site</li>
		<li><code>example_drive_download.py</code>: exemplo de download de arquivo remoto para disco local</li>
		<li><code>example_drive_upload.py</code>: exemplo de upload de arquivo local para o SharePoint</li>
		<li><code>example_drive_read_write.py</code>: leitura e sobrescrita de conteudo textual de arquivo</li>
		<li><code>example_list_get.py</code>: consulta de itens de lista para analise e validacao</li>
		<li><code>example_list_create.py</code>: criacao de novos itens em lista SharePoint</li>
		<li><code>example_list_update.py</code>: atualizacao de campos em itens existentes</li>
		<li><code>bulk_create_example.json</code>: modelo de entrada para o fluxo de criacao em lote</li>
	</ul>
</details>

<details>
	<summary><strong>notebooks/</strong>: validacao interativa passo a passo em ambiente exploratorio</summary>
	<ul>
		<li><code>graph_auth_site_attributes.ipynb</code>: roteiro end-to-end para autenticar, inspecionar site e testar edicao de conteudo</li>
	</ul>
</details>

<details>
	<summary><strong>docs/</strong>: guias de configuracao, operacao e referencia do projeto</summary>
	<ul>
		<li><code>getting_started.md</code>: visao geral do fluxo, prerequisitos e primeiros testes</li>
		<li><code>setup_cli.md</code>: configuracao do ambiente usando Azure CLI e PowerShell</li>
		<li><code>setup_portal.md</code>: configuracao manual no portal do Azure</li>
		<li><code>setup_delegated_auth.md</code>: orientacoes para autenticacao delegada com usuario</li>
		<li><code>bulk_create_apps.md</code>: detalhes do processo de provisionamento em lote</li>
	</ul>
</details>

<details>
	<summary><strong>downloads/</strong>: artefatos locais de teste (diretorio ignorado no git)</summary>
	<ul>
		<li>Armazena downloads, uploads de teste e arquivos temporarios gerados nas validacoes</li>
	</ul>
</details>

<ul>
	<li><code>.env.example</code>: modelo de configuracao de ambiente para armazenamento de tokens e senhas</li>
	<li><code>pyproject.toml</code>: manifesto do projeto com dependencias, metadados e configuracoes de ferramentas</li>
	<li><code>LICENSE</code>: termos de licenciamento do repositorio sob GPL v3.0</li>
</ul>

[⬆ Voltar ao topo](#topo)

---

<a id="sec-pre-requisitos"></a>

## Pré-requisitos

| Requisito | Observações |
|---|---|
| Python ≥ 3.11 | Testado com 3.11+ |
| [UV](https://docs.astral.sh/uv/) | Gerenciador de pacotes e ambiente virtual |
| Registro de aplicativo no Microsoft Entra ID | Configure `Sites.Selected` e inscreva os sites necessários |

> Este repositório **não usa** permissões amplas como `Sites.Read.All` ou `Sites.ReadWrite.All` para acesso a dados no SharePoint.

[⬆ Voltar ao topo](#topo)

---

<a id="sec-inicio-rapido"></a>

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

[⬆ Voltar ao topo](#topo)

---

<a id="sec-notebooks-interativos"></a>

## Notebooks interativos (alternativa aos examples)

Além dos scripts em `examples/`, você pode usar notebooks para validar o fluxo de forma
interativa, inspecionando respostas e DataFrames a cada etapa.

Notebook principal:

- `notebooks/graph_auth_site_attributes.ipynb`

Esse notebook executa um fluxo end-to-end para:

1. carregar `.env` e autenticar no Graph;
2. consultar atributos e conteúdo do site (drives/lists);
3. testar operações de conteúdo no drive (write, update, load e download);
4. testar criação e atualização de itens em lista com visualização tabular.

Quando usar notebooks em vez dos examples:

- quando você quer depurar autenticação passo a passo;
- quando precisa validar transformação e inspeção de dados em DataFrames;
- quando deseja testar rapidamente mudanças no processo de edição de conteúdo.

> [!WARNING]
> **Boas práticas de higiene (credenciais e dados sensíveis)**
>
> - Nunca imprima valores de variáveis sensíveis do `.env` (como `AZURE_CLIENT_SECRET`).
> - Evite exibir tokens, headers de autorização ou payloads contendo segredos.
> - Limpe os outputs do notebook antes de commit (`Clear All Outputs`) para não versionar dados sensíveis.
> - Mantenha o arquivo `.env` fora do versionamento e use apenas `.env.example` no repositório.
> - Se houver exposição acidental de segredo em output/código, revogue e gere novas credenciais imediatamente.

[⬆ Voltar ao topo](#topo)

---

<a id="sec-executando-testes"></a>

## Executando testes

```bash
uv run pytest
```

Relatório de cobertura é impresso automaticamente. Os testes usam mocking e **não**
requerem credenciais reais.

[⬆ Voltar ao topo](#topo)

---

<a id="sec-criacao-em-lote"></a>

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

[⬆ Voltar ao topo](#topo)

---

<a id="sec-visao-geral-modulos"></a>

## Visão geral dos módulos

### `auth.py`
Obtém um token Bearer para a Microsoft Graph API usando o fluxo OAuth 2.0 de
**credenciais do cliente** via [MSAL](https://github.com/AzureAD/microsoft-authentication-library-for-python).

Para o fluxo delegado, veja [docs/setup_delegated_auth.md](docs/setup_delegated_auth.md).

### `auth.py`
`GraphClient` é o cliente principal do Microsoft Graph. Ele gerencia a sessão
HTTP autenticada, expõe os helpers `get`, `post`, `patch`, `put_bytes` e
`get_raw`, e possui um `GraphAuthenticator` associado para descoberta do site.

### `drive.py`
Operações de biblioteca de documentos:

| Método | Descrição |
|---|---|
| `GraphDrive.list_drive_items(folder_path)` | Lista os filhos de uma pasta |
| `GraphDrive.download_file(item_id, local_path)` | Baixa um arquivo para disco |
| `GraphDrive.upload_file(local_path, remote_folder)` | Envia um arquivo local (≤ 4 MB) |
| `GraphDrive.read_file_content(item_id)` | Retorna o conteúdo textual do arquivo |
| `GraphDrive.write_file_content(item_id, content)` | Sobrescreve o conteúdo textual do arquivo |

### `lists.py`
Operações de listas do SharePoint:

| Método | Descrição |
|---|---|
| `GraphList.get_list_columns(names)` | Recupera colunas da lista, opcionalmente filtradas por nome |
| `GraphList.get_list_views()` | Lista as views da lista |
| `GraphList.get_list_view_columns(view_id)` | Lista as colunas visíveis em uma view |
| `GraphList.get_list_items(select, fields_only=False, include_title=False, include_item_id=False)` | Recupera itens da lista com seleção opcional de campos |
| `GraphList.create_list_item(fields)` | Cria um novo item |
| `GraphList.update_list_item(item_id, fields)` | Atualiza campos de um item existente |

### `auth.py`
`GraphAuthenticator` concentra a descoberta do site:

| Método | Descrição |
|---|---|
| `GraphAuthenticator.get_site_contents()` | Retorna metadados do site, drives e lists |
| `GraphAuthenticator.list_site_drives()` | Lista os drives do site |
| `GraphAuthenticator.list_site_lists()` | Lista as lists do site |

[⬆ Voltar ao topo](#topo)

---

<a id="sec-documentacao-adicional"></a>

## Documentação adicional

- [docs/getting_started.md](docs/getting_started.md) — visão geral, papéis administrativos e permissões
- [docs/setup_portal.md](docs/setup_portal.md) — configuração manual pelo portal
- [docs/setup_cli.md](docs/setup_cli.md) — configuração via Azure CLI e PowerShell
- [docs/setup_delegated_auth.md](docs/setup_delegated_auth.md) — fluxo delegado com login interativo
- [docs/bulk_create_apps.md](docs/bulk_create_apps.md) — criação em lote de aplicações

[⬆ Voltar ao topo](#topo)

---

<a id="sec-licenca"></a>

## Licença

Este projeto é licenciado sob a **GNU General Public License v3.0**.
Consulte [LICENSE](LICENSE) para o texto completo.

[⬆ Voltar ao topo](#topo)
