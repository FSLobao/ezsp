<a id="topo"></a>

# MSGraphClient — Camada de abstração MSAL para Microsoft Graph e SharePoint

Uma biblioteca Python para abstrair fluxos de autenticação com **MSAL** e simplificar
a integração com **Microsoft Graph API** em aplicações desktop e mobile.
As operações cobertas incluem gerenciamento de biblioteca de documentos (drive)
e manipulação de listas do SharePoint.

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
		<li><strong>python/</strong>: implementacao do cliente Graph e operacoes de SharePoint
			<ul>
				<li><code>__init__.py</code>: ponto de entrada do pacote para importacoes publicas</li>
				<li><code>auth.py</code>: autenticacao MSAL (app_only/delegated) — valida credenciais e adquire tokens</li>
				<li><code>client.py</code>: ponto de entrada principal — le .env, gerencia sessao HTTP e descoberta do site</li>
				<li><code>drive.py</code>: listagem, upload, download e leitura/escrita de conteudo em bibliotecas de documentos</li>
				<li><code>lists.py</code>: consulta de colunas/views, validacao tipada e operacoes de create/update em itens de lista</li>
			</ul>
		</li>
	</ul>
</details>

<details>
	<summary><strong>tests/</strong>: testes automatizados de unidade e comportamento</summary>
	<ul>
		<li><code>test_auth.py</code>: validacao dos fluxos de autenticacao e aquisicao de token</li>
		<li><code>test_graph_client.py</code>: cobertura do GraphClient (sessao HTTP, helpers get/post/patch)</li>
		<li><code>test_site.py</code>: cobertura da descoberta e metadados do site</li>
		<li><code>test_drive.py</code>: cobertura das operacoes de arquivos e bibliotecas</li>
		<li><code>test_lists.py</code>: cobertura das operacoes em listas, validacao tipada e metadados</li>
	</ul>
</details>

<details>
	<summary><strong>examples/</strong>: scripts de referencia para execucao rapida por caso de uso</summary>
	<ul>
		<li><code>example_site_contents.py</code>: demonstra consulta de conteudo do site (drives e lists)</li>
		<li><code>example_delegated_site_contents.py</code>: mesmo fluxo com autenticacao delegada</li>
		<li><code>example_drive_list.py</code>: demonstra listagem de itens no drive do site</li>
		<li><code>example_drive_download.py</code>: exemplo de download de arquivo remoto para disco local</li>
		<li><code>example_drive_upload.py</code>: exemplo de upload de arquivo local para o SharePoint</li>
		<li><code>example_drive_read_write.py</code>: leitura e sobrescrita de conteudo textual de arquivo</li>
		<li><code>example_list_get.py</code>: consulta de itens de lista para analise e validacao</li>
		<li><code>example_list_create.py</code>: criacao de novos itens em lista SharePoint</li>
		<li><code>example_list_update.py</code>: atualizacao de campos em itens existentes</li>
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
| Python ≥ 3.11 | Testado com 3.14 |
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
cd MSGraphClient
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

Variáveis opcionais para modo de autenticação:

| Variável | Descrição |
|---|---|
| `GRAPH_AUTH_MODE` | Modo de autenticação (`client_credentials` ou `delegated`) |
| `AZURE_REDIRECT_URI` | URI de redirecionamento para login interativo em modo `delegated` |
| `GRAPH_DELEGATED_LOGIN_MODE` | `interactive` (browser local) ou `device_code` |
| `GRAPH_DELEGATED_SCOPES` | Escopos delegados (separados por espaço ou vírgula) |

> **Encontrando IDs** — veja [docs/getting_started.md](docs/getting_started.md).

### 3. Escolher o modelo de autenticação

- **`client_credentials`**: indicado para automação sem interação do usuário.
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

### `client.py`
`GraphClient` é o ponto de entrada principal. Lê variáveis de ambiente do `.env`,
cria um `GraphAuthenticator` interno, gerencia a sessão HTTP autenticada e expõe
helpers de requisição (`get`, `post`, `patch`, `put_bytes`, `get_raw`) além de
métodos de descoberta do site.

| Método / Atributo | Descrição |
|---|---|
| `GraphClient()` | Lê `.env`, autentica e carrega metadados do site |
| `client.site_graph_id` | ID Graph do site conectado |
| `client.site_name` | Nome interno do site |
| `client.site_display_name` | Nome de exibição do site |
| `client.site_web_url` | URL do site no SharePoint |
| `client.site_drives` | Lista de drives do site |
| `client.site_lists` | Lista de lists do site |
| `client.get_site_contents()` | Retorna metadados do site, drives e lists |
| `client.refresh_site_info()` | Recarrega metadados do site a partir do Graph |

### `auth.py`
`GraphAuthenticator` é responsável exclusivamente pela autenticação MSAL.
Recebe credenciais explícitas (não lê `.env`) e adquire tokens OAuth 2.0
via fluxo de credenciais do cliente ou fluxo delegado interativo.

Para o fluxo delegado, veja [docs/setup_delegated_auth.md](docs/setup_delegated_auth.md).

### `drive.py`
Operações de biblioteca de documentos. Requer `drive_id` explícito na construção.

```python
drive = GraphDrive(drive_id=os.environ["SHAREPOINT_DRIVE_ID"], client=client)
```

| Método | Descrição |
|---|---|
| `GraphDrive.list_drive_items(folder_path)` | Lista os filhos de uma pasta |
| `GraphDrive.download_file(item_id, local_path)` | Baixa um arquivo para disco |
| `GraphDrive.upload_file(local_path, remote_folder)` | Envia um arquivo local (≤ 4 MB) |
| `GraphDrive.read_file_content(item_id)` | Retorna o conteúdo textual do arquivo |
| `GraphDrive.write_file_content(item_id, content)` | Sobrescreve o conteúdo textual do arquivo |

### `lists.py`
Operações de listas do SharePoint. Requer `list_id` explícito na construção.

```python
list_client = GraphList(list_id=os.environ["SHAREPOINT_LIST_ID"], client=client)
```

| Método | Descrição |
|---|---|
| `GraphList.get_columns(names=None)` | Recupera colunas da lista, opcionalmente filtradas por nome |
| `GraphList.get_views()` | Lista as views da lista |
| `GraphList.get_view_columns(view_id)` | Lista as colunas visíveis em uma view |
| `GraphList.get_schema()` | Retorna schema de colunas editáveis (display_name, tipo e regras) |
| `GraphList.get_field_types()` | Retorna mapeamento `displayName -> type` |
| `GraphList.get_items(select=None, include_id=True)` | Recupera itens com chaves `displayName` e paginação automática |
| `GraphList.get_items_dataframe(select=None, include_id=True)` | Recupera itens diretamente em DataFrame pandas |
| `GraphList.get_item_template(include_optional=True)` | Gera template de item com colunas editáveis |
| `GraphList.validate_item(data)` | Valida tipos e regras antes de persistir |
| `GraphList.save_item(data)` | Cria ou atualiza item conforme presença de `_id` |
| `GraphList.save_items(items)` | Persiste múltiplos itens, interrompendo no primeiro erro |
| `GraphList.save_dataframe(dataframe)` | Persiste linhas de DataFrame no formato da API nova |

A validação em `validate_item` utiliza metadados extraídos da definição de coluna no Graph:

| Tipo | Regras aplicadas |
|---|---|
| `text` | Deve ser `str`; rejeita quebras de linha; respeita `max_length` da coluna (padrão 255) |
| `note` | Deve ser `str`; permite quebras de linha; respeita `max_length` (padrão 63999) |
| `number` | Deve ser `int`/`float`; rejeita valores fora de `minimum`/`maximum` quando definidos |
| `boolean` | Deve ser `bool` |
| `dateTime` | Deve ser `str`, `datetime` ou `date`; valida parsing ISO |
| `choice` | Deve ser `str`; rejeita valores fora da lista quando `allowTextEntry` é `False` |

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
