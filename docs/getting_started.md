# Guia de Início

Este guia descreve perfis e permissões necessárias e diferentes escopos que podem ser utilizados para acessar o SharePoint via Microsoft Graph API. Ele é destinado a desenvolvedores e administradores que estão configurando ou gerenciando o acesso de aplicativos ao SharePoint usando a Microsoft Graph API.

Guias adicionais são providos para configuração passo a passo, de acordo com a abordagem preferida:

**Autenticação por credenciais do cliente (app-only) — execução sem supervisão:**
- [**Via Portal e Navegador**](setup_portal.md) — usando portal do Azure e Microsoft Graph Explorer, sem instalação de ferramentas de linha de comando.
- [**Via Azure CLI e PowerShell**](setup_cli.md) — usando linha de comando.

**Autenticação delegada (usuário) — com login interativo e trilha de auditoria por usuário:**
- [**Autenticação Delegada**](setup_delegated_auth.md) — fluxo de código de autorização OAuth 2.0; o aplicativo age em nome do usuário autenticado.

---

## Criação em lote de aplicações

Se você precisa criar **múltiplas aplicações Azure AD** automaticamente, use:

- **PowerShell (recomendado)**: `.\.\src/bulkCreate/Bulk-CreateApps.ps1 -InputPath config.json`
  - Já instalado no Windows, requer só Microsoft.Graph SDK
- **Python (alternativa)**: `python -m bulkCreate.bulk_create_apps config.json`
  - Requer Azure CLI ou PowerShell SDK instalados

**Dica de sessão**: Autentique uma vez com `Connect-MgGraph` e execute múltiplas vezes com `-SkipLogin` (PowerShell) ou `--skip-login` (Python).

Veja [bulk_create_apps.md](bulk_create_apps.md) para instruções detalhadas, exemplos de entrada JSON e como recuperar as credenciais geradas.

---

## Funções envolvidas

Três funções distintas participam da configuração de acesso de aplicativo ao SharePoint
através da Microsoft Graph API. As etapas na Parte 2 são marcadas com o distintivo da
função responsável.

| Distintivo | Função | Função no Microsoft Entra (PT) | Função no Microsoft Entra (EN) | Descrição |
|---|---|---|---|---|
| 🔧 **Equipe de desenvolvimento** | Desenvolvedor de aplicativo | **Desenvolvedor de aplicativos** (ou usuário comum, se o tenant permitir registro de apps) | **Application Developer** (or regular user, if tenant allows app registration) | Cria o registro de aplicativo e escreve o código. |
| 🔑 **Administrador Entra** | Administrador do Azure AD / Entra ID | **Administrador global** ou **Administrador de função com privilégios** | **Global Administrator** or **Privileged Role Administrator** | Concede consentimento de administrador para permissões de API no portal do Azure. |
| 🛡️ **Administrador SP** | Administrador do tenant (locatário) do SharePoint | **Administrador do SharePoint** (portal do Azure) | **SharePoint Administrator** (portal) / **SharePoint Service Administrator** (Graph API e PowerShell) | Inscreve sites específicos para o aplicativo via Graph API. |

> ⚠️ **Nota importante sobre o perfil 🔑 Administrador Entra:** A função **Administrador de Aplicativos** (*Application Administrator*) e a função **Administrador de Aplicativos de Nuvem** (*Cloud Application Administrator*) **não são suficientes** para conceder consentimento para `Sites.Selected`. A documentação oficial da Microsoft estabelece uma exceção explícita para essas funções: elas não podem consentir com permissões de aplicativo do **Microsoft Graph** — categoria à qual `Sites.Selected` pertence. Somente o **Administrador global** ou o **Administrador de função com privilégios** possuem a permissão `managePermissionGrantsForAll.microsoft-company-admin` necessária para este consentimento.

> ⚠️ **Nota importante sobre o perfil 🛡️ Administrador SP:** A função necessária é **Administrador do SharePoint** no nível do tenant (locatário) do Microsoft 365. **Isso não é o mesmo que um administrador de coleção de sites (proprietário de site).** Um proprietário de site gerencia usuários dentro de um site, mas não pode conceder acesso à API do Graph — isso requer direitos de administrador do SharePoint no nível do tenant.

## Permissões disponíveis da API Microsoft Graph para SharePoint

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

## Fluxo de autenticação: credenciais do cliente vs. autenticação delegada (usuário)

Este projeto suporta os fluxos OAuth 2.0 de **credenciais do cliente (client credentials)** e **autenticação delegada (authorization code)**. Em ambos os casos, o acesso aos dados do SharePoint continua restrito por `Sites.Selected` e por inscrição explícita de cada site.

No fluxo de **credenciais do cliente**, o aplicativo se autentica como ele mesmo (o registro de aplicativo) sem contexto de usuário:

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

Ambos os fluxos usam o mesmo modelo de permissão `Sites.Selected`, exigem inscrição explícita do site e evitam autorizações amplas no nível do tenant. A diferença é **quem o aplicativo representa** (ele mesmo vs. um usuário) e como essa identidade aparece nos logs de auditoria.

No fluxo delegado, o acesso efetivo em runtime é a interseção entre: 1) a concessão `read` ou `write` dada ao aplicativo no site e 2) as permissões que o usuário autenticado já possui nesse mesmo site.

Mudar para autenticação delegada requer:

1. Mudança da aquisição de token em `src/msgraphclient/auth.py` do endpoint de
   credenciais do cliente para o fluxo de código de autorização (usando uma biblioteca como
   [MSAL for Python](https://github.com/AzureAD/microsoft-authentication-library-for-python)
   com `acquire_token_interactive()`).
2. Adição de uma **URI de Redirecionamento** no registro de aplicativo (ex:
   `http://localhost:8000`) para tratar o callback OAuth.
3. Usuários fazendo login quando o aplicativo executa.

Para a maioria dos cenários, **credenciais do cliente é mais simples**. Use
autenticação delegada apenas se seus requisitos de governança ou conformidade mandatarem uma
trilha de auditoria vinculando cada ação do SharePoint a um usuário nomeado.
