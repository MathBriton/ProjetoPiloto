# SWnet — Scaffold inicial

Estrutura decidida: backend em `src/`, frontend em `SwnetWeb/`.

## Stack

- **Backend**: .NET 10 Web API + EF Core + SQL Server, JWT, RBAC por permissão (`[RequerPermissao]`)
- **Frontend**: React 19 + TypeScript + Vite + Tailwind v4, dark mode desde o início
- **Identidade visual**: branco + azul (tokens `swnet-*` em `SwnetWeb/src/index.css`)

## Login para testar (fictício)

```
usuário: admin
senha:   Admin@123
```

- No **frontend**, o login já está mockado (`USAR_MOCK = true` em `SwnetWeb/src/services/autenticacaoApi.ts`), então dá pra testar a navegação, sidebar, troca de planta e dark mode sem precisar do backend/banco rodando.
- No **backend**, esse mesmo usuário é criado automaticamente pelo seed de desenvolvimento (`SwNet.Infraestrutura/Contexto/SwNetDbSeeder.cs`) na primeira vez que a API sobe em ambiente `Development` — desde que o banco `SWNET` já exista (rode as migrations antes, ver abaixo).

Quando o backend estiver pronto para uso real, troque `USAR_MOCK` para `false` em `autenticacaoApi.ts`.

## Rodando o backend

```bash
cd src
dotnet restore
dotnet tool restore            # instala o dotnet-ef local (dotnet-tools.json)
dotnet run --project SwNet.Api # em Development aplica as migrations e roda o seed
dotnet test                    # testes de integração de autenticação/autorização (SQLite em memória)
```

Ajuste a connection string em `SwNet.Api/appsettings.json` (`ConnectionStrings:DefaultConnection`) se seu SQL Server não for local/Windows Auth.

A chave JWT de desenvolvimento fica em `appsettings.Development.json`. Fora de Development a API **não sobe** sem `Jwt:ChaveSecreta` (mínimo 32 caracteres) — defina por variável de ambiente (`Jwt__ChaveSecreta`) ou `dotnet user-secrets`.

## Autenticação e autorização

| Endpoint | Acesso | O que faz |
|---|---|---|
| `POST /api/auth/login` | anônimo, rate limit | Valida usuário/senha (BCrypt), bloqueia após 5 falhas por 15 min, devolve access token (JWT, 15 min) + refresh token (7 dias) |
| `POST /api/auth/refresh` | anônimo, rate limit | Rotação: o refresh enviado é revogado e um novo é emitido. Reapresentar um refresh já usado revoga a sessão inteira (detecção de roubo) |
| `POST /api/auth/logout` | anônimo | Revoga o refresh token da sessão |
| `POST /api/auth/logout-todas` | logado | Revoga todas as sessões do usuário |
| `GET /api/auth/me` | logado | Usuário + permissões atuais lidas do banco |
| `POST /api/auth/trocar-planta` | logado | Reemite os tokens com outra planta (só plantas vinculadas ao usuário) |
| `POST /api/auth/alterar-senha` | logado | Valida a política de senha e encerra todas as sessões |

- **Tudo exige login por padrão** (`FallbackPolicy`); endpoints públicos usam `[AllowAnonymous]`.
- **Permissões**: `[RequerPermissao(Permissoes.AcessosUsuariosEditar)]` — a policy é criada sob demanda pelo `PermissaoPolicyProvider`, não precisa registrar no `Program.cs`. Novas chaves vão no catálogo `SwNet.Aplicacao/Autorizacao/Permissoes.cs` (o seed sincroniza com o banco).
- **Permissões efetivas** = perfil + concedidas ao usuário − negadas ao usuário (`UsuarioPermissao.Concedida = false` sempre vence).
- O refresh token é guardado no banco só como hash SHA-256 (tabela `RefreshTokens`).
- No frontend, `apiClient.ts` renova o token automaticamente ao receber 401 (uma renovação por vez) e volta para o login se a sessão tiver sido revogada.

## Rodando o frontend

```bash
cd SwnetWeb
npm install
npm run dev
```

Abre em `http://localhost:5173`. O `vite.config.ts` já tem proxy de `/api` para `https://localhost:7000` (ajuste a porta conforme o `launchSettings.json` do `SwNet.Api` quando for integrar de verdade).

## O que já está implementado

- Login com autenticação (Strategy pattern por método — hoje só `SISTEMA`/BCrypt), refresh token com rotação, logout, troca de senha e autorização por permissão (RBAC via claims no JWT) — ver seção acima
- Sidebar lateral que expande/retrai (`SwnetWeb/src/components/layout/Sidebar.tsx`)
- Topbar com pesquisa global (input pronto, endpoint ainda não ligado) e troca de planta/fábrica (`TopBar.tsx` + `PlantaContext.tsx`)
- Menu "Acessos Usuários" navegando para tela própria, protegida por `acessos-usuarios.visualizar`, com listagem (fallback mock se a API não responder)
- Dark mode com persistência e leitura da preferência do SO (`ThemeContext.tsx`)

## Pendências conhecidas (próximos passos)

- Endpoints de inserir/alterar usuário (`UsuariosController`) — DTOs já prontos em `UsuarioDTOs.cs`
- Endpoint de pesquisa global (`/api/pesquisa`) — input da topbar já existe, só falta ligar
- Troca de planta no frontend chamar `POST /api/auth/trocar-planta` (hoje só troca localmente)
- Tela para forçar troca de senha quando `usuario.senhaProvisoria = true`
- Estratégia de autenticação AD/LDAP (`IEstrategiaAutenticacao` já é o ponto de extensão)
