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
# criar a migration inicial (ainda não gerada neste scaffold)
dotnet ef migrations add InicialControleDeAcesso -p SwNet.Infraestrutura -s SwNet.Api
dotnet run --project SwNet.Api
```

Ajuste a connection string em `SwNet.Api/appsettings.json` (`ConnectionStrings:DefaultConnection`) se seu SQL Server não for local/Windows Auth.

## Rodando o frontend

```bash
cd SwnetWeb
npm install
npm run dev
```

Abre em `http://localhost:5173`. O `vite.config.ts` já tem proxy de `/api` para `https://localhost:7000` (ajuste a porta conforme o `launchSettings.json` do `SwNet.Api` quando for integrar de verdade).

## O que já está implementado

- Login com autenticação (Strategy pattern por método — hoje só `SISTEMA`/BCrypt) e autorização por permissão (RBAC via claims no JWT)
- Sidebar lateral que expande/retrai (`SwnetWeb/src/components/layout/Sidebar.tsx`)
- Topbar com pesquisa global (input pronto, endpoint ainda não ligado) e troca de planta/fábrica (`TopBar.tsx` + `PlantaContext.tsx`)
- Menu "Acessos Usuários" navegando para tela própria, protegida por `acessos-usuarios.visualizar`, com listagem (fallback mock se a API não responder)
- Dark mode com persistência e leitura da preferência do SO (`ThemeContext.tsx`)

## Pendências conhecidas (próximos passos)

- Migration inicial do EF Core ainda não gerada
- Endpoints de inserir/alterar usuário (`UsuariosController`) — DTOs já prontos em `UsuarioDTOs.cs`
- Endpoint de pesquisa global (`/api/pesquisa`) — input da topbar já existe, só falta ligar
- Refresh token com rotação real (hoje stub 501 em `AuthController.Refresh`)
- Estratégia de autenticação AD/LDAP (`IEstrategiaAutenticacao` já é o ponto de extensão)
