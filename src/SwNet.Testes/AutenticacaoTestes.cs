using System.IdentityModel.Tokens.Jwt;
using System.Net;
using System.Net.Http.Json;
using Microsoft.EntityFrameworkCore;
using SwNet.Aplicacao.Autorizacao;
using SwNet.Aplicacao.DTOs.Autenticacao;
using SwNet.Dominio.Entidades;

namespace SwNet.Testes;

public class AutenticacaoTestes
{
    [Fact]
    public async Task Login_valido_retorna_tokens_e_permissoes_do_perfil()
    {
        using var api = new ApiFactory();

        var retorno = await api.LogarAsync();

        Assert.True(retorno.Sucesso);
        Assert.False(string.IsNullOrEmpty(retorno.Token));
        Assert.False(string.IsNullOrEmpty(retorno.RefreshToken));
        Assert.Contains(Permissoes.AcessosUsuariosEditar, retorno.Usuario!.Permissoes);
        Assert.Equal(2, retorno.Usuario.Plantas.Count);

        var jwt = new JwtSecurityTokenHandler().ReadJwtToken(retorno.Token);
        Assert.Equal(retorno.Usuario.Id.ToString(), jwt.Subject);
        Assert.Contains(jwt.Claims, c => c.Type == ClaimsSwNet.Permissao && c.Value == Permissoes.AcessosUsuariosVisualizar);
    }

    [Fact]
    public async Task Login_com_senha_errada_retorna_401()
    {
        using var api = new ApiFactory();

        var resposta = await api.CreateClient().PostAsJsonAsync("/api/auth/login", new { login = "admin", senha = "errada" });

        Assert.Equal(HttpStatusCode.Unauthorized, resposta.StatusCode);
        var corpo = await resposta.Content.ReadFromJsonAsync<LoginRetornoDTO>();
        Assert.False(corpo!.Sucesso);
        Assert.Null(corpo.Token);
    }

    [Fact]
    public async Task Bloqueia_usuario_apos_5_tentativas_invalidas()
    {
        using var api = new ApiFactory();
        var cliente = api.CreateClient();

        for (var i = 0; i < 5; i++)
            await cliente.PostAsJsonAsync("/api/auth/login", new { login = "admin", senha = "errada" });

        // Mesmo com a senha certa, continua bloqueado
        var resposta = await cliente.PostAsJsonAsync("/api/auth/login", new { login = "admin", senha = "Admin@123" });
        Assert.Equal(HttpStatusCode.Unauthorized, resposta.StatusCode);
        var corpo = await resposta.Content.ReadFromJsonAsync<LoginRetornoDTO>();
        Assert.Contains("bloqueado", corpo!.MensagemErro);
    }

    [Fact]
    public async Task Endpoint_protegido_exige_token()
    {
        using var api = new ApiFactory();

        Assert.Equal(HttpStatusCode.Unauthorized, (await api.CreateClient().GetAsync("/api/auth/me")).StatusCode);
        Assert.Equal(HttpStatusCode.Unauthorized, (await api.CreateClient().GetAsync("/api/plantas/minhas")).StatusCode);

        var login = await api.LogarAsync();
        var me = await api.ClienteAutenticado(login.Token!).GetFromJsonAsync<UsuarioLogadoDTO>("/api/auth/me");
        Assert.Equal("admin", me!.Login);
    }

    [Fact]
    public async Task Token_adulterado_e_rejeitado()
    {
        using var api = new ApiFactory();
        var login = await api.LogarAsync();

        var adulterado = login.Token![..^4] + "AAAA";
        var resposta = await api.ClienteAutenticado(adulterado).GetAsync("/api/auth/me");

        Assert.Equal(HttpStatusCode.Unauthorized, resposta.StatusCode);
    }

    [Fact]
    public async Task RequerPermissao_retorna_403_para_usuario_sem_a_permissao()
    {
        using var api = new ApiFactory();
        await CriarUsuarioAsync(api, "operador", "Operador@123", perfil: "Operador");

        var admin = await api.LogarAsync();
        var operador = await api.LogarAsync("operador", "Operador@123");

        Assert.Equal(HttpStatusCode.OK, (await api.ClienteAutenticado(admin.Token!).GetAsync("/api/usuarios")).StatusCode);
        Assert.Equal(HttpStatusCode.Forbidden, (await api.ClienteAutenticado(operador.Token!).GetAsync("/api/usuarios")).StatusCode);
    }

    [Fact]
    public async Task Permissao_negada_no_usuario_sobrepoe_o_perfil()
    {
        using var api = new ApiFactory();
        await api.ExecutarNoBancoAsync(async db =>
        {
            var admin = await db.Usuarios.SingleAsync(u => u.Login == "admin");
            var permissao = await db.Permissoes.SingleAsync(p => p.Chave == Permissoes.AcessosUsuariosVisualizar);
            db.UsuarioPermissoes.Add(new UsuarioPermissao { UsuarioId = admin.Id, PermissaoId = permissao.Id, Concedida = false });
            await db.SaveChangesAsync();
        });

        var login = await api.LogarAsync();

        Assert.DoesNotContain(Permissoes.AcessosUsuariosVisualizar, login.Usuario!.Permissoes);
        Assert.Equal(HttpStatusCode.Forbidden, (await api.ClienteAutenticado(login.Token!).GetAsync("/api/usuarios")).StatusCode);
    }

    [Fact]
    public async Task Refresh_rotaciona_o_token_e_reuso_revoga_a_familia()
    {
        using var api = new ApiFactory();
        var cliente = api.CreateClient();
        var login = await api.LogarAsync();

        // 1ª renovação: ok, recebe um refresh novo
        var r1 = await cliente.PostAsJsonAsync("/api/auth/refresh", new { refreshToken = login.RefreshToken });
        Assert.Equal(HttpStatusCode.OK, r1.StatusCode);
        var renovado = (await r1.Content.ReadFromJsonAsync<LoginRetornoDTO>())!;
        Assert.NotEqual(login.RefreshToken, renovado.RefreshToken);

        // Reuso do token antigo: rejeitado...
        var reuso = await cliente.PostAsJsonAsync("/api/auth/refresh", new { refreshToken = login.RefreshToken });
        Assert.Equal(HttpStatusCode.Unauthorized, reuso.StatusCode);

        // ...e derruba também o token novo da mesma família
        var r2 = await cliente.PostAsJsonAsync("/api/auth/refresh", new { refreshToken = renovado.RefreshToken });
        Assert.Equal(HttpStatusCode.Unauthorized, r2.StatusCode);
    }

    [Fact]
    public async Task Refresh_nao_e_armazenado_em_claro()
    {
        using var api = new ApiFactory();
        var login = await api.LogarAsync();

        await api.ExecutarNoBancoAsync(async db =>
        {
            Assert.False(await db.RefreshTokens.AnyAsync(rt => rt.TokenHash == login.RefreshToken));
            Assert.Equal(1, await db.RefreshTokens.CountAsync());
        });
    }

    [Fact]
    public async Task Logout_invalida_o_refresh_token()
    {
        using var api = new ApiFactory();
        var cliente = api.CreateClient();
        var login = await api.LogarAsync();

        var logout = await cliente.PostAsJsonAsync("/api/auth/logout", new { refreshToken = login.RefreshToken });
        Assert.Equal(HttpStatusCode.NoContent, logout.StatusCode);

        var refresh = await cliente.PostAsJsonAsync("/api/auth/refresh", new { refreshToken = login.RefreshToken });
        Assert.Equal(HttpStatusCode.Unauthorized, refresh.StatusCode);
    }

    [Fact]
    public async Task Trocar_planta_reemite_token_com_nova_planta_e_mantem_no_refresh()
    {
        using var api = new ApiFactory();
        var login = await api.LogarAsync();
        var outraPlanta = login.Usuario!.Plantas.Single(p => p.Id != login.Usuario.PlantaAtualId);
        var cliente = api.ClienteAutenticado(login.Token!);

        var resposta = await cliente.PostAsJsonAsync("/api/auth/trocar-planta",
            new { plantaId = outraPlanta.Id, refreshToken = login.RefreshToken });
        Assert.Equal(HttpStatusCode.OK, resposta.StatusCode);
        var trocado = (await resposta.Content.ReadFromJsonAsync<LoginRetornoDTO>())!;
        Assert.Equal(outraPlanta.Id, trocado.Usuario!.PlantaAtualId);

        var renovado = await (await api.CreateClient().PostAsJsonAsync("/api/auth/refresh", new { refreshToken = trocado.RefreshToken }))
            .Content.ReadFromJsonAsync<LoginRetornoDTO>();
        Assert.Equal(outraPlanta.Id, renovado!.Usuario!.PlantaAtualId);

        // Planta sem vínculo → 403
        var semAcesso = await api.ClienteAutenticado(renovado.Token!).PostAsJsonAsync("/api/auth/trocar-planta",
            new { plantaId = 999, refreshToken = renovado.RefreshToken });
        Assert.Equal(HttpStatusCode.Forbidden, semAcesso.StatusCode);
    }

    [Fact]
    public async Task Alterar_senha_valida_politica_e_encerra_sessoes()
    {
        using var api = new ApiFactory();
        var login = await api.LogarAsync();
        var cliente = api.ClienteAutenticado(login.Token!);

        var fraca = await cliente.PostAsJsonAsync("/api/auth/alterar-senha", new { senhaAtual = "Admin@123", novaSenha = "123" });
        Assert.Equal(HttpStatusCode.BadRequest, fraca.StatusCode);

        var ok = await cliente.PostAsJsonAsync("/api/auth/alterar-senha", new { senhaAtual = "Admin@123", novaSenha = "NovaSenha2026" });
        Assert.Equal(HttpStatusCode.OK, ok.StatusCode);

        var refresh = await api.CreateClient().PostAsJsonAsync("/api/auth/refresh", new { refreshToken = login.RefreshToken });
        Assert.Equal(HttpStatusCode.Unauthorized, refresh.StatusCode);

        var novoLogin = await api.LogarAsync("admin", "NovaSenha2026");
        Assert.True(novoLogin.Sucesso);
    }

    private static Task CriarUsuarioAsync(ApiFactory api, string login, string senha, string perfil) =>
        api.ExecutarNoBancoAsync(async db =>
        {
            var novoPerfil = new Perfil { Nome = perfil };
            db.Perfis.Add(novoPerfil);
            await db.SaveChangesAsync();

            var usuario = new Usuario
            {
                Login = login,
                NomeCompleto = login,
                Email = $"{login}@swnet.local",
                SenhaHash = BCrypt.Net.BCrypt.HashPassword(senha),
                PerfilId = novoPerfil.Id
            };
            db.Usuarios.Add(usuario);
            await db.SaveChangesAsync();

            db.UsuarioPlantas.Add(new UsuarioPlanta { UsuarioId = usuario.Id, PlantaId = (await db.Plantas.FirstAsync()).Id, Padrao = true });
            await db.SaveChangesAsync();
        });
}
