using System.Net.Http.Headers;
using System.Net.Http.Json;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc.Testing;
using Microsoft.Data.Sqlite;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Infrastructure;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection.Extensions;
using SwNet.Aplicacao.DTOs.Autenticacao;
using SwNet.Infraestrutura.Contexto;

namespace SwNet.Testes;

/// <summary>
/// Sobe a API de verdade (pipeline completo: JWT, policies, rate limit) trocando
/// o SQL Server por um SQLite em memória isolado por instância.
/// </summary>
public sealed class ApiFactory : WebApplicationFactory<Program>
{
    private readonly SqliteConnection _conexao = new("DataSource=:memory:");

    public ApiFactory() => _conexao.Open();

    protected override void ConfigureWebHost(IWebHostBuilder builder)
    {
        builder.UseEnvironment("Testing");
        builder.UseSetting("Jwt:ChaveSecreta", "chave-de-teste-com-pelo-menos-32-caracteres!!");
        builder.UseSetting("Jwt:Emissor", "SwNet.Api");
        builder.UseSetting("Jwt:Audiencia", "SwnetWeb");
        builder.UseSetting("Autenticacao:LimiteRequisicoesPorMinuto", "1000");

        builder.ConfigureServices(services =>
        {
            services.RemoveAll<DbContextOptions<SwNetDbContext>>();
            services.RemoveAll<IDbContextOptionsConfiguration<SwNetDbContext>>();
            services.AddDbContext<SwNetDbContext>(opt => opt.UseSqlite(_conexao));

            using var escopo = services.BuildServiceProvider().CreateScope();
            var contexto = escopo.ServiceProvider.GetRequiredService<SwNetDbContext>();
            contexto.Database.EnsureCreated();
            SwNetDbSeeder.SemearAsync(contexto).GetAwaiter().GetResult();
        });
    }

    public async Task ExecutarNoBancoAsync(Func<SwNetDbContext, Task> acao)
    {
        using var escopo = Services.CreateScope();
        await acao(escopo.ServiceProvider.GetRequiredService<SwNetDbContext>());
    }

    public async Task<LoginRetornoDTO> LogarAsync(string login = SwNetDbSeeder.LoginAdmin, string senha = SwNetDbSeeder.SenhaAdmin)
    {
        var resposta = await CreateClient().PostAsJsonAsync("/api/auth/login", new { login, senha });
        resposta.EnsureSuccessStatusCode();
        return (await resposta.Content.ReadFromJsonAsync<LoginRetornoDTO>())!;
    }

    public HttpClient ClienteAutenticado(string token)
    {
        var cliente = CreateClient();
        cliente.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
        return cliente;
    }

    protected override void Dispose(bool disposing)
    {
        base.Dispose(disposing);
        if (disposing) _conexao.Dispose();
    }
}
