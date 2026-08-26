using System.Text;
using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.AspNetCore.Authorization;
using Microsoft.EntityFrameworkCore;
using Microsoft.IdentityModel.Tokens;
using SwNet.Aplicacao.Autorizacao;
using SwNet.Aplicacao.Interfaces;
using SwNet.Aplicacao.Servicos;
using SwNet.Infraestrutura.Contexto;
using SwNet.Infraestrutura.Repositorios;
using SwNet.Infraestrutura.Servicos;

var builder = WebApplication.CreateBuilder(args);

// Banco de dados (SQL Server)
builder.Services.AddDbContext<SwNetDbContext>(opt =>
    opt.UseSqlServer(builder.Configuration.GetConnectionString("DefaultConnection")));

// Aplicação / Infraestrutura
builder.Services.AddScoped<IUsuarioRepositorio, UsuarioRepositorio>();
builder.Services.AddScoped<ITokenServico, TokenServico>();
builder.Services.AddScoped<IEstrategiaAutenticacao, SistemaEstrategiaAutenticacao>();
builder.Services.AddScoped<AutenticacaoServico>();

// Autenticação JWT
var jwtChave = builder.Configuration["Jwt:ChaveSecreta"]!;
builder.Services.AddAuthentication(JwtBearerDefaults.AuthenticationScheme)
    .AddJwtBearer(opt =>
    {
        opt.TokenValidationParameters = new TokenValidationParameters
        {
            ValidateIssuer = true,
            ValidateAudience = true,
            ValidateLifetime = true,
            ValidateIssuerSigningKey = true,
            ValidIssuer = builder.Configuration["Jwt:Emissor"],
            ValidAudience = builder.Configuration["Jwt:Audiencia"],
            IssuerSigningKey = new SymmetricSecurityKey(Encoding.UTF8.GetBytes(jwtChave))
        };
    });

// Autorização: cada policy "permissao:<chave>" exige a claim correspondente.
// [RequerPermissao("acessos-usuarios.editar")] cria essa policy automaticamente.
builder.Services.AddAuthorization(opt =>
{
    opt.AddPolicy("permissao:acessos-usuarios.visualizar", p =>
        p.RequireClaim(RequerPermissaoAttribute.ClaimPermissao, "acessos-usuarios.visualizar"));
    opt.AddPolicy("permissao:acessos-usuarios.editar", p =>
        p.RequireClaim(RequerPermissaoAttribute.ClaimPermissao, "acessos-usuarios.editar"));
});

builder.Services.AddCors(opt =>
{
    opt.AddPolicy("SwnetWeb", policy =>
    {
        var origens = builder.Configuration.GetSection("Cors:OrigensPermitidas").Get<string[]>() ?? [];
        policy.WithOrigins(origens).AllowAnyHeader().AllowAnyMethod();
    });
});

builder.Services.AddControllers();
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

var app = builder.Build();

if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();

    // Aplica migrations e cria o usuário fictício (admin / Admin@123) para teste.
    using var escopo = app.Services.CreateScope();
    var contexto = escopo.ServiceProvider.GetRequiredService<SwNetDbContext>();
    await SwNetDbSeeder.ExecutarAsync(contexto);
}

app.UseHttpsRedirection();
app.UseCors("SwnetWeb");
app.UseAuthentication();
app.UseAuthorization();
app.MapControllers();

app.Run();
