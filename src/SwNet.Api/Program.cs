using System.Text;
using System.Threading.RateLimiting;
using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.AspNetCore.Authorization;
using Microsoft.EntityFrameworkCore;
using Microsoft.IdentityModel.Tokens;
using Microsoft.OpenApi;
using SwNet.Aplicacao.Autorizacao;
using SwNet.Aplicacao.Configuracao;
using SwNet.Aplicacao.Interfaces;
using SwNet.Aplicacao.Servicos;
using SwNet.Infraestrutura.Contexto;
using SwNet.Infraestrutura.Repositorios;
using SwNet.Infraestrutura.Servicos;

var builder = WebApplication.CreateBuilder(args);

// Banco de dados (SQL Server)
builder.Services.AddDbContext<SwNetDbContext>(opt =>
    opt.UseSqlServer(builder.Configuration.GetConnectionString("DefaultConnection")));

// Opções validadas no startup: a API nem sobe com chave JWT fraca/ausente
builder.Services.AddOptions<JwtOpcoes>()
    .Bind(builder.Configuration.GetSection(JwtOpcoes.Secao))
    .ValidateDataAnnotations()
    .ValidateOnStart();
builder.Services.AddOptions<AutenticacaoOpcoes>()
    .Bind(builder.Configuration.GetSection(AutenticacaoOpcoes.Secao))
    .ValidateDataAnnotations()
    .ValidateOnStart();

// Aplicação / Infraestrutura
builder.Services.AddSingleton(TimeProvider.System);
builder.Services.AddScoped<IUsuarioRepositorio, UsuarioRepositorio>();
builder.Services.AddScoped<IRefreshTokenRepositorio, RefreshTokenRepositorio>();
builder.Services.AddSingleton<ITokenServico, TokenServico>();
builder.Services.AddScoped<IEstrategiaAutenticacao, SistemaEstrategiaAutenticacao>();
builder.Services.AddScoped<AutenticacaoServico>();

// Autenticação JWT
var jwt = builder.Configuration.GetSection(JwtOpcoes.Secao).Get<JwtOpcoes>() ?? new JwtOpcoes();
builder.Services.AddAuthentication(JwtBearerDefaults.AuthenticationScheme)
    .AddJwtBearer(opt =>
    {
        // Mantém os nomes curtos das claims ("sub", "permissao", ...) em vez de mapear para URIs
        opt.MapInboundClaims = false;
        opt.TokenValidationParameters = new TokenValidationParameters
        {
            ValidateIssuer = true,
            ValidateAudience = true,
            ValidateLifetime = true,
            ValidateIssuerSigningKey = true,
            ValidIssuer = jwt.Emissor,
            ValidAudience = jwt.Audiencia,
            IssuerSigningKey = new SymmetricSecurityKey(Encoding.UTF8.GetBytes(jwt.ChaveSecreta)),
            ValidAlgorithms = [SecurityAlgorithms.HmacSha256],
            NameClaimType = ClaimsSwNet.Login,
            ClockSkew = TimeSpan.FromSeconds(30)
        };
        opt.Events = new JwtBearerEvents
        {
            // Avisa o frontend que o token venceu (para disparar o refresh)
            OnChallenge = ctx =>
            {
                if (ctx.AuthenticateFailure is SecurityTokenExpiredException)
                    ctx.Response.Headers.Append("X-Token-Expirado", "true");
                return Task.CompletedTask;
            }
        };
    });

// Autorização:
//  - [RequerPermissao("chave")] → policy "permissao:chave" criada sob demanda pelo PermissaoPolicyProvider
//  - FallbackPolicy: todo endpoint exige login, exceto os marcados com [AllowAnonymous]
builder.Services.AddSingleton<IAuthorizationPolicyProvider, PermissaoPolicyProvider>();
builder.Services.AddSingleton<IAuthorizationHandler, PermissaoAuthorizationHandler>();
builder.Services.AddAuthorization(opt =>
{
    opt.FallbackPolicy = new AuthorizationPolicyBuilder().RequireAuthenticatedUser().Build();
});

// Rate limit no login/refresh: 10 requisições por minuto por IP (mitiga força bruta)
builder.Services.AddRateLimiter(opt =>
{
    opt.RejectionStatusCode = StatusCodes.Status429TooManyRequests;
    opt.AddPolicy("autenticacao", ctx =>
        RateLimitPartition.GetFixedWindowLimiter(
            ctx.Connection.RemoteIpAddress?.ToString() ?? "desconhecido",
            _ => new FixedWindowRateLimiterOptions
            {
                PermitLimit = builder.Configuration.GetValue("Autenticacao:LimiteRequisicoesPorMinuto", 10),
                Window = TimeSpan.FromMinutes(1),
                QueueLimit = 0
            }));
});

builder.Services.AddCors(opt =>
{
    opt.AddPolicy("SwnetWeb", policy =>
    {
        var origens = builder.Configuration.GetSection("Cors:OrigensPermitidas").Get<string[]>() ?? [];
        policy.WithOrigins(origens).AllowAnyHeader().AllowAnyMethod().WithExposedHeaders("X-Token-Expirado");
    });
});

builder.Services.AddProblemDetails();
builder.Services.AddControllers();
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen(opt =>
{
    opt.SwaggerDoc("v1", new OpenApiInfo { Title = "SwNet API", Version = "v1" });
    opt.AddSecurityDefinition("Bearer", new OpenApiSecurityScheme
    {
        Type = SecuritySchemeType.Http,
        Scheme = "bearer",
        BearerFormat = "JWT",
        Description = "Cole aqui o token retornado por POST /api/auth/login"
    });
    opt.AddSecurityRequirement(documento => new OpenApiSecurityRequirement
    {
        [new OpenApiSecuritySchemeReference("Bearer", documento)] = []
    });
});

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
app.UseRateLimiter();
app.MapControllers();

app.Run();

/// <summary>Exposto para os testes de integração (WebApplicationFactory).</summary>
public partial class Program;
