using Microsoft.AspNetCore.Authorization;
using Microsoft.Extensions.Options;

namespace SwNet.Aplicacao.Autorizacao;

/// <summary>Requisito: o usuário autenticado precisa ter a claim "permissao" com a chave informada.</summary>
public sealed class PermissaoRequirement : IAuthorizationRequirement
{
    public PermissaoRequirement(string chave) => Chave = chave;
    public string Chave { get; }
}

public sealed class PermissaoAuthorizationHandler : AuthorizationHandler<PermissaoRequirement>
{
    protected override Task HandleRequirementAsync(AuthorizationHandlerContext context, PermissaoRequirement requirement)
    {
        if (context.User.Identity?.IsAuthenticated == true &&
            context.User.HasClaim(ClaimsSwNet.Permissao, requirement.Chave))
        {
            context.Succeed(requirement);
        }

        return Task.CompletedTask;
    }
}

/// <summary>
/// Cria sob demanda as policies "permissao:&lt;chave&gt;" usadas pelo [RequerPermissao].
/// Qualquer outra policy cai no provider padrão (policies nomeadas, default, fallback).
/// </summary>
public sealed class PermissaoPolicyProvider : IAuthorizationPolicyProvider
{
    private readonly DefaultAuthorizationPolicyProvider _padrao;

    public PermissaoPolicyProvider(IOptions<AuthorizationOptions> opcoes) =>
        _padrao = new DefaultAuthorizationPolicyProvider(opcoes);

    public Task<AuthorizationPolicy?> GetPolicyAsync(string policyName)
    {
        if (policyName.StartsWith(RequerPermissaoAttribute.PrefixoPolicy, StringComparison.Ordinal))
        {
            var chave = policyName[RequerPermissaoAttribute.PrefixoPolicy.Length..];
            var policy = new AuthorizationPolicyBuilder()
                .RequireAuthenticatedUser()
                .AddRequirements(new PermissaoRequirement(chave))
                .Build();
            return Task.FromResult<AuthorizationPolicy?>(policy);
        }

        return _padrao.GetPolicyAsync(policyName);
    }

    public Task<AuthorizationPolicy> GetDefaultPolicyAsync() => _padrao.GetDefaultPolicyAsync();
    public Task<AuthorizationPolicy?> GetFallbackPolicyAsync() => _padrao.GetFallbackPolicyAsync();
}
