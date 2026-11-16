using Microsoft.AspNetCore.Authorization;

namespace SwNet.Aplicacao.Autorizacao;

/// <summary>
/// Uso: [RequerPermissao(Permissoes.AcessosUsuariosEditar)] em cima do endpoint/controller.
/// Gera a policy "permissao:&lt;chave&gt;", resolvida dinamicamente pelo
/// <see cref="PermissaoPolicyProvider"/> — não é preciso registrar cada policy no Program.cs.
/// </summary>
[AttributeUsage(AttributeTargets.Class | AttributeTargets.Method, AllowMultiple = true, Inherited = true)]
public class RequerPermissaoAttribute : AuthorizeAttribute
{
    public const string PrefixoPolicy = "permissao:";

    public RequerPermissaoAttribute(string chave)
    {
        if (string.IsNullOrWhiteSpace(chave))
            throw new ArgumentException("Informe a chave da permissão.", nameof(chave));

        Chave = chave;
        Policy = PrefixoPolicy + chave;
    }

    public string Chave { get; }
}
