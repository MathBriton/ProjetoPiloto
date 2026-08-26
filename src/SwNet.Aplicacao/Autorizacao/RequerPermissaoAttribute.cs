using Microsoft.AspNetCore.Authorization;

namespace SwNet.Aplicacao.Autorizacao;

/// <summary>
/// Uso: [RequerPermissao("acessos-usuarios.editar")] em cima do endpoint/controller.
/// A permissão é resolvida contra as claims "permissao" emitidas no JWT
/// (perfil + overrides de UsuarioPermissao), nunca contra o token em si.
/// </summary>
public class RequerPermissaoAttribute : AuthorizeAttribute
{
    public const string ClaimPermissao = "permissao";

    public RequerPermissaoAttribute(string chave) : base()
    {
        Policy = $"{ClaimPermissao}:{chave}";
    }
}
