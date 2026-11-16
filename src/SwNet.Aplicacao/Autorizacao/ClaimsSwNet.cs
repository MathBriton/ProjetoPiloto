using System.Security.Claims;

namespace SwNet.Aplicacao.Autorizacao;

/// <summary>Nomes das claims emitidas no JWT (sem o mapeamento legado do .NET para URIs do WS-*).</summary>
public static class ClaimsSwNet
{
    public const string UsuarioId = "sub";
    public const string Login = "unique_name";
    public const string Nome = "nome";
    public const string PlantaAtual = "plantaAtual";
    public const string Permissao = "permissao";

    public static int ObterUsuarioId(this ClaimsPrincipal usuario) =>
        int.TryParse(usuario.FindFirstValue(UsuarioId), out var id)
            ? id
            : throw new InvalidOperationException("Token sem a claim 'sub'.");

    public static int ObterPlantaAtualId(this ClaimsPrincipal usuario) =>
        int.TryParse(usuario.FindFirstValue(PlantaAtual), out var id) ? id : 0;

    public static bool TemPermissao(this ClaimsPrincipal usuario, string chave) =>
        usuario.HasClaim(Permissao, chave);
}
