using System.IdentityModel.Tokens.Jwt;
using System.Security.Claims;
using System.Security.Cryptography;
using Microsoft.Extensions.Configuration;
using Microsoft.IdentityModel.Tokens;
using SwNet.Aplicacao.Autorizacao;
using SwNet.Aplicacao.Interfaces;
using SwNet.Dominio.Entidades;

namespace SwNet.Infraestrutura.Servicos;

/// <summary>
/// Emite JWT com claims mínimas de identidade + uma claim "permissao" por
/// permissão efetiva do usuário. Nada de sessão de ~50 campos como no legado:
/// o resto é resolvido no backend a cada requisição via [RequerPermissao].
/// </summary>
public class TokenServico : ITokenServico
{
    private readonly IConfiguration _configuracao;

    public TokenServico(IConfiguration configuracao) => _configuracao = configuracao;

    public (string token, DateTime expiraEm) GerarToken(Usuario usuario, IEnumerable<string> permissoes, int plantaAtualId)
    {
        var chave = _configuracao["Jwt:ChaveSecreta"]!;
        var emissor = _configuracao["Jwt:Emissor"];
        var audiencia = _configuracao["Jwt:Audiencia"];
        var minutosExpiracao = int.Parse(_configuracao["Jwt:MinutosExpiracao"] ?? "60");

        var claims = new List<Claim>
        {
            new(JwtRegisteredClaimNames.Sub, usuario.Id.ToString()),
            new(JwtRegisteredClaimNames.UniqueName, usuario.Login),
            new("nome", usuario.NomeCompleto),
            new("plantaAtual", plantaAtualId.ToString())
        };
        claims.AddRange(permissoes.Select(p => new Claim(RequerPermissaoAttribute.ClaimPermissao, p)));

        var expiraEm = DateTime.UtcNow.AddMinutes(minutosExpiracao);
        var credenciais = new SigningCredentials(
            new SymmetricSecurityKey(System.Text.Encoding.UTF8.GetBytes(chave)),
            SecurityAlgorithms.HmacSha256);

        var token = new JwtSecurityToken(emissor, audiencia, claims, expires: expiraEm, signingCredentials: credenciais);
        return (new JwtSecurityTokenHandler().WriteToken(token), expiraEm);
    }

    public string GerarRefreshToken() => Convert.ToBase64String(RandomNumberGenerator.GetBytes(64));
}
