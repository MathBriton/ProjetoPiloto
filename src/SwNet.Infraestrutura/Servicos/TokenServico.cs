using System.IdentityModel.Tokens.Jwt;
using System.Security.Claims;
using System.Security.Cryptography;
using System.Text;
using Microsoft.Extensions.Options;
using Microsoft.IdentityModel.Tokens;
using SwNet.Aplicacao.Autorizacao;
using SwNet.Aplicacao.Configuracao;
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
    private readonly JwtOpcoes _opcoes;
    private readonly TimeProvider _relogio;
    private readonly SigningCredentials _credenciais;

    public TokenServico(IOptions<JwtOpcoes> opcoes, TimeProvider relogio)
    {
        _opcoes = opcoes.Value;
        _relogio = relogio;
        _credenciais = new SigningCredentials(
            new SymmetricSecurityKey(Encoding.UTF8.GetBytes(_opcoes.ChaveSecreta)),
            SecurityAlgorithms.HmacSha256);
    }

    public (string token, DateTime expiraEm) GerarToken(Usuario usuario, IEnumerable<string> permissoes, int plantaAtualId)
    {
        var agora = _relogio.GetUtcNow().UtcDateTime;
        var expiraEm = agora.AddMinutes(_opcoes.MinutosExpiracao);

        var claims = new List<Claim>
        {
            new(ClaimsSwNet.UsuarioId, usuario.Id.ToString()),
            new(ClaimsSwNet.Login, usuario.Login),
            new(ClaimsSwNet.Nome, usuario.NomeCompleto),
            new(ClaimsSwNet.PlantaAtual, plantaAtualId.ToString()),
            new(JwtRegisteredClaimNames.Jti, Guid.NewGuid().ToString("N"))
        };
        claims.AddRange(permissoes.Select(p => new Claim(ClaimsSwNet.Permissao, p)));

        var token = new JwtSecurityToken(
            issuer: _opcoes.Emissor,
            audience: _opcoes.Audiencia,
            claims: claims,
            notBefore: agora,
            expires: expiraEm,
            signingCredentials: _credenciais);

        return (new JwtSecurityTokenHandler().WriteToken(token), expiraEm);
    }

    public string GerarRefreshToken() =>
        Base64UrlEncoder.Encode(RandomNumberGenerator.GetBytes(64));

    public string CalcularHash(string refreshToken) =>
        Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(refreshToken)));
}
