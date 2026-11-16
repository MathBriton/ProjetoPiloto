using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.RateLimiting;
using SwNet.Aplicacao.Autorizacao;
using SwNet.Aplicacao.DTOs.Autenticacao;
using SwNet.Aplicacao.Servicos;

namespace SwNet.Api.Controllers;

[ApiController]
[Route("api/auth")]
public class AuthController : ControllerBase
{
    private readonly AutenticacaoServico _autenticacaoServico;

    public AuthController(AutenticacaoServico autenticacaoServico) =>
        _autenticacaoServico = autenticacaoServico;

    private string? Ip => HttpContext.Connection.RemoteIpAddress?.ToString();

    /// <summary>
    /// Login com usuário/senha do sistema (BCrypt). Ver seed do banco para
    /// o usuário fictício de teste (usuário: admin / senha: Admin@123).
    /// </summary>
    [HttpPost("login")]
    [AllowAnonymous]
    [EnableRateLimiting("autenticacao")]
    public async Task<ActionResult<LoginRetornoDTO>> Login([FromBody] LoginDTO dto)
    {
        var resultado = await _autenticacaoServico.AutenticarAsync(dto, Ip);
        return resultado.Sucesso ? Ok(resultado) : Unauthorized(resultado);
    }

    /// <summary>Troca o refresh token por um novo par access/refresh (rotação). O token enviado deixa de valer.</summary>
    [HttpPost("refresh")]
    [AllowAnonymous]
    [EnableRateLimiting("autenticacao")]
    public async Task<ActionResult<LoginRetornoDTO>> Refresh([FromBody] RefreshTokenDTO dto)
    {
        var resultado = await _autenticacaoServico.RenovarAsync(dto.RefreshToken, Ip);
        return resultado.Sucesso ? Ok(resultado) : Unauthorized(resultado);
    }

    /// <summary>Logout: invalida o refresh token (e as renovações dele). Funciona mesmo com access token vencido.</summary>
    [HttpPost("logout")]
    [AllowAnonymous]
    public async Task<IActionResult> Logout([FromBody] RefreshTokenDTO dto)
    {
        await _autenticacaoServico.SairAsync(dto.RefreshToken, Ip);
        return NoContent();
    }

    /// <summary>Encerra as sessões do usuário logado em todos os dispositivos.</summary>
    [HttpPost("logout-todas")]
    public async Task<IActionResult> LogoutTodas()
    {
        await _autenticacaoServico.SairDeTodasAsSessoesAsync(User.ObterUsuarioId(), Ip);
        return NoContent();
    }

    /// <summary>Dados atualizados do usuário logado (permissões lidas do banco, não do token).</summary>
    [HttpGet("me")]
    public async Task<ActionResult<UsuarioLogadoDTO>> Me()
    {
        var usuario = await _autenticacaoServico.ObterUsuarioLogadoAsync(User.ObterUsuarioId(), User.ObterPlantaAtualId());
        return usuario is null ? Unauthorized() : Ok(usuario);
    }

    /// <summary>Troca a planta ativa: reemite access + refresh token com a nova claim "plantaAtual".</summary>
    [HttpPost("trocar-planta")]
    public async Task<ActionResult<LoginRetornoDTO>> TrocarPlanta([FromBody] TrocarPlantaDTO dto)
    {
        var resultado = await _autenticacaoServico.TrocarPlantaAsync(User.ObterUsuarioId(), dto, Ip);
        return resultado.Sucesso ? Ok(resultado) : StatusCode(StatusCodes.Status403Forbidden, resultado);
    }

    /// <summary>Troca a senha do próprio usuário. Encerra todas as sessões abertas (é preciso logar de novo).</summary>
    [HttpPost("alterar-senha")]
    [EnableRateLimiting("autenticacao")]
    public async Task<ActionResult<BaseErrorDTO>> AlterarSenha([FromBody] AlterarSenhaDTO dto)
    {
        var resultado = await _autenticacaoServico.AlterarSenhaAsync(User.ObterUsuarioId(), dto, Ip);
        return resultado.Sucesso ? Ok(resultado) : BadRequest(resultado);
    }
}
