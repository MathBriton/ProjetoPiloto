using Microsoft.AspNetCore.Mvc;
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

    /// <summary>
    /// Login com usuário/senha do sistema (BCrypt). Ver seed do banco para
    /// o usuário fictício de teste (usuário: admin / senha: Admin@123).
    /// </summary>
    [HttpPost("login")]
    public async Task<ActionResult<LoginRetornoDTO>> Login([FromBody] LoginDTO dto)
    {
        var resultado = await _autenticacaoServico.AutenticarAsync(dto);
        return resultado.Sucesso ? Ok(resultado) : Unauthorized(resultado);
    }

    // TODO: POST api/auth/refresh — rotação real de refresh token (hoje stub 501 no scaffold do frontend)
    [HttpPost("refresh")]
    public IActionResult Refresh() => StatusCode(501, new { mensagem = "Rotação de refresh token ainda não implementada." });
}
