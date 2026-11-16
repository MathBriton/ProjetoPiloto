using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SwNet.Aplicacao.Autorizacao;
using SwNet.Aplicacao.DTOs.Autenticacao;
using SwNet.Infraestrutura.Contexto;
using Microsoft.EntityFrameworkCore;

namespace SwNet.Api.Controllers;

[ApiController]
[Route("api/plantas")]
[Authorize]
public class PlantasController : ControllerBase
{
    private readonly SwNetDbContext _contexto;

    public PlantasController(SwNetDbContext contexto) => _contexto = contexto;

    /// <summary>Plantas do usuário logado, para o seletor "trocar de planta" na topbar.</summary>
    [HttpGet("minhas")]
    public async Task<ActionResult<List<PlantaDTO>>> MinhasPlantas()
    {
        var usuarioId = User.ObterUsuarioId();

        var plantas = await _contexto.UsuarioPlantas
            .Where(up => up.UsuarioId == usuarioId && up.Planta!.Ativa)
            .Select(up => new PlantaDTO { Id = up.Planta!.Id, Codigo = up.Planta.Codigo, Nome = up.Planta.Nome })
            .ToListAsync();

        return Ok(plantas);
    }
}
