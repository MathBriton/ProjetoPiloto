using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SwNet.Aplicacao.Autorizacao;
using SwNet.Aplicacao.DTOs.Usuario;
using SwNet.Aplicacao.Interfaces;

namespace SwNet.Api.Controllers;

/// <summary>
/// Backend da tela "Acessos Usuários" da sidebar: CRUD de usuário, perfil
/// e vínculo com plantas. Protegido por RBAC via [RequerPermissao].
/// </summary>
[ApiController]
[Route("api/usuarios")]
[Authorize]
public class UsuariosController : ControllerBase
{
    private readonly IUsuarioRepositorio _usuarioRepositorio;

    public UsuariosController(IUsuarioRepositorio usuarioRepositorio) =>
        _usuarioRepositorio = usuarioRepositorio;

    [HttpGet]
    [RequerPermissao("acessos-usuarios.visualizar")]
    public async Task<ActionResult<UsuarioObterDTO>> Listar()
    {
        var usuarios = await _usuarioRepositorio.ListarAsync();
        return Ok(new UsuarioObterDTO
        {
            Usuarios = usuarios.Select(u => new UsuarioDTO
            {
                Id = u.Id,
                Login = u.Login,
                NomeCompleto = u.NomeCompleto,
                Email = u.Email,
                Ativo = u.Ativo,
                Perfil = u.Perfil?.Nome ?? string.Empty,
                UltimoLoginEm = u.UltimoLoginEm
            }).ToList()
        });
    }

    // TODO: POST/PUT com [RequerPermissao("acessos-usuarios.editar")] — inserir/alterar
    // usuário, perfil e vínculo de plantas (UsuarioInserirDTO / UsuarioAlterarDTO já prontos)
}
