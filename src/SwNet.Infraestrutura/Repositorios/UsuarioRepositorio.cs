using Microsoft.EntityFrameworkCore;
using SwNet.Aplicacao.Interfaces;
using SwNet.Dominio.Entidades;
using SwNet.Infraestrutura.Contexto;

namespace SwNet.Infraestrutura.Repositorios;

public class UsuarioRepositorio : IUsuarioRepositorio
{
    private const int MaxTentativas = 5;
    private static readonly TimeSpan TempoBloqueio = TimeSpan.FromMinutes(15);

    private readonly SwNetDbContext _contexto;

    public UsuarioRepositorio(SwNetDbContext contexto) => _contexto = contexto;

    public Task<Usuario?> ObterPorLoginAsync(string login) =>
        _contexto.Usuarios
            .Include(u => u.Perfil).ThenInclude(p => p!.PerfilPermissoes).ThenInclude(pp => pp.Permissao)
            .Include(u => u.UsuarioPermissoes).ThenInclude(up => up.Permissao)
            .Include(u => u.UsuarioPlantas).ThenInclude(up => up.Planta)
            .FirstOrDefaultAsync(u => u.Login == login);

    public Task<Usuario?> ObterPorIdComPermissoesAsync(int id) =>
        _contexto.Usuarios
            .Include(u => u.Perfil)
            .Include(u => u.UsuarioPlantas).ThenInclude(up => up.Planta)
            .FirstOrDefaultAsync(u => u.Id == id);

    public Task<List<Usuario>> ListarAsync() =>
        _contexto.Usuarios.Include(u => u.Perfil).OrderBy(u => u.NomeCompleto).ToListAsync();

    public async Task<int> InserirAsync(Usuario usuario)
    {
        _contexto.Usuarios.Add(usuario);
        await _contexto.SaveChangesAsync();
        return usuario.Id;
    }

    public Task AtualizarAsync(Usuario usuario)
    {
        _contexto.Usuarios.Update(usuario);
        return _contexto.SaveChangesAsync();
    }

    public async Task RegistrarTentativaFalhaAsync(int usuarioId)
    {
        var usuario = await _contexto.Usuarios.FindAsync(usuarioId);
        if (usuario is null) return;

        usuario.TentativasFalhas++;
        if (usuario.TentativasFalhas >= MaxTentativas)
            usuario.BloqueadoAte = DateTime.UtcNow.Add(TempoBloqueio);

        await _contexto.SaveChangesAsync();
    }

    public async Task RegistrarLoginSucedidoAsync(int usuarioId)
    {
        var usuario = await _contexto.Usuarios.FindAsync(usuarioId);
        if (usuario is null) return;

        usuario.TentativasFalhas = 0;
        usuario.BloqueadoAte = null;
        usuario.UltimoLoginEm = DateTime.UtcNow;

        await _contexto.SaveChangesAsync();
    }
}
