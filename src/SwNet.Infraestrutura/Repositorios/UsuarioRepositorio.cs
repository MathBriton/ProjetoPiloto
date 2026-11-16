using Microsoft.EntityFrameworkCore;
using SwNet.Aplicacao.Interfaces;
using SwNet.Dominio.Entidades;
using SwNet.Infraestrutura.Contexto;

namespace SwNet.Infraestrutura.Repositorios;

public class UsuarioRepositorio : IUsuarioRepositorio
{
    private readonly SwNetDbContext _contexto;

    public UsuarioRepositorio(SwNetDbContext contexto) => _contexto = contexto;

    private IQueryable<Usuario> ComPermissoes() =>
        _contexto.Usuarios
            .Include(u => u.Perfil).ThenInclude(p => p!.PerfilPermissoes).ThenInclude(pp => pp.Permissao)
            .Include(u => u.UsuarioPermissoes).ThenInclude(up => up.Permissao)
            .Include(u => u.UsuarioPlantas).ThenInclude(up => up.Planta)
            .AsSplitQuery();

    public Task<Usuario?> ObterPorLoginAsync(string login) =>
        ComPermissoes().FirstOrDefaultAsync(u => u.Login == login);

    public Task<Usuario?> ObterPorIdComPermissoesAsync(int id) =>
        ComPermissoes().FirstOrDefaultAsync(u => u.Id == id);

    public Task<List<Usuario>> ListarAsync() =>
        _contexto.Usuarios.AsNoTracking().Include(u => u.Perfil).OrderBy(u => u.NomeCompleto).ToListAsync();

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

    public Task SalvarAsync() => _contexto.SaveChangesAsync();
}
