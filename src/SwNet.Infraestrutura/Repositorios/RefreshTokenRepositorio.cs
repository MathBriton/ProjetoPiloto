using Microsoft.EntityFrameworkCore;
using SwNet.Aplicacao.Interfaces;
using SwNet.Dominio.Entidades;
using SwNet.Infraestrutura.Contexto;

namespace SwNet.Infraestrutura.Repositorios;

public class RefreshTokenRepositorio : IRefreshTokenRepositorio
{
    private readonly SwNetDbContext _contexto;

    public RefreshTokenRepositorio(SwNetDbContext contexto) => _contexto = contexto;

    public Task<RefreshToken?> ObterPorHashAsync(string tokenHash) =>
        _contexto.RefreshTokens.FirstOrDefaultAsync(rt => rt.TokenHash == tokenHash);

    public void Adicionar(RefreshToken refreshToken) => _contexto.RefreshTokens.Add(refreshToken);

    public async Task RevogarFamiliaAsync(Guid familia, DateTime agora, string? ip, string motivo)
    {
        var ativos = await _contexto.RefreshTokens
            .Where(rt => rt.Familia == familia && rt.RevogadoEm == null)
            .ToListAsync();

        foreach (var token in ativos)
            token.Revogar(agora, ip, motivo);
    }

    public async Task RevogarTodosDoUsuarioAsync(int usuarioId, DateTime agora, string? ip, string motivo)
    {
        var ativos = await _contexto.RefreshTokens
            .Where(rt => rt.UsuarioId == usuarioId && rt.RevogadoEm == null)
            .ToListAsync();

        foreach (var token in ativos)
            token.Revogar(agora, ip, motivo);
    }
}
