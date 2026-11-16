using SwNet.Dominio.Entidades;

namespace SwNet.Aplicacao.Interfaces;

public interface IUsuarioRepositorio
{
    /// <summary>Usuário com perfil, permissões e plantas carregados (entidade rastreada).</summary>
    Task<Usuario?> ObterPorLoginAsync(string login);

    /// <summary>Usuário com perfil, permissões e plantas carregados (entidade rastreada).</summary>
    Task<Usuario?> ObterPorIdComPermissoesAsync(int id);

    Task<List<Usuario>> ListarAsync();
    Task<int> InserirAsync(Usuario usuario);
    Task AtualizarAsync(Usuario usuario);

    /// <summary>Persiste alterações feitas em entidades rastreadas (unidade de trabalho do request).</summary>
    Task SalvarAsync();
}

public interface IRefreshTokenRepositorio
{
    Task<RefreshToken?> ObterPorHashAsync(string tokenHash);
    void Adicionar(RefreshToken refreshToken);
    Task RevogarFamiliaAsync(Guid familia, DateTime agora, string? ip, string motivo);
    Task RevogarTodosDoUsuarioAsync(int usuarioId, DateTime agora, string? ip, string motivo);
}

public interface ITokenServico
{
    (string token, DateTime expiraEm) GerarToken(Usuario usuario, IEnumerable<string> permissoes, int plantaAtualId);

    /// <summary>Token aleatório (512 bits) para devolver ao cliente.</summary>
    string GerarRefreshToken();

    /// <summary>Hash SHA-256 (hex) do refresh token — é isso que vai para o banco.</summary>
    string CalcularHash(string refreshToken);
}
