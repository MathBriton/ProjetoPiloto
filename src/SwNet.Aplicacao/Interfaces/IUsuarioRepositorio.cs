using SwNet.Dominio.Entidades;

namespace SwNet.Aplicacao.Interfaces;

public interface IUsuarioRepositorio
{
    Task<Usuario?> ObterPorLoginAsync(string login);
    Task<Usuario?> ObterPorIdComPermissoesAsync(int id);
    Task<List<Usuario>> ListarAsync();
    Task<int> InserirAsync(Usuario usuario);
    Task AtualizarAsync(Usuario usuario);
    Task RegistrarTentativaFalhaAsync(int usuarioId);
    Task RegistrarLoginSucedidoAsync(int usuarioId);
}

public interface ITokenServico
{
    (string token, DateTime expiraEm) GerarToken(Usuario usuario, IEnumerable<string> permissoes, int plantaAtualId);
    string GerarRefreshToken();
}
