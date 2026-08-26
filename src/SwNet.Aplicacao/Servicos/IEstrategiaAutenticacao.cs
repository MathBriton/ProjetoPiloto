using SwNet.Dominio.Entidades;

namespace SwNet.Aplicacao.Servicos;

/// <summary>
/// Estratégia de autenticação (Sistema/BCrypt, AD/LDAP, ...). Cada método do
/// legado (Autenticacao.asp) vira uma implementação; o AutenticacaoServico
/// só orquestra bloqueio/tentativas/token, sem duplicar a lógica de cada uma.
/// </summary>
public interface IEstrategiaAutenticacao
{
    string Metodo { get; }
    Task<bool> ValidarCredenciaisAsync(Usuario usuario, string senha);
}

/// <summary>
/// Autenticação local com hash BCrypt. Implementação padrão deste scaffold.
/// </summary>
public class SistemaEstrategiaAutenticacao : IEstrategiaAutenticacao
{
    public string Metodo => "SISTEMA";

    public Task<bool> ValidarCredenciaisAsync(Usuario usuario, string senha)
    {
        var valido = BCrypt.Net.BCrypt.Verify(senha, usuario.SenhaHash);
        return Task.FromResult(valido);
    }
}
