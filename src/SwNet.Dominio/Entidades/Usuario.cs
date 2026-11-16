namespace SwNet.Dominio.Entidades;

/// <summary>
/// Usuário do sistema. Autenticação local (BCrypt) neste scaffold;
/// estratégia AD/Windows entra depois via IEstrategiaAutenticacao.
/// </summary>
public class Usuario
{
    public int Id { get; set; }
    public string Login { get; set; } = string.Empty;
    public string NomeCompleto { get; set; } = string.Empty;
    public string Email { get; set; } = string.Empty;
    public string SenhaHash { get; set; } = string.Empty;
    public bool Ativo { get; set; } = true;
    public bool SenhaProvisoria { get; set; }
    public int TentativasFalhas { get; set; }
    public DateTime? BloqueadoAte { get; set; }
    public DateTime? UltimoLoginEm { get; set; }
    public DateTime CriadoEm { get; set; } = DateTime.UtcNow;

    public int PerfilId { get; set; }
    public Perfil? Perfil { get; set; }

    public ICollection<UsuarioPlanta> UsuarioPlantas { get; set; } = new List<UsuarioPlanta>();
    public ICollection<UsuarioPermissao> UsuarioPermissoes { get; set; } = new List<UsuarioPermissao>();
    public ICollection<RefreshToken> RefreshTokens { get; set; } = new List<RefreshToken>();

    public bool EstaBloqueado(DateTime agora) => BloqueadoAte is not null && BloqueadoAte > agora;

    /// <summary>Conta uma falha de login; ao atingir o limite, bloqueia. Retorna true se bloqueou agora.</summary>
    public bool RegistrarFalhaLogin(DateTime agora, int maxTentativas, TimeSpan tempoBloqueio)
    {
        // Bloqueio anterior já venceu: recomeça a contagem
        if (BloqueadoAte is not null && BloqueadoAte <= agora)
        {
            BloqueadoAte = null;
            TentativasFalhas = 0;
        }

        TentativasFalhas++;
        if (TentativasFalhas < maxTentativas) return false;

        BloqueadoAte = agora.Add(tempoBloqueio);
        TentativasFalhas = 0;
        return true;
    }

    public void RegistrarLoginSucedido(DateTime agora)
    {
        TentativasFalhas = 0;
        BloqueadoAte = null;
        UltimoLoginEm = agora;
    }
}
