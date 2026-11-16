namespace SwNet.Dominio.Entidades;

/// <summary>
/// Refresh token persistido para rotação. Guardamos apenas o hash SHA-256 do
/// token (nunca o valor em claro). Cada login abre uma "família"; a cada
/// renovação o token atual é revogado e substituído por um novo da mesma
/// família. Se um token já revogado for reapresentado (reuso = possível
/// roubo), a família inteira é revogada e o usuário precisa logar de novo.
/// </summary>
public class RefreshToken
{
    public long Id { get; set; }
    public int UsuarioId { get; set; }
    public Usuario? Usuario { get; set; }

    public string TokenHash { get; set; } = string.Empty;
    public Guid Familia { get; set; }

    /// <summary>Planta selecionada na sessão, preservada ao renovar o token.</summary>
    public int PlantaId { get; set; }

    public DateTime CriadoEm { get; set; } = DateTime.UtcNow;
    public string? CriadoPorIp { get; set; }
    public DateTime ExpiraEm { get; set; }

    public DateTime? RevogadoEm { get; set; }
    public string? RevogadoPorIp { get; set; }
    public string? MotivoRevogacao { get; set; }
    public string? SubstituidoPorHash { get; set; }

    public bool Expirado(DateTime agora) => agora >= ExpiraEm;
    public bool Revogado => RevogadoEm is not null;
    public bool Ativo(DateTime agora) => !Revogado && !Expirado(agora);

    public void Revogar(DateTime agora, string? ip, string motivo, string? substituidoPorHash = null)
    {
        RevogadoEm = agora;
        RevogadoPorIp = ip;
        MotivoRevogacao = motivo;
        SubstituidoPorHash = substituidoPorHash;
    }
}
