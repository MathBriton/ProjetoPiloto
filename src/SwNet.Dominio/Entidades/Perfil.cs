namespace SwNet.Dominio.Entidades;

public class Perfil
{
    public int Id { get; set; }
    public string Nome { get; set; } = string.Empty;
    public string? Descricao { get; set; }

    public ICollection<PerfilPermissao> PerfilPermissoes { get; set; } = new List<PerfilPermissao>();
    public ICollection<Usuario> Usuarios { get; set; } = new List<Usuario>();
}
