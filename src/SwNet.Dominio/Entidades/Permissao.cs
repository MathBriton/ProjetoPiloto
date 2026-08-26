namespace SwNet.Dominio.Entidades;

/// <summary>
/// Catálogo de permissões (chave usada em [RequerPermissao("chave")]).
/// Ex: "acessos-usuarios.visualizar", "acessos-usuarios.editar".
/// </summary>
public class Permissao
{
    public int Id { get; set; }
    public string Chave { get; set; } = string.Empty;
    public string Descricao { get; set; } = string.Empty;
    public string? Modulo { get; set; }

    public ICollection<PerfilPermissao> PerfilPermissoes { get; set; } = new List<PerfilPermissao>();
    public ICollection<UsuarioPermissao> UsuarioPermissoes { get; set; } = new List<UsuarioPermissao>();
}

public class PerfilPermissao
{
    public int PerfilId { get; set; }
    public Perfil? Perfil { get; set; }
    public int PermissaoId { get; set; }
    public Permissao? Permissao { get; set; }
}

/// <summary>
/// Permissão concedida/negada diretamente ao usuário, sobrepondo o perfil
/// (equivalente às flags especiais de CRM_DTI do legado).
/// </summary>
public class UsuarioPermissao
{
    public int UsuarioId { get; set; }
    public Usuario? Usuario { get; set; }
    public int PermissaoId { get; set; }
    public Permissao? Permissao { get; set; }
    public bool Concedida { get; set; }
}
