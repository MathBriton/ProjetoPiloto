namespace SwNet.Dominio.Entidades;

/// <summary>
/// Fábrica/unidade do grupo (ex.: Juiz de Fora, ...). Usada pelo seletor
/// de planta na topbar para trocar o contexto de dados do usuário.
/// </summary>
public class Planta
{
    public int Id { get; set; }
    public string Codigo { get; set; } = string.Empty;
    public string Nome { get; set; } = string.Empty;
    public bool Ativa { get; set; } = true;

    public ICollection<UsuarioPlanta> UsuarioPlantas { get; set; } = new List<UsuarioPlanta>();
}

/// <summary>
/// Plantas às quais o usuário tem acesso, e qual é a planta padrão dele.
/// </summary>
public class UsuarioPlanta
{
    public int UsuarioId { get; set; }
    public Usuario? Usuario { get; set; }
    public int PlantaId { get; set; }
    public Planta? Planta { get; set; }
    public bool Padrao { get; set; }
}
