using SwNet.Aplicacao.DTOs.Autenticacao;

namespace SwNet.Aplicacao.DTOs.Usuario;

public class UsuarioDTO
{
    public int Id { get; set; }
    public string Login { get; set; } = string.Empty;
    public string NomeCompleto { get; set; } = string.Empty;
    public string Email { get; set; } = string.Empty;
    public bool Ativo { get; set; }
    public string Perfil { get; set; } = string.Empty;
    public DateTime? UltimoLoginEm { get; set; }
}

public class UsuarioObterDTO : BaseErrorDTO
{
    public List<UsuarioDTO> Usuarios { get; set; } = new();
}

public class UsuarioInserirDTO
{
    public string Login { get; set; } = string.Empty;
    public string NomeCompleto { get; set; } = string.Empty;
    public string Email { get; set; } = string.Empty;
    public int PerfilId { get; set; }
    public List<int> PlantaIds { get; set; } = new();
}

public class UsuarioAlterarDTO : UsuarioInserirDTO
{
    public int Id { get; set; }
    public bool Ativo { get; set; }
}

public class UsuarioResultadoDTO : BaseErrorDTO
{
    public int Id { get; set; }
}
