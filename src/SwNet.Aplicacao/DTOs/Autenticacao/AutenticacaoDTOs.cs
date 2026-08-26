namespace SwNet.Aplicacao.DTOs.Autenticacao;

public class LoginDTO
{
    public string Login { get; set; } = string.Empty;
    public string Senha { get; set; } = string.Empty;
}

public class LoginRetornoDTO : BaseErrorDTO
{
    public string? Token { get; set; }
    public string? RefreshToken { get; set; }
    public DateTime ExpiraEm { get; set; }
    public UsuarioLogadoDTO? Usuario { get; set; }
}

public class UsuarioLogadoDTO
{
    public int Id { get; set; }
    public string Login { get; set; } = string.Empty;
    public string NomeCompleto { get; set; } = string.Empty;
    public string Perfil { get; set; } = string.Empty;
    public List<string> Permissoes { get; set; } = new();
    public List<PlantaDTO> Plantas { get; set; } = new();
    public int PlantaAtualId { get; set; }
}

public class PlantaDTO
{
    public int Id { get; set; }
    public string Codigo { get; set; } = string.Empty;
    public string Nome { get; set; } = string.Empty;
}

public class BaseErrorDTO
{
    public bool Sucesso { get; set; } = true;
    public string? MensagemErro { get; set; }
}
