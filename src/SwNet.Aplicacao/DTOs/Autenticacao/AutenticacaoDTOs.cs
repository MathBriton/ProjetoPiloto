using System.ComponentModel.DataAnnotations;

namespace SwNet.Aplicacao.DTOs.Autenticacao;

public class LoginDTO
{
    [Required, MaxLength(60)]
    public string Login { get; set; } = string.Empty;

    [Required, MaxLength(128)]
    public string Senha { get; set; } = string.Empty;
}

public class RefreshTokenDTO
{
    [Required]
    public string RefreshToken { get; set; } = string.Empty;
}

public class TrocarPlantaDTO : RefreshTokenDTO
{
    [Range(1, int.MaxValue)]
    public int PlantaId { get; set; }
}

public class AlterarSenhaDTO
{
    [Required]
    public string SenhaAtual { get; set; } = string.Empty;

    [Required, MaxLength(128)]
    public string NovaSenha { get; set; } = string.Empty;
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
    public bool SenhaProvisoria { get; set; }
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
