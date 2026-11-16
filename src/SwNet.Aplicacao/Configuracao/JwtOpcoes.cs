using System.ComponentModel.DataAnnotations;

namespace SwNet.Aplicacao.Configuracao;

/// <summary>Seção "Jwt" do appsettings. Validada no startup (ValidateOnStart).</summary>
public class JwtOpcoes
{
    public const string Secao = "Jwt";

    /// <summary>Chave HMAC-SHA256. Mínimo de 32 caracteres (256 bits).
    /// Em produção, defina via variável de ambiente Jwt__ChaveSecreta ou user-secrets.</summary>
    [Required, MinLength(32)]
    public string ChaveSecreta { get; set; } = string.Empty;

    [Required]
    public string Emissor { get; set; } = string.Empty;

    [Required]
    public string Audiencia { get; set; } = string.Empty;

    /// <summary>Validade do access token. Curta de propósito: a renovação é via refresh token.</summary>
    [Range(1, 1440)]
    public int MinutosExpiracao { get; set; } = 15;
}
