using System.ComponentModel.DataAnnotations;

namespace SwNet.Aplicacao.Configuracao;

/// <summary>Seção "Autenticacao" do appsettings: bloqueio, refresh token e política de senha.</summary>
public class AutenticacaoOpcoes
{
    public const string Secao = "Autenticacao";

    [Range(1, 50)]
    public int MaxTentativasFalhas { get; set; } = 5;

    [Range(1, 1440)]
    public int MinutosBloqueio { get; set; } = 15;

    [Range(1, 90)]
    public int DiasValidadeRefreshToken { get; set; } = 7;

    [Range(6, 128)]
    public int TamanhoMinimoSenha { get; set; } = 8;
}
