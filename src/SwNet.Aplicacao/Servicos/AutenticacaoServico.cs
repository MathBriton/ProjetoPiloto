using SwNet.Aplicacao.DTOs.Autenticacao;
using SwNet.Aplicacao.Interfaces;

namespace SwNet.Aplicacao.Servicos;

public class AutenticacaoServico
{
    private const int MaxTentativas = 5;
    private static readonly TimeSpan TempoBloqueio = TimeSpan.FromMinutes(15);

    private readonly IUsuarioRepositorio _usuarioRepositorio;
    private readonly ITokenServico _tokenServico;
    private readonly IEnumerable<IEstrategiaAutenticacao> _estrategias;

    public AutenticacaoServico(
        IUsuarioRepositorio usuarioRepositorio,
        ITokenServico tokenServico,
        IEnumerable<IEstrategiaAutenticacao> estrategias)
    {
        _usuarioRepositorio = usuarioRepositorio;
        _tokenServico = tokenServico;
        _estrategias = estrategias;
    }

    public async Task<LoginRetornoDTO> AutenticarAsync(LoginDTO dto, string metodo = "SISTEMA")
    {
        var usuario = await _usuarioRepositorio.ObterPorLoginAsync(dto.Login);

        if (usuario is null || !usuario.Ativo)
            return Erro("Usuário ou senha inválidos.");

        if (usuario.BloqueadoAte is not null && usuario.BloqueadoAte > DateTime.UtcNow)
            return Erro($"Usuário bloqueado até {usuario.BloqueadoAte:HH:mm}. Tente novamente mais tarde.");

        var estrategia = _estrategias.FirstOrDefault(e => e.Metodo == metodo)
            ?? throw new InvalidOperationException($"Estratégia de autenticação '{metodo}' não configurada.");

        var credenciaisValidas = await estrategia.ValidarCredenciaisAsync(usuario, dto.Senha);

        if (!credenciaisValidas)
        {
            await _usuarioRepositorio.RegistrarTentativaFalhaAsync(usuario.Id);
            var restantes = MaxTentativas - (usuario.TentativasFalhas + 1);
            return Erro(restantes > 0
                ? "Usuário ou senha inválidos."
                : $"Usuário bloqueado por {TempoBloqueio.Minutes} minutos após {MaxTentativas} tentativas.");
        }

        await _usuarioRepositorio.RegistrarLoginSucedidoAsync(usuario.Id);

        var permissoes = usuario.UsuarioPermissoes
            .Where(up => up.Concedida)
            .Select(up => up.Permissao!.Chave)
            .Union(usuario.Perfil?.PerfilPermissoes.Select(pp => pp.Permissao!.Chave) ?? Enumerable.Empty<string>())
            .Distinct()
            .ToList();

        var plantaAtual = usuario.UsuarioPlantas.FirstOrDefault(up => up.Padrao)
            ?? usuario.UsuarioPlantas.FirstOrDefault();

        var (token, expiraEm) = _tokenServico.GerarToken(usuario, permissoes, plantaAtual?.PlantaId ?? 0);

        return new LoginRetornoDTO
        {
            Sucesso = true,
            Token = token,
            RefreshToken = _tokenServico.GerarRefreshToken(),
            ExpiraEm = expiraEm,
            Usuario = new UsuarioLogadoDTO
            {
                Id = usuario.Id,
                Login = usuario.Login,
                NomeCompleto = usuario.NomeCompleto,
                Perfil = usuario.Perfil?.Nome ?? string.Empty,
                Permissoes = permissoes,
                PlantaAtualId = plantaAtual?.PlantaId ?? 0,
                Plantas = usuario.UsuarioPlantas.Select(up => new PlantaDTO
                {
                    Id = up.Planta!.Id,
                    Codigo = up.Planta.Codigo,
                    Nome = up.Planta.Nome
                }).ToList()
            }
        };
    }

    private static LoginRetornoDTO Erro(string mensagem) => new()
    {
        Sucesso = false,
        MensagemErro = mensagem
    };
}
