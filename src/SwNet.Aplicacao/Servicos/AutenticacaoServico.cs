using Microsoft.Extensions.Options;
using SwNet.Aplicacao.Autorizacao;
using SwNet.Aplicacao.Configuracao;
using SwNet.Aplicacao.DTOs.Autenticacao;
using SwNet.Aplicacao.Interfaces;
using SwNet.Dominio.Entidades;

namespace SwNet.Aplicacao.Servicos;

/// <summary>
/// Orquestra login, bloqueio por tentativas, emissão de JWT e o ciclo de vida
/// do refresh token (rotação + detecção de reuso). A validação da senha em si
/// fica na <see cref="IEstrategiaAutenticacao"/> do método escolhido.
/// </summary>
public class AutenticacaoServico
{
    private const string MensagemCredenciaisInvalidas = "Usuário ou senha inválidos.";
    private const string MensagemSessaoInvalida = "Sessão expirada ou inválida. Faça login novamente.";

    // Hash usado quando o login não existe, para o tempo de resposta não revelar
    // se o usuário existe ou não (mitiga enumeração de usuários por timing).
    private static readonly string HashFicticio = BCrypt.Net.BCrypt.HashPassword("swnet-hash-ficticio");

    private readonly IUsuarioRepositorio _usuarioRepositorio;
    private readonly IRefreshTokenRepositorio _refreshTokenRepositorio;
    private readonly ITokenServico _tokenServico;
    private readonly IEnumerable<IEstrategiaAutenticacao> _estrategias;
    private readonly AutenticacaoOpcoes _opcoes;
    private readonly TimeProvider _relogio;

    public AutenticacaoServico(
        IUsuarioRepositorio usuarioRepositorio,
        IRefreshTokenRepositorio refreshTokenRepositorio,
        ITokenServico tokenServico,
        IEnumerable<IEstrategiaAutenticacao> estrategias,
        IOptions<AutenticacaoOpcoes> opcoes,
        TimeProvider relogio)
    {
        _usuarioRepositorio = usuarioRepositorio;
        _refreshTokenRepositorio = refreshTokenRepositorio;
        _tokenServico = tokenServico;
        _estrategias = estrategias;
        _opcoes = opcoes.Value;
        _relogio = relogio;
    }

    private DateTime Agora => _relogio.GetUtcNow().UtcDateTime;

    public async Task<LoginRetornoDTO> AutenticarAsync(LoginDTO dto, string? ip = null, string metodo = "SISTEMA")
    {
        var usuario = await _usuarioRepositorio.ObterPorLoginAsync(dto.Login.Trim());

        if (usuario is null || !usuario.Ativo)
        {
            BCrypt.Net.BCrypt.Verify(dto.Senha, HashFicticio);
            return Erro(MensagemCredenciaisInvalidas);
        }

        if (usuario.EstaBloqueado(Agora))
            return Erro("Usuário temporariamente bloqueado por excesso de tentativas. Tente novamente mais tarde.");

        var estrategia = _estrategias.FirstOrDefault(e => e.Metodo == metodo)
            ?? throw new InvalidOperationException($"Estratégia de autenticação '{metodo}' não configurada.");

        if (!await estrategia.ValidarCredenciaisAsync(usuario, dto.Senha))
        {
            var bloqueou = usuario.RegistrarFalhaLogin(
                Agora, _opcoes.MaxTentativasFalhas, TimeSpan.FromMinutes(_opcoes.MinutosBloqueio));
            await _usuarioRepositorio.SalvarAsync();

            return Erro(bloqueou
                ? $"Usuário bloqueado por {_opcoes.MinutosBloqueio} minutos após {_opcoes.MaxTentativasFalhas} tentativas inválidas."
                : MensagemCredenciaisInvalidas);
        }

        usuario.RegistrarLoginSucedido(Agora);

        var plantaInicial = usuario.UsuarioPlantas.FirstOrDefault(up => up.Padrao)?.PlantaId
            ?? usuario.UsuarioPlantas.FirstOrDefault()?.PlantaId
            ?? 0;

        return await EmitirSessaoAsync(usuario, plantaInicial, Guid.NewGuid(), ip);
    }

    /// <summary>
    /// Troca um refresh token válido por um novo par (access + refresh). O token
    /// apresentado é revogado; se ele já estava revogado, é reuso: revoga a família toda.
    /// </summary>
    public async Task<LoginRetornoDTO> RenovarAsync(string refreshToken, string? ip = null)
    {
        var atual = await _refreshTokenRepositorio.ObterPorHashAsync(_tokenServico.CalcularHash(refreshToken));
        if (atual is null)
            return Erro(MensagemSessaoInvalida);

        if (atual.Revogado)
        {
            await _refreshTokenRepositorio.RevogarFamiliaAsync(atual.Familia, Agora, ip, "Reuso de refresh token detectado");
            await _usuarioRepositorio.SalvarAsync();
            return Erro(MensagemSessaoInvalida);
        }

        if (atual.Expirado(Agora))
            return Erro(MensagemSessaoInvalida);

        var usuario = await _usuarioRepositorio.ObterPorIdComPermissoesAsync(atual.UsuarioId);
        if (usuario is null || !usuario.Ativo || usuario.EstaBloqueado(Agora))
        {
            await _refreshTokenRepositorio.RevogarFamiliaAsync(atual.Familia, Agora, ip, "Usuário inativo ou bloqueado");
            await _usuarioRepositorio.SalvarAsync();
            return Erro(MensagemSessaoInvalida);
        }

        // Se perdeu o acesso à planta da sessão, volta para a padrão
        var plantaId = usuario.UsuarioPlantas.Any(up => up.PlantaId == atual.PlantaId)
            ? atual.PlantaId
            : usuario.UsuarioPlantas.FirstOrDefault(up => up.Padrao)?.PlantaId ?? usuario.UsuarioPlantas.FirstOrDefault()?.PlantaId ?? 0;

        return await EmitirSessaoAsync(usuario, plantaId, atual.Familia, ip, tokenAnterior: atual, motivoRevogacao: "Substituído na renovação");
    }

    /// <summary>Reemite os tokens com outra planta ativa (seletor da topbar).</summary>
    public async Task<LoginRetornoDTO> TrocarPlantaAsync(int usuarioId, TrocarPlantaDTO dto, string? ip = null)
    {
        var atual = await _refreshTokenRepositorio.ObterPorHashAsync(_tokenServico.CalcularHash(dto.RefreshToken));
        if (atual is null || atual.UsuarioId != usuarioId || !atual.Ativo(Agora))
            return Erro(MensagemSessaoInvalida);

        var usuario = await _usuarioRepositorio.ObterPorIdComPermissoesAsync(usuarioId);
        if (usuario is null || !usuario.Ativo)
            return Erro(MensagemSessaoInvalida);

        if (usuario.UsuarioPlantas.All(up => up.PlantaId != dto.PlantaId))
            return Erro("Você não tem acesso a esta planta.");

        return await EmitirSessaoAsync(usuario, dto.PlantaId, atual.Familia, ip, tokenAnterior: atual, motivoRevogacao: "Troca de planta");
    }

    /// <summary>Logout: revoga a família do refresh token (todas as renovações daquela sessão).</summary>
    public async Task SairAsync(string refreshToken, string? ip = null)
    {
        var atual = await _refreshTokenRepositorio.ObterPorHashAsync(_tokenServico.CalcularHash(refreshToken));
        if (atual is null) return;

        await _refreshTokenRepositorio.RevogarFamiliaAsync(atual.Familia, Agora, ip, "Logout");
        await _usuarioRepositorio.SalvarAsync();
    }

    /// <summary>Encerra todas as sessões do usuário (todos os dispositivos).</summary>
    public async Task SairDeTodasAsSessoesAsync(int usuarioId, string? ip = null)
    {
        await _refreshTokenRepositorio.RevogarTodosDoUsuarioAsync(usuarioId, Agora, ip, "Logout de todas as sessões");
        await _usuarioRepositorio.SalvarAsync();
    }

    public async Task<UsuarioLogadoDTO?> ObterUsuarioLogadoAsync(int usuarioId, int plantaAtualId)
    {
        var usuario = await _usuarioRepositorio.ObterPorIdComPermissoesAsync(usuarioId);
        return usuario is null || !usuario.Ativo
            ? null
            : MontarUsuarioLogado(usuario, PermissoesEfetivas.Calcular(usuario), plantaAtualId);
    }

    /// <summary>Troca de senha do próprio usuário. Encerra as demais sessões ao concluir.</summary>
    public async Task<BaseErrorDTO> AlterarSenhaAsync(int usuarioId, AlterarSenhaDTO dto, string? ip = null)
    {
        var usuario = await _usuarioRepositorio.ObterPorIdComPermissoesAsync(usuarioId);
        if (usuario is null || !usuario.Ativo)
            return new BaseErrorDTO { Sucesso = false, MensagemErro = MensagemSessaoInvalida };

        if (!BCrypt.Net.BCrypt.Verify(dto.SenhaAtual, usuario.SenhaHash))
            return new BaseErrorDTO { Sucesso = false, MensagemErro = "Senha atual incorreta." };

        var erroPolitica = ValidarPoliticaSenha(dto.NovaSenha, usuario.Login);
        if (erroPolitica is not null)
            return new BaseErrorDTO { Sucesso = false, MensagemErro = erroPolitica };

        if (BCrypt.Net.BCrypt.Verify(dto.NovaSenha, usuario.SenhaHash))
            return new BaseErrorDTO { Sucesso = false, MensagemErro = "A nova senha deve ser diferente da atual." };

        usuario.SenhaHash = BCrypt.Net.BCrypt.HashPassword(dto.NovaSenha);
        usuario.SenhaProvisoria = false;

        await _refreshTokenRepositorio.RevogarTodosDoUsuarioAsync(usuarioId, Agora, ip, "Senha alterada");
        await _usuarioRepositorio.SalvarAsync();

        return new BaseErrorDTO();
    }

    public string? ValidarPoliticaSenha(string senha, string login)
    {
        if (senha.Length < _opcoes.TamanhoMinimoSenha)
            return $"A senha deve ter pelo menos {_opcoes.TamanhoMinimoSenha} caracteres.";
        if (!senha.Any(char.IsUpper) || !senha.Any(char.IsLower) || !senha.Any(char.IsDigit))
            return "A senha deve conter letras maiúsculas, minúsculas e números.";
        if (senha.Contains(login, StringComparison.OrdinalIgnoreCase))
            return "A senha não pode conter o login.";
        return null;
    }

    private async Task<LoginRetornoDTO> EmitirSessaoAsync(
        Usuario usuario, int plantaId, Guid familia, string? ip,
        RefreshToken? tokenAnterior = null, string? motivoRevogacao = null)
    {
        var permissoes = PermissoesEfetivas.Calcular(usuario);
        var (token, expiraEm) = _tokenServico.GerarToken(usuario, permissoes, plantaId);

        var refreshToken = _tokenServico.GerarRefreshToken();
        var refreshHash = _tokenServico.CalcularHash(refreshToken);

        tokenAnterior?.Revogar(Agora, ip, motivoRevogacao ?? "Substituído", refreshHash);

        _refreshTokenRepositorio.Adicionar(new RefreshToken
        {
            UsuarioId = usuario.Id,
            TokenHash = refreshHash,
            Familia = familia,
            PlantaId = plantaId,
            CriadoEm = Agora,
            CriadoPorIp = ip,
            ExpiraEm = Agora.AddDays(_opcoes.DiasValidadeRefreshToken)
        });

        await _usuarioRepositorio.SalvarAsync();

        return new LoginRetornoDTO
        {
            Sucesso = true,
            Token = token,
            RefreshToken = refreshToken,
            ExpiraEm = expiraEm,
            Usuario = MontarUsuarioLogado(usuario, permissoes, plantaId)
        };
    }

    private static UsuarioLogadoDTO MontarUsuarioLogado(Usuario usuario, List<string> permissoes, int plantaAtualId) => new()
    {
        Id = usuario.Id,
        Login = usuario.Login,
        NomeCompleto = usuario.NomeCompleto,
        Perfil = usuario.Perfil?.Nome ?? string.Empty,
        SenhaProvisoria = usuario.SenhaProvisoria,
        Permissoes = permissoes,
        PlantaAtualId = plantaAtualId,
        Plantas = usuario.UsuarioPlantas
            .Where(up => up.Planta is not null)
            .Select(up => new PlantaDTO { Id = up.Planta!.Id, Codigo = up.Planta.Codigo, Nome = up.Planta.Nome })
            .ToList()
    };

    private static LoginRetornoDTO Erro(string mensagem) => new() { Sucesso = false, MensagemErro = mensagem };
}
