using SwNet.Dominio.Entidades;

namespace SwNet.Aplicacao.Autorizacao;

/// <summary>
/// Permissões efetivas = permissões do perfil + concedidas ao usuário − negadas ao usuário.
/// A negação direta (UsuarioPermissao.Concedida = false) sempre vence o perfil.
/// Exige que Perfil.PerfilPermissoes.Permissao e UsuarioPermissoes.Permissao estejam carregados.
/// </summary>
public static class PermissoesEfetivas
{
    public static List<string> Calcular(Usuario usuario)
    {
        var doPerfil = usuario.Perfil?.PerfilPermissoes.Select(pp => pp.Permissao!.Chave) ?? [];
        var concedidas = usuario.UsuarioPermissoes.Where(up => up.Concedida).Select(up => up.Permissao!.Chave);
        var negadas = usuario.UsuarioPermissoes.Where(up => !up.Concedida).Select(up => up.Permissao!.Chave).ToHashSet();

        return doPerfil.Union(concedidas)
            .Where(chave => !negadas.Contains(chave))
            .Distinct()
            .Order(StringComparer.Ordinal)
            .ToList();
    }
}
