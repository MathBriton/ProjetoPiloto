namespace SwNet.Aplicacao.Autorizacao;

/// <summary>
/// Catálogo de permissões do sistema. Use as constantes no [RequerPermissao]
/// em vez de strings soltas. O seed de desenvolvimento cria todas na tabela
/// Permissoes e vincula ao perfil Administrador.
/// </summary>
public static class Permissoes
{
    public const string AcessosUsuariosVisualizar = "acessos-usuarios.visualizar";
    public const string AcessosUsuariosEditar = "acessos-usuarios.editar";
    public const string UsuariosIncluir = "usuarios.incluir";
    public const string UsuariosAlterar = "usuarios.alterar";
    public const string UsuariosAtivarDesativar = "usuarios.ativar-desativar";
    public const string UsuariosCancelar = "usuarios.cancelar";
    public const string UsuariosCriacaoAd = "usuarios.criacao-ad";
    public const string UsuariosRevisao = "usuarios.revisao";
    public const string MenusGerenciar = "menus.gerenciar";
    public const string PermissoesEspeciaisEditar = "permissoes-especiais.editar";
    public const string PermissoesRelatorio = "permissoes.relatorio";
    public const string RecebimentoControladoria = "recebimento.controladoria";

    public static readonly IReadOnlyList<(string Chave, string Descricao, string Modulo)> Todas =
    [
        (AcessosUsuariosVisualizar, "Visualizar usuários", "Acessos"),
        (AcessosUsuariosEditar, "Editar usuários", "Acessos"),
        (UsuariosIncluir, "Incluir usuário", "Administrador"),
        (UsuariosAlterar, "Alterar usuário", "Administrador"),
        (UsuariosAtivarDesativar, "Ativar/desativar usuário", "Administrador"),
        (UsuariosCancelar, "Cancelar usuários", "Administrador"),
        (UsuariosCriacaoAd, "Criação de usuário no AD", "Administrador"),
        (UsuariosRevisao, "Revisão de usuários", "Administrador"),
        (MenusGerenciar, "Gerenciar menus", "Administrador"),
        (PermissoesEspeciaisEditar, "Editar permissões especiais", "Administrador"),
        (PermissoesRelatorio, "Relatório de permissões", "Administrador"),
        (RecebimentoControladoria, "Controladoria no Controle de Recebimento", "Compras"),
    ];
}
