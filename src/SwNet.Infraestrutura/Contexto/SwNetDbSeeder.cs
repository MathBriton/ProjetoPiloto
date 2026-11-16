using Microsoft.EntityFrameworkCore;
using SwNet.Aplicacao.Autorizacao;
using SwNet.Dominio.Entidades;

namespace SwNet.Infraestrutura.Contexto;

/// <summary>
/// Seed de desenvolvimento. Cria um usuário fictício para testar o login
/// ponta a ponta: login "admin" / senha "Admin@123".
/// Chamar apenas em ambiente de desenvolvimento (ver Program.cs).
/// </summary>
public static class SwNetDbSeeder
{
    public const string LoginAdmin = "admin";
    public const string SenhaAdmin = "Admin@123";

    /// <summary>Aplica as migrations e semeia os dados.</summary>
    public static async Task ExecutarAsync(SwNetDbContext contexto)
    {
        await contexto.Database.MigrateAsync();
        await SemearAsync(contexto);
    }

    /// <summary>
    /// Idempotente: sincroniza o catálogo de permissões (Permissoes.Todas) e
    /// garante perfil Administrador com todas elas, plantas e usuário admin.
    /// </summary>
    public static async Task SemearAsync(SwNetDbContext contexto)
    {
        // 1) Catálogo de permissões
        var existentes = await contexto.Permissoes.ToDictionaryAsync(p => p.Chave);
        foreach (var (chave, descricao, modulo) in Permissoes.Todas)
        {
            if (!existentes.ContainsKey(chave))
                contexto.Permissoes.Add(new Permissao { Chave = chave, Descricao = descricao, Modulo = modulo });
        }
        await contexto.SaveChangesAsync();

        // 2) Perfil Administrador com todas as permissões
        var perfilAdmin = await contexto.Perfis.Include(p => p.PerfilPermissoes)
            .FirstOrDefaultAsync(p => p.Nome == "Administrador");
        if (perfilAdmin is null)
        {
            perfilAdmin = new Perfil { Nome = "Administrador", Descricao = "Acesso total ao sistema" };
            contexto.Perfis.Add(perfilAdmin);
            await contexto.SaveChangesAsync();
        }

        var idsNoPerfil = perfilAdmin.PerfilPermissoes.Select(pp => pp.PermissaoId).ToHashSet();
        foreach (var permissao in await contexto.Permissoes.ToListAsync())
        {
            if (!idsNoPerfil.Contains(permissao.Id))
                contexto.PerfilPermissoes.Add(new PerfilPermissao { PerfilId = perfilAdmin.Id, PermissaoId = permissao.Id });
        }

        // 3) Plantas
        if (!await contexto.Plantas.AnyAsync())
        {
            contexto.Plantas.AddRange(
                new Planta { Codigo = "JF01", Nome = "Juiz de Fora" },
                new Planta { Codigo = "SP01", Nome = "São Paulo" });
        }
        await contexto.SaveChangesAsync();

        // 4) Usuário admin
        if (await contexto.Usuarios.AnyAsync(u => u.Login == LoginAdmin)) return;

        var usuarioAdmin = new Usuario
        {
            Login = LoginAdmin,
            NomeCompleto = "Administrador (fictício)",
            Email = "admin@swnet.local",
            SenhaHash = BCrypt.Net.BCrypt.HashPassword(SenhaAdmin),
            Ativo = true,
            PerfilId = perfilAdmin.Id
        };
        contexto.Usuarios.Add(usuarioAdmin);
        await contexto.SaveChangesAsync();

        var plantas = await contexto.Plantas.OrderBy(p => p.Id).ToListAsync();
        contexto.UsuarioPlantas.AddRange(plantas.Select((p, i) =>
            new UsuarioPlanta { UsuarioId = usuarioAdmin.Id, PlantaId = p.Id, Padrao = i == 0 }));

        await contexto.SaveChangesAsync();
    }
}
