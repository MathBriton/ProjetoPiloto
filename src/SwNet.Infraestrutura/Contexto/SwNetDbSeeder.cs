using Microsoft.EntityFrameworkCore;
using SwNet.Dominio.Entidades;

namespace SwNet.Infraestrutura.Contexto;

/// <summary>
/// Seed de desenvolvimento. Cria um usuário fictício para testar o login
/// ponta a ponta: login "admin" / senha "Admin@123".
/// Chamar apenas em ambiente de desenvolvimento (ver Program.cs).
/// </summary>
public static class SwNetDbSeeder
{
    public static async Task ExecutarAsync(SwNetDbContext contexto)
    {
        await contexto.Database.MigrateAsync();

        if (await contexto.Permissoes.AnyAsync()) return; // já semeado

        var permissaoVisualizar = new Permissao { Chave = "acessos-usuarios.visualizar", Descricao = "Visualizar usuários", Modulo = "Acessos" };
        var permissaoEditar = new Permissao { Chave = "acessos-usuarios.editar", Descricao = "Editar usuários", Modulo = "Acessos" };
        contexto.Permissoes.AddRange(permissaoVisualizar, permissaoEditar);

        var perfilAdmin = new Perfil { Nome = "Administrador", Descricao = "Acesso total ao sistema" };
        contexto.Perfis.Add(perfilAdmin);

        var plantaJf = new Planta { Codigo = "JF01", Nome = "Juiz de Fora" };
        var plantaSp = new Planta { Codigo = "SP01", Nome = "São Paulo" };
        contexto.Plantas.AddRange(plantaJf, plantaSp);

        await contexto.SaveChangesAsync();

        contexto.PerfilPermissoes.AddRange(
            new PerfilPermissao { PerfilId = perfilAdmin.Id, PermissaoId = permissaoVisualizar.Id },
            new PerfilPermissao { PerfilId = perfilAdmin.Id, PermissaoId = permissaoEditar.Id }
        );

        var usuarioAdmin = new Usuario
        {
            Login = "admin",
            NomeCompleto = "Administrador (fictício)",
            Email = "admin@swnet.local",
            SenhaHash = BCrypt.Net.BCrypt.HashPassword("Admin@123"),
            Ativo = true,
            PerfilId = perfilAdmin.Id
        };
        contexto.Usuarios.Add(usuarioAdmin);
        await contexto.SaveChangesAsync();

        contexto.UsuarioPlantas.AddRange(
            new UsuarioPlanta { UsuarioId = usuarioAdmin.Id, PlantaId = plantaJf.Id, Padrao = true },
            new UsuarioPlanta { UsuarioId = usuarioAdmin.Id, PlantaId = plantaSp.Id, Padrao = false }
        );

        await contexto.SaveChangesAsync();
    }
}
