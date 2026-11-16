using Microsoft.EntityFrameworkCore;
using SwNet.Dominio.Entidades;

namespace SwNet.Infraestrutura.Contexto;

public class SwNetDbContext : DbContext
{
    public SwNetDbContext(DbContextOptions<SwNetDbContext> options) : base(options) { }

    public DbSet<Usuario> Usuarios => Set<Usuario>();
    public DbSet<Perfil> Perfis => Set<Perfil>();
    public DbSet<Permissao> Permissoes => Set<Permissao>();
    public DbSet<PerfilPermissao> PerfilPermissoes => Set<PerfilPermissao>();
    public DbSet<UsuarioPermissao> UsuarioPermissoes => Set<UsuarioPermissao>();
    public DbSet<Planta> Plantas => Set<Planta>();
    public DbSet<UsuarioPlanta> UsuarioPlantas => Set<UsuarioPlanta>();
    public DbSet<RefreshToken> RefreshTokens => Set<RefreshToken>();

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        modelBuilder.Entity<Usuario>(e =>
        {
            e.HasIndex(u => u.Login).IsUnique();
            e.Property(u => u.Login).HasMaxLength(60);
            e.Property(u => u.NomeCompleto).HasMaxLength(150);
            e.Property(u => u.Email).HasMaxLength(150);
            e.Property(u => u.SenhaHash).HasMaxLength(100);
            e.HasOne(u => u.Perfil).WithMany(p => p.Usuarios).HasForeignKey(u => u.PerfilId);
        });

        modelBuilder.Entity<Permissao>(e =>
        {
            e.HasIndex(p => p.Chave).IsUnique();
            e.Property(p => p.Chave).HasMaxLength(120);
            e.Property(p => p.Descricao).HasMaxLength(200);
            e.Property(p => p.Modulo).HasMaxLength(60);
        });

        modelBuilder.Entity<PerfilPermissao>(e =>
        {
            e.HasKey(pp => new { pp.PerfilId, pp.PermissaoId });
            e.HasOne(pp => pp.Perfil).WithMany(p => p.PerfilPermissoes).HasForeignKey(pp => pp.PerfilId);
            e.HasOne(pp => pp.Permissao).WithMany(p => p.PerfilPermissoes).HasForeignKey(pp => pp.PermissaoId);
        });

        modelBuilder.Entity<UsuarioPermissao>(e =>
        {
            e.HasKey(up => new { up.UsuarioId, up.PermissaoId });
            e.HasOne(up => up.Usuario).WithMany(u => u.UsuarioPermissoes).HasForeignKey(up => up.UsuarioId);
            e.HasOne(up => up.Permissao).WithMany(p => p.UsuarioPermissoes).HasForeignKey(up => up.PermissaoId);
        });

        modelBuilder.Entity<Planta>(e =>
        {
            e.HasIndex(p => p.Codigo).IsUnique();
            e.Property(p => p.Codigo).HasMaxLength(20);
            e.Property(p => p.Nome).HasMaxLength(120);
        });

        modelBuilder.Entity<UsuarioPlanta>(e =>
        {
            e.HasKey(up => new { up.UsuarioId, up.PlantaId });
            e.HasOne(up => up.Usuario).WithMany(u => u.UsuarioPlantas).HasForeignKey(up => up.UsuarioId);
            e.HasOne(up => up.Planta).WithMany(p => p.UsuarioPlantas).HasForeignKey(up => up.PlantaId);
        });

        modelBuilder.Entity<Perfil>(e =>
        {
            e.HasIndex(p => p.Nome).IsUnique();
            e.Property(p => p.Nome).HasMaxLength(80);
            e.Property(p => p.Descricao).HasMaxLength(200);
        });

        modelBuilder.Entity<RefreshToken>(e =>
        {
            e.ToTable("RefreshTokens");
            e.HasIndex(rt => rt.TokenHash).IsUnique();
            e.HasIndex(rt => rt.Familia);
            e.HasIndex(rt => new { rt.UsuarioId, rt.RevogadoEm });
            e.Property(rt => rt.TokenHash).HasMaxLength(64).IsRequired();
            e.Property(rt => rt.SubstituidoPorHash).HasMaxLength(64);
            e.Property(rt => rt.CriadoPorIp).HasMaxLength(45);
            e.Property(rt => rt.RevogadoPorIp).HasMaxLength(45);
            e.Property(rt => rt.MotivoRevogacao).HasMaxLength(120);
            e.HasOne(rt => rt.Usuario).WithMany(u => u.RefreshTokens)
                .HasForeignKey(rt => rt.UsuarioId).OnDelete(DeleteBehavior.Cascade);
        });
    }
}
