using Microsoft.EntityFrameworkCore;

namespace wyplaty.Entities
{
    public class FrachtDbContext : DbContext
    {
        private string _connectionString = "Server=(localdb)\\MSSQLLocalDB;Database=frachts;Trusted_Connection=True;";
        public DbSet<Fracht> Frachts { get; set; }
        public DbSet<Driver> Drivers { get; set; }



        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            optionsBuilder.UseSqlServer(_connectionString);
        }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<Fracht>(entity =>
            {
                entity.HasIndex(e => e.OrderNumber).IsUnique();
            });
        }
    }

}
