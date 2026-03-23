
using Microsoft.EntityFrameworkCore;
using gool_4337Group_26.Models;

namespace gool_4337Group_26.Data
{
    public class AppDbContext : DbContext
    {
        public DbSet<Client> Clients { get; set; }

        protected override void OnConfiguring(DbContextOptionsBuilder options)
            => options.UseSqlite("Data Source=clients.db");
    }
}
