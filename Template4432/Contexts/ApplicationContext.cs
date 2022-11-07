using System.Data.Entity;
using Template4432.Models;

namespace Template4432.Contexts
{
    public class ApplicationContext : DbContext
    {
        public DbSet<SkiService> SkiServices { get; set; }
        
        public ApplicationContext() : base("SkiRentPoint") { }
    }
}