using Microsoft.EntityFrameworkCore;

namespace ExcelReader;

public class ApplicationContext : DbContext
{
    public ApplicationContext()
    {
        Database.EnsureCreated();
    }
    public DbSet<Subject> Subjects => Set<Subject>();

    protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
    {
        optionsBuilder.UseSqlite("Data Source=Subjects.db");
    }
}