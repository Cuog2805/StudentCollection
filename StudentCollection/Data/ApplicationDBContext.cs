using Microsoft.EntityFrameworkCore;
using StudentCollection.Models;

namespace StudentCollection.Data
{
    public class ApplicationDBContext: DbContext
    {
        public ApplicationDBContext(DbContextOptions<ApplicationDBContext> option): base(option)
        {
        }
        public DbSet<Student> Students { get; set; }
        public DbSet<User> Users { get; set; }
        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<Student>().HasKey(m => m.StudentID);
            modelBuilder.Entity<User>().HasKey(m => m.UserID);

            modelBuilder.Entity<User>()
                .HasMany(u => u.Students)
                .WithOne(s => s.User)
                .IsRequired(false);

            base.OnModelCreating(modelBuilder);
        }
    }
}
