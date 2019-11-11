using System.Data.Entity;

namespace HAA.DataModel
{
    public class WebApiDbEntities : DbContext
    {
        public WebApiDbEntities()
            : base("name=WebApiDbEntities")
        {
        }

        public virtual DbSet<User> Users { get; set; }

        public virtual DbSet<Element> Elements { get; set; }
        
        public virtual DbSet<Token> Tokens { get; set; }

        public virtual DbSet<SpeakerConfig> SpeakerConfigs { get; set; }

        public virtual DbSet<Project> Projects { get; set; }

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            base.OnModelCreating(modelBuilder);

            modelBuilder.Entity<SpeakerConfig>()
                .HasKey(y => y.Index)
                .HasRequired(x => x.Project)
                .WithMany()
                .HasForeignKey(y => y.ProjectId);
        }
    }
}