namespace HAA.DataModel.Migrations
{
    using System.Data.Entity.Migrations;

    internal sealed class Configuration : DbMigrationsConfiguration<HAA.DataModel.WebApiDbEntities>
    {
        public Configuration()
        {
            AutomaticMigrationsEnabled = true;
        }
    }
}
