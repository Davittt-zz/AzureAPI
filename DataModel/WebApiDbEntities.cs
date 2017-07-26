namespace DataModel
{
	using System;
	using System.Data.Entity;
	using System.Linq;

	public class WebApiDbEntities : DbContext
	{
		public WebApiDbEntities()
			: base("name=WebApiDbEntities")
		{ 
			 
		}

		public virtual DbSet<User> Users { get; set; }
		//	public virtual DbSet<UserActivity> UserActivities { get; set; }
		public virtual DbSet<Element> Elements { get; set; }
		public virtual DbSet<Token> Tokens { get; set; }
	}
}