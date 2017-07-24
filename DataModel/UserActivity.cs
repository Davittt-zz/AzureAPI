
namespace DataModel
{
	using System;
	using System.Collections.Generic;

	public class UserActivity
	{
		public int Id { get; set; }
		public int UserId { get; set; }
		public string Username { get; set; }

		public string Activity { get; set; }
		public DateTime Date { get; set; }
	}
}
