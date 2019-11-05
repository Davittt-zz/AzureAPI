using System;

namespace HAA.DataModel
{
    public class Element
	{
		public int Id { get; set; }
		public string Name { get; set; }
		public DateTime Created { get; set; }
		public DateTime Modified { get; set; }
		public bool Active { get; set; }
	}
}
