
namespace DataModel
{
	using System;
    using System.Collections.Generic;

    public partial class Token
    {
        public int ID { get; set; }
        public int UserID { get; set; }
        public int StoreID { get; set; }
        public string TokenID { get; set; }
        public Nullable<System.DateTime> Created { get; set; }
        public Nullable<System.DateTime> LastUpdate { get; set; }
    }
}
