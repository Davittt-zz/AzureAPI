using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace HAA.DataModel
{
    public enum UserType { User = 0, Admin = 1, SuperUser = 2 }

    [Table("Users")]
    public class User
    {
        [Key]
        public int Index { get; set; }

        public string AccountNum { get; set; }
        
        public string Address1 { get; set; }
        
        public string Address2 { get; set; }
        
        public string Company { get; set; }
        
        public DateTime? CreateDate { get; set; }

        public DateTime? DateORecord { get; set; }
        
        public string Designer { get; set; }
        
        public string DesignerEmail { get; set; }
        
        public string DrawRefIndex { get; set; }
        
        public DateTime? Expiration { get; set; }

        public string HAACert { get; set; }

        public string Logo { get; set; }

        public string Note1 { get; set; }

        public string NumberOfProjects { get; set; }

        public string Phone { get; set; }

        public string Postal { get; set; }

        public string State { get; set; }

        public string UserID { get; set; }
        
        public string Password { get; set; }
    }
}