namespace HAA.DataModel
{
    using System;
    using System.ComponentModel.DataAnnotations.Schema;

    [Table("Token")]
    public class Token
    {
        public int Id { get; set; }
        public int UserId { get; set; }
        public string AuthToken { get; set; }
        public DateTime IssuedOn { get; set; }
        public DateTime ExpiresOn { get; set; }
    }
}
