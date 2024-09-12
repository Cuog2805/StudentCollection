namespace StudentCollection.Models
{
    public class User
    {
        public int UserID { get; set; }
        public string UserName { get; set; }
        public string Password { get; set; }
        public string? FilePath { get; set; }
        public virtual ICollection<Student> Students { get; set; }
    }
}
