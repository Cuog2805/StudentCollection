namespace StudentCollection.Models
{
    public class Student
    {
        public int StudentID { get; set; }
        public int Stt { get; set; }
        public string? Name { get; set; }
        public string? Class { get; set; }
        public DateTime Birth { get; set; }
        public string? Gender { get; set; }
        public string? CurrentResidence { get; set; }
        public string? PermanentResidece { get; set; }
        public string? BirthPlace { get; set; }
        public string? FatherName { get; set; }
        public string? MotherName { get; set; }
        public string? PhoneNumber { get; set; }
        public int? UserID { get; set; }
        public virtual User User { get; set; }
    }
}
