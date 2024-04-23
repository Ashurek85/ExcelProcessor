namespace ExcelProcessor.Examples.Reader.Reader.Entities
{
    public class StudentContext
    {
        public string University { get; set; }
        public DateTime GeneratedAt { get; set; }
        public List<Student> Students { get; set; }
    }
}
