namespace ParserOpenXML
{
    public class Chapter
    {
        public Chapter(string _num, string _title)
        {
            num = _num;
            title = _title;
        }

        public string num { get; set; }
        public string title { get; set; }
    }
}
