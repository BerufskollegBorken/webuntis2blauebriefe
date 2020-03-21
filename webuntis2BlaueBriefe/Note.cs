namespace webuntis2BlaueBriefe
{
    public class Note
    {
        public string Klartext { get; private set; }
        public string Stufe { get; private set; }

        public Note(string stufe, string klartext)
        {
            Klartext = klartext;
            Stufe = stufe;
        }
    }
}