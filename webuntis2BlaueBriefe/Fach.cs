namespace webuntis2BlaueBriefe
{
    public class Fach
    {
        public Fach()
        {
        }

        public Fach(string kürzelUntis, string bezeichnungImZeugnis, string noteJetzt, string noteHalbjahr)
        {
            KürzelUntis = kürzelUntis;
            BezeichnungImZeugnis = bezeichnungImZeugnis;
            NoteJetzt = noteJetzt;
            NoteHalbjahr = noteHalbjahr;
        }

        public int IdUntis { get; set; }
        public string KürzelUntis { get; set; }
        public string LangnameUntis { get; set; }
        public string BezeichnungImZeugnis { get; set; }
        public string Statistikname { get; set; }
        public string NoteHalbjahr { get; internal set; }
        public string NoteJetzt { get; internal set; }
    }
}