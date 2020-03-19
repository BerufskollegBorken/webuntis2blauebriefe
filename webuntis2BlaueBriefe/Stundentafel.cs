using System.Collections.Generic;

namespace webuntis2BlaueBriefe
{
    public class Stundentafel
    {
        public Stundentafel()
        {
            Fachs = new List<Fach>();
            Fachklasses = new List<string>();
        }
        public int IdAtlantis { get; set; }
        public string Name { get; set; }
        public List<Fach> Fachs { get; internal set; }
        public List<string> Fachklasses { get; internal set; }
        public string GemeinsamesPräfixAllerKlassen { get; internal set; }
        public int IdUntis { get; internal set; }
        public string Langname { get; internal set; }
        public string Bemerkung { get; internal set; }
        public int AnzahlJahrgänge { get; internal set; }
    }
}