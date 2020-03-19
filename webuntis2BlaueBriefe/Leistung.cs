// Published under the terms of GPLv3 Stefan Bäumer 2019.

using System;
using System.Linq;

namespace webuntis2BlaueBriefe
{
    public class Leistung
    {
        public DateTime Datum { get; internal set; }
        public string Name { get; internal set; }
        public string Klasse { get; internal set; }
        public Fach Fach { get; internal set; }
        public string Prüfungsart { get; internal set; }
        public string Note { get; internal set; }
        public string Bemerkung { get; internal set; }
        public string Benutzer { get; internal set; }
        public int SchlüsselExtern { get; internal set; }
        public int LeistungId { get; internal set; }
        public bool ReligionAbgewählt { get; internal set; }

        internal bool IstKeinDiff(Stundentafels stundentafels)
        {
            if ((from s in stundentafels
                 from f in s.Fachs
                 where f.KürzelUntis == Fach.KürzelUntis
                 where Klasse.StartsWith(s.Name)
                 select f).Any())
            {
                return true;
            }
            else
            {
                return false;
            }            
        }
    }
}