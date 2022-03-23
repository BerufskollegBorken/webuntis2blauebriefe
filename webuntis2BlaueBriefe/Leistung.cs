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
        public string BlauerBriefNote { get; internal set; }
        public string Bemerkung { get; internal set; }
        public string Benutzer { get; internal set; }
        public int SchlüsselExtern { get; internal set; }
        public int LeistungId { get; internal set; }
        public bool ReligionAbgewählt { get; internal set; }
        public string Halbjahresgesamtnote { get; internal set; }

        internal bool IstKeinDiff(Klasses klasses)
        {
            foreach (var klasse in klasses)
            {
                if (klasse.NameUntis == this.Klasse)
                {
                    foreach (var fach in klasse.Stundentafel.Fachs)
                    {
                        if (fach.BezeichnungImZeugnis == this.Name)
                        {
                            return false;
                        }
                    }
                }
            }
            return true;
        }
    }
}