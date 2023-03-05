using System;
using System.Linq;

namespace webuntis2BlaueBriefe
{
    public class Fach
    {
        public int IdUntis { get; set; }
        public string KürzelUntis { get; set; }
        public string LangnameUntis { get; set; }
        public string BezeichnungImZeugnis { get; set; }
        public string Statistikname { get; set; }
        public int NoteHalbjahr { get; internal set; }
        public int NoteJetzt { get; internal set; }
        public bool NeuHinzugekommenesDefizitFach { get; internal set; }
        public bool NochmaligeVerschlechterungAuf6 { get; internal set; }

        public Fach()
        {
        }

        public Fach(string kürzelUntis, string bezeichnungImZeugnis, int noteJetzt, int noteHalbjahr)
        {
            KürzelUntis = kürzelUntis;
            BezeichnungImZeugnis = bezeichnungImZeugnis;
        }
    }
}