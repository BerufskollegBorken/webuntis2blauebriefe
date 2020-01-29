using System;
using System.Collections.Generic;

namespace webuntis2BlaueBriefe
{
    internal class Schueler
    {
        public int IdAtlantis { get; internal set; }
        public string Art { get; internal set; }
        public string Nachname { get; internal set; }
        public string SorgeberechtigtJn { get; internal set; }
        public string Anrede { get; internal set; }
        public string Vorname { get; internal set; }
        public object Telefons { get; internal set; }
        public string Plz { get; internal set; }
        public string Ort { get; internal set; }
        public string Strasse { get; internal set; }
        public string Email { get; internal set; }
        public string Klasse { get; private set; }
        public string Jahrgang { get; private set; }
        public DateTime Geburtsdatum { get; private set; }
        public bool Volljaehrig { get; private set; }
        public string GeschlechtMw { get; private set; }
        public List<string> Fachs { get; private set; }
        public string Typ { get; private set; }
        public string EVorname { get; private set; }
        public string ENachname { get; private set; }

        public Schueler(int idAtlantis, string typ, string klasse, string jahrgang, string nachname, string vorname, string enachname, string evorname, DateTime geburtsdatum, bool volljaehrig, string geschlechtMw, string sorgeberechtigtJn, string anrede, string plz, string ort, string strasse, List<string> fachs)
        {
            IdAtlantis = idAtlantis;
            Typ = typ;
            Klasse = klasse;
            Jahrgang = jahrgang;
            Nachname = nachname;
            Vorname = vorname;
            ENachname = enachname;
            EVorname = evorname;
            Geburtsdatum = geburtsdatum;
            Volljaehrig = volljaehrig;
            GeschlechtMw = geschlechtMw;
            SorgeberechtigtJn = sorgeberechtigtJn;
            Anrede = anrede;
            Plz = plz;
            Ort = ort;
            Strasse = strasse;
            Fachs = fachs;
        }
    }
}