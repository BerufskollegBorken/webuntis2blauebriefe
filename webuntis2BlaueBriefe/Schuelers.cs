using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Odbc;
using System.Linq;

namespace webuntis2BlaueBriefe
{
    internal class Schuelers : List<Schueler>
    {        
        public Schuelers()
        {
        }

        public Schuelers(string aktSj, string connectionStringAtlantis, DefizitäreLeistungen defizitäreLeistungen, Klasses klasses, Lehrers lehrers)
        {
            foreach (var idAtlantis in (from t in defizitäreLeistungen select t.SchlüsselExtern).Distinct().ToList())
            {
                using (OdbcConnection connection = new OdbcConnection(connectionStringAtlantis))
                {
                    connection.Open();

                    if (connection != null)
                    {
                        OdbcCommand command = connection.CreateCommand();
                        command.CommandText = @"SELECT 
DBA.schue_sj.pu_id as ID,
DBA.schue_sj.s_jahrgang AS Jahrgang,
DBA.adresse.s_typ_adr as Typ,
DBA.klasse.klasse as Klasse,
DBA.schueler.name_1 as Vorname,
DBA.schueler.name_2 as Nachname,
DBA.adresse.name_2 AS EVorname,
DBA.adresse.name_1 AS ENachname,
DBA.schueler.dat_geburt as Geburtsdatum,
DBA.schueler.s_geschl as Geschlecht,
DBA.adresse.strasse AS Strasse,
DBA.adresse.plz AS Plz,
DBA.adresse.ort AS Ort,
DBA.adresse.sorge_berechtigt_jn,
DBA.adresse.s_anrede,
DBA.schueler.s_erzb_1_art,
DBA.schueler.s_erzb_2_art,
DBA.schueler.id_hauptadresse,
DBA.adresse.hauptadresse_jn,
DBA.adresse.anrede_text,
DBA.schueler.anrede_text,
DBA.adresse.name_3,
DBA.adresse.plz_postfach as PlzPostfach,
DBA.adresse.postfach as Postfach,
DBA.adresse.s_titel_ad,
DBA.adresse.s_sorgerecht,
DBA.adresse.brief_adresse,
DBA.schue_sj.kl_id, 
DBA.adresse.s_famstand_adr
FROM((DBA.schue_sj JOIN DBA.klasse ON DBA.schue_sj.kl_id = DBA.klasse.kl_id) JOIN DBA.schueler ON DBA.schue_sj.pu_id = DBA.schueler.pu_id) JOIN DBA.adresse ON DBA.schueler.pu_id = DBA.adresse.pu_id
WHERE vorgang_schuljahr = '" + aktSj + @"' AND schue_sj.pu_id = " + idAtlantis + ";";

                        OdbcDataReader reader = command.ExecuteReader();

                        int fCount = reader.FieldCount;

                        var x = (from s in this where s.IdAtlantis == idAtlantis select s).FirstOrDefault();

                        Schueler schueler = new Schueler();
                        schueler.Sorgeberechtigte = new Sorgeberechtigte();

                        while (reader.Read())
                        {
                            schueler.IdAtlantis = Convert.ToInt32(reader.GetValue(0));
                            var sorgeberechtigtJn = reader.GetValue(13).ToString();
                            var anrede = reader.GetValue(14).ToString();

                            if (sorgeberechtigtJn == "N")
                            {
                                schueler.Vorname = reader.GetValue(5).ToString();
                                schueler.Nachname = reader.GetValue(4).ToString();
                                schueler.Strasse = reader.GetValue(10).ToString();
                                schueler.Plz = reader.GetValue(11).ToString();
                                schueler.Ort = reader.GetValue(12).ToString();                                
                                schueler.Geburtsdatum = DateTime.ParseExact(reader.GetValue(8).ToString(), "dd.MM.yyyy HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
                                schueler.Volljaehrig = schueler.Geburtsdatum.AddYears(18) > DateTime.Now ? false : true;
                                schueler.Geschlecht = reader.GetValue(9).ToString();                                
                                schueler.Jahrgang = reader.GetValue(1).ToString();
                                schueler.Anrede = anrede;
                                schueler.Typ = reader.GetValue(2).ToString(); // 0 = Schüler V = Vater  M = Mutter
                                schueler.Klasse = reader.GetValue(3).ToString();
                                schueler.Fachs = new Fachs();
                                schueler.Klassenleitung = (from k in klasses where k.NameUntis == schueler.Klasse select k.Klassenleitungen[0].Vorname + " " + k.Klassenleitungen[0].Nachname).FirstOrDefault();
                                schueler.KlassenleitungMail = (from k in klasses where k.NameUntis == schueler.Klasse select k.Klassenleitungen[0].Mail).FirstOrDefault();
                                schueler.KlassenleitungMw = (from l in lehrers where l.Vorname + " " + l.Nachname == schueler.Klassenleitung select l.Anrede).FirstOrDefault();
                            }

                            if (sorgeberechtigtJn == "J" && anrede == "H")
                            {
                                var vorname = reader.GetValue(6).ToString();
                                var nachname = reader.GetValue(7).ToString();
                                var strasse = reader.GetValue(10).ToString();
                                var plz = reader.GetValue(11).ToString();
                                var ort = reader.GetValue(12).ToString();
                                schueler.Sorgeberechtigte.Add(new Sorgeberechtigt(vorname, nachname, strasse, plz, ort));                                
                            }

                            if (sorgeberechtigtJn == "J" && anrede == "F")
                            {
                                var vorname = reader.GetValue(6).ToString();
                                var nachname = reader.GetValue(7).ToString();
                                var strasse = reader.GetValue(10).ToString();
                                var plz = reader.GetValue(11).ToString();
                                var ort = reader.GetValue(12).ToString();
                                schueler.Sorgeberechtigte.Add(new Sorgeberechtigt(vorname, nachname, strasse, plz, ort));
                            }
                        }

                        defizitäreLeistungen.Get(schueler);

                        this.Add(schueler);

                        reader.Close();
                        command.Dispose();
                    }
                }
            }
        }

        internal void RenderBriefe()
        {
            foreach (var schueler in this)
            {
                Console.Write((schueler.Nachname + " " + schueler.Vorname + " (" + schueler.Klasse + "):").PadRight(35));

                if ((from f in schueler.Fachs
                     where !Global.Mangelhaft.Contains(f.NoteHalbjahr)
                     where !Global.Ungenügend.Contains(f.NoteHalbjahr)
                     select f).Any())
                {
                     Console.Write(("HZ: Kein Defizit. ").PadRight(25));
                    
                    if ((from f in schueler.Fachs
                         where Global.Mangelhaft.Contains(f.NoteJetzt)
                         select f).Count() == 1)
                    {
                        Console.Write("Jetzt eine 5: Mitteilung über Leistungsstand ...");
                        schueler.RenderMitteilung("M");
                        Console.WriteLine("... ok.");
                    }

                    // Jetzt zwei oder mehr 5: Gefährdung

                    if ((from f in schueler.Fachs
                         where Global.Mangelhaft.Contains(f.NoteJetzt)
                         select f).Count() > 1)
                    {
                        Console.Write("Jetzt mehr als eine 5: Gefährdung ...");
                        schueler.RenderMitteilung("G");
                        Console.WriteLine("... ok.");                        
                    }

                    // Jetzt eine 6 oder mehr: Gefährdung

                    if ((from f in schueler.Fachs
                         where Global.Ungenügend.Contains(f.NoteJetzt)
                         select f).Count() > 0)
                    {
                        Console.Write("Jetzt 1 oder mehr 6en: Gefährdung ...");
                        schueler.RenderMitteilung("G");
                        Console.WriteLine("... ok.");
                    }
                }
                
                if ((from f in schueler.Fachs
                     where Global.Mangelhaft.Contains(f.NoteHalbjahr)                     
                     select f).Count() == 1)
                {
                    Console.Write(("HZ: Eine 5. ").PadRight(25));
                    
                    if ((from f in schueler.Fachs
                         where Global.Mangelhaft.Contains(f.NoteJetzt)
                         select f).Count() > (from f in schueler.Fachs
                                              where Global.Mangelhaft.Contains(f.NoteHalbjahr)
                                              select f).Count())
                    {
                        Console.WriteLine(" Jetzt eine oder mehrere zusätzliche 5en: Gefährdung ...");
                        schueler.RenderMitteilung("G");
                        Console.WriteLine("... ok.");
                    }
                    
                    if ((from f in schueler.Fachs
                         where Global.Ungenügend.Contains(f.NoteJetzt)
                         select f).Count() > (from f in schueler.Fachs
                                              where Global.Ungenügend.Contains(f.NoteHalbjahr) select f).Count())
                    {
                        Console.WriteLine(" Jetzt eine oder mehr zusätzliche 6: Gefährdung ...");
                        schueler.RenderMitteilung("G");
                        Console.WriteLine("... ok.");
                    }
                }
                
                if ((from f in schueler.Fachs where Global.Ungenügend.Contains(f.NoteHalbjahr) select f).Count() >= 1 || 
                    (from f in schueler.Fachs where Global.Mangelhaft.Contains(f.NoteHalbjahr) select f).Count() > 1)
                {
                    Console.Write(("HZ: Zwei oder mehr 5er oder eine 6. ").PadRight(25));
                    
                    if ((from f in schueler.Fachs
                         where Global.Ungenügend.Contains(f.NoteJetzt)
                         where Global.Mangelhaft.Contains(f.NoteJetzt)
                         select f).Count() > (from f in schueler.Fachs
                                                where Global.Ungenügend.Contains(f.NoteHalbjahr)
                                              where Global.Mangelhaft.Contains(f.NoteHalbjahr)
                                              select f).Count())
                    {
                        Console.Write("Jetzt eine oder mehrere zusätzliche 5 oder 6: Gefährdung ...");
                        schueler.RenderGefährdung();
                        Console.WriteLine("... ok.");
                    }

                    //Abschlussklasse erhalten keine Benachrichtigung
                }
            }
        }

        internal Schuelers FilterDefizitschüler(DefizitäreLeistungen defizitäreLeistungen, Fachs fachs)
        {
            Schuelers schuelersMitDefiziten = new Schuelers();

            foreach (var schueler in this)
            {
                schueler.GetDefizitfächer(defizitäreLeistungen, fachs);
                                
                if (schueler.Fachs.Count > 0)
                {
                    schuelersMitDefiziten.Add(schueler);

                }
            }
            return schuelersMitDefiziten;
        }
    }
}