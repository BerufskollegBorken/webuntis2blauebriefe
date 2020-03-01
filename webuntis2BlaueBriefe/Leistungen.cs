// Published under the terms of GPLv3 Stefan Bäumer 2019.

using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Odbc;
using System.IO;
using System.Linq;
using Microsoft.Office.Interop.Excel;

namespace webuntis2BlaueBriefe
{
    public class Leistungen : List<Leistung>
    {
        public Leistungen(string datei, Fachs fachs)
        {
            using (StreamReader reader = new StreamReader(datei))
            {
                string überschrift = reader.ReadLine();

                Console.Write("Leistungsdaten aus Webuntis ".PadRight(70, '.'));

                while (true)
                {
                    string line = reader.ReadLine();

                    try
                    {
                        if (line != null)
                        {
                            Leistung leistung = new Leistung();
                            var x = line.Split('\t');

                            leistung.Datum = DateTime.ParseExact(x[0], "dd.MM.yyyy", System.Globalization.CultureInfo.InvariantCulture);
                            leistung.Name = x[1];
                            leistung.Klasse = x[2];
                            leistung.Fach = (from f in fachs where f.KürzelUntis.ToString() == x[3] select f.BezeichnungImZeugnis.ToString()).FirstOrDefault();
                            leistung.Prüfungsart = x[4];
                            leistung.Note = x[5].Substring(0, 1);
                            leistung.Bemerkung = x[6];
                            leistung.Benutzer = x[7];
                            leistung.SchlüsselExtern = Convert.ToInt32(x[8]);

                            if (leistung.Prüfungsart.Contains("Blau") && leistung.Note == "1")
                            {
                                this.Add(leistung);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }

                    if (line == null)
                    {
                        break;
                    }
                }
                Console.WriteLine((" " + this.Count.ToString()).PadLeft(30, '.'));
            }
        }

        internal void GetSchuelerMitBlauemBrief(string connectionStringAtlantis, string aktSj)
        {
            Schuelers schuelers = new Schuelers();

            foreach (var person in this)
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
WHERE vorgang_schuljahr = '" + aktSj + @"' AND schue_sj.pu_id = " + person.SchlüsselExtern + ";";

                        OdbcDataReader reader = command.ExecuteReader();

                        int fCount = reader.FieldCount;


                        while (reader.Read())
                        {
                            var idAtlantis = Convert.ToInt32(reader.GetValue(0));
                            var jahrgang = reader.GetValue(1).ToString();
                            string typ = reader.GetValue(2).ToString(); // 0 = Schüler V = Vater  M = Mutter
                            var klasse = reader.GetValue(3).ToString();
                            var vorname = reader.GetValue(5).ToString();
                            var nachname = reader.GetValue(4).ToString();
                            var evorname = reader.GetValue(7).ToString();
                            var enachname = reader.GetValue(6).ToString();
                            var geburtsdatum = DateTime.ParseExact(reader.GetValue(8).ToString(), "dd.MM.yyyy HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
                            var volljaehrig = geburtsdatum.AddYears(18) > DateTime.Now ? false : true;
                            var geschlecht34 = reader.GetValue(9).ToString();
                            var geschlechtMw = geschlecht34 == "3" ? "m" : "w";
                            var strasse = reader.GetValue(10).ToString();
                            var plz = reader.GetValue(11).ToString();
                            var ort = reader.GetValue(12).ToString();
                            var sorgeberechtigtJn = reader.GetValue(13).ToString();
                            var anrede = reader.GetValue(14).ToString();

                            var fachs = (from t in this where t.SchlüsselExtern == idAtlantis select t.Fach).ToList();

                            if (volljaehrig)
                            {
                                Console.WriteLine("Der Schüler " + vorname + " " + nachname + " (Klasse:" + klasse + ") soll gemahnt werden, obwohl er volljährig ist. Der Schüler wird ignoriert.");
                            }
                            else
                            {
                                if (!jahrgang.EndsWith("1"))
                                {
                                    Console.WriteLine("Der Schüler " + vorname + " " + nachname + " (Klasse:" + klasse + ") soll gemahnt werden, obwohl er nicht im ersten Jahrgang ist. Das wird ignoriert.");
                                }
                                else
                                {
                                    schuelers.Add(new Schueler(idAtlantis, typ, klasse, jahrgang, nachname, vorname, evorname, enachname, geburtsdatum, volljaehrig, geschlechtMw, sorgeberechtigtJn, anrede, plz, ort, strasse, fachs));
                                }
                            }
                        }
                        reader.Close();
                        command.Dispose();
                    }
                }
            }
            // Für jeden Schüler ...
            
            try
            {
                foreach (var schueler in (from s in schuelers select s.IdAtlantis).Distinct())
                {
                    // ... für jeden Erziehungsberechtigten dieses Schülers ...

                    int zeile = 1;



                    using (StreamWriter outputFile = new StreamWriter(@"C:\Users\bm\Berufskolleg Borken\Terminplanung - Documents\Blaue Briefe\Steuerdatei.csv"))
                    {

                        outputFile.WriteLine("Anrede,Klasse,Nachname,Vorname,ENachname,EVorname,Strasse,Plz,Ort,SorgeberechtigtJn,Jahrgang,Fach1,Fach2,Fach3,Fach4");

                        foreach (var erziehungsberechtigter in (from e in schuelers where e.IdAtlantis == schueler where e.Typ != "0" select e).ToList())
                            outputFile.WriteLine(
                                erziehungsberechtigter.Anrede + "," +
                                erziehungsberechtigter.Klasse + "," +
                                erziehungsberechtigter.Nachname + "," +
                                erziehungsberechtigter.Vorname + "," +
                                erziehungsberechtigter.ENachname + "," +
                                erziehungsberechtigter.EVorname + "," +
                                erziehungsberechtigter.Strasse + "," +
                                erziehungsberechtigter.Plz + "," +
                                erziehungsberechtigter.Ort + "," +
                                erziehungsberechtigter.SorgeberechtigtJn + "," +
                                erziehungsberechtigter.Jahrgang + "," +
                                RenderFachs(erziehungsberechtigter.Fachs)
                                );
                    }

                    EditorOeffnen(@"C:\Users\bm\Berufskolleg Borken\Terminplanung - Documents\Blaue Briefe\Steuerdatei.csv");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Console.ReadKey();
            }
        }

        private string RenderFachs(List<string> fachs)
        {
            string f = "";

            foreach (var fach in fachs)
            {
                f += fach + ",";
            }

            return f.TrimEnd(',');
        }

        private void EditorOeffnen(string pfad)
        {
            try
            {
                System.Diagnostics.Process.Start(@"C:\Program Files (x86)\Notepad++\Notepad++.exe", pfad);
            }
            catch (Exception)
            {
                System.Diagnostics.Process.Start("Notepad.exe", pfad);
            }
        }
    }
}