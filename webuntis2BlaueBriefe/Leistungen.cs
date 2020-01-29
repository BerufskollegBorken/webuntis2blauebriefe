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
        public Leistungen(string datei)
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
                            leistung.Fach = x[3];
                            leistung.Prüfungsart = x[4];
                            leistung.Note = x[5].Substring(0, 1);
                            leistung.Bemerkung = x[6];
                            leistung.Benutzer = x[7];
                            leistung.SchlüsselExtern = Convert.ToInt32(x[8]);
                            this.Add(leistung);
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
                    DataSet dataSet = new DataSet();
                    OdbcDataAdapter schuelerAdapter = new OdbcDataAdapter(@"SELECT 
DBA.schue_sj.pu_id as ID,
DBA.schue_sj.kl_id, 
DBA.schue_sj.s_jahrgang AS Jahrgang,
DBA.klasse.klasse as Klasse,
DBA.schueler.name_1 as Vorname,
DBA.schueler.name_2 as Nachname,
DBA.schueler.anrede_text,
DBA.schueler.dat_geburt as Geburtsdatum,
DBA.schueler.s_geschl as Geschlecht
DBA.schueler.s_erzb_1_art,
DBA.schueler.s_erzb_2_art,
DBA.schueler.id_hauptadresse,
DBA.adresse.hauptadresse_jn,
DBA.adresse.s_anrede,
DBA.adresse.anrede_text,
DBA.adresse.name_1 AS EVorname,
DBA.adresse.name_2 AS ENachname,
DBA.adresse.name_3,
DBA.adresse.strasse AS Strasse,
DBA.adresse.plz AS Plz,
DBA.adresse.ort AS Ort,
DBA.adresse.s_typ_adr as Typ,
DBA.adresse.plz_postfach as Postfach,
DBA.adresse.postfach as Postfach,
DBA.adresse.s_titel_ad,
DBA.adresse.sorge_berechtigt_jn,
DBA.adresse.s_sorgerecht,
DBA.adresse.brief_adresse,
DBA.adresse.s_famstand_adr
FROM((DBA.schue_sj JOIN DBA.klasse ON DBA.schue_sj.kl_id = DBA.klasse.kl_id) JOIN DBA.schueler ON DBA.schue_sj.pu_id = DBA.schueler.pu_id) JOIN DBA.adresse ON DBA.schueler.pu_id = DBA.adresse.pu_id
WHERE vorgang_schuljahr = '" + aktSj + @"' AND schue_sj.pu_id = " + person.SchlüsselExtern + ", connection);", connectionStringAtlantis);
                    connection.Open();
                    schuelerAdapter.Fill(dataSet, "DBA.klasse");

                    foreach (DataRow theRow in dataSet.Tables["DBA.klasse"].Rows)
                    {
                        var idAtlantis = theRow["IdAtlantis"] == null ? -99 : Convert.ToInt32(theRow["IdAtlantis"]);
                        string typ = theRow["Typ"] == null ? "" : theRow["Typ"].ToString(); // 0 = Schüler V = Vater  M = Mutter
                        var klasse = theRow["Klasse"] == null ? "" : theRow["Klasse"].ToString();
                        var jahrgang = theRow["Jahrgang"] == null ? "" : theRow["Jahrgang"].ToString();
                        var nachname = theRow["Nachname"] == null || theRow["Nachname"].ToString() == "" ? "NN" : theRow["Nachname"].ToString();
                        var vorname = theRow["Vorname"] == null ? "" : theRow["Vorname"].ToString();
                        var enachname = theRow["ENachname"] == null || theRow["ENachname"].ToString() == "" ? "NN" : theRow["Nachname"].ToString();
                        var evorname = theRow["EVorname"] == null ? "" : theRow["EVorname"].ToString();
                        var geburtsdatum = theRow["Geburtsdatum"].ToString().Length < 3 ? new DateTime() : DateTime.ParseExact(theRow["Geburtsdatum"].ToString(), "dd.MM.yyyy HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
                        var volljaehrig = geburtsdatum.AddYears(18) > DateTime.Now ? false : true;
                        var geschlecht34 = theRow["Geschlecht"] == null ? "" : theRow["Geschlecht"].ToString();
                        var geschlechtMw = geschlecht34 == "3" ? "m" : "w";
                        var sorgeberechtigtJn = theRow["SorgeberechtigtJn"] == null ? "" : theRow["SorgeberechtigtJn"].ToString();
                        var anrede = theRow["Anrede"] == null ? "" : theRow["Anrede"].ToString();
                        var plz = theRow["plz"] == null ? "" : theRow["plz"].ToString();
                        var ort = theRow["ort"] == null ? "" : theRow["ort"].ToString();
                        var strasse = theRow["strasse"] == null ? "" : theRow["strasse"].ToString();
                        var fachs = (from t in this where t.SchlüsselExtern == idAtlantis select t.Fach).ToList();
                        if (typ == "0")
                        {
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
                    }
                    connection.Close();
                }            
            }

            // Für jeden Schüler ...

            foreach (var schueler in (from s in schuelers select s.IdAtlantis).Distinct())
            {
                // ... für jeden Erziehungsberechtigten dieses Schülers ...

                Application excel = new Application();
                Workbook workbook = excel.Workbooks.Open(@"C:\Users\bm\Berufskolleg Borken\Terminplanung - Documents\Blaue Briefe\Steuerdatei.xlsx");
                Worksheet worksheet = (Worksheet)workbook.Worksheets.get_Item(1);

                int zeile = 0;

                foreach (var erziehungsberechtigter in (from e in schuelers where e.IdAtlantis == schueler where e.Typ != "0" where e.SorgeberechtigtJn == "J" select e))
                {
                    worksheet.Cells[zeile, 1] = erziehungsberechtigter.IdAtlantis;
                    worksheet.Cells[zeile, 2] = erziehungsberechtigter.Vorname;
                    worksheet.Cells[zeile, 3] = erziehungsberechtigter.Nachname;
                    worksheet.Cells[zeile, 4] = erziehungsberechtigter.Strasse;
                    worksheet.Cells[zeile, 5] = erziehungsberechtigter.Plz;
                    worksheet.Cells[zeile, 6] = erziehungsberechtigter.Ort;
                    worksheet.Cells[zeile, 7] = erziehungsberechtigter.EVorname;
                    worksheet.Cells[zeile, 8] = erziehungsberechtigter.ENachname;

                    for (int i = 0; i < erziehungsberechtigter.Fachs.Count; i++)
                    {
                        worksheet.Cells[zeile, 9 + i] = erziehungsberechtigter.Fachs[i];
                    }

                    zeile++;
                }
                workbook.SaveAs(@"C:\Users\bm\Berufskolleg Borken\Terminplanung - Documents\Blaue Briefe\Steuerdatei.xlsx");

                workbook.Close();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                excel.Quit();
            }
        }
    }
}
                