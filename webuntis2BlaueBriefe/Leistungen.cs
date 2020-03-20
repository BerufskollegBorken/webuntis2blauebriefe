// Published under the terms of GPLv3 Stefan Bäumer 2019.

using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Odbc;
using System.IO;
using System.Linq;

namespace webuntis2BlaueBriefe
{
    public class DefizitäreLeistungen : List<Leistung>
    {
        public DefizitäreLeistungen(Fachs fachs, Stundentafels stundentafels)
        {
            using (StreamReader reader = new StreamReader(Global.InputNotenCsv))
            {
                string überschrift = reader.ReadLine();
                
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
                            leistung.Fach = (from f in fachs where f.KürzelUntis.ToString() == x[3] select f).FirstOrDefault();
                            leistung.Prüfungsart = x[4];
                            leistung.Note = x[5];
                            leistung.Bemerkung = x[6];
                            leistung.Benutzer = x[7];
                            leistung.SchlüsselExtern = Convert.ToInt32(x[8]);

                            // Nur Halbjahresnoten und Blaue Briefe sind relevant. Differenzierungsbereich zählt nicht.
                            if (Global.Mangelhaft.Contains(leistung.Note) || Global.Ungenügend.Contains(leistung.Note))
                            {
                                if (leistung.Prüfungsart == Global.BlaueBriefe)
                                {
                                    if (leistung.IstKeinDiff(stundentafels))
                                    {
                                        this.Add(leistung);
                                    }
                                    else
                                    {
                                        Console.WriteLine("ACHTUNG: Mahnung im Diff-Bereich." + leistung.Klasse + " " + leistung.Fach);
                                        Console.ReadKey();
                                    }                               
                                }                               
                            }
                            if (leistung.Prüfungsart == Global.Halbjahreszeugnis)
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
                Console.WriteLine(("Leistungsdaten " + ".".PadRight(this.Count / 150, '.')).PadRight(48, '.') + (" " + this.Count).ToString().PadLeft(4), '.');
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