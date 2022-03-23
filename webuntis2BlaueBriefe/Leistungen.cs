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
        public DefizitäreLeistungen(Fachs fachs, Klasses klasses)
        {
            using (StreamReader reader = new StreamReader(Global.InputNotenCsv))
            {
                string überschrift = reader.ReadLine();

                int i = 1;

                Leistung leistung = new Leistung();

                while (true)
                {
                    string line = reader.ReadLine();

                    try
                    {
                        if (line != null)
                        {                            
                            var x = line.Split('\t');
                            i++;

                            if (i==2629)
                            {
                                string a = "";
                            }
                            if (x.Length == 10)
                            {
                                leistung = new Leistung();
                                leistung.Datum = DateTime.ParseExact(x[0], "dd.MM.yyyy", System.Globalization.CultureInfo.InvariantCulture);
                                leistung.Name = x[1];
                                leistung.Klasse = x[2];
                                leistung.Fach = (from f in fachs where f.KürzelUntis.ToString() == x[3] select f).FirstOrDefault();
                                leistung.Prüfungsart = x[4];
                                leistung.BlauerBriefNote = x[5];
                                leistung.Halbjahresgesamtnote = x[9];
                                leistung.Bemerkung = x[6];
                                leistung.Benutzer = x[7];
                                leistung.SchlüsselExtern = Convert.ToInt32(x[8]);
                            }

                            // Wenn in den Bemerkungen eine zusätzlicher Umbruch eingebaut wurde:

                            if (x.Length == 7)
                            {
                                leistung = new Leistung();
                                leistung.Datum = DateTime.ParseExact(x[0], "dd.MM.yyyy", System.Globalization.CultureInfo.InvariantCulture);
                                leistung.Name = x[1];
                                leistung.Klasse = x[2];
                                leistung.Fach = (from f in fachs where f.KürzelUntis.ToString() == x[3] select f).FirstOrDefault();
                                leistung.BlauerBriefNote = x[5];
                                leistung.Bemerkung = x[6];
                                Console.WriteLine("\n\n  [!] Achtung: In den Zeilen " + (i - 1) + "-" + i + " hat vermutlich die Lehrkraft eine Bemerkung mit einem Zeilen-");
                                Console.Write("      umbruch eingebaut. Es wird nun versucht trotzdem korrekt zu importieren ... ");
                            }

                            if (x.Length == 4)
                            {
                                leistung.Benutzer = x[1];
                                leistung.SchlüsselExtern = Convert.ToInt32(x[2]);
                                leistung.Halbjahresgesamtnote = x[3];
                                Console.WriteLine("hat geklappt.\n");                                
                            }

                            if (x.Length < 4)
                            {
                                Console.WriteLine("\n\n[!] MarksPerLesson.CSV: In der Zeile " + i + " stimmt die Anzahl der Spalten nicht. Das kann passieren, wenn z. B. die Lehrkraft bei einer Bemerkung einen Umbruch eingibt. Mit Suchen & Ersetzen kann die Datei MarksPerLesson.CSV korrigiert werden.");
                                Console.ReadKey();                                
                                throw new Exception("\n\n[!] MarksPerLesson.CSV: In der Zeile " + i + " stimmt die Anzahl der Spalten nicht. Das kann passieren, wenn z. B. die Lehrkraft bei einer Bemerkung einen Umbruch eingibt. Mit Suchen & Ersetzen kann die Datei MarksPerLesson.CSV korrigiert werden.");
                            }

                            // Nur Halbjahresnoten und Blaue Briefe sind relevant. Differenzierungsbereich zählt nicht.

                            if (Global.Mangelhaft.Contains(leistung.BlauerBriefNote) || Global.Ungenügend.Contains(leistung.BlauerBriefNote))
                            {
                                if (leistung.Prüfungsart == Global.BlaueBriefe)
                                {
                                    if (leistung.IstKeinDiff(klasses))
                                    {
                                        this.Add(leistung);
                                    }
                                    else
                                    {
                                        Console.WriteLine("ACHTUNG: Mahnung im Diff-Bereich. " + leistung.Klasse + ": " + leistung.Fach.BezeichnungImZeugnis + " [ENTER]");
                                        Console.ReadKey();
                                    }                               
                                }                               
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