﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace webuntis2BlaueBriefe
{
    class Program
    {
        public static string User = System.Security.Principal.WindowsIdentity.GetCurrent().Name.ToUpper().Split('\\')[1];
        public static string Folder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\BlaueBriefe-" + DateTime.Now.ToString("yyyyMMdd-hhmm");

        public static List<string> AktSj = new List<string>
                {
                    (DateTime.Now.Month >= 8 ? DateTime.Now.Year : DateTime.Now.Year - 1).ToString(),
                    (DateTime.Now.Month >= 8 ? DateTime.Now.Year + 1 - 2000 : DateTime.Now.Year - 2000).ToString()
                };

        static void Main(string[] args)
        {
            string steuerdatei = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), DateTime.Now.ToString("yyMMdd-HHmmss") + "_webuntisnoten2atlantis_" + System.Security.Principal.WindowsIdentity.GetCurrent().Name.ToUpper().Split('\\')[1] + ".csv");

            try
            {
                Global.Zeilen = new List<string>() { "AnDieErziehungsberechtigtenVon; anrede; anredeLerncoaching; vorname; nachname; dichSie; plz; straße; ort; klasse; heute; betreff; absatz1; fächer; absatz2; absatz3; klassenleitung; klassenlehrerIn; hinweis"};

                Console.WriteLine(" Webuntis2BlaueBriefe | Published under the terms of GPLv3 | Stefan Bäumer 2022 | Version 20230310");
                Console.WriteLine("====================================================================================================");
                Console.WriteLine("");

                string sourceMarksPerLesson = CheckFile(User, "MarksPerLesson");

                Periodes periodes = new Periodes();
                Leistungen defizitäreWebuntisLeistungen = new Leistungen(sourceMarksPerLesson);
                Fachs fachs = new Fachs(Global.ConnectionStringUntis, defizitäreWebuntisLeistungen);
                Lehrers lehrers = new Lehrers(periodes);
                Klasses klasses = new Klasses(lehrers, periodes, defizitäreWebuntisLeistungen);
                
                Leistungen atlantisLeistungen = new Leistungen(Global.ConnectionStringAtlantis, AktSj, defizitäreWebuntisLeistungen);
                
                Schuelers schuelerMitDefiziten = new Schuelers(defizitäreWebuntisLeistungen, atlantisLeistungen, klasses, lehrers, fachs);

                foreach (var sd in schuelerMitDefiziten)
                {
                    var hzAnzahl5en = (from s in sd.DefizitäreLeistungen where s.NoteHalbjahr == 5 select s).Count();
                    var hzAnzahl6en = (from s in sd.DefizitäreLeistungen where s.NoteHalbjahr == 6 select s).Count();
                    var jetztAnzahl5en = (from s in sd.DefizitäreLeistungen where s.NoteJetzt == 5 select s).Count();
                    var jetztAnzahl6en = (from s in sd.DefizitäreLeistungen where s.NoteJetzt == 6 select s).Count();
                    var nochWeitereDefiziteHinzugekommen = (from s in sd.DefizitäreLeistungen where s.NeueDefizitLeistung select s).Any();
                    var bereitsImHalbjahrGefährdet = (from s in sd.DefizitäreLeistungen where s.NoteHalbjahr >= 5 select s.NoteHalbjahr).Sum() >= 6 ? true : false;
                    var bereitsImHalbjahrEine5 = (from s in sd.DefizitäreLeistungen where s.NoteHalbjahr == 5 select s).Count() == 1 ? true : false;
                    var imHalbjahrKeinDefizit = (from s in sd.DefizitäreLeistungen where s.NoteHalbjahr >=5 select s.NoteHalbjahr).Any() ? false : true;

                    Console.Write(sd.Nachname + "," + sd.Vorname + "," + (sd.Volljaehrig ? " Vollj. " : " Mindj. ") + " (" + sd.Klasse + "):");

                    if (!nochWeitereDefiziteHinzugekommen)
                    {
                        Console.WriteLine("keine weiteren Defizite seit dem Halbjahr, keine Mitteilung.");
                    }

                    // HZ: kein Defizit; 
                    
                    if (imHalbjahrKeinDefizit && nochWeitereDefiziteHinzugekommen)
                    {
                        Console.Write("imHalbjahrKeinDefizit,");
                        
                        //jetzt eine 5: Mitteilung über Leistungsstand

                        if ((from s in sd.DefizitäreLeistungen where s.NeueDefizitLeistung select s.NoteJetzt).Sum() == 5)
                        {
                            Console.Write("jetzt eine 5,");
                                                        
                            sd.RenderMitteilung("M", Folder);
                        }

                        // HZ kein Defizit; jetzt zwei oder mehr 5: Gefährdung
                        // HZ kein Defizit; jetzt eine 6 oder mehr: Gefährdung

                        if ((from s in sd.DefizitäreLeistungen where s.NeueDefizitLeistung select s.NoteJetzt).Sum() >= 6)
                        {
                            Console.Write("jetzt zwei oder mehr 5 oder eine 6,");
                            sd.RenderMitteilung("G", Folder);
                        }
                    }

                    // HZ eine 5; jetzt eine oder mehrere zusätzliche 5en: Gefährdung
                    // HZ eine 5; jetzt eine oder mehrere zusätzliche 6en: Gefährdung

                    if (bereitsImHalbjahrEine5 && nochWeitereDefiziteHinzugekommen)
                    {
                        Console.Write("bereits im Halbjahr eine 5, jetzt eine o. mehrere zusätzliche 5en o. 6en,");
                        sd.RenderMitteilung("G", Folder);
                    }

                    // HZ: Zwei oder mehr 5en oder mindestens eine 6; jetzt eine oder zusätzliche 5en oder 6er: Gefährdung

                    if (bereitsImHalbjahrGefährdet && nochWeitereDefiziteHinzugekommen)
                    {
                        Console.Write("bereits im Halbjahr gefährdet, jetzt eine o. mehrere zusätzliche 5en o. 6en,");
                        sd.RenderMitteilung("G", Folder);
                    } 
                }

                Console.WriteLine("");
                Console.WriteLine("Verarbeitung beendet. ENTER");
                Process.Start(Folder);
                Console.ReadKey();
            }
            catch(IOException ex)
            {
                Console.WriteLine("");
                Console.WriteLine("");
                if (ex.ToString().Contains("bereits vorhanden"))
                {
                    Console.WriteLine("FEHLER: Die Datei existiert bereits. Bitte zuerst löschen. Dann erneut starten.");
                }
                else
                {
                    Console.WriteLine(ex);
                }            
                Console.ReadKey();
                Environment.Exit(0);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Heiliger Bimbam! Es ist etwas schiefgelaufen! Die Verarbeitung wird gestoppt.");
                Console.WriteLine("");
                Console.WriteLine(ex);
                Console.ReadKey();
                Environment.Exit(0);
            }
        }

        private static void CsvDateiExistiert()
        {
            if (!File.Exists(Global.InputNotenCsv))
            {
                RenderNotenexportCsv(Global.InputNotenCsv);
            }
            else
            {
                if (System.IO.File.GetLastWriteTime(Global.InputNotenCsv).Date != DateTime.Now.Date)
                {
                    RenderNotenexportCsv(Global.InputNotenCsv);
                }
            }
        }

        private static void RenderNotenexportCsv(string inputNotenCsv)
        {
            Console.WriteLine("Die Datei " + inputNotenCsv + " existiert nicht.");
            Console.WriteLine("Exportieren Sie die Datei aus dem Digitalen Klassenbuch, indem Sie als Administrator");
            Console.WriteLine(" 1. Klassenbuch > Berichte klicken");
            Console.WriteLine(" 2. Alle Klassen auswählen und als Zeitraum das ganze Schuljahr wählen");
            Console.WriteLine(" 3. Unter \"Noten\" die Prüfungsart \"Alle\" auswählen.");
            Console.WriteLine(" 4. Hinter \"Noten pro Schüler\" auf CSV klicken");
            Console.WriteLine(" 5. Die Datei \"MarksPerLesson.csv\" auf dem Desktop speichern");
            Console.WriteLine("ENTER beendet das Programm");
            Console.ReadKey();
            Environment.Exit(0);
        }

        private static string CheckFile(string user, string kriterium)
        {
            var sourceFile = (from f in Directory.GetFiles(@"c:\users\" + user + @"\Downloads", "*.csv", SearchOption.AllDirectories) where f.Contains(kriterium) orderby File.GetLastWriteTime(f) select f).LastOrDefault();

            if ((sourceFile == null || System.IO.File.GetLastWriteTime(sourceFile).Date != DateTime.Now.Date))
            {
                Console.WriteLine("");
                Console.WriteLine(" Die " + kriterium + "<...>.csv" + (sourceFile == null ? " existiert nicht im Download-Ordner" : " im Download-Ordner ist nicht von heute. \n Es werden keine Daten aus der Datei importiert") + ".");
                Console.WriteLine(" Exportieren Sie die Datei frisch aus Webuntis, indem Sie als Administrator:");

                if (kriterium.Contains("MarksPerLesson"))
                {
                    Console.WriteLine("   1. Klassenbuch > Berichte klicken");
                    Console.WriteLine("   2. Alle Klassen auswählen und ggfs. den Zeitraum einschränken");
                    Console.WriteLine("   3. Unter \"Noten\" die Prüfungsart (-Alle-) auswählen");
                    Console.WriteLine("   4. Unter \"Noten\" den Haken bei Notennamen ausgeben _NICHT_ setzen");
                    Console.WriteLine("   5. Hinter \"Noten pro Schüler\" auf CSV klicken");
                    Console.WriteLine("   6. Die Datei \"MarksPerLesson<...>.CSV\" im Download-Ordner zu speichern");
                    Console.WriteLine(" ");
                    Console.WriteLine(" ENTER beendet das Programm.");
                    Console.ReadKey();
                    Environment.Exit(0);
                }

                if (kriterium.Contains("AbsenceTimesTotal"))
                {
                    Console.WriteLine("   1. Administration > Export klicken");
                    Console.WriteLine("   2. Zeitraum begrenzen, also die Woche der Zeugniskonferenz und vergange Abschnitte herauslassen");
                    Console.WriteLine("   2. Das CSV-Icon hinter Gesamtfehlzeiten klicken");
                    Console.WriteLine("   4. Die Datei \"AbsenceTimesTotal<...>.CSV\" im Download-Ordner zu speichern");
                }
                Console.WriteLine(" ");
                sourceFile = null;
            }

            if (sourceFile != null)
            {
                Console.WriteLine("Ausgewertete Datei: " + (Path.GetFileName(sourceFile) + " ").PadRight(53, '.') + ". Erstell-/Bearbeitungszeitpunkt heute um " + System.IO.File.GetLastWriteTime(sourceFile).ToShortTimeString());
            }

            return sourceFile;
        }
    }
}