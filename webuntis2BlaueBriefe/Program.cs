using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace webuntis2BlaueBriefe
{
    class Program
    {        
        static void Main(string[] args)
        {
            System.Net.ServicePointManager.ServerCertificateValidationCallback = ((sender, certificate, chain, sslPolicyErrors) => true);

            string steuerdatei = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), DateTime.Now.ToString("yyMMdd-HHmmss") + "_webuntisnoten2atlantis_" + System.Security.Principal.WindowsIdentity.GetCurrent().Name.ToUpper().Split('\\')[1] + ".csv");

            try
            {
                Global.Zeilen = new List<string>() { "AnDieErziehungsberechtigtenVon; anrede; anredeLerncoaching; vorname; nachname; dichSie; plz; straße; ort; klasse; heute; betreff; absatz1; fächer; absatz2; absatz3; klassenleitung; klassenlehrerIn; hinweis"};

                Console.WriteLine(" Webuntis2BlaueBriefe | Published under the terms of GPLv3 | Stefan Bäumer 2020 | Version 202000322");
                Console.WriteLine("====================================================================================================");
                Console.WriteLine("");

                CsvDateiExistiert();
                
                Fachs fachs = new Fachs(Global.ConnectionStringUntis);
                Stundentafels stundentafels = new Stundentafels(fachs);
                Periodes periodes = new Periodes();
                Lehrers lehrers = new Lehrers(periodes);
                Klasses klasses = new Klasses(lehrers, periodes);                
                DefizitäreLeistungen defizitäreLeistungen = new DefizitäreLeistungen(fachs,stundentafels);
                Schuelers schuelerMitDefiziten = new Schuelers(defizitäreLeistungen, klasses, lehrers, fachs);
                
                schuelerMitDefiziten.RenderBriefeUndSteuerdatei(steuerdatei);

                schuelerMitDefiziten.MailAnKlassenlehrer();
                Console.WriteLine("Verarbeitung beendet. ENTER");
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
            Console.WriteLine("Exportieren Sie die Datei aus dem Digitalen Klassenbuch, indem Sie");
            Console.WriteLine(" 1. Klassenbuch > Berichte klicken");
            Console.WriteLine(" 2. Alle Klassen auswählen");
            Console.WriteLine(" 3. Unter \"Noten\" die Prüfungsart alle Prüfungsarten auswählen");
            Console.WriteLine(" 4. Hinter \"Noten pro Schüler\" auf CSV klicken.");
            Console.WriteLine(" 5. Die Datei \"MarksPerLesson.csv\" auf dem Desktop speichern.");
            Console.WriteLine("ENTER beendet das Programm.");
            Console.ReadKey();
            Environment.Exit(0);
        }
    }
}