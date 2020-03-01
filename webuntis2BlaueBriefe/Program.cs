using System;
using System.Collections.Generic;
using System.IO;

namespace webuntis2BlaueBriefe
{
    class Program
    {
        public const string ConnectionStringAtlantis = @"Dsn=Atlantis9;uid=DBA";
        public const string ConnectionStringUntis = @"Provider = Microsoft.Jet.OLEDB.4.0; Data Source=M:\\Data\\gpUntis.mdb;";

        static void Main(string[] args)
        {
            try
            {
                string inputNotenCsv = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\MarksPerLesson.csv";
               
                List<string> aktSj = new List<string>
                {
                    (DateTime.Now.Month >= 8 ? DateTime.Now.Year : DateTime.Now.Year - 1).ToString(),
                    (DateTime.Now.Month >= 8 ? DateTime.Now.Year + 1 - 2000 : DateTime.Now.Year - 2000).ToString()
                };

                Console.WriteLine(" Webuntis2BlaueBriefe | Published under the terms of GPLv3 | Stefan Bäumer 2020 | Version 202000301");
                Console.WriteLine("====================================================================================================");
                Console.WriteLine("");

                int sj = (DateTime.Now.Month >= 8 ? DateTime.Now.Year : DateTime.Now.Year - 1);
                string aktSjUntis = sj.ToString() + (sj + 1);

                Fachs fachs = new Fachs(aktSjUntis, ConnectionStringUntis);

                if (!File.Exists(inputNotenCsv))
                {
                    RenderNotenexportCsv(inputNotenCsv);
                }
                else
                {
                    if (System.IO.File.GetLastWriteTime(inputNotenCsv).Date != DateTime.Now.Date)
                    {
                        RenderNotenexportCsv(inputNotenCsv);
                    }
                }
                                
                Console.WriteLine("");                
                Leistungen alleWebuntisLeistungen = new Leistungen(inputNotenCsv, fachs);
                
                alleWebuntisLeistungen.GetSchuelerMitBlauemBrief(ConnectionStringAtlantis, aktSj[0] + "/" + aktSj[1]);                
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
        private static void RenderNotenexportCsv(string inputNotenCsv)
        {
            Console.WriteLine("Die Datei " + inputNotenCsv + " existiert nicht.");
            Console.WriteLine("Exportieren Sie die Datei aus dem Digitalen Klassenbuch, indem Sie");
            Console.WriteLine(" 1. Klassenbuch > Berichte klicken");
            Console.WriteLine(" 2. Alle Klassen auswählen");
            Console.WriteLine(" 3. Unter \"Noten\" die Prüfungsart (\"Blauer Brief\") auswählen");
            Console.WriteLine(" 4. Hinter \"Noten pro Schüler\" auf CSV klicken.");
            Console.WriteLine(" 5. Die Datei \"MarksPerLesson.csv\" auf dem Desktop speichern.");
            Console.WriteLine("ENTER beendet das Programm.");
            Console.ReadKey();
            Environment.Exit(0);
        }
    }
}