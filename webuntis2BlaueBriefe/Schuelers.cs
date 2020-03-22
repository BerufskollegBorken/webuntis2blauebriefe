using Microsoft.Exchange.WebServices.Data;
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

        public Schuelers(DefizitäreLeistungen defizitäreLeistungen, Klasses klasses, Lehrers lehrers, Fachs fachs)
        {
            foreach (var idAtlantis in (from t in defizitäreLeistungen select t.SchlüsselExtern).Distinct().ToList())
            {
                using (OdbcConnection connection = new OdbcConnection(Global.ConnectionStringAtlantis))
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
WHERE vorgang_schuljahr = '" + Global.AktSjAtlantis + @"' AND schue_sj.pu_id = " + idAtlantis + ";";

                        OdbcDataReader reader = command.ExecuteReader();

                        int fCount = reader.FieldCount;

                        var x = (from s in this where s.IdAtlantis == idAtlantis select s).FirstOrDefault();

                        Schueler schueler = new Schueler()
                        {
                            Sorgeberechtigte = new Sorgeberechtigte()
                        };

                        while (reader.Read())
                        {
                            schueler.IdAtlantis = Convert.ToInt32(reader.GetValue(0));
                            var sorgeberechtigtJn = reader.GetValue(13).ToString();
                            var anrede = reader.GetValue(14).ToString();
                            var typ = reader.GetValue(2).ToString();
                            
                            if (sorgeberechtigtJn == "N" && typ == "0")
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

                        schueler.GetDefizitfächer(defizitäreLeistungen, fachs);
                        schueler.Dateien = new List<string>();
                        if (schueler.Fachs.Count > 0)
                        {
                            this.Add(schueler);
                        }
                        reader.Close();
                        command.Dispose();                        
                    }
                }                
            }
            Console.WriteLine(("Schüler mit Defiziten " + ".".PadRight(this.Count / 150, '.')).PadRight(48, '.') + (" " + this.Count).ToString().PadLeft(4), '.');
        }

        internal void MailAnKlassenlehrer()
        {
            foreach (var klasse in (from s in this select s.Klasse).Distinct())
            {
                var schülerDieserKlasse = (from s in this
                                           where s.Klasse == klasse
                                           select s).OrderBy(x => x.Nachname).ThenBy(x => x.Vorname).ToList();

                Mail(schülerDieserKlasse);
            }
        }

        private void Mail(List<Schueler> schülerDieserKlasse)
        {
            ExchangeService exchangeService = new ExchangeService()
            {
                UseDefaultCredentials = true,
                TraceEnabled = false,
                TraceFlags = TraceFlags.All,
                Url = new Uri("https://ex01.bkb.local/EWS/Exchange.asmx")
            };

            EmailMessage message = new EmailMessage(exchangeService);

            message.ToRecipients.Add("stefan.baeumer@berufskolleg-borken.de");
            message.BccRecipients.Add(schülerDieserKlasse[0].KlassenleitungMail);

            foreach (var s in schülerDieserKlasse)
            {
                foreach (var datei in s.Dateien)
                {
                    message.Attachments.AddFileAttachment(datei);
                }
            }
                 
            message.Subject = "Blaue Briefe - BITTE KONTROLLIEREN";

            message.Body = @"Guten Tag " + schülerDieserKlasse[0].Klassenleitung + "," +
                "<br><br>Sie erhalten diese Mail in Ihrer Eigenschaft als Klassenleitung der Klasse " + schülerDieserKlasse[0].Klasse + "." +
                "<br><br>" +
                "Bitte prüfen Sie die im Folgenden automatisch erstellten und aufgelisteten Blauen Briefe gewissenhaft. Die Verantwortung für die Richtigkeit liegt ganz allein bei Ihnen. Achten Sie darauf, dass kein Fach des Differenzierungsbereichs angemahnt wird.</br>" +
                "<br><table border = 1><tr><td>Name</td><td>Vollj.</td><td>Halbjahreszeugnis</td><td>Aktueller Notenstand<br> aller abweichenden <br>Fächer</td><td>Gefährdung / Mitteilung Leistungsstand</td><td>Anschrift(en)</td></tr>";

            foreach (var s in schülerDieserKlasse)
            {
                message.Body += "<tr>" + s.Protokoll + "</tr>";
            }

            message.Body += "</table></br>" +
                "Wenn Sie sich nicht zeitnah zurückmelden, bestätigen damit die Richtigkeit. Die Briefe werden dann alsbald verschickt." +
                "</br></br>" +
                "Stefan Bäumer" +
                "";

            message.Save(WellKnownFolderName.Drafts);
            //message.SendAndSaveCopy();
            
            Console.WriteLine(schülerDieserKlasse[0].Klasse + " " + schülerDieserKlasse[0].Klassenleitung  + ": Mail gesendet.");
        }

        internal void RenderBriefe()
        {
            Console.WriteLine(this[0].Klasse + "\n" + "=".PadRight(this[0].Klasse.Length - 1, '='));

            foreach (var schueler in this)
            {
                schueler.RenderBrief();
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