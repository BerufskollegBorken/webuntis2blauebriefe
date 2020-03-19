using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;

namespace webuntis2BlaueBriefe
{
    internal class Schueler
    {
        public int IdAtlantis { get; internal set; }
        public string Art { get; internal set; }
        public string Nachname { get; internal set; }
        public string Anrede { get; internal set; }
        public string Vorname { get; internal set; }
        public object Telefons { get; internal set; }
        public string Plz { get; internal set; }
        public string Ort { get; internal set; }
        public string Strasse { get; internal set; }
        public string Email { get; internal set; }
        public string Klasse { get; set; }
        public string Jahrgang { get; set; }
        public DateTime Geburtsdatum { get; set; }
        public bool Volljaehrig { get; set; }
        public string GeschlechtMw { get; set; }
        public Fachs Fachs { get; set; }
        public string Typ { get; set; }       
        public string MAnrede { get; internal set; }
        public string MSorgeberechtigtJn { get; internal set; }
        public string MOrt { get; internal set; }
        public string MPlz { get; internal set; }
        public string MStrasse { get; internal set; }
        public string MNachname { get; internal set; }
        public string MVorname { get; internal set; }
        public string VAnrede { get; internal set; }
        public string VSorgeberechtigtJn { get; internal set; }
        public string VOrt { get; internal set; }
        public string VPlz { get; internal set; }
        public string VStrasse { get; internal set; }
        public string VNachname { get; internal set; }
        public string VVorname { get; internal set; }
        public string Geschlecht { get; internal set; }
        public Sorgeberechtigte Sorgeberechtigte { get; internal set; }
        public string Klassenleitung { get; internal set; }
        public string KlassenleitungMw { get; internal set; }
        public string KlassenleitungMail { get; internal set; }

        internal void RenderMitteilung(string art)
        {
            // Für jede unterschiedliche Adresse

            foreach (var strasse in (from s in this.Sorgeberechtigte select s.Strasse).Distinct().ToList())
            {
                var sorgeberechtigter = (from s in this.Sorgeberechtigte where s.Strasse == strasse select s).FirstOrDefault();

                var origFileName = "Blaue Briefe.docx";

                var fileName = @"c:\\users\\bm\\Desktop\\" + DateTime.Now.ToString("yyyyMMdd") + "-" + Nachname + "-" + Vorname + "-Mitteilung-Leistungsstand" + ".docx";

                System.IO.File.Copy(origFileName.ToString(), fileName.ToString());

                Application wordApp = new Application { Visible = true };
                Document aDoc = wordApp.Documents.Open(fileName, ReadOnly: false, Visible: true);
                aDoc.Activate();
                
                if (Volljaehrig)
                {
                    FindAndReplace(wordApp, "<AnDieErziehungsberechtigtenVon>", "");
                }
                else
                {
                    FindAndReplace(wordApp, "<AnDieErziehungsberechtigtenVon>", "An die Erziehungsberechtigten von");
                }

                FindAndReplace(wordApp, "<anrede>", GetAnrede());
                FindAndReplace(wordApp, "<vorname>", Vorname);
                FindAndReplace(wordApp, "<nachname>", Nachname);
                FindAndReplace(wordApp, "<plz>", sorgeberechtigter.Plz);
                FindAndReplace(wordApp, "<straße>", sorgeberechtigter.Strasse);
                FindAndReplace(wordApp, "<ort>", sorgeberechtigter.Ort);
                FindAndReplace(wordApp, "<klasse>", Klasse);
                FindAndReplace(wordApp, "<heute>", DateTime.Now.ToShortDateString());
                FindAndReplace(wordApp, "<betreff>", art == "G" ? "Gefährdung der Versetzung" : "Mitteilung über den Leistungsstand");
                FindAndReplace(wordApp, "<absatz1>", GetAbsatz1(art));
                FindAndReplace(wordApp, "<fächer>", RenderFächer());
                FindAndReplace(wordApp, "<absatz2>", GetAbsatz2(art));
                FindAndReplace(wordApp, "<absatz3>", GetAbsatz3());
                FindAndReplace(wordApp, "<klassenleitung>", Klassenleitung);
                FindAndReplace(wordApp, "<klassenlehrerIn>", KlassenleitungMw == "Herr" ? "Klassenlehrer" : "Klassenlehrerin");
                
                FindAndReplace(wordApp, "<hinweis>", GetHinweis());
                aDoc.Save();
                aDoc.Close();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(aDoc);
                aDoc = null;
                GC.Collect();
            }
        }

        private object GetAbsatz3()
        {
            return "Wir laden Sie zu einem Beratungsgespräch ein. Stimmen Sie bitte den Gesprächstermin mit " + (KlassenleitungMw == "Herr" ? "dem Klassenlehrer " : "der Klassenlehrerin") + " " + Klassenleitung + " (" + KlassenleitungMail + ") ab.";
        }

        private object GetAbsatz2(string art)
        {
            if (art == "G")
            {
                return "abweichend von " + (this.Fachs.Count() > 1 ? "den" : "der") + " im letzten Zeugnis erteilten Note" + (this.Fachs.Count() > 1 ? "n" : "") + " nicht mehr " + (this.Fachs.Count() > 1 ? "ausreichen" : "ausreicht") + ".";
            }
            else
            {
                return "abweichend von " + (this.Fachs.Count() > 1 ? "den" : "der") + " im letzten Zeugnis erteilten Note" + (this.Fachs.Count() > 1 ? "n" : "") + " nicht mehr " + (this.Fachs.Count() > 1 ? "ausreichen" : "ausreicht") + ". Stellt sich eine weitere nicht ausreichende Leistung ein, ist die Versetzung gefährdet.";
            }            
        }

        private object GetAbsatz1(string art)
        {
            if (!Volljaehrig)
            {
                if (Geschlecht.ToLower() == "m")
                {
                    return "Sie werden darüber unterrichtet, dass die Leistung" + (this.Fachs.Count() > 1 ? "en" : "") + " Ihres Sohnes " + Vorname + ", Klasse " + Klasse + ", in " + (this.Fachs.Count() > 1 ? "den Fächern" : "dem Fach");
                }
                else
                {
                    return "Sie werden darüber unterrichtet, dass die Leistung" + (this.Fachs.Count() > 1 ? "en" : "") + " Ihrer Tochter " + Vorname + ", Klasse " + Klasse + ", in " + (this.Fachs.Count() > 1 ? "den Fächern" : "dem Fach");
                }
            }
            else
            {
                return "Sie werden darüber unterrichtet, dass Ihre Leistung" + (this.Fachs.Count() > 1 ? "en" : "") + " in " + (this.Fachs.Count() > 1 ? "den Fächern" : "dem Fach");
            }
        }

        private object GetIhreTochterIhrSohn()
        {
            if (!Volljaehrig)
            {
                if (Geschlecht.ToLower() == "m")
                {
                    return "die Leistung Ihres Sohnes " + Vorname + ", Klasse " + Klasse + ", ";
                }
                else
                {
                    return "die Leistung Ihrer Tochter " + Vorname + ", Klasse " + Klasse + ", ";
                }
            }
            else
            {
                return @"Ihre 
Leistung";
            }
        }

        private object GetHinweis()
        {
            if (!Volljaehrig)
            {
                if (Geschlecht.ToLower() == "m")
                {
                    return "Ihr Sohn die Klasse zurzeit wiederholt,";
                }
                else
                {
                    return "Ihre Tochter die Klasse zurzeit wiederholt,";
                }
            }
            else
            {
                return "Sie die Klasse zurzeit wiederholen,";
            }
        }

        private object GetAnrede()
        {
            if (Volljaehrig)
            {
                if (Geschlecht.ToLower() == "m")
                {
                    return "Sehr geehrter Herr " + Vorname + " " + Nachname + ",";
                }
                else
                {
                    return "Sehr geehrte Frau " + Vorname + " " + Nachname + ",";
                }
            }
            else
            {
                return "Sehr geehrte Erziehungsberechtigte,";
            }
        }

        private string RenderFächer()
        {
            string x = "";

            foreach (var fach in Fachs)
            {
                x += " " + fach.BezeichnungImZeugnis + " (" + (Global.Mangelhaft.Contains(fach.NoteJetzt) ? "mangelhaft":"") + (Global.Ungenügend.Contains(fach.NoteJetzt) ? "ungenügend" : "") + ")," ;
            }
            return x.TrimEnd(',');
        }

        internal void RenderGefährdung()
        {
            //System.IO.File.Copy(origFileName.ToString(), fileName.ToString());

            //Application wordApp = new Microsoft.Office.Interop.Word.Application { Visible = true };
            //Document aDoc = wordApp.Documents.Open(fileName, ReadOnly: false, Visible: true);
            //aDoc.Activate();

            //FindAndReplace(wordApp, "<vorname>", Vorname);
            //FindAndReplace(wordApp, "<nachname>", Nachname);
            //FindAndReplace(wordApp, "<plz>", Adresse.Plz);
            //FindAndReplace(wordApp, "<straße>", Adresse.Strasse);
            //FindAndReplace(wordApp, "<ort>", Adresse.Ort);
            //FindAndReplace(wordApp, "<klasse>", Klasse.NameUntis);
            //FindAndReplace(wordApp, "<klassenleitung>", Klasse.Klassenleitungen[0].Anrede + " " + Klasse.Klassenleitungen[0].Nachname);
            //FindAndReplace(wordApp, "<mahnung>", RenderBisherigeMaßnahmen());
            //FindAndReplace(wordApp, "<heute>", DateTime.Now.ToShortDateString());

            //for (int i = 0; i < AbwesenheitenSeitLetzterMaßnahme.Count; i++)
            //{
            //    string fehltage = AbwesenheitenSeitLetzterMaßnahme[i].Datum.ToShortDateString() + " (" + AbwesenheitenSeitLetzterMaßnahme[i].Fehlstunden + "), " + "<fehltage>";
            //    FindAndReplace(wordApp, "<fehltage>", fehltage.TrimEnd(','));
            //}

            //FindAndReplace(wordApp, ", <fehltage>", "");

            //aDoc.Save();
            //aDoc.Close();
            //System.Runtime.InteropServices.Marshal.ReleaseComObject(aDoc);
            //aDoc = null;
            //GC.Collect();

            //return fileName;
        }

        private static void FindAndReplace(Application doc, object findText, object replaceWithText)
        {
            //options
            object matchCase = false;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundsLike = false;
            object matchAllWordForms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiacritics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;
            //execute find and replace
            try
            {
                doc.Selection.Find.Execute(ref findText, ref matchCase, ref matchWholeWord,
                ref matchWildCards, ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format, ref replaceWithText, ref replace,
                ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Console.ReadKey();
            }
        }

        internal void GetDefizitfächer(DefizitäreLeistungen defizitäreLeistungen, Fachs fachs)
        {
            Fachs fachss = new Fachs();

            // Suche alle defizitären Fächer dieses Schülers

            var defizitäreFächerDiesesSchülers = (from d in defizitäreLeistungen
                                                  where d.SchlüsselExtern == IdAtlantis
                                                  where d.Prüfungsart.Contains("laue")
                                                  select d.Fach.KürzelUntis).Distinct().ToList();

            var noteJetzt = "";
            var noteHalbjahr = "";

            foreach (var dFach in defizitäreFächerDiesesSchülers)
            {
                foreach (var d in defizitäreLeistungen)
                {
                    if (d.SchlüsselExtern == IdAtlantis)
                    {
                        if (d.Fach.KürzelUntis == dFach)
                        {
                            if (d.Prüfungsart.Contains("laue"))
                            {
                                noteJetzt = d.Note;
                            }
                            if (d.Prüfungsart.Contains("albjahres"))
                            {
                                noteHalbjahr = d.Note;
                            }
                        }
                    }
                }

                // Nur Fächer mit Defizit werden gesetzt

                if (noteJetzt != null)
                {
                    this.Fachs.Add(new Fach(dFach, (from f in fachs where f.KürzelUntis == dFach select f.BezeichnungImZeugnis).FirstOrDefault(), noteJetzt,noteHalbjahr));
                }
            }
        }
    }
}