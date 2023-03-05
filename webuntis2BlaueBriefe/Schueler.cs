using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
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
        public Leistungen DefizitäreLeistungen { get; set; }
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
        public string Protokoll { get; private set; }
        public List<string> Dateien { get; set; }

        internal void RenderMitteilung(string art, string folder)
        {
            System.IO.Directory.CreateDirectory(folder);

            // Für jede unterschiedliche Adresse
            
            var x = (from s in this.Sorgeberechtigte select s.Strasse).Distinct().Count();

            var sss = (from s in this.Sorgeberechtigte select s.Strasse).Distinct().ToList();

            if (Volljaehrig)
            {
                sss = new List<string>() { Strasse };
            }

            foreach (var strasse in sss)
            {
                var sorgeberechtigter = (from s in this.Sorgeberechtigte where s.Strasse == strasse select s).FirstOrDefault();

                var origFileName = "Blaue Briefe.docx";
                
                var fileName = folder + "\\" + (Volljaehrig?"V-":"M-") + DateTime.Now.ToString("yyyyMMdd") + "-" + Klasse + "-" + Nachname + "-" + Vorname + (x > 1 ? strasse : "") + (art == "G" ? "-Gefährdung.docx" : "-Mitteilung.docx");

                Dateien.Add(fileName);

                if (File.Exists(fileName));
                {
                    File.Delete(fileName);
                }

                System.IO.File.Copy(origFileName.ToString(), fileName.ToString());

                object oMissing = System.Reflection.Missing.Value;

                Application wordApp = new Application { Visible = true };
                Document doc = wordApp.Documents.Open(fileName, ReadOnly: false, Visible: true);
                doc.Activate();

                if (Volljaehrig)
                {
                    FindAndReplace(wordApp, doc, "<AnDieErziehungsberechtigtenVon>", "");                    
                }
                else
                {
                    FindAndReplace(wordApp, doc, "<AnDieErziehungsberechtigtenVon>", "An die Erziehungsberechtigten von");
                }

                FindAndReplace(wordApp, doc,"<anrede>", GetAnrede());
                FindAndReplace(wordApp, doc,"<anredeLerncoaching>", GetAnredeLerncoaching());
                FindAndReplace(wordApp, doc,"<vorname>", Vorname);
                FindAndReplace(wordApp, doc,"<nachname>", Nachname);
                FindAndReplace(wordApp, doc,"<dichSie>", Volljaehrig ? "Sie" : "Dich");
                
                if (!Volljaehrig)
                {
                    FindAndReplace(wordApp, doc,"<plz>", sorgeberechtigter.Plz);
                    FindAndReplace(wordApp, doc,"<straße>", sorgeberechtigter.Strasse);
                    FindAndReplace(wordApp, doc,"<ort>", sorgeberechtigter.Ort);                    
                }
                else
                {
                    FindAndReplace(wordApp, doc,"<plz>", "");
                    FindAndReplace(wordApp, doc,"<straße>", "!!! Kein Briefversand bei Volljährigen !!!");
                    FindAndReplace(wordApp, doc,"<ort>", "");                    
                }
                FindAndReplace(wordApp, doc,"<klasse>", Klasse);
                FindAndReplace(wordApp, doc,"<heute>", DateTime.Now.ToShortDateString());
                FindAndReplace(wordApp, doc,"<betreff>", art == "M" ? "Mitteilung über den Leistungsstand" : "Gefährdung der Versetzung");
                FindAndReplace(wordApp, doc,"<absatz1>", GetAbsatz1(art));
                FindAndReplace(wordApp, doc,"<fächer>", RenderFächer(art));
                FindAndReplace(wordApp, doc,"<absatz2>", GetAbsatz2(art));
                FindAndReplace(wordApp, doc,"<absatz3>", GetAbsatz3());
                FindAndReplace(wordApp, doc,"<klassenleitung>", Klassenleitung);
                FindAndReplace(wordApp, doc,"<klassenlehrerIn>", KlassenleitungMw == "Herr" ? "Klassenlehrer" : "Klassenlehrerin");
                FindAndReplace(wordApp, doc,"<hinweis>", GetHinweis());
                FindAndReplace(wordApp, doc,"<footer>", "");
                
                doc.ExportAsFixedFormat(fileName+".pdf", WdExportFormat.wdExportFormatPDF, false, WdExportOptimizeFor.wdExportOptimizeForOnScreen,
                    WdExportRange.wdExportAllDocument, 1, 1, WdExportItem.wdExportDocumentContent, true, true,
                    WdExportCreateBookmarks.wdExportCreateHeadingBookmarks, true, true, false, ref oMissing);
                doc.Save();                             
                doc.Close();            
                System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);
                doc = null;
                GC.Collect();
                wordApp.Quit();
            }
            if (art == "M")
            {
                Console.WriteLine("Mitteilung über den Leistungsstand");
            }
            else
            {
                Console.WriteLine("Gefährdung");
                
            }
        }

        private object GetAnredeLerncoaching()
        {
            string x = "";
            
            x += "Liebe" + (GeschlechtMw == "M" ? "r " : " ") + Vorname;

            if (!Volljaehrig)
            {
                x += ",\r\nliebe Erziehungsberechtigte";
            }

            x += "!";
            return x;
        }

        private string RenderGefährdungNeu()
        {
            string x = "Neu hinzukommende Gefährdung: ";
                        
            foreach (var item in (from f in DefizitäreLeistungen where f.NeueDefizitLeistung select f).ToList())
            {
                //x += " " + item.KürzelUntis + "(" + (from g in Global.Noten where item.NoteJetzt == g.Stufe select g.Klartext).FirstOrDefault() + "),";
            }
            return x.TrimEnd(',');
        }

        private object GetAbsatz3()
        {
            return "Wir laden Sie zu einem Beratungsgespräch ein. Stimmen Sie bitte den Gesprächster- min mit " + (KlassenleitungMw == "Herr" ? "dem Klassenlehrer" : "der Klassenlehrerin") + " " + Klassenleitung + " (" + KlassenleitungMail + ") ab.";
        }

        private object GetAbsatz2(string art)
        {            
            if(art == "M")
            {
                return "abweichend von " + ((from f in DefizitäreLeistungen where f.NeueDefizitLeistung select f).Count() > 1 ? "den" : "der") + " im letzten Zeugnis erteilten Note" + ((from f in DefizitäreLeistungen where f.NeueDefizitLeistung select f).Count() > 1 ? "n" : "") + " nicht mehr " + ((from f in DefizitäreLeistungen where f.NeueDefizitLeistung select f).Count() > 1 ? "ausreichen" : "ausreicht") + ". Stellt sich eine weitere nicht ausreichende Leistung ein, ist die Versetzung gefährdet.";
            }
            if (art == "G")
            {
                return "abweichend von " + ((from f in DefizitäreLeistungen where f.NeueDefizitLeistung select f).Count() > 1 ? "den" : "der") + " im letzten Zeugnis erteilten Note" + ((from f in DefizitäreLeistungen where f.NeueDefizitLeistung select f).Count() > 1 ? "n" : "") + " nicht mehr " + ((from f in DefizitäreLeistungen where f.NeueDefizitLeistung select f).Count() > 1 ? "ausreichen" : "ausreicht") + ".";
            }
            else
            {
                return "abweichend von der im letzten Zeugnis erteilten Note nur noch ungenügend ist.";
            }
        }

        private object GetAbsatz1(string art)
        {
            if (!Volljaehrig)
            {
                if (Geschlecht.ToLower() == "m")
                {
                    return "Sie werden darüber unterrichtet, dass die Leistung" + ((from d in DefizitäreLeistungen where d.NeueDefizitLeistung select d).Count() > 1 ? "en" : "") + " Ihres Sohnes " + Vorname + ", Klasse " + Klasse + ", in " + ((from d in DefizitäreLeistungen where d.NeueDefizitLeistung select d).Count() > 1 ? "den Fächern" : "dem Fach");
                }
                else
                {
                    return "Sie werden darüber unterrichtet, dass die Leistung" + ((from d in DefizitäreLeistungen where d.NeueDefizitLeistung select d).Count() > 1 ? "en" : "") + " Ihrer Tochter " + Vorname + ", Klasse " + Klasse + ", in " + ((from d in DefizitäreLeistungen where d.NeueDefizitLeistung select d).Count() > 1 ? "den Fächern" : "dem Fach");
                }
            }
            else
            {
                return "Sie werden darüber unterrichtet, dass Ihre Leistung" + ((from d in DefizitäreLeistungen where d.NeueDefizitLeistung select d).Count() > 1 ? "en" : "") + " in " + ((from f in DefizitäreLeistungen where f.NeueDefizitLeistung select f).Count() > 1 ? "den Fächern" : "dem Fach");
            }
        }

        internal void RenderBrief(string folder)
        {
            string footer = "";// (Volljaehrig ? "Vollj.;" : "Minderj.;" ) + Klasse + ";" + Nachname + "," + Vorname + "; HZ: " + RenderNotenHz() + "; Jetzt: " + RenderNotenJetzt() + "; ";
            Console.Write(footer, folder);
            Console.Write(Klasse + ";" + Nachname + ";" + Vorname + ";");

            if ((from f in DefizitäreLeistungen
                 where f.NoteHalbjahr == 5                
                 select f).Count() == 0)
            {
                if ((from f in DefizitäreLeistungen
                     where f.NoteHalbjahr == 6
                     select f).Count() == 0)
                {
                    // HZ: kein Defizit; jetzt eine 5: Mitteilung über Leistungsstand

                    if ((from f in DefizitäreLeistungen
                         where f.NoteJetzt == 5
                         select f).Count() == 1)
                    {
                        if ((from f in DefizitäreLeistungen
                             where f.NoteJetzt == 6
                             select f).Count() == 0)
                        {
                            //RenderMitteilung("M", footer, folder);
                        }
                    }

                    // HZ kein Defizit; jetzt zwei oder mehr 5: Gefährdung

                   if ((from f in DefizitäreLeistungen
                         where f.NoteJetzt == 5
                         select f).Count() > 1)
                    {
                        if ((from f in DefizitäreLeistungen
                             where f.NoteJetzt == 6
                             select f).Count() == 0)
                        {
                            //RenderMitteilung("G", footer, folder);
                        }
                    }

                    // HZ: kein Defizit; jetzt eine 6 oder mehr: Gefährdung

                    if ((from f in DefizitäreLeistungen
                         where f.NoteJetzt == 6
                         select f).Count() > 0)
                    {
                        //RenderMitteilung("G", footer, folder);
                    }
                }   
            }
            
            // HZ eine 5; jetzt eine oder mehrere zusätzliche 5en: Gefährdung
            
            if ((from f in DefizitäreLeistungen
                 where f.NoteJetzt == 5
                 select f).Count() == 1)
            {                
                if ((from f in DefizitäreLeistungen
                     where f.NoteJetzt == 6
                     select f).Count() == 0)
                {
                    if ((from f in DefizitäreLeistungen
                         where f.NoteJetzt == 6
                         select f).Count() == 0)
                    {
                        if ((from f in DefizitäreLeistungen
                             where f.NoteJetzt == 5
                             select f).Count() > (from f in DefizitäreLeistungen
                                                  where f.NoteJetzt == 5
                                                  select f).Count())
                        {
                            //RenderMitteilung("G", footer, folder);
                        }
                    }
                }
                
                // HZ eine 5; jetzt eine oder mehrere zusätzliche 6en: Gefährdung
                
                if ((from f in DefizitäreLeistungen
                     where f.NoteJetzt == 6
                     select f).Count() > (from f in DefizitäreLeistungen
                                          where f.NoteJetzt == 6
                                          select f).Count())
                {
                    //RenderMitteilung("G", footer, folder);                    
                }
            }

            // HZ: Zwei oder mehr 5er oder eine 6. Jetzt eine oder mehrere zusätzliche 5 oder 6: Gefährdung

            if ((from f in DefizitäreLeistungen where f.NoteJetzt == 6 select f).Count() >= 1 ||
                (from f in DefizitäreLeistungen where f.NoteJetzt == 5 select f).Count() > 1)
            {
                var anzahlHzDefizite5 = (from f in DefizitäreLeistungen
                                        where f.NoteHalbjahr == 5
                                        select f).Count();

                var anzahlJetztDefizite5 = (from f in DefizitäreLeistungen
                                           where f.NoteJetzt == 5
                                           select f).Count();

                var anzahlHzDefizite6 = (from f in DefizitäreLeistungen
                                        where f.NoteHalbjahr == 6 
                                        select f).Count();

                var anzahlJetztDefizite6 = (from f in DefizitäreLeistungen
                                           where f.NoteJetzt == 6 
                                           select f).Count();

                if (anzahlJetztDefizite5 > anzahlHzDefizite5 || anzahlJetztDefizite6 > anzahlHzDefizite6)
                {
                    //RenderMitteilung("G", footer, folder);             
                }
                //Abschlussklasse erhalten keine Benachrichtigung
            }
            Console.WriteLine("ok");
        }        

        private string RenderNotenHz()
        {
            string x = "";

            if((from f in DefizitäreLeistungen
                                  where f.NoteHalbjahr == 5 
                                  || f.NoteHalbjahr == 6
                                  select f).Count()  == 0)
            {
                x = "";                
            }

            x += "";

            foreach (var item in DefizitäreLeistungen)
            {
                x += " " + item.Fach + "(" + item.NoteHalbjahr + "),";
            }
            return x.TrimEnd(',');
        }

        //private string RenderNotenJetzt()
        //{
        //    string x = "";

        //    foreach (var item in Fachs)
        //    {
        //        x += " " + item.KürzelUntis + "(" + (from g in Global.Noten where item.NoteJetzt == g.Stufe select g.Klartext).FirstOrDefault() + "),";
        //    }
        //    return x.TrimEnd(',');
        //}

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

        private string RenderFächer(string art)
        {
            string x = "";

            if (art == "V")
            {
                foreach (var dl in (from d in DefizitäreLeistungen where d.NoteHalbjahr == 5 where d.NoteJetzt == 6 select d).ToList())
                {
                    x += " " + dl.BezeichnungImZeugnis + " (" + NoteKlartext(dl.NoteJetzt) + ")\r\n";
                }
            }
            else
            {
                foreach (var dl in (from d in DefizitäreLeistungen where d.NeueDefizitLeistung select d).ToList())
                {
                    x += " " + dl.BezeichnungImZeugnis + " (" + NoteKlartext(dl.NoteJetzt) + ")\r\n";
                }
            }            
            return x.Replace(" **)","");
        }

        private string NoteKlartext(int noteJetzt)
        {
            if (noteJetzt == 6)
            {
                return "ungenügend";
            }
            if (noteJetzt == 5)
            {
                return "mangelhaft";
            }
            if (noteJetzt == 4)
            {
                return "ausreichend";
            }
            if (noteJetzt == 3)
            {
                return "befriedigend";
            }
            if (noteJetzt == 2)
            {
                return "gut";
            }
            if (noteJetzt == 1)
            {
                return "sehr gut";
            }
            return "fehler";
        }

        private static void FindAndReplace(Application app, Document doc, object findText, object replaceWithText)
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
                // Der neue Text darf nur 255 Zeichen lang sein.

                if (replaceWithText.ToString().Length < 255)
                {
                    app.Selection.Find.Execute(ref findText, ref matchCase, ref matchWholeWord,
                    ref matchWildCards, ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format, ref replaceWithText, ref replace,
                    ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl);
                }
                else
                {
                    object empty = "";
                    Bookmark bm = doc.Bookmarks["faecher"];
                    Range range = bm.Range;
                    range.Text = replaceWithText.ToString();
                    doc.Bookmarks.Add("faecher", range);
                    app.Selection.Find.Execute(ref findText, ref matchCase, ref matchWholeWord,
                    ref matchWildCards, ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format, ref empty, ref replace,
                    ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl);
                }
            }
            catch (Exception ex)
            {                
                Console.WriteLine(ex);
                Console.ReadKey();
            }
        }

        //internal void GetDefizitfächer(DefizitäreLeistungen defizitäreLeistungen, DefizitäreLeistungen defizitäreAtlantisLeistungen, Fachs fachs)
        //{
        //    Fachs fachss = new Fachs();

        //    // Suche alle defizitären Fächer dieses Schülers
            
        //    var defizitäreFächerDiesesSchülers = (from d in defizitäreLeistungen
        //                                          where d.SchlüsselExtern == IdAtlantis
        //                                          where Global.BlaueBriefe.Contains(d.Prüfungsart)
        //                                          //where !d.Fach.BezeichnungImZeugnis.Contains("Förder")
        //                                          select d.Fach.KürzelUntis).Distinct().ToList();

        //    var noteJetzt = "";
        //    var noteHalbjahr = "";

        //    foreach (var dFach in defizitäreFächerDiesesSchülers)
        //    {
        //        foreach (var d in defizitäreLeistungen)
        //        {
        //            if (d.SchlüsselExtern == IdAtlantis)
        //            {
        //                if (d.Fach.KürzelUntis == dFach)
        //                {
        //                    if (d.Prüfungsart.Contains("laue"))
        //                    {
        //                        noteJetzt = d.BlauerBriefNote;
        //                        noteHalbjahr = d.GetNoteHalbjahr(defizitäreAtlantisLeistungen);
        //                    }                            
        //                }
        //            }
        //        }
                
        //        if (noteJetzt != null)
        //        {
        //            this.Fachs.Add(new Fach(dFach, (from f in fachs where f.KürzelUntis == dFach select f.BezeichnungImZeugnis).FirstOrDefault(), noteJetzt, noteHalbjahr));
        //        }
        //    }
        //}
    }
}