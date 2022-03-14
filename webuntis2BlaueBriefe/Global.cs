using System;
using System.Collections.Generic;
using System.Data.SqlClient;

namespace webuntis2BlaueBriefe
{
    internal class Global
    {
        public static List<string> Mangelhaft
        {
            get
            {
                return new List<string>() { "3.0", "2.0", "1.0" };
            }
        }
        public static List<string> Ungenügend
        {
            get
            {
                return new List<string>() { "0.0" };
            }
        }

        public static List<Note> Noten
        {
            get
            {
                List<Note> n = new List<Note>
                {
                    new Note("15.0", "sehr gut"),
                    new Note("14.0", "sehr gut"),
                    new Note("13.0", "sehr gut"),
                    new Note("12.0", "gut"),
                    new Note("11.0", "gut"),
                    new Note("10.0", "gut"),
                    new Note("9.0", "befriedigend"),
                    new Note("8.0", "befriedigend"),
                    new Note("7.0", "befriedigend"),
                    new Note("6.0", "ausreichend"),
                    new Note("5.0", "ausreichend"),
                    new Note("4.0", "ausreichend"),
                    new Note("3.0", "mangelhaft"),
                    new Note("2.0", "mangelhaft"),
                    new Note("1.0", "mangelhaft"),
                    new Note("0.0", "ungenügend")
                };
                return n;
            }
        }

        public static string Halbjahreszeugnis
        { get { return "Halbjahreszeugnis"; } }
        public static string BlaueBriefe
        { get { return "Mahnung gem. §50 (4) SchulG (Blauer Brief)"; } }

        public static string ConnectionStringAtlantis = @"Dsn=Atlantis9;uid=DBA";
        public static string ConnectionStringUntis = @"Data Source=SQL01\UNTIS;Initial Catalog=master;Integrated Security=True";

        public static string InputNotenCsv = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\MarksPerLesson.csv";

        public static List<string> Zeilen;

        public static string AktSjUntis
        {
            get
            {
                int sj = (DateTime.Now.Month >= 8 ? DateTime.Now.Year : DateTime.Now.Year - 1);
                return sj.ToString() + (sj + 1);
            }
        }

        public static string AktSjAtlantis
        {
            get
            {
                List<string> aktSj = new List<string>
                {
                    (DateTime.Now.Month >= 8 ? DateTime.Now.Year : DateTime.Now.Year - 1).ToString(),
                    (DateTime.Now.Month >= 8 ? DateTime.Now.Year + 1 - 2000 : DateTime.Now.Year - 2000).ToString()
                };
                return aktSj[0] + "/" + aktSj[1];
            }
        }
        
        public static string SafeGetString(SqlDataReader reader, int colIndex)
        {
            if (!reader.IsDBNull(colIndex))
                return reader.GetString(colIndex);
            return string.Empty;
        }
    }
}