using System.Collections.Generic;
using System.Data.OleDb;

namespace webuntis2BlaueBriefe
{
    internal class Global
    {
        public static List<string> Mangelhaft { get; internal set; }
        public static List<string> Ungenügend { get; internal set; }
        public static string Halbjahreszeugnis { get; internal set; }
        public static string BlaueBriefe { get; internal set; }

        public static string SafeGetString(OleDbDataReader reader, int colIndex)
        {
            if (!reader.IsDBNull(colIndex))
                return reader.GetString(colIndex);
            return string.Empty;
        }
    }
}