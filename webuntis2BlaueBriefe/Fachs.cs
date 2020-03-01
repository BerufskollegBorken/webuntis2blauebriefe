using System;
using System.Collections.Generic;
using System.Data.OleDb;

namespace webuntis2BlaueBriefe
{
    public class Fachs:List<Fach>
    {
        public Fachs(string aktSj, string connectionStringUntis)
        {
            using (OleDbConnection oleDbConnection = new OleDbConnection(connectionStringUntis))
            {
                try
                {
                    string queryString = @"SELECT DISTINCT 
                                            Subjects.Subject_ID,
                                            Subjects.Name,
                                            Subjects.Longname,
                                            Subjects.Text
                                            FROM Subjects 
                                            WHERE Schoolyear_id = " + aktSj + " AND Deleted=No ORDER BY Name;";

                    OleDbCommand oleDbCommand = new OleDbCommand(queryString, oleDbConnection);
                    oleDbConnection.Open();
                    OleDbDataReader oleDbDataReader = oleDbCommand.ExecuteReader();

                    while (oleDbDataReader.Read())
                    {
                        Fach fach = new Fach()
                        {
                            IdUntis = oleDbDataReader.GetInt32(0),
                            KürzelUntis = SafeGetString(oleDbDataReader, 1),
                            LangnameUntis = SafeGetString(oleDbDataReader, 2),
                            BezeichnungImZeugnis = SafeGetString(oleDbDataReader, 3)                            
                        };
                        this.Add(fach);
                    };

                    oleDbDataReader.Close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex);
                    Console.ReadKey();
                }
                finally
                {
                    oleDbConnection.Close();
                }
            }
        }

        private object SafeGetString(OleDbDataReader reader, int colIndex)
        {
            if (!reader.IsDBNull(colIndex))
                return (String)reader.GetString(colIndex);
            return string.Empty;
        }
    }
}