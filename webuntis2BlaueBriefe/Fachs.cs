using System;
using System.Collections.Generic;
using System.Data.SqlClient;

namespace webuntis2BlaueBriefe
{
    public class Fachs:List<Fach>
    {
        public Fachs()
        {
        }

        public Fachs(string connectionStringUntis)
        {
            using (SqlConnection sqlConnection = new SqlConnection(Global.ConnectionStringUntis))
            {
                try
                {
                    string queryString = @"SELECT DISTINCT 
                                            Subjects.Subject_ID,
                                            Subjects.Name,
                                            Subjects.Longname,
                                            Subjects.Text
                                            FROM Subjects 
                                            WHERE Schoolyear_id = " + Global.AktSjUntis + " AND (Deleted='false') ORDER BY Name;";

                    SqlCommand odbcCommand = new SqlCommand(queryString, sqlConnection);
                    sqlConnection.Open();
                    SqlDataReader sqlDataReader = odbcCommand.ExecuteReader();

                    while (sqlDataReader.Read())
                    {
                        Fach fach = new Fach()
                        {
                            IdUntis = sqlDataReader.GetInt32(0),
                            KürzelUntis = SafeGetString(sqlDataReader, 1).ToString(),
                            LangnameUntis = SafeGetString(sqlDataReader, 2).ToString(),
                            BezeichnungImZeugnis = SafeGetString(sqlDataReader, 3).ToString()                            
                        };
                        this.Add(fach);
                    };

                    sqlDataReader.Close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex);
                    Console.ReadKey();
                }
                finally
                {
                    sqlConnection.Close();
                }
            }
        }
                
        private object SafeGetString(SqlDataReader reader, int colIndex)
        {
            if (!reader.IsDBNull(colIndex))
                return (String)reader.GetString(colIndex);
            return string.Empty;
        }
    }
}