using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;

namespace webuntis2BlaueBriefe
{
    public class Stundentafels : List<Stundentafel>
    {
        public Stundentafels(Fachs fachs)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(Global.ConnectionStringUntis))
                {
                    string queryString = @"SELECT 
PeriodsTable.PERIODS_TABLE_ID, 
PeriodsTable.Name, 
PeriodsTable.Longname, 
PeriodsTable.PerTabElement1
FROM PeriodsTable
WHERE (((PeriodsTable.SCHOOLYEAR_ID)=" + Global.AktSjUntis + ") AND ((PeriodsTable.Deleted)='false')) ORDER BY Name;";

                    SqlCommand odbcCommand = new SqlCommand(queryString, sqlConnection);
                    sqlConnection.Open();
                    SqlDataReader sqlDataReader = odbcCommand.ExecuteReader();

                    while (sqlDataReader.Read())
                    { 
                        try
                        {
                            if (!(from s in this where s.IdUntis == sqlDataReader.GetInt32(0) select s).Any())
                            {
                                Stundentafel stundentafel = new Stundentafel();

                                stundentafel.IdUntis = sqlDataReader.GetInt32(0);
                                stundentafel.Name = Global.SafeGetString(sqlDataReader, 1);
                                stundentafel.Langname = Global.SafeGetString(sqlDataReader, 2);
                                var elemente = (Global.SafeGetString(sqlDataReader, 3)).Split(',');

                                for (int i = 0; i < elemente.Count(); i++)
                                {
                                    if (elemente[i].Length < 2)
                                    {
                                    }
                                    else
                                    {
                                        var teile = elemente[i].Split('~');

                                        if (teile[19] == "F")
                                        {
                                            var fa = (from f in fachs where f.KürzelUntis == teile[2] select f).FirstOrDefault();
                                            if (fa == null)
                                            {
                                                fa = (from f in fachs where f.LangnameUntis == teile[2] select f).FirstOrDefault();
                                            }
                                            else
                                            {
                                                Fach fach = new Fach()
                                                {
                                                    KürzelUntis = fa.KürzelUntis
                                                };
                                                stundentafel.Fachs.Add(fach);
                                            }
                                        }
                                    }
                                }
                                this.Add(stundentafel);
                            }                                   
                        }
                        catch (Exception ex)
                        {
                        }
                    };
                    
                    sqlDataReader.Close();
                    sqlConnection.Close();
                }
              
            }
            catch (Exception ex)
            {
              
            }
        }
    }
}