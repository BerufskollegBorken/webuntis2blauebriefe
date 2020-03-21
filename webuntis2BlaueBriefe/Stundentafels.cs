using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;

namespace webuntis2BlaueBriefe
{
    public class Stundentafels : List<Stundentafel>
    {
        public Stundentafels(Fachs fachs)
        {
            try
            {
                using (OleDbConnection oleDbConnection = new OleDbConnection(Global.ConnectionStringUntis))
                {
                    string queryString = @"SELECT 
PeriodsTable.PERIODS_TABLE_ID, 
PeriodsTable.Name, 
PeriodsTable.Longname, 
PeriodsTable.PerTabElement1
FROM PeriodsTable
WHERE (((PeriodsTable.SCHOOLYEAR_ID)=" + Global.AktSjUntis + ") AND ((PeriodsTable.Deleted)=No)) ORDER BY Name;";

                    OleDbCommand oleDbCommand = new OleDbCommand(queryString, oleDbConnection);
                    oleDbConnection.Open();
                    OleDbDataReader oleDbDataReader = oleDbCommand.ExecuteReader();

                    while (oleDbDataReader.Read())
                    { 
                        try
                        {
                            if (!(from s in this where s.IdUntis == oleDbDataReader.GetInt32(0) select s).Any())
                            {
                                Stundentafel stundentafel = new Stundentafel();

                                stundentafel.IdUntis = oleDbDataReader.GetInt32(0);
                                stundentafel.Name = Global.SafeGetString(oleDbDataReader, 1);
                                stundentafel.Langname = Global.SafeGetString(oleDbDataReader, 2);
                                var elemente = (Global.SafeGetString(oleDbDataReader, 3)).Split(',');

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
                    
                    oleDbDataReader.Close();
                    oleDbConnection.Close();
                }
              
            }
            catch (Exception ex)
            {
              
            }
        }
    }
}