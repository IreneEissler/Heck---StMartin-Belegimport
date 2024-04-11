using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Sagede.OfficeLine.Data;
using Sagede.OfficeLine.Engine;
using System.Runtime.InteropServices;

namespace StMartinExport
{
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class KundenFcts
    {
        public bool KundenExportExtern(OLSysIInterop80.Mandant goMandant)
        {
            Mandant mandant = (goMandant as Sagede.OfficeLine.Interop80.Mandant).GetRealObject;
            if (KundenExport(false, mandant))
            {
                System.Diagnostics.Process.Start("http://stmartinsapotheke.eu/webios/spezial/x_import_kunden.pl");
                return true;
            }
            else return false;
        }

        public static bool KundenExport(bool bDebug, Mandant mandant)
        {
            string fileName = "Kundenstamm" + DateTime.Now.ToString("_yyyyMMdd_HHmmss") + ".csv";
            string sQry;
            string Zeile;

            try
            {
                GlobalFcts.mandant = mandant;
                string appPath = GlobalFcts.mandant.MainDevice.Lookup.GetString("strValue", "WUDGrundlagen", " Mandant = " + GlobalFcts.mandant.Id + " AND strKey = 'ExportPfadLokal' AND UserName = 'All' AND Owner = 'StMartinExport'", "");

                if (bDebug) GlobalFcts.writeLog("Beginn Kundendatei erstellen");
                //Kundendaten holen
                sQry = " SELECT K.Kto, A.Adresse, A.Name1, A.Name2, A.LieferStrasse, A.LieferPLZ, A.LieferOrt, A.Telefon, A.Telefax, A.Email, K.USER_WebKennwort, ISNULL(V.Matchcode, '') AS [Therapeut]";
                sQry += " FROM KHKKontokorrent K INNER JOIN KHKAdressen A";
                sQry += " ON A.Adresse = K.Adresse AND A.Mandant = K.mandant ";
                sQry += " LEFT OUTER JOIN KHKVertreter V ON V.Vertreternummer = K.Vertreter AND V.Mandant = K.Mandant ";
                sQry += " WHERE A.Email IS NOT NULL AND LEN(A.Email) > 0 AND K.USER_WebShop = -1 AND  K.USER_WebKennwort IS NOT NULL ";
                sQry += " AND K.KtoArt = 'D' AND A.Mandant = " + GlobalFcts.mandant.Id;

                IGenericCommand command = GlobalFcts.mandant.MainDevice.GenericConnection.CreateSqlStringCommand();
                command.CommandText = sQry;

                //Datei schreiben
                using (IGenericReader reader = command.ExecuteReader())
                {
                    StreamWriter sw = new StreamWriter(appPath + fileName, false, Encoding.GetEncoding(1252));

                    while (reader.Read())
                    {
                        Zeile = reader.GetValue("Kto").ToString();
                        Zeile += ";" + reader.GetValue("Adresse").ToString();
                        Zeile += ";" + reader.GetValue("Name1").ToString().Replace(";", ",").Replace("'", "\\'");
                        Zeile += ";" + reader.GetValue("Name2").ToString().Replace(";", ",").Replace("'", "\\'");
                        Zeile += ";" + reader.GetValue("LieferStrasse").ToString().Replace(";", ",").Replace("'", "\\'");
                        Zeile += ";" + reader.GetValue("LieferPLZ").ToString();
                        Zeile += ";" + reader.GetValue("LieferOrt").ToString().Replace(";", ",").Replace("'", "\\'");
                        Zeile += ";" + reader.GetValue("Telefon").ToString().Replace("'", "\\'");
                        Zeile += ";" + reader.GetValue("Telefax").ToString().Replace("'", "\\'");
                        Zeile += ";" + reader.GetValue("Email").ToString().Replace(";", ",");
                        Zeile += ";" + reader.GetValue("USER_WebKennwort").ToString().Replace("'", "\\'");
                        Zeile += ";" + reader.GetValue("Therapeut").ToString().Replace(";", ",").Replace("'", "\\'");

                        sw.WriteLine(Zeile, Encoding.GetEncoding(1252));
                    }
                    sw.Close();

                    //prüfen, ob Datei erstellt wurde
                    FileInfo fil = new FileInfo(appPath + fileName);
                    if (!fil.Exists)
                    {
                        GlobalFcts.writeLog("Fehler beim Erstellen der Kundendatei " + appPath + fileName);
                        return false;
                    }
                    else
                    {
                        //Datei auf den FTP-Server laden
                        if (FTPFcts.DateiHochladen(appPath, fileName, "Kundenstamm.csv", bDebug))
                        {
                            fil.MoveTo(appPath + "Versandt\\" + fileName);
                        }
                        else
                        {
                            fil.MoveTo(appPath + "Error\\" + fileName);
                        }
                    }
                    if (bDebug) GlobalFcts.writeLog("Kundendatei erstellt: " + appPath + fileName);

                }
                return true;
            }
            catch (Exception ex)
            {
                GlobalFcts.writeLog("Fehler beim Export der Kundendaten: " + ex.Message);
                return false;
            }


        }
    }
}