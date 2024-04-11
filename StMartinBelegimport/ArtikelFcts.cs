using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Sagede.OfficeLine.Data;

namespace StMartinBelegimport
{
    class ArtikelFcts
    {
        public static void ArtikelExport()
        {
            string appPath = Properties.Settings.Default.ArtikelPfadLokal + "\\";
            string fileName = "Artikelstamm_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".csv";
            string PreislisteID = Properties.Settings.Default.Preisliste;
            string sQry;
            string Zeile;

            try
            {
                if (Properties.Settings.Default.Debug == 1) GlobalFcts.writeLog("Beginn Artikeldaten exportieren");
                //Artikeldaten holen
                sQry = " SELECT A.Artikelnummer, A.Bezeichnung1, A.Bezeichnung2, A.USER_WebShopArtikelgruppe AS Artikelgruppe, PA.Einzelpreis, ISNULL(SK.Lieferung, 0) AS Steuercode ";
                sQry += " FROM KHKArtikel A INNER JOIN KHKPreislistenArtikel PA";
                sQry += " ON A.Artikelnummer = PA.Artikelnummer AND A.Mandant = PA.mandant ";
                sQry += " LEFT OUTER JOIN KHKSteuerklassen SK ON A.Steuerklasse = SK.Steuerklasse  AND SK.Land = '*'";
                sQry += " WHERE PA.ListeID = " + PreislisteID + " AND A.USER_WebShop= -1 AND PA.AbMenge = 0 AND A.Mandant = " + GlobalFcts.mandant.Id;

                IGenericCommand command = GlobalFcts.mandant.MainDevice.GenericConnection.CreateSqlStringCommand();
                command.CommandText = sQry;

                //Datei schreiben
                using (IGenericReader reader = command.ExecuteReader())
                {
                    StreamWriter sw = new StreamWriter(appPath + fileName, false, Encoding.GetEncoding(1252));

                    while (reader.Read())
                    {
                        Zeile = reader.GetValue("Artikelnummer").ToString();
                        Zeile += ";" + reader.GetValue("Bezeichnung1").ToString().Replace(";", ",");
                        Zeile += ";" + reader.GetValue("Bezeichnung2").ToString().Replace(";", ",");
                        Zeile += ";" + reader.GetValue("Artikelgruppe").ToString();
                        if (Convert.ToInt16(GlobalFcts.vntLookup("IstBruttopreis","KHKPreislisten", "ID = " + PreislisteID + " AND Mandant = " + GlobalFcts.mandant.Id)) == 0) 
                        {   //Preisliste ist netto -> in Bruttopreis umrechnen
                            decimal Steuerprozent = Convert.ToDecimal(GlobalFcts.vntLookup("Steuersatz", "KHKSteuertabelle", "Steuercode = " +reader.GetValue("Steuercode").ToString() ));
                            Zeile += ";" + (reader.GetDecimal("Einzelpreis")* (1+ Steuerprozent/100)).ToString("0.00");
                        }
                        else
                        {
                            Zeile += ";" + reader.GetDecimal("Einzelpreis").ToString("0.##");
                        }                        
                        sw.WriteLine(Zeile, Encoding.GetEncoding(1252));
                    }
                    sw.Close();
                }

                //prüfen, ob Datei erstellt wurde
                FileInfo fil = new FileInfo(appPath + fileName);
                if (!fil.Exists)
                {
                    GlobalFcts.writeLog("Fehler beim Erstellen der Artikeldatei " + appPath + fileName);
                }
                else
                {
                    //Datei auf den FTP-Server laden
                    if (FTPFcts.DateiHochladen(fileName, "Artikelstamm.csv"))
                    {
                        fil.MoveTo(appPath + "Versandt\\" + fileName);
                    }
                    else
                    {
                        fil.MoveTo(appPath + "Error\\" + fileName);
                    }
                }
                if (Properties.Settings.Default.Debug == 1) GlobalFcts.writeLog("Ende Artikeldaten exportieren: " + appPath + fileName);
            }
            catch (Exception ex)
            {
                GlobalFcts.writeLog("Fehler beim Export der Artikeldaten: " + ex.Message);
            }



        }
    }
}
