
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.IO;
using Sagede.OfficeLine.Wawi.BelegEngine;
using Sagede.OfficeLine.Engine;
using Sagede.OfficeLine.Data;
using Sagede.OfficeLine.Shared;

namespace StMartinBelegimport
{
    class GlobalFcts
    {
        public static Mandant mandant;
        public static Session goSession;

        internal static void connectToOL()
        {
            string database = Properties.Settings.Default.Datenbank;
            string benutzer = Properties.Settings.Default.OLBenutzer;
            string kennwort = Properties.Settings.Default.OLKennwort;
            short mandantId = Properties.Settings.Default.Mandant;
            try
            {
                //Mandantenobjekt erzeugen
                if (mandant == null)
                {
                    goSession = ApplicationEngine.CreateSession(database, ApplicationToken.Abf, null, new Sagede.Core.Tools.NamePasswordCredential(benutzer, kennwort));
                    mandant = goSession.CreateMandant(mandantId);
                }
                if (mandant == null)
                {
                    GlobalFcts.writeLog("Fehler bei der Verbindung zur Office Line." );
                }
                else
                {
                    if (Properties.Settings.Default.Debug == 1) GlobalFcts.writeLog("Verbindung zur Office Line erfolgreich hergestellt.");
                }
            }
            catch (Exception ex)
            {
                GlobalFcts.writeLog("Fehler bei der Verbindung zur OfficeLine: " + ex.Message);
            }
        }

        internal static string vntLookup(string Feld, string Tabelle, string Condition)
        {
            //Debug.Assert(!String.IsNullOrEmpty(Feld), "Feld is null or empty.");
            //Debug.Assert(!String.IsNullOrEmpty(Tabelle), "Tabelle is null or empty.");
            //Debug.Assert(!String.IsNullOrEmpty(Condition), "Condition is null or empty.");
            string sQry;

            try
            {
                sQry = " SELECT " + Feld + " FROM " + Tabelle + " WHERE " + Condition;
                IGenericCommand command = mandant.MainDevice.GenericConnection.CreateSqlStringCommand();
                command.CommandText = sQry;
                using (IGenericReader reader = command.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        return reader.GetValue(0).ToString();
                    }
                    else
                    {
                        return null;
                    }
                }
            }
            catch (Exception ex)
            {
                GlobalFcts.writeLog("Fehler in vntLookup: " + ex.Message);
                return null;
            }
        }

        internal static void writeLog(string errorDescription)
        {
            String appPfad;
            StreamWriter sw;

            appPfad = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            FileInfo fil = new FileInfo(appPfad + "\\StMBelegimport.log");
            if (!fil.Exists)
            {
                File.Create(appPfad + "\\StMBelegimport.log").Close();
                fil = new FileInfo(appPfad + "\\StMBelegimport.log");
            }
            if (fil.Length > 100000000) //größer 100 MB neue Datei erzeugen
            {
                fil.MoveTo(appPfad + "\\StMBelegimport" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".log");
                File.Create(appPfad + "\\StMBelegimport.log").Close();
            }

            sw = File.AppendText(appPfad + "\\StMBelegimport.log");
            sw.WriteLine(DateTime.Now + " " + errorDescription);
            sw.Close();
        }

        internal static void disconnectFromOL()
        {
            if (!(mandant== null))
            {
                mandant.Dispose();
                   }
            mandant = null;
            if (!(goSession == null))
            {
                goSession.Dispose();
            }
            goSession = null;
        }

        public static string naechsterTeilstring(ref string sZeile)
        {
            string returnValue;
            if (sZeile == null || sZeile.Length == 0)
            {
                return "";
            }
            else
            {
                // Teilstring abtrennen
                if (sZeile.Substring(0, 1) == "\"") // Teilstring beginnt mit "
                {
                    sZeile = sZeile.Substring(1);
                    if (sZeile.Contains("\""))
                    {
                        returnValue = sZeile.Substring(0, sZeile.IndexOf("\""));
                        sZeile = sZeile.Substring(sZeile.IndexOf("\"") + 1);
                        if (sZeile.Contains(";"))
                        {
                            sZeile = sZeile.Substring(sZeile.IndexOf(";") + 1);
                        }
                        else
                        {
                            sZeile = "";
                        }
                        return returnValue;
                    }
                    else
                        return "";
                }
                else
                {
                    if (sZeile.Contains(";"))
                    {
                        returnValue = sZeile.Substring(0, sZeile.IndexOf(";"));
                        sZeile = sZeile.Substring(sZeile.IndexOf(";") + 1);
                        return returnValue;
                    }
                    else
                    {
                        returnValue = sZeile;
                        sZeile = "";
                        return returnValue;
                    }
                }
            }
        }
    }
}
