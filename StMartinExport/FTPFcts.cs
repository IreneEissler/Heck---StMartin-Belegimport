using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using System.IO;

namespace StMartinExport
{
    class FTPFcts
    {
        public static bool DateiHochladen(string appPath,string filename, string zielFilename, bool bDebug)
        {
            string URIstring = GlobalFcts.mandant.MainDevice.Lookup.GetString("strValue", "WUDGrundlagen", " Mandant = " + GlobalFcts.mandant.Id + " AND strKey = 'ExportPfadFTP' AND UserName = 'All' AND Owner = 'StMartinExport'", ""); 
            try
            {
                //FTP-Verbindung herstellen
                Uri serverUri = new Uri(URIstring);

                if (bDebug) GlobalFcts.writeLog("Beginn Datei " + filename + " hochladen.");
                // The serverUri parameter should start with the ftp:// scheme.
                if (serverUri.Scheme != Uri.UriSchemeFtp)
                {
                    GlobalFcts.writeLog("FTP-Verbindung fehlerhaft");
                    return false;
                }
                else
                {
                    NetworkCredential cred = new NetworkCredential(GlobalFcts.mandant.MainDevice.Lookup.GetString("strValue", "WUDGrundlagen", " Mandant = " + GlobalFcts.mandant.Id + 
                                " AND strKey = 'FTPUser' AND UserName = 'All' AND Owner = 'StMartinExport'", ""), GlobalFcts.mandant.MainDevice.Lookup.GetString("strValue", "WUDGrundlagen", 
                                " Mandant = " + GlobalFcts.mandant.Id + " AND strKey = 'FTPKennwort' AND UserName = 'All' AND Owner = 'StMartinExport'", ""));

                    //Datei hochladen
                    WebClient request = new WebClient();
                    request.Credentials = cred;
                    request.UploadFile(URIstring + zielFilename, appPath + "\\" + filename);

                    //Prüfen, ob Datei angekommen ist.
                    FtpWebRequest ftprequest = (FtpWebRequest)WebRequest.Create(URIstring + zielFilename);
                    ftprequest.Credentials = cred;
                    ftprequest.Method = WebRequestMethods.Ftp.ListDirectory;

                    FtpWebResponse response = (FtpWebResponse)ftprequest.GetResponse();

                    Stream responseStream = response.GetResponseStream();
                    StreamReader reader = new StreamReader(responseStream);
                    if (reader.ReadLine() != null)
                    {
                        if (bDebug) GlobalFcts.writeLog("Datei " + filename + " hochgeladen.");
                        reader.Close();
                        response.Close();
                        return true;
                    }
                    else
                    {
                        GlobalFcts.writeLog("Fehler beim Hochladen der Datei " + filename);
                        reader.Close();
                        response.Close();
                        return false;
                    }
                }
            }
            catch (WebException e)
            {
                GlobalFcts.writeLog("Fehler beim Hochladen der Datei: " + e.ToString());
                return false;
            }
        }

   

    }
}
