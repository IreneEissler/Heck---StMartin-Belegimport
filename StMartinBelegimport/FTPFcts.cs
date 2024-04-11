using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using System.IO;

namespace StMartinBelegimport
{
    class FTPFcts
    {
        public static bool DateiHochladen(string filename, string zielFilename)
        {
            string URIstring = Properties.Settings.Default.ArtikelPfadFTP;
            try
            {
                //FTP-Verbindung herstellen
                Uri serverUri = new Uri(URIstring);

                if (Properties.Settings.Default.Debug == 1) GlobalFcts.writeLog("Beginn Datei " + filename + " hochladen.");
                // The serverUri parameter should start with the ftp:// scheme.
                if (serverUri.Scheme != Uri.UriSchemeFtp)
                {
                    GlobalFcts.writeLog("FTP-Verbindung fehlerhaft");
                    return false;
                }
                else
                {
                    NetworkCredential cred = new NetworkCredential(Properties.Settings.Default.FTPUser, Properties.Settings.Default.FTPKennwort);

                    //Datei hochladen
                    WebClient request = new WebClient();
                    request.Credentials = cred;
                    request.UploadFile(URIstring + zielFilename, Properties.Settings.Default.ArtikelPfadLokal + "\\" + filename);

                    //Prüfen, ob Datei angekommen ist.
                    FtpWebRequest ftprequest = (FtpWebRequest)WebRequest.Create(URIstring + zielFilename);
                    ftprequest.Credentials = cred;
                    ftprequest.Method = WebRequestMethods.Ftp.ListDirectory;

                    FtpWebResponse response = (FtpWebResponse)ftprequest.GetResponse();

                    Stream responseStream = response.GetResponseStream();
                    StreamReader reader = new StreamReader(responseStream);
                    if (reader.ReadLine() != null)
                    {
                        if (Properties.Settings.Default.Debug == 1) GlobalFcts.writeLog("Datei " + filename + " hochgeladen.");
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

        public static void GetBestellungen()
        {
            string URIstring = Properties.Settings.Default.BelegPfadFTP;//"ftp://stmartinsapotheke.eu/webios/_data/_stmartinsapotheke.eu/_export/_asc/";

            try
            {
                //FTP-Verbindung herstellen
                Uri serverUri = new Uri(URIstring);

                if (Properties.Settings.Default.Debug == 1) GlobalFcts.writeLog("Beginn Dateien herunterladen.");
                // The serverUri parameter should start with the ftp:// scheme.
                if (serverUri.Scheme != Uri.UriSchemeFtp)
                {
                    GlobalFcts.writeLog("FTP-Verbindung fehlerhaft");
                }
                else
                {
                    //Neue Bestellungen auflisten
                    NetworkCredential cred = new NetworkCredential(Properties.Settings.Default.FTPUser, Properties.Settings.Default.FTPKennwort);
                    FtpWebRequest ftprequest = (FtpWebRequest)WebRequest.Create(URIstring );
                    ftprequest.Credentials = cred;
                    ftprequest.Method = WebRequestMethods.Ftp.ListDirectory;

                    FtpWebResponse response = (FtpWebResponse)ftprequest.GetResponse();

                    Stream responseStream = response.GetResponseStream();
                    StreamReader reader = new StreamReader(responseStream);

                    string filename;

                    filename = reader.ReadLine();
                    while (!(filename == null))
                    {
                        if (filename.Substring(0,Math.Min(filename.Length, 11)) == "Bestellung_")
                        {
                            if (Properties.Settings.Default.Debug == 1) GlobalFcts.writeLog("Datei " + filename + " herunterladen");

                            // nächste Datei herunterladen
                            WebClient request = new WebClient();
                            request.Credentials = cred;
                            request.DownloadFile(URIstring + filename, Properties.Settings.Default.BelegPfadLokal + "\\" + filename);

                            //Prüfen, ob Datei angekommen ist.
                            FileInfo fil = new FileInfo(Properties.Settings.Default.BelegPfadLokal + "\\" + filename);
                            if (!fil.Exists)
                            {
                                GlobalFcts.writeLog("Fehler beim Herunterladen der Datei " + filename);
                            }
                            else
                            {
                                //Datei erfolgreich heruntergeladen -> vom Server löschen
                                FtpWebRequest ftprequestDel = (FtpWebRequest)WebRequest.Create(URIstring + filename);
                                ftprequestDel.Credentials = cred;
                                ftprequestDel.Method = WebRequestMethods.Ftp.DeleteFile;
                                FtpWebResponse responseDel = (FtpWebResponse)ftprequestDel.GetResponse();
                                if (Properties.Settings.Default.Debug == 1) GlobalFcts.writeLog("Delete status: {0} " + responseDel.StatusDescription);
                                responseDel.Close();
                            }
                           
                        }
                        filename = reader.ReadLine();
                    }
                    reader.Close();
                    response.Close();
                    if (Properties.Settings.Default.Debug == 1) GlobalFcts.writeLog("Ende Dateien herunterladen");
                }

            }
            catch (WebException e)
            {
                GlobalFcts.writeLog("Fehler beim Herunterladen der Belegdateien: " + e.ToString());
            }

        }

    }
}
