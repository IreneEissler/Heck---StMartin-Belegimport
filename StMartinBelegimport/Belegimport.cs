using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Timers;

namespace StMartinBelegimport
{
    public partial class Belegimport : ServiceBase
    {

        private static System.Timers.Timer aTimer;

        public Belegimport()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            try
            {
                if (Properties.Settings.Default.Debug == 1) GlobalFcts.writeLog("OnStart");

                Properties.Settings.Default.LetzterImport = DateTime.Now.ToString();
                DateTime exportDatum;
                DateTime letzterExport;
                exportDatum = Convert.ToDateTime( Properties.Settings.Default.ExportZeit);
                //letzterExport = DateTime.Today.AddDays(-2); //vorgestern
                letzterExport = DateTime.Today; //heute
                letzterExport = letzterExport.AddHours(exportDatum.Hour).AddMinutes(exportDatum.Minute); //Plus Minuten

                // letzte mögliche Exportzeit vor aktuellem Datum suchen.
                if (letzterExport > DateTime.Now)
                {
                    while (letzterExport > DateTime.Now)
                    {
                        letzterExport = letzterExport.AddMinutes(-Properties.Settings.Default.IntervallExport);
                    }
                }
                else
                {
                    while (letzterExport < DateTime.Now)
                    {
                        letzterExport = letzterExport.AddMinutes(Properties.Settings.Default.IntervallExport);
                    }
                    letzterExport = letzterExport.AddMinutes(-Properties.Settings.Default.IntervallExport);
                }
                Properties.Settings.Default.LetzterExport = letzterExport.ToString();
                Properties.Settings.Default.Save();

                aTimer = new System.Timers.Timer(10000);

                //// Hook up the Elapsed event for the timer.
                aTimer.Elapsed += new ElapsedEventHandler(OnTimedEvent);


                //' Set the Interval to 2 seconds (2000 milliseconds).
                aTimer.Interval = 60000; //einmal pro Minute
                aTimer.Enabled = true;
                aTimer.Start();

                GC.KeepAlive(aTimer);
                if (Properties.Settings.Default.Debug == 1) GlobalFcts.writeLog(" Ende OnStart");
            }

            catch (Exception ex)
            {
                GlobalFcts.writeLog("Fehler beim Start: " + ex.Message);
            }

        }

        private static void OnTimedEvent(object source, ElapsedEventArgs e)
        {
            aTimer.Stop();

            Action();

            aTimer.Start();

        }

        internal static void Action()
        {
            try
            {

                if (DateTime.Now > Properties.Settings.Default.LetzteAnmeldung.AddMinutes(3))
                {
                    //viermal am Tag die Verbindung neu aufbauen.
                    GlobalFcts.disconnectFromOL();

                    Properties.Settings.Default.LetzteAnmeldung = DateTime.Now;
                }

                //Belegintervall überschritten? - Import Bestellungen
                if (Convert.ToDateTime(Properties.Settings.Default.LetzterImport).AddMinutes(Properties.Settings.Default.IntervallImport) < DateTime.Now && Properties.Settings.Default.ABImportAktiv == 1)
                {
                    GlobalFcts.connectToOL();

                    FTPFcts.GetBestellungen();
                    if (GlobalFcts.mandant != null)
                    {
                        BelegFcts.BelegImport();
                        Properties.Settings.Default.LetzterImport = Convert.ToDateTime(Properties.Settings.Default.LetzterImport).AddMinutes(Properties.Settings.Default.IntervallImport).ToString();
                        Properties.Settings.Default.Save();
                    }
                }

                //Belegintervall überschritten? - Import Rechnungen
                if (Convert.ToDateTime(Properties.Settings.Default.LetzterImport).AddMinutes(Properties.Settings.Default.IntervallImport) < DateTime.Now && Properties.Settings.Default.RechnungsImportAktiv == 1)
                {
                    GlobalFcts.connectToOL();

                    if (GlobalFcts.mandant != null)
                    {
                        RechnungsImport.BelegImport();
                        Properties.Settings.Default.LetzterImport = Convert.ToDateTime(Properties.Settings.Default.LetzterImport).AddMinutes(Properties.Settings.Default.IntervallImport).ToString();
                        Properties.Settings.Default.Save();
                    }
                }

                // Export Artikel
                if (Convert.ToDateTime(Properties.Settings.Default.LetzterExport).AddMinutes(Properties.Settings.Default.IntervallExport) < DateTime.Now && Properties.Settings.Default.ArtikelExportAktiv == 1)
                {

                    GlobalFcts.connectToOL();
                    if (GlobalFcts.mandant != null)
                    {
                        ArtikelFcts.ArtikelExport();
                        KundenFcts.KundenExport();
                        Properties.Settings.Default.LetzterExport = Convert.ToDateTime(Properties.Settings.Default.LetzterExport).AddMinutes(Properties.Settings.Default.IntervallExport).ToString();
                        Properties.Settings.Default.Save();
                    }
                }


            }
            catch (Exception ex)
            {
                GlobalFcts.writeLog("Fehler bei der Verarbeitung: " + ex.Message);
                GlobalFcts.mandant = null;
                GlobalFcts.goSession = null;
                aTimer.Start();
            }
        }

        protected override void OnStop()
        {
            GlobalFcts.writeLog("OnStop");
        }
    }
}
