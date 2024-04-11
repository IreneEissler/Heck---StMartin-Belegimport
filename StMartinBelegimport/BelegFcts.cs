using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Data.SqlClient;
using Sagede.OfficeLine.Wawi.BelegEngine;
using Sagede.OfficeLine.Wawi.Tools;


namespace StMartinBelegimport
{
    class BelegFcts
    {

        public static void BelegImport()
        {
            string returnvalue;
            try
            {
                if (Properties.Settings.Default.Debug == 1) GlobalFcts.writeLog("Beginn Belegimport");

                // Neue Importdateien suchen
                string importPfad = Properties.Settings.Default.BelegPfadLokal;//System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                if (importPfad.Length > 0)
                {
                    if (Directory.Exists(importPfad))
                    {
                        DirectoryInfo di = new DirectoryInfo(importPfad);

                        //Alle Products-Importdateien durchgehen
                        FileInfo[] files = di.GetFiles("Bestellung_*.csv");
                        foreach (FileInfo file in files)
                        {
                            //Datei verarbeiten
                            returnvalue = importBelegDatei(importPfad + "\\" + file.Name);
                            //Datei verschieben
                            if (returnvalue.Length > 0)
                            {
                                file.MoveTo(importPfad + "\\Error\\" + file.Name);
                            }
                            else
                            {
                                //Datei ins Verarbeitet-Verzeichnis verschieben
                                file.MoveTo(importPfad + "\\Verarbeitet\\" + file.Name.Substring(0,file.Name.Length-4) + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".csv");
                            }

                        }
                    }
                }
                if (Properties.Settings.Default.Debug == 1) GlobalFcts.writeLog("Ende Belegimport");
            }
            catch (Exception ex)
            {
                GlobalFcts.writeLog("Fehler beim Import der Belege: " + ex.Message);
            }
        }

        private static string importBelegDatei(string fileName)
        {
            string sZeile;
            int lCount = 0;
            Beleg beleg;
            StreamReader sr = null;

            try
            {
                sr = new StreamReader(fileName, Encoding.GetEncoding(28591));
                sZeile = sr.ReadLine();
                if (sZeile != null) //Datei enthält Daten, Erste Zeile Kopfzeile
                {
                    string KundenNr = "";
                    string BelegDatum = "";
                    string Belegnummer = "";
                    string Bemerkung = "";
                    string sKto;
                    string sFehler = "";


                    if (!sZeile.Contains(";"))
                    {
                        GlobalFcts.writeLog("Fehler in Datei " + fileName + ", Kopfzeile " + lCount + ", Feld KundenNr.");
                    }
                    else
                    {
                        KundenNr = GlobalFcts.naechsterTeilstring(ref sZeile);
                        if (!sZeile.Contains(";"))
                        {
                            GlobalFcts.writeLog("Fehler in Datei " + fileName + ", Kopfzeile " + lCount + ", Feld BelegDatum.");
                            sFehler = "Fehler";
                        }
                        else
                        {
                            BelegDatum = GlobalFcts.naechsterTeilstring(ref sZeile);

                            if (!sZeile.Contains(";"))
                            {
                                GlobalFcts.writeLog("Fehler in Datei " + fileName + ", Kopfzeile " + lCount + ", Feld Belegnummer.");
                                sFehler = "Fehler";
                            }
                            else
                            {
                                Belegnummer = GlobalFcts.naechsterTeilstring(ref sZeile);

                                if (sZeile.Length > 0)
                                {
                                    Bemerkung = GlobalFcts.naechsterTeilstring(ref sZeile);
                                }
                            }
                        }
                    }
                    if (sFehler != "")
                    {
                        sr.Close();
                        return "Fehler";
                    }
                    else
                    {
                        sKto = GlobalFcts.vntLookup("Kto", "KHKKontokorrent", "Kto = '" + KundenNr + "' AND KtoArt = 'D' AND Mandant = " + GlobalFcts.mandant.Id);
                        if (sKto == null || sKto.Length == 0)
                        {
                            GlobalFcts.writeLog("Kundennummer nicht gefunden: " + KundenNr);
                            sr.Close();
                            return "Kundennummer nicht gefunden: " + KundenNr;
                        }

                        beleg = new Beleg(GlobalFcts.mandant, Erfassungsart.Verkauf);
                        beleg.Initialize("VVA", Convert.ToDateTime(BelegDatum), Convert.ToInt16(Convert.ToDateTime(BelegDatum).Year));
                        if (beleg.Errors.NumberOfErrors > 0)
                        {
                            GlobalFcts.writeLog("Fehler beim Initialisieren des Belegs: " + beleg.Errors.GetDescriptionSummary());
                            sr.Close();
                            return "Fehler beim Initialisieren des Belegs: " + beleg.Errors.GetDescriptionSummary();
                        }

                        if (!beleg.SetKonto(sKto, false))
                        {
                            GlobalFcts.writeLog("Fehler beim Setzen der Kundennummer " + sKto);
                            sr.Close();
                            return "Fehler beim Setzen der Kundennummer " + sKto;
                        }
                        beleg.RefreshPreiskennzeichen(true); //Zunächst Bruttopreise
                        beleg.Bearbeiter = GlobalFcts.mandant.Benutzer.Name;
                        beleg.Matchcode = "Webshop " + Belegnummer;
                        beleg.HauptvorgangsMatchcode = "Webshop " + Belegnummer;

                        beleg.ReadStandardTexte(true);

                        //Bemerkung in Textposition schreiben
                        if (Bemerkung.Length > 0)
                        {
                            BelegPosition position = new BelegPosition(beleg);
                            position.Initialize(Positionstyp.Texte);
                            position.Langtext = Bemerkung;
                            beleg.Positionen.Add(position);
                        }
                    }

                    //Positionen einlesen und anfügen
                    sZeile = sr.ReadLine();
                    lCount = 1;
                    while (sZeile != null && sFehler == "")
                    {
                        string Position = "";
                        string ArtikelNr = "";
                        string Menge = "";
                        string Einzelpreis = "";

                        if (sZeile.Length == 0) { } // leere Zeile, ignorieren
                        else
                        {

                            if (!sZeile.Contains(";"))
                            {
                                GlobalFcts.writeLog("Fehler in Datei " + fileName + ", Zeile " + lCount + ", Feld Position.");
                                sFehler = "Fehler";
                            }
                            else
                            {
                                Position = GlobalFcts.naechsterTeilstring(ref sZeile);

                                if (!sZeile.Contains(";"))
                                {
                                    GlobalFcts.writeLog("Fehler in Datei " + fileName + ", Zeile " + lCount + ", Feld ArtikelNr.");
                                    sFehler = "Fehler";
                                }
                                else
                                {
                                    ArtikelNr = GlobalFcts.naechsterTeilstring(ref sZeile);

                                    if (!sZeile.Contains(";"))
                                    {
                                        GlobalFcts.writeLog("Fehler in Datei " + fileName + ", Zeile " + lCount + ", Feld Menge.");
                                        sFehler = "Fehler";
                                    }
                                    else
                                    {
                                        Menge = GlobalFcts.naechsterTeilstring(ref sZeile);

                                        if (!(sZeile.Length > 0))
                                        {
                                            GlobalFcts.writeLog("Fehler in Datei " + fileName + ", Zeile " + lCount + ", Feld Einzelpreis.");
                                            sFehler = "Fehler";
                                        }
                                        else
                                        {
                                            Einzelpreis = GlobalFcts.naechsterTeilstring(ref sZeile);


                                            // Position zu Beleg hinzufügen
                                            BelegPosition position = new BelegPosition(beleg);
                                            if (GlobalFcts.vntLookup("Artikelnummer", "KHKArtikel", "Artikelnummer = '" + ArtikelNr + "' AND Mandant = " + GlobalFcts.mandant.Id) == null)
                                            {
                                                //Artikel nicht vorhanden, Artikeldaten als Textposition übergeben.
                                                GlobalFcts.writeLog("Fehler in Beleg " + Belegnummer + ": Artikel " + ArtikelNr + " nicht vorhanden.");

                                                return "Fehler in Beleg " + Belegnummer + ": Artikel " + ArtikelNr + " nicht vorhanden.";
                                            }
                                            else
                                            {
                                                position.Initialize(Positionstyp.Artikel);
                                                position.SetArtikel(ArtikelNr, 0);

                                                if (beleg.Errors.NumberOfErrors > 0)
                                                {
                                                    GlobalFcts.writeLog("Fehler beim Setzen der Position " + Position + ", " + ArtikelNr + ": " + beleg.Errors.GetDescriptionSummary());
                                                    return "Fehler beim Setzen der Position " + Position + ", " + ArtikelNr + ": " + beleg.Errors.GetDescriptionSummary();
                                                }

                                                position.Einzelpreis = Convert.ToDecimal(Einzelpreis.Replace(".", ","));
                                                position.IstEinzelpreisManuell = true;
                                                position.Menge = Convert.ToDecimal(Menge.Replace(".", ","));
                                                position.Calculate();

                                                //Hinzufügen der oben definierten Position zum Beleg
                                                beleg.Positionen.Add(position);

                                            }
                                        }
                                    }
                                }
                            }
                        }
                        sZeile = sr.ReadLine();
                        lCount = lCount + 1;
                    }
                    sr.Close();
                    if (sFehler == "")
                    {


                        beleg.Renumber();
                        if (!(beleg.Calculate(true)))
                        {
                            GlobalFcts.writeLog("Fehler beim Neuberechnen des Beleges " + Belegnummer + ": " + beleg.Errors.GetDescriptionSummary());
                            return "Fehler beim Neuberechnen des Beleges " + Belegnummer + ": " + beleg.Errors.GetDescriptionSummary();
                        }
                        beleg.RefreshPreiskennzeichen(false); //von Brutto auf Netto umstellen
                        if (!(beleg.Calculate(true)))
                        {
                            GlobalFcts.writeLog("Fehler beim Neuberechnen des Beleges " + Belegnummer + ": " + beleg.Errors.GetDescriptionSummary());
                            return "Fehler beim Neuberechnen des Beleges " + Belegnummer + ": " + beleg.Errors.GetDescriptionSummary();
                        }
                        beleg.Speichermode = BelegSpeicherstatus.Geparkt;
                        if (beleg.Validate())
                        {
                            if (beleg.Save(false))
                            {
                                GlobalFcts.writeLog("Beleg " + Belegnummer + " erfolgreich gespeichert.");
                            }
                            else
                            {
                                GlobalFcts.writeLog("Fehler beim Speichern des Beleges" + Belegnummer + ": " + beleg.Errors.GetDescriptionSummary());
                                return "Fehler beim Speichern des Beleges" + Belegnummer + ": " + beleg.Errors.GetDescriptionSummary();
                            }
                        }
                        else
                        {
                            GlobalFcts.writeLog("Fehler beim Validieren des Beleges" + Belegnummer + ": " + beleg.Errors.GetDescriptionSummary());
                            return "Fehler bei der Validierung des Beleges" + Belegnummer + ": " + beleg.Errors.GetDescriptionSummary();
                        }
                        GlobalFcts.writeLog("Aus Datei " + fileName + " wurden " + (lCount - 1).ToString() + " Zeilen erfolgreich importiert.");
                        return "";
                    }
                    else
                    {
                        return sFehler;
                    }

                }
                else
                {
                    GlobalFcts.writeLog("Datei " + fileName + " enthält keine Daten.");
                    sr.Close();
                    return "Datei " + fileName + " enthält keine Daten.";
                }
            }

            catch (Exception ex)
            {
                GlobalFcts.writeLog("Fehler in Datei " + fileName + ": " + ex.Message);
                if (sr != null)
                {
                    sr.Close();                   
                }

                return ex.Message;
            }
        }

    }
}
