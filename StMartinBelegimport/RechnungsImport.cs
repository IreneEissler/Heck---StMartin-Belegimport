﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

using Sagede.OfficeLine.Wawi.BelegEngine;
using Sagede.OfficeLine.Wawi.Tools;

namespace StMartinBelegimport
{
    class RechnungsImport
    {
        public static void BelegImport()
        {
            string returnvalue;
            try
            {
                if (Properties.Settings.Default.Debug == 1) GlobalFcts.writeLog("Beginn Rechnungsimport");

                // Neue Importdateien suchen
                string importPfad = Properties.Settings.Default.RechnungsPfadLokal;//System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                if (importPfad.Length > 0)
                {
                    if (Directory.Exists(importPfad))
                    {
                        DirectoryInfo di = new DirectoryInfo(importPfad);

                        //Alle Products-Importdateien durchgehen
                        FileInfo[] files = di.GetFiles("*.csv");
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
                                file.MoveTo(importPfad + "\\Verarbeitet\\" + file.Name.Substring(0, file.Name.Length - 4) + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".csv");
                            }

                        }
                    }
                }
                if (Properties.Settings.Default.Debug == 1) GlobalFcts.writeLog("Ende Rechnungsimport");
            }
            catch (Exception ex)
            {
                GlobalFcts.writeLog("Fehler beim Import der Rechnungen: " + ex.Message);
            }
        }

        private static string importBelegDatei(string fileName)
        {
            string sZeile;
            int lCount = 0;
            Beleg beleg;
            StreamReader sr = null;
            string dummy;

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
                        GlobalFcts.writeLog("Fehler in Datei " + fileName + ", Kopfzeile " + lCount + ", Feld Belegdatum.");
                    }
                    else
                    {
                        KundenNr = GlobalFcts.naechsterTeilstring(ref sZeile); // 1. Spalte Kundennummer


                        if (!sZeile.Contains(";"))
                        {
                            GlobalFcts.writeLog("Fehler in Datei " + fileName + ", Kopfzeile " + lCount + ", Feld Kundennummer.");
                            sFehler = "Fehler";
                        }
                        else
                        {
                            BelegDatum = GlobalFcts.naechsterTeilstring(ref sZeile); // 2. Spalte Belegdatum

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
                        bool bOK = !GlobalFcts.mandant.MainDevice.Lookup.RowExists("Belid", "KHKVKBelege", "Matchcode = '" + "ProKas " + Belegnummer + "' and Belegkennzeichen = 'VSD' and Mandant =" + GlobalFcts.mandant.Id);
                        if (!bOK)
                        {
                            GlobalFcts.writeLog("Beleg " + Belegnummer + " wurde bereits eingelesen" );
                            sr.Close();
                            return "Beleg " + Belegnummer + " wurde bereits eingelesen";
                        }
                        beleg = new Beleg(GlobalFcts.mandant, Erfassungsart.Verkauf);
                        beleg.Initialize("VSD", Convert.ToDateTime(BelegDatum), Convert.ToInt16(Convert.ToDateTime(BelegDatum).Year));
                        if (beleg.Errors.NumberOfErrors > 0)
                        {
                            GlobalFcts.writeLog("Fehler beim Initialisieren der Rechnung: " + beleg.Errors.GetDescriptionSummary());
                            sr.Close();
                            return "Fehler beim Initialisieren der Rechnung: " + beleg.Errors.GetDescriptionSummary();
                        }

                        if (!beleg.SetKonto(sKto, false))
                        {
                            GlobalFcts.writeLog("Fehler beim Setzen der Kundennummer " + sKto);
                            sr.Close();
                            return "Fehler beim Setzen der Kundennummer " + sKto;
                        }
                        beleg.RefreshPreiskennzeichen(true); //Zunächst Bruttopreise
                        beleg.Bearbeiter = GlobalFcts.mandant.Benutzer.Name;
                        beleg.Matchcode = "ProKas " + Belegnummer;
                        beleg.HauptvorgangsMatchcode = "ProKas " + Belegnummer;
                        beleg.UserProperties["USER_ImportiertProKas"].Value = -1;
                        beleg.ReadStandardTexte(true);
                        beleg.Referenznummer = Belegnummer;

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

                        if (beleg.Validate())
                        {
                            if (beleg.Save(true))
                            {
                                GlobalFcts.writeLog("Beleg " + Belegnummer + " erfolgreich gespeichert.");
                                beleg.ReweUebergabe(beleg.Belegdatum);
                                writeBuchungsdatei(Belegnummer, beleg);
                            }
                            else  // Fehler beim Speichern.
                            {
                                GlobalFcts.writeLog("Fehler beim Speichern des Beleges" + Belegnummer + ": " + beleg.Errors.GetDescriptionSummary());
                                beleg.Memo = "Fehler beim Speichern des Beleges" + Belegnummer + ": " + beleg.Errors.GetDescriptionSummary();
                                // Falls Fehler beim Speichern, beleg parken
                                beleg.Speichermode = BelegSpeicherstatus.Geparkt;
                                if (beleg.Save(true))
                                {
                                    GlobalFcts.writeLog("Beleg " + Belegnummer + " erfolgreich gespeichert.");
                                    beleg.ReweUebergabe(beleg.Belegdatum);

                                }
                                else
                                {
                                    //auch Parken nicht möglich
                                    GlobalFcts.writeLog("Fehler beim Parken des Beleges" + Belegnummer + ": " + beleg.Errors.GetDescriptionSummary());
                                    return "Fehler beim Parken des Beleges" + Belegnummer + ": " + beleg.Errors.GetDescriptionSummary();
                                }
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

        private static void writeBuchungsdatei(string belegnummerProkas, Beleg beleg)
        {
            try
            {
                string sZeile = "";
                sZeile += beleg.Belegdatum.ToShortDateString() + ";";
                sZeile += beleg.Belegdatum.ToShortDateString() + ";";
                sZeile += belegnummerProkas + ";";
                sZeile += "\"" + Properties.Settings.Default.Buchungstext + " " + belegnummerProkas + "\";";
                sZeile += beleg.VKRechnungsempfaenger + ";";
                sZeile += Properties.Settings.Default.SachkontoAusbuchung + ";";
                sZeile += beleg.BelegnummerFormatiert + ";";
                sZeile += beleg.Bruttobetrag + ";";
                sZeile += beleg.BruttobetragEW + ";";
                sZeile += "-1";

                string appPfad = Properties.Settings.Default.RechnungsPfadLokal + "\\Buchungsdatei";
                string fileandpath = appPfad + "\\OPAusbuchungProkasImport_" + beleg.Periode + ".txt";
                FileInfo fil = new FileInfo(fileandpath);
                if (!fil.Exists)
                {
                    File.Create(fileandpath).Close();
                    StreamWriter swh = File.AppendText(fileandpath);
                    swh.WriteLine("Buchungsdatum;Belegdatum;Belegnummer;Buchungstext;KtoHaben;KtoSoll;OPNummer;Buchungsbetrag;BuchungsbetragEW;ProkasImport");
                    swh.Close();
                    fil = new FileInfo(fileandpath);
                }

                StreamWriter sw = File.AppendText(fileandpath);
                sw.WriteLine(sZeile);
                sw.Close();

            }
            catch (Exception ex)
            {
                GlobalFcts.writeLog("Fehler beim Erstellen des Buchungssatzes für Datei " + belegnummerProkas + ": " + ex.Message);
            }
        }
    }
}
