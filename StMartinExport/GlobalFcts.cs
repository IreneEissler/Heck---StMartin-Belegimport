using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.IO;
using Sagede.OfficeLine.Engine;

namespace StMartinExport
{
    class GlobalFcts
    {
        public static Mandant mandant;
        public static Session goSession;
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
                fil.MoveTo(appPfad + "\\StMBelegimport" + DateTime.Now.ToString("_yyyyMMdd_HHmmss") + ".log");
                File.Create(appPfad + "\\StMBelegimport.log").Close();
            }

            sw = File.AppendText(appPfad + "\\StMBelegimport.log");
            sw.WriteLine(DateTime.Now + " " + errorDescription);
            sw.Close();
        }
    }
}
