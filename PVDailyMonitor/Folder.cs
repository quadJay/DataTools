using System;
using System.Configuration;
using System.IO;

namespace PVDailyMonitor
{
    internal class Folder
    {
        private string xmlPath;
        private string reportPath;
        private static Program p;

        public Folder(Program program)
        {
            xmlPath = ConfigurationManager.AppSettings["xmlPath"];
            reportPath = ConfigurationManager.AppSettings["reportPath"];
            p = program;
        }

        /*
         * Function    : CreateDirectory()
         * Description : Creates the directory specified by the "xmlPath" in file App.config.
         * Params      : (DateTime) date
		 *				 (string) endingNum
         * Return      :
         */

        public void CreateDirectory(DateTime date, string endingNum)
        {
            DateTime currDate = date;

            // Create a unique path to store the XML data into
            xmlPath = Path.Combine(xmlPath, currDate.ToString("yyyyMMdd") + @"_" + endingNum);

            try
            {
                p.logger("[Folder.cs] Directory to store data : " + xmlPath);
                Directory.CreateDirectory(xmlPath);
            }
            catch (Exception e)
            {
                p.logger("[Folder.cs] [Error] Directory: " + xmlPath + " could not be created");
                p.logger(e.ToString());
            }
        }

        /*
         * Function    : DeleteDirectory()
         * Description : Deletes the directory and its contents specified by the "xmlPath" in file App.config.
         * Params      : (string) xmlPath
         * Return      :
         */

        public void DeleteDirectory(string xmlPath)
        {
            p.logger("\n[Folder.cs] Deleting all directories and files within " + xmlPath);
            DirectoryInfo clean_dir = new DirectoryInfo(xmlPath);

            foreach (FileInfo f in clean_dir.GetFiles())
            {
                f.Delete();
            }

            foreach (DirectoryInfo dir in clean_dir.GetDirectories())
            {
                dir.Delete(true);
            }

            p.logger("[Folder.cs] Deleting process is complete");
        }

        /*
         * Function    : GetXmlPath()
         * Description : Private variable xmlPath's accessor
         * Params      :
         * Return      : (string) xmlPath
         */

        public string GetXmlPath()
        {
            return xmlPath;
        }

        /*
         * Function    : reportPath()
         * Description : Private variable reportPath's accessor
         * Params      : N/A
         * Return      : (string) reportPath
         */

        public string GetReportPath()
        {
            return reportPath;
        }
    }
}