using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;

namespace PVDailyMonitor
{
    internal class Program
    {
        private static string loggerFile { get; set; }
        private static Program p = new Program();

        private static void Main(string[] args)
        {
            // Create two dictionary instances to hold weekly and daily data.
            // The key is a string for the product names.
            // The value is a list of rowInfo objects to hold rows of data.
            Dictionary<string, List<rowInfo>> weeklyRowDict = new Dictionary<string, List<rowInfo>>();
            Dictionary<string, List<rowInfo>> dailyRowDict = new Dictionary<string, List<rowInfo>>();
            Dictionary<string, List<rowInfo>> wigigRowDict = new Dictionary<string, List<rowInfo>>();
            string[] cx = ConfigurationManager.AppSettings["CX"].Split(',');
            string[] wz = ConfigurationManager.AppSettings["WZ"].Split(',');
            string[] fs = ConfigurationManager.AppSettings["FS"].Split(',');
            string[] productList = cx.Union(wz).Union(fs).ToArray();
            string[] wigigProductList = ConfigurationManager.AppSettings["CX_WIGIG"].Split(',');

            foreach (string product in productList)
            {
                weeklyRowDict[product] = new List<rowInfo>();
                dailyRowDict[product] = new List<rowInfo>();
            }

            foreach (string product in wigigProductList)
            {
                wigigRowDict[product] = new List<rowInfo>();
            }

            // startDate/finishDate specifies the date range of data to be gathered.
            DateTime currDate = DateTime.Now.AddDays(-1);
            DateTime startDate = new DateTime();
            DateTime finishDate = new DateTime();

            // Create a unique identifier to be the ending of the report's name
            Random RandomGenerator = new Random();
            int rn = RandomGenerator.Next(0, 1000);
            string endingNum = rn.ToString().PadLeft(3, '0');
            string[] currReportList = Directory.GetFiles(ConfigurationManager.AppSettings["reportPath"]);

            // Check if there are any filename duplicates and regenerate endingNum if there exists a duplicate
            bool dup;
            do
            {
                dup = false;
                foreach (string report in currReportList)
                {
                    string[] substrs = report.Split('\\');
                    substrs = substrs[substrs.Length - 1].Split('_');

                    if (substrs[substrs.Length - 1].IndexOf(endingNum, StringComparison.Ordinal) == 0 &&
                        currDate.ToString("yyyyMMdd").Equals(substrs[0]))
                    {
                        dup = true;
                        break;
                    }
                }

                if (dup)
                {
                    rn = RandomGenerator.Next(0, 1000);
                    endingNum = rn.ToString().PadLeft(3, '0');
                }
            } while (dup);

            string logPath = ConfigurationManager.AppSettings.Get(@"logPath");
            loggerFile = Path.Combine(logPath, currDate.ToString(@"yyyyMMdd") + @"_" + endingNum + @"_log.txt");

            p.logger("########################################### OUTPUT/ERROR LOG ###########################################");
            p.GetDateRange(ref startDate, ref finishDate, ref weeklyRowDict, ref dailyRowDict, ref wigigRowDict, currDate, endingNum);

            // Begin gathering data and insert it into a Excel spreadsheet
            p.DoWork(startDate, finishDate, ref weeklyRowDict, ref dailyRowDict, ref wigigRowDict, endingNum);
        }

        /*
         * Function    : getDateRange()
         * Description : Sets the start and finish date. Gathers the latest report's data if it exists.
         * Params      : (ref DateTime) startDate
		 *				 (ref DateTime) finishDate
		 * 				 (ref Dictionary<string, List<rowInfo>>) weeklyRowDict
		 *				 (ref Dictionary<string, List<rowInfo>>) dailyRowDict
         *				 (ref Dictionary<string, List<rowInfo>>) wigigRowDict
		 *				 (DateTime) currDate
         *				 (string) endingNum
         * Return      :
         */

        public void GetDateRange(ref DateTime startDate, ref DateTime finishDate, ref Dictionary<string, List<rowInfo>> weeklyRowDict, ref Dictionary<string, List<rowInfo>> dailyRowDict, ref Dictionary<string, List<rowInfo>> wigigRowDict, DateTime currDate, string endingNum)
        {
            p.logger("[MAIN] Getting date range");

            // Get all reports from the current month/year
            string currDateString = currDate.ToString("yyyyMM");
            int currYear = Int32.Parse(currDate.ToString("yyyy"));
            int currMonth = Int32.Parse(currDate.ToString("MM"));
            int currDay = Int32.Parse(currDate.ToString("dd"));

            // default date range
            startDate = new DateTime(currYear, currMonth, 1, 0, 0, 0);
            finishDate = new DateTime(currYear, currMonth, currDay, 23, 59, 59);

            // to find existing daily reports for the month
            string[] cx = Directory.GetFiles(ConfigurationManager.AppSettings["reportPath"], currDateString + "*_CX*");
            string[] wz = Directory.GetFiles(ConfigurationManager.AppSettings["reportPath"], currDateString + "*_WZ*");
            string[] fs = Directory.GetFiles(ConfigurationManager.AppSettings["reportPath"], currDateString + "*_FS*");

            // Find the latest report date by search for the max
            DateTime cxMaxDate = startDate;
            DateTime wzMaxDate = startDate;
            DateTime fsMaxDate = startDate;

            // Iterate through every report file to check for the latest report
            foreach (string c in cx)
            {
                // Get the eight characters of the date (yyyyMMdd) to compare dates
                string tempDateString = c.Substring(c.IndexOf(currDateString, StringComparison.Ordinal), 8);
                DateTime dummyDate = new DateTime(Int32.Parse(tempDateString.Substring(0, 4)), Int32.Parse(tempDateString.Substring(4, 2)), Int32.Parse(tempDateString.Substring(6, 2)), 0, 0, 0);

                if (DateTime.Compare(cxMaxDate, dummyDate) < 0)
                {
                    cxMaxDate = dummyDate;
                }
            }
            foreach (string w in wz)
            {
                string tempDateString = w.Substring(w.IndexOf(currDateString, StringComparison.Ordinal), 8);
                DateTime dummyDate = new DateTime(Int32.Parse(tempDateString.Substring(0, 4)), Int32.Parse(tempDateString.Substring(4, 2)), Int32.Parse(tempDateString.Substring(6, 2)), 0, 0, 0);

                if (DateTime.Compare(wzMaxDate, dummyDate) < 0)
                {
                    wzMaxDate = dummyDate;
                }
            }
            foreach (string w in fs)
            {
                string tempDateString = w.Substring(w.IndexOf(currDateString, StringComparison.Ordinal), 8);
                DateTime dummyDate = new DateTime(Int32.Parse(tempDateString.Substring(0, 4)), Int32.Parse(tempDateString.Substring(4, 2)), Int32.Parse(tempDateString.Substring(6, 2)), 0, 0, 0);

                if (DateTime.Compare(fsMaxDate, dummyDate) < 0)
                {
                    fsMaxDate = dummyDate;
                }
            }

            string dateString = startDate.ToString("yyyyMMdd");
            if (cx.Length > 0 || wz.Length > 0 || fs.Length > 0)
            {
                DateTime minMaxDate = cxMaxDate;
                if (minMaxDate.CompareTo(wzMaxDate) > 0)
                {
                    minMaxDate = wzMaxDate;
                }
                if (minMaxDate.CompareTo(fsMaxDate) > 0)
                {
                    minMaxDate = fsMaxDate;
                }
                dateString = minMaxDate.ToString("yyyyMMdd");
                startDate = minMaxDate.AddDays(1);
            }
            DataRetrieval data = new DataRetrieval(startDate, finishDate, p);
            p.logger("[MAIN] Finished getting date range");
            data.GatherSpreadsheet(ref weeklyRowDict, ref dailyRowDict, ref wigigRowDict, dateString);
        }

        /*
         * Function    : DoWork()
         * Description : Calls functions in a sequence. Creates a directory, extract/gather data, insert
		 *				 data into an Excel spreadsheet then delete the directory.
         * Params      : (DateTime) startDate
		 *				 (DateTime) finishDate
		 * 				 (ref Dictionary<string, List<rowInfo>>) weeklyRowDict
		 *				 (ref Dictionary<string, List<rowInfo>>) dailyRowDict
         *				 (ref Dictionary<string, List<rowInfo>>) wigigRowDict
         *				 (string) endingNum
         * Return      :
         */

        public void DoWork(DateTime startDate, DateTime finishDate, ref Dictionary<string, List<rowInfo>> weeklyRowDict, ref Dictionary<string, List<rowInfo>> dailyRowDict, ref Dictionary<string, List<rowInfo>> wigigRowDict, string endingNum)
        {
            Folder dir = new Folder(p);
            DataRetrieval data = new DataRetrieval(startDate, finishDate, p);
            Report report = new Report(p);

            var timer = System.Diagnostics.Stopwatch.StartNew();
            dir.CreateDirectory(finishDate, endingNum);
            data.GatherData(ref weeklyRowDict, ref dailyRowDict, ref wigigRowDict, dir.GetXmlPath(), endingNum);
            try
            {
                report.CreateReport(ref weeklyRowDict, ref dailyRowDict, ref wigigRowDict, dir.GetReportPath(), dir.GetXmlPath());
                report.CreateGraph(ref weeklyRowDict, ref dailyRowDict, ref wigigRowDict, dir.GetReportPath(), dir.GetXmlPath());
            }
            catch (Exception e)
            {
                string logPath = ConfigurationManager.AppSettings.Get(@"logPath");
                Directory.CreateDirectory(logPath);
                string timeStamp = DateTime.Now.ToString("yyyyMMddHHmmssfff");

                string[] msg = { "" };
                msg[0] = @"Message:" + e.Message + @" StackTrace:" + e.StackTrace + @" TargetSite:" + e.TargetSite + @" Source:" + e.Source;
                File.AppendAllLines(Path.Combine(logPath, timeStamp + ".txt"), msg);
            }
            dir.DeleteDirectory(dir.GetXmlPath());
            timer.Stop();

            TimeSpan t = TimeSpan.FromMilliseconds(timer.ElapsedMilliseconds);
            string duration = string.Format("{0:D2}h:{1:D2}m:{2:D2}s:{3:D3}ms",
                t.Hours,
                t.Minutes,
                t.Seconds,
                t.Milliseconds);

            p.logger("[MAIN] All processes has been completed");
            p.logger("[MAIN] Time elapsed is " + duration);
            p.logger("[MAIN] Process completed");
        }

        /*
         * Function    : logger()
         * Description : Takes in a string as a parameter and write it into the log file.
         * Params      : (string) lines
         * Return      :
         */

        public void logger(string lines)
        {
            StreamWriter log = new StreamWriter(loggerFile, true);
            log.WriteLine(lines);
            log.Close();
        }
    }
}