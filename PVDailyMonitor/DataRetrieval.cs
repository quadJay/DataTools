using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices;
using System.Xml;

namespace PVDailyMonitor
{
    internal class DataRetrieval
    {
        [DllImport("user32.dll")]
        private static extern int GetWindowThreadProcessId(int hWnd, out int lpdwProcessId);

        private string[] productList;
        private string[] pullPathList;
        private string xmlPath;
        private DateTime startDate;
        private DateTime finishDate;
        private static Program p;

        /*
         * DataRetrieval's default constructor
         */

        public DataRetrieval(DateTime sd, DateTime fd, Program program)
        {
            // Set properties productList, pullPathList from App.config.
            // Set properties startDate, finishDate. Values passed in from instance
            // declaration in Program.cs.
            string[] cx = ConfigurationManager.AppSettings["CX"].Split(',');
            string[] wz = ConfigurationManager.AppSettings["WZ"].Split(',');
            string[] cx_wigig = ConfigurationManager.AppSettings["CX_WIGIG"].Split(',');
            string[] fs = ConfigurationManager.AppSettings["FS"].Split(',');
            productList = cx.Union(wz).Union(cx_wigig).Union(fs).ToArray();
            pullPathList = ConfigurationManager.AppSettings["pullPathList"].Split(',');
            startDate = sd;
            finishDate = fd;
            p = program;
        }

        /*
        * Function    : GetProcess()
        * Description : Returns a Process variable of an Application instance.
        * Params      : (Application) app
        * Return      :
        */

        private Process GetProcess(Microsoft.Office.Interop.Excel.Application app)
        {
            int id;
            GetWindowThreadProcessId(app.Hwnd, out id);
            return Process.GetProcessById(id);
        }

        /*
         * Function    : GetDateRange()
         * Description : Set the list of string containing the eight characters dates (yyyyMMdd) to dateRange.
         * Params      : (ref List<string>) dateRange
         * Return      :
         */

        public void GetDateRange(ref List<string> dateRange)
        {
            if (startDate.Date.ToString("d").Equals(finishDate.Date.ToString("d")))
            {
                string tempDate = startDate.ToString("yyyyMMdd");
                dateRange.Add(tempDate);
            }
            else
            {
                DateTime tempDate = startDate;

                while (!tempDate.Date.ToString("d").Equals(finishDate.Date.ToString("d")))
                {
                    string date = tempDate.ToString("yyyyMMdd");
                    dateRange.Add(date);
                    tempDate = tempDate.AddDays(1);
                }

                string lastDate = tempDate.ToString("yyyyMMdd");
                dateRange.Add(lastDate);
            }

            p.logger("[DataRetrieval.cs] Total # of days in range : [" + dateRange.Count.ToString() + "]");
        }

        /*
         * Function    : GetWorkDay()
         * Description : Gets the day of the week as an integer ranging from Sunday to Saturday, 1 to 7 respectively.
         * Params      : (DateTime) workDate
         * Return      :
         */

        public string GetWorkDay(DateTime workDate)
        {
            if (workDate.DayOfWeek == DayOfWeek.Sunday)
            {
                return "1";
            }
            else if (workDate.DayOfWeek == DayOfWeek.Monday)
            {
                return "2";
            }
            else if (workDate.DayOfWeek == DayOfWeek.Tuesday)
            {
                return "3";
            }
            else if (workDate.DayOfWeek == DayOfWeek.Wednesday)
            {
                return "4";
            }
            else if (workDate.DayOfWeek == DayOfWeek.Thursday)
            {
                return "5";
            }
            else if (workDate.DayOfWeek == DayOfWeek.Friday)
            {
                return "6";
            }
            else
            {
                return "7";
            }
        }

        /*
         * Function    : UpdateRow()
         * Description : Check/Updates rows.
         * Params      : (string) productName,
		 *			 	 (ref Dictionary<string, List<rowInfo>>) rowDict,
		 *			 	 (int) i,
		 *			 	 (XmlNode) uut
         * Return      :
         */

        public void UpdateRow(string productName, ref Dictionary<string, List<rowInfo>> rowDict, int i, XmlNode uut)
        {
            string[] errorCodeList = ConfigurationManager.AppSettings["errorCodeList"].Split(',');

            if (uut["Indicator_1A2A"].InnerText.Equals("1A"))
            {
                rowDict[productName][i].TotalBoards += 1;
            }

            foreach (string errorCode in errorCodeList)
            {
                if (string.Equals(uut["ErrorCode"].InnerText, errorCode,
                StringComparison.CurrentCultureIgnoreCase))
                {
                    rowDict[productName][i].FailedBoards += 1;
                }
            }
        }

        /*
         * Function    : CreateRow()
         * Description : Creates a new rowInfo object to be added onto the weeklyRowDict/dailyRowDict
		 *				 for a product name passed.
         * Params      : (string) productName
		 *				 (bool) weeklyRowExist
		 *				 (bool) dailyRowExist
		 *				 (bool) wigigRowExist
		 *				 (ref Dictionary<string, List<rowInfo>>) weeklyRowDict
		 *				 (ref Dictionary<string, List<rowInfo>>) dailyRowDict
		 *				 (ref Dictionary<string, List<rowInfo>>) wigigRowDict
		 *				 (XmlNode) uut
		 *				 (FileInfo) zip
		 *				 (string) workWeek
         *				 (string) workDay
         * Return      :
         */

        public void CreateRow(string productName, bool weeklyRowExist, bool dailyRowExist, bool wigigRowExist, ref Dictionary<string, List<rowInfo>> weeklyRowDict, ref Dictionary<string, List<rowInfo>> dailyRowDict, ref Dictionary<string, List<rowInfo>> wigigRowDict, XmlNode uut, FileInfo zip, string workWeek, string workDay)
        {
            // Create the rowInfo object
            rowInfo weeklyRow = new rowInfo();
            rowInfo dailyRow = new rowInfo();
            rowInfo wigigRow = new rowInfo();

            // Set attributes then check/update attributes totalboards and failedboards
            weeklyRow.WorkYear = dailyRow.WorkYear = wigigRow.WorkYear = uut["TestTime"].InnerText.Substring(6, 4);
            weeklyRow.WorkWeek = dailyRow.WorkWeek = wigigRow.WorkWeek = workWeek;
            weeklyRow.ODM = dailyRow.ODM = wigigRow.ODM = zip.Name.Substring(8, 2);

            if (uut["Indicator_1A2A"].InnerText.Equals("1A"))
            {
                weeklyRow.TotalBoards += 1;
                dailyRow.TotalBoards += 1;
                wigigRow.TotalBoards += 1;
            }

            if (string.Equals(uut["ErrorCode"].InnerText, "WIFI_FT_ERR-034",
                StringComparison.CurrentCultureIgnoreCase))
            {
                weeklyRow.FailedBoards += 1;
                dailyRow.FailedBoards += 1;
            }
            else if (string.Equals(uut["ErrorCode"].InnerText, "WIGIG_FT-ERR-128",
                StringComparison.CurrentCultureIgnoreCase))
            {
                wigigRow.FailedBoards += 1;
            }

            // Check a couple attributes for values of Null, Empty, WhiteSpace
            // to avoid adding empty rows
            if (!string.IsNullOrEmpty(weeklyRow.WorkYear)
                && !string.IsNullOrEmpty(weeklyRow.WorkWeek)
                && !string.IsNullOrWhiteSpace(weeklyRow.WorkYear)
                && !string.IsNullOrWhiteSpace(weeklyRow.WorkWeek))
            {
                if (!weeklyRowExist)
                {
                    weeklyRowDict[productName].Add(weeklyRow);
                }

                if (!dailyRowExist)
                {
                    dailyRow.WorkDay = workDay;
                    dailyRowDict[productName].Add(dailyRow);
                }

                if (!wigigRowExist)
                {
                    wigigRow.WorkDay = workDay;
                    wigigRowDict[productName].Add(wigigRow);
                }
            }
        }

        /*
         * Function    : GatherData()
         * Description : Extracting XML files from zip file. Read the XML files to parse data and insert them to
		 *				 the dictionaries accordingly.
         * Params      : (Dictionary<string, List<rowInfo>>) weeklyRowDict
		 *			 	 (Dictionary<string, List<rowInfo>>) dailyRowDict
         *			 	 (Dictionary<string, List<rowInfo>>) wigigRowDict
         *               (string) path
         *               (string) endingNum
         * Return      :
         */

        public void GatherData(ref Dictionary<string, List<rowInfo>> weeklyRowDict, ref Dictionary<string, List<rowInfo>> dailyRowDict, ref Dictionary<string, List<rowInfo>> wigigRowDict, string path, string endingNum)
        {
            p.logger("[DataRetrieval.cs] Initiating gathering data");
            p.logger("\n[DataRetrieval.cs] Gathering data with the date range of " + startDate.ToString("MMM d yyyy") + " to " + finishDate.ToString("MMM d yyyy"));

            xmlPath = path;
            int xmlProcessed = 0;
            List<string> dateRange = new List<string>();
            GetDateRange(ref dateRange);

            CultureInfo culInfo = new CultureInfo("en-Us");
            Calendar cal = culInfo.Calendar;
            CalendarWeekRule calWeekRule = culInfo.DateTimeFormat.CalendarWeekRule;
            DayOfWeek firstDayOfWeek = culInfo.DateTimeFormat.FirstDayOfWeek;

            // Start iterating through each product's directory and gather/read XML data
            foreach (string pullPath in pullPathList)
            {
                foreach (string product in productList)
                {
                    string productDir = Path.Combine(pullPath, product);
                    p.logger("Gathering from : " + productDir);
                    if (Directory.Exists(productDir) == false)
                    {
                        continue;
                    }

                    DirectoryInfo pathInfo = new DirectoryInfo(productDir);
                    IEnumerable<FileInfo> zips = new FileInfo[] { };
                    try
                    {
                        zips = pathInfo.GetFiles("*.zip", SearchOption.AllDirectories).OrderBy(p => p.CreationTimeUtc);
                    }
                    catch (Exception e)
                    {
                        p.logger("[DataRetrieval.cs] " + e.ToString());
                    }

                    string[] foqm = { "FQM", "FOQM", "322", "WFFOQM" };
                    string[] wgfoqm = { "WGFOQM" };

                    string tempXmlPath = Path.Combine(xmlPath, product);
                    foreach (FileInfo zip in zips)
                    {
                        try
                        {
                            string zipDate = zip.Name.Substring(0, 8);
                            if (dateRange.Contains(zipDate))
                            {
                                Console.WriteLine("\nExtracting [{0}]\n", zip.FullName);

                                try
                                {
                                    ZipFile.ExtractToDirectory(zip.FullName, tempXmlPath);
                                }
                                catch (Exception e)
                                {
                                    p.logger("[DataRetrieval.cs] [Error] Can't extract the zip file :" + zip.FullName);
                                    p.logger("[DataRetrieval.cs] " + e.ToString());
                                }

                                DirectoryInfo xmlDirInfo = new DirectoryInfo(tempXmlPath);

                                IEnumerable<DirectoryInfo> stations = xmlDirInfo.GetDirectories("*", SearchOption.TopDirectoryOnly);

                                foreach (DirectoryInfo station in stations)
                                {
                                    if (foqm.Contains(station.Name) == true)
                                    {
                                        IEnumerable<FileInfo> xmls = station.GetFiles("*.xml", SearchOption.AllDirectories);
                                        XmlDocument doc = new XmlDocument();
                                        foreach (FileInfo xml in xmls)
                                        {
                                            Console.WriteLine("XML : [{0}]", xml.Name);

                                            xmlProcessed++;
                                            doc.Load(xml.FullName);
                                            XmlNode uut = doc.SelectSingleNode("/Log/Test/UUT");

                                            string productName = uut["ProductName"].InnerText;
                                            string testTime = uut["TestTime"].InnerText;
                                            DateTime workDate = new DateTime(Int32.Parse(testTime.Substring(6, 4)), Int32.Parse(testTime.Substring(0, 2)), Int32.Parse(testTime.Substring(3, 2)));
                                            string workDay = GetWorkDay(workDate);

                                            // Check if any rows exists for the product. If there is then check if the current XML
                                            // matches the existing rows to be added else add a new row.
                                            if (weeklyRowDict[productName].Count != 0 || dailyRowDict[productName].Count != 0)
                                            {
                                                bool weeklyRowExist = false;
                                                bool dailyRowExist = false;

                                                for (int i = 0; i < weeklyRowDict[productName].Count; i++)
                                                {
                                                    if (weeklyRowDict[productName][i].WorkYear.Equals(testTime.Substring(6, 4))
                                                        && weeklyRowDict[productName][i].WorkWeek.Equals(cal.GetWeekOfYear(workDate, calWeekRule, firstDayOfWeek).ToString())
                                                        && weeklyRowDict[productName][i].ODM.Equals(zip.Name.Substring(8, 2)))
                                                    {
                                                        weeklyRowExist = true;
                                                        UpdateRow(productName, ref weeklyRowDict, i, uut);
                                                        break;
                                                    }
                                                }

                                                for (int i = 0; i < dailyRowDict[productName].Count; i++)
                                                {
                                                    if (dailyRowDict[productName][i].WorkYear.Equals(testTime.Substring(6, 4))
                                                        && dailyRowDict[productName][i].WorkWeek.Equals(cal.GetWeekOfYear(workDate, calWeekRule, firstDayOfWeek).ToString())
                                                        && dailyRowDict[productName][i].ODM.Equals(zip.Name.Substring(8, 2))
                                                        && dailyRowDict[productName][i].WorkDay.Equals(workDay))
                                                    {
                                                        dailyRowExist = true;
                                                        UpdateRow(productName, ref dailyRowDict, i, uut);
                                                        break;
                                                    }
                                                }

                                                // Create and add the first row for weeklyRowDict/dailyRowDict whichever is nonexistent
                                                if (!weeklyRowExist || !dailyRowExist)
                                                {
                                                    CreateRow(productName, weeklyRowExist, dailyRowExist, true, ref weeklyRowDict, ref dailyRowDict, ref wigigRowDict, uut, zip, cal.GetWeekOfYear(workDate, calWeekRule, firstDayOfWeek).ToString(), workDay);
                                                }
                                            }
                                            else
                                            {
                                                // Create and add the first row for both weeklyRowDict/dailyRowDict
                                                CreateRow(productName, false, false, true, ref weeklyRowDict, ref dailyRowDict, ref wigigRowDict, uut, zip, cal.GetWeekOfYear(workDate, calWeekRule, firstDayOfWeek).ToString(), workDay);
                                            }
                                        }
                                    }
                                    if (wgfoqm.Contains(station.Name))
                                    {
                                        IEnumerable<FileInfo> xmls = station.GetFiles("*.xml", SearchOption.AllDirectories);
                                        XmlDocument doc = new XmlDocument();
                                        foreach (FileInfo xml in xmls)
                                        {
                                            Console.WriteLine("XML : [{0}]", xml.Name);

                                            xmlProcessed++;
                                            doc.Load(xml.FullName);
                                            XmlNode uut = doc.SelectSingleNode("/Log/Test/UUT");
                                            string productName = uut["ProductName"].InnerText;
                                            string testTime = uut["TestTime"].InnerText;
                                            DateTime workDate = new DateTime(Int32.Parse(testTime.Substring(6, 4)), Int32.Parse(testTime.Substring(0, 2)), Int32.Parse(testTime.Substring(3, 2)));
                                            string workDay = GetWorkDay(workDate);

                                            if (wigigRowDict[productName].Count != 0)
                                            {
                                                bool wigigRowExist = false;

                                                for (int i = 0; i < wigigRowDict[productName].Count; i++)
                                                {
                                                    if (wigigRowDict[productName][i].WorkYear.Equals(testTime.Substring(6, 4))
                                                        && wigigRowDict[productName][i].WorkWeek.Equals(cal.GetWeekOfYear(workDate, calWeekRule, firstDayOfWeek).ToString())
                                                        && wigigRowDict[productName][i].ODM.Equals(zip.Name.Substring(8, 2))
                                                        && wigigRowDict[productName][i].WorkDay.Equals(workDay))
                                                    {
                                                        wigigRowExist = true;
                                                        UpdateRow(productName, ref wigigRowDict, i, uut);
                                                        break;
                                                    }
                                                }

                                                if (!wigigRowExist)
                                                {
                                                    CreateRow(productName, true, true, wigigRowExist, ref weeklyRowDict, ref dailyRowDict, ref wigigRowDict, uut, zip, cal.GetWeekOfYear(workDate, calWeekRule, firstDayOfWeek).ToString(), workDay);
                                                }
                                            }
                                            else
                                            {
                                                CreateRow(productName, true, true, false, ref weeklyRowDict, ref dailyRowDict, ref wigigRowDict, uut, zip, cal.GetWeekOfYear(workDate, calWeekRule, firstDayOfWeek).ToString(), workDay);
                                            }
                                        }
                                    }
                                }

                                // Delete all directories and XML files to clear up space for more data
                                Folder dir = new Folder(p);
                                dir.DeleteDirectory(tempXmlPath);
                            }
                        }
                        catch (Exception e)
                        {
                            p.logger("[DataRetrieval.cs] [Error] Can't extract the zip file :" + zip.FullName);
                            p.logger(e.ToString());
                        }
                    }
                }
            }

            p.logger("\n[DataRetrieval.cs] Completed gathering data.");
            p.logger("[DataRetrieval.cs] XML files read : " + xmlProcessed);

            // Calculate the failure rates for dictionaries
            foreach (KeyValuePair<string, List<rowInfo>> entry in weeklyRowDict)
            {
                if (weeklyRowDict[entry.Key].Count > 0)
                {
                    for (int i = 0; i < weeklyRowDict[entry.Key].Count; i++)
                    {
                        weeklyRowDict[entry.Key][i].GetFailureRate();
                    }
                }
            }

            foreach (KeyValuePair<string, List<rowInfo>> entry in dailyRowDict)
            {
                if (dailyRowDict[entry.Key].Count > 0)
                {
                    for (int i = 0; i < dailyRowDict[entry.Key].Count; i++)
                    {
                        dailyRowDict[entry.Key][i].GetFailureRate();
                    }
                }
            }

            foreach (KeyValuePair<string, List<rowInfo>> entry in wigigRowDict)
            {
                if (wigigRowDict[entry.Key].Count > 0)
                {
                    for (int i = 0; i < wigigRowDict[entry.Key].Count; i++)
                    {
                        wigigRowDict[entry.Key][i].GetFailureRate();
                    }
                }
            }

            p.logger("[DataRetrieval.cs] Completed calculating the failure rates.");
        }

        /*
         * Function    : gatherSpreadsheet()
         * Description : Gets data from excel spreadsheet and store it into the dictionaries accordingly.
         * Params      : (ref Dictionary<string, List<rowInfo>>) weeklyRowDict
		 *				 (ref Dictionary<string, List<rowInfo>>) dailyRowDict
         *				 (ref Dictionary<string, List<rowInfo>>) wigigRowDict
		 *				 (string) dateString
         * Return      :
         */

        public void GatherSpreadsheet(ref Dictionary<string, List<rowInfo>> weeklyRowDict, ref Dictionary<string, List<rowInfo>> dailyRowDict, ref Dictionary<string, List<rowInfo>> wigigRowDict, string dateString)
        {
            p.logger("[DataRetrieval.cs] Gathering data from spreadsheets");
            string[] cx = Directory.GetFiles(ConfigurationManager.AppSettings["reportPath"], dateString + "_CX*");
            string[] wz = Directory.GetFiles(ConfigurationManager.AppSettings["reportPath"], dateString + "_WZ*");
            string[] cx_wigig = Directory.GetFiles(ConfigurationManager.AppSettings["reportPath"], dateString + "_WIGIG_CX*");
            string[] fs = Directory.GetFiles(ConfigurationManager.AppSettings["reportPath"], dateString + "_FS*");
            string[] odmList = { "CX", "WZ", "FS" };

            foreach (string ODM in odmList)
            {
                var application = new Microsoft.Office.Interop.Excel.Application();
                application.Visible = false;
                p.logger("[DataRetrieval.cs] Gathering data from " + ODM);
                // Open up the Excel sheet for the current ODM
                Workbooks workbooks = application.Workbooks;
                Workbook workbook = null;

                if (ODM.Equals("CX") && cx.Length > 0)
                {
                    workbook = workbooks.Open(cx[0], ReadOnly: false);
                }
                else if (ODM.Equals("WZ") && wz.Length > 0)
                {
                    workbook = workbooks.Open(wz[0], ReadOnly: false);
                }
                else if (ODM.Equals("FS") && fs.Length > 0)
                {
                    workbook = workbooks.Open(fs[0], ReadOnly: false);
                }

                Process pid = GetProcess(application);

                if (workbook == null)
                {
                    continue;
                }
                var worksheets = workbook.Worksheets;
                // Iterate over the products' worksheets to insert data
                for (int i = 1; i <= worksheets.Count; i++)
                {
                    Worksheet worksheet = (Worksheet)workbook.Worksheets[i];
                    var rowCountTemp = worksheet.UsedRange;
                    var rowCountTemp2 = rowCountTemp.Rows;
                    int rowCount = rowCountTemp2.Count - 1;
                    bool insertDailyData = false;
                    int titleCount = 0;

                    if (rowCountTemp != null)
                    {
                        Marshal.ReleaseComObject(rowCountTemp);
                    }

                    if (rowCountTemp2 != null)
                    {
                        Marshal.ReleaseComObject(rowCountTemp2);
                    }

                    // Start at the 2nd position because the first position contains the column titles.
                    // Iterate over all of the used rows to get data and save data to dictionaries.
                    for (int j = 2; j < (rowCount + 2); j++)
                    {
                        int k = 0;
                        var rangeTemp = worksheet.UsedRange;
                        Range range1 = rangeTemp.Rows[j];
                        rowInfo row = new rowInfo();

                        if (rangeTemp != null)
                        {
                            Marshal.ReleaseComObject(rangeTemp);
                        }

                        foreach (Range r in range1.Cells)
                        {
                            string value = r.Text;

                            if ((string.IsNullOrEmpty(value) || string.IsNullOrWhiteSpace(value)) && (titleCount < 6))
                            {
                                insertDailyData = true;

                                if (value.Equals("Work Year")
                                    || value.Equals("Work Week")
                                    || value.Equals("Work Day")
                                    || value.Equals("Failed Boards")
                                    || value.Equals("Total Boards")
                                    || value.Equals("Failure Rate (%)"))
                                {
                                    titleCount++;
                                }
                            }

                            if (!insertDailyData)
                            {
                                switch (k)
                                {
                                    case 0:
                                        row.WorkYear = value;
                                        break;

                                    case 1:
                                        row.WorkWeek = value;
                                        break;

                                    case 2:
                                        int temp;
                                        Int32.TryParse(value, out temp);
                                        row.FailedBoards = temp;
                                        break;

                                    case 3:
                                        int temp2;
                                        Int32.TryParse(value, out temp2);
                                        row.TotalBoards = temp2;
                                        break;

                                    case 4:
                                        row.FailureRate = value;
                                        break;

                                    case 5:
                                        row.Goal = value;
                                        break;
                                }
                            }
                            else
                            {
                                if (string.IsNullOrEmpty(value)
                                    || string.IsNullOrWhiteSpace(value)
                                    || value.Equals("Work Year")
                                    || value.Equals("Work Week")
                                    || value.Equals("Work Day")
                                    || value.Equals("Failed Boards")
                                    || value.Equals("Total Boards")
                                    || value.Equals("Failure Rate (%)"))
                                {
                                    continue;
                                }
                                else
                                {
                                    switch (k)
                                    {
                                        case 0:
                                            row.WorkYear = value;
                                            break;

                                        case 1:
                                            row.WorkWeek = value;
                                            break;

                                        case 2:
                                            row.WorkDay = value;
                                            break;

                                        case 3:
                                            int temp;
                                            Int32.TryParse(value, out temp);
                                            row.FailedBoards = temp;
                                            break;

                                        case 4:
                                            int temp2;
                                            Int32.TryParse(value, out temp2);
                                            row.TotalBoards = temp2;
                                            break;

                                        case 5:
                                            row.FailureRate = value;
                                            break;
                                    }
                                }
                            }

                            k++;

                            if (r != null)
                            {
                                Marshal.ReleaseComObject(r);
                            }
                        }

                        row.ODM = ODM;

                        if (string.IsNullOrEmpty(row.WorkYear)
                            || string.IsNullOrEmpty(row.WorkWeek)
                            || string.IsNullOrWhiteSpace(row.WorkYear)
                            || string.IsNullOrWhiteSpace(row.WorkWeek))
                        {
                            continue;
                        }
                        if (!insertDailyData)
                        {
                            weeklyRowDict[worksheet.Name].Add(row);
                        }
                        else
                        {
                            dailyRowDict[worksheet.Name].Add(row);
                        }

                        if (range1 != null)
                        {
                            Marshal.ReleaseComObject(range1);
                        }
                    }

                    if (worksheet != null)
                    {
                        Marshal.ReleaseComObject(worksheet);
                    }
                }

                // Clean up the COM objects to remove "EXCEL.EXE" process
                if (worksheets != null)
                {
                    Marshal.ReleaseComObject(worksheets);
                }

                if (workbooks != null)
                {
                    Marshal.ReleaseComObject(workbooks);
                }

                if (workbook != null)
                {
                    workbook.Close(Type.Missing, Type.Missing, Type.Missing);
                    Marshal.ReleaseComObject(workbook);
                }

                if (application != null)
                {
                    application.Quit();
                    Marshal.ReleaseComObject(application);
                }

                pid.Kill();
                p.logger("[DataRetrieval.cs] Finished gathering data for " + ODM);
            }

            if (cx_wigig.Length <= 0)
            {
                return;
            }
            // Gather CX WIGIG data from spreadsheet
            var app = new Microsoft.Office.Interop.Excel.Application();
            app.Visible = false;
            Workbooks wbs = app.Workbooks;
            Workbook wb = wbs.Open(cx_wigig[0], ReadOnly: false);
            Process pidWigig = GetProcess(app);
            p.logger("[DataRetrieval.cs] Gathering data for WIGIG");
            var ws = wb.Worksheets;
            for (int i = 1; i <= ws.Count; i++)
            {
                Worksheet worksheet = (Worksheet)wb.Worksheets[i];
                var rowCountTemp = worksheet.UsedRange;
                var rowCountTemp2 = rowCountTemp.Rows;
                int rowCount = rowCountTemp2.Count - 1;

                if (rowCountTemp != null)
                {
                    Marshal.ReleaseComObject(rowCountTemp);
                }

                if (rowCountTemp2 != null)
                {
                    Marshal.ReleaseComObject(rowCountTemp2);
                }

                for (int j = 2; j < (rowCount + 2); j++)
                {
                    int k = 0;
                    var rangeTemp = worksheet.UsedRange;
                    Range range1 = rangeTemp.Rows[j];
                    rowInfo row = new rowInfo();

                    if (rangeTemp != null)
                    {
                        Marshal.ReleaseComObject(rangeTemp);
                    }

                    foreach (Range r in range1.Cells)
                    {
                        string value = r.Text;

                        switch (k)
                        {
                            case 0:
                                row.WorkYear = value;
                                break;

                            case 1:
                                row.WorkWeek = value;
                                break;

                            case 2:
                                row.WorkDay = value;
                                break;

                            case 3:
                                int temp;
                                Int32.TryParse(value, out temp);
                                row.FailedBoards = temp;
                                break;

                            case 4:
                                int temp2;
                                Int32.TryParse(value, out temp2);
                                row.TotalBoards = temp2;
                                break;

                            case 5:
                                row.FailureRate = value;
                                break;
                        }

                        k++;

                        if (r != null)
                        {
                            Marshal.ReleaseComObject(r);
                        }
                    }

                    row.ODM = "CX";

                    if (string.IsNullOrEmpty(row.WorkYear)
                        || string.IsNullOrEmpty(row.WorkWeek)
                        || string.IsNullOrWhiteSpace(row.WorkYear)
                        || string.IsNullOrWhiteSpace(row.WorkWeek))
                    {
                        continue;
                    }
                    else
                    {
                        wigigRowDict[worksheet.Name].Add(row);
                    }

                    if (range1 != null)
                    {
                        Marshal.ReleaseComObject(range1);
                    }
                }

                if (worksheet != null)
                {
                    Marshal.ReleaseComObject(worksheet);
                }
            }

            if (ws != null)
            {
                Marshal.ReleaseComObject(ws);
            }

            if (wbs != null)
            {
                Marshal.ReleaseComObject(wbs);
            }

            if (wb != null)
            {
                wb.Close(Type.Missing, Type.Missing, Type.Missing);
                Marshal.ReleaseComObject(wb);
            }

            if (app != null)
            {
                app.Quit();
                Marshal.ReleaseComObject(app);
            }

            pidWigig.Kill();
            p.logger("[DataRetrieval.cs] Finished gathering data for WIGIG");
        }
    }
}