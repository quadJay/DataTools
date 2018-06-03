using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;

namespace PVDailyMonitor
{
    internal class Report
    {
        [DllImport("user32.dll")]
        private static extern int GetWindowThreadProcessId(int hWnd, out int lpdwProcessId);

        private string fileName = ConfigurationManager.AppSettings["reportPath"];
        private string topLeft;
        private string bottomRight;
        private const string weeklyGraphTitle = "TX Power Failures - FOQM - (Weekly)";
        private const string weeklyXAxis = "Workweek";
        private const string weeklyYAxis = "Failure Rate (%)";
        private const string dailyGraphTitle = "TX Power Failures - FOQM - (Daily)";
        private const string dailyXAxis = "Workday/Workweek";
        private const string dailyYAxis = "Failure Rate (%)";
        private const string wigigGraphTitle = "TX Power Failures - WGFOQM - (Daily)";
        private const string wigigXAxis = "Workday/Workweek";
        private const string wigigYAxis = "Failure Rate (%)";
        private string[] odmList = { "CX", "WZ", "FS" };

        private string[] weeklyColumnNames =
        {
            "Work Year", "Work Week", "Failed Boards", "Total Boards", "Failure Rate (%)", "Threshold (%)"
        };

        private string[] dailyColumnNames =
        {
            "Work Year", "Work Week", "Work Day", "Failed Boards", "Total Boards", "Failure Rate (%)"
        };

        private string[] columnLetters =
        {
            "A", "B", "C", "D", "E", "F"
        };

        private static Program p;

        public Report(Program program)
        {
            topLeft = "";
            bottomRight = "";
            p = program;
        }

        /*
         * Function    : GetProcess()
         * Description : Returns a Process variable of an Application instance.
         * Params      : (Application) app
         * Return      :
         */

        private Process GetProcess(Application app)
        {
            int id;
            GetWindowThreadProcessId(app.Hwnd, out id);
            return Process.GetProcessById(id);
        }

        /*
         * Function    : CreateReport()
         * Description : Generates an excel spreadsheet then inserts rows of data into it.
         * Params      : (Dictionary<string, List<rowInfo>>) weeklyRowDict
		 *				 (Dictionary<string, List<rowInfo>>) dailyRowDict
		 *				 (Dictionary<string, List<rowInfo>>) wigigRowDict
         *               (string) reportPath
         *               (string) xmlPath
         * Return      :
         */

        public void CreateReport(ref Dictionary<string, List<rowInfo>> weeklyRowDict, ref Dictionary<string, List<rowInfo>> dailyRowDict, ref Dictionary<string, List<rowInfo>> wigigRowDict, string reportPath, string xmlPath)
        {
            // The variable fileName is the full path to the excel report file.
            string removeString = ConfigurationManager.AppSettings["xmlPath"];
            xmlPath = xmlPath.Remove(xmlPath.IndexOf(removeString, StringComparison.Ordinal), removeString.Length);
            string num = xmlPath.Substring(xmlPath.Length - 3);
            string dateString = xmlPath.Remove(xmlPath.IndexOf(num, StringComparison.Ordinal), num.Length);

            foreach (string odm in odmList)
            {
                fileName = reportPath + dateString + odm + "_" + num + ".xlsx";
                string[] productList = ConfigurationManager.AppSettings[odm].Split(',');
                var application = new Application();
                application.SheetsInNewWorkbook = productList.Length;
                application.Visible = false;
                Workbooks workbooks = application.Workbooks;
                Workbook workbook = workbooks.Add(Missing.Value);

                // Sort every products' rows by year, week, and day if applicable.
                foreach (var product in productList)
                {
                    weeklyRowDict[product] = weeklyRowDict[product].OrderBy(x => x.WorkYear).ThenBy(x => x.WorkWeek).ToList();
                    dailyRowDict[product] = dailyRowDict[product].OrderBy(x => x.WorkYear).ThenBy(x => x.WorkWeek).ThenBy(x => x.WorkDay).ToList();
                }

                var worksheets = workbook.Sheets;
                for (int i = 1; i <= worksheets.Count; i++)
                {
                    Worksheet worksheet = (Worksheet)workbook.Worksheets[i];
                    worksheet.Name = productList[i - 1];

                    // Set column titles for weekly data
                    Range aRange = null;
                    for (int j = 0; j < weeklyColumnNames.Length; j++)
                    {
                        aRange = worksheet.get_Range(columnLetters[j] + "2", columnLetters[j] + "2");
                        Object[] args = new object[1];
                        args[0] = weeklyColumnNames[j];
                        var aRangeType = aRange.GetType();
                        var aRangeCols = aRange.EntireColumn;
                        var aRangeInterior = aRange.Interior;
                        aRangeType.InvokeMember("Value", BindingFlags.SetProperty, null, aRange, args);
                        aRangeCols.AutoFit();
                        aRangeInterior.Color = ColorTranslator.ToOle(Color.Aqua);

                        if (aRangeCols != null)
                        {
                            Marshal.ReleaseComObject(aRangeCols);
                        }

                        if (aRangeInterior != null)
                        {
                            Marshal.ReleaseComObject(aRangeInterior);
                        }

                        if (aRange != null)
                        {
                            Marshal.ReleaseComObject(aRange);
                        }
                    }

                    // Insert data into the Excel spreadsheet
                    int rowCount = 2;
                    if (weeklyRowDict[productList[i - 1]].Count > 0)
                    {
                        // Iterating through all of the rows for the product
                        for (int k = 0; k < weeklyRowDict[productList[i - 1]].Count; k++)
                        {
                            // Iterating through all of the columns of the row for the product
                            if (weeklyRowDict[productList[i - 1]][k].ODM.Equals(odm, StringComparison.CurrentCultureIgnoreCase))
                            {
                                rowCount++;
                                for (int l = 0; l < columnLetters.Length; l++)
                                {
                                    string currPos = rowCount.ToString();
                                    aRange = worksheet.get_Range(columnLetters[l] + currPos, columnLetters[l] + currPos);

                                    Object[] args = new object[1];

                                    if (l == 0)
                                    {
                                        args[0] = weeklyRowDict[productList[i - 1]][k].WorkYear;
                                    }
                                    else if (l == 1)
                                    {
                                        args[0] = weeklyRowDict[productList[i - 1]][k].WorkWeek;
                                    }
                                    else if (l == 2)
                                    {
                                        args[0] = weeklyRowDict[productList[i - 1]][k].FailedBoards;
                                    }
                                    else if (l == 3)
                                    {
                                        args[0] = weeklyRowDict[productList[i - 1]][k].TotalBoards;
                                    }
                                    else if (l == 4)
                                    {
                                        args[0] = weeklyRowDict[productList[i - 1]][k].FailureRate;
                                    }
                                    else if (l == 5)
                                    {
                                        args[0] = weeklyRowDict[productList[i - 1]][k].Goal;
                                    }

                                    var aRangeType = aRange.GetType();
                                    aRangeType.InvokeMember("Value", BindingFlags.SetProperty, null, aRange, args);

                                    if (aRange != null)
                                    {
                                        Marshal.ReleaseComObject(aRange);
                                    }
                                }
                            }
                        }
                    }

                    // Add 20 spaces for the gap between the two data analysis
                    rowCount += 20;

                    // Set column titles for daily data
                    for (int j = 0; j < dailyColumnNames.Length; j++)
                    {
                        aRange = worksheet.get_Range(columnLetters[j] + rowCount, columnLetters[j] + rowCount);
                        Object[] args = new object[1];
                        args[0] = dailyColumnNames[j];
                        var aRangeType = aRange.GetType();
                        var aRangeCols = aRange.EntireColumn;
                        var aRangeInterior = aRange.Interior;
                        aRangeType.InvokeMember("Value", BindingFlags.SetProperty, null, aRange, args);
                        aRangeCols.AutoFit();
                        aRangeInterior.Color = ColorTranslator.ToOle(Color.Aqua);

                        if (aRangeCols != null)
                        {
                            Marshal.ReleaseComObject(aRangeCols);
                        }

                        if (aRangeInterior != null)
                        {
                            Marshal.ReleaseComObject(aRangeInterior);
                        }

                        if (aRange != null)
                        {
                            Marshal.ReleaseComObject(aRange);
                        }
                    }

                    if (dailyRowDict[productList[i - 1]].Count > 0)
                    {
                        // Iterating through all of the rows for the product
                        for (int k = 0; k < dailyRowDict[productList[i - 1]].Count; k++)
                        {
                            // Iterating through all of the columns of the row for the product
                            if (dailyRowDict[productList[i - 1]][k].ODM.Equals(odm, StringComparison.CurrentCultureIgnoreCase))
                            {
                                rowCount++;
                                for (int l = 0; l < columnLetters.Length; l++)
                                {
                                    string currPos = rowCount.ToString();
                                    aRange = worksheet.get_Range(columnLetters[l] + currPos, columnLetters[l] + currPos);

                                    Object[] args = new object[1];

                                    if (l == 0)
                                    {
                                        args[0] = dailyRowDict[productList[i - 1]][k].WorkYear;
                                    }
                                    else if (l == 1)
                                    {
                                        args[0] = dailyRowDict[productList[i - 1]][k].WorkWeek;
                                    }
                                    else if (l == 2)
                                    {
                                        args[0] = dailyRowDict[productList[i - 1]][k].WorkDay;
                                    }
                                    else if (l == 3)
                                    {
                                        args[0] = dailyRowDict[productList[i - 1]][k].FailedBoards;
                                    }
                                    else if (l == 4)
                                    {
                                        args[0] = dailyRowDict[productList[i - 1]][k].TotalBoards;
                                    }
                                    else if (l == 5)
                                    {
                                        args[0] = dailyRowDict[productList[i - 1]][k].FailureRate;
                                    }

                                    var aRangeType = aRange.GetType();
                                    aRange.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, aRange, args);

                                    if (aRange != null)
                                    {
                                        Marshal.ReleaseComObject(aRange);
                                    }
                                }
                            }
                        }
                    }

                    if (worksheet != null)
                    {
                        Marshal.ReleaseComObject(worksheet);
                    }

                    if (aRange != null)
                    {
                        Marshal.ReleaseComObject(aRange);
                    }
                }

                application.DisplayAlerts = false;
                workbook.SaveAs(fileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing);

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

                Process pid = GetProcess(application);

                if (application != null)
                {
                    application.Quit();
                    Marshal.ReleaseComObject(application);
                }

                pid.Kill();
            }

            // Create spreadsheet for WIGIG products
            fileName = reportPath + dateString + "WIGIG_CX" + "_" + num + ".xlsx";
            string[] wigigProductList = ConfigurationManager.AppSettings["CX_WIGIG"].Split(',');
            var app = new Application();
            app.SheetsInNewWorkbook = wigigProductList.Length;
            app.Visible = false;
            Workbooks wbs = app.Workbooks;
            Workbook wb = wbs.Add(Missing.Value);

            // Sort every WIGIG products's rows by year, week, and day if applicable
            foreach (var product in wigigProductList)
            {
                wigigRowDict[product] = wigigRowDict[product].OrderBy(x => x.WorkYear).ThenBy(x => x.WorkWeek).ThenBy(x => x.WorkDay).ToList();
            }

            var ws = wb.Sheets;
            for (int i = 1; i <= ws.Count; i++)
            {
                Worksheet worksheet = (Worksheet)wb.Worksheets[i];
                worksheet.Name = wigigProductList[i - 1];

                // Set column titles for wigig daily data
                Range aRange = null;
                for (int j = 0; j < dailyColumnNames.Length; j++)
                {
                    aRange = worksheet.get_Range(columnLetters[j] + "2", columnLetters[j] + "2");
                    Object[] args = new object[1];
                    args[0] = dailyColumnNames[j];
                    var aRangeType = aRange.GetType();
                    var aRangeCols = aRange.EntireColumn;
                    var aRangeInterior = aRange.Interior;
                    aRangeType.InvokeMember("Value", BindingFlags.SetProperty, null, aRange, args);
                    aRangeCols.AutoFit();
                    aRangeInterior.Color = ColorTranslator.ToOle(Color.Aqua);

                    if (aRangeCols != null)
                    {
                        Marshal.ReleaseComObject(aRangeCols);
                    }

                    if (aRangeInterior != null)
                    {
                        Marshal.ReleaseComObject(aRangeInterior);
                    }

                    if (aRange != null)
                    {
                        Marshal.ReleaseComObject(aRange);
                    }
                }

                // Insert data into the Excel Spreadsheet
                int rowCount = 2;
                if (wigigRowDict[wigigProductList[i - 1]].Count > 0)
                {
                    // Iterating through all of the rows for the product
                    for (int k = 0; k < wigigRowDict[wigigProductList[i - 1]].Count; k++)
                    {
                        // Iterating through all of the columns of the row for the product
                        rowCount++;
                        for (int l = 0; l < columnLetters.Length; l++)
                        {
                            string currPos = rowCount.ToString();
                            aRange = worksheet.get_Range(columnLetters[l] + currPos, columnLetters[l] + currPos);

                            Object[] args = new object[1];

                            if (l == 0)
                            {
                                args[0] = wigigRowDict[wigigProductList[i - 1]][k].WorkYear;
                            }
                            else if (l == 1)
                            {
                                args[0] = wigigRowDict[wigigProductList[i - 1]][k].WorkWeek;
                            }
                            else if (l == 2)
                            {
                                args[0] = wigigRowDict[wigigProductList[i - 1]][k].WorkDay;
                            }
                            else if (l == 3)
                            {
                                args[0] = wigigRowDict[wigigProductList[i - 1]][k].FailedBoards;
                            }
                            else if (l == 4)
                            {
                                args[0] = wigigRowDict[wigigProductList[i - 1]][k].TotalBoards;
                            }
                            else if (l == 5)
                            {
                                args[0] = wigigRowDict[wigigProductList[i - 1]][k].FailureRate;
                            }

                            var aRangeType = aRange.GetType();
                            aRange.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, aRange, args);

                            if (aRange != null)
                            {
                                Marshal.ReleaseComObject(aRange);
                            }
                        }
                    }
                }

                if (worksheet != null)
                {
                    Marshal.ReleaseComObject(worksheet);
                }

                if (aRange != null)
                {
                    Marshal.ReleaseComObject(aRange);
                }
            }

            app.DisplayAlerts = false;
            wb.SaveAs(fileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing);

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

            Process pidWigig = GetProcess(app);

            if (app != null)
            {
                app.Quit();
                Marshal.ReleaseComObject(app);
            }

            pidWigig.Kill();
            p.logger("[Report.cs] Compeleted inserting data into excel spreadsheet");
        }

        /*
         * Function    : CreateGraph()
         * Description : Create graphs based on the row of data in the saved excel spreadsheet.
         * Params      : (Dictionary<string, List<rowInfo>>) weeklyRowDict
		 *				 (Dictionary<string, List<rowInfo>>) dailyRowDict
		 *				 (Dictionary<string, List<rowInfo>>) wigigRowDict
         *               (string) reportPath
         *               (string) xmlPath
         * Return      :
         */

        public void CreateGraph(ref Dictionary<string, List<rowInfo>> weeklyRowDict, ref Dictionary<string, List<rowInfo>> dailyRowDict, ref Dictionary<string, List<rowInfo>> wigigRowDict, string reportPath, string xmlPath)
        {
            p.logger("[Report.cs] Creating all of the graphs for the Excel spreadsheets");

            string removeString = ConfigurationManager.AppSettings["xmlPath"];
            xmlPath = xmlPath.Remove(xmlPath.IndexOf(removeString, StringComparison.Ordinal), removeString.Length);
            string num = xmlPath.Substring(xmlPath.Length - 3);
            string dateString = xmlPath.Remove(xmlPath.IndexOf(num, StringComparison.Ordinal), num.Length);

            foreach (string odm in odmList)
            {
                string[] productList = ConfigurationManager.AppSettings[odm].Split(',');
                fileName = reportPath + dateString + odm + "_" + num + ".xlsx";

                // Open the excel worksheets
                var application = new Application();
                application.Visible = false;
                var workbooks = application.Workbooks;
                var workbook = workbooks.Open(fileName, ReadOnly: false);

                var worksheets = workbook.Sheets;
                for (int i = 1; i <= worksheets.Count; i++)
                {
                    int rowCount = 2;
                    foreach (var row in weeklyRowDict[productList[i - 1]])
                    {
                        if (row.ODM.Equals(odm))
                        {
                            rowCount++;
                        }
                    }

                    int rowCountDaily = rowCount + 20;
                    foreach (var row in dailyRowDict[productList[i - 1]])
                    {
                        if (row.ODM.Equals(odm))
                        {
                            rowCountDaily++;
                        }
                    }

                    var worksheet = workbook.Worksheets[i] as Worksheet;

                    // Add chart
                    var charts = worksheet.ChartObjects() as ChartObjects;
                    var chartObject = charts.Add(400, 30, 300, 300) as ChartObject;
                    var chart = chartObject.Chart;
                    var chartObjectDaily = charts.Add(400, (rowCount + 20) * 14, 600, 300) as ChartObject;
                    var chartDaily = chartObjectDaily.Chart;

                    // Set chart range. Getting Columns E and F for failure rate and threshold respectively.
                    topLeft = "E3";
                    bottomRight = "E" + rowCount.ToString();
                    string topLeftDaily = "F" + (rowCount + 21).ToString();
                    string bottomRightDaily = "F" + rowCountDaily.ToString();

                    Range range;
                    Range rangeDaily;

                    if (rowCount == 3)
                    {
                        range = worksheet.get_Range(topLeft, topLeft);
                    }
                    else
                    {
                        range = worksheet.get_Range(topLeft, bottomRight);
                    }

                    rangeDaily = worksheet.get_Range(topLeftDaily, bottomRightDaily);

                    chart.SetSourceData(range);
                    chartDaily.SetSourceData(rangeDaily);

                    // Set chart properties
                    chart.ChartType = XlChartType.xlLineMarkers;
                    chart.ChartArea.Border.Color = ColorTranslator.ToOle(Color.Blue);
                    chart.Legend.Border.Color = ColorTranslator.ToOle(Color.Blue);
                    chart.ChartWizard(Source: range,
                        Title: weeklyGraphTitle,
                        CategoryTitle: weeklyXAxis,
                        ValueTitle: weeklyYAxis);

                    //chartDaily.ChartTitle.Text = dailyGraphTitle;
                    chartDaily.ChartType = XlChartType.xlLineMarkers;
                    chartDaily.ChartArea.Border.Color = ColorTranslator.ToOle(Color.Blue);
                    chartDaily.Legend.Border.Color = ColorTranslator.ToOle(Color.Blue);
                    chartDaily.ChartWizard(Source: range,
                        Title: dailyGraphTitle,
                        CategoryTitle: dailyXAxis,
                        ValueTitle: dailyYAxis);

                    // This is where we setting the work week for the X-axis
                    var xAxisValues = new List<string>();

                    foreach (var row in weeklyRowDict[productList[i - 1]])
                    {
                        if (row.ODM.Equals(odm))
                        {
                            xAxisValues.Add(row.WorkWeek);
                        }
                    }

                    if (xAxisValues.Count > 0)
                    {
                        Series series1 = (Series)chart.SeriesCollection(1);
                        series1.XValues = xAxisValues.ToArray();
                        series1.Name = "Failure Rate";

                        if (rowCount == 3)
                        {
                            SeriesCollection seriesCollection = chart.SeriesCollection();
                            Series series2 = seriesCollection.NewSeries();
                            series2.Values = worksheet.get_Range("F3", "F3");
                            series2.Name = "Threshold";
                        }
                        else
                        {
                            SeriesCollection seriesCollection = chart.SeriesCollection();
                            Series series2 = seriesCollection.NewSeries();
                            series2.Values = worksheet.get_Range("F3", "F" + rowCount.ToString());
                            series2.Name = "Threshold";
                        }

                        p.logger("[Report.cs] There's data for " + productList[i - 1]);
                    }
                    else
                    {
                        p.logger("[Report.cs] No data for " + productList[i - 1]);
                    }

                    Series seriesDaily1 = (Series)chartDaily.SeriesCollection(1);
                    seriesDaily1.Values = worksheet.get_Range(topLeftDaily, bottomRightDaily);
                    seriesDaily1.XValues = worksheet.get_Range("B" + (rowCount + 21).ToString(), "C" + rowCountDaily.ToString());
                    seriesDaily1.Name = "Failure Rate";

                    if (range != null)
                    {
                        Marshal.ReleaseComObject(range);
                    }

                    if (rangeDaily != null)
                    {
                        Marshal.ReleaseComObject(rangeDaily);
                    }
                }

                // Save
                application.DisplayAlerts = false;
                workbook.Save();

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

                Process pid = GetProcess(application);

                if (application != null)
                {
                    application.Quit();
                    Marshal.ReleaseComObject(application);
                }

                pid.Kill();
            }

            // Create graphs for CX WIGIG products
            string[] wigigProductList = ConfigurationManager.AppSettings["CX_WIGIG"].Split(',');
            fileName = reportPath + dateString + "WIGIG_CX" + "_" + num + ".xlsx";

            var app = new Application();
            app.Visible = false;
            var wbs = app.Workbooks;
            var wb = wbs.Open(fileName, ReadOnly: false);

            var ws = wb.Sheets;
            for (int i = 1; i <= ws.Count; i++)
            {
                int rowCount = 2;
                foreach (var row in wigigRowDict[wigigProductList[i - 1]])
                {
                    rowCount++;
                }

                var worksheet = wb.Worksheets[i] as Worksheet;

                // Add chart
                var charts = worksheet.ChartObjects() as ChartObjects;
                var chartObject = charts.Add(400, 30, 300, 300) as ChartObject;
                var chart = chartObject.Chart;

                // Set chart range. Getting Columns E and F for failure rate and threshold respectively.
                topLeft = "F3";
                bottomRight = "F" + rowCount.ToString();

                Range range;

                if (rowCount == 3)
                {
                    range = worksheet.get_Range(topLeft, topLeft);
                }
                else
                {
                    range = worksheet.get_Range(topLeft, bottomRight);
                }

                chart.SetSourceData(range);

                // Set chart properties
                chart.ChartType = XlChartType.xlLineMarkers;
                chart.ChartArea.Border.Color = ColorTranslator.ToOle(Color.Blue);
                chart.Legend.Border.Color = ColorTranslator.ToOle(Color.Blue);
                chart.ChartWizard(Source: range,
                    Title: wigigGraphTitle,
                    CategoryTitle: wigigXAxis,
                    ValueTitle: wigigYAxis);

                // This is where we setting the work week for the X-axis
                var xAxisValues = new List<string>();

                foreach (var row in wigigRowDict[wigigProductList[i - 1]])
                {
                    xAxisValues.Add(row.WorkWeek);
                }

                if (xAxisValues.Count > 0)
                {
                    Series series1 = (Series)chart.SeriesCollection(1);
                    series1.Values = worksheet.get_Range(topLeft, bottomRight);
                    series1.XValues = worksheet.get_Range("B" + rowCount.ToString(), "C" + rowCount.ToString());
                    series1.Name = "Failure Rate";
                    p.logger("[Report.cs] There's data for " + wigigProductList[i - 1]);
                }
                else
                {
                    p.logger("[Report.cs] No data for " + wigigProductList[i - 1]);
                }

                if (range != null)
                {
                    Marshal.ReleaseComObject(range);
                }
            }

            // Save
            app.DisplayAlerts = false;
            wb.Save();

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

            Process pidWigig = GetProcess(app);

            if (app != null)
            {
                app.Quit();
                Marshal.ReleaseComObject(app);
            }

            pidWigig.Kill();

            p.logger("[DataRetrieval.cx] Finished gathering data from spreadsheets");
        }
    }
}