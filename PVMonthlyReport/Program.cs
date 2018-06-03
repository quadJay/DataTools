using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace PVMonthlyReport
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            try
            {
                Outlook.Application outlookApp = new Outlook.Application();

                string reportPath = ConfigurationManager.AppSettings.Get(@"reportPath");
                string reportDBPath = ConfigurationManager.AppSettings.Get(@"reportDBPath");
                string[] emails = ConfigurationManager.AppSettings.Get(@"emails").Split(',');

                DateTime currDate = DateTime.Now.AddMonths(-1);
                DateTime lastDate = new DateTime(Int32.Parse(currDate.ToString("yyyy")), Int32.Parse(currDate.ToString("MM")), 1);
                lastDate = lastDate.AddMonths(1).AddDays(-1);

                string currDateString = lastDate.ToString("yyyyMMdd");

                HashSet<string> attachments = new HashSet<string>();

                DirectoryInfo reportDirInfo = new DirectoryInfo(reportPath);
                IEnumerable<FileInfo> xlsxs;
                xlsxs = reportDirInfo.GetFiles("*.xlsx", SearchOption.TopDirectoryOnly).Where(x => (x.Name.Contains(currDateString)));
                foreach (FileInfo xlsx in xlsxs)
                {
                    string outFile = xlsx.FullName.Replace(currDateString, lastDate.ToString(@"MMMMyyyy"));
                    string[] substrs = xlsx.Name.Split(new string[] { "_" }, StringSplitOptions.RemoveEmptyEntries);
                    outFile = outFile.Replace(reportPath, reportDBPath);
                    outFile = outFile.Replace(substrs[substrs.Count() - 1], @"TX_PV_Monthly_Report.xlsx");
                    File.Copy(xlsx.FullName, outFile, true);
                    attachments.Add(outFile);
                }
                string subject = lastDate.ToString("y") + " TX PV Report";
                string body = "Attached to this email are reports for Gemtek, Azurewave, and Brazil Flextronics. The reports contains information for " + lastDate.ToString("y") + "."
                                + "\nThese reports are up to date and are all based on raw data from ODMs.\n\nPlease let me know if any of the data is incorrect or questionable.\n\nThanks.";
                send_email(ref outlookApp, emails, subject, body, ref attachments);
                Console.WriteLine("done with email");
            }

            // Simple error handler.
            catch (Exception e)
            {
                string logPath = ConfigurationManager.AppSettings.Get(@"logPath");
                Directory.CreateDirectory(logPath);
                string timeStamp = DateTime.Now.ToString("yyyyMMddHHmmssfff");

                string[] msg = { "" };
                msg[0] = @"Message:" + e.Message + @" StackTrace:" + e.StackTrace + @" TargetSite:" + e.TargetSite + @" Source:" + e.Source;
                File.AppendAllLines(Path.Combine(logPath, timeStamp + ".txt"), msg);
            }
        }

        public static void send_email(ref Outlook.Application outlookApp, string[] emails, string subject, string body, ref HashSet<string> attachments)
        {
            Outlook.MailItem mail = outlookApp.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;

            mail.Subject = subject;
            mail.Body = body;
            Outlook.AddressEntry currentUser = outlookApp.Session.CurrentUser.AddressEntry;
            if (currentUser.Type == "EX")
            {
                foreach (string email in emails)
                {
                    mail.Recipients.Add(email);
                }
                mail.Recipients.ResolveAll();
                foreach (string attachment in attachments)
                {
                    mail.Attachments.Add(attachment, Outlook.OlAttachmentType.olByValue, Type.Missing, Type.Missing);
                }
                mail.Send();
            }
        }
    }
}