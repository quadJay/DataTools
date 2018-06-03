using System;

namespace PVDailyMonitor
{
    internal class rowInfo
    {
        public string WorkYear { get; set; }
        public string WorkWeek { get; set; }
        public string WorkDay { get; set; }
        public int FailedBoards { get; set; }
        public int TotalBoards { get; set; }
        public string FailureRate { get; set; }
        public string Goal { get; set; }
        public string ODM { get; set; }

        /*
         * rowInfo's default constructor
         */

        public rowInfo()
        {
            FailedBoards = 0;
            TotalBoards = 0;
            Goal = "2";
        }

        /*
         * Function    : GetFailureRate()
         * Description : Calculate the failure rate with the formula failed board / total boards
         * Params      :
         * Return      : (string) FailureRate
         */

        public string GetFailureRate()
        {
            if (TotalBoards == 0)
            {
                FailureRate = "0";
            }
            else
            {
                double percentage = (Convert.ToDouble(FailedBoards) / Convert.ToDouble(TotalBoards)) * 100;
                FailureRate = FormatString(percentage);
            }
            return FailureRate;
        }

        /*
         * Function    : FormatString()
         * Description : Converts double variable to a string with two decimal places
         * Params      : (double) percent
         * Return      : (string) s
         */

        public static string FormatString(double percent)
        {
            string s = string.Format("{0:0.00}", percent);

            if (s.EndsWith("00", StringComparison.Ordinal))
            {
                return ((int)percent).ToString();
            }
            else
            {
                return s;
            }
        }
    }
}