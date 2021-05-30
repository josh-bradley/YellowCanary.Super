using System;
using System.Data.Common;

namespace YellowCanary.Super
{
    public enum Quarter
    {
        Q1,
        Q2,
        Q3,
        Q4
    }

    public record YearQuarter(Quarter quarter, int year);
    
    public static class QuarterUtil
    {
        public static YearQuarter GetDistributionQuarter(DateTime date)
        {
            var quarter = Quarter.Q4;
            if (date.DayOfYear is >= 29 and <= 119)
            {
                quarter = Quarter.Q1;
            }
            
            if (date.DayOfYear is >= 120 and <= 210)
            {
                quarter = Quarter.Q2;
            }
            
            if (date.DayOfYear is >= 211 and <= 302)
            {
                quarter = Quarter.Q3;
            }

            var year = (date.DayOfYear is >= 0 and <= 28) ? date.Year - 1 :  date.Year;
            
            return new(quarter, year);
        }
        
        public static YearQuarter GetQuarter(DateTime date)
        {
            var quarter = Quarter.Q4;
            if (date.DayOfYear is >= 1 and <= 91)
            {
                quarter = Quarter.Q1;
            }
            
            if (date.DayOfYear is >= 92 and <= 182)
            {
                quarter = Quarter.Q2;
            }
            
            if (date.DayOfYear is >= 183 and <= 274)
            {
                quarter = Quarter.Q3;
            }

            return new(quarter, date.Year);
        }
    }
}