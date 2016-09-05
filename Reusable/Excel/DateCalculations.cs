using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Reusable.Excel
{
    public static class Date
    {
        public static DateTime CalculatePriorWeekdayDate(DateTime day)
        {
            //Set day - 1 on weeked, move back to friday,
            //else move back 1

            switch (day.DayOfWeek)
            {
                case DayOfWeek.Monday:
                    return day.Subtract(new System.TimeSpan(3, 0, 0, 0));
                case DayOfWeek.Sunday:
                    return day.Subtract(new System.TimeSpan(2, 0, 0, 0));
                default:
                    return day.Subtract(new System.TimeSpan(1, 0, 0, 0));
            }
        }
    }
}
