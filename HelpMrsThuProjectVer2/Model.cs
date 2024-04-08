using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HelpMrsThuProjectVer2
{
    internal class Model
    {
    }

    public class RangeTimeModel
    {
        public DateTime FromDate { set; get; }
        public DateTime ToDate { set; get; }
        public int Row { set; get; }
    }

    public enum MonthColumn
    {
        January = 1,
        February = 2,
        March = 3,
        April = 4,
        May = 5,
        June = 6,
        July = 7,
        August = 8,
        September = 9,
        October = 10,
        November = 11,
        December = 12

    }
}
