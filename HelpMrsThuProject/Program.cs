using HelpMrsThuProject;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
    static int WorkingDaysInMonth(int year, int month, List<DateTime> publicHolidays, DateTime startDate, DateTime endDate)
    {
        // Get the number of days in the month
        int numDays = DateTime.DaysInMonth(year, month);

        // Initialize a counter for working days
        int workingDays = 0;

        // Iterate through each day in the month
        for (int day = 1; day <= numDays; day++)
        {
            DateTime date = new DateTime(year, month, day);
            if (month == startDate.Month && day < startDate.Day)
            {
                continue;
            }

            if (month == endDate.Month && day > endDate.Day)
            {
                continue;
            }

            // Check if the day is a weekday (Monday to Friday) and not a public holiday
            if (date.DayOfWeek != DayOfWeek.Saturday && date.DayOfWeek != DayOfWeek.Sunday && !publicHolidays.Contains(date))
            {
                workingDays++;
            }
        }

        return workingDays;
    }

    static Dictionary<Tuple<int, int>, int> WorkingDaysInDuration(DateTime startDate, DateTime endDate, List<DateTime> publicHolidays)
    {
        // Initialize a dictionary to store the count of working days for each month
        Dictionary<Tuple<int, int>, int> workingDaysPerMonth = new Dictionary<Tuple<int, int>, int>();

        // Iterate through each month within the duration
        while (startDate <= endDate)
        {
            int year = startDate.Year;
            int month = startDate.Month;
            int daysInMonth = DateTime.DaysInMonth(year, month);
            int workingDays = WorkingDaysInMonth(year, month, publicHolidays, startDate, endDate);

            // Exclude start and end dates if they are weekends or public holidays
            if ((startDate.DayOfWeek == DayOfWeek.Saturday || startDate.DayOfWeek == DayOfWeek.Sunday) || publicHolidays.Contains(startDate))
            {
                workingDays--;
            }
            if ((endDate.DayOfWeek == DayOfWeek.Saturday || endDate.DayOfWeek == DayOfWeek.Sunday) || publicHolidays.Contains(endDate))
            {
                workingDays--;
            }
            // Adjust the count if start and end dates are in the same month
            if (startDate.Month == endDate.Month && startDate.Year == endDate.Year)
            {
                if ((startDate.DayOfWeek == DayOfWeek.Saturday || startDate.DayOfWeek == DayOfWeek.Sunday) || publicHolidays.Contains(startDate) || (endDate.DayOfWeek == DayOfWeek.Saturday || endDate.DayOfWeek == DayOfWeek.Sunday) || publicHolidays.Contains(endDate))
                {
                    workingDays = workingDays - (daysInMonth - endDate.Day + startDate.Day) + 1;
                }
                else
                {
                    workingDays = workingDays - (daysInMonth - endDate.Day + startDate.Day);
                }
            }
            // Store the count in the dictionary
            workingDaysPerMonth.Add(new Tuple<int, int>(year, month), workingDays);
            // Move to the next month
            startDate = startDate.AddMonths(1).AddDays(-startDate.Day + 1);
        }

        return workingDaysPerMonth;
    }

        // Example usage:



var excel = GetOpeningExcelFile("testMrsThu", "testData");

var startRow = 8;
List<RangeTimeModel> lstRangeTimeModel = new List<RangeTimeModel>();
while (excel.Item3.Cells[startRow, 3].Value != null && excel.Item3.Cells[startRow, 4].Value != null)
{
    lstRangeTimeModel.Add(new RangeTimeModel
    {
        FromDate = excel.Item3.Cells[startRow, 3].Value,
        ToDate = excel.Item3.Cells[startRow, 4].Value,
        Row = startRow
    });
    startRow++;
}

foreach(var rangeTimeModel in lstRangeTimeModel)
{
    DateTime startDate = rangeTimeModel.FromDate;
    DateTime endDate = rangeTimeModel.ToDate;
    List<DateTime> publicHolidays = new List<DateTime>
    {
        //new DateTime(2024, 5, 1) // International labor day
        // Add more public holidays here if needed
    };
    Dictionary<Tuple<int, int>, int> result = WorkingDaysInDuration(startDate, endDate, publicHolidays);
    Console.WriteLine($"Number of working days from {startDate:MM yyyy dd} tp {endDate:MM yyyy dd}");
    foreach (var entry in result)
    {
        Console.WriteLine($"{new DateTime(entry.Key.Item1, entry.Key.Item2, 1):MM yyyy}: {entry.Value} working days");
    }
}

Console.WriteLine("");


static Tuple<Excel.Workbook, Excel.Worksheet, Excel.Range> GetOpeningExcelFile(string name, string sheetname)
{
    bool wasFoundRunning = false;
    Excel.Workbook workbook = null;
    Excel.Worksheet worksheet = null;
    Excel.Range range = null;
    try
    {
        var xlApp = (Application)Marshal2.GetActiveObject("Excel.Application");
        Excel.Workbooks xlBooks = xlApp.Workbooks;
        foreach(Excel.Workbook xlbook in xlBooks)
        {
            if (xlbook.Name.ToLower().Trim().Contains(name.ToLower().Trim()))
            {
                workbook = xlbook;
                worksheet = (Excel.Worksheet) workbook.Worksheets[sheetname];
                range = worksheet.Cells[1, 1];

            }
        }
        var numBooks = xlBooks.Count;
        wasFoundRunning = true;

        xlApp.Visible = true;
    }
    catch (Exception e)
    {
        //Log.Error(e.Message);
        Console.WriteLine(e.Message);
        wasFoundRunning = false;
    }

    return Tuple.Create(workbook, worksheet, range);
}

