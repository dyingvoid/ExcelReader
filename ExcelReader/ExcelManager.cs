using Microsoft.EntityFrameworkCore.Sqlite.Query.Internal;
using OfficeOpenXml;

namespace ExcelReader;

public static class ExcelManager
{
    private static string _borderBackgroundColor = "FF92D050";
    public static List<Schedule> ReadExcel(FileInfo excelFile)
    {
        var weekSchedule = new List<Schedule>();

        using (var package = new ExcelPackage(excelFile))
        {
            var sheets = package.Workbook.Worksheets;

            foreach (var sheet in sheets)
            {
                var borders = FindDaysBorders(sheet);
                var groupsRowNumber = borders[DayOfWeek.Monday].Item1.Row();
                var groups = FindGroupNamesCells(sheet, groupsRowNumber);
                var smt = ParseDay(sheet, borders[DayOfWeek.Monday], groups["ПИ 101"]);
                var weekSch = ParseWeek(sheet, borders, groups["ПИ 101"]);
            }
        }

        return weekSchedule;
    }

    private static Dictionary<DayOfWeek, Dictionary<int, List<string>>> ParseWeek(ExcelWorksheet sheet, 
        Dictionary<DayOfWeek, Tuple<ExcelRange, ExcelRange>> daysBorders, ExcelRange columnCell)
    {
        var dict = new Dictionary<DayOfWeek, Dictionary<int, List<string>>>();

        for (var day = DayOfWeek.Monday; day <= DayOfWeek.Saturday; day++)
        {
            var daySubject = ParseDay(sheet, daysBorders[day], columnCell);
            dict.TryAdd(day, daySubject);
        }
        
        return dict;
    }

    private static Dictionary<int, List<string>> ParseDay(ExcelWorksheet sheet, 
        Tuple<ExcelRange, ExcelRange> yAxisBorders, ExcelRange columnCell)
    {
        var daySchedule = new Dictionary<int, List<string>>();

        var rowStart = yAxisBorders.Item1.Row() + 1;
        var rowEnd = yAxisBorders.Item2.Row();

        var columnOfSubjectNumber = 2;
        for (var row = rowStart; row <= rowEnd; ++row)
        {
            var value = sheet.Cells[row, columnOfSubjectNumber].Value;
            
            if (value != null && int.TryParse(value.ToString(), out var number))
                daySchedule.TryAdd(number, ParseClassCells(sheet, row, columnCell));
        }

        return daySchedule;
    }

    private static List<string> ParseClassCells(ExcelWorksheet sheet, int topBorderRow, ExcelRange columnCell)
    {
        var list = new List<string>();
        
        // Sometimes there is no bottom border in cells, so we count top border of this class time
        // And top border of next class time
        int topBorderCounter = 0;

        // Classes in one class time divided by a thinner border
        // We store style of main border to define iterating range correctly
        var cell = sheet.Cells[topBorderRow, columnCell.Column()];
        string mainBorderStyle = cell.Style.Border.Top.Style.ToString();

        // There may be several subgroups in one group. Exact number is unknown, way without number is not implemented
        for (var column = columnCell.Column(); column <= columnCell.Column() + 1; column++)
        {
            while (true)
            {
                cell = sheet.Cells[topBorderRow, column];
            
                if (cell.Style.Border.Top.Style.ToString() == mainBorderStyle)
                    topBorderCounter++;
                if (topBorderCounter > 1)
                    break;

                if (cell.Value != null)
                {
                    string subject = cell.Value.ToString() + sheet.Cells[topBorderRow + 1, column].Value;
                    if(column == columnCell.Column())
                        list.Add("1 группа: " + subject);
                    else
                        list.Add("2 группа: " + subject);

                    topBorderRow++;
                }

                if (cell.Style.Border.Bottom.Style.ToString() == mainBorderStyle)
                    break;

                topBorderRow++;
            }
        }

        return list;
    }

    private static Dictionary<string, ExcelRange> FindGroupNamesCells(ExcelWorksheet sheet, int groupsRowNumber)
    {
        var groups = new Dictionary<string, ExcelRange>();
        var cells = 
            FindAllCellsColumns(sheet, groupsRowNumber, cell => cell.Value != null);
        foreach (var cell in cells)
        {
            groups.Add((string)cell.Value, cell);
        }

        return groups;
    }

    private static Dictionary<DayOfWeek, Tuple<ExcelRange, ExcelRange>> FindDaysBorders(ExcelWorksheet sheet)
    {
        var dayBorders = new Dictionary<DayOfWeek, Tuple<ExcelRange, ExcelRange>>();
        
        // Better to check that style is not None, than this
        var list = FindDaySeparatorRows(sheet);

        var startDay = DayOfWeek.Monday;
        int i = 0;
        for (var day = startDay; day <= DayOfWeek.Saturday; ++day)
        {
            dayBorders.Add(day, Tuple.Create(list[i], list[i+1]));
            i++;
        }
        
        return dayBorders;
    }

    private static List<ExcelRange?> FindDaySeparatorRows(ExcelWorksheet sheet)
    {
        var cellsWithBorder = FindAllCellsRows(sheet, sheet.Dimension.Start.Column,
            range => range.Style.Border.Bottom.Style.ToString() != "None");
        var list = FindAllCellsRows(sheet, sheet.Dimension.Start.Column,
            range => range.Style.Fill.BackgroundColor.Rgb == _borderBackgroundColor);

        list.Insert(0, cellsWithBorder.First());
        list.Add(cellsWithBorder.Last());
        
        return list;
    }

    private static List<ExcelRange?> FindAllCellsRows(ExcelWorksheet sheet, int column, Predicate<ExcelRange> predicate)
    {
        var list = new List<ExcelRange?>();

        for (var row = sheet.Dimension.Start.Row; row <= sheet.Dimension.End.Row; ++row)
        {
            var cell = sheet.Cells[row, column];
            
            if (predicate(cell))
                list.Add(cell);
        }

        return list;
    }

    private static List<ExcelRange?> FindAllCellsColumns(ExcelWorksheet sheet, int row, Predicate<ExcelRange> predicate)
    {
        var list = new List<ExcelRange?>();

        for (var column = sheet.Dimension.Start.Column; column <= sheet.Dimension.End.Column; ++column)
        {
            var cell = sheet.Cells[row, column];
            
            if(predicate(cell))
                list.Add(cell);
        }

        return list;
    }
}