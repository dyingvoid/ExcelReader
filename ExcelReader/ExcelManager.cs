using OfficeOpenXml;
using OfficeOpenXml.Drawing.Vml;

namespace ExcelReader;

public static class ExcelManager
{
    // Group can have 2 subjects at the same time for subgroups
    private static int _numberOfColumnsForGroup = 2;
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
            }
        }

        return weekSchedule;
    }

    private static List<Subject> ParseDay(ExcelWorksheet sheet,Tuple<ExcelRange, ExcelRange> verticalBorders, ExcelRange columnCell)
    {
        var list = new List<Subject>();

        var rowStart = verticalBorders.Item1.Row() + 1;
        var rowEnd = verticalBorders.Item2.Row();
        var columnStart = columnCell.Column();
        var columnEnd = columnStart + 1;

        var columnOfSubjectNumber = 2;
        for (var row = rowStart; row <= rowEnd; ++row)
        {
            if (int.TryParse(sheet.Cells[row, columnOfSubjectNumber].Value.ToString(), out var number))
            {
                ParseClassCells(sheet, row, columnCell);
            }
        }

        return list;
    }

    private static List<Subject> ParseClassCells(ExcelWorksheet sheet, int topBorderRow, ExcelRange columnCell)
    {
        var list = new List<Subject>();
        var test = new List<string>();

        var row = topBorderRow;
        var cell = sheet.Cells[row, columnCell.Column()];
        while (cell.Style.Border.Bottom.Style.ToString() == "None")
        {
            var value = cell.Value;
            if (value != null)
                test.Add(value.ToString());
            row++;
            cell = sheet.Cells[row, columnCell.Column()];
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

    private static Dictionary<DayOfWeek, Tuple<ExcelRange, ExcelRange?>> FindDaysBorders(ExcelWorksheet sheet)
    {
        var dayBorders = new Dictionary<DayOfWeek, Tuple<ExcelRange, ExcelRange?>>();
        
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