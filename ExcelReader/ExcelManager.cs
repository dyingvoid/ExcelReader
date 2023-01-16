using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ExcelReader;

public static class ExcelManager
{
    private static string _borderBackgroundColor = "FF92D050";
    private static string _mainBorderStyle;
    private static ExcelBorderStyle _mainBorderSt;
    private static Dictionary<int, List<Tuple<int, bool>>> test = new Dictionary<int, List<Tuple<int, bool>>>();
    public static Dictionary<string, Dictionary<DayOfWeek, Dictionary<int, List<string>>>> ReadExcel(FileInfo excelFile)
    {
        var groupsSchedule = new Dictionary<string, Dictionary<DayOfWeek, Dictionary<int, List<string>>>>();

        using (var package = new ExcelPackage(excelFile))
        {
            var sheets = package.Workbook.Worksheets;

            foreach (var sheet in sheets)
            {
                var borders = FindDaysBorders(sheet);
                var groupsRowNumber = borders[DayOfWeek.Monday].Item1.Row();
                DefineMainBorderStyle(sheet.Cells[groupsRowNumber, 1]);
                _mainBorderStyle = _mainBorderSt.ToString();
                var groups = FindGroupNamesCells(sheet, groupsRowNumber);
                foreach (var (group, groupColumnCell) in groups)
                {
                    groupsSchedule.Add(group, ParseWeek(sheet, borders, groupColumnCell));
                    var x = test;
                    x.Clear();
                }
                // Test(sheet);
            }
        }
        
        return groupsSchedule;
    }

    private static void Test(ExcelWorksheet sheet)
    {
        // Sometimes there are invisible borders that different from mainBorderStyle,
        // They can hide under it, this may cause errors
        var l = new List<int>() { 3, 4, 5, 6, 7, 8 };
        foreach (var i in l)
        {
            var list = FindRangesOfLessonInClassTime(sheet, 15, 140, sheet.Cells[15, i]);
        }
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

    private static Dictionary<int, Tuple<int, int>> FindClassesBorders(ExcelWorksheet sheet, 
        Tuple<ExcelRange, ExcelRange> dayBorders,
        ExcelRange columnCell)
    {
        var dict = new Dictionary<int, Tuple<int, int>>();
        // 'B' Column
        int column = 2;
        ExcelRange storedCell = null;
        
        for (var row = dayBorders.Item1.Row() + 1; row <= dayBorders.Item2.Row(); row++)
        {
            var value = sheet.Cells[row, column].Value;
            
            if (value != null && int.TryParse(value.ToString(), out var number))
            {
                if (storedCell != null)
                {
                    dict.Add(number - 1, Tuple.Create(storedCell.Row(), row - 1));
                }
                storedCell = sheet.Cells[row, column];
            }
        }
        dict.Add(int.Parse(storedCell.Value.ToString()), Tuple.Create(storedCell.Row(), dayBorders.Item2.Row() - 1));
        
        return dict;
    }

    private static Dictionary<int, List<string>> ParseDay(ExcelWorksheet sheet, 
        Tuple<ExcelRange, ExcelRange> yAxisBorders, ExcelRange columnCell)
    {
        var daySchedule = new Dictionary<int, List<string>>();
        var testDict = FindClassesBorders(sheet, yAxisBorders, columnCell);

        foreach (var (classNumber, borders) in testDict)
        {
            daySchedule.TryAdd(classNumber, ParseClassCells(sheet, borders, columnCell));
        }

        return daySchedule;
    }

    private static List<Tuple<int, bool>> FindRangesOfLessonInClassTime(ExcelWorksheet sheet, 
        int topBorderRow, int bottomBorderRow, ExcelRange columnCell)
    {
        var list = new List<Tuple<int, bool>>();

        for (var row = topBorderRow + 1; row < bottomBorderRow; row++)
        {
            var cell = sheet.Cells[row, columnCell.Column()];
            var bottomBorderStyle = cell.Style.Border.Bottom.Style;
            var topBorderStyle = cell.Style.Border.Top.Style;

            if (bottomBorderStyle != ExcelBorderStyle.None)
                list.Add(Tuple.Create(row, true));
            
            if(topBorderStyle != ExcelBorderStyle.None)
                list.Add(Tuple.Create(row, false));
        }

        /*if (list.Count > 1)
        {
            for (var i = 0; i < list.Count - 1; i++)
            {
                if (list[i + 1].Item1 == list[i].Item1 - 1 && list[i + 1].Item2 == false && list[i].Item2)
                    list.RemoveAt(i + 1);
            }
        }*/
        if(test.ContainsKey(columnCell.Column()))
            test[columnCell.Column()].AddRange(list);
        else
        {
            test.Add(columnCell.Column(), list);
        }
        return list;
    }

    private static List<string> ParseClassCells(ExcelWorksheet sheet, Tuple<int, int> borders,
        ExcelRange columnCell)
    {
        var list = new List<string>();

        // There may be several subgroups in one group. Exact number is unknown, way without number is not implemented
        for (var column = columnCell.Column(); column <= columnCell.Column() + 1; column++)
        {
            var test = FindRangesOfLessonInClassTime(sheet, borders.Item1, 
                borders.Item2, 
                sheet.Cells[columnCell.Row(), column]);
            for (var row = borders.Item1; row <= borders.Item2; ++row)
            {
                var cell = sheet.Cells[row, column];
                
                if (cell.Value != null)
                {
                    string subject = cell.Value.ToString() + sheet.Cells[row + 1, column].Value;
                    if(column == columnCell.Column())
                        list.Add("1 группа: " + subject);
                    else
                        list.Add("2 группа: " + subject);

                    row++;
                }
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

    private static void DefineMainBorderStyle(ExcelRange cellWithMainBorder)
    {
        _mainBorderSt = cellWithMainBorder.Style.Border.Bottom.Style;
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