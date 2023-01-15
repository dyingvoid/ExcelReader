using OfficeOpenXml;

namespace ExcelReader;

public class Subject
{
    public Subject(string name, string teacher, TimeOnly start, TimeOnly end, int numberInDay)
    {
        Name = name;
        Teacher = teacher;
        Start = start;
        End = end;
        NumberInDay = numberInDay;
    }

    public Subject(ExcelRange subjectCell)
    {
        
    }
    
    public int  Id { get; set; }
    public string Name { get; set; }
    public string Teacher { get; set; }
    public TimeOnly Start { get; set; }
    public TimeOnly End { get; set; }
    public int NumberInDay { get; set; }
}