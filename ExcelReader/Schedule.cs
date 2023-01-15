using System.Runtime.Intrinsics.Arm;

namespace ExcelReader;

public class Schedule
{
    public Schedule(string groupName, Dictionary<DayOfWeek, List<Subject>> weekSchedule)
    {
        GroupName = groupName;
        WeekSchedule = weekSchedule;
    }
    public string GroupName { get; set; }
    public Dictionary<DayOfWeek, List<Subject>> WeekSchedule { get; set; }
}