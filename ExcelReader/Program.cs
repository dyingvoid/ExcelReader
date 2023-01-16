using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.Unicode;
using ExcelReader;
using OfficeOpenXml;

public class Program
{
    public static void Main()
    {
        string dbPathName = @"C:\Users\Dying\RiderProjects\ExcelReader\ExcelReader\Test.sqlite";
        ExcelWork();
    }

    public static void Entity()
    {
        using (var db = new ApplicationContext())
        {
            var programming = new Subject("Programming", "Kaa", 
                new TimeOnly(8, 0), new TimeOnly(9, 30),
                1);

            db.Subjects.Add(programming);
            db.SaveChanges();
            Console.WriteLine("Success!");

            var subjects = db.Subjects.ToList();
            foreach (var subject in subjects)
            {
                Console.WriteLine(subject.Name);
            }
        }
    }

    public static void DataBaseWork(string dbPathName)
    {
        var dbWorker = new DbManager(dbPathName);
    }

    public static void ExcelWork()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        string filePath = @"C:\Users\Dying\Downloads\Расписание ИИТ 1 сем 22-23.xlsx";
        var file = new FileInfo(filePath);
        var weekSchedule = ExcelManager.ReadExcel(file);
        DictToJson(weekSchedule);
    }
    
    public static void DictToJson<TDict>(TDict schedule)
    {
        var options = new JsonSerializerOptions
        {
            Encoder = JavaScriptEncoder.Create(UnicodeRanges.BasicLatin, UnicodeRanges.Cyrillic),
            WriteIndented = true
        };
        string information = JsonSerializer.Serialize(schedule, options);
        File.WriteAllText(@"..\..\..\testJson.json", information);
    }
}