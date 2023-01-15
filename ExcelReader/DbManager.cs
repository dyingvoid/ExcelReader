using System.Data;
using System.Data.SQLite;

namespace ExcelReader;

public class DbManager
{
    private SQLiteConnection? _dbConnection;
    public DbManager(string dbPathName)
    {
        CreateDataBase(dbPathName);
        _dbConnection = CreateConnection(dbPathName);

        if (_dbConnection.State != ConnectionState.Closed)
        {
            _dbConnection.Close();
        }
    }

    public bool CreateDataBase(string dbPathName)
    {
        if (new FileInfo(dbPathName).Exists)
        {
            Console.WriteLine("DataBase already exists.");
            return false;
        }

        SQLiteConnection.CreateFile(dbPathName);
        return true;
    }

    private SQLiteConnection CreateConnection(string dbPathName)
    {
        var sqliteConnection = new SQLiteConnection($"Data Source=" +
                                                    $"{dbPathName};Version=3;New=True;Compress=True;");
        try
        {
            sqliteConnection.Open();
        }
        catch(Exception ex)
        {
            Console.WriteLine($"{ex.Source} {ex.Message}");
            throw;
        }

        return sqliteConnection;
    }
}