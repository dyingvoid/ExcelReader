using OfficeOpenXml;

namespace ExcelReader;

public static class Extensions
{
    public static void ChangeCellAddressLetter(this ExcelRange cell, string letters)
    {
        cell.Address = letters + 
                       cell.Address.Substring(letters.Length, cell.Address.Length - letters.Length);
    }

    public static int Row(this ExcelRange cell)
    {
        return int.Parse(string.Concat(cell.Address.Where(c => char.IsDigit(c))));
    }

    public static int Column(this ExcelRange cell)
    {
        int column = 0;
        string letterAdress = string.Concat(cell.Address.Where(c => char.IsLetter(c)));
        foreach (var c in letterAdress)
        {
            column += char.ToUpper(c) - 64;
        }

        return column;
    }
}