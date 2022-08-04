using ExcelReader;
using OfficeOpenXml;

DatabaseManager.CheckDatabase();
GetExcelHandler.ReadExcelSheet();

var listOfPeople = DatabaseManager.Get();

if (listOfPeople.Any())
{
    TableVisualtion.ShowTable(listOfPeople);
}

Console.WriteLine("Press any key to exit the program");
Console.ReadKey();






