using ExcelReader;
using OfficeOpenXml;

DatabaseManager.CheckDatabase();
GetExcelHandler.ReadExcelSheet();

var listOfPeople = DatabaseManager.Get();

TableVisualtion.ShowTable(listOfPeople);





