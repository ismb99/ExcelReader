using ExcelReader;
using OfficeOpenXml;

//string file = @"C:\Users\im_07\OneDrive\desktop\data.xlsx";

DatabaseManager databaseManager = new();
databaseManager.CheckDatabase();
GetExcelHandler getExcel = new();
getExcel.ReadExcelSheet();
//ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
//FileInfo file = new FileInfo(@"C:\Users\im_07\documents\ExcelToDB.xlsx");

//List<Person> people = GetSetupData();

//await SaveExcelFile(people, file);


//async Task SaveExcelFile(List<Person> people, FileInfo file)
//{
//    DeleteIfExists(file);
//    using (var package = new ExcelPackage(file))
//    {
//        ExcelWorksheet workSheet = package.Workbook.Worksheets.Add("MainReport");
//        var range = workSheet.Cells["A1"].LoadFromCollection(people, true);
//        range.AutoFitColumns();
//        await package.SaveAsync();
//    }
//}

//void DeleteIfExists(FileInfo file)
//{
//    if (file.Exists)
//    {
//        file.Delete();
//    }
//}

//static List<Person> GetSetupData()
//{
//    var output = new List<Person>()
//    {
//        new Person()
//        {
//            Id = 1,
//            FirstName = "Ismail",
//            LastName = "Mohamed"
//        },
//         new Person()
//        {
//            Id = 1,
//            FirstName = "James",
//            LastName = "Smith"
//        },
//          new Person()
//        {
//            Id = 1,
//            FirstName = "Mike",
//            LastName = "Jones"
//        }
//    };
//    return output;
//}







//const string file = @"C:\Users\im_07\documents\ExcelToDB.xlsx";











//// path to your excel file
////string path = "C:/****/sample_data.xlsx";
//FileInfo file = new FileInfo(@"C:\Users\im_07\documents\ExcelToDB.xlsx");
//List<Person> personList = new();

//ExcelPackage package = new ExcelPackage(file);
//ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();

//// get number of rows and columns in the sheet
//int rows = worksheet.Dimension.Rows; // 12
//int columns = worksheet.Dimension.Columns; // 3

//// loop through the worksheet rows and columns
//for (int i = 2; i <= rows; i++)
//{
//    for (int j = 1; j <= columns; j++)
//    {

//        string content = worksheet.Cells[i, j].Value.ToString();

//        personList.Add(new Person
//        {
//            Id = i,
//            FirstName = content,
//            LastName = content,
//        });


//        //Console.WriteLine($"Content = {content}");
//        /* Do something ...*/
//    }
//}

//foreach (var p in personList)
//{
//    Console.WriteLine($"{p.Id} {p.FirstName} {p.LastName}");
//}