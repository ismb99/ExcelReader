using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReader
{
    public class GetExcelHandler
    {
        public static List<T> GetList<T>(ExcelWorksheet worksheet) where T : class
        {
            List<T> list = new();

            var columnInformation = Enumerable.Range(1, worksheet.Dimension.Columns).ToList().Select(n =>
                new
                {
                    Index = n,
                    ColumnName = worksheet.Cells[1, n].Value.ToString()
                });

            for (int row = 2; row <= worksheet.Dimension.Rows; row++)
            {
                T obj = (T)Activator.CreateInstance(typeof(T));
                if (obj is not null)
                {
                    foreach (var prop in typeof(T).GetProperties())
                    {
                        if (columnInformation is not null)
                        {
                            int col = columnInformation.SingleOrDefault(c => c.ColumnName == prop?.Name).Index;
                            var val = worksheet.Cells[row, col].Value;
                            var propType = prop.PropertyType;

                            prop.SetValue(obj, Convert.ChangeType(val, propType));
                        }
                        
                    }
                    list.Add(obj);
                }
               
            }
            return list;
        }


        public static void ReadExcelSheet()
        {
            FileInfo file = new FileInfo(@"C:\Users\jazzy\source\repos\ExcelReader\ExcelToDB.xlsx");

            using (ExcelPackage excelPackage = new ExcelPackage(file))
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                var excelSheet = excelPackage.Workbook.Worksheets.First();
                var personDtoList = GetList<PersonDto>(excelSheet);

                List<Person> mappedPersonList = Mapper.FromDtoMapper(personDtoList);

                DatabaseManager.WriteExcelToDatabase(mappedPersonList);
            }
        }
    }
}
