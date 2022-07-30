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
        public List<T> GetList<T>(ExcelWorksheet worksheet)
        {
            List<T> list = new List<T>();
            var columnInformation = Enumerable.Range(1, worksheet.Dimension.Columns).ToList().Select(n =>
                new
                {
                    Index = n,
                    ColumnName = worksheet.Cells[1, n].Value.ToString()
                });
            //2
            for (int row = 2; row <= worksheet.Dimension.Rows; row++)
            {
                T obj = (T)Activator.CreateInstance(typeof(T));
                foreach (var prop in typeof(T).GetProperties())
                {
                    int col = columnInformation.SingleOrDefault(c => c.ColumnName == prop.Name).Index;
                    var val = worksheet.Cells[row, col].Value;
                    var propType = prop.PropertyType;


                    prop.SetValue(obj, Convert.ChangeType(val, propType));
                }
                list.Add(obj);
            }
            return list;
        }

        public void ReadExcelSheet()
        {
            DatabaseManager database = new();
            FileInfo file = new FileInfo(@"C:\Users\im_07\documents\ExcelToDB.xlsx");

            using (ExcelPackage excelPackage = new ExcelPackage(file))
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                var excelSheet = excelPackage.Workbook.Worksheets.First();
                var person = GetList<Person>(excelSheet);
                database.WriteExcelToDatabase(person);
            }
        }
       
    }
}
