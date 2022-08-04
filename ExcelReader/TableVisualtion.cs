using ConsoleTableExt;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReader
{
    public class TableVisualtion
    {
        public static void ShowTable(List<Person> person)
        {
            Console.WriteLine("\n\n");

            ConsoleTableBuilder
                .From(person)
                .WithTitle("Excel People")
                .ExportAndWriteLine();
            Console.WriteLine("\n\n");
        }
    }
}
