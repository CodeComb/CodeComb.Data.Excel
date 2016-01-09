using System;
using CodeComb.Data.Excel;

namespace CodeComb.Data.Excel.Sample
{
    public class Program
    {
        public static void Main(string[] args)
        {
            using (var x = new ExcelStream(@"c:\excel\1.xlsx")) // Open excel file
            using (var sheet = x.LoadSheet("Sheet1")) // var sheet = x.LoadSheet(1)
            {
                // Removing sheet2
                x.RemoveSheet(2);

                // Creating sheet
                var s = x.CreateSheet("Hello");
                s.Add(new Infrastructure.Row
                {
                    "Code Comb"
                });
                s.SaveChanges();

                // Writing data into sheet
                sheet.Add(new Infrastructure.Row
                {
                    "Hello world!"
                });
                sheet.SaveChanges();

                // Reading the data from sheet
                foreach (var a in sheet)
                {
                    foreach (var b in a)
                        Console.Write(b + '\t');
                    Console.Write("\r\n");
                }
            }
            Console.ReadKey();
        }
    }
}
