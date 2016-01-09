using System;
using CodeComb.Data.Excel;

namespace CodeComb.Data.Excel.Sample
{
    public class Program
    {
        public static void Main(string[] args)
        {
            using (var x = new ExcelStream(@"c:\excel\1.xlsx"))
            using (var sheet = x.LoadSheet("Sheet1")) // var sheet = x.LoadSheet(1)
            {
                sheet.Add(new Infrastructure.Row
                {
                    "Hello world!"
                });
                sheet.SaveChanges();

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
