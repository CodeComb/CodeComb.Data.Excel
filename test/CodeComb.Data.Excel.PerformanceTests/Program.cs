using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using CodeComb.Data.Excel;

namespace CodeComb.Data.Excel.PerformanceTests
{
    public class Program
    {
        public static void Main(string[] args)
        {
            var time1 = DateTime.Now;
            using (var excel = ExcelStream.Create("test.xlsx"))
            using (var sheet1 = excel.LoadSheet(1))
            {
                var rnd = new Random();
                for (var i = 1; i <= 100000; i++)
                {
                    sheet1.Add(new Infrastructure.Row
                    {
                        i.ToString(),
                        rnd.Next(0, 100000000).ToString(),
                        Guid.NewGuid().ToString()
                    });
                }
                sheet1.SaveChanges();
            }
            var time2 = DateTime.Now;
            //System.IO.File.Delete("test.xlsx");
            Console.WriteLine("Writing 100,000: " + (time2 - time1));
            Console.Read();
        }
    }
}
