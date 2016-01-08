using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.IO;
using CodeComb.Data.Excel.Infrastructure;

namespace CodeComb.Data.Excel.Sample
{
    public class Program
    {
        public static void Main(string[] args)
        {
            var txt = File.ReadAllText(@"c:\excel\sharedStrings.xml");
            var sheettxt = File.ReadAllText(@"c:\excel\worksheets\sheet1.xml");
            var ss = new SharedStrings(txt);
            var sheet = new SheetWithoutHDR(sheettxt, ss);
            foreach(var x in sheet)
            {
                foreach(var y in x)
                {
                    Console.Write(y + "\t");
                }
                Console.Write("\r\n");
            }
            Console.ReadKey();
        }
    }
}
