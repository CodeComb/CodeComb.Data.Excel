# CodeComb.Data.Excel

Excel data provider for vNext (dnxcore).

It is under developing now, we will open source after some tests and refactory.

## Samples

### Create a new workbook

```c#
using System;
using CodeComb.Data.Excel;

namespace CodeComb.Data.Excel.Sample
{
    public class Program
    {
        public static void Main(string[] args)
        {
            using (var x = ExcelStream.Create(@"c:\excel\test.xlsx"))
            using (var sheet = x.LoadSheet(1))
            {
                sheet.Add(new Infrastructure.Row
                {
                    "Create test"
                });
                sheet.SaveChanges();
            }
        }
    }
}

```

### Load an existed workbook

```
using (var x = new ExcelStream(@"c:\excel\somebook.xlsx"))
{
    ...
}
```

### Read the sheet

```c#
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
```

### Updates deletions and additions

```c#
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
                // Adding a row
                sheet.Add(new Infrastructure.Row
                {
                    null,
                    null,
                    "Hello world!"
                });
                
                sheet.RemoveAt(1); // Remove the row 1
                
                sheet.SaveChanges(); // Save changes
            }
            Console.ReadKey();
        }
    }
}

```

### Create a sheet

```c#
using (var x = new ExcelStream(@"c:\excel\1.xlsx"))
using (var sheet = x.CreateSheet("Sheet2"))
{
    sheet.Add(new Infrastructure.Row
    {
        "Code Comb"
    });
    sheet.SaveChanges();
}
```

### Remove a sheet

```c#
using (var x = new ExcelStream(@"c:\excel\1.xlsx"))
{
    x.RemoveSheet("Sheet1"); // x.RemoveSheet(1);
}
```

### Row 1 as header

```c#
using System;
using System.Linq;
using CodeComb.Data.Excel;

namespace CodeComb.Data.Excel.Sample
{
    public class Program
    {
        public static void Main(string[] args)
        {
            using (var x = new ExcelStream(@"c:\excel\1.xlsx"))
            using (var sheet = x.LoadSheetHDR("Sheet1")) // var sheet = x.LoadSheetHDR(1)
            {
                foreach (var row in sheet)
                    Console.WriteLine($"{row["Name"]} {row["Sex"]} {row["Name"]} {row["School"]} {row["Score"]}");
                Console.WriteLine($"Average score: {sheet.Average(a => Convert.ToDouble(a["Score"]))}");
            }
            Console.ReadKey();
        }
    }
}
```
