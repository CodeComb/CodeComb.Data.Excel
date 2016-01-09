using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.IO;
using System.IO.Compression;
using System.Xml;
using CodeComb.Data.Excel.Infrastructure;

namespace CodeComb.Data.Excel
{
    public class ExcelStream : IDisposable
    {
        private FileStream file;
        public ZipArchive ZipArchive { get; set; }
        private SharedStrings sharedStrings;
        private SharedStrings cachedSharedStrings
        {
            get
            {
                if (sharedStrings == null)
                {
                    var e = ZipArchive.GetEntry("xl/sharedStrings.xml");
                    using (var stream = e.Open())
                    {
                        var sr = new StreamReader(stream);
                        var result = sr.ReadToEnd();
                        sharedStrings = new SharedStrings(result);
                    }
                }
                return sharedStrings;
            }
        }

        public List<WorkBook> WorkBook { get; set; } = new List<WorkBook>();

        public ExcelStream(string path)
        {
            file = File.Open(path, FileMode.Open , FileAccess.ReadWrite);
            ZipArchive = new ZipArchive(file, ZipArchiveMode.Read | ZipArchiveMode.Update);
            var e = ZipArchive.GetEntry("xl/workbook.xml");
            using (var stream = e.Open())
            {
                var sr = new StreamReader(stream);
                var result = sr.ReadToEnd();
                var xd = new XmlDocument();
                xd.LoadXml(result);
                var tmp = xd.GetElementsByTagName("sheet");
                foreach(XmlNode x in tmp)
                {
                    var name = x.Attributes["name"].Value;
                    var sheetId = x.Attributes["sheetId"].Value;
                    WorkBook.Add(new WorkBook
                    {
                        Name = name,
                        Id = Convert.ToUInt64(sheetId)
                    });
                }
            }
        }

        public SheetWithoutHDR LoadSheet(string name)
        {
            var worksheet = WorkBook.Where(x => x.Name == name).First();
            var e = ZipArchive.GetEntry($"xl/worksheets/{worksheet.FileName}");
            using (var stream = e.Open())
            {
                var sr = new StreamReader(stream);
                var result = sr.ReadToEnd();
                return new SheetWithoutHDR(worksheet.Id, result, this ,cachedSharedStrings);
            }
        }

        public SheetWithoutHDR LoadSheet(ulong Id)
        {
            var worksheet = WorkBook.Where(x => x.Id == Id).First();
            var e = ZipArchive.GetEntry($"xl/worksheets/{worksheet.FileName}");
            using (var stream = e.Open())
            {
                var sr = new StreamReader(stream);
                var result = sr.ReadToEnd();
                return new SheetWithoutHDR(worksheet.Id, result, this, cachedSharedStrings);
            }
        }

        public SheetHDR LoadSheetHDR(string name)
        {
            var worksheet = WorkBook.Where(x => x.Name == name).First();
            var e = ZipArchive.GetEntry($"xl/worksheets/{worksheet.FileName}");
            using (var stream = e.Open())
            {
                var sr = new StreamReader(stream);
                var result = sr.ReadToEnd();
                return new SheetHDR(worksheet.Id, result, this, cachedSharedStrings);
            }
        }

        public SheetHDR LoadSheetHDR(ulong Id)
        {
            var worksheet = WorkBook.Where(x => x.Id == Id).First();
            var e = ZipArchive.GetEntry($"xl/worksheets/{worksheet.FileName}");
            using (var stream = e.Open())
            {
                var sr = new StreamReader(stream);
                var result = sr.ReadToEnd();
                return new SheetHDR(worksheet.Id, result, this, cachedSharedStrings);
            }
        }

        public void Dispose()
        {
            ZipArchive.Dispose();
            file.Dispose();
        }
    }
}
