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
        private ZipArchive archive;
        private SharedStrings sharedStrings;
        private SharedStrings cachedSharedStrings
        {
            get
            {
                if (sharedStrings == null)
                {
                    var e = archive.GetEntry("xl/sharedStrings.xml");
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

        public List<WorkBook> WorkBook { get; set; } = new List<Infrastructure.WorkBook>();

        public ExcelStream(string path)
        {
            file = File.Open(path, FileMode.Open);
            archive = new ZipArchive(file);
            var e = archive.GetEntry("xl/workbook.xml");
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
            var filename = WorkBook.Where(x => x.Name == name).First();
            var e = archive.GetEntry($"xl/worksheets/{filename.FileName}");
            using (var stream = e.Open())
            {
                var sr = new StreamReader(stream);
                var result = sr.ReadToEnd();
                return new SheetWithoutHDR(result, cachedSharedStrings);
            }
        }

        public SheetWithoutHDR LoadSheet(ulong Id)
        {
            var filename = WorkBook.Where(x => x.Id == Id).First();
            var e = archive.GetEntry($"xl/worksheets/{filename.FileName}");
            using (var stream = e.Open())
            {
                var sr = new StreamReader(stream);
                var result = sr.ReadToEnd();
                return new SheetWithoutHDR(result, cachedSharedStrings);
            }
        }

        public SheetHDR LoadSheetHDR(string name)
        {
            var filename = WorkBook.Where(x => x.Name == name).First();
            var e = archive.GetEntry($"xl/worksheets/{filename.FileName}");
            using (var stream = e.Open())
            {
                var sr = new StreamReader(stream);
                var result = sr.ReadToEnd();
                return new SheetHDR(result, cachedSharedStrings);
            }
        }

        public SheetHDR LoadSheetHDR(ulong Id)
        {
            var filename = WorkBook.Where(x => x.Id == Id).First();
            var e = archive.GetEntry($"xl/worksheets/{filename.FileName}");
            using (var stream = e.Open())
            {
                var sr = new StreamReader(stream);
                var result = sr.ReadToEnd();
                return new SheetHDR(result, cachedSharedStrings);
            }
        }

        public void Dispose()
        {
            archive.Dispose();
            file.Dispose();
        }
    }
}
