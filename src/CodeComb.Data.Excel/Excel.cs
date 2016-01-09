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
                    // 如果sharedStrings.xml不存在，则创建
                    if (e == null)
                    {
                        // 创建sharedStrings.xml
                        e = ZipArchive.CreateEntry("xl/sharedStrings.xml", CompressionLevel.Optimal);
                        using (var stream = e.Open())
                        using (var sw = new StreamWriter(stream))
                        {
                            sw.Write(@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><sst xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" count=""0""></sst>");
                        }

                        // 同时需要向[Content_Types].xml添加sharedStrings.xml的信息
                        var e2 = ZipArchive.GetEntry("[Content_Types].xml");
                        using (var stream = e2.Open())
                        {
                            var sr = new StreamReader(stream);
                            var result = sr.ReadToEnd();
                            var xd = new XmlDocument();
                            xd.LoadXml(result);
                            var element = xd.CreateElement("Override", xd.DocumentElement.NamespaceURI);
                            element.SetAttribute("PartName", "/xl/sharedStrings.xml");
                            element.SetAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml");
                            var tmp = xd.GetElementsByTagName("Types")
                                .Cast<XmlNode>()
                                .First()
                                .AppendChild(element);
                            stream.Position = 0;
                            stream.SetLength(0);
                            xd.Save(stream);
                        }
                    }
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

        public void RemoveSheet(string name)
        {
            var sheetId = WorkBook.Where(x => x.Name == name)
                .Select(x => x.Id)
                .First();
            RemoveSheet(sheetId);
        }

        public void RemoveSheet(ulong Id)
        {
            // 从ExcelStream对象中删除
            WorkBook.Remove(WorkBook.Where(x => x.Id == Id).First());

            // 从workbook.xml中删除
            var e = ZipArchive.GetEntry("xl/workbook.xml");
            using (var stream = e.Open())
            {
                var sr = new StreamReader(stream);
                var result = sr.ReadToEnd();
                var xd = new XmlDocument();
                xd.LoadXml(result);
                var tmp = xd.GetElementsByTagName("sheet")
                    .Cast<XmlNode>()
                    .Where(x => x.Attributes["sheetId"].Value == Id.ToString())
                    .Single();
                tmp.ParentNode.RemoveChild(tmp);
                stream.Position = 0;
                stream.SetLength(0);
                xd.Save(stream);
            }

            // 删除sheetX.xml
            var e2 = ZipArchive.GetEntry($"xl/worksheets/sheet{Id}.xml");
            e2.Delete();

            // 从[Content_Types].xml中删除
            var e3 = ZipArchive.GetEntry("[Content_Types].xml");
            using (var stream = e3.Open())
            {
                var sr = new StreamReader(stream);
                var result = sr.ReadToEnd();
                var xd = new XmlDocument();
                xd.LoadXml(result);
                var tmp = xd.GetElementsByTagName("Override")
                    .Cast<XmlNode>()
                    .Where(x => x.Attributes["PartName"].Value == $"/xl/worksheets/sheet{Id}.xml")
                    .Single();
                tmp.ParentNode.RemoveChild(tmp);
                stream.Position = 0;
                stream.SetLength(0);
                xd.Save(stream);
            }
        }

        public void Dispose()
        {
            ZipArchive.Dispose();
            file.Dispose();
        }
    }
}
