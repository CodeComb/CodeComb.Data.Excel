using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace CodeComb.Data.Excel.Infrastructure
{
    public class Sheet : List<Row>, IDisposable
    {
        protected SharedStrings StringDictionary { get; set; }

        public void Dispose()
        {
            this.Clear();
            GC.Collect();
        }

        public void SaveChanges()
        {
            var tmp = this.Cast<List<string>>()
                .ToList();
            if (this is SheetHDR)
                tmp.Insert(0, (this as SheetHDR).Header.ToList());
            var row = 1ul;
            RowNumber col;
            foreach (var x in this)
            {
                col = new RowNumber("A");
                foreach (var y in x)
                {
                    var innerText = ""; 
                    try
                    {
                        // 如果是数值类型，则直接写入xml
                        if (x.Contains("."))
                        {
                            // 如果是小数尝试转换为double
                        }
                        else
                        {
                            // 如果是整数尝试转换为long
                        }
                    }
                    catch
                    {
                        // 否则需要将字符串添加到sharedStrings.xml中，并生成索引
                        if (!StringDictionary.Contains(y))
                            StringDictionary.Add(y);
                        innerText = StringDictionary.IndexOf(y).ToString();
                    }
                    col++;
                }
                row++;
            }
        }
    }
}