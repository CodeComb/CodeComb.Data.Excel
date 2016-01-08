using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Data;
using System.Data.Common;
using System.Xml;
using System.Xml.Linq;

namespace CodeComb.Data.Excel.Infrastructure
{
    public class SheetHDR : Sheet
    {
        public SheetHDR(string XmlSource, SharedStrings StringDictionary)
        {
            var xd = new XmlDocument();
            xd.LoadXml(XmlSource);
            var rows = xd.GetElementsByTagName("row");
            // 遍历row标签
            var flag = false;
            Header header = new Header();
            foreach (XmlNode x in rows)
            {
                var cols = x.ChildNodes;
                var objs = new Row(header);
                // 遍历c标签
                foreach (XmlNode y in cols)
                {
                    string value = null;
                    // 如果是字符串类型，则需要从字典中查询
                    foreach (XmlAttribute z in y.Attributes)
                    {
                        if (z.Name == "t" && z.Value == "s")
                        {
                            var index = Convert.ToUInt64(y.FirstChild.InnerText);
                            value = StringDictionary[index];
                            break;
                        }
                    }
                    // 否则其中的v标签值即为单元格内容
                    if (string.IsNullOrEmpty(value))
                        value = y.FirstChild.InnerText;

                    if (!flag)
                    {
                        header.Add(value);
                        continue;
                    }
                    objs.Add(value);
                }
                if (!flag)
                {
                    while (header.LastOrDefault() == null)
                        header.RemoveAt(header.Count - 1);
                    flag = true;
                    continue;
                }
                // 去掉末尾的null
                while (objs.LastOrDefault() == null)
                    objs.RemoveAt(objs.Count - 1);
                this.Add(objs);
                GC.Collect();
            }
        }
    }
}
