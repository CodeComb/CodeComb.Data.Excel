﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Data;
using System.Data.Common;
using System.Xml;
using System.Xml.Linq;

namespace CodeComb.Data.Excel.Infrastructure
{
    public class SheetWithoutHDR : Sheet
    {
        public SheetWithoutHDR(ulong Id, string XmlSource, ExcelStream Excel, SharedStrings stringDictionary)
            : base(Id, Excel, stringDictionary)
        {
            var xd = new XmlDocument();
            xd.LoadXml(XmlSource);
            var rows = xd.GetElementsByTagName("row");
            // 遍历row标签
            foreach (XmlNode x in rows)
            {
                var cols = x.ChildNodes;
                var objs = new Row();
                // 遍历c标签
                foreach (XmlNode y in cols)
                {
                    string value = null;
                    // 如果是字符串类型，则需要从字典中查询
                    if (y.Attributes["t"]?.Value == "s")
                    {
                        var index = Convert.ToUInt64(y.FirstChild.InnerText);
                        value = StringDictionary[index];
                    }
                    // 否则其中的v标签值即为单元格内容
                    else
                    {
                        value = y.InnerText;
                    }

                    objs.Add(value, y.Attributes["r"].Value);
                }

                // 去掉末尾的null
                while (objs.LastOrDefault() == null)
                    objs.RemoveAt(objs.Count - 1);
                if (objs.Count > 0)
                    this.Add(objs);
            }
            GC.Collect();
        }
    }
}