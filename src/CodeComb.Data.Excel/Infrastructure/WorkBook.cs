using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Xml;

namespace CodeComb.Data.Excel.Infrastructure
{
    public class WorkBook
    {
        public WorkBook(string XmlSource)
        {
            var xd = new XmlDocument();
            xd.LoadXml(XmlSource);
            var tmp = xd.GetElementsByTagName("sheet");
            foreach(XmlNode x in tmp)
            {

            }
        }
    }
}
