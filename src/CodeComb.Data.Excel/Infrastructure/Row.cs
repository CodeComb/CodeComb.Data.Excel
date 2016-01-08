using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace CodeComb.Data.Excel.Infrastructure
{
    public class Row : List<string>
    {
        private Header header;

        public Row(Header Header = null)
        {
            header = Header;
        }

        public string this[string index]
        {
            get
            {
                try
                {
                    return this[header.IndexOf(index)];
                }
                catch
                {
                    return null;
                }
            }
        }
    }
}
