using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace CodeComb.Data.Excel.Infrastructure
{
    public class Sheet : List<Row>, IDisposable
    {
        public void Dispose()
        {
            this.Clear();
            GC.Collect();
        }
    }
}
