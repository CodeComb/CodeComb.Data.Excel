using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace CodeComb.Data.Excel.Infrastructure
{
    public class WorkSheet
    {
        private WorkBook wb;

        public string Title { get; set; }

        public Sheet Sheet { get; private set; }
    }
}
