using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace Budget_Forms_Core
{
    internal class WorkbookResult
    {
        public bool Success { get; set; }
        public string? Message { get; set; }//declare as nullable per debug
    }
}
