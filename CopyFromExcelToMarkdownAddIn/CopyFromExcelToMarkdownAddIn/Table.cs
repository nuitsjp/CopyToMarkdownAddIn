using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CopyFromExcelToMarkdownAddIn
{
    public class Table
    {
        public Row Header { get; }
        public IList<Row> Rows { get; } = new List<Row>();
    }
}
