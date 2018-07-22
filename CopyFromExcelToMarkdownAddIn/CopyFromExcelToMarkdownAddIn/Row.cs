using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CopyFromExcelToMarkdownAddIn
{
    public class Row : List<Cell>
    {
        public Row(IEnumerable<Cell> cells)
        {
            AddRange(cells);
        }
    }
}
