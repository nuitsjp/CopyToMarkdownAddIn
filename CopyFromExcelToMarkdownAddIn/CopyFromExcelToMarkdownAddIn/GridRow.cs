using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CopyFromExcelToMarkdownAddIn
{
    public class GridRow : List<GridCell>
    {
        public GridRow(IEnumerable<GridCell> cells) => AddRange(cells);
    }
}
