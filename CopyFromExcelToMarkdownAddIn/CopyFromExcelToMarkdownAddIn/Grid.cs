using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CopyFromExcelToMarkdownAddIn
{
    public class Grid
    {
        private readonly List<GridRow> _rows = new List<GridRow>();

        public IReadOnlyList<GridRow> Rows => _rows;

        public bool HasAlignmentRows => 1 < _rows.Count && _rows[1].Count(x => !x.IsAlignment) == 0;

        public void AddRow(GridRow row) => _rows.Add(row);
   }
}
