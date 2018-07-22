using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CopyFromExcelToMarkdownAddIn
{
    public class Grid
    {
        private readonly List<Row> _rows = new List<Row>();

        public IReadOnlyList<Row> Rows => _rows;

        public void AddRow(Row row) => _rows.Add(row);
    }
}
