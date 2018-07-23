using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CopyFromExcelToMarkdownAddIn
{
    public class Table
    {
        private readonly List<TableRow> _rows = new List<TableRow>();
        public IReadOnlyList<TableRow> Rows => _rows;
        public void AddRow(TableRow row) => _rows.Add(row);
    }
}
