using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CopyFromExcelToMarkdownAddIn
{
    public class TableRow : List<TableCell>
    {
        public void AddCell(TableCell tableCell) => Add(tableCell);
    }
}
