using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CopyFromExcelToMarkdownAddIn
{
    public class TableCell
    {
        public Alignment Alignment { get; }
        public string Value { get; }

        public TableCell(string value, Alignment alignment)
        {
            Value = value;
            Alignment = alignment;
        }
    }
}
