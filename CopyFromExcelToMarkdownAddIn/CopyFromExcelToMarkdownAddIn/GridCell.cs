using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CopyFromExcelToMarkdownAddIn
{
    public class GridCell
    {
        public string Value { get; }

        public bool IsAlignment =>
            System.Text.RegularExpressions.Regex.IsMatch(Value, "^:?-+:?$");

        public GridCell(string value)
        {
            Value = value;
        }
    }
}
