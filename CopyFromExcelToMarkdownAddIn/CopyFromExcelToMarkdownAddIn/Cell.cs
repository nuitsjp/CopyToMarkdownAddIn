using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CopyFromExcelToMarkdownAddIn
{
    public class Cell
    {
        public string Value { get; }

        public bool IsAlignment =>
            System.Text.RegularExpressions.Regex.IsMatch(Value, "^:?-+:?$");

        public Cell(string value)
        {
            Value = value;
        }
    }
}
