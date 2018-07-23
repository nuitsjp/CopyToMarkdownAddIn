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

        public bool IsAlignment { get; }
            

        public Alignment Alignment { get; }

        public GridCell(string value)
        {
            Value = value;
            if (System.Text.RegularExpressions.Regex.IsMatch(Value, "^-+$"))
            {
                IsAlignment = true;
                Alignment = Alignment.Undefined;
            }
            else if (System.Text.RegularExpressions.Regex.IsMatch(Value, "^:?-+$"))
            {
                IsAlignment = true;
                Alignment = Alignment.Left;
            }
            else if (System.Text.RegularExpressions.Regex.IsMatch(Value, "^:-+:$"))
            {
                IsAlignment = true;
                Alignment = Alignment.Center;
            }
            else if (System.Text.RegularExpressions.Regex.IsMatch(Value, "^-+:?$"))
            {
                IsAlignment = true;
                Alignment = Alignment.Right;
            }
            else
            {
                IsAlignment = false;
                Alignment = Alignment.Undefined;
            }
        }
    }
}
