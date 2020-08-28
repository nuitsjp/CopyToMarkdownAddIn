using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace CopyFromExcelToMarkdownAddIn
{
    public static class RangeExtensions
    {
        private const string ExcelEnterNotation = "\n";
        private const string MarkdownEnterNotation = "<br>";
        private const string ExcelPipeNotation = "|";
        private const string MarkdownPipeNotation = "&#124;";


        public static string FormatText(this Range range)
        {
            if (range == null || range.Text == null)
            {
                return string.Empty;
            }
            else
            {
                return range.Text
                    .Replace(ExcelEnterNotation, MarkdownEnterNotation)
                    .Replace(ExcelPipeNotation, MarkdownPipeNotation);
            }
        }

    }
}
