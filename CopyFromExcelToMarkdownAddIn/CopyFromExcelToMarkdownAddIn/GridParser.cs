using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CopyFromExcelToMarkdownAddIn
{
    public class GridParser
    {
        public Grid Parse(string markdown)
        {
            var grid = new Grid();

            using (var reader = new StringReader(markdown))
            for(var line = reader.ReadLine(); line != null; line = reader.ReadLine())
            {
                grid.AddRow(new Row(TrimPipe(line).Split('|').Select(x => new Cell(x))));
            }

            return grid;
        }

        private string TrimPipe(string line)
        {
            var result = line.Trim();

            if (result.StartsWith("|"))
                result = result.Substring(1);

            if (result.EndsWith("|"))
                result = result.Substring(0, result.Length - 1);

            return result;
        }
    }
}
