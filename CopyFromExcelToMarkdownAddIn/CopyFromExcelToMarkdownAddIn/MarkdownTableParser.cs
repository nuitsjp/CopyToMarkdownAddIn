using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CopyFromExcelToMarkdownAddIn
{
    public class MarkdownTableParser
    {
        public Table Parse(string markdown)
        {
            var table = new Table();
            using (var reader = new StringReader(markdown))
            {
                var line = reader.ReadLine();
                if (line != null)
                {

                }
            }
                return table;
        }
    }
}
