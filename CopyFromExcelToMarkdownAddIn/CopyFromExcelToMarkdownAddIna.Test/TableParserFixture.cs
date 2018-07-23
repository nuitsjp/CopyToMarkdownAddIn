using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CopyFromExcelToMarkdownAddIn;
using Xunit;

namespace CopyFromExcelToMarkdownAddIna.Test
{
    public class TableParserFixture
    {
        [Fact]
        public void Case01()
        {
            var grid = new Grid();
            grid.AddRow(new GridRow(new []{new GridCell("a"), new GridCell("b") }));
            grid.AddRow(new GridRow(new[] { new GridCell(":-"), new GridCell("-:") }));
            grid.AddRow(new GridRow(new[] { new GridCell("1"), new GridCell("2") }));
        }
    }
}
