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
        public void HasAlignmentRows()
        {
            var grid = new Grid();
            grid.AddRow(new GridRow(new []{new GridCell("a") }));
            grid.AddRow(new GridRow(new[] { new GridCell("-"), new GridCell(":-"), new GridCell(":-:"), new GridCell("-:") }));
            grid.AddRow(new GridRow(new[] { new GridCell("1"), new GridCell("2"), new GridCell("3"), new GridCell("4"), new GridCell("5") }));
            var table = new TableParser().Parse(grid);

            Assert.NotNull(table);
            Assert.Equal(2, table.Rows.Count);

            Assert.NotNull(table.Rows[0]);
            Assert.Single(table.Rows[0]);
            Assert.NotNull(table.Rows[0][0]);
            Assert.Equal("a", table.Rows[0][0].Value);
            Assert.Equal(Alignment.Undefined, table.Rows[0][0].Alignment);

            Assert.NotNull(table.Rows[1]);
            Assert.Equal(5, table.Rows[1].Count);

            Assert.NotNull(table.Rows[1][0]);
            Assert.Equal("1", table.Rows[1][0].Value);
            Assert.Equal(Alignment.Undefined, table.Rows[1][0].Alignment);

            Assert.NotNull(table.Rows[1][1]);
            Assert.Equal("2", table.Rows[1][1].Value);
            Assert.Equal(Alignment.Left, table.Rows[1][1].Alignment);

            Assert.NotNull(table.Rows[1][2]);
            Assert.Equal("3", table.Rows[1][2].Value);
            Assert.Equal(Alignment.Center, table.Rows[1][2].Alignment);

            Assert.NotNull(table.Rows[1][3]);
            Assert.Equal("4", table.Rows[1][3].Value);
            Assert.Equal(Alignment.Right, table.Rows[1][3].Alignment);

            Assert.NotNull(table.Rows[1][4]);
            Assert.Equal("5", table.Rows[1][4].Value);
            Assert.Equal(Alignment.Undefined, table.Rows[1][4].Alignment);

        }

        [Fact]
        public void HasNotAlignmentRows()
        {
            var grid = new Grid();
            grid.AddRow(new GridRow(new[] { new GridCell("a") }));
            grid.AddRow(new GridRow(new[] { new GridCell("-"), new GridCell(":-a") }));
            var table = new TableParser().Parse(grid);

            Assert.NotNull(table);
            Assert.Equal(2, table.Rows.Count);

            Assert.NotNull(table.Rows[0]);
            Assert.Single(table.Rows[0]);
            Assert.NotNull(table.Rows[0][0]);
            Assert.Equal("a", table.Rows[0][0].Value);
            Assert.Equal(Alignment.Undefined, table.Rows[0][0].Alignment);

            Assert.NotNull(table.Rows[1]);
            Assert.Equal(2, table.Rows[1].Count);

            Assert.NotNull(table.Rows[1][0]);
            Assert.Equal("-", table.Rows[1][0].Value);
            Assert.Equal(Alignment.Undefined, table.Rows[1][0].Alignment);

            Assert.NotNull(table.Rows[1][1]);
            Assert.Equal(":-a", table.Rows[1][1].Value);
            Assert.Equal(Alignment.Undefined, table.Rows[1][1].Alignment);
        }
    }
}
