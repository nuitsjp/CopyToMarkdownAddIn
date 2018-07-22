using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CopyFromExcelToMarkdownAddIn;
using Xunit;

namespace CopyFromExcelToMarkdownAddIna.Test
{
    public class GridParserFixture
    {
        [Fact]
        public void Case01()
        {
            var grid = new GridParser().Parse(Properties.Resources.GridParserFixture_Case01);
            Assert.NotNull(grid);
            Assert.Equal(3, grid.Rows.Count);

            var row0 = grid.Rows[0];
            Assert.NotNull(row0);
            Assert.Single(row0);
            Assert.NotNull(row0[0]);
            Assert.Equal("0-1", row0[0].Value);

            var row1 = grid.Rows[1];
            Assert.NotNull(row1);
            Assert.Equal(2, row1.Count);
            Assert.NotNull(row1[0]);
            Assert.Equal("1-1", row1[0].Value);
            Assert.NotNull(row1[1]);
            Assert.Equal("1-2", row1[1].Value);

            var row2 = grid.Rows[2];
            Assert.NotNull(row2);
            Assert.Equal(3, row2.Count);
            Assert.NotNull(row2[0]);
            Assert.Equal("2-1", row2[0].Value);
            Assert.NotNull(row2[1]);
            Assert.Equal(string.Empty, row2[1].Value);
            Assert.NotNull(row2[2]);
            Assert.Equal("2-3", row2[2].Value);

        }
    }
}
