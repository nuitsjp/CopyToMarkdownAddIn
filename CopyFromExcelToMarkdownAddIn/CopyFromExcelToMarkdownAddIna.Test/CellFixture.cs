using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CopyFromExcelToMarkdownAddIn;
using Xunit;

namespace CopyFromExcelToMarkdownAddIna.Test
{
    public class CellFixture
    {
        [Fact]
        public void IsAlignment()
        {
            Assert.True(new Cell("-").IsAlignment);
            Assert.True(new Cell(":-").IsAlignment);
            Assert.True(new Cell(":-:").IsAlignment);
            Assert.True(new Cell("-:").IsAlignment);
            Assert.True(new Cell("--").IsAlignment);
            Assert.False(new Cell("-:-").IsAlignment);
            Assert.False(new Cell(string.Empty).IsAlignment);
            Assert.False(new Cell(":").IsAlignment);
            Assert.False(new Cell("a").IsAlignment);
            Assert.False(new Cell("-a").IsAlignment);
            Assert.False(new Cell("a-").IsAlignment);
        }
    }
}
