using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CopyFromExcelToMarkdownAddIn;
using Xunit;

namespace CopyFromExcelToMarkdownAddIna.Test
{
    public class GridCellFixture
    {
        [Fact]
        public void IsAlignment()
        {
            Assert.True(new GridCell("-").IsAlignment);
            Assert.True(new GridCell(":-").IsAlignment);
            Assert.True(new GridCell(":-:").IsAlignment);
            Assert.True(new GridCell("-:").IsAlignment);
            Assert.True(new GridCell("--").IsAlignment);
            Assert.False(new GridCell("-:-").IsAlignment);
            Assert.False(new GridCell(string.Empty).IsAlignment);
            Assert.False(new GridCell(":").IsAlignment);
            Assert.False(new GridCell("a").IsAlignment);
            Assert.False(new GridCell("-a").IsAlignment);
            Assert.False(new GridCell("a-").IsAlignment);
        }
    }
}
