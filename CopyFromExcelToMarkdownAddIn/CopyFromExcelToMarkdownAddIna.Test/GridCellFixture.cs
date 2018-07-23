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
            Assert.True(new GridCell("--").IsAlignment);
            Assert.True(new GridCell(":-").IsAlignment);
            Assert.True(new GridCell(":--").IsAlignment);
            Assert.True(new GridCell(":-:").IsAlignment);
            Assert.True(new GridCell(":--:").IsAlignment);
            Assert.True(new GridCell("-:").IsAlignment);
            Assert.True(new GridCell("--:").IsAlignment);
            Assert.False(new GridCell("-:-").IsAlignment);
            Assert.False(new GridCell(string.Empty).IsAlignment);
            Assert.False(new GridCell(":").IsAlignment);
            Assert.False(new GridCell("a").IsAlignment);
            Assert.False(new GridCell("-a").IsAlignment);
            Assert.False(new GridCell("a-").IsAlignment);
        }

        [Fact]
        public void UndefinedAlignment()
        {
            Assert.Equal(Alignment.Undefined, new GridCell("-").Alignment);
            Assert.Equal(Alignment.Undefined, new GridCell("--").Alignment);
        }

        [Fact]
        public void LeftAlignment()
        {
            Assert.Equal(Alignment.Left, new GridCell(":-").Alignment);
            Assert.Equal(Alignment.Left, new GridCell(":--").Alignment);
        }

        [Fact]
        public void CenterAlignment()
        {
            Assert.Equal(Alignment.Center, new GridCell(":-:").Alignment);
            Assert.Equal(Alignment.Center, new GridCell(":--:").Alignment);
        }

        [Fact]
        public void RightAlignment()
        {
            Assert.Equal(Alignment.Right, new GridCell("-:").Alignment);
            Assert.Equal(Alignment.Right, new GridCell("--:").Alignment);
        }
    }
}
