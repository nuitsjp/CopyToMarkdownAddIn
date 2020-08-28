using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CopyFromExcelToMarkdownAddIn;
using FluentAssertions;
using Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;
using Moq;
using Xunit;

namespace CopyFromExcelToMarkdownAddIna.Test
{
    public class RangeExtensionsTest
    {
        [Fact]
        public void WhenRangeIsNull()
        {
            var range = (Range)null;
            range.FormatText()
                .Should().NotBeNull()
                .And.BeEmpty();
        }

        [Fact]
        public void WhenRangeTextIsNull()
        {
            var range = new Mock<Range>();
            range.SetupGet(x => x.Text).Returns(null);
            range.Object.FormatText()
                .Should().NotBeNull()
                .And.BeEmpty();
        }

        [Fact]
        public void WhenRangeTextIsEmpty()
        {
            var range = new Mock<Range>();
            range.SetupGet(x => x.Text).Returns(string.Empty);
            range.Object.FormatText()
                .Should().NotBeNull()
                .And.BeEmpty();
        }

        [Fact]
        public void WhenTextCoantaintsEnter()
        {
            var range = new Mock<Range>();
            range.SetupGet(x => x.Text).Returns("a\nb");
            range.Object.FormatText()
                .Should().Be("a<br>b");
        }

        [Fact]
        public void WhenTextCoantaintsPipe()
        {
            var range = new Mock<Range>();
            range.SetupGet(x => x.Text).Returns("a|b");
            range.Object.FormatText()
                .Should().Be("a&#124;b");
        }
    }
}


