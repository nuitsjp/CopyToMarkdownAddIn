using System;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;

namespace CopyFromExcelToMarkdownAddIn
{
    public partial class ThisAddIn
    {
        /// <summary>
        /// Minimum required select the number of rows.
        /// </summary>
        private const int MinRowCount = 3;
        /// <summary>
        /// Alignment Left
        /// </summary>
        private const int AlignmentLeft = 1;
        /// <summary>
        /// Alignment Center
        /// </summary>
        private const int AlignmentCenter = -4108;
        /// <summary>
        /// Alignment Right
        /// </summary>
        private const int AlignmentRight = -4152;

        /// <summary>
        /// Button in ContextMenu.
        /// </summary>
        private CommandBarButton _copyToMarkdownButton;
        /// <summary>
        /// Startup event handler
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            // Create button in ContextMenu
            var menuItem = MsoControlType.msoControlButton;
            _copyToMarkdownButton = (CommandBarButton)Application.CommandBars["GridCell"].Controls.Add(menuItem, missing, missing, 1, true);
            _copyToMarkdownButton.Style = MsoButtonStyle.msoButtonCaption;
            _copyToMarkdownButton.Caption = "Copy to Markdown";
            _copyToMarkdownButton.Tag = "0";
            _copyToMarkdownButton.Click += CopyToMarkdown;
        }

        /// <summary>
        /// Shutdown event handler
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
        }

        /// <summary>
        /// CopyToMarkdown
        /// </summary>
        /// <param name="ctrl"></param>
        /// <param name="cancelDefault"></param>
        private void CopyToMarkdown(CommandBarButton ctrl, ref bool cancelDefault)
        {
            var range = Application.Selection as Range;
            if (range == null)
            {
                MessageBox.Show(Properties.Resources.UnselectedErrorMessage);
                return;
            }
            
            var rowsCount = range.Rows.Count;
            if (rowsCount < MinRowCount)
            {
                MessageBox.Show(Properties.Resources.UnselectedErrorMessage);
                return;
            }

            var columnsCount = range.Count / rowsCount;
            var resultBuffer = new StringBuilder();
            var separatorBuffer = new StringBuilder();
            for (int x = 1; x <= columnsCount; x++)
            {
                var cell = (Range)range.Cells[1, x];

                resultBuffer.Append("|");
                resultBuffer.Append(FormatText(cell));
                switch ((int)cell.HorizontalAlignment)
                {
                    case AlignmentCenter:
                        separatorBuffer.Append("|:-:");
                        break;
                    case AlignmentRight:
                        separatorBuffer.Append("|--:");
                        break;
                    default:
                        separatorBuffer.Append("|:--");
                        break;
                }
            }
            // Partition of the header and data lines.
            // Process only after the first line.
            resultBuffer.Append("|");
            resultBuffer.Append(Environment.NewLine);
            separatorBuffer.Append("|");
            separatorBuffer.Append(Environment.NewLine);
            resultBuffer.Append(separatorBuffer);

            for (int y = 2; y <= rowsCount; y++)
            {
                for (int x = 1; x <= columnsCount; x++)
                {
                    var cell = (Range)range.Cells[y, x];

                    resultBuffer.Append("|");
                    resultBuffer.Append(FormatText(cell));
                }
                resultBuffer.Append("|");
                resultBuffer.Append(Environment.NewLine);
            }
            Clipboard.SetText(resultBuffer.ToString());
        }

        private static string FormatText(Range range)
        {
            if (range == null || range.Text == null)
            {
                return string.Empty;
            }
            else
            {
                return range.Text.Replace("\n", "<br>");
            }
        }


        #region VSTO で生成されたコード

        /// <summary>
        /// デザイナーのサポートに必要なメソッドです。
        /// このメソッドの内容をコード エディターで変更しないでください。
        /// </summary>
        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }
        
        #endregion
    }
}
