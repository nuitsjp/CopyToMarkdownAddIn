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
        /// Alignment Undefined
        /// </summary>
        private const int AlignmentUndefined = 1;
        /// <summary>
        /// Alignment Left
        /// </summary>
        private const int AlignmentLeft = -4131;
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
        /// Button in ContextMenu.
        /// </summary>
        private CommandBarButton _copyFromMarkdownButton;
        /// <summary>
        /// Startup event handler
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            // Create button in ContextMenu
            _copyToMarkdownButton = (CommandBarButton)Application.CommandBars["Cell"].Controls.Add(MsoControlType.msoControlButton, missing, missing, 1, true);
            _copyToMarkdownButton.Style = MsoButtonStyle.msoButtonCaption;
            _copyToMarkdownButton.Caption = "Copy to Markdown";
            _copyToMarkdownButton.Tag = "0";
            _copyToMarkdownButton.Click += CopyToMarkdown;

            _copyFromMarkdownButton = (CommandBarButton)Application.CommandBars["Cell"].Controls.Add(MsoControlType.msoControlButton, missing, missing, 2, true);
            _copyFromMarkdownButton.Style = MsoButtonStyle.msoButtonCaption;
            _copyFromMarkdownButton.Caption = "Copy from Markdown";
            _copyFromMarkdownButton.Tag = "1";
            _copyFromMarkdownButton.Click += CopyFromMarkdown;
        }

        private void CopyFromMarkdown(CommandBarButton ctrl, ref bool canceldefault)
        {
            var text = Clipboard.GetText();
            if(string.IsNullOrEmpty(text))
                return;

            var range = Application.Selection as Range;
            if (range == null)
            {
                MessageBox.Show(Properties.Resources.UnselectedErrorMessage);
                return;
            }

            var table = new TableParser().Parse(new GridParser().Parse(text));
            var activeSheet = (Worksheet)Application.ActiveSheet;

            for (var i = 0; i < table.Rows.Count; i++)
            {
                var row = table.Rows[i];
                for (var j = 0; j < row.Count; j++)
                {
                    var cell = row[j];
                    var activeSheetCell =  (Range)activeSheet.Cells[range.Row + i, range.Column + j];
                    activeSheetCell.Value2 = cell.Value.Replace("<br>", Environment.NewLine).Replace("<br/>", Environment.NewLine);
                    switch (cell.Alignment)
                    {
                        case Alignment.Undefined:
                            activeSheetCell.HorizontalAlignment = AlignmentUndefined;
                            break;
                        case Alignment.Left:
                            activeSheetCell.HorizontalAlignment = AlignmentLeft;
                            break;
                        case Alignment.Center:
                            activeSheetCell.HorizontalAlignment = AlignmentCenter;
                            break;
                        case Alignment.Right:
                            activeSheetCell.HorizontalAlignment = AlignmentRight;
                            break;
                        default:
                            throw new ArgumentOutOfRangeException();
                    }
                }
            }
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
                    case AlignmentLeft:
                        separatorBuffer.Append("|:--");
                        break;
                    case AlignmentCenter:
                        separatorBuffer.Append("|:-:");
                        break;
                    case AlignmentRight:
                        separatorBuffer.Append("|--:");
                        break;
                    default:
                        separatorBuffer.Append("|--");
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
