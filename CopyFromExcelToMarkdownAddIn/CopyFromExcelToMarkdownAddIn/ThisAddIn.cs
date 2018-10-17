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
        /// Button in ContextMenu for Cell.
        /// </summary>
        private CommandBarButton _copyToMarkdownButtonForCell;
        /// <summary>
        /// Button in ContextMenu for Table.
        /// </summary>
        private CommandBarButton _copyToMarkDownButtonForTable;

        /// <summary>
        /// Button in ContextMenu for Cell.
        /// </summary>
        private CommandBarButton _copyFromMarkdownButtonForCell;
        /// <summary>
        /// Button in ContextMenu for Table.
        /// </summary>
        private CommandBarButton _copyFromMarkdownButtonForTable;

        /// <summary>
        /// Startup event handler
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            const string CELL = "Cell";
            const string TABLE = "List Range Popup";

            // Create button in ContextMenu for Cell
            _copyToMarkdownButtonForCell = CreateCopyToMarkdownButton(CELL, "0");
            _copyFromMarkdownButtonForCell = CreateCopyFromMarkdownButton(CELL, "1");
            // Create button in ContextMenu for Table
            _copyToMarkDownButtonForTable = CreateCopyToMarkdownButton(TABLE, "2");
            _copyFromMarkdownButtonForTable = CreateCopyFromMarkdownButton(TABLE, "3");
        }

        private CommandBarButton CreateCopyToMarkdownButton(string commandBarsKey, string tag)
        {
            var copyToMarkdownButton = (CommandBarButton)Application.CommandBars[commandBarsKey].Controls.Add(MsoControlType.msoControlButton, missing, missing, 1, true);
            copyToMarkdownButton.Style = MsoButtonStyle.msoButtonCaption;
            copyToMarkdownButton.Caption = "Copy to Markdown";
            copyToMarkdownButton.Tag = tag;
            copyToMarkdownButton.Click += CopyToMarkdown;
            return copyToMarkdownButton;
        }
        private CommandBarButton CreateCopyFromMarkdownButton(string commandBarsKey, string tag)
        {
            var copyFromMarkdownButton = (CommandBarButton)Application.CommandBars[commandBarsKey].Controls.Add(MsoControlType.msoControlButton, missing, missing, 2, true);
            copyFromMarkdownButton.Style = MsoButtonStyle.msoButtonCaption;
            copyFromMarkdownButton.Caption = "Paste from Markdown";
            copyFromMarkdownButton.Tag = tag;
            copyFromMarkdownButton.Click += CopyFromMarkdown;
            return copyFromMarkdownButton;
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
                    activeSheetCell.Value2 = cell.Value.Replace("<br>", "\n").Replace("<br/>", "\n");
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
