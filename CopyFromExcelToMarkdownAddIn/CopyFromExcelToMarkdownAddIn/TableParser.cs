using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CopyFromExcelToMarkdownAddIn
{
    public class TableParser
    {
        public Table Parse(Grid grid)
        {
            if (grid.HasAlignmentRows)
            {
                return ParseHasAlignmentRows(grid);
            }
            else
            {
                return ParseHasNotAlignmentRows(grid);
            }
        }

        private Table ParseHasAlignmentRows(Grid grid)
        {
            var table = new Table();
            table.AddRow(ParseRow(grid, 0));
            for (var i = 2; i < grid.Rows.Count; i++)
            {
                table.AddRow(ParseRow(grid, i));
            }
            return table;
        }

        private TableRow ParseRow(Grid grid, int rowIndex)
        {
            var tableRow = new TableRow();
            var gridRow = grid.Rows[rowIndex];
            for (int i = 0; i < gridRow.Count; i++)
            {
                tableRow.AddCell(new TableCell(gridRow[i].Value, GetAlignment(grid, i)));
            }

            return tableRow;
        }

        private Alignment GetAlignment(Grid grid, int columnIndex)
        {
            var row = grid.Rows[1];
            if (columnIndex < row.Count)
            {
                return row[columnIndex].Alignment;
            }
            else
            {
                return Alignment.Undefined;
            }
        }

        private Table ParseHasNotAlignmentRows(Grid grid)
        {
            var table = new Table();
            for (var i = 0; i < grid.Rows.Count; i++)
            {
                table.AddRow(ParseRow(grid, i));
            }
            return table;
        }
    }
}
