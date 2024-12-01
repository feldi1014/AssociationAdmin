using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SplitOrders
{
    internal class WorksheetBuilder
    {
        internal void Fill(IXLWorksheet worksheet, List<DataRow> rows, DataTable table)
        {
            // Kopfzeile schreiben
            for (int i = 0; i < table.Columns.Count; i++)
            {
                worksheet.Cell(1, i + 1).Value = table.Columns[i].ColumnName;
            }

            // Tabelle füllen
            for (int rowIndex = 0; rowIndex < rows.Count(); rowIndex++)
            {
                for (int i = 0; i < table.Columns.Count; i++)
                {
                    worksheet.Cell(rowIndex + 2, i + 1).Value = rows.ElementAt(rowIndex)[i].ToString();
                }
            }
        }
    }
}
