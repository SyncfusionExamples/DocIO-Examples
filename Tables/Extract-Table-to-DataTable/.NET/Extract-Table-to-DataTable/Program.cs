using System.Text;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.XlsIO;

class Program
{
    static void Main(string[] args)
    {
        // Load existing word document
        using (FileStream inputfileStream = new FileStream(Path.GetFullPath(@"../../../Data/Input.docx"), FileMode.Open))
        {
            using (WordDocument document = new WordDocument(inputfileStream, FormatType.Automatic))
            {
                using (ExcelEngine engine = new ExcelEngine())
                {
                    IApplication app = engine.Excel;
                    app.DefaultVersion = ExcelVersion.Excel2016;

                    // Create one sheet to start with; we’ll add sheets as we find more tables.
                    IWorkbook workbook = app.Workbooks.Create(1);
                    int sheetIndex = 0;
                    int tableNumber = 0;

                    // Get table entities in word document
                    List<Entity> entities = document.FindAllItemsByProperty(EntityType.Table, null, null);

                    foreach (Entity entity in entities)
                    {
                        WTable wTable = (WTable)entity;

                        if (sheetIndex >= workbook.Worksheets.Count)
                            workbook.Worksheets.Create();

                        IWorksheet worksheet = workbook.Worksheets[sheetIndex++];
                        worksheet.Name = $"Table{++tableNumber}";

                        // Export with merges starting
                        ExportWordTableToExcelMerged(wTable, worksheet);

                        // Formatting
                        worksheet.UsedRange.AutofitRows();
                        worksheet.UsedRange.AutofitColumns();
                    }
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"../../../Output/Result.xlsx"), FileMode.Create))
                    {
                        workbook.SaveAs(outputStream);
                    }
                    workbook.Close();
                }                    
            }          
        }        
    }

    /// <summary>
	/// Writes a Word table to the worksheet, preserving horizontal and vertical merges.
	/// startRow/startCol are 1-based Excel coordinates where the table should be placed.
	/// </summary>
	static void ExportWordTableToExcelMerged(IWTable table, IWorksheet worksheet)
    {
        for (int r = 0; r < table.Rows.Count; r++)
        {
            WTableRow wRow = (WTableRow)table.Rows[r];

            // Map Word's logical grid to Excel columns using GridSpan
            int gridCol = 1;

            for (int i = 0; i < wRow.Cells.Count; i++)
            {
                WTableCell wCell = wRow.Cells[i];

                // Horizontal width in grid columns
                int hSpan = (int)wCell.GridSpan;

                // Merge flags
                CellMerge vFlag = wCell.CellFormat.VerticalMerge;   // None | Start | Continue
                CellMerge hFlag = wCell.CellFormat.HorizontalMerge;

                // Excel start cell for this Word cell
                int xRow = r + 1;
                int excelStartColIndex = gridCol;

                // Compute vertical span when this cell is the START of a vertical merge
                int vSpan = 1;
                if (vFlag == CellMerge.Start)
                {
                    // Count how many subsequent rows continue the merge at the same grid column
                    for (int nr = r + 1; nr < table.Rows.Count; nr++)
                    {
                        WTableRow nextRow = (WTableRow)table.Rows[nr];
                        WTableCell nextCell = GetCellAtGridColumn(nextRow, excelStartColIndex); // 1-based grid col
                        if (nextCell != null && nextCell.CellFormat.VerticalMerge == CellMerge.Continue)
                            vSpan++;
                        else
                            break;
                    }
                }
                if (hFlag == CellMerge.Start) 
                {
                    for( int nc = i + 1; nc < wRow.Cells.Count; nc++)
                    {
                        WTableCell cell = wRow.Cells[nc];
                        if (cell != null && cell.CellFormat.HorizontalMerge == CellMerge.Continue)
                            hSpan += cell.GridSpan;
                        else
                            break;
                    }
                }

                // Is Start or None of a merge region
                bool isCotinued =
                    (vFlag != CellMerge.Continue) &&
                    (hFlag != CellMerge.Continue);

                if (isCotinued)
                {
                    int vMergeEndIndex = xRow + vSpan - 1;
                    int hMergeEndColIndex = excelStartColIndex + hSpan - 1;

                    // Merge in Excel if region spans multiple cells
                    if (vMergeEndIndex > xRow || hMergeEndColIndex > excelStartColIndex)
                        worksheet.Range[xRow, excelStartColIndex, vMergeEndIndex, hMergeEndColIndex].Merge();

                    // Write the visible text to the top-left Excel cell
                    IRange range = worksheet.Range[xRow, excelStartColIndex];
                    range.Text = BuildCellText(wCell);

                    // Format styling
                    range.CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
                    range.CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter;

                    worksheet.Range[xRow, excelStartColIndex, vMergeEndIndex, hMergeEndColIndex].CellStyle.Borders[ExcelBordersIndex.EdgeLeft].LineStyle = ExcelLineStyle.Thin;
                    worksheet.Range[xRow, excelStartColIndex, vMergeEndIndex, hMergeEndColIndex].CellStyle.Borders[ExcelBordersIndex.EdgeRight].LineStyle = ExcelLineStyle.Thin;
                    worksheet.Range[xRow, excelStartColIndex, vMergeEndIndex, hMergeEndColIndex].CellStyle.Borders[ExcelBordersIndex.EdgeTop].LineStyle = ExcelLineStyle.Thin;
                    worksheet.Range[xRow, excelStartColIndex, vMergeEndIndex, hMergeEndColIndex].CellStyle.Borders[ExcelBordersIndex.EdgeBottom].LineStyle = ExcelLineStyle.Thin;

                }

                // Advance Excel column cursor by the horizontal span of this Word cell
                gridCol += hSpan;
            }
        }
    }

    /// <summary>
    /// Returns the WTableCell occupying the given 1-based "grid column" in this row,
    /// taking each cell's GridSpan into account.
    /// </summary>
    static WTableCell GetCellAtGridColumn(WTableRow row, int gridColumn)
    {
        int cursor = 1; // 1-based grid column within the table
        foreach (WTableCell c in row.Cells)
        {
            int span = (int)c.GridSpan;
            int start = cursor;
            int end = cursor + span - 1;
            if (gridColumn >= start && gridColumn <= end)
                return c;
            cursor += span;
        }
        return null;
    }

    /// <summary>
    /// Concatenate all paragraph texts in a Word cell (one per line).
    /// </summary>
    static string BuildCellText(WTableCell cell)
    {
        StringBuilder sb = new StringBuilder();
        for (int p = 0; p < cell.Paragraphs.Count; p++)
        {
            WParagraph para = cell.Paragraphs[p];
            string text = para.Text?.TrimEnd();
            if (!string.IsNullOrEmpty(text))
            {
                if (sb.Length > 0) sb.AppendLine();
                sb.Append(text);
            }
        }
        return sb.ToString();
    }
}