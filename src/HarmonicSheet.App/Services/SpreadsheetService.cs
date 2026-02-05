using ClosedXML.Excel;
using HarmonicSheet.Models;

namespace HarmonicSheet.Services;

/// <summary>
/// スプレッドシートファイル操作サービスの実装
/// </summary>
public class SpreadsheetService : ISpreadsheetService
{
    public async Task SaveToExcel(List<SpreadsheetRow> rows, string filePath)
    {
        await Task.Run(() =>
        {
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Sheet1");

            foreach (var row in rows)
            {
                foreach (var cell in row.Cells)
                {
                    var xlCell = worksheet.Cell(cell.Row, cell.Column - 'A' + 1);

                    if (!string.IsNullOrEmpty(cell.Formula))
                    {
                        xlCell.FormulaA1 = cell.Formula;
                    }
                    else if (!string.IsNullOrEmpty(cell.Value))
                    {
                        // 数値かどうか判定
                        if (double.TryParse(cell.Value, out var numValue))
                        {
                            xlCell.Value = numValue;
                        }
                        else
                        {
                            xlCell.Value = cell.Value;
                        }
                    }
                }
            }

            // シニア向けに見やすいフォント設定
            worksheet.Style.Font.FontSize = 14;
            worksheet.Style.Font.FontName = "Yu Gothic UI";
            worksheet.Columns().AdjustToContents();

            workbook.SaveAs(filePath);
        });
    }

    public async Task<List<SpreadsheetRow>> LoadFromExcel(string filePath)
    {
        return await Task.Run(() =>
        {
            var rows = new List<SpreadsheetRow>();

            using var workbook = new XLWorkbook(filePath);
            var worksheet = workbook.Worksheet(1);

            var usedRange = worksheet.RangeUsed();
            if (usedRange == null)
            {
                // 空のファイルの場合、デフォルトのグリッドを返す
                return CreateEmptyGrid(10, 5);
            }

            var lastRow = Math.Max(usedRange.LastRow().RowNumber(), 10);
            var lastCol = Math.Max(usedRange.LastColumn().ColumnNumber(), 5);

            for (int i = 1; i <= lastRow; i++)
            {
                var row = new SpreadsheetRow { RowNumber = i };

                for (int j = 1; j <= lastCol; j++)
                {
                    var xlCell = worksheet.Cell(i, j);
                    var cell = new SpreadsheetCell
                    {
                        Column = (char)('A' + j - 1),
                        Row = i
                    };

                    if (xlCell.HasFormula)
                    {
                        cell.Formula = xlCell.FormulaA1;
                        cell.Value = xlCell.Value.ToString();
                        cell.DisplayValue = xlCell.Value.ToString();
                    }
                    else
                    {
                        cell.Value = xlCell.Value.ToString();
                        cell.DisplayValue = cell.Value;
                    }

                    row.Cells.Add(cell);
                }

                rows.Add(row);
            }

            return rows;
        });
    }

    private List<SpreadsheetRow> CreateEmptyGrid(int rowCount, int colCount)
    {
        var rows = new List<SpreadsheetRow>();

        for (int i = 0; i < rowCount; i++)
        {
            var row = new SpreadsheetRow { RowNumber = i + 1 };
            for (int j = 0; j < colCount; j++)
            {
                row.Cells.Add(new SpreadsheetCell
                {
                    Column = (char)('A' + j),
                    Row = i + 1,
                    Value = string.Empty
                });
            }
            rows.Add(row);
        }

        return rows;
    }
}
