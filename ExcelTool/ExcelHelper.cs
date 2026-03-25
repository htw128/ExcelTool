using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using ExcelTool.Parser;

namespace ExcelTool
{
    public static class ExcelHelper
    {
        public static List<ParsedSheet> ParseAllSheets(string fileName)
        {
            List<ParsedSheet> result = [];
            try
            {
                using FileStream fs = File.OpenRead(fileName);
                XSSFWorkbook wk = new(fs);
                int sheetCount = wk.NumberOfSheets;

                for (int sheetNum = 0; sheetNum < sheetCount; sheetNum++)
                {
                    ISheet sheet = wk.GetSheetAt(sheetNum);
                    string sheetName = sheet.SheetName;

                    // # 开头的 sheet 跳过
                    if (string.IsNullOrEmpty(sheetName) || sheetName.StartsWith('#'))
                        continue;

                    var parsed = ParseSheet(sheet, sheetName);
                    if (parsed != null)
                        result.Add(parsed);
                }
            }
            catch (Exception ex)
            {
                ex.ToString().WriteErrorLine();
            }
            return result;
        }

        static ParsedSheet ParseSheet(ISheet sheet, string sheetName)
        {
            try
            {
                IRow nameRow = sheet.GetRow(0);
                IRow typeRow = sheet.GetRow(1);
                IRow descRow = sheet.GetRow(2);

                var headers = new List<TableExcelHeader>();
                for (int j = 0; j < nameRow.LastCellNum; j++)
                {
                    string fieldName = nameRow.GetCell(j)?.ToString() ?? "";
                    string fieldType = typeRow.GetCell(j)?.ToString() ?? "";
                    string fieldDesc = descRow.GetCell(j)?.ToString() ?? "";

                    if (string.IsNullOrEmpty(fieldName))
                    {
                        $"列 {j} 字段名为空".WriteErrorLine();
                        continue;
                    }

                    headers.Add(new TableExcelHeader
                    {
                        FieldName = fieldName,
                        FieldType = fieldType,
                        FieldDesc = fieldDesc,
                    });
                }

                // 数据从第 6 行开始（0-indexed）
                var tableRows = new List<TableExcelRow>();
                for (int i = 5; i <= sheet.LastRowNum; i++)
                {
                    IRow row = sheet.GetRow(i);
                    if (row == null) continue;

                    var tableExcelRow = new TableExcelRow();
                    for (int j = 0; j < headers.Count; j++)
                        tableExcelRow.Add(GetCellValue(row.GetCell(j)));

                    tableRows.Add(tableExcelRow);
                }

                return new ParsedSheet
                {
                    SheetName = sheetName,
                    Headers   = headers,
                    Data      = new TableExcelData(headers, tableRows),
                };
            }
            catch (Exception ex)
            {
                ex.ToString().WriteErrorLine();
                return null;
            }
        }
        
        private static string GetCellValue(ICell cell)
        {
            if (cell == null)
                return "";

            switch (cell.CellType)
            {
                case CellType.String:
                    return cell.StringCellValue;

                case CellType.Numeric:
                    return DateUtil.IsCellDateFormatted(cell) ? cell.DateCellValue.ToString() : cell.NumericCellValue.ToString();

                case CellType.Boolean:
                    return cell.BooleanCellValue ? "1" : "0";

                case CellType.Formula:
                    return cell.ToString();

                default:
                    return "";
            }
        }
    }

    public class ParsedSheet
    {
        public string SheetName { get; init; }
        public List<TableExcelHeader> Headers { get; init; }
        public TableExcelData Data { get; init; }
    }
}
