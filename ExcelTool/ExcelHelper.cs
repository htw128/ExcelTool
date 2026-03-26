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
        public static List<ParsedSheet> ParseAllSheets(string fileName, out List<ParsedEnum> enumSheets)
        {
            List<ParsedSheet> result = [];
            enumSheets = [];
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

                    if (sheetName.Equals("__enums__", StringComparison.OrdinalIgnoreCase))
                    {
                        enumSheets = ParseEnumSheet(sheet);
                        continue;
                    }

                    ParsedSheet parsed = ParseSheet(sheet, sheetName);
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

                List<TableExcelHeader> headers = [];
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
                List<TableExcelRow> tableRows = [];
                for (int i = 5; i <= sheet.LastRowNum; i++)
                {
                    IRow row = sheet.GetRow(i);
                    if (row == null) continue;

                    TableExcelRow tableExcelRow = new();
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

        static List<ParsedEnum> ParseEnumSheet(ISheet sheet)
        {
            List<ParsedEnum> result = [];
            // 列索引约定：A=0 Id, B=1 EnumName, C=2 items.Name, D=3 items.Value, E=4 items.Desc
            // 数据从第 6 行开始（0-indexed = 5）
            ParsedEnum currentEnum = null;

            for (int i = 5; i <= sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row == null) continue;

                string idStr      = GetCellValue(row.GetCell(0));
                string enumName   = GetCellValue(row.GetCell(1));
                string memberKey  = GetCellValue(row.GetCell(2));
                string memberVal  = GetCellValue(row.GetCell(3));
                string memberDesc = GetCellValue(row.GetCell(4));

                // 成员名为空则跳过此行
                if (string.IsNullOrEmpty(memberKey)) continue;

                // Id / EnumName 非空时，开启新枚举
                if (!string.IsNullOrEmpty(enumName))
                {
                    currentEnum = new ParsedEnum
                    {
                        Id       = string.IsNullOrEmpty(idStr) ? 0 : Convert.ToUInt32(double.Parse(idStr)),
                        EnumName = enumName,
                    };
                    result.Add(currentEnum);
                }

                // 如果没有任何枚举上下文就跳过
                currentEnum?.Members.Add(new ParsedEnumMember
                {
                    Key   = memberKey,
                    Value = string.IsNullOrEmpty(memberVal) ? null : Convert.ToInt32(double.Parse(memberVal)),
                    Desc  = memberDesc,
                });
            }

            return result;
        }

        static string GetCellValue(ICell cell)
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
